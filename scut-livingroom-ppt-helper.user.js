// ==UserScript==
// @name         华园视频(SCUT) PPT 导出助手 (PDF / PPTX)
// @namespace    https://github.com/peter-erer/buaa-spoc-helper
// @version      1.0.0
// @description  基于 MIT 开源项目 peter-erer/buaa-spoc-helper 适配华南理工大学华园视频（video.jw.scut.edu.cn 及 WebVPN），一键导出课程回放 PPT 为 PDF / PPTX，并支持导出字幕与笔记
// @author       micraow, Peter-erer (SCUT livingroom adaptation)
// @license      MIT
// @match        *://video.jw.scut.edu.cn/livingroom*
// @match        *://video-jw-443.webvpn.scut.edu.cn/livingroom*
// @require      https://cdn.jsdelivr.net/npm/pptxgenjs@3.9.0/dist/pptxgen.bundle.js
// @require      https://cdn.jsdelivr.net/npm/pdf-lib@1.17.1/dist/pdf-lib.min.js
// @grant        GM_xmlhttpRequest
// @grant        unsafeWindow
// @connect      *
// @run-at       document-start
// ==/UserScript==

/*
 * Credits & License:
 * Derived from peter-erer/buaa-spoc-helper (https://github.com/peter-erer/buaa-spoc-helper)
 * Original authors: micraow, Peter-erer (MIT License)
 * Adapted for SCUT 华园视频 (video.jw.scut.edu.cn / video-jw-443.webvpn.scut.edu.cn):
 * - Added livingroom direct PDF (pdf-lib) & PPTX (PptxGenJS) dual-button export
 * - Added CORS-safe image fetching via GM_xmlhttpRequest
 * - Added document-start XHR/fetch interception & proactive API fallback
 */

(function () {
  "use strict";

  const pageWindow = typeof unsafeWindow !== "undefined" ? unsafeWindow : window;
  if (pageWindow.__SCUT_PPT_HELPER_LOADED__) return;
  pageWindow.__SCUT_PPT_HELPER_LOADED__ = true;

  const IMAGE_FETCH_CONCURRENCY = 4;

  const state = {
    pptSlides: [],
    subtitleData: null,
    subInfo: null,
    capturedAuthHeader: "",
    lastCapturedPptUrl: "",
    busy: false,
    collapsed: false,
    ui: null,
  };

  // Install network interceptor immediately at document-start so no initial API response is missed.
  installNetworkInterceptor();

  ready(() => {
    initPanel();
    scheduleAutoProbe();
  });

  function ready(callback) {
    if (document.readyState === "loading") {
      document.addEventListener("DOMContentLoaded", callback, { once: true });
    } else {
      callback();
    }
  }

  function sleep(ms) {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  function getUrlParam(key) {
    const searchParams = new URLSearchParams(window.location.search);
    if (searchParams.has(key)) return searchParams.get(key) || "";
    const hash = window.location.hash || "";
    const queryIndex = hash.indexOf("?");
    if (queryIndex >= 0) {
      const hashParams = new URLSearchParams(hash.slice(queryIndex + 1));
      if (hashParams.has(key)) return hashParams.get(key) || "";
    }
    return "";
  }

  function escapeHtml(text) {
    return String(text || "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function sanitizeFilename(text, max = 90) {
    return (
      String(text || "SCUT_PPT")
        .replace(/[\\/:*?"<>|]+/g, "_")
        .replace(/\s+/g, " ")
        .trim()
        .slice(0, max) || "SCUT_PPT"
    );
  }

  function formatSrtTime(seconds) {
    const safeSeconds = Number.isFinite(seconds) ? Math.max(0, seconds) : 0;
    const totalMs = Math.floor(safeSeconds * 1000);
    const hours = String(Math.floor(totalMs / 3600000)).padStart(2, "0");
    const minutes = String(Math.floor((totalMs % 3600000) / 60000)).padStart(2, "0");
    const secs = String(Math.floor((totalMs % 60000) / 1000)).padStart(2, "0");
    const ms = String(totalMs % 1000).padStart(3, "0");
    return `${hours}:${minutes}:${secs},${ms}`;
  }

  function formatClockTime(seconds) {
    const safeSeconds = Number.isFinite(seconds) ? Math.max(0, Math.floor(seconds)) : 0;
    const hours = String(Math.floor(safeSeconds / 3600)).padStart(2, "0");
    const minutes = String(Math.floor((safeSeconds % 3600) / 60)).padStart(2, "0");
    const secs = String(safeSeconds % 60).padStart(2, "0");
    return `${hours}:${minutes}:${secs}`;
  }

  function normalizeImageUrl(rawUrl) {
    const trimmed = String(rawUrl || "").trim();
    if (!trimmed) return "";
    try {
      const resolved = new URL(trimmed, window.location.origin);
      // If accessed via SCUT WebVPN, rewrite internal video.jw.scut.edu.cn URLs to current WebVPN origin.
      if (
        window.location.hostname.includes("webvpn.scut.edu.cn") &&
        resolved.hostname === "video.jw.scut.edu.cn"
      ) {
        return `${window.location.origin}${resolved.pathname}${resolved.search}${resolved.hash}`;
      }
      return resolved.href;
    } catch {
      return trimmed;
    }
  }

  function getExportBaseName() {
    if (state.subInfo) {
      const parts = [state.subInfo.course_title, state.subInfo.sub_title].filter(Boolean);
      if (parts.length) return sanitizeFilename(parts.join("_"));
    }
    const pageTitle = (document.title || "").trim();
    const subId = getUrlParam("sub_id");
    if (pageTitle && pageTitle !== "华园视频" && pageTitle !== "课程回放") {
      return sanitizeFilename(pageTitle);
    }
    return sanitizeFilename(subId ? `华园视频_PPT_${subId}` : "华园视频_PPT");
  }

  function getJwtTokenFromCookie() {
    try {
      const decoded = decodeURIComponent(document.cookie || "");
      const match = decoded.match(/eyJ[\w-]+\.[\w-]+\.[\w-]+/);
      return match ? match[0] : "";
    } catch {
      return "";
    }
  }

  // ==================== Network Interception ====================

  function installNetworkInterceptor() {
    // 1. Hook XMLHttpRequest on the real page window
    const XHR = pageWindow.XMLHttpRequest;
    if (XHR && XHR.prototype) {
      const originalOpen = XHR.prototype.open;
      const originalSetHeader = XHR.prototype.setRequestHeader;
      const originalSend = XHR.prototype.send;

      XHR.prototype.open = function (_method, url) {
        this.__scutHelperUrl = String(url || "");
        return originalOpen.apply(this, arguments);
      };

      XHR.prototype.setRequestHeader = function (name, value) {
        try {
          if (String(name || "").toLowerCase() === "authorization" && value) {
            state.capturedAuthHeader = String(value);
          }
        } catch {
          // ignore
        }
        return originalSetHeader.apply(this, arguments);
      };

      XHR.prototype.send = function () {
        this.addEventListener(
          "load",
          function () {
            try {
              const url = this.__scutHelperUrl || this.responseURL || "";
              if (!url) return;
              if (this.responseType && this.responseType !== "" && this.responseType !== "text") {
                return;
              }
              handleInterceptedResponse(url, this.responseText);
            } catch {
              // ignore
            }
          },
          { once: true }
        );
        return originalSend.apply(this, arguments);
      };
    }

    // 2. Hook fetch on the real page window
    if (typeof pageWindow.fetch === "function") {
      const originalFetch = pageWindow.fetch;
      pageWindow.fetch = function (input, init) {
        try {
          const headers = init && init.headers;
          if (headers) {
            if (headers instanceof Headers) {
              const auth = headers.get("Authorization") || headers.get("authorization");
              if (auth) state.capturedAuthHeader = auth;
            } else if (typeof headers === "object") {
              const auth = headers.Authorization || headers.authorization;
              if (auth) state.capturedAuthHeader = String(auth);
            }
          }
        } catch {
          // ignore
        }

        const url =
          typeof input === "string"
            ? input
            : input && typeof input.url === "string"
              ? input.url
              : "";

        return originalFetch.apply(this, arguments).then((response) => {
          try {
            const finalUrl = response.url || url;
            if (isInterestingUrl(finalUrl)) {
              response
                .clone()
                .text()
                .then((text) => handleInterceptedResponse(finalUrl, text))
                .catch(() => {});
            }
          } catch {
            // ignore
          }
          return response;
        });
      };
    }
  }

  function isInterestingUrl(url) {
    return (
      url.includes("search-ppt") ||
      url.includes("pptnote") ||
      url.includes("search-trans-result") ||
      url.includes("get-sub-info")
    );
  }

  function handleInterceptedResponse(url, text) {
    if (!isInterestingUrl(url) || !text) return;

    let payload;
    try {
      payload = JSON.parse(text);
    } catch {
      return;
    }

    if (url.includes("get-sub-info")) {
      const info = payload && (payload.data || payload.result || payload);
      if (info && typeof info === "object" && (info.course_title || info.sub_title || info.resource_guid)) {
        state.subInfo = info;
      }
      return;
    }

    if (url.includes("search-trans-result")) {
      if (Array.isArray(payload.list) && payload.list.length > 0) {
        state.subtitleData = payload.list;
        updateUi();
      }
      return;
    }

    if (url.includes("search-ppt") || url.includes("pptnote")) {
      const slides = extractSlidesFromPayload(payload);
      if (slides.length > 0) {
        state.pptSlides = slides;
        state.lastCapturedPptUrl = url;
        updateUi();
      } else {
        console.info("[SCUT PPT Helper] 捕获到接口但未提取出 PPT 图片:", url, payload);
      }
    }
  }

  function extractSlidesFromPayload(payload) {
    if (!payload || typeof payload !== "object") return [];

    // Locate the array of slide records
    const candidates = [
      payload.list,
      payload.data && payload.data.list,
      payload.data && payload.data.ppt_list,
      payload.data && payload.data.slides,
      Array.isArray(payload.data) ? payload.data : null,
      payload.result && payload.result.list,
    ];

    const rawList = candidates.find((arr) => Array.isArray(arr) && arr.length > 0);
    if (!rawList) return [];

    const slides = [];
    let lastUrl = "";

    rawList.forEach((item) => {
      if (!item || typeof item !== "object") return;

      let content = {};
      if (typeof item.content === "string") {
        try {
          content = JSON.parse(item.content);
        } catch {
          content = {};
        }
      } else if (item.content && typeof item.content === "object") {
        content = item.content;
      }

      const rawUrl =
        content.pptimgurl ||
        content.pptImgUrl ||
        content.imageUrl ||
        content.img_url ||
        content.ppt_img_url ||
        content.url ||
        item.pptimgurl ||
        item.pptImgUrl ||
        item.imageUrl ||
        item.img_url ||
        item.ppt_img_url ||
        item.pic_url ||
        item.url ||
        "";

      const imageUrl = normalizeImageUrl(rawUrl);
      if (!imageUrl) return;
      // Skip consecutive duplicate URLs
      if (imageUrl === lastUrl) return;
      lastUrl = imageUrl;

      const sec = Number(
        item.created_sec ?? item.BeginSec ?? item.begin_sec ?? content.created_sec ?? 0
      );
      slides.push({
        imageUrl,
        index: slides.length + 1,
        timeText: formatClockTime(sec),
      });
    });

    return slides;
  }

  // ==================== Proactive API Fallback ====================

  function scheduleAutoProbe() {
    setTimeout(() => {
      if (!state.pptSlides.length && getUrlParam("course_id") && getUrlParam("sub_id")) {
        proactivelyFetchPpt({ silent: true });
      }
    }, 3000);
  }

  async function apiGetJson(path) {
    const url = new URL(path, window.location.origin).href;
    const headers = {};
    const auth = state.capturedAuthHeader || (getJwtTokenFromCookie() ? `Bearer ${getJwtTokenFromCookie()}` : "");
    if (auth) headers.Authorization = auth;

    const response = await fetch(url, {
      method: "GET",
      credentials: "include",
      headers,
    });
    if (!response.ok) {
      throw new Error(`HTTP ${response.status}`);
    }
    return response.json();
  }

  async function proactivelyFetchPpt({ silent = false } = {}) {
    const courseId = getUrlParam("course_id");
    const subId = getUrlParam("sub_id");
    if (!courseId || !subId) {
      if (!silent) {
        setStatus("当前网址中未找到 course_id 或 sub_id，请先点击页面内的「PPT」标签触发加载。", "warning");
      }
      return;
    }

    if (!silent) {
      setStatus("正在主动向华园视频接口请求当前课时的 PPT 数据…");
    }

    try {
      // 1. Fetch sub info for title & resource_guid if missing or switched to another lesson
      if (!state.subInfo || String(state.subInfo.sub_id || state.subInfo.id || "") !== String(subId)) {
        try {
          const infoResp = await apiGetJson(
            `/courseapi/v3/portal-home-setting/get-sub-info?course_id=${encodeURIComponent(courseId)}&sub_id=${encodeURIComponent(subId)}`
          );
          if (infoResp && infoResp.data) {
            state.subInfo = { ...infoResp.data, sub_id: infoResp.data.sub_id || subId };
          }
        } catch {
          // Non-fatal: continue to query search-ppt directly
        }
      }

      const guid = state.subInfo && state.subInfo.resource_guid ? state.subInfo.resource_guid : "";
      const endpoints = [];
      if (guid) {
        endpoints.push(
          `/pptnote/v1/schedule/search-ppt?course_id=${encodeURIComponent(courseId)}&sub_id=${encodeURIComponent(subId)}&resource_guid=${encodeURIComponent(guid)}`
        );
      }
      endpoints.push(
        `/pptnote/v1/schedule/search-ppt?course_id=${encodeURIComponent(courseId)}&sub_id=${encodeURIComponent(subId)}`
      );
      endpoints.push(
        `/pptnote/v1/get/course?course_id=${encodeURIComponent(courseId)}&sub_id=${encodeURIComponent(subId)}`
      );

      for (const endpoint of endpoints) {
        try {
          const data = await apiGetJson(endpoint);
          const slides = extractSlidesFromPayload(data);
          if (slides.length > 0) {
            state.pptSlides = slides;
            state.lastCapturedPptUrl = endpoint;
            updateUi();
            if (!silent) {
              setStatus(`已成功拉取 ${slides.length} 页 PPT，现在可以导出 PDF 或 PPTX。`);
            }
            return;
          }
        } catch {
          // Try next endpoint
        }
      }

      if (!silent) {
        setStatus(
          "主动拉取未获取到 PPT 列表。请点击一下页面里的「PPT / 课件」标签，或刷新页面后再试。",
          "warning"
        );
      }
    } catch (error) {
      console.warn("[SCUT PPT Helper] 主动拉取 PPT 失败", error);
      if (!silent) {
        setStatus("主动拉取失败，请点击页面内的「PPT」标签触发加载。", "warning");
      }
    }
  }

  // ==================== CORS-Safe Image Fetching & Processing ====================

  function fetchImageBlob(url) {
    return new Promise((resolve, reject) => {
      if (typeof GM_xmlhttpRequest === "function") {
        GM_xmlhttpRequest({
          method: "GET",
          url,
          responseType: "blob",
          timeout: 25000,
          headers: {
            Referer: window.location.href,
          },
          onload: (res) => {
            if (res.status >= 200 && res.status < 300 && res.response) {
              resolve(res.response);
            } else {
              reject(new Error(`图片请求失败 (HTTP ${res.status})`));
            }
          },
          onerror: () => reject(new Error("图片网络请求出错")),
          ontimeout: () => reject(new Error("图片请求超时")),
        });
        return;
      }

      fetch(url, { credentials: "include" })
        .then((res) => {
          if (!res.ok) throw new Error(`HTTP ${res.status}`);
          return res.blob();
        })
        .then(resolve)
        .catch(reject);
    });
  }

  function blobToArrayBuffer(blob) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = () => reject(new Error("读取图片字节失败"));
      reader.readAsArrayBuffer(blob);
    });
  }

  function blobToDataUrl(blob) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(String(reader.result || ""));
      reader.onerror = () => reject(new Error("读取图片 DataURL 失败"));
      reader.readAsDataURL(blob);
    });
  }

  function detectImageMime(bytes) {
    if (bytes.length >= 3 && bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff) {
      return "image/jpeg";
    }
    if (
      bytes.length >= 8 &&
      bytes[0] === 0x89 &&
      bytes[1] === 0x50 &&
      bytes[2] === 0x4e &&
      bytes[3] === 0x47
    ) {
      return "image/png";
    }
    return "unknown";
  }

  function loadImageFromBlob(blob) {
    return new Promise((resolve, reject) => {
      const objectUrl = URL.createObjectURL(blob);
      const img = new Image();
      img.onload = () => {
        URL.revokeObjectURL(objectUrl);
        resolve(img);
      };
      img.onerror = () => {
        URL.revokeObjectURL(objectUrl);
        reject(new Error("图片解码失败"));
      };
      img.src = objectUrl;
    });
  }

  function clamp255(value) {
    return value < 0 ? 0 : value > 255 ? 255 : Math.round(value);
  }

  function applySharpen(imageData, width, height, strength) {
    const source = new Uint8ClampedArray(imageData.data);
    const target = imageData.data;
    const kernel = [0, -strength, 0, -strength, 1 + 4 * strength, -strength, 0, -strength, 0];

    for (let y = 1; y < height - 1; y += 1) {
      for (let x = 1; x < width - 1; x += 1) {
        const targetIndex = (y * width + x) * 4;
        for (let channel = 0; channel < 3; channel += 1) {
          let sum = 0;
          let k = 0;
          for (let oy = -1; oy <= 1; oy += 1) {
            for (let ox = -1; ox <= 1; ox += 1) {
              sum += source[((y + oy) * width + (x + ox)) * 4 + channel] * kernel[k];
              k += 1;
            }
          }
          target[targetIndex + channel] = clamp255(sum);
        }
      }
    }
  }

  async function processBlobWithCanvas(blob, { enhance = false, quality = 0.92, maxDimension = 2560 } = {}) {
    const img = await loadImageFromBlob(blob);
    const naturalWidth = img.naturalWidth || 1280;
    const naturalHeight = img.naturalHeight || 720;

    const scale = Math.min(1, maxDimension / Math.max(naturalWidth, naturalHeight));
    const width = Math.max(1, Math.round(naturalWidth * scale));
    const height = Math.max(1, Math.round(naturalHeight * scale));

    const canvas = document.createElement("canvas");
    canvas.width = width;
    canvas.height = height;
    const ctx = canvas.getContext("2d", { willReadFrequently: enhance });
    if (!ctx) throw new Error("无法创建 Canvas");

    // White background so transparent PNGs don't turn black in JPEG
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, width, height);
    ctx.drawImage(img, 0, 0, width, height);

    if (enhance) {
      const brightness = 1.18;
      const yellowFix = 0.03;
      const sharpness = 0.25;
      const imageData = ctx.getImageData(0, 0, width, height);
      const data = imageData.data;
      const redGain = brightness * (1 - yellowFix * 0.5);
      const greenGain = brightness * (1 - yellowFix * 0.25);
      const blueGain = brightness * (1 + yellowFix);
      for (let i = 0; i < data.length; i += 4) {
        data[i] = clamp255(data[i] * redGain);
        data[i + 1] = clamp255(data[i + 1] * greenGain);
        data[i + 2] = clamp255(data[i + 2] * blueGain);
      }
      applySharpen(imageData, width, height, sharpness);
      ctx.putImageData(imageData, 0, 0);
    }

    const outBlob = await new Promise((resolve, reject) => {
      canvas.toBlob(
        (b) => (b ? resolve(b) : reject(new Error("Canvas 转码失败"))),
        "image/jpeg",
        quality
      );
    });
    const bytes = new Uint8Array(await blobToArrayBuffer(outBlob));
    const dataUrl = await blobToDataUrl(outBlob);
    return { bytes, dataUrl, mime: "image/jpeg", width, height };
  }

  async function prepareSlideAsset(slide, enhance) {
    const blob = await fetchImageBlob(slide.imageUrl);
    const rawBytes = new Uint8Array(await blobToArrayBuffer(blob));
    const mime = detectImageMime(rawBytes);

    if (enhance || mime === "unknown") {
      return processBlobWithCanvas(blob, { enhance });
    }

    const img = await loadImageFromBlob(blob);
    const width = img.naturalWidth || 1280;
    const height = img.naturalHeight || 720;
    const typedBlob = new Blob([rawBytes], { type: mime });
    const dataUrl = await blobToDataUrl(typedBlob);
    return { bytes: rawBytes, dataUrl, mime, width, height };
  }

  async function runWithConcurrency(items, concurrency, worker, onProgress) {
    const results = new Array(items.length);
    let nextIndex = 0;
    let completed = 0;

    async function runner() {
      while (nextIndex < items.length) {
        const current = nextIndex;
        nextIndex += 1;
        try {
          results[current] = await worker(items[current], current);
        } catch (error) {
          console.warn(`[SCUT PPT Helper] 第 ${current + 1} 页处理失败`, error);
          results[current] = null;
        }
        completed += 1;
        if (onProgress) onProgress(completed, items.length);
      }
    }

    const workers = Array.from({ length: Math.min(concurrency, items.length) }, () => runner());
    await Promise.all(workers);
    return results;
  }

  function downloadBlob(blob, fileName) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 5000);
  }

  function confirmLargeEnhancement(count) {
    return (
      count <= 50 ||
      window.confirm(`当前共有 ${count} 页 PPT，开启图像增强会耗时较久并占用较多内存，是否继续？`)
    );
  }

  // ==================== Export 1: Direct PDF (pdf-lib) ====================

  async function exportDirectPdf() {
    const slides = state.pptSlides;
    if (!slides.length) {
      window.alert("还没有捕获到 PPT 数据，请先点击页面内的「PPT」标签，或点击面板里的「主动拉取 PPT」。");
      return;
    }

    // Fallback to print-preview PDF if pdf-lib failed to load from CDN
    const pdfLibRef = typeof PDFLib !== "undefined" ? PDFLib : window.PDFLib;
    if (!pdfLibRef || !pdfLibRef.PDFDocument) {
      setStatus("pdf-lib 未加载成功，自动切换为横版打印页（可在打印窗口选择「另存为 PDF」）。");
      await exportPrintPreviewPdf();
      return;
    }

    const enhance = Boolean(state.ui?.enhanceCheckbox?.checked) && confirmLargeEnhancement(slides.length);
    setBusyState(true, "正在准备导出 PDF…");

    try {
      const assets = await runWithConcurrency(
        slides,
        IMAGE_FETCH_CONCURRENCY,
        (slide) => prepareSlideAsset(slide, enhance),
        (done, total) => {
          setStatus(`正在下载${enhance ? "并增强" : ""} PPT 图片…`);
          setProgress(done, total, enhance ? "下载并增强图片" : "下载图片");
        }
      );

      const validAssets = assets.filter(Boolean);
      if (!validAssets.length) {
        throw new Error("所有 PPT 图片下载均失败，请检查登录状态或网络连接。");
      }

      setStatus(`正在合成 PDF（共 ${validAssets.length} 页）…`);
      setProgress(null, null, "合成 PDF");
      const pdfDoc = await pdfLibRef.PDFDocument.create();
      pdfDoc.setTitle(getExportBaseName());
      pdfDoc.setCreator("SCUT PPT Helper (based on buaa-spoc-helper)");

      const targetWidth = 842; // A4 landscape width in points (72 DPI), ideal for iPad/GoodNotes pen thickness
      for (let i = 0; i < validAssets.length; i += 1) {
        const asset = validAssets[i];
        const embedded =
          asset.mime === "image/png"
            ? await pdfDoc.embedPng(asset.bytes)
            : await pdfDoc.embedJpg(asset.bytes);

        const ratio = embedded.height / embedded.width;
        const pageWidth = targetWidth;
        const pageHeight = Math.round(targetWidth * ratio * 100) / 100;
        const page = pdfDoc.addPage([pageWidth, pageHeight]);
        page.drawImage(embedded, {
          x: 0,
          y: 0,
          width: pageWidth,
          height: pageHeight,
        });

        if (i % 10 === 0) await sleep(0);
      }

      const pdfBytes = await pdfDoc.save();
      const fileName = `${getExportBaseName()}.pdf`;
      downloadBlob(new Blob([pdfBytes], { type: "application/pdf" }), fileName);

      const failedCount = slides.length - validAssets.length;
      setStatus(
        failedCount > 0
          ? `PDF 已导出：${fileName}（成功 ${validAssets.length} 页，跳过 ${failedCount} 页失败图片）。`
          : `PDF 导出完成：${fileName}（共 ${validAssets.length} 页）。`
      );
    } catch (error) {
      console.error("[SCUT PPT Helper] PDF 导出失败", error);
      setStatus(`PDF 导出失败：${error.message || "请重试或改用打印版 PDF"}`, "warning");
    } finally {
      setBusyState(false);
    }
  }

  // ==================== Export 2: Direct PPTX (PptxGenJS) ====================

  async function exportDirectPptx() {
    const slides = state.pptSlides;
    if (!slides.length) {
      window.alert("还没有捕获到 PPT 数据，请先点击页面内的「PPT」标签，或点击面板里的「主动拉取 PPT」。");
      return;
    }

    const PptxCtor = typeof PptxGenJS !== "undefined" ? PptxGenJS : window.PptxGenJS;
    if (typeof PptxCtor !== "function") {
      window.alert("PptxGenJS 库未能从 CDN 加载，请检查网络连接后刷新页面重试。");
      return;
    }

    const enhance = Boolean(state.ui?.enhanceCheckbox?.checked) && confirmLargeEnhancement(slides.length);
    setBusyState(true, "正在准备导出 PPTX…");

    try {
      const assets = await runWithConcurrency(
        slides,
        IMAGE_FETCH_CONCURRENCY,
        (slide) => prepareSlideAsset(slide, enhance),
        (done, total) => {
          setStatus(`正在下载${enhance ? "并增强" : ""} PPT 图片…`);
          setProgress(done, total, enhance ? "下载并增强图片" : "下载图片");
        }
      );

      const validAssets = assets.filter(Boolean);
      if (!validAssets.length) {
        throw new Error("所有 PPT 图片下载均失败，请检查登录状态或网络连接。");
      }

      setStatus(`正在打包 PPTX（共 ${validAssets.length} 页）…`);
      setProgress(null, null, "打包 PPTX");

      // Choose 4:3 or 16:9 based on the first slide's aspect ratio
      const firstRatio = validAssets[0].width / validAssets[0].height;
      const is4x3 = firstRatio < 1.5;
      const slideW = 10;
      const slideH = is4x3 ? 7.5 : 5.625;

      const ppt = new PptxCtor();
      ppt.layout = is4x3 ? "LAYOUT_4x3" : "LAYOUT_16x9";
      ppt.title = getExportBaseName();
      ppt.subject = (state.subInfo && state.subInfo.course_title) || "SCUT Course";
      ppt.author = (state.subInfo && state.subInfo.lecturer_name) || "SCUT";

      validAssets.forEach((asset) => {
        const slide = ppt.addSlide();
        const imgRatio = asset.width / asset.height;
        const slideRatio = slideW / slideH;
        let w;
        let h;
        if (imgRatio > slideRatio) {
          w = slideW;
          h = slideW / imgRatio;
        } else {
          h = slideH;
          w = slideH * imgRatio;
        }
        const x = (slideW - w) / 2;
        const y = (slideH - h) / 2;
        slide.addImage({
          data: asset.dataUrl,
          x,
          y,
          w,
          h,
        });
      });

      const fileName = `${getExportBaseName()}.pptx`;
      await ppt.writeFile({ fileName });

      const failedCount = slides.length - validAssets.length;
      setStatus(
        failedCount > 0
          ? `PPTX 已导出：${fileName}（成功 ${validAssets.length} 页，跳过 ${failedCount} 页失败图片）。`
          : `PPTX 导出完成：${fileName}（共 ${validAssets.length} 页）。`
      );
    } catch (error) {
      console.error("[SCUT PPT Helper] PPTX 导出失败", error);
      setStatus(`PPTX 导出失败：${error.message || "请稍后重试"}`, "warning");
    } finally {
      setBusyState(false);
    }
  }

  // ==================== Export 3: Print-Preview PDF (from buaa-spoc-helper) ====================

  async function exportPrintPreviewPdf() {
    const slides = state.pptSlides;
    if (!slides.length) {
      window.alert("还没有捕获到 PPT 数据，请先点击页面内的「PPT」标签。");
      return;
    }

    const title = escapeHtml(getExportBaseName());
    const generatedAt = escapeHtml(new Date().toLocaleString("zh-CN"));

    const slidesMarkup = slides
      .map(
        (slide) => `
          <section class="slide-page">
            <img class="ppt-image" src="${escapeHtml(slide.imageUrl)}" alt="第 ${slide.index} 页">
            <div class="slide-meta">第 ${slide.index} 页 · ${escapeHtml(slide.timeText)}</div>
          </section>
        `
      )
      .join("");

    const html = `
      <!DOCTYPE html>
      <html lang="zh-CN">
        <head>
          <meta charset="UTF-8">
          <title>${title} - 横版 PDF</title>
          <style>
            * { box-sizing: border-box; }
            html, body {
              margin: 0;
              min-height: 100%;
              background: #eef2f7;
              color: #0f172a;
              font-family: "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif;
            }
            .toolbar {
              position: sticky;
              top: 0;
              z-index: 10;
              display: flex;
              flex-wrap: wrap;
              align-items: center;
              justify-content: space-between;
              gap: 12px;
              padding: 14px 18px;
              background: rgba(255,255,255,0.96);
              border-bottom: 1px solid rgba(148, 163, 184, 0.35);
            }
            .toolbar strong { display: block; font-size: 15px; }
            .toolbar span { display: block; margin-top: 3px; color: #475569; font-size: 12px; }
            .toolbar button {
              border: 0;
              border-radius: 10px;
              padding: 9px 14px;
              background: #2563eb;
              color: white;
              cursor: pointer;
            }
            .slide-page {
              position: relative;
              width: min(1280px, calc(100vw - 48px));
              aspect-ratio: 16 / 9;
              margin: 24px auto;
              display: flex;
              align-items: center;
              justify-content: center;
              overflow: hidden;
              background: white;
              box-shadow: 0 18px 42px rgba(15, 23, 42, 0.12);
            }
            .ppt-image {
              display: block;
              width: 100%;
              height: 100%;
              object-fit: contain;
              background: white;
            }
            .slide-meta {
              position: absolute;
              right: 10px;
              bottom: 8px;
              padding: 4px 7px;
              border-radius: 999px;
              background: rgba(15, 23, 42, 0.68);
              color: white;
              font-size: 11px;
            }
            @media print {
              @page { size: A4 landscape; margin: 0; }
              html, body { width: 297mm; height: 210mm; margin: 0 !important; padding: 0 !important; background: white !important; }
              .toolbar, .slide-meta { display: none !important; }
              .slide-page {
                width: 297mm;
                height: 210mm;
                margin: 0 !important;
                box-shadow: none !important;
                break-after: page;
                page-break-after: always;
              }
              .slide-page:last-child { break-after: auto; page-break-after: auto; }
            }
          </style>
        </head>
        <body>
          <div class="toolbar">
            <div>
              <strong>${title}</strong>
              <span>共 ${slides.length} 页 · 生成时间：${generatedAt}</span>
            </div>
            <button type="button" onclick="window.print()">打印 / 另存为 PDF</button>
          </div>
          ${slidesMarkup}
          <script>
            window.onload = function () {
              setTimeout(function () { window.print(); }, 400);
            };
          <\/script>
        </body>
      </html>
    `;

    const popup = window.open("", "_blank");
    if (!popup) {
      window.alert("浏览器拦截了弹窗，请允许当前站点弹窗后再试。");
      return;
    }
    const blobUrl = URL.createObjectURL(new Blob([html], { type: "text/html;charset=utf-8" }));
    popup.location.replace(blobUrl);
    setTimeout(() => URL.revokeObjectURL(blobUrl), 60000);
  }

  // ==================== Export 4: Subtitles (from buaa-spoc-helper) ====================

  function getSubtitleLines() {
    return (state.subtitleData || []).reduce((all, chapter) => {
      if (Array.isArray(chapter.all_content)) {
        all.push(...chapter.all_content);
      }
      return all;
    }, []);
  }

  function mergeSubtitleLines(rawLines) {
    if (!rawLines.length) return [];
    const merged = [];
    let current = { ...rawLines[0] };
    for (let i = 1; i < rawLines.length; i += 1) {
      const next = rawLines[i];
      const gap = Number(next.BeginSec || 0) - Number(current.EndSec || current.BeginSec || 0);
      const currentText = String(current.Text || "");
      if (gap < 1 && currentText.length < 20) {
        current.Text = `${currentText}，${next.Text || ""}`;
        current.EndSec = next.EndSec;
      } else {
        merged.push(current);
        current = { ...next };
      }
    }
    merged.push(current);
    return merged;
  }

  function exportSubtitle(type) {
    const rawLines = getSubtitleLines();
    if (!rawLines.length) {
      window.alert("还没有捕获到字幕数据，请先点击页面内的语音/字幕标签。");
      return;
    }
    const lines = type === "txt" ? mergeSubtitleLines(rawLines) : rawLines;
    const baseName = getExportBaseName();

    let content = type === "txt" ? `课程笔记 - ${baseName}\n\n` : "";
    lines.forEach((line, index) => {
      const begin = Number(line.BeginSec || 0);
      const end = Number(line.EndSec || begin + 2);
      const text = String(line.Text || "").trim();
      if (!text) return;
      if (type === "srt") {
        content += `${index + 1}\n${formatSrtTime(begin)} --> ${formatSrtTime(end)}\n${text}\n\n`;
      } else {
        content += `[${formatClockTime(begin)}] ${text}\n`;
      }
    });

    downloadBlob(
      new Blob([content], { type: "text/plain;charset=utf-8" }),
      `${baseName}_${type === "srt" ? "字幕" : "笔记"}.${type}`
    );
  }

  // ==================== Floating UI (redesigned, Shadow DOM isolated) ====================

  const POS_KEY = "__scut_ppt_helper_pos_v2__";
  const COLLAPSE_KEY = "__scut_ppt_helper_collapsed_v2__";
  const INITIAL_HINT = "点一下网页里的「PPT」标签，或点右侧 ↻ 主动拉取课件。";

  const ICONS = {
    slides:
      '<path d="M2 3h20"/><path d="M21 3v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V3"/><path d="m7 21 5-5 5 5"/>',
    file:
      '<path d="M15 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7Z"/><path d="M14 2v4a2 2 0 0 0 2 2h4"/><path d="M9 13h6"/><path d="M9 17h4"/>',
    layout:
      '<rect x="3" y="4" width="18" height="14" rx="2"/><path d="M7 9h6"/><path d="M7 13h10"/><path d="M12 18v3"/><path d="M8 21h8"/>',
    download:
      '<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><path d="m7 10 5 5 5-5"/><path d="M12 15V3"/>',
    refresh:
      '<path d="M3 12a9 9 0 0 1 9-9 9.75 9.75 0 0 1 6.74 2.74L21 8"/><path d="M21 3v5h-5"/><path d="M21 12a9 9 0 0 1-9 9 9.75 9.75 0 0 1-6.74-2.74L3 16"/><path d="M8 16H3v5"/>',
    minus: '<path d="M5 12h14"/>',
    chevron: '<path d="m6 9 6 6 6-6"/>',
    printer:
      '<path d="M6 9V2h12v7"/><path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/><rect x="6" y="14" width="12" height="8" rx="1"/>',
    captions:
      '<rect x="3" y="5" width="18" height="14" rx="2"/><path d="M7 15h4"/><path d="M15 15h2"/><path d="M7 11h2"/><path d="M13 11h4"/>',
    notes: '<path d="M17 6H3"/><path d="M21 12H3"/><path d="M15 18H3"/>',
    sparkles:
      '<path d="M12 3 13.9 8.8 20 10l-6.1 1.2L12 17l-1.9-5.8L4 10l6.1-1.2Z"/><path d="M19 17v4"/><path d="M17 19h4"/>',
  };

  function icon(name, cls = "ic") {
    return `<svg class="${cls}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${ICONS[name] || ""}</svg>`;
  }

  const PANEL_CSS = `
    :host { all: initial; }
    *, *::before, *::after { box-sizing: border-box; }

    .root {
      --bg: rgba(255, 255, 255, 0.84);
      --bg-solid: #ffffff;
      --surface: rgba(15, 23, 42, 0.035);
      --surface-hover: rgba(15, 23, 42, 0.07);
      --border: rgba(15, 23, 42, 0.08);
      --text: #0f172a;
      --text-2: #475569;
      --text-3: #94a3b8;
      --accent: #4f46e5;
      --accent-2: #8b5cf6;
      --pdf: #e11d48;
      --pptx: #ea580c;
      --ok: #10b981;
      --warn: #f59e0b;
      --warn-text: #b45309;
      --ok-text: #047857;
      --shadow:
        0 1px 2px rgba(15, 23, 42, 0.06),
        0 14px 34px -10px rgba(15, 23, 42, 0.25),
        0 30px 60px -30px rgba(79, 70, 229, 0.35);
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif;
      font-size: 13px;
      line-height: 1.45;
      color: var(--text);
      -webkit-font-smoothing: antialiased;
    }

    @media (prefers-color-scheme: dark) {
      .root {
        --bg: rgba(17, 24, 39, 0.82);
        --bg-solid: #161e2e;
        --surface: rgba(255, 255, 255, 0.045);
        --surface-hover: rgba(255, 255, 255, 0.09);
        --border: rgba(255, 255, 255, 0.09);
        --text: #f1f5f9;
        --text-2: #cbd5e1;
        --text-3: #64748b;
        --warn-text: #fbbf24;
        --ok-text: #34d399;
        --shadow: 0 1px 2px rgba(0, 0, 0, 0.3), 0 18px 44px -10px rgba(0, 0, 0, 0.6);
      }
    }

    .ic { width: 16px; height: 16px; flex: none; display: block; }

    /* ---------- Panel ---------- */
    .panel {
      position: fixed;
      z-index: 2147483000;
      width: 304px;
      background: var(--bg);
      -webkit-backdrop-filter: blur(22px) saturate(1.7);
      backdrop-filter: blur(22px) saturate(1.7);
      border: 1px solid var(--border);
      border-radius: 20px;
      box-shadow: var(--shadow);
      overflow: hidden;
      animation: pop 0.32s cubic-bezier(0.2, 0.9, 0.3, 1.15);
    }
    .panel[hidden], .fab[hidden] { display: none !important; }

    @keyframes pop {
      from { opacity: 0; transform: translateY(10px) scale(0.96); }
      to { opacity: 1; transform: none; }
    }

    .header {
      display: flex;
      align-items: center;
      gap: 10px;
      padding: 14px 12px 10px 14px;
      cursor: grab;
      user-select: none;
      touch-action: none;
    }
    .header.dragging { cursor: grabbing; }

    .logo {
      width: 34px;
      height: 34px;
      flex: none;
      display: grid;
      place-items: center;
      border-radius: 11px;
      color: #fff;
      background: linear-gradient(135deg, var(--accent), var(--accent-2));
      box-shadow: 0 6px 16px -6px rgba(79, 70, 229, 0.7), inset 0 1px 0 rgba(255, 255, 255, 0.25);
    }
    .logo .ic { width: 18px; height: 18px; }

    .title { font-size: 14px; font-weight: 650; letter-spacing: 0.2px; }
    .subtitle { margin-top: 1px; font-size: 11px; color: var(--text-3); }

    .icon-btn {
      flex: none;
      width: 30px;
      height: 30px;
      display: grid;
      place-items: center;
      padding: 0;
      border: 0;
      border-radius: 9px;
      color: var(--text-2);
      background: transparent;
      cursor: pointer;
      transition: background 0.15s, color 0.15s;
    }
    .icon-btn:hover:not(:disabled) { background: var(--surface-hover); color: var(--text); }
    .icon-btn:disabled { opacity: 0.4; cursor: not-allowed; }
    .header .icon-btn { margin-left: auto; }

    .body { padding: 2px 14px 14px; }

    /* ---------- Hero ---------- */
    .hero {
      display: flex;
      align-items: center;
      gap: 12px;
      padding: 12px 10px 12px 14px;
      border: 1px solid var(--border);
      border-radius: 15px;
      background: var(--surface);
    }
    .hero-main { flex: 1; min-width: 0; }
    .count {
      font-size: 28px;
      font-weight: 700;
      line-height: 1;
      letter-spacing: -0.6px;
      font-variant-numeric: tabular-nums;
    }
    .count small { margin-left: 4px; font-size: 12px; font-weight: 500; letter-spacing: 0; color: var(--text-3); }
    .count.empty { color: var(--text-3); }

    .state { display: flex; align-items: center; gap: 7px; margin-top: 7px; font-size: 11.5px; color: var(--text-2); }
    .dot { width: 7px; height: 7px; flex: none; border-radius: 50%; background: var(--warn); box-shadow: 0 0 0 3px rgba(245, 158, 11, 0.2); }
    .dot.waiting { animation: pulse 1.6s ease-in-out infinite; }
    .dot.ready { background: var(--ok); box-shadow: 0 0 0 3px rgba(16, 185, 129, 0.2); }
    .dot.busy { background: var(--accent); box-shadow: 0 0 0 3px rgba(79, 70, 229, 0.22); animation: pulse 0.9s ease-in-out infinite; }

    @keyframes pulse { 50% { opacity: 0.35; } }

    .chips { display: flex; gap: 6px; margin-top: 8px; flex-wrap: wrap; }
    .chip {
      display: inline-flex;
      align-items: center;
      gap: 4px;
      padding: 2px 8px;
      border-radius: 99px;
      font-size: 10.5px;
      color: var(--text-3);
      background: var(--surface-hover);
    }
    .chip.on { color: var(--ok-text); background: rgba(16, 185, 129, 0.12); }

    .hero .icon-btn { width: 34px; height: 34px; border-radius: 11px; background: var(--bg-solid); border: 1px solid var(--border); }
    .spin .ic { animation: spin 0.8s linear infinite; }
    @keyframes spin { to { transform: rotate(360deg); } }

    /* ---------- Export cards ---------- */
    .exports { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-top: 12px; }
    .export {
      position: relative;
      display: flex;
      flex-direction: column;
      align-items: flex-start;
      gap: 9px;
      padding: 12px;
      border: 1px solid var(--border);
      border-radius: 15px;
      background: var(--bg-solid);
      color: var(--text);
      font: inherit;
      text-align: left;
      cursor: pointer;
      transition: transform 0.18s ease, box-shadow 0.18s ease, border-color 0.18s ease;
    }
    .export:hover:not(:disabled) {
      transform: translateY(-2px);
      border-color: var(--c);
      box-shadow: 0 12px 24px -14px var(--c);
    }
    .export:active:not(:disabled) { transform: translateY(0); }
    .export:disabled { opacity: 0.45; cursor: not-allowed; }
    .export.pdf { --c: var(--pdf); }
    .export.pptx { --c: var(--pptx); }

    .file-badge {
      width: 32px;
      height: 32px;
      display: grid;
      place-items: center;
      border-radius: 10px;
      color: #fff;
      background: var(--c);
      box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.25);
    }
    .file-badge .ic { width: 17px; height: 17px; }
    .export-name { font-size: 13px; font-weight: 650; }
    .export-desc { margin-top: 1px; font-size: 11px; color: var(--text-3); }
    .export .arrow { position: absolute; top: 12px; right: 12px; color: var(--text-3); transition: color 0.18s, transform 0.18s; }
    .export:hover:not(:disabled) .arrow { color: var(--c); transform: translateY(2px); }

    /* ---------- Progress ---------- */
    .progress { margin-top: 12px; }
    .track { height: 6px; overflow: hidden; border-radius: 99px; background: var(--surface-hover); }
    .fill {
      width: 0;
      height: 100%;
      border-radius: inherit;
      background: linear-gradient(90deg, var(--accent), var(--accent-2));
      transition: width 0.25s ease;
    }
    .fill.indeterminate { width: 38% !important; animation: indet 1.15s ease-in-out infinite; }
    @keyframes indet { from { transform: translateX(-100%); } to { transform: translateX(270%); } }
    .progress-meta {
      display: flex;
      justify-content: space-between;
      margin-top: 6px;
      font-size: 11px;
      color: var(--text-3);
      font-variant-numeric: tabular-nums;
    }

    /* ---------- Status ---------- */
    .status { margin-top: 10px; font-size: 11.5px; line-height: 1.55; color: var(--text-2); user-select: text; }
    .status.warning { color: var(--warn-text); }
    .status.success { color: var(--ok-text); }

    /* ---------- More ---------- */
    details.more { margin-top: 12px; padding-top: 6px; border-top: 1px solid var(--border); }
    summary {
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: 6px 4px;
      list-style: none;
      font-size: 12px;
      color: var(--text-2);
      cursor: pointer;
      user-select: none;
      border-radius: 8px;
    }
    summary::-webkit-details-marker { display: none; }
    summary:hover { color: var(--text); }
    summary .ic { transition: transform 0.2s; }
    details[open] summary .ic { transform: rotate(180deg); }

    .list { display: flex; flex-direction: column; gap: 1px; margin-top: 4px; }
    .row {
      display: flex;
      align-items: center;
      gap: 10px;
      width: 100%;
      padding: 8px;
      border: 0;
      border-radius: 10px;
      background: transparent;
      color: var(--text);
      font: inherit;
      font-size: 12.5px;
      text-align: left;
      cursor: pointer;
      transition: background 0.15s;
    }
    .row:hover:not(:disabled) { background: var(--surface-hover); }
    .row:disabled { opacity: 0.4; cursor: not-allowed; }
    .row > .ic { color: var(--text-2); }
    .row .hint { margin-left: auto; font-size: 11px; color: var(--text-3); }

    .switch { position: relative; flex: none; width: 34px; height: 20px; margin-left: auto; }
    .switch input { position: absolute; width: 0; height: 0; opacity: 0; }
    .slider { position: absolute; inset: 0; border-radius: 99px; background: rgba(148, 163, 184, 0.55); transition: background 0.2s; }
    .slider::before {
      content: "";
      position: absolute;
      top: 2px;
      left: 2px;
      width: 16px;
      height: 16px;
      border-radius: 50%;
      background: #fff;
      box-shadow: 0 1px 3px rgba(0, 0, 0, 0.25);
      transition: transform 0.2s cubic-bezier(0.3, 1.3, 0.5, 1);
    }
    .switch input:checked + .slider { background: var(--accent); }
    .switch input:checked + .slider::before { transform: translateX(14px); }
    .switch input:focus-visible + .slider { outline: 2px solid var(--accent); outline-offset: 2px; }

    .credit { margin-top: 8px; padding: 0 8px; font-size: 10.5px; color: var(--text-3); }
    .credit a { color: inherit; }

    /* ---------- Collapsed FAB ---------- */
    .fab {
      position: fixed;
      z-index: 2147483000;
      width: 50px;
      height: 50px;
      display: grid;
      place-items: center;
      padding: 0;
      border: 0;
      border-radius: 17px;
      color: #fff;
      background: linear-gradient(135deg, var(--accent), var(--accent-2));
      box-shadow: 0 12px 26px -8px rgba(79, 70, 229, 0.7), inset 0 1px 0 rgba(255, 255, 255, 0.25);
      cursor: pointer;
      touch-action: none;
      animation: pop 0.28s cubic-bezier(0.2, 0.9, 0.3, 1.15);
      transition: transform 0.15s;
    }
    .fab:hover { transform: scale(1.06); }
    .fab .ic { width: 23px; height: 23px; }
    .fab-num {
      position: absolute;
      top: -6px;
      right: -6px;
      min-width: 21px;
      height: 21px;
      padding: 0 6px;
      display: grid;
      place-items: center;
      border: 2px solid var(--bg-solid);
      border-radius: 99px;
      background: var(--ok);
      color: #fff;
      font-size: 10.5px;
      font-weight: 700;
      font-variant-numeric: tabular-nums;
    }
    .fab-num:empty { display: none; }

    button:focus-visible, summary:focus-visible { outline: 2px solid var(--accent); outline-offset: 2px; }

    @media (prefers-reduced-motion: reduce) {
      *, *::before, *::after { animation: none !important; transition: none !important; }
    }
  `;

  function readStore(key, fallback) {
    try {
      const raw = window.localStorage.getItem(key);
      return raw ? JSON.parse(raw) : fallback;
    } catch {
      return fallback;
    }
  }

  function writeStore(key, value) {
    try {
      window.localStorage.setItem(key, JSON.stringify(value));
    } catch {
      // ignore
    }
  }

  function initPanel() {
    const existing = document.getElementById("scut-ppt-helper-host");
    if (existing) existing.remove();

    const host = document.createElement("div");
    host.id = "scut-ppt-helper-host";
    const shadow = host.attachShadow({ mode: "open" });

    shadow.innerHTML = `
      <style>${PANEL_CSS}</style>
      <div class="root">
        <section class="panel" role="dialog" aria-label="华园视频 PPT 导出助手">
          <header class="header">
            <div class="logo">${icon("slides")}</div>
            <div>
              <div class="title">PPT 导出助手</div>
              <div class="subtitle">华园视频 · SCUT</div>
            </div>
            <button type="button" class="icon-btn" data-action="collapse" data-nodrag title="收起为悬浮按钮">${icon("minus")}</button>
          </header>

          <div class="body">
            <div class="hero">
              <div class="hero-main">
                <div class="count empty" data-role="count">—</div>
                <div class="state"><span class="dot waiting" data-role="dot"></span><span data-role="state-text">等待捕获课件</span></div>
                <div class="chips">
                  <span class="chip" data-role="chip-ppt">课件</span>
                  <span class="chip" data-role="chip-sub">字幕</span>
                </div>
              </div>
              <button type="button" class="icon-btn" data-action="probe" title="主动拉取 PPT">${icon("refresh")}</button>
            </div>

            <div class="exports">
              <button type="button" class="export pdf" data-action="pdf" disabled>
                <span class="file-badge">${icon("file")}</span>
                <span>
                  <div class="export-name">导出 PDF</div>
                  <div class="export-desc">原图无损 · 适合阅读</div>
                </span>
                <span class="arrow">${icon("download")}</span>
              </button>
              <button type="button" class="export pptx" data-action="pptx" disabled>
                <span class="file-badge">${icon("layout")}</span>
                <span>
                  <div class="export-name">导出 PPTX</div>
                  <div class="export-desc">逐页可编辑</div>
                </span>
                <span class="arrow">${icon("download")}</span>
              </button>
            </div>

            <div class="progress" data-role="progress" hidden>
              <div class="track"><div class="fill" data-role="fill"></div></div>
              <div class="progress-meta"><span data-role="progress-label">处理中</span><span data-role="progress-num"></span></div>
            </div>

            <div class="status" data-role="status">${escapeHtml(INITIAL_HINT)}</div>

            <details class="more">
              <summary>更多选项 ${icon("chevron")}</summary>
              <div class="list">
                <label class="row">
                  ${icon("sparkles")}
                  <span>图片增强<span class="hint" style="margin-left:6px">调亮 / 去黄 / 锐化</span></span>
                  <span class="switch"><input type="checkbox" data-role="enhance"><span class="slider"></span></span>
                </label>
                <button type="button" class="row" data-action="print-pdf" disabled>${icon("printer")}<span>打印预览版 PDF</span><span class="hint">浏览器打印</span></button>
                <button type="button" class="row" data-action="srt" disabled>${icon("captions")}<span>导出 SRT 字幕</span><span class="hint">.srt</span></button>
                <button type="button" class="row" data-action="txt" disabled>${icon("notes")}<span>导出 TXT 笔记</span><span class="hint">.txt</span></button>
              </div>
              <div class="credit">基于 <a href="https://github.com/peter-erer/buaa-spoc-helper" target="_blank" rel="noopener noreferrer">buaa-spoc-helper</a> 改编 · MIT</div>
            </details>
          </div>
        </section>

        <button type="button" class="fab" data-action="expand" title="展开 PPT 导出助手" hidden>
          ${icon("slides")}
          <span class="fab-num" data-role="fab-num"></span>
        </button>
      </div>
    `;

    document.body.appendChild(host);

    const $ = (selector) => shadow.querySelector(selector);
    state.ui = {
      host,
      shadow,
      panel: $(".panel"),
      header: $(".header"),
      fab: $(".fab"),
      fabNum: $('[data-role="fab-num"]'),
      count: $('[data-role="count"]'),
      dot: $('[data-role="dot"]'),
      stateText: $('[data-role="state-text"]'),
      chipPpt: $('[data-role="chip-ppt"]'),
      chipSub: $('[data-role="chip-sub"]'),
      status: $('[data-role="status"]'),
      progress: $('[data-role="progress"]'),
      fill: $('[data-role="fill"]'),
      progressLabel: $('[data-role="progress-label"]'),
      progressNum: $('[data-role="progress-num"]'),
      collapseBtn: $('[data-action="collapse"]'),
      probeBtn: $('[data-action="probe"]'),
      pdfBtn: $('[data-action="pdf"]'),
      pptxBtn: $('[data-action="pptx"]'),
      printPdfBtn: $('[data-action="print-pdf"]'),
      srtBtn: $('[data-action="srt"]'),
      txtBtn: $('[data-action="txt"]'),
      enhanceCheckbox: $('[data-role="enhance"]'),
      pos: readStore(POS_KEY, { right: 20, top: 96 }),
      probing: false,
    };

    const ui = state.ui;

    ui.collapseBtn.addEventListener("click", () => setCollapsed(true));
    ui.fab.addEventListener("click", () => {
      if (ui.fab.__dragged) {
        ui.fab.__dragged = false;
        return;
      }
      setCollapsed(false);
    });
    ui.probeBtn.addEventListener("click", async () => {
      if (ui.probing) return;
      ui.probing = true;
      ui.probeBtn.classList.add("spin");
      updateUi();
      try {
        await proactivelyFetchPpt({ silent: false });
      } finally {
        ui.probing = false;
        ui.probeBtn.classList.remove("spin");
        updateUi();
      }
    });
    ui.pdfBtn.addEventListener("click", exportDirectPdf);
    ui.pptxBtn.addEventListener("click", exportDirectPptx);
    ui.printPdfBtn.addEventListener("click", exportPrintPreviewPdf);
    ui.srtBtn.addEventListener("click", () => exportSubtitle("srt"));
    ui.txtBtn.addEventListener("click", () => exportSubtitle("txt"));

    makeDraggable(ui.panel, ui.header);
    makeDraggable(ui.fab, ui.fab);
    window.addEventListener("resize", () => applyPosition());

    state.collapsed = Boolean(readStore(COLLAPSE_KEY, false));
    setCollapsed(state.collapsed, { persist: false });
    updateUi();
  }

  function setCollapsed(collapsed, { persist = true } = {}) {
    const ui = state.ui;
    if (!ui) return;
    state.collapsed = collapsed;
    ui.panel.hidden = collapsed;
    ui.fab.hidden = !collapsed;
    if (persist) writeStore(COLLAPSE_KEY, collapsed);
    applyPosition();
  }

  function applyPosition() {
    const ui = state.ui;
    if (!ui) return;
    const el = state.collapsed ? ui.fab : ui.panel;
    const width = el.offsetWidth || (state.collapsed ? 50 : 304);
    const height = el.offsetHeight || (state.collapsed ? 50 : 360);
    const right = Math.min(Math.max(8, Number(ui.pos.right) || 20), Math.max(8, window.innerWidth - width - 8));
    const top = Math.min(Math.max(8, Number(ui.pos.top) || 96), Math.max(8, window.innerHeight - height - 8));
    [ui.panel, ui.fab].forEach((node) => {
      node.style.right = `${right}px`;
      node.style.top = `${top}px`;
      node.style.left = "auto";
    });
  }

  function makeDraggable(el, handle) {
    let drag = null;

    handle.addEventListener("pointerdown", (event) => {
      if (event.button !== 0) return;
      if (event.target instanceof Element && event.target.closest("[data-nodrag]")) return;
      const rect = el.getBoundingClientRect();
      drag = { x: event.clientX, y: event.clientY, left: rect.left, top: rect.top, moved: false };
      try {
        handle.setPointerCapture(event.pointerId);
      } catch {
        // ignore
      }
    });

    handle.addEventListener("pointermove", (event) => {
      if (!drag) return;
      const dx = event.clientX - drag.x;
      const dy = event.clientY - drag.y;
      if (!drag.moved && Math.hypot(dx, dy) < 4) return;
      drag.moved = true;
      handle.classList.add("dragging");
      const width = el.offsetWidth;
      const height = el.offsetHeight;
      const left = Math.min(window.innerWidth - width - 8, Math.max(8, drag.left + dx));
      const top = Math.min(window.innerHeight - height - 8, Math.max(8, drag.top + dy));
      state.ui.pos = { right: window.innerWidth - left - width, top };
      applyPosition();
    });

    const end = () => {
      if (!drag) return;
      const moved = drag.moved;
      drag = null;
      handle.classList.remove("dragging");
      if (moved) {
        el.__dragged = true;
        setTimeout(() => {
          el.__dragged = false;
        }, 0);
        writeStore(POS_KEY, state.ui.pos);
      }
    };
    handle.addEventListener("pointerup", end);
    handle.addEventListener("pointercancel", end);
  }

  function setStatus(text, tone) {
    if (!state.ui || !state.ui.status) return;
    let resolvedTone = tone;
    if (!resolvedTone || resolvedTone === "info") {
      resolvedTone = /完成|已导出|已成功|已捕获/.test(text) ? "success" : "info";
    }
    state.ui.status.textContent = text;
    state.ui.status.classList.toggle("warning", resolvedTone === "warning");
    state.ui.status.classList.toggle("success", resolvedTone === "success");
  }

  function setProgress(done, total, label) {
    const ui = state.ui;
    if (!ui) return;
    ui.progress.hidden = false;
    if (label) ui.progressLabel.textContent = label;
    if (Number.isFinite(done) && Number.isFinite(total) && total > 0) {
      ui.fill.classList.remove("indeterminate");
      ui.fill.style.width = `${Math.round((done / total) * 100)}%`;
      ui.progressNum.textContent = `${done} / ${total}`;
    } else {
      ui.fill.classList.add("indeterminate");
      ui.progressNum.textContent = "";
    }
  }

  function hideProgress() {
    const ui = state.ui;
    if (!ui) return;
    ui.progress.hidden = true;
    ui.fill.style.width = "0";
    ui.fill.classList.remove("indeterminate");
  }

  function setBusyState(busy, statusText) {
    state.busy = busy;
    if (statusText) setStatus(statusText);
    if (busy) {
      setProgress(null, null, "准备中");
    } else {
      hideProgress();
    }
    updateUi();
  }

  function updateUi() {
    const ui = state.ui;
    if (!ui) return;
    const pptCount = state.pptSlides.length;
    const pptReady = pptCount > 0;
    const subReady = Array.isArray(state.subtitleData) && state.subtitleData.length > 0;

    if (pptReady) {
      ui.count.innerHTML = `${pptCount}<small>页</small>`;
      ui.count.classList.remove("empty");
    } else {
      ui.count.textContent = "—";
      ui.count.classList.add("empty");
    }

    ui.dot.className = "dot";
    if (state.busy || ui.probing) {
      ui.dot.classList.add("busy");
      ui.stateText.textContent = state.busy ? "正在导出…" : "正在拉取课件…";
    } else if (pptReady) {
      ui.dot.classList.add("ready");
      ui.stateText.textContent = "课件已就绪";
    } else {
      ui.dot.classList.add("waiting");
      ui.stateText.textContent = "等待捕获课件";
    }

    ui.chipPpt.classList.toggle("on", pptReady);
    ui.chipPpt.textContent = pptReady ? "✓ 课件" : "课件";
    ui.chipSub.classList.toggle("on", subReady);
    ui.chipSub.textContent = subReady ? "✓ 字幕" : "字幕";

    ui.fabNum.textContent = pptReady ? String(pptCount) : "";

    ui.pdfBtn.disabled = state.busy || !pptReady;
    ui.pptxBtn.disabled = state.busy || !pptReady;
    ui.printPdfBtn.disabled = state.busy || !pptReady;
    ui.probeBtn.disabled = state.busy || ui.probing;
    ui.srtBtn.disabled = state.busy || !subReady;
    ui.txtBtn.disabled = state.busy || !subReady;

    if (!state.busy && pptReady && ui.status.textContent === INITIAL_HINT) {
      setStatus(`已捕获 ${pptCount} 页课件，点击上方按钮即可一键导出。`);
    }
  }
})();

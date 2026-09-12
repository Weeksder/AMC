/* AMC Productivity Studio — browser port (JSZip + optional Tesseract.js) */
(function () {
  "use strict";

  // Bump this whenever you re-upload amc-tools.js. If Chrome and Edge show
  // different version strings, one browser is still using a cached script.
  var TOOL_VERSION = "2026-09-12-editname";
  var WHITEBG_MAX_BYTES = 500 * 1024; // 500 KB
  try {
    console.log("[AMC Studio] script version", TOOL_VERSION);
  } catch (e) {}

  function $(id) {
    return document.getElementById(id);
  }

  /** Yield so the browser can paint labels/status before heavy zip work. */
  function yieldToUI() {
    return new Promise(function (resolve) {
      setTimeout(resolve, 0);
    });
  }

  /** Load Tesseract only when OCR is used (not on every page open). */
  var tesseractLoadPromise = null;
  function ensureTesseract() {
    if (typeof Tesseract !== "undefined") return Promise.resolve();
    if (tesseractLoadPromise) return tesseractLoadPromise;
    tesseractLoadPromise = new Promise(function (resolve, reject) {
      var s = document.createElement("script");
      s.src = "https://cdn.jsdelivr.net/npm/tesseract.js@5/dist/tesseract.min.js";
      s.async = true;
      s.onload = function () {
        resolve();
      };
      s.onerror = function () {
        tesseractLoadPromise = null;
        reject(new Error("Could not load OCR engine (need internet once)."));
      };
      document.head.appendChild(s);
    });
    return tesseractLoadPromise;
  }

  function setStatus(el, msg, kind) {
    if (!el) return;
    el.textContent = msg || "";
    el.className = "status" + (kind ? " " + kind : "");
  }

  function downloadBlob(blob, name) {
    try {
      // Prefer msSaveOrOpenBlob on legacy Edge
      if (window.navigator && window.navigator.msSaveOrOpenBlob) {
        window.navigator.msSaveOrOpenBlob(blob, name);
        return;
      }
    } catch (e) {}
    var a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = name || "download.pptx";
    a.rel = "noopener";
    a.style.display = "none";
    document.body.appendChild(a);
    // Synchronous click inside user gesture chain (required by Brave/Edge)
    a.click();
    setTimeout(function () {
      URL.revokeObjectURL(a.href);
      a.remove();
    }, 2500);
  }

  /**
   * Prefer writing over an existing file handle (Chrome/Edge).
   * Else save-picker. Else normal download (Safari).
   */
  async function savePptxResult(blob, fileName, fileHandle) {
    var name = fileName || "result.pptx";
    if (fileHandle && fileHandle.createWritable) {
      try {
        var writable = await fileHandle.createWritable();
        await writable.write(blob);
        await writable.close();
        return { mode: "replaced", name: name };
      } catch (e) {
        console.warn("Overwrite handle failed, falling back to download:", e);
      }
    }
    if (window.showSaveFilePicker) {
      try {
        var handle = await window.showSaveFilePicker({
          suggestedName: name,
          types: [
            {
              description: "PowerPoint",
              accept: {
                "application/vnd.openxmlformats-officedocument.presentationml.presentation": [
                  ".pptx",
                ],
              },
            },
          ],
        });
        var w2 = await handle.createWritable();
        await w2.write(blob);
        await w2.close();
        return { mode: "saved", name: name };
      } catch (err) {
        if (err && err.name === "AbortError") return { mode: "cancelled" };
      }
    }
    downloadBlob(blob, name);
    return { mode: "download", name: name };
  }

  async function pickPptxWithHandle() {
    return pickFileWithHandle(false);
  }

  /** @param {boolean} allowPdf — Source can be PPTX or PDF */
  async function pickFileWithHandle(allowPdf) {
    if (!window.showOpenFilePicker) return null;
    try {
      var types = [
        {
          description: "PowerPoint",
          accept: {
            "application/vnd.openxmlformats-officedocument.presentationml.presentation": [
              ".pptx",
            ],
          },
        },
      ];
      if (allowPdf) {
        types.push({
          description: "PDF",
          accept: { "application/pdf": [".pdf"] },
        });
      }
      var handles = await window.showOpenFilePicker({
        multiple: false,
        types: types,
      });
      return handles && handles[0] ? handles[0] : null;
    } catch (e) {
      if (e && e.name === "AbortError") return null;
      return null;
    }
  }

  function isPptxName(name) {
    return /\.pptx$/i.test(name || "");
  }
  function isPdfName(name) {
    return /\.pdf$/i.test(name || "");
  }
  function isSourceFile(f) {
    return f && (isPptxName(f.name) || isPdfName(f.name));
  }

  /** Lazy-load Mozilla PDF.js (renders pages as screenshots including text). */
  var pdfJsLoadPromise = null;
  function ensurePdfJs() {
    if (window.pdfjsLib) return Promise.resolve(window.pdfjsLib);
    if (pdfJsLoadPromise) return pdfJsLoadPromise;
    pdfJsLoadPromise = new Promise(function (resolve, reject) {
      var s = document.createElement("script");
      s.src = "https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.min.js";
      s.async = true;
      s.onload = function () {
        try {
          if (!window.pdfjsLib) throw new Error("pdfjsLib not found");
          window.pdfjsLib.GlobalWorkerOptions.workerSrc =
            "https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js";
          resolve(window.pdfjsLib);
        } catch (e) {
          pdfJsLoadPromise = null;
          reject(e);
        }
      };
      s.onerror = function () {
        pdfJsLoadPromise = null;
        reject(new Error("Could not load PDF.js (need internet once)."));
      };
      document.head.appendChild(s);
    });
    return pdfJsLoadPromise;
  }

  function canvasToJpegBytes(canvas, quality) {
    return new Promise(function (resolve, reject) {
      if (!canvas.toBlob) {
        try {
          var dataUrl = canvas.toDataURL("image/jpeg", quality || 0.92);
          var bin = atob(dataUrl.split(",")[1]);
          var bytes = new Uint8Array(bin.length);
          for (var i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
          resolve(bytes);
        } catch (e) {
          reject(e);
        }
        return;
      }
      canvas.toBlob(
        function (blob) {
          if (!blob) {
            reject(new Error("Could not encode page image."));
            return;
          }
          blob.arrayBuffer().then(
            function (ab) {
              resolve(new Uint8Array(ab));
            },
            reject
          );
        },
        "image/jpeg",
        quality || 0.92
      );
    });
  }

  /**
   * Render selected PDF pages to JPEG bytes (full visual snapshot: graphics + text).
   * @returns {Promise<Array<{bytes: Uint8Array, page: number}>>}
   */
  async function renderPdfPagesToImages(arrayBuffer, pageNums, onProgress) {
    var pdfjsLib = await ensurePdfJs();
    var pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
    var total = pdf.numPages;
    var pages = pageNums.filter(function (n) {
      return n >= 1 && n <= total;
    });
    if (!pages.length) {
      throw new Error("No valid PDF pages. This file has " + total + " page(s).");
    }

    var out = [];
    for (var i = 0; i < pages.length; i++) {
      var pageNum = pages[i];
      if (onProgress) onProgress(i + 1, pages.length, pageNum);
      var page = await pdf.getPage(pageNum);
      var baseVp = page.getViewport({ scale: 1 });
      // ~1920px wide for sharp text; clamp for memory
      var scale = 1920 / Math.max(baseVp.width, 1);
      scale = Math.min(Math.max(scale, 1.5), 2.75);
      var viewport = page.getViewport({ scale: scale });
      var canvas = document.createElement("canvas");
      canvas.width = Math.floor(viewport.width);
      canvas.height = Math.floor(viewport.height);
      var ctx = canvas.getContext("2d", { alpha: false });
      ctx.fillStyle = "#ffffff";
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      await page
        .render({
          canvasContext: ctx,
          viewport: viewport,
          intent: "display",
        })
        .promise;
      var bytes = await canvasToJpegBytes(canvas, 0.92);
      // free canvas memory
      canvas.width = 0;
      canvas.height = 0;
      out.push({ bytes: bytes, page: pageNum });
      await yieldToUI();
    }
    return out;
  }

  function makeFullImageSlideXml(cx, cy) {
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
      'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
      "<p:cSld><p:spTree>" +
      "<p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>" +
      "<p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/>" +
      "<a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></p:grpSpPr>" +
      "<p:pic>" +
      "<p:nvPicPr><p:cNvPr id=\"2\" name=\"PDF Page\"/>" +
      "<p:cNvPicPr><a:picLocks noChangeAspect=\"0\"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>" +
      '<p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
      "<p:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"" +
      cx +
      '" cy="' +
      cy +
      '"/></a:xfrm>' +
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>' +
      "</p:pic></p:spTree></p:cSld>" +
      "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>"
    );
  }

  function makeImageSlideRels(mediaFileName) {
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/' +
      mediaFileName +
      '"/>' +
      "</Relationships>"
    );
  }

  /**
   * Append full-slide screenshot images into blank/target PPTX.
   * @param {File|ArrayBuffer} tgtFile
   * @param {Array<{bytes: Uint8Array}>} pageImages
   * @param {number} insertAfter
   */
  async function mergePdfImagesIntoTarget(tgtFile, pageImages, insertAfter) {
    if (!pageImages || !pageImages.length) {
      throw new Error("No PDF pages rendered.");
    }
    var tgtData =
      tgtFile instanceof ArrayBuffer
        ? tgtFile
        : await readFileAsArrayBuffer(tgtFile);
    var tgtZip = await JSZip.loadAsync(tgtData, { checkCRC32: false });
    var tgtCount = slideNames(tgtZip).length;
    if (insertAfter > tgtCount) insertAfter = tgtCount;

    var prsXml = await tgtZip.file("ppt/presentation.xml").async("string");
    var cx = 12192000;
    var cy = 6858000;
    var sz = prsXml.match(/<p:sldSz\b[^/]*\/>/);
    if (sz) {
      var cxM = /cx="(\d+)"/.exec(sz[0]);
      var cyM = /cy="(\d+)"/.exec(sz[0]);
      if (cxM) cx = parseInt(cxM[1], 10);
      if (cyM) cy = parseInt(cyM[1], 10);
    }

    var addedMeta = [];
    var nextSlideNum = tgtCount + 1;
    for (var i = 0; i < pageImages.length; i++) {
      var newNum = nextSlideNum++;
      var mediaName = "pdf_page_" + newNum + ".jpg";
      var mediaPath = "ppt/media/" + mediaName;
      var c = 1;
      while (tgtZip.file(mediaPath)) {
        mediaName = "pdf_page_" + newNum + "_" + c + ".jpg";
        mediaPath = "ppt/media/" + mediaName;
        c++;
      }
      tgtZip.file(mediaPath, pageImages[i].bytes);
      tgtZip.file(
        "ppt/slides/slide" + newNum + ".xml",
        makeFullImageSlideXml(cx, cy)
      );
      tgtZip.file(
        "ppt/slides/_rels/slide" + newNum + ".xml.rels",
        makeImageSlideRels(mediaName)
      );
      addedMeta.push({ newNum: newNum });
    }

    // presentation.xml.rels
    var prsRelsXml = await tgtZip
      .file("ppt/_rels/presentation.xml.rels")
      .async("string");
    var maxRid = 0;
    prsRelsXml.replace(/Id="rId(\d+)"/g, function (_, n) {
      maxRid = Math.max(maxRid, parseInt(n, 10));
      return _;
    });
    var relChunks = [];
    addedMeta.forEach(function (meta, idx) {
      meta.rid = "rId" + (maxRid + idx + 1);
      relChunks.push(
        '<Relationship Id="' +
          meta.rid +
          '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide' +
          meta.newNum +
          '.xml"/>'
      );
    });
    prsRelsXml = prsRelsXml.replace(
      "</Relationships>",
      relChunks.join("") + "</Relationships>"
    );
    tgtZip.file("ppt/_rels/presentation.xml.rels", prsRelsXml);

    // sldIdLst
    var maxId = 255;
    prsXml.replace(/<p:sldId\b[^>]*\bid="(\d+)"/g, function (_, n) {
      maxId = Math.max(maxId, parseInt(n, 10));
      return _;
    });
    var newSldTags = addedMeta.map(function (meta, idx) {
      meta.id = maxId + idx + 1;
      return '<p:sldId id="' + meta.id + '" r:id="' + meta.rid + '"/>';
    });
    var existingTags = [];
    prsXml.replace(/<p:sldId\b[^/]*\/>/g, function (tag) {
      existingTags.push(tag);
      return tag;
    });
    var ordered;
    if (insertAfter < tgtCount && existingTags.length) {
      ordered = existingTags
        .slice(0, insertAfter)
        .concat(newSldTags)
        .concat(existingTags.slice(insertAfter));
    } else {
      ordered = existingTags.concat(newSldTags);
    }
    var newList = "<p:sldIdLst>" + ordered.join("") + "</p:sldIdLst>";
    if (/<p:sldIdLst\s*\/>/.test(prsXml)) {
      prsXml = prsXml.replace(/<p:sldIdLst\s*\/>/, newList);
    } else if (/<p:sldIdLst[\s>][\s\S]*?<\/p:sldIdLst>/.test(prsXml)) {
      prsXml = prsXml.replace(/<p:sldIdLst[\s>][\s\S]*?<\/p:sldIdLst>/, newList);
    } else if (/<\/p:sldMasterIdLst>/.test(prsXml)) {
      prsXml = prsXml.replace(
        /<\/p:sldMasterIdLst>/,
        "</p:sldMasterIdLst>" + newList
      );
    } else {
      throw new Error("Could not insert sldIdLst into presentation.xml");
    }
    if (prsXml.indexOf("<?xml") !== 0) {
      prsXml =
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + prsXml;
    }
    tgtZip.file("ppt/presentation.xml", prsXml);

    // Content_Types
    var ctXml = await tgtZip.file("[Content_Types].xml").async("string");
    if (!/Extension="jpeg"/i.test(ctXml) && !/Extension="jpg"/i.test(ctXml)) {
      ctXml = ctXml.replace(
        "</Types>",
        '<Default Extension="jpeg" ContentType="image/jpeg"/></Types>'
      );
    }
    if (!/Extension="jpg"/i.test(ctXml)) {
      ctXml = ctXml.replace(
        "</Types>",
        '<Default Extension="jpg" ContentType="image/jpeg"/></Types>'
      );
    }
    var ctAdds = [];
    addedMeta.forEach(function (meta) {
      var part = "/ppt/slides/slide" + meta.newNum + ".xml";
      if (ctXml.indexOf('PartName="' + part + '"') < 0) {
        ctAdds.push(
          '<Override PartName="' +
            part +
            '" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        );
      }
    });
    if (ctAdds.length) {
      ctXml = ctXml.replace("</Types>", ctAdds.join("") + "</Types>");
    }
    tgtZip.file("[Content_Types].xml", ctXml);

    var ab = await tgtZip.generateAsync({
      type: "arraybuffer",
      compression: "DEFLATE",
      compressionOptions: { level: 6 },
    });
    if (!ab || ab.byteLength < 2000) {
      throw new Error("Output file was empty — PDF merge failed.");
    }
    return {
      blob: new Blob([ab], {
        type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      }),
      count: addedMeta.length,
    };
  }

  function wireDrop(zone, input, onFiles, opts) {
    opts = opts || {};
    function pick() {
      input.click();
    }
    // Merge zones register their own click (File System Access API)
    if (!opts.skipClick) {
      zone.addEventListener("click", function (e) {
        if (e.target === input) return;
        pick();
      });
    }
    input.addEventListener("change", function () {
      if (input.files && input.files.length) onFiles(input.files);
    });
    ["dragenter", "dragover"].forEach(function (ev) {
      zone.addEventListener(ev, function (e) {
        e.preventDefault();
        zone.classList.add("drag");
      });
    });
    ["dragleave", "drop"].forEach(function (ev) {
      zone.addEventListener(ev, function (e) {
        e.preventDefault();
        if (ev === "dragleave") zone.classList.remove("drag");
      });
    });
    zone.addEventListener("drop", function (e) {
      zone.classList.remove("drag");
      if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files.length) {
        onFiles(e.dataTransfer.files);
      }
    });
  }

  // ---- Tabs ----
  document.querySelectorAll(".tab").forEach(function (btn) {
    btn.addEventListener("click", function () {
      document.querySelectorAll(".tab").forEach(function (b) {
        b.classList.remove("active");
      });
      document.querySelectorAll(".panel").forEach(function (p) {
        p.classList.remove("active");
      });
      btn.classList.add("active");
      $("panel-" + btn.dataset.tab).classList.add("active");
    });
  });

  // ---- OCR ----
  var ocrText = $("ocrText");
  var ocrStatus = $("ocrStatus");
  var ocrFile = $("ocrFile");
  var ocrPanel = $("panel-ocr");
  var ocrPasteHint = $("ocrPasteHint");
  var ocrBusy = false;

  function isOcrPanelActive() {
    return !!(ocrPanel && ocrPanel.classList.contains("active"));
  }

  function ensureOcrTabActive() {
    var tab = document.querySelector('.tab[data-tab="ocr"]');
    if (tab && !isOcrPanelActive()) tab.click();
  }

  /**
   * Pull an image blob from a paste/drop event.
   * Win+Shift+S / Snipping Tool often use files[], empty type, or HTML image.
   */
  function imageBlobFromClipboardData(cd) {
    if (!cd) return null;
    var i, it, f, type;

    // 1) DataTransferItemList — preferred
    if (cd.items && cd.items.length) {
      for (i = 0; i < cd.items.length; i++) {
        it = cd.items[i];
        type = (it.type || "").toLowerCase();
        if (it.kind === "file" && type.indexOf("image/") === 0) {
          f = it.getAsFile();
          if (f) return f;
        }
      }
      // Screenshot with empty MIME (some Edge/Chrome builds)
      for (i = 0; i < cd.items.length; i++) {
        it = cd.items[i];
        if (it.kind === "file") {
          f = it.getAsFile();
          if (f && f.size > 0 && (!f.type || f.type.indexOf("image/") === 0 || f.type === "application/octet-stream")) {
            // Prefer if name looks like image or type empty (common for screen capture)
            if (!f.type || f.type.indexOf("image/") === 0 || /\.(png|jpe?g|gif|webp|bmp)$/i.test(f.name || "")) {
              return f;
            }
            // Last resort: unnamed clipboard file with size (screen grab)
            if (!f.type && (!f.name || f.name === "image.png" || f.name === "blob")) {
              return f;
            }
          }
        }
      }
    }

    // 2) files list (Snipping Tool / Explorer copy)
    if (cd.files && cd.files.length) {
      for (i = 0; i < cd.files.length; i++) {
        f = cd.files[i];
        if (!f) continue;
        if (!f.type || f.type.indexOf("image/") === 0) return f;
        if (/\.(png|jpe?g|gif|webp|bmp)$/i.test(f.name || "")) return f;
      }
    }

    // 3) HTML clipboard with <img src="data:image/...">
    try {
      var html = cd.getData && cd.getData("text/html");
      if (html && /data:image\//i.test(html)) {
        var m = html.match(/src\s*=\s*["'](data:image\/[a-zA-Z0-9.+-]+;base64,[^"']+)["']/i);
        if (m && m[1]) {
          return dataUrlToBlob(m[1]);
        }
      }
    } catch (e) {}

    return null;
  }

  function dataUrlToBlob(dataUrl) {
    try {
      var parts = String(dataUrl).split(",");
      var meta = parts[0] || "";
      var b64 = parts[1] || "";
      var mime = (meta.match(/data:([^;]+)/) || [])[1] || "image/png";
      var bin = atob(b64);
      var arr = new Uint8Array(bin.length);
      for (var i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
      return new Blob([arr], { type: mime });
    } catch (e) {
      return null;
    }
  }

  /** Clipboard API (async) — works when paste event has no image items. */
  async function imageBlobFromClipboardApi() {
    if (!navigator.clipboard || !navigator.clipboard.read) return null;
    try {
      var items = await navigator.clipboard.read();
      for (var i = 0; i < items.length; i++) {
        var item = items[i];
        var types = item.types || [];
        for (var t = 0; t < types.length; t++) {
          if (String(types[t]).indexOf("image/") === 0) {
            return await item.getType(types[t]);
          }
        }
      }
    } catch (e) {
      // NotAllowedError / insecure context — caller shows message
      return null;
    }
    return null;
  }

  async function runOcrOnBlob(blob) {
    if (!blob) {
      setStatus(ocrStatus, "No image to OCR.", "err");
      return;
    }
    if (ocrBusy) {
      setStatus(ocrStatus, "OCR already running…", "err");
      return;
    }
    ocrBusy = true;
    ensureOcrTabActive();
    setStatus(ocrStatus, "Loading OCR engine…");
    try {
      await ensureTesseract();
    } catch (e) {
      ocrBusy = false;
      setStatus(ocrStatus, e.message || "OCR engine failed to load.", "err");
      return;
    }
    if (typeof Tesseract === "undefined") {
      ocrBusy = false;
      setStatus(ocrStatus, "Tesseract.js failed to load (need internet once for CDN).", "err");
      return;
    }
    setStatus(ocrStatus, "Running OCR…");
    try {
      var result = await Tesseract.recognize(blob, "eng", {
        logger: function (m) {
          if (m.status === "recognizing text" && m.progress != null) {
            setStatus(ocrStatus, "OCR " + Math.round(m.progress * 100) + "%…");
          }
        },
      });
      var text = (result && result.data && result.data.text ? result.data.text : "").trim();
      if (ocrText) ocrText.value = text || "[No text detected]";
      setStatus(ocrStatus, text ? "Done — you can Copy." : "No text found.", text ? "ok" : "err");
    } catch (e) {
      console.error(e);
      setStatus(ocrStatus, "OCR error: " + (e.message || e), "err");
    }
    ocrBusy = false;
  }

  if (ocrFile) {
    ocrFile.addEventListener("change", function () {
      var f = ocrFile.files && ocrFile.files[0];
      if (f) runOcrOnBlob(f);
      // allow re-choosing the same file
      try {
        ocrFile.value = "";
      } catch (e) {}
    });
  }

  async function handleOcrPasteEvent(e) {
    if (!isOcrPanelActive()) return false;
    var blob = imageBlobFromClipboardData(e.clipboardData || e.originalEvent && e.originalEvent.clipboardData);
    if (blob) {
      e.preventDefault();
      e.stopPropagation();
      runOcrOnBlob(blob);
      return true;
    }
    return false;
  }

  // Capture phase so we win over textarea default paste when an image is present
  document.addEventListener(
    "paste",
    function (e) {
      if (handleWhitebgPasteEvent(e)) return;
      handleOcrPasteEvent(e);
    },
    true
  );
  window.addEventListener(
    "paste",
    function (e) {
      if (handleWhitebgPasteEvent(e)) return;
      handleOcrPasteEvent(e);
    },
    true
  );

  // "Paste image" button — Clipboard API or guided Ctrl+V
  if (ocrPasteHint) {
    ocrPasteHint.addEventListener("click", async function () {
      ensureOcrTabActive();
      setStatus(ocrStatus, "Reading clipboard…");
      var blob = await imageBlobFromClipboardApi();
      if (blob) {
        runOcrOnBlob(blob);
        return;
      }
      // Focus a target so the next Ctrl+V is received inside this frame
      if (ocrText) {
        try {
          ocrText.focus();
        } catch (e) {}
      }
      setStatus(
        ocrStatus,
        "No image in clipboard API. Click here, then press Ctrl+V (Win+Shift+S first to capture).",
        "err"
      );
    });
  }

  // Drag & drop image onto Text Extractor panel
  if (ocrPanel) {
    ["dragenter", "dragover"].forEach(function (ev) {
      ocrPanel.addEventListener(ev, function (e) {
        if (!isOcrPanelActive()) return;
        e.preventDefault();
        e.stopPropagation();
      });
    });
    ocrPanel.addEventListener("drop", function (e) {
      if (!isOcrPanelActive()) return;
      e.preventDefault();
      e.stopPropagation();
      var blob = imageBlobFromClipboardData(e.dataTransfer);
      if (blob) {
        runOcrOnBlob(blob);
      } else {
        setStatus(ocrStatus, "Drop a PNG/JPG image file.", "err");
      }
    });
  }

  var ocrCopyBtn = $("ocrCopy");
  if (ocrCopyBtn) {
    ocrCopyBtn.addEventListener("click", async function () {
      try {
        await navigator.clipboard.writeText((ocrText && ocrText.value) || "");
        setStatus(ocrStatus, "Copied.", "ok");
      } catch (e) {
        if (ocrText) ocrText.select();
        document.execCommand("copy");
        setStatus(ocrStatus, "Copied.", "ok");
      }
    });
  }
  var ocrClearBtn = $("ocrClear");
  if (ocrClearBtn) {
    ocrClearBtn.addEventListener("click", function () {
      if (ocrText) ocrText.value = "";
      setStatus(ocrStatus, "Ready");
    });
  }

  // ---- PNG White Background ----
  var whitebgPanel = $("panel-whitebg");
  var whitebgStatus = $("whitebgStatus");
  var whitebgDrop = $("whitebgDrop");
  var whitebgLabel = $("whitebgLabel");
  var whitebgDropTitle = $("whitebgDropTitle");
  var whitebgFile = $("whitebgFile");
  var whitebgDownload = $("whitebgDownload");
  var whitebgClear = $("whitebgClear");
  var whitebgClip = $("whitebgClip");
  var whitebgActions = $("whitebgActions");
  var whitebgPreviews = $("whitebgPreviews");
  var whitebgOrig = $("whitebgOrig");
  var whitebgResult = $("whitebgResult");
  var whitebgCropStage = $("whitebgCropStage");
  var whitebgCropWrap = $("whitebgCropWrap");
  var whitebgCropImg = $("whitebgCropImg");
  var whitebgCropBox = $("whitebgCropBox");
  var whitebgCropApply = $("whitebgCropApply");
  var whitebgCropCancel = $("whitebgCropCancel");
  var whitebgSourceBlob = null;
  var whitebgSourceName = "image.png";
  var whitebgResultBlob = null;
  var whitebgBusy = false;
  var WHITEBG_HINT = "Click / tap · drag & drop · Ctrl+V / paste";
  // Crop state in natural image pixels: {x, y, w, h, imgW, imgH}
  var whitebgCrop = null;
  var whitebgCropDrag = null;

  function isWhitebgPanelActive() {
    return !!(whitebgPanel && whitebgPanel.classList.contains("active"));
  }

  function ensureWhitebgTabActive() {
    var tab = document.querySelector('.tab[data-tab="whitebg"]');
    if (tab && !isWhitebgPanelActive()) tab.click();
  }

  function revokeImgUrl(img) {
    if (!img) return;
    var u = img.getAttribute("src");
    if (u && u.indexOf("blob:") === 0) {
      try {
        URL.revokeObjectURL(u);
      } catch (e) {}
    }
  }

  function showWhitebgStatus(msg, kind) {
    if (!whitebgStatus) return;
    if (!msg) {
      whitebgStatus.hidden = true;
      whitebgStatus.textContent = "";
      whitebgStatus.className = "status";
      return;
    }
    whitebgStatus.hidden = false;
    setStatus(whitebgStatus, msg, kind);
  }

  function exitWhitebgCropMode() {
    whitebgCropDrag = null;
    if (whitebgCropStage) whitebgCropStage.hidden = true;
    revokeImgUrl(whitebgCropImg);
    if (whitebgCropImg) whitebgCropImg.removeAttribute("src");
    if (whitebgPreviews) whitebgPreviews.hidden = !whitebgSourceBlob;
    if (whitebgActions) whitebgActions.hidden = !whitebgSourceBlob;
  }

  function clearWhitebg() {
    exitWhitebgCropMode();
    whitebgSourceBlob = null;
    whitebgResultBlob = null;
    whitebgSourceName = "image.png";
    whitebgCrop = null;
    if (whitebgDropTitle) whitebgDropTitle.textContent = "Add PNG";
    if (whitebgLabel) whitebgLabel.textContent = WHITEBG_HINT;
    if (whitebgDownload) whitebgDownload.disabled = true;
    if (whitebgActions) whitebgActions.hidden = true;
    if (whitebgPreviews) whitebgPreviews.hidden = true;
    revokeImgUrl(whitebgOrig);
    revokeImgUrl(whitebgResult);
    revokeImgUrl(whitebgCropImg);
    if (whitebgOrig) whitebgOrig.removeAttribute("src");
    if (whitebgResult) whitebgResult.removeAttribute("src");
    if (whitebgCropImg) whitebgCropImg.removeAttribute("src");
    showWhitebgStatus("");
  }

  function whitebgOutName(name) {
    var baseName = (name || "image.png").replace(/\.[^.]+$/, "");
    if (!baseName) baseName = "image";
    return baseName + "_whitebg.png";
  }

  /** True if pixel looks like a transparent / checker / plain light backdrop. */
  function isBackdropPixel(r, g, b, a) {
    if (a < 40) return true;
    var maxc = Math.max(r, g, b);
    var minc = Math.min(r, g, b);
    var chroma = maxc - minc;
    var lum = (r + g + b) / 3;
    // Light gray / white / typical PNG checkerboard squares (incl. ~231 gray)
    if (chroma <= 22 && lum >= 175) return true;
    // Near-white with slight tint (export artifacts)
    if (chroma <= 30 && lum >= 235) return true;
    // Semi-transparent light pixels
    if (a < 230 && chroma <= 28 && lum >= 190) return true;
    return false;
  }

  /**
   * Mark backdrop pixels connected to the image border (4-connected).
   * Handles real transparency and baked-in checkerboards without eating
   * interior whites (eyes/teeth) surrounded by the subject.
   */
  function markBorderConnectedBackdrop(data, w, h) {
    var d = data.data;
    var n = w * h;
    var bg = new Uint8Array(n); // 1 = backdrop
    var queue = new Int32Array(n);
    var qh = 0;
    var qt = 0;

    function tryPush(x, y) {
      if (x < 0 || y < 0 || x >= w || y >= h) return;
      var i = y * w + x;
      if (bg[i]) return;
      var o = i * 4;
      if (!isBackdropPixel(d[o], d[o + 1], d[o + 2], d[o + 3])) return;
      bg[i] = 1;
      queue[qt++] = i;
    }

    var x, y;
    for (x = 0; x < w; x++) {
      tryPush(x, 0);
      tryPush(x, h - 1);
    }
    for (y = 0; y < h; y++) {
      tryPush(0, y);
      tryPush(w - 1, y);
    }

    while (qh < qt) {
      var i = queue[qh++];
      var cx = i % w;
      var cy = (i / w) | 0;
      tryPush(cx - 1, cy);
      tryPush(cx + 1, cy);
      tryPush(cx, cy - 1);
      tryPush(cx, cy + 1);
    }
    return bg;
  }

  /**
   * Process image:
   * - If transparent / checkerboard backdrop detected → solid white BG
   * - Otherwise leave colors as-is (still force full opacity)
   * Returns Promise<Blob> PNG.
   */
  function convertToWhiteBackground(blob) {
    return new Promise(function (resolve, reject) {
      if (!blob) {
        reject(new Error("No image."));
        return;
      }
      var url = URL.createObjectURL(blob);
      var img = new Image();
      img.onload = function () {
        try {
          var w = img.naturalWidth || img.width;
          var h = img.naturalHeight || img.height;
          if (!w || !h) {
            URL.revokeObjectURL(url);
            reject(new Error("Could not read image size."));
            return;
          }
          var canvas = document.createElement("canvas");
          canvas.width = w;
          canvas.height = h;
          var ctx = canvas.getContext("2d");
          ctx.clearRect(0, 0, w, h);
          ctx.drawImage(img, 0, 0);
          URL.revokeObjectURL(url);

          var imageData = ctx.getImageData(0, 0, w, h);
          var d = imageData.data;
          var bg = markBorderConnectedBackdrop(imageData, w, h);
          var i, o, a, r, g, b, af;
          var transparentCount = 0;
          var bgCount = 0;
          var total = w * h;

          for (i = 0; i < total; i++) {
            o = i * 4;
            if (d[o + 3] < 250) transparentCount++;
            if (bg[i]) bgCount++;
          }

          // Only change background when image looks transparent / checkerboard
          var needsWhiteBg =
            transparentCount > total * 0.005 || bgCount > total * 0.02;

          if (needsWhiteBg) {
            for (i = 0; i < total; i++) {
              o = i * 4;
              r = d[o];
              g = d[o + 1];
              b = d[o + 2];
              a = d[o + 3];
              if (bg[i] || a < 40) {
                d[o] = 255;
                d[o + 1] = 255;
                d[o + 2] = 255;
                d[o + 3] = 255;
              } else if (a < 255) {
                af = a / 255;
                d[o] = Math.round(r * af + 255 * (1 - af));
                d[o + 1] = Math.round(g * af + 255 * (1 - af));
                d[o + 2] = Math.round(b * af + 255 * (1 - af));
                d[o + 3] = 255;
              } else {
                d[o + 3] = 255;
              }
            }
            ctx.putImageData(imageData, 0, 0);
          } else {
            // No transparent backdrop — leave pixels as-is, ensure opaque PNG
            for (i = 0; i < total; i++) {
              d[i * 4 + 3] = 255;
            }
            ctx.putImageData(imageData, 0, 0);
          }
          canvasToPngBlob(canvas).then(resolve, reject);
        } catch (err) {
          try {
            URL.revokeObjectURL(url);
          } catch (e2) {}
          reject(err);
        }
      };
      img.onerror = function () {
        URL.revokeObjectURL(url);
        reject(new Error("Could not load image."));
      };
      img.src = url;
    });
  }

  function canvasToPngBlob(canvas) {
    return new Promise(function (resolve, reject) {
      if (canvas.toBlob) {
        canvas.toBlob(
          function (out) {
            if (out) resolve(out);
            else reject(new Error("Could not export PNG."));
          },
          "image/png"
        );
      } else {
        try {
          var dataUrl = canvas.toDataURL("image/png");
          var bin = atob(dataUrl.split(",")[1]);
          var arr = new Uint8Array(bin.length);
          for (var i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
          resolve(new Blob([arr], { type: "image/png" }));
        } catch (e) {
          reject(e);
        }
      }
    });
  }

  function loadImageFromBlob(blob) {
    return new Promise(function (resolve, reject) {
      var url = URL.createObjectURL(blob);
      var img = new Image();
      img.onload = function () {
        URL.revokeObjectURL(url);
        resolve(img);
      };
      img.onerror = function () {
        URL.revokeObjectURL(url);
        reject(new Error("Could not load image for resize."));
      };
      img.src = url;
    });
  }

  /**
   * If blob is over maxBytes, scale dimensions down until PNG is ≤ maxBytes.
   * Leaves smaller files unchanged.
   */
  async function shrinkPngToMaxBytes(blob, maxBytes) {
    if (!blob || blob.size <= maxBytes) return blob;

    var img = await loadImageFromBlob(blob);
    var srcW = img.naturalWidth || img.width;
    var srcH = img.naturalHeight || img.height;
    if (!srcW || !srcH) return blob;

    var lo = 0.15;
    var hi = 1;
    var best = null;
    var attempt;

    // Binary search scale factor (PNG size mainly tracks pixel count)
    for (attempt = 0; attempt < 12; attempt++) {
      var scale = (lo + hi) / 2;
      var w = Math.max(1, Math.round(srcW * scale));
      var h = Math.max(1, Math.round(srcH * scale));
      var canvas = document.createElement("canvas");
      canvas.width = w;
      canvas.height = h;
      var ctx = canvas.getContext("2d");
      ctx.fillStyle = "#ffffff";
      ctx.fillRect(0, 0, w, h);
      ctx.imageSmoothingEnabled = true;
      ctx.imageSmoothingQuality = "high";
      ctx.drawImage(img, 0, 0, w, h);
      var candidate = await canvasToPngBlob(canvas);
      if (candidate.size <= maxBytes) {
        best = candidate;
        lo = scale; // try larger (better quality) while staying under cap
      } else {
        hi = scale;
      }
    }

    if (best) return best;

    // Last resort: very small scale
    var w2 = Math.max(1, Math.round(srcW * 0.12));
    var h2 = Math.max(1, Math.round(srcH * 0.12));
    var c2 = document.createElement("canvas");
    c2.width = w2;
    c2.height = h2;
    var ctx2 = c2.getContext("2d");
    ctx2.fillStyle = "#ffffff";
    ctx2.fillRect(0, 0, w2, h2);
    ctx2.drawImage(img, 0, 0, w2, h2);
    return canvasToPngBlob(c2);
  }

  async function runWhitebgConvert() {
    if (!whitebgSourceBlob || whitebgBusy) return;
    whitebgBusy = true;
    if (whitebgDownload) whitebgDownload.disabled = true;
    try {
      var out = await convertToWhiteBackground(whitebgSourceBlob);
      if (out.size > WHITEBG_MAX_BYTES) {
        out = await shrinkPngToMaxBytes(out, WHITEBG_MAX_BYTES);
      }
      whitebgResultBlob = out;
      revokeImgUrl(whitebgResult);
      if (whitebgResult) whitebgResult.src = URL.createObjectURL(out);
      if (whitebgPreviews) whitebgPreviews.hidden = false;
      if (whitebgActions) whitebgActions.hidden = false;
      if (whitebgDownload) whitebgDownload.disabled = false;
      showWhitebgStatus("");
    } catch (e) {
      console.error(e);
      showWhitebgStatus("Error: " + (e.message || e), "err");
    }
    whitebgBusy = false;
  }

  function getCropDisplayMetrics() {
    if (!whitebgCropImg || !whitebgCrop) return null;
    var rect = whitebgCropImg.getBoundingClientRect();
    var natW = whitebgCrop.imgW;
    var natH = whitebgCrop.imgH;
    if (!rect.width || !rect.height || !natW || !natH) return null;
    return {
      left: rect.left,
      top: rect.top,
      width: rect.width,
      height: rect.height,
      scaleX: rect.width / natW,
      scaleY: rect.height / natH,
    };
  }

  function layoutWhitebgCropBox() {
    if (!whitebgCropBox || !whitebgCrop) return;
    var m = getCropDisplayMetrics();
    if (!m) return;
    whitebgCropBox.style.left = whitebgCrop.x * m.scaleX + "px";
    whitebgCropBox.style.top = whitebgCrop.y * m.scaleY + "px";
    whitebgCropBox.style.width = whitebgCrop.w * m.scaleX + "px";
    whitebgCropBox.style.height = whitebgCrop.h * m.scaleY + "px";
  }

  function clampWhitebgCrop() {
    if (!whitebgCrop) return;
    var minSize = 20;
    var imgW = whitebgCrop.imgW;
    var imgH = whitebgCrop.imgH;
    whitebgCrop.w = Math.max(minSize, Math.min(whitebgCrop.w, imgW));
    whitebgCrop.h = Math.max(minSize, Math.min(whitebgCrop.h, imgH));
    whitebgCrop.x = Math.max(0, Math.min(whitebgCrop.x, imgW - whitebgCrop.w));
    whitebgCrop.y = Math.max(0, Math.min(whitebgCrop.y, imgH - whitebgCrop.h));
  }

  function openWhitebgCropMode() {
    if (!whitebgSourceBlob || !whitebgCropImg || !whitebgCropStage) return;
    exitWhitebgCropMode();
    if (whitebgPreviews) whitebgPreviews.hidden = true;
    if (whitebgActions) whitebgActions.hidden = true;
    whitebgCropStage.hidden = false;

    revokeImgUrl(whitebgCropImg);
    var url = URL.createObjectURL(whitebgSourceBlob);
    whitebgCropImg.onload = function () {
      var imgW = whitebgCropImg.naturalWidth || whitebgCropImg.width;
      var imgH = whitebgCropImg.naturalHeight || whitebgCropImg.height;
      // Default selection: centered 80% box
      var w = Math.round(imgW * 0.8);
      var h = Math.round(imgH * 0.8);
      whitebgCrop = {
        x: Math.round((imgW - w) / 2),
        y: Math.round((imgH - h) / 2),
        w: w,
        h: h,
        imgW: imgW,
        imgH: imgH,
      };
      // next frame after layout
      requestAnimationFrame(function () {
        layoutWhitebgCropBox();
      });
    };
    whitebgCropImg.src = url;
  }

  async function applyWhitebgCrop() {
    if (!whitebgSourceBlob || !whitebgCrop) return;
    try {
      var img = await loadImageFromBlob(whitebgSourceBlob);
      var sx = Math.round(whitebgCrop.x);
      var sy = Math.round(whitebgCrop.y);
      var sw = Math.round(whitebgCrop.w);
      var sh = Math.round(whitebgCrop.h);
      sx = Math.max(0, Math.min(sx, (img.naturalWidth || img.width) - 1));
      sy = Math.max(0, Math.min(sy, (img.naturalHeight || img.height) - 1));
      sw = Math.max(1, Math.min(sw, (img.naturalWidth || img.width) - sx));
      sh = Math.max(1, Math.min(sh, (img.naturalHeight || img.height) - sy));

      var canvas = document.createElement("canvas");
      canvas.width = sw;
      canvas.height = sh;
      var ctx = canvas.getContext("2d");
      ctx.clearRect(0, 0, sw, sh);
      ctx.drawImage(img, sx, sy, sw, sh, 0, 0, sw, sh);
      var cropped = await canvasToPngBlob(canvas);
      exitWhitebgCropMode();
      // Treat cropped image as new source
      whitebgSourceBlob = cropped;
      whitebgResultBlob = null;
      revokeImgUrl(whitebgOrig);
      if (whitebgOrig) whitebgOrig.src = URL.createObjectURL(cropped);
      if (whitebgLabel) whitebgLabel.textContent = whitebgSourceName + " (clipped)";
      await runWhitebgConvert();
    } catch (e) {
      console.error(e);
      showWhitebgStatus("Clip error: " + (e.message || e), "err");
    }
  }

  function whitebgPointerPos(e) {
    if (e.touches && e.touches.length) {
      return { x: e.touches[0].clientX, y: e.touches[0].clientY };
    }
    if (e.changedTouches && e.changedTouches.length) {
      return { x: e.changedTouches[0].clientX, y: e.changedTouches[0].clientY };
    }
    return { x: e.clientX, y: e.clientY };
  }

  function startWhitebgCropDrag(e, handle) {
    if (!whitebgCrop) return;
    e.preventDefault();
    e.stopPropagation();
    var pos = whitebgPointerPos(e);
    whitebgCropDrag = {
      handle: handle || "move",
      startX: pos.x,
      startY: pos.y,
      orig: {
        x: whitebgCrop.x,
        y: whitebgCrop.y,
        w: whitebgCrop.w,
        h: whitebgCrop.h,
      },
    };
  }

  function moveWhitebgCropDrag(e) {
    if (!whitebgCropDrag || !whitebgCrop) return;
    e.preventDefault();
    var m = getCropDisplayMetrics();
    if (!m) return;
    var pos = whitebgPointerPos(e);
    var dx = (pos.x - whitebgCropDrag.startX) / m.scaleX;
    var dy = (pos.y - whitebgCropDrag.startY) / m.scaleY;
    var o = whitebgCropDrag.orig;
    var h = whitebgCropDrag.handle;
    var minSize = 20;

    var x = o.x;
    var y = o.y;
    var w = o.w;
    var hh = o.h;

    if (h === "move") {
      x = o.x + dx;
      y = o.y + dy;
    } else {
      if (h.indexOf("e") >= 0) w = o.w + dx;
      if (h.indexOf("s") >= 0) hh = o.h + dy;
      if (h.indexOf("w") >= 0) {
        x = o.x + dx;
        w = o.w - dx;
      }
      if (h.indexOf("n") >= 0) {
        y = o.y + dy;
        hh = o.h - dy;
      }
    }

    // Normalize negative sizes
    if (w < minSize) {
      if (h.indexOf("w") >= 0) x = o.x + o.w - minSize;
      w = minSize;
    }
    if (hh < minSize) {
      if (h.indexOf("n") >= 0) y = o.y + o.h - minSize;
      hh = minSize;
    }

    whitebgCrop.x = x;
    whitebgCrop.y = y;
    whitebgCrop.w = w;
    whitebgCrop.h = hh;
    clampWhitebgCrop();
    layoutWhitebgCropBox();
  }

  function endWhitebgCropDrag() {
    whitebgCropDrag = null;
  }

  async function loadWhitebgBlob(blob, name) {
    if (!blob) return;
    var type = (blob.type || "").toLowerCase();
    var nm = name || blob.name || "image.png";
    if (type && type.indexOf("image/") !== 0 && type !== "application/octet-stream") {
      showWhitebgStatus("Use a PNG or image file.", "err");
      return;
    }
    exitWhitebgCropMode();
    whitebgSourceBlob = blob;
    whitebgSourceName = nm;
    whitebgResultBlob = null;
    whitebgCrop = null;
    if (whitebgDropTitle) whitebgDropTitle.textContent = "Add PNG";
    if (whitebgLabel) whitebgLabel.textContent = nm;
    if (whitebgDownload) whitebgDownload.disabled = true;
    revokeImgUrl(whitebgOrig);
    revokeImgUrl(whitebgResult);
    if (whitebgOrig) whitebgOrig.src = URL.createObjectURL(blob);
    if (whitebgResult) whitebgResult.removeAttribute("src");
    if (whitebgPreviews) whitebgPreviews.hidden = false;
    if (whitebgActions) whitebgActions.hidden = false;
    showWhitebgStatus("");
    await runWhitebgConvert();
  }

  function handleWhitebgPasteEvent(e) {
    if (!isWhitebgPanelActive()) return false;
    // Don't steal paste while typing isn't relevant; always ok on this panel
    var blob = imageBlobFromClipboardData(
      e.clipboardData || (e.originalEvent && e.originalEvent.clipboardData)
    );
    if (blob) {
      e.preventDefault();
      e.stopPropagation();
      loadWhitebgBlob(blob, "pasted.png");
      return true;
    }
    return false;
  }

  // One dotted box: click/tap to upload, drag-drop, focus + Ctrl+V / paste
  if (whitebgDrop && whitebgFile) {
    wireDrop(whitebgDrop, whitebgFile, function (files) {
      var f = files && files[0];
      if (f) loadWhitebgBlob(f, f.name);
    });
  }

  if (whitebgClip) {
    whitebgClip.addEventListener("click", function () {
      if (!whitebgSourceBlob) return;
      openWhitebgCropMode();
    });
  }
  if (whitebgCropApply) {
    whitebgCropApply.addEventListener("click", function () {
      applyWhitebgCrop();
    });
  }
  if (whitebgCropCancel) {
    whitebgCropCancel.addEventListener("click", function () {
      exitWhitebgCropMode();
      if (whitebgSourceBlob) {
        if (whitebgPreviews) whitebgPreviews.hidden = false;
        if (whitebgActions) whitebgActions.hidden = false;
      }
    });
  }

  // Crop drag (mouse + touch)
  if (whitebgCropBox) {
    whitebgCropBox.addEventListener("mousedown", function (e) {
      if (e.target && e.target.getAttribute("data-handle")) return;
      startWhitebgCropDrag(e, "move");
    });
    whitebgCropBox.addEventListener(
      "touchstart",
      function (e) {
        if (e.target && e.target.getAttribute("data-handle")) return;
        startWhitebgCropDrag(e, "move");
      },
      { passive: false }
    );
  }
  if (whitebgCropStage) {
    whitebgCropStage.querySelectorAll(".whitebg-handle").forEach(function (handle) {
      var kind = handle.getAttribute("data-handle");
      handle.addEventListener("mousedown", function (e) {
        startWhitebgCropDrag(e, kind);
      });
      handle.addEventListener(
        "touchstart",
        function (e) {
          startWhitebgCropDrag(e, kind);
        },
        { passive: false }
      );
    });
  }
  window.addEventListener("mousemove", moveWhitebgCropDrag);
  window.addEventListener("mouseup", endWhitebgCropDrag);
  window.addEventListener(
    "touchmove",
    function (e) {
      if (whitebgCropDrag) moveWhitebgCropDrag(e);
    },
    { passive: false }
  );
  window.addEventListener("touchend", endWhitebgCropDrag);
  window.addEventListener("touchcancel", endWhitebgCropDrag);
  window.addEventListener("resize", function () {
    if (whitebgCropStage && !whitebgCropStage.hidden) layoutWhitebgCropBox();
  });

  if (whitebgDownload) {
    whitebgDownload.addEventListener("click", function () {
      if (!whitebgResultBlob) return;
      downloadBlob(whitebgResultBlob, whitebgOutName(whitebgSourceName));
    });
  }
  if (whitebgClear) {
    whitebgClear.addEventListener("click", clearWhitebg);
  }

  if (whitebgPanel) {
    ["dragenter", "dragover"].forEach(function (ev) {
      whitebgPanel.addEventListener(ev, function (e) {
        if (!isWhitebgPanelActive()) return;
        e.preventDefault();
        e.stopPropagation();
      });
    });
    whitebgPanel.addEventListener("drop", function (e) {
      if (!isWhitebgPanelActive()) return;
      // Don't steal drops while cropping
      if (whitebgCropStage && !whitebgCropStage.hidden) return;
      e.preventDefault();
      e.stopPropagation();
      var blob = imageBlobFromClipboardData(e.dataTransfer);
      if (blob) {
        loadWhitebgBlob(
          blob,
          blob.name && blob.name !== "blob" ? blob.name : "dropped.png"
        );
      } else if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0]) {
        loadWhitebgBlob(e.dataTransfer.files[0], e.dataTransfer.files[0].name);
      }
    });
  }

  // ---- PPTX helpers ----
  function slideNames(zip) {
    return Object.keys(zip.files)
      .filter(function (n) {
        return /^ppt\/slides\/slide\d+\.xml$/.test(n);
      })
      .sort(function (a, b) {
        return parseInt(a.match(/\d+/)[0], 10) - parseInt(b.match(/\d+/)[0], 10);
      });
  }

  function parseSlideList(input, maxSlides) {
    var out = [];
    String(input || "")
      .split(",")
      .forEach(function (part) {
        part = part.trim();
        if (!part) return;
        if (part.indexOf("-") >= 0) {
          var bits = part.split("-");
          var a = parseInt(bits[0], 10);
          var b = parseInt(bits[1], 10);
          if (isNaN(a) || isNaN(b)) return;
          for (var n = a; n <= b; n++) {
            if (n >= 1 && n <= maxSlides) out.push(n);
          }
        } else {
          var n2 = parseInt(part, 10);
          if (!isNaN(n2) && n2 >= 1 && n2 <= maxSlides) out.push(n2);
        }
      });
    return out;
  }

  function localAll(doc, name) {
    return Array.prototype.filter.call(doc.getElementsByTagName("*"), function (el) {
      return el.localName === name;
    });
  }

  // ---- Merge slides ----
  var mergeSourceFile = null;
  var mergeSourceBuf = null; // cached ArrayBuffer — avoid re-reading / re-parsing freezes
  var mergeTargetFile = null;
  var mergeTargetHandle = null; // FileSystemFileHandle for overwrite (Chrome/Edge)
  var mergeSourceLoadId = 0;

  function setMergeSource(file) {
    mergeSourceFile = file;
    mergeSourceBuf = null;
    $("mergeSourceLabel").textContent = file ? file.name : "Drop or click";
  }
  function setMergeTarget(file, handle) {
    mergeTargetFile = file;
    mergeTargetHandle = handle || null;
    var label = file ? file.name : "Drop or click";
    if (file && handle) label = file.name + " (will replace this file)";
    else if (file) label = file.name + " (will download result)";
    $("mergeTargetLabel").textContent = label;
  }

  async function getMergeSourceBuffer() {
    if (mergeSourceBuf) return mergeSourceBuf;
    if (!mergeSourceFile) throw new Error("No source file.");
    mergeSourceBuf = await readFileAsArrayBuffer(mergeSourceFile);
    return mergeSourceBuf;
  }

  /**
   * Instant UI feedback, then count slides/pages in the background so the page
   * does not freeze while a multi‑MB PPTX/PDF is parsed.
   */
  async function loadSourceFromFile(f) {
    if (!isSourceFile(f)) return;
    var loadId = ++mergeSourceLoadId;
    setMergeSource(f);
    if ($("mergeSlides")) {
      $("mergeSlides").value = "";
      $("mergeSlides").placeholder = isPdfName(f.name)
        ? "counting pages…"
        : "counting slides…";
    }
    setStatus(
      $("mergeStatus"),
      "Source selected: " +
        f.name +
        (isPdfName(f.name)
          ? " — counting PDF pages…"
          : " — counting slides…"),
      "ok"
    );
    await yieldToUI();
    await yieldToUI();
    try {
      var buf = await readFileAsArrayBuffer(f);
      if (loadId !== mergeSourceLoadId) return;
      mergeSourceBuf = buf;
      await yieldToUI();

      var c = 0;
      if (isPdfName(f.name)) {
        var pdfjsLib = await ensurePdfJs();
        if (loadId !== mergeSourceLoadId) return;
        // slice() so PDF.js cannot detach the cached buffer
        var pdf = await pdfjsLib.getDocument({ data: buf.slice(0) }).promise;
        if (loadId !== mergeSourceLoadId) return;
        c = pdf.numPages || 0;
      } else {
        var z = await JSZip.loadAsync(buf, { checkCRC32: false });
        if (loadId !== mergeSourceLoadId) return;
        c = slideNames(z).length;
      }

      if (!c) {
        setStatus(
          $("mergeStatus"),
          isPdfName(f.name)
            ? "Source PDF has 0 pages — pick a different file."
            : "Source has 0 slides — pick a different PPTX.",
          "err"
        );
        return;
      }
      if ($("mergeSlides")) {
        $("mergeSlides").value = "1-" + c;
        $("mergeSlides").placeholder = "1-10";
      }
      setStatus(
        $("mergeStatus"),
        isPdfName(f.name)
          ? "Source PDF: " + c + " page(s)"
          : "Source: " + c + " slides",
        "ok"
      );
    } catch (e) {
      if (loadId !== mergeSourceLoadId) return;
      console.error(e);
      setStatus(
        $("mergeStatus"),
        "Could not read source: " + (e.message || e),
        "err"
      );
    }
  }
  async function loadTargetFromFile(f, handle) {
    if (!f || !/\.pptx$/i.test(f.name || "")) return;
    setMergeTarget(f, handle);
    setStatus($("mergeStatus"), "Reading target…", "ok");
    await yieldToUI();
    try {
      var z = await JSZip.loadAsync(await readFileAsArrayBuffer(f), {
        checkCRC32: false,
      });
      var c = slideNames(z).length;
      // Blank shell (0 slides) → insert at start; otherwise append after last slide
      $("mergeInsert").value = String(c);
      var msg =
        c === 0
          ? "Target: blank (0 slides)"
          : "Target: " + c + " slides";
      setStatus($("mergeStatus"), msg, "ok");
    } catch (e) {
      setStatus($("mergeStatus"), "Could not read target: " + (e.message || e), "err");
    }
  }

  // Prefer File System Access for TARGET so we can overwrite the real Desktop file
  $("mergeTargetDrop").addEventListener("click", async function (e) {
    if (e.target && e.target.tagName === "INPUT") return;
    e.preventDefault();
    e.stopPropagation();
    var handle = await pickPptxWithHandle();
    if (handle) {
      var file = await handle.getFile();
      await loadTargetFromFile(file, handle);
      return;
    }
    $("mergeTarget").value = "";
    $("mergeTarget").click();
  });
  $("mergeSourceDrop").addEventListener("click", async function (e) {
    if (e.target && e.target.tagName === "INPUT") return;
    e.preventDefault();
    e.stopPropagation();
    var handle = await pickFileWithHandle(true);
    if (handle) {
      var file = await handle.getFile();
      await loadSourceFromFile(file);
      return;
    }
    $("mergeSource").value = "";
    $("mergeSource").click();
  });

  wireDrop(
    $("mergeSourceDrop"),
    $("mergeSource"),
    function (files) {
      var f = files[0];
      if (f) loadSourceFromFile(f);
    },
    { skipClick: true }
  );
  wireDrop(
    $("mergeTargetDrop"),
    $("mergeTarget"),
    function (files) {
      var f = files[0];
      // Drop does not give a writeable handle in browsers
      if (f) loadTargetFromFile(f, null);
    },
    { skipClick: true }
  );

  function readFileAsArrayBuffer(file) {
    if (file.arrayBuffer) return file.arrayBuffer();
    return new Promise(function (resolve, reject) {
      var fr = new FileReader();
      fr.onload = function () {
        resolve(fr.result);
      };
      fr.onerror = function () {
        reject(fr.error || new Error("Could not read file"));
      };
      fr.readAsArrayBuffer(file);
    });
  }

  function eachRelationship(relsXml, fn) {
    // Self-closing and paired tags
    relsXml.replace(/<Relationship\b[^>]*\/>/g, function (tag) {
      fn(tag);
      return tag;
    });
    relsXml.replace(/<Relationship\b[^>]*>[\s\S]*?<\/Relationship>/g, function (tag) {
      fn(tag);
      return tag;
    });
  }

  function isSlideRelationship(tag) {
    // Type must be .../relationships/slide  (not slideMaster / slideLayout / notesSlide)
    var typeM = /Type="([^"]+)"/.exec(tag);
    if (!typeM) return false;
    var t = typeM[1];
    return /\/relationships\/slide$/.test(t);
  }

  function normalizeSlidePath(target) {
    var path = String(target || "").replace(/\\/g, "/").replace(/^\//, "");
    if (path.indexOf("ppt/") === 0) return path;
    if (path.indexOf("slides/") === 0) return "ppt/" + path;
    if (path.indexOf("../") === 0) return "ppt/" + path.replace(/^(\.\.\/)+/, "");
    return "ppt/" + path;
  }

  /**
   * Extract selected slides from SOURCE into a new PPTX.
   * Clones the source package and keeps only chosen slides — media paths untouched.
   */
  async function extractSlidesFromSource(srcFile, nums) {
    var buf = await readFileAsArrayBuffer(srcFile);
    var zip = await JSZip.loadAsync(buf);

    var prsFile = zip.file("ppt/presentation.xml");
    var relsFile = zip.file("ppt/_rels/presentation.xml.rels");
    if (!prsFile || !relsFile) {
      throw new Error("Not a valid .pptx (missing presentation.xml).");
    }

    var prsXml = await prsFile.async("string");
    var prsRels = await relsFile.async("string");

    // rId -> ppt/slides/slideN.xml
    var ridToTarget = {};
    eachRelationship(prsRels, function (tag) {
      if (!isSlideRelationship(tag)) return;
      var idM = /Id="([^"]+)"/.exec(tag);
      var tM = /Target="([^"]+)"/.exec(tag);
      if (idM && tM) ridToTarget[idM[1]] = normalizeSlidePath(tM[1]);
    });

    // Document order of slides (self-closing or paired sldId)
    var order = [];
    function pushSldId(tag) {
      var ridM = /r:id="([^"]+)"/.exec(tag) || /r:id='([^']+)'/.exec(tag);
      if (!ridM) return;
      var rid = ridM[1];
      var path = ridToTarget[rid];
      if (!path) return;
      // Prefer actual zip path if case differs
      if (!zip.file(path)) {
        var alt = slideNames(zip).find(function (p) {
          return p.toLowerCase() === path.toLowerCase();
        });
        if (alt) path = alt;
      }
      if (!zip.file(path)) return;
      order.push({ rid: rid, path: path });
    }
    prsXml.replace(/<p:sldId\b[^/]*\/>/g, function (tag) {
      pushSldId(tag);
      return tag;
    });
    prsXml.replace(/<p:sldId\b[^>]*>[\s\S]*?<\/p:sldId>/g, function (tag) {
      pushSldId(tag);
      return tag;
    });

    // Fallback: match slide files to relationship targets by path
    if (!order.length) {
      var files = slideNames(zip);
      if (!files.length) throw new Error("No slides found inside this PowerPoint.");
      var pathToRid = {};
      Object.keys(ridToTarget).forEach(function (rid) {
        var p = ridToTarget[rid];
        pathToRid[p] = rid;
        pathToRid[normalizeSlidePath(p)] = rid;
        pathToRid[p.replace(/^ppt\//, "")] = rid;
      });
      files.forEach(function (path) {
        var rid =
          pathToRid[path] ||
          pathToRid[path.replace(/^ppt\//, "")] ||
          pathToRid["slides/" + path.split("/").pop()];
        if (rid) order.push({ rid: rid, path: path });
      });
    }

    if (!order.length) {
      throw new Error(
        "Could not read slide list from this PowerPoint (relationships/sldIdLst)."
      );
    }

    var selected = [];
    nums.forEach(function (n) {
      if (n >= 1 && n <= order.length) selected.push(order[n - 1]);
    });
    if (!selected.length) {
      throw new Error(
        "No slides matched. This file has " +
          order.length +
          " slides. Try 1-" +
          order.length +
          "."
      );
    }

    var keepRid = {};
    var keepSlidePath = {};
    selected.forEach(function (e) {
      keepRid[e.rid] = true;
      keepSlidePath[e.path] = true;
      // also basename match
      keepSlidePath[e.path.split("/").pop()] = true;
    });

    // Rebuild sldIdLst
    var newTags = selected.map(function (e, i) {
      return '<p:sldId id="' + (256 + i) + '" r:id="' + e.rid + '"/>';
    });
    var newList = "<p:sldIdLst>" + newTags.join("") + "</p:sldIdLst>";
    if (/<p:sldIdLst\/>/.test(prsXml)) {
      prsXml = prsXml.replace(/<p:sldIdLst\/>/, newList);
    } else if (/<p:sldIdLst[\s>][\s\S]*?<\/p:sldIdLst>/.test(prsXml)) {
      prsXml = prsXml.replace(/<p:sldIdLst[\s>][\s\S]*?<\/p:sldIdLst>/, newList);
    } else if (/<\/p:sldMasterIdLst>/.test(prsXml)) {
      prsXml = prsXml.replace(/<\/p:sldMasterIdLst>/, "</p:sldMasterIdLst>" + newList);
    } else {
      throw new Error("presentation.xml has no sldIdLst — cannot extract.");
    }
    zip.file("ppt/presentation.xml", prsXml);

    // Filter presentation relationships (handle /> and ></Relationship>)
    function filterRelTag(tag) {
      if (!isSlideRelationship(tag)) return tag;
      var idM = /Id="([^"]+)"/.exec(tag);
      if (idM && !keepRid[idM[1]]) return "";
      return tag;
    }
    prsRels = prsRels.replace(/<Relationship\b[^>]*\/>/g, filterRelTag);
    prsRels = prsRels.replace(/<Relationship\b[^>]*>[\s\S]*?<\/Relationship>/g, filterRelTag);
    zip.file("ppt/_rels/presentation.xml.rels", prsRels);

    // Remove slide parts not kept; clean kept slides (watermark strip + source/note fix)
    var cleanJobs = [];
    slideNames(zip).forEach(function (path) {
      var base = path.split("/").pop();
      if (keepSlidePath[path] || keepSlidePath[base]) {
        cleanJobs.push(
          zip
            .file(path)
            .async("string")
            .then(function (xml) {
              zip.file(path, cleanSlideXmlLikeDesktop(xml));
            })
        );
        return;
      }
      zip.remove(path);
      var relsPath = "ppt/slides/_rels/" + base + ".rels";
      if (zip.file(relsPath)) zip.remove(relsPath);
    });
    await Promise.all(cleanJobs);

    // Verify kept slides still exist
    var kept = 0;
    selected.forEach(function (e) {
      if (zip.file(e.path)) kept++;
    });
    if (kept === 0) {
      throw new Error("Internal error: all slides were removed. Please try again or report this file.");
    }

    // Content types: drop overrides for removed slides
    var ctXml = await zip.file("[Content_Types].xml").async("string");
    ctXml = ctXml.replace(
      /<Override\b[^>]*PartName="\/ppt\/slides\/slide\d+\.xml"[^>]*\/>/g,
      function (tag) {
        var m = /PartName="\/ppt\/slides\/(slide\d+\.xml)"/.exec(tag);
        if (!m) return tag;
        var base = m[1];
        var full = "ppt/slides/" + base;
        return keepSlidePath[full] || keepSlidePath[base] ? tag : "";
      }
    );
    zip.file("[Content_Types].xml", ctXml);

    var blob = await zip.generateAsync({
      type: "blob",
      mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      compression: "DEFLATE",
    });
    if (!blob || blob.size < 2000) {
      throw new Error("Output file was empty — extraction failed.");
    }
    return { blob: blob, count: selected.length, totalInSource: order.length };
  }

  function pathSlideNum(path) {
    var m = String(path).match(/slide(\d+)\.xml$/i);
    return m ? m[1] : "1";
  }

  /**
   * Desktop AMC Pdf's modification cleanup — STRING ONLY (no DOMParser).
   * Browser DOM serialize breaks OOXML namespaces and kills pictures.
   * Matches Python: fix Source:/Note:/short text box margins + bullets.
   */
  function cleanSlideXmlLikeDesktop(xml) {
    try {
      if (!xml) return xml;
      var lower = xml.toLowerCase();
      // Desktop only rewrites boxes that contain Source:/Note: or short text.
      // Without DOM we only auto-fix when Source:/Note: appear (safe for AMC).
      if (lower.indexOf("source:") < 0 && lower.indexOf("note:") < 0) {
        return xml;
      }

      // bodyPr: zero existing insets
      xml = xml.replace(/<a:bodyPr\b([^>]*?)(\/?>)/g, function (all, attrs, end) {
        ["lIns", "tIns", "rIns", "bIns"].forEach(function (a) {
          var re = new RegExp("\\s" + a + '="[^"]*"', "g");
          if (re.test(attrs)) {
            attrs = attrs.replace(re, " " + a + '="0"');
          }
        });
        return "<a:bodyPr" + attrs + end;
      });

      // pPr: lvl=0, drop marL/indent, algn=l
      xml = xml.replace(/<a:pPr\b([^>]*?)(\/?>)/g, function (all, attrs, end) {
        if (/\slvl="/.test(attrs)) attrs = attrs.replace(/\slvl="[^"]*"/g, ' lvl="0"');
        attrs = attrs.replace(/\smarL="[^"]*"/g, "");
        attrs = attrs.replace(/\sindent="[^"]*"/g, "");
        if (/\salgn="/.test(attrs)) attrs = attrs.replace(/\salgn="[^"]*"/g, ' algn="l"');
        return "<a:pPr" + attrs + end;
      });

      xml = xml.replace(
        /<a:bu(?:None|SzPct|SzPts|Char|Blip|Font|Clr|FontTx|ClrTx)\b[^>]*\/>/g,
        ""
      );
      xml = xml.replace(
        /<a:bu(?:None|SzPct|SzPts|Char|Blip|Font|Clr|FontTx|ClrTx)\b[^>]*>[\s\S]*?<\/a:bu(?:None|SzPct|SzPts|Char|Blip|Font|Clr|FontTx|ClrTx)>/g,
        ""
      );
      return xml;
    } catch (e) {
      console.warn("cleanSlideXmlLikeDesktop failed", e);
      return xml;
    }
  }

  function escapeRegExp(s) {
    return String(s).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  }

  /**
   * clean selected source slides → insert into target (blank shell).
   * Media remapped with _src_ prefix. Speaker notes / footnotes are preserved.
   */
  async function mergePresentations(srcFile, tgtFile, nums, insertAfter) {
    var srcData =
      srcFile instanceof ArrayBuffer
        ? srcFile
        : await readFileAsArrayBuffer(srcFile);
    var tgtData =
      tgtFile instanceof ArrayBuffer
        ? tgtFile
        : await readFileAsArrayBuffer(tgtFile);
    var srcZip = await JSZip.loadAsync(srcData, { checkCRC32: false });
    var tgtZip = await JSZip.loadAsync(tgtData, { checkCRC32: false });
    var tgtCount = slideNames(tgtZip).length;
    if (insertAfter > tgtCount) insertAfter = tgtCount;

    // Match target slide size (and notes size) to source
    try {
      var srcPrs = await srcZip.file("ppt/presentation.xml").async("string");
      var tgtPrs0 = await tgtZip.file("ppt/presentation.xml").async("string");
      var sz = srcPrs.match(/<p:sldSz\b[^/]*\/>/);
      if (sz && /<p:sldSz\b[^/]*\/>/.test(tgtPrs0)) {
        tgtPrs0 = tgtPrs0.replace(/<p:sldSz\b[^/]*\/>/, sz[0]);
      }
      var nsz = srcPrs.match(/<p:notesSz\b[^/]*\/>/);
      if (nsz && /<p:notesSz\b[^/]*\/>/.test(tgtPrs0)) {
        tgtPrs0 = tgtPrs0.replace(/<p:notesSz\b[^/]*\/>/, nsz[0]);
      }
      tgtZip.file("ppt/presentation.xml", tgtPrs0);
    } catch (e) {}

    // Media: _src_ + original name; longest-first remap
    var mediaMap = {};
    var srcMedia = Object.keys(srcZip.files)
      .filter(function (n) {
        return n.indexOf("ppt/media/") === 0 && !srcZip.files[n].dir;
      })
      .sort();
    for (var mi = 0; mi < srcMedia.length; mi++) {
      var mp = srcMedia[mi];
      var base = mp.split("/").pop();
      var dot = base.lastIndexOf(".");
      var stem = dot >= 0 ? base.slice(0, dot) : base;
      var ext = dot >= 0 ? base.slice(dot) : "";
      var newBase = "_src_" + base;
      var newPath = "ppt/media/" + newBase;
      var c = 1;
      while (tgtZip.file(newPath)) {
        newBase = "_src_" + stem + "_" + c + ext;
        newPath = "ppt/media/" + newBase;
        c++;
      }
      tgtZip.file(newPath, await srcZip.file(mp).async("uint8array"));
      mediaMap[base] = newBase;
    }

    function remapMedia(text) {
      if (!text) return text;
      Object.keys(mediaMap)
        .sort(function (a, b) {
          return b.length - a.length;
        })
        .forEach(function (oldN) {
          var neu = mediaMap[oldN];
          text = text.split("../media/" + oldN).join("../media/" + neu);
          text = text.split("/media/" + oldN).join("/media/" + neu);
        });
      return text;
    }

    function sanitizeContentTypes(ctXml) {
      ctXml = ctXml.replace(
        /<Types(?:\s+xmlns="[^"]*")+/g,
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"'
      );
      if (!/^[\s\S]*<Types\s+xmlns=/.test(ctXml)) {
        ctXml = ctXml.replace(
          /<Types\b/,
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"'
        );
      }
      ctXml = ctXml.replace(
        /<Types\s+xmlns="http:\/\/schemas\.openxmlformats\.org\/package\/2006\/content-types"\s+xmlns="[^"]*"/g,
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"'
      );
      return ctXml;
    }

    function ensureImageDefaults(ctXml) {
      ctXml = sanitizeContentTypes(ctXml);
      var defaults = {
        png: "image/png",
        jpg: "image/jpeg",
        jpeg: "image/jpeg",
        gif: "image/gif",
        emf: "image/x-emf",
        wmf: "image/x-wmf",
        tiff: "image/tiff",
        bin: "application/octet-stream",
      };
      Object.keys(defaults).forEach(function (ext) {
        if (!new RegExp('Extension="' + ext + '"', "i").test(ctXml)) {
          ctXml = ctXml.replace(
            "</Types>",
            '<Default Extension="' +
              ext +
              '" ContentType="' +
              defaults[ext] +
              '"/></Types>'
          );
        }
      });
      return ctXml;
    }

    /**
     * Copy notes as-is from source (default formatting / bold / spaces unchanged).
     * Notes rels only link to the parent slide (no notesMaster into blank = no Repair).
     */
    var addedMeta = [];
    var nextSlideNum = tgtCount + 1;
    var notesCounter = 1;
    Object.keys(tgtZip.files).forEach(function (n) {
      var m = n.match(/^ppt\/notesSlides\/notesSlide(\d+)\.xml$/i);
      if (m) notesCounter = Math.max(notesCounter, parseInt(m[1], 10) + 1);
    });
    var notesPartsAdded = [];

    for (var i = 0; i < nums.length; i++) {
      var sn = nums[i];
      var sp = "ppt/slides/slide" + sn + ".xml";
      var rp = "ppt/slides/_rels/slide" + sn + ".xml.rels";
      if (!srcZip.file(sp)) continue;

      var newNum = nextSlideNum++;
      var nsp = "ppt/slides/slide" + newNum + ".xml";
      var nrp = "ppt/slides/_rels/slide" + newNum + ".xml.rels";

      var xml = await srcZip.file(sp).async("string");
      xml = cleanSlideXmlLikeDesktop(xml);
      xml = remapMedia(xml);
      xml = xml.replace(/<p:extLst>[\s\S]*?<\/p:extLst>/g, "");
      if (xml.indexOf("<?xml") !== 0) {
        xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + xml;
      }
      tgtZip.file(nsp, xml);

      if (srcZip.file(rp)) {
        var rels = await srcZip.file(rp).async("string");
        rels = remapMedia(rels);

        var notesMatch = rels.match(
          /Target="([^"]*notesSlides\/notesSlide\d+\.xml)"/i
        );
        if (notesMatch) {
          var notesTarget = notesMatch[1].replace(/^\.\.\//, "ppt/");
          if (notesTarget.indexOf("ppt/") !== 0) {
            notesTarget = "ppt/notesSlides/" + notesTarget.split("/").pop();
          }
          if (srcZip.file(notesTarget)) {
            var newNotesName = "notesSlide" + notesCounter + ".xml";
            var newNotesPath = "ppt/notesSlides/" + newNotesName;
            // Exact source notes XML — no rewrite of bold/spacing
            var srcNotesXml = await srcZip.file(notesTarget).async("string");
            srcNotesXml = remapMedia(srcNotesXml);
            if (srcNotesXml.indexOf("<?xml") !== 0) {
              srcNotesXml =
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
                srcNotesXml;
            }
            tgtZip.file(newNotesPath, srcNotesXml);
            notesPartsAdded.push(newNotesPath);

            tgtZip.file(
              "ppt/notesSlides/_rels/" + newNotesName + ".rels",
              '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/slide' +
                newNum +
                '.xml"/>' +
                "</Relationships>"
            );

            rels = rels.replace(/<Relationship\b[^>]*notesMaster[^>]*\/>/gi, "");
            if (/notesSlide/i.test(rels)) {
              rels = rels.replace(
                /Target="[^"]*notesSlides\/notesSlide\d+\.xml"/gi,
                'Target="../notesSlides/' + newNotesName + '"'
              );
            } else {
              rels = rels.replace(
                "</Relationships>",
                '<Relationship Id="rIdNotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/' +
                  newNotesName +
                  '"/></Relationships>'
              );
            }
            notesCounter++;
          } else {
            rels = rels.replace(/<Relationship\b[^>]*notesSlide[^>]*\/>/gi, "");
          }
        }

        rels = rels.replace(/<Relationship\b[^>]*notesMaster[^>]*\/>/gi, "");

        if (rels.indexOf("<?xml") !== 0) {
          rels =
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + rels;
        }
        tgtZip.file(nrp, rels);
      } else {
        tgtZip.file(
          nrp,
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
        );
      }

      addedMeta.push({ newNum: newNum });
    }

    if (!addedMeta.length) {
      throw new Error(
        "No slides were copied. Check slide numbers (source may use different numbering)."
      );
    }

    // presentation.xml.rels — slides only
    var prsRelsXml = await tgtZip
      .file("ppt/_rels/presentation.xml.rels")
      .async("string");
    prsRelsXml = prsRelsXml.replace(
      /<Relationship\b[^>]*notesMaster[^>]*\/>/gi,
      ""
    );
    var maxRid = 0;
    prsRelsXml.replace(/Id="rId(\d+)"/g, function (_, n) {
      maxRid = Math.max(maxRid, parseInt(n, 10));
      return _;
    });

    var relChunks = [];
    addedMeta.forEach(function (meta, idx) {
      meta.rid = "rId" + (maxRid + idx + 1);
      relChunks.push(
        '<Relationship Id="' +
          meta.rid +
          '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide' +
          meta.newNum +
          '.xml"/>'
      );
    });
    if (prsRelsXml.indexOf("</Relationships>") < 0) {
      throw new Error("Invalid presentation.xml.rels in target");
    }
    prsRelsXml = prsRelsXml.replace(
      "</Relationships>",
      relChunks.join("") + "</Relationships>"
    );
    tgtZip.file("ppt/_rels/presentation.xml.rels", prsRelsXml);

    // presentation.xml sldIdLst — no notesMasterIdLst
    var prsXml = await tgtZip.file("ppt/presentation.xml").async("string");
    prsXml = prsXml.replace(
      /<p:notesMasterIdLst>[\s\S]*?<\/p:notesMasterIdLst>/g,
      ""
    );
    prsXml = prsXml.replace(/<p:notesMasterIdLst\s*\/>/g, "");

    var maxId = 255;
    prsXml.replace(/<p:sldId\b[^>]*\bid="(\d+)"/g, function (_, n) {
      maxId = Math.max(maxId, parseInt(n, 10));
      return _;
    });

    var newSldTags = addedMeta.map(function (meta, idx) {
      meta.id = maxId + idx + 1;
      return '<p:sldId id="' + meta.id + '" r:id="' + meta.rid + '"/>';
    });

    var existingTags = [];
    prsXml.replace(/<p:sldId\b[^/]*\/>/g, function (tag) {
      existingTags.push(tag);
      return tag;
    });

    var ordered;
    if (insertAfter < tgtCount && existingTags.length) {
      ordered = existingTags
        .slice(0, insertAfter)
        .concat(newSldTags)
        .concat(existingTags.slice(insertAfter));
    } else {
      ordered = existingTags.concat(newSldTags);
    }
    var newList = "<p:sldIdLst>" + ordered.join("") + "</p:sldIdLst>";

    if (/<p:sldIdLst\s*\/>/.test(prsXml)) {
      prsXml = prsXml.replace(/<p:sldIdLst\s*\/>/, newList);
    } else if (/<p:sldIdLst[\s>][\s\S]*?<\/p:sldIdLst>/.test(prsXml)) {
      prsXml = prsXml.replace(/<p:sldIdLst[\s>][\s\S]*?<\/p:sldIdLst>/, newList);
    } else if (/<\/p:sldMasterIdLst>/.test(prsXml)) {
      prsXml = prsXml.replace(
        /<\/p:sldMasterIdLst>/,
        "</p:sldMasterIdLst>" + newList
      );
    } else {
      throw new Error("Could not insert sldIdLst into presentation.xml");
    }
    if (prsXml.indexOf("<?xml") !== 0) {
      prsXml =
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + prsXml;
    }
    tgtZip.file("ppt/presentation.xml", prsXml);

    // Content_Types
    var ctXml = await tgtZip.file("[Content_Types].xml").async("string");
    ctXml = ensureImageDefaults(ctXml);
    var ctAdds = [];
    addedMeta.forEach(function (meta) {
      var part = "/ppt/slides/slide" + meta.newNum + ".xml";
      if (ctXml.indexOf('PartName="' + part + '"') < 0) {
        ctAdds.push(
          '<Override PartName="' +
            part +
            '" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        );
      }
    });
    notesPartsAdded.forEach(function (npath) {
      var part = "/" + npath;
      if (ctXml.indexOf('PartName="' + part + '"') < 0) {
        ctAdds.push(
          '<Override PartName="' +
            part +
            '" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>'
        );
      }
    });
    // Strip any leftover notesMaster override from blank target (we don't ship one)
    ctXml = ctXml.replace(
      /<Override[^>]*notesMaster[^>]*\/>/gi,
      ""
    );
    if (ctAdds.length) {
      ctXml = ctXml.replace("</Types>", ctAdds.join("") + "</Types>");
    }
    ctXml = sanitizeContentTypes(ctXml);
    if ((ctXml.match(/\sxmlns="/g) || []).length > 1) {
      ctXml = ctXml.replace(
        /<Types[^>]*>/,
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      );
    }
    if (ctXml.indexOf("</Types>") < 0) {
      throw new Error("Content_Types.xml became invalid during merge.");
    }
    tgtZip.file("[Content_Types].xml", ctXml);

    var ab = await tgtZip.generateAsync({
      type: "arraybuffer",
      compression: "DEFLATE",
      compressionOptions: { level: 6 },
    });
    if (!ab || ab.byteLength < 2000) {
      throw new Error("Output file was empty — merge failed.");
    }
    var blob = new Blob([ab], {
      type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    });
    return {
      blob: blob,
      count: addedMeta.length,
      notesCopied: notesPartsAdded.length,
    };
  }

  function blankTargetFromB64() {
    if (typeof window.AMC_BLANK_TARGET_B64 !== "string" || !window.AMC_BLANK_TARGET_B64) {
      return null;
    }
    try {
      var bin = atob(window.AMC_BLANK_TARGET_B64);
      var bytes = new Uint8Array(bin.length);
      for (var i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
      return new File([bytes], "blank-target.pptx", {
        type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      });
    } catch (e) {
      console.warn("embedded blank decode failed", e);
      return null;
    }
  }

  // Load built-in blank target (fetch file, else embedded base64 — no network required)
  async function ensureBlankTarget() {
    var lastErr = null;

    // 1) Embedded base64 (always works, even offline / file://)
    try {
      var emb = blankTargetFromB64();
      if (emb) {
        await loadTargetFromFile(emb, null);
        $("mergeTargetLabel").textContent = "blank-target.pptx (built-in)";
        $("mergeInsert").value = "0";
        setStatus($("mergeStatus"), "", "ok");
        return true;
      }
    } catch (e) {
      lastErr = e;
    }

    // 2) Fetch blank-target.pptx next to this page
    var urls = ["blank-target.pptx", "./blank-target.pptx"];
    try {
      urls.push(new URL("blank-target.pptx", location.href).href);
    } catch (e) {}

    for (var i = 0; i < urls.length; i++) {
      try {
        var res = await fetch(urls[i], { cache: "no-store" });
        if (!res.ok) throw new Error("HTTP " + res.status + " for " + urls[i]);
        var buf = await res.arrayBuffer();
        if (!buf || buf.byteLength < 1000) throw new Error("File too small / empty");
        var file = new File([buf], "blank-target.pptx", {
          type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        });
        await loadTargetFromFile(file, null);
        $("mergeTargetLabel").textContent = "blank-target.pptx (built-in)";
        $("mergeInsert").value = "0";
        setStatus($("mergeStatus"), "", "ok");
        return true;
      } catch (e) {
        lastErr = e;
        console.warn("blank-target load failed:", urls[i], e);
      }
    }
    setStatus(
      $("mergeStatus"),
      "No blank target — click Target and pick any blank .pptx, or re-upload blank-target-data.js. (" +
        (lastErr && lastErr.message ? lastErr.message : "load failed") +
        ")",
      "err"
    );
    return false;
  }

  function makeEditOutName(srcName) {
    var base = (srcName || "slides")
      .replace(/\.pptx$/i, "")
      .replace(/\.pdf$/i, "");
    base = base.replace(/_extracted_[\d\-T:.]+$/i, "");
    base = base.replace(/_EDIT$/i, "");
    return base + "_EDIT.pptx";
  }

  /**
   * Shared extract path used by both buttons (PPTX slides or PDF pages → blank PPTX).
   * @returns {{ blob: Blob, count: number, outName: string }}
   */
  async function runExtractSlides() {
    if (!mergeSourceFile) {
      throw new Error("Choose a Source PPTX or PDF file first.");
    }
    if (typeof JSZip === "undefined") {
      throw new Error("JSZip missing.");
    }
    if (!mergeTargetFile) {
      var ok = await ensureBlankTarget();
      if (!ok || !mergeTargetFile) {
        throw new Error(
          "No blank target. Upload blank-target.pptx next to amc-studio.html, or click Target and pick a blank .pptx."
        );
      }
    }

    setStatus($("mergeStatus"), "Reading source…");
    await yieldToUI();
    var srcBuf = await getMergeSourceBuffer();
    var sourceIsPdf = isPdfName(mergeSourceFile.name);
    var srcCount = 0;

    if (sourceIsPdf) {
      var pdfjsLib = await ensurePdfJs();
      var pdfDoc = await pdfjsLib
        .getDocument({ data: srcBuf.slice(0) })
        .promise;
      srcCount = pdfDoc.numPages || 0;
      if (!srcCount) throw new Error("No pages found in PDF.");
    } else {
      var probe = await JSZip.loadAsync(srcBuf, { checkCRC32: false });
      srcCount = slideNames(probe).length;
      if (!srcCount) throw new Error("No slides found in source file.");
    }

    var rawList = ($("mergeSlides") && $("mergeSlides").value) || "";
    if (!String(rawList).trim()) {
      rawList = "1-" + srcCount;
      if ($("mergeSlides")) $("mergeSlides").value = rawList;
    }
    var nums = parseSlideList(rawList, srcCount);
    if (!nums.length) {
      throw new Error(
        "No valid " +
          (sourceIsPdf ? "page" : "slide") +
          " numbers. This source has " +
          srcCount +
          " (try 1-" +
          srcCount +
          ")."
      );
    }

    var insertAfter = parseInt(($("mergeInsert") && $("mergeInsert").value) || "0", 10);
    if (isNaN(insertAfter) || insertAfter < 0) insertAfter = 0;

    var result;
    if (sourceIsPdf) {
      setStatus(
        $("mergeStatus"),
        "Rendering PDF pages as images (includes text overlays)…"
      );
      await yieldToUI();
      var images = await renderPdfPagesToImages(
        srcBuf.slice(0),
        nums,
        function (done, total, pageNum) {
          setStatus(
            $("mergeStatus"),
            "Rendering PDF page " +
              pageNum +
              " (" +
              done +
              "/" +
              total +
              ")…"
          );
        }
      );
      setStatus(
        $("mergeStatus"),
        "Inserting " + images.length + " page image(s) into blank target…"
      );
      await yieldToUI();
      result = await mergePdfImagesIntoTarget(
        mergeTargetFile,
        images,
        insertAfter
      );
    } else {
      setStatus(
        $("mergeStatus"),
        "Cleaning " + nums.length + " slide(s) + inserting into blank target…"
      );
      await yieldToUI();
      result = await mergePresentations(
        srcBuf,
        mergeTargetFile,
        nums,
        insertAfter
      );
    }

    return {
      blob: result.blob,
      count: result.count,
      outName: makeEditOutName(mergeSourceFile.name),
    };
  }

  function setMergeButtonsDisabled(disabled) {
    if ($("mergeGo")) $("mergeGo").disabled = disabled;
    if ($("mergeGoViewer")) $("mergeGoViewer").disabled = disabled;
  }

  $("mergeGo").addEventListener("click", async function () {
    setMergeButtonsDisabled(true);
    setStatus($("mergeStatus"), "Cleaning + inserting into blank target…");
    try {
      var result = await runExtractSlides();
      downloadBlob(result.blob, result.outName);
      setStatus(
        $("mergeStatus"),
        "Downloaded “" +
          result.outName +
          "” (" +
          result.count +
          " slide(s) into blank target). Check Downloads.",
        "ok"
      );
    } catch (e) {
      console.error(e);
      setStatus($("mergeStatus"), "Error: " + (e.message || e), "err");
    }
    setMergeButtonsDisabled(false);
  });

  // Red button: extract only → open Notes Viewer (no download; SAVE from viewer later)
  if ($("mergeGoViewer")) {
    $("mergeGoViewer").addEventListener("click", async function () {
      setMergeButtonsDisabled(true);
      setStatus($("mergeStatus"), "Extracting, then opening Notes Viewer…");

      // Open placeholder tab immediately while we still have the click gesture
      var pre = null;
      try {
        if (window.PptxNotesLoader && PptxNotesLoader.openPlaceholderTab) {
          pre = PptxNotesLoader.openPlaceholderTab();
        }
      } catch (e) {
        pre = null;
      }

      try {
        var result = await runExtractSlides();
        // Do NOT download here — user edits in Notes Viewer and SAVE from there

        var file = new File([result.blob], result.outName, {
          type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        });

        if (!window.PptxNotesLoader || !PptxNotesLoader.openPptxInNewTab) {
          throw new Error(
            "Notes Viewer loader missing — re-upload pptx-loader.js next to amc-studio.html."
          );
        }

        setStatus($("mergeStatus"), "Opening Notes Viewer with “" + result.outName + "”…");
        await yieldToUI();

        var opened = await PptxNotesLoader.openPptxInNewTab(file, "viewer.html", {
          preOpenedWindow: pre,
          sameTabFallback: true,
          forceNewTab: true,
        });

        setStatus(
          $("mergeStatus"),
          "Opened Notes Viewer with “" +
            result.outName +
            "” (" +
            result.count +
            " slides). Edit notes, then SAVE there — no download yet" +
            (opened && opened.mode === "same-tab" ? " (this tab)." : "."),
          "ok"
        );
      } catch (e) {
        console.error(e);
        if (pre && !pre.closed) {
          try {
            pre.close();
          } catch (err) {}
        }
        setStatus($("mergeStatus"), "Error: " + (e.message || e), "err");
      }
      setMergeButtonsDisabled(false);
    });
  }

  // Preload built-in blank target (0-slide shell) — quiet unless error
  ensureBlankTarget();

  // ---- Image strip ----
  var stripFiles = [];
  function renderStripList() {
    var ul = $("stripList");
    ul.innerHTML = "";
    stripFiles.forEach(function (f) {
      var li = document.createElement("li");
      li.textContent = f.name;
      ul.appendChild(li);
    });
  }
  wireDrop($("stripDrop"), $("stripFiles"), function (files) {
    Array.prototype.forEach.call(files, function (f) {
      if (/\.pptx$/i.test(f.name)) stripFiles.push(f);
    });
    renderStripList();
    setStatus($("stripStatus"), stripFiles.length + " file(s) ready.", "ok");
  });
  $("stripClear").addEventListener("click", function () {
    stripFiles = [];
    renderStripList();
    setStatus($("stripStatus"), "Cleared.");
  });

  /**
   * Thorough image strip (string-based — no DOM corruption):
   * 1) Remove all picture shapes (p:pic)
   * 2) Clear background image fills (p:bg with blipFill) — e.g. full-slide photos
   * 3) Replace solid/gradient slide backgrounds with white (avoids dark leftover pages)
   * 4) Remove picture fills on shapes (a:blipFill → noFill)
   */
  function stripImagesFromSlideXml(xml) {
    if (!xml) return xml;

    // 1) Picture shapes (loop in case of nesting / multiple)
    var prev;
    do {
      prev = xml;
      xml = xml.replace(/<[a-zA-Z0-9]+:pic\b[\s\S]*?<\/[a-zA-Z0-9]+:pic>/g, "");
      xml = xml.replace(/<pic\b[\s\S]*?<\/pic>/g, "");
    } while (xml !== prev);

    // 2–3) Slide backgrounds: image OR colored solid/gradient → plain white
    // (slide 16 in AMC decks often uses dark schemeClr accent after the photo is gone)
    xml = xml.replace(/<p:bg\b[\s\S]*?<\/p:bg>/g, function (bg) {
      if (
        /blipFill|a:blip\b|solidFill|gradFill|pattFill/i.test(bg)
      ) {
        return (
          "<p:bg><p:bgPr>" +
          '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>' +
          "<a:effectLst/></p:bgPr></p:bg>"
        );
      }
      return bg;
    });

    // 4) Picture-as-shape-fill (and any leftover blip fills not inside bg we already rewrote)
    do {
      prev = xml;
      xml = xml.replace(/<a:blipFill\b[\s\S]*?<\/a:blipFill>/g, "<a:noFill/>");
    } while (xml !== prev);

    // Self-closing / rare forms
    xml = xml.replace(/<a:blipFill\b[^>]*\/>/g, "<a:noFill/>");

    return xml;
  }

  /** Drop image relationships from a slide .rels so PowerPoint does not look for missing media. */
  function stripImageRels(relsXml) {
    if (!relsXml) return relsXml;
    relsXml = relsXml.replace(
      /<Relationship\b[^>]*Type="[^"]*\/image"[^>]*\/>/gi,
      ""
    );
    relsXml = relsXml.replace(
      /<Relationship\b[^>]*Type="[^"]*\/image"[^>]*>[\s\S]*?<\/Relationship>/gi,
      ""
    );
    return relsXml;
  }

  async function stripImagesFromFile(file) {
    var zip = await JSZip.loadAsync(await readFileAsArrayBuffer(file), {
      checkCRC32: false,
    });
    var slides = slideNames(zip);
    for (var i = 0; i < slides.length; i++) {
      var path = slides[i];
      var xml = await zip.file(path).async("string");
      xml = stripImagesFromSlideXml(xml);
      if (xml.indexOf("<?xml") !== 0) {
        xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + xml;
      }
      zip.file(path, xml);

      // Clean slide relationships that pointed at images
      var base = path.split("/").pop();
      var relsPath = "ppt/slides/_rels/" + base + ".rels";
      if (zip.file(relsPath)) {
        var rels = await zip.file(relsPath).async("string");
        zip.file(relsPath, stripImageRels(rels));
      }
    }

    // Remove media parts (images) — optional size win; slides no longer reference them
    Object.keys(zip.files).forEach(function (n) {
      if (n.indexOf("ppt/media/") === 0 && !zip.files[n].dir) {
        zip.remove(n);
      }
    });

    var ab = await zip.generateAsync({
      type: "arraybuffer",
      compression: "DEFLATE",
      compressionOptions: { level: 6 },
    });
    return new Blob([ab], {
      type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    });
  }

  $("stripGo").addEventListener("click", async function () {
    if (!stripFiles.length) {
      setStatus($("stripStatus"), "Add at least one PPTX.", "err");
      return;
    }
    $("stripGo").disabled = true;
    var ok = 0;
    for (var i = 0; i < stripFiles.length; i++) {
      var f = stripFiles[i];
      setStatus($("stripStatus"), "Processing " + f.name + "…");
      try {
        var blob = await stripImagesFromFile(f);
        downloadBlob(blob, f.name.replace(/\.pptx$/i, "") + "_stripped.pptx");
        ok++;
      } catch (e) {
        console.error(e);
      }
    }
    setStatus($("stripStatus"), "Done — " + ok + " file(s) downloaded.", "ok");
    $("stripGo").disabled = false;
  });

  // ---- Special strip ----
  var specialFiles = [];
  function renderSpecialList() {
    var ul = $("specialList");
    ul.innerHTML = "";
    specialFiles.forEach(function (f) {
      var li = document.createElement("li");
      li.textContent = f.name;
      ul.appendChild(li);
    });
  }
  wireDrop($("specialDrop"), $("specialFiles"), function (files) {
    Array.prototype.forEach.call(files, function (f) {
      if (/\.pptx$/i.test(f.name)) specialFiles.push(f);
    });
    renderSpecialList();
    setStatus($("specialStatus"), specialFiles.length + " file(s) ready.", "ok");
  });
  $("specialClear").addEventListener("click", function () {
    specialFiles = [];
    renderSpecialList();
    setStatus($("specialStatus"), "Cleared.");
  });

  async function specialStripFile(file) {
    var zip = await JSZip.loadAsync(file);
    var slides = slideNames(zip);
    for (var i = 0; i < slides.length; i++) {
      var path = slides[i];
      var xml = await zip.file(path).async("string");
      var doc = new DOMParser().parseFromString(xml, "application/xml");
      // Remove sp, pic, graphicFrame under spTree
      ["sp", "pic", "graphicFrame", "cxnSp", "grpSp"].forEach(function (name) {
        localAll(doc, name).forEach(function (el) {
          // keep only if not inside notes — these are slide parts
          var p = el.parentNode;
          if (p) p.removeChild(el);
        });
      });
      localAll(doc, "bg").forEach(function (el) {
        if (el.parentNode) el.parentNode.removeChild(el);
      });
      zip.file(path, new XMLSerializer().serializeToString(doc));
    }
    // Drop media files
    Object.keys(zip.files).forEach(function (n) {
      if (n.indexOf("ppt/media/") === 0 && !zip.files[n].dir) {
        zip.remove(n);
      }
    });
    // Clean image relationships on slides
    Object.keys(zip.files).forEach(function (n) {
      if (!/^ppt\/slides\/_rels\/slide\d+\.xml\.rels$/.test(n)) return;
    });
    var relPaths = Object.keys(zip.files).filter(function (n) {
      return /^ppt\/slides\/_rels\/slide\d+\.xml\.rels$/.test(n);
    });
    for (var r = 0; r < relPaths.length; r++) {
      var rp = relPaths[r];
      var relsXml = await zip.file(rp).async("string");
      var rdoc = new DOMParser().parseFromString(relsXml, "application/xml");
      localAll(rdoc, "Relationship").forEach(function (rel) {
        var t = (rel.getAttribute("Type") || "").toLowerCase();
        if (t.indexOf("image") >= 0 || t.indexOf("media") >= 0) {
          if (rel.parentNode) rel.parentNode.removeChild(rel);
        }
      });
      zip.file(rp, new XMLSerializer().serializeToString(rdoc));
    }
    return zip.generateAsync({
      type: "blob",
      mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    });
  }

  $("specialGo").addEventListener("click", async function () {
    if (!specialFiles.length) {
      setStatus($("specialStatus"), "Add at least one PPTX.", "err");
      return;
    }
    $("specialGo").disabled = true;
    var ok = 0;
    for (var i = 0; i < specialFiles.length; i++) {
      var f = specialFiles[i];
      setStatus($("specialStatus"), "Processing " + f.name + "…");
      try {
        var blob = await specialStripFile(f);
        downloadBlob(blob, f.name.replace(/\.pptx$/i, "") + "_COMPLETE_STRIP.pptx");
        ok++;
      } catch (e) {
        console.error(e);
      }
    }
    setStatus($("specialStatus"), "Done — " + ok + " file(s) downloaded.", "ok");
    $("specialGo").disabled = false;
  });

  // ---- Rename _EDIT (filename only; file bytes unchanged) ----
  var editnameFiles = [];

  /** True when the name already ends with _EDIT.pptx */
  function pptxAlreadyEditNamed(name) {
    return /_EDIT\.pptx$/i.test(name || "");
  }

  /**
   * Put _EDIT at the very end of the stem, just before .pptx.
   * Cleans leftover _EDIT / trailing underscores (e.g. foo__EDIT_.pptx → foo_EDIT.pptx).
   */
  function makeRenameEditName(srcName) {
    var base = String(srcName || "slides").replace(/\.pptx$/i, "");
    base = base.replace(/_+EDIT_*$/i, "");
    base = base.replace(/_+$/g, "");
    if (!base) base = "slides";
    return base + "_EDIT.pptx";
  }

  function renderEditnameList() {
    var ul = $("editnameList");
    if (!ul) return;
    ul.innerHTML = "";
    editnameFiles.forEach(function (f) {
      var li = document.createElement("li");
      if (pptxAlreadyEditNamed(f.name)) {
        li.className = "rename-skip";
        li.textContent = f.name + "  (already has _EDIT)";
      } else {
        li.textContent = f.name + "  →  " + makeRenameEditName(f.name);
      }
      ul.appendChild(li);
    });
  }

  function addEditnameFiles(files) {
    var seen = {};
    editnameFiles.forEach(function (f) {
      seen[(f.name || "").toLowerCase()] = true;
    });
    Array.prototype.forEach.call(files, function (f) {
      if (!f || !/\.pptx$/i.test(f.name)) return;
      var key = f.name.toLowerCase();
      if (seen[key]) return;
      seen[key] = true;
      editnameFiles.push(f);
    });
    renderEditnameList();
    var n = editnameFiles.length;
    setStatus(
      $("editnameStatus"),
      n ? n + " file(s) ready." : "No .pptx files in that drop.",
      n ? "ok" : "err"
    );
  }

  if ($("editnameDrop") && $("editnameFiles")) {
    wireDrop($("editnameDrop"), $("editnameFiles"), addEditnameFiles);
  }
  if ($("editnameClear")) {
    $("editnameClear").addEventListener("click", function () {
      editnameFiles = [];
      renderEditnameList();
      setStatus($("editnameStatus"), "Cleared.");
    });
  }
  if ($("editnameGo")) {
    $("editnameGo").addEventListener("click", async function () {
      if (!editnameFiles.length) {
        setStatus($("editnameStatus"), "Add at least one PPTX.", "err");
        return;
      }
      var toRename = [];
      var skipped = 0;
      editnameFiles.forEach(function (f) {
        if (pptxAlreadyEditNamed(f.name)) skipped++;
        else toRename.push(f);
      });
      if (!toRename.length) {
        setStatus(
          $("editnameStatus"),
          skipped
            ? "Nothing to do — every file already ends with _EDIT.pptx."
            : "Add at least one PPTX.",
          skipped ? "ok" : "err"
        );
        return;
      }

      $("editnameGo").disabled = true;
      var ok = 0;
      var failed = 0;
      try {
        if (toRename.length === 1) {
          var one = toRename[0];
          var oneOut = makeRenameEditName(one.name);
          setStatus($("editnameStatus"), "Saving " + oneOut + "…");
          await yieldToUI();
          try {
            var buf = await readFileAsArrayBuffer(one);
            downloadBlob(
              new Blob([buf], {
                type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
              }),
              oneOut
            );
            ok = 1;
          } catch (e) {
            console.error(e);
            failed = 1;
          }
        } else {
          if (typeof JSZip === "undefined") {
            throw new Error("JSZip missing.");
          }
          var zip = new JSZip();
          var usedNames = {};
          for (var i = 0; i < toRename.length; i++) {
            var f = toRename[i];
            var outName = makeRenameEditName(f.name);
            var unique = outName;
            var n = 2;
            while (usedNames[unique.toLowerCase()]) {
              unique = outName.replace(/\.pptx$/i, "") + "_" + n + ".pptx";
              n++;
            }
            usedNames[unique.toLowerCase()] = true;
            setStatus(
              $("editnameStatus"),
              "Adding " + (i + 1) + "/" + toRename.length + " — " + unique + "…"
            );
            await yieldToUI();
            try {
              zip.file(unique, await readFileAsArrayBuffer(f));
              ok++;
            } catch (e) {
              console.error(e);
              failed++;
            }
          }
          if (ok) {
            setStatus($("editnameStatus"), "Building zip…");
            await yieldToUI();
            var zipBlob = await zip.generateAsync({
              type: "blob",
              compression: "DEFLATE",
              compressionOptions: { level: 1 },
            });
            downloadBlob(zipBlob, "AMC_EDIT_renamed.zip");
          }
        }
      } catch (e) {
        console.error(e);
        setStatus($("editnameStatus"), "Error: " + (e.message || e), "err");
        $("editnameGo").disabled = false;
        return;
      }

      var parts = [];
      if (ok) {
        parts.push(
          ok === 1
            ? "Downloaded 1 renamed PowerPoint."
            : "Downloaded " + ok + " renamed files in AMC_EDIT_renamed.zip."
        );
      }
      if (skipped) parts.push("Skipped " + skipped + " already named _EDIT.");
      if (failed) parts.push("Failed " + failed + ".");
      setStatus(
        $("editnameStatus"),
        parts.join(" ") || "Nothing downloaded.",
        failed && !ok ? "err" : "ok"
      );
      $("editnameGo").disabled = false;
    });
  }

  // ---- Text formatter ----
  function formatMarkdown(text) {
    var lines = String(text || "").replace(/\r\n/g, "\n").split("\n");
    var html = lines
      .map(function (line) {
        var esc = line
          .replace(/&/g, "&amp;")
          .replace(/</g, "&lt;")
          .replace(/>/g, "&gt;");
        esc = esc.replace(/\*\*(.+?)\*\*/g, "<strong>$1</strong>");
        return esc || " ";
      })
      .join("<br>");
    return html;
  }

  $("fmtGo").addEventListener("click", function () {
    $("fmtOut").innerHTML = formatMarkdown($("fmtIn").value);
    setStatus($("fmtStatus"), "Formatted.", "ok");
  });
  $("fmtCopy").addEventListener("click", async function () {
    var t = $("fmtOut").innerText || "";
    try {
      await navigator.clipboard.writeText(t);
      setStatus($("fmtStatus"), "Copied.", "ok");
    } catch (e) {
      setStatus($("fmtStatus"), "Copy failed — select and copy manually.", "err");
    }
  });
  $("fmtTxt").addEventListener("click", function () {
    var t = $("fmtOut").innerText || $("fmtIn").value || "";
    downloadBlob(new Blob([t], { type: "text/plain" }), "formatted.txt");
    setStatus($("fmtStatus"), "Downloaded formatted.txt", "ok");
  });
  $("fmtClear").addEventListener("click", function () {
    $("fmtIn").value = "";
    $("fmtOut").innerHTML = "";
    setStatus($("fmtStatus"), "Cleared.");
  });
})();

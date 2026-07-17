/**
 * Shared PPTX → slides loader + IndexedDB handoff.
 * Tuned for Safari / iOS / iPadOS as well as desktop browsers.
 */
(function (global) {
  const OPEN_DB = "pptxNotesViewerHandoff";
  const OPEN_STORE = "decks";

  function isAppleMobile() {
    var ua = navigator.userAgent || "";
    var iOS = /iPad|iPhone|iPod/.test(ua);
    // iPadOS 13+ desktop mode reports as Mac with touch
    var iPadOS =
      navigator.platform === "MacIntel" && navigator.maxTouchPoints > 1;
    return iOS || iPadOS;
  }

  function xmlLocalAll(doc, localName) {
    return Array.prototype.filter.call(doc.getElementsByTagName("*"), function (el) {
      return el.localName === localName;
    });
  }

  function escapeHtmlText(s) {
    return String(s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function notesXmlToHtml(xmlText) {
    if (!xmlText) return "";
    var doc = new DOMParser().parseFromString(xmlText, "application/xml");
    var parts = [];
    xmlLocalAll(doc, "p").forEach(function (p) {
      var level = 0;
      Array.prototype.forEach.call(p.childNodes, function (ch) {
        if (ch.nodeType === 1 && ch.localName === "pPr" && ch.getAttribute("lvl")) {
          level = parseInt(ch.getAttribute("lvl"), 10) || 0;
        }
      });
      var lineHtml = "";
      var plain = "";
      var allBold = true;
      var sawText = false;

      function walkRun(rNode) {
        var bold = false;
        var text = "";
        Array.prototype.forEach.call(rNode.childNodes, function (ch) {
          if (ch.nodeType !== 1) return;
          if (ch.localName === "rPr") {
            if (ch.getAttribute("b") === "1" || ch.getAttribute("b") === "true") bold = true;
          }
          if (ch.localName === "t") text += ch.textContent || "";
        });
        if (!text) return;
        sawText = true;
        plain += text;
        var esc = escapeHtmlText(text);
        lineHtml += bold ? "<strong>" + esc + "</strong>" : esc;
        if (!bold && text.trim()) allBold = false;
      }

      function flushLine() {
        var t = plain.trim();
        if (!t && !lineHtml) {
          parts.push('<p class="notes-blank"><br></p>');
        } else if (/^[=─\-_]{6,}$/.test(t)) {
          parts.push('<hr class="notes-sep">');
        } else if (level > 0) {
          parts.push(
            '<p class="notes-bullet" data-level="' +
              level +
              '" style="margin-left:' +
              level * 1.25 +
              'em">' +
              lineHtml +
              "</p>"
          );
        } else if (allBold && sawText && t.length < 80 && !/[.!?]$/.test(t)) {
          parts.push('<h3 class="notes-h"><strong>' + escapeHtmlText(t) + "</strong></h3>");
        } else {
          parts.push("<p>" + lineHtml + "</p>");
        }
        lineHtml = "";
        plain = "";
        allBold = true;
        sawText = false;
      }

      var hasContent = false;
      Array.prototype.forEach.call(p.childNodes, function (ch) {
        if (ch.nodeType !== 1) return;
        if (ch.localName === "r") {
          hasContent = true;
          walkRun(ch);
        } else if (ch.localName === "br") {
          hasContent = true;
          flushLine();
        } else if (ch.localName === "fld") {
          xmlLocalAll(ch, "t").forEach(function (tEl) {
            var text = tEl.textContent || "";
            if (!text) return;
            hasContent = true;
            sawText = true;
            allBold = false;
            plain += text;
            lineHtml += escapeHtmlText(text);
          });
        }
      });
      if (hasContent || plain || lineHtml) flushLine();
      else parts.push('<p class="notes-blank"><br></p>');
    });
    while (parts.length && parts[parts.length - 1] === '<p class="notes-blank"><br></p>') {
      parts.pop();
    }
    return parts.join("\n");
  }

  function relsMap(relsXml) {
    var doc = new DOMParser().parseFromString(relsXml, "application/xml");
    var map = {};
    xmlLocalAll(doc, "Relationship").forEach(function (r) {
      map[r.getAttribute("Id")] = r.getAttribute("Target");
    });
    return map;
  }

  function joinPptPath(baseDir, target) {
    var parts = (baseDir + "/" + target).split("/");
    var out = [];
    parts.forEach(function (p) {
      if (!p || p === ".") return;
      if (p === "..") out.pop();
      else out.push(p);
    });
    return out.join("/");
  }

  function extMime(path) {
    var lower = path.toLowerCase();
    if (lower.endsWith(".png")) return "image/png";
    if (lower.endsWith(".jpg") || lower.endsWith(".jpeg")) return "image/jpeg";
    if (lower.endsWith(".gif")) return "image/gif";
    if (lower.endsWith(".webp")) return "image/webp";
    return "application/octet-stream";
  }

  function blobToDataUrl(blob) {
    return new Promise(function (resolve, reject) {
      var reader = new FileReader();
      reader.onload = function () {
        resolve(reader.result);
      };
      reader.onerror = reject;
      reader.readAsDataURL(blob);
    });
  }

  function safeDeckKey(name) {
    return (
      String(name || "deck")
        .replace(/[<>:"/\\|?*]+/g, "_")
        .replace(/\.pptx$/i, "")
        .trim() || "deck"
    );
  }

  async function pptxFileToSlides(file) {
    if (typeof JSZip === "undefined") {
      throw new Error("JSZip failed to load — check jszip.min.js is online.");
    }
    if (!file) throw new Error("No file selected.");

    var zip;
    try {
      zip = await JSZip.loadAsync(await file.arrayBuffer());
    } catch (e) {
      throw new Error("Could not read that file. Use a .pptx (not .ppt) from Files / iCloud / Drive.");
    }

    var presRelsFile = zip.file("ppt/_rels/presentation.xml.rels");
    var presFile = zip.file("ppt/presentation.xml");
    if (!presRelsFile || !presFile) {
      throw new Error("Not a valid .pptx PowerPoint file.");
    }

    var idToTarget = relsMap(await presRelsFile.async("text"));
    var presDoc = new DOMParser().parseFromString(await presFile.async("text"), "application/xml");
    var R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    var slideTargets = [];
    xmlLocalAll(presDoc, "sldId").forEach(function (el) {
      var rid = el.getAttributeNS(R_NS, "id") || el.getAttribute("r:id");
      if (!rid || !idToTarget[rid]) return;
      var t = String(idToTarget[rid]).replace(/\\/g, "/").replace(/^\//, "");
      if (!t.startsWith("ppt/")) t = "ppt/" + t;
      slideTargets.push(t);
    });

    var slides = [];
    for (var i = 0; i < slideTargets.length; i++) {
      var slidePath = slideTargets[i];
      var slideFile = zip.file(slidePath);
      if (!slideFile) continue;
      var baseDir = slidePath.split("/").slice(0, -1).join("/");
      var relsPath = baseDir + "/_rels/" + slidePath.split("/").pop() + ".rels";
      var relsFile = zip.file(relsPath);
      var rels = relsFile ? relsMap(await relsFile.async("text")) : {};
      var slideDoc = new DOMParser().parseFromString(await slideFile.async("text"), "application/xml");

      var imgDataUrl = "";
      var blips = xmlLocalAll(slideDoc, "blip");
      for (var b = 0; b < blips.length; b++) {
        var embed =
          blips[b].getAttribute("r:embed") || blips[b].getAttributeNS(R_NS, "embed");
        if (!embed || !rels[embed]) continue;
        var mediaPath = joinPptPath(baseDir, rels[embed]);
        var mediaFile = zip.file(mediaPath);
        if (!mediaFile) continue;
        var mime = extMime(mediaPath);
        if (!mime.startsWith("image/")) continue;
        var u8 = await mediaFile.async("uint8array");
        imgDataUrl = await blobToDataUrl(new Blob([u8], { type: mime }));
        break;
      }
      if (!imgDataUrl) {
        imgDataUrl =
          "data:image/svg+xml," +
          encodeURIComponent(
            '<svg xmlns="http://www.w3.org/2000/svg" width="960" height="540">' +
              '<rect width="100%" height="100%" fill="#e0e0e0"/>' +
              '<text x="50%" y="50%" text-anchor="middle" fill="#666" font-size="22" font-family="-apple-system,sans-serif">Slide ' +
              (i + 1) +
              "</text></svg>"
          );
      }

      var notesHtml = "";
      var notesRid = Object.keys(rels).find(function (id) {
        return (rels[id] || "").toLowerCase().indexOf("notesslide") >= 0;
      });
      if (notesRid) {
        var notesPath = joinPptPath(baseDir, rels[notesRid]);
        var notesFile = zip.file(notesPath);
        if (notesFile) notesHtml = notesXmlToHtml(await notesFile.async("text"));
      }

      slides.push({
        n: i + 1,
        img: imgDataUrl,
        notesHtml: notesHtml,
        hasNotes: !!(notesHtml && notesHtml.replace(/<[^>]+>/g, "").trim()),
      });
    }
    if (!slides.length) throw new Error("No slides found in this PowerPoint file.");
    return slides;
  }

  function idbOpen() {
    return new Promise(function (resolve, reject) {
      var req = indexedDB.open(OPEN_DB, 1);
      req.onupgradeneeded = function () {
        var db = req.result;
        if (!db.objectStoreNames.contains(OPEN_STORE)) db.createObjectStore(OPEN_STORE);
      };
      req.onsuccess = function () {
        resolve(req.result);
      };
      req.onerror = function () {
        reject(req.error || new Error("IndexedDB unavailable (Private Browsing?)"));
      };
    });
  }

  async function handoffPut(id, payload) {
    var db = await idbOpen();
    return new Promise(function (resolve, reject) {
      var tx = db.transaction(OPEN_STORE, "readwrite");
      tx.objectStore(OPEN_STORE).put(payload, id);
      tx.oncomplete = function () {
        resolve();
      };
      tx.onerror = function () {
        reject(tx.error);
      };
    });
  }

  async function handoffTake(id) {
    var db = await idbOpen();
    return new Promise(function (resolve, reject) {
      var tx = db.transaction(OPEN_STORE, "readwrite");
      var store = tx.objectStore(OPEN_STORE);
      var getReq = store.get(id);
      getReq.onsuccess = function () {
        var val = getReq.result;
        try {
          store.delete(id);
        } catch (e) {}
        resolve(val || null);
      };
      getReq.onerror = function () {
        reject(getReq.error);
      };
    });
  }

  function viewerUrlWithOpen(viewerUrl, handoffId) {
    var base = viewerUrl || "viewer.html";
    var url = new URL(base, location.href);
    url.searchParams.set("open", handoffId);
    return url.toString();
  }

  /**
   * Open a blank tab/window during the user tap (Safari-safe), before any await.
   * Returns a Window or null.
   */
  function openPlaceholderTab() {
    try {
      // Prefer a real page shell so iOS doesn’t discard about:blank as easily
      var w = window.open("about:blank", "_blank");
      if (w) {
        try {
          w.document.write(
            "<!DOCTYPE html><title>Loading…</title><body style='font-family:-apple-system,sans-serif;padding:24px;color:#333'>" +
              "Loading presentation…</body>"
          );
          w.document.close();
        } catch (e) {}
      }
      return w;
    } catch (e) {
      return null;
    }
  }

  /**
   * @param {File} file
   * @param {string} [viewerUrl]
   * @param {{ preOpenedWindow?: Window|null, sameTabFallback?: boolean }} [opts]
   */
  async function openPptxInNewTab(file, viewerUrl, opts) {
    opts = opts || {};
    var pre = opts.preOpenedWindow || null;
    var sameTabFallback = opts.sameTabFallback !== false;

    // On Apple mobile, new tabs after async often fail — same-tab is more reliable
    var preferSameTab = isAppleMobile() && opts.forceNewTab !== true;

    var slides = await pptxFileToSlides(file);
    var deckKey = safeDeckKey(file.name);
    var handoffId = String(Date.now()) + "_" + Math.random().toString(36).slice(2, 8);

    try {
      await handoffPut(handoffId, {
        deckKey: deckKey,
        title: deckKey,
        slides: slides,
        createdAt: Date.now(),
      });
    } catch (e) {
      if (pre && !pre.closed) {
        try {
          pre.close();
        } catch (err) {}
      }
      throw new Error(
        "Could not store the presentation (storage full or Private Browsing). Try a smaller file or leave Private mode."
      );
    }

    var url = viewerUrlWithOpen(viewerUrl, handoffId);

    if (preferSameTab) {
      if (pre && !pre.closed) {
        try {
          pre.close();
        } catch (e) {}
      }
      location.href = url;
      return { deckKey: deckKey, slides: slides.length, mode: "same-tab" };
    }

    if (pre && !pre.closed) {
      try {
        pre.location.href = url;
        return { deckKey: deckKey, slides: slides.length, mode: "new-tab" };
      } catch (e) {
        /* fall through */
      }
    }

    var w = null;
    try {
      w = window.open(url, "_blank");
    } catch (e) {
      w = null;
    }

    if (w) {
      return { deckKey: deckKey, slides: slides.length, mode: "new-tab" };
    }

    if (sameTabFallback) {
      location.href = url;
      return { deckKey: deckKey, slides: slides.length, mode: "same-tab" };
    }

    throw new Error("Pop-up blocked. Allow pop-ups, or open the file again to load in this tab.");
  }

  global.PptxNotesLoader = {
    pptxFileToSlides: pptxFileToSlides,
    openPptxInNewTab: openPptxInNewTab,
    openPlaceholderTab: openPlaceholderTab,
    handoffTake: handoffTake,
    handoffPut: handoffPut,
    safeDeckKey: safeDeckKey,
    isAppleMobile: isAppleMobile,
  };
})(typeof window !== "undefined" ? window : globalThis);

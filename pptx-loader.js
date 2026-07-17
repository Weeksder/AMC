/**
 * Shared PPTX → slides loader + IndexedDB handoff for multi-tab open.
 */
(function (global) {
  const OPEN_DB = "pptxNotesViewerHandoff";
  const OPEN_STORE = "decks";

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
    const doc = new DOMParser().parseFromString(xmlText, "application/xml");
    const parts = [];
    xmlLocalAll(doc, "p").forEach(function (p) {
      let level = 0;
      Array.prototype.forEach.call(p.childNodes, function (ch) {
        if (ch.nodeType === 1 && ch.localName === "pPr" && ch.getAttribute("lvl")) {
          level = parseInt(ch.getAttribute("lvl"), 10) || 0;
        }
      });
      let lineHtml = "";
      let plain = "";
      let allBold = true;
      let sawText = false;

      function walkRun(rNode) {
        let bold = false;
        let text = "";
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
        const esc = escapeHtmlText(text);
        lineHtml += bold ? "<strong>" + esc + "</strong>" : esc;
        if (!bold && text.trim()) allBold = false;
      }

      function flushLine() {
        const t = plain.trim();
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

      let hasContent = false;
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
            const text = tEl.textContent || "";
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
    const doc = new DOMParser().parseFromString(relsXml, "application/xml");
    const map = {};
    xmlLocalAll(doc, "Relationship").forEach(function (r) {
      map[r.getAttribute("Id")] = r.getAttribute("Target");
    });
    return map;
  }

  function joinPptPath(baseDir, target) {
    const parts = (baseDir + "/" + target).split("/");
    const out = [];
    parts.forEach(function (p) {
      if (!p || p === ".") return;
      if (p === "..") out.pop();
      else out.push(p);
    });
    return out.join("/");
  }

  function extMime(path) {
    const lower = path.toLowerCase();
    if (lower.endsWith(".png")) return "image/png";
    if (lower.endsWith(".jpg") || lower.endsWith(".jpeg")) return "image/jpeg";
    if (lower.endsWith(".gif")) return "image/gif";
    if (lower.endsWith(".webp")) return "image/webp";
    return "application/octet-stream";
  }

  function blobToDataUrl(blob) {
    return new Promise(function (resolve, reject) {
      const reader = new FileReader();
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
      throw new Error("JSZip failed to load");
    }
    const zip = await JSZip.loadAsync(await file.arrayBuffer());
    const presRelsFile = zip.file("ppt/_rels/presentation.xml.rels");
    const presFile = zip.file("ppt/presentation.xml");
    if (!presRelsFile || !presFile) throw new Error("Not a valid .pptx file");

    const idToTarget = relsMap(await presRelsFile.async("text"));
    const presDoc = new DOMParser().parseFromString(await presFile.async("text"), "application/xml");
    const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const slideTargets = [];
    xmlLocalAll(presDoc, "sldId").forEach(function (el) {
      const rid = el.getAttributeNS(R_NS, "id") || el.getAttribute("r:id");
      if (!rid || !idToTarget[rid]) return;
      let t = String(idToTarget[rid]).replace(/\\/g, "/").replace(/^\//, "");
      if (!t.startsWith("ppt/")) t = "ppt/" + t;
      slideTargets.push(t);
    });

    const slides = [];
    for (let i = 0; i < slideTargets.length; i++) {
      const slidePath = slideTargets[i];
      const slideFile = zip.file(slidePath);
      if (!slideFile) continue;
      const baseDir = slidePath.split("/").slice(0, -1).join("/");
      const relsPath = baseDir + "/_rels/" + slidePath.split("/").pop() + ".rels";
      const relsFile = zip.file(relsPath);
      const rels = relsFile ? relsMap(await relsFile.async("text")) : {};
      const slideDoc = new DOMParser().parseFromString(await slideFile.async("text"), "application/xml");

      let imgDataUrl = "";
      const blips = xmlLocalAll(slideDoc, "blip");
      for (let b = 0; b < blips.length; b++) {
        const embed =
          blips[b].getAttribute("r:embed") ||
          blips[b].getAttributeNS(R_NS, "embed");
        if (!embed || !rels[embed]) continue;
        const mediaPath = joinPptPath(baseDir, rels[embed]);
        const mediaFile = zip.file(mediaPath);
        if (!mediaFile) continue;
        const mime = extMime(mediaPath);
        if (!mime.startsWith("image/")) continue;
        const u8 = await mediaFile.async("uint8array");
        imgDataUrl = await blobToDataUrl(new Blob([u8], { type: mime }));
        break;
      }
      if (!imgDataUrl) {
        imgDataUrl =
          "data:image/svg+xml," +
          encodeURIComponent(
            '<svg xmlns="http://www.w3.org/2000/svg" width="960" height="540">' +
              '<rect width="100%" height="100%" fill="#e0e0e0"/>' +
              '<text x="50%" y="50%" text-anchor="middle" fill="#666" font-size="22" font-family="Segoe UI,sans-serif">Slide ' +
              (i + 1) +
              "</text></svg>"
          );
      }

      let notesHtml = "";
      const notesRid = Object.keys(rels).find(function (id) {
        return (rels[id] || "").toLowerCase().indexOf("notesslide") >= 0;
      });
      if (notesRid) {
        const notesPath = joinPptPath(baseDir, rels[notesRid]);
        const notesFile = zip.file(notesPath);
        if (notesFile) notesHtml = notesXmlToHtml(await notesFile.async("text"));
      }

      slides.push({
        n: i + 1,
        img: imgDataUrl,
        notesHtml: notesHtml,
        hasNotes: !!(notesHtml && notesHtml.replace(/<[^>]+>/g, "").trim()),
      });
    }
    if (!slides.length) throw new Error("No slides found in this PowerPoint file");
    return slides;
  }

  function idbOpen() {
    return new Promise(function (resolve, reject) {
      const req = indexedDB.open(OPEN_DB, 1);
      req.onupgradeneeded = function () {
        const db = req.result;
        if (!db.objectStoreNames.contains(OPEN_STORE)) db.createObjectStore(OPEN_STORE);
      };
      req.onsuccess = function () {
        resolve(req.result);
      };
      req.onerror = function () {
        reject(req.error);
      };
    });
  }

  async function handoffPut(id, payload) {
    const db = await idbOpen();
    return new Promise(function (resolve, reject) {
      const tx = db.transaction(OPEN_STORE, "readwrite");
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
    const db = await idbOpen();
    return new Promise(function (resolve, reject) {
      const tx = db.transaction(OPEN_STORE, "readwrite");
      const store = tx.objectStore(OPEN_STORE);
      const getReq = store.get(id);
      getReq.onsuccess = function () {
        const val = getReq.result;
        store.delete(id);
        resolve(val || null);
      };
      getReq.onerror = function () {
        reject(getReq.error);
      };
    });
  }

  /**
   * Load a PPTX File, store handoff, open viewer in a new tab.
   * @param {File} file
   * @param {string} [viewerUrl] default viewer.html relative to current page
   */
  async function openPptxInNewTab(file, viewerUrl) {
    const slides = await pptxFileToSlides(file);
    const deckKey = safeDeckKey(file.name);
    const handoffId = String(Date.now()) + "_" + Math.random().toString(36).slice(2, 8);
    await handoffPut(handoffId, {
      deckKey: deckKey,
      title: deckKey,
      slides: slides,
      createdAt: Date.now(),
    });
    const base = viewerUrl || "viewer.html";
    const url = new URL(base, location.href);
    url.searchParams.set("open", handoffId);
    const w = window.open(url.toString(), "_blank");
    if (!w) {
      throw new Error("Pop-up blocked. Allow pop-ups for this site, then try again.");
    }
    return { deckKey: deckKey, slides: slides.length };
  }

  global.PptxNotesLoader = {
    pptxFileToSlides: pptxFileToSlides,
    openPptxInNewTab: openPptxInNewTab,
    handoffTake: handoffTake,
    handoffPut: handoffPut,
    safeDeckKey: safeDeckKey,
  };
})(typeof window !== "undefined" ? window : globalThis);

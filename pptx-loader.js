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
        } else {
          // Desktop-style: always plain paragraphs (bold stays as <strong>, not headings)
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
    // Collapse consecutive blank lines → max one (removes extra vertical gaps)
    var collapsed = [];
    parts.forEach(function (p) {
      var isBlank = p.indexOf("notes-blank") >= 0;
      if (isBlank && collapsed.length && collapsed[collapsed.length - 1].indexOf("notes-blank") >= 0) {
        return;
      }
      collapsed.push(p);
    });
    while (
      collapsed.length &&
      collapsed[collapsed.length - 1].indexOf("notes-blank") >= 0
    ) {
      collapsed.pop();
    }
    return collapsed.join("\n");
  }

  /**
   * Normalize pasted ChatGPT / Docs / Slides HTML into simple desktop-style
   * notes: one visual line → one <p>, keep only bold/italic/underline.
   */
  function normalizeNotesPasteHtml(html, plainText) {
    // Prefer structured HTML when present; fall back to plain text lines
    if (html && /<[a-z][\s\S]*>/i.test(html)) {
      try {
        var doc = new DOMParser().parseFromString(html, "text/html");
        var body = doc.body;
        // Strip Google/Office junk wrappers
        body.querySelectorAll("style, meta, link, script, xml, o\\:p").forEach(function (el) {
          el.remove();
        });
        var lines = [];
        var cur = [];

        function flush() {
          lines.push(cur);
          cur = [];
        }
        function pushText(t, bold, italic, under) {
          if (!t) return;
          var parts = String(t).split(/\r\n|\n|\r/);
          for (var i = 0; i < parts.length; i++) {
            if (i > 0) flush();
            if (parts[i].length) {
              cur.push({ t: parts[i], b: !!bold, i: !!italic, u: !!under });
            }
          }
        }
        function walk(n, bold, italic, under) {
          if (n.nodeType === 3) {
            pushText(n.nodeValue, bold, italic, under);
            return;
          }
          if (n.nodeType !== 1) return;
          var tag = n.tagName.toLowerCase();
          if (tag === "br") {
            flush();
            return;
          }
          if (tag === "style" || tag === "script" || tag === "meta") return;
          var isBlock =
            tag === "p" ||
            tag === "div" ||
            tag === "li" ||
            tag === "h1" ||
            tag === "h2" ||
            tag === "h3" ||
            tag === "h4" ||
            tag === "tr" ||
            tag === "blockquote";
          if (isBlock && cur.length) flush();
          var b = bold || tag === "b" || tag === "strong";
          var it = italic || tag === "i" || tag === "em";
          var u = under || tag === "u";
          // Inline style font-weight bold
          try {
            var fw = n.style && n.style.fontWeight;
            if (fw === "bold" || parseInt(fw, 10) >= 600) b = true;
          } catch (e) {}
          Array.prototype.forEach.call(n.childNodes, function (ch) {
            walk(ch, b, it, u);
          });
          if (isBlock) flush();
        }
        Array.prototype.forEach.call(body.childNodes, function (ch) {
          walk(ch, false, false, false);
        });
        if (cur.length) flush();
        while (lines.length && lines[lines.length - 1].length === 0) lines.pop();

        if (lines.length) {
          return lines
            .map(function (runs) {
              if (!runs.length) return "<p><br></p>";
              var inner = runs
                .map(function (r) {
                  var esc = escapeHtmlText(r.t);
                  if (r.b) esc = "<strong>" + esc + "</strong>";
                  if (r.i) esc = "<em>" + esc + "</em>";
                  if (r.u) esc = "<u>" + esc + "</u>";
                  return esc;
                })
                .join("");
              return "<p>" + inner + "</p>";
            })
            .join("");
        }
      } catch (e) {
        /* fall through to plain */
      }
    }

    var text = plainText != null ? String(plainText) : "";
    if (!text.trim()) return "<p><br></p>";
    return text
      .replace(/\r\n/g, "\n")
      .replace(/\r/g, "\n")
      .split("\n")
      .map(function (line) {
        if (!line.length) return "<p><br></p>";
        return "<p>" + escapeHtmlText(line) + "</p>";
      })
      .join("");
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

  function firstLocal(el, name) {
    if (!el) return null;
    for (var i = 0; i < el.childNodes.length; i++) {
      var ch = el.childNodes[i];
      if (ch.nodeType === 1 && ch.localName === name) return ch;
    }
    return null;
  }

  function attrInt(el, name, fallback) {
    if (!el) return fallback;
    var v = el.getAttribute(name);
    if (v == null || v === "") return fallback;
    var n = parseInt(v, 10);
    return isNaN(n) ? fallback : n;
  }

  function ptToEmu(pt) {
    return pt * 12700;
  }

  function szToEmu(szHundredths, fallbackPt) {
    var sz = parseInt(szHundredths, 10);
    if (!sz || sz <= 0) return ptToEmu(fallbackPt == null ? 18 : fallbackPt);
    return ptToEmu(sz / 100);
  }

  function isChromePlaceholderShape(sp) {
    var phs = xmlLocalAll(sp, "ph");
    for (var i = 0; i < phs.length; i++) {
      var type = String(phs[i].getAttribute("type") || "").toLowerCase();
      if (type === "ftr" || type === "sldnum" || type === "dt" || type === "hdr") return true;
    }
    return false;
  }

  function shapeXfrm(sp) {
    var spPr = firstLocal(sp, "spPr") || xmlLocalAll(sp, "spPr")[0];
    var xfrm = spPr ? firstLocal(spPr, "xfrm") || xmlLocalAll(spPr, "xfrm")[0] : null;
    var off = xfrm ? firstLocal(xfrm, "off") || xmlLocalAll(xfrm, "off")[0] : null;
    var ext = xfrm ? firstLocal(xfrm, "ext") || xmlLocalAll(xfrm, "ext")[0] : null;
    return {
      x: attrInt(off, "x", 0),
      y: attrInt(off, "y", 0),
      cx: attrInt(ext, "cx", 0),
      cy: attrInt(ext, "cy", 0),
    };
  }

  function svgToDataUrl(svg) {
    return "data:image/svg+xml;charset=utf-8," + encodeURIComponent(svg);
  }

  function runFontName(rPr) {
    if (!rPr) return "Calibri";
    var latin = xmlLocalAll(rPr, "latin")[0];
    var face = latin && latin.getAttribute("typeface");
    if (face && face.charAt(0) !== "+") return face;
    return "Calibri";
  }

  function runIsBold(rPr) {
    if (!rPr) return false;
    var b = rPr.getAttribute("b");
    return b === "1" || b === "true";
  }

  function runFillColor(rPr) {
    if (!rPr) return "#111111";
    var srgb = xmlLocalAll(rPr, "srgbClr")[0];
    if (srgb && srgb.getAttribute("val")) {
      var hex = String(srgb.getAttribute("val")).replace("#", "");
      if (/^[0-9a-fA-F]{6}$/.test(hex)) {
        var r = parseInt(hex.slice(0, 2), 16);
        var g = parseInt(hex.slice(2, 4), 16);
        var b = parseInt(hex.slice(4, 6), 16);
        if (r + g + b >= 720) return "#111111";
        return "#" + hex;
      }
    }
    return "#111111";
  }

  function estimateTextWidthEmu(text, fontEmu) {
    return String(text || "").length * fontEmu * 0.5;
  }

  function wrapWordsToWidth(text, fontEmu, maxW) {
    var words = String(text || "").split(/\s+/).filter(Boolean);
    if (!words.length) return [""];
    var lines = [];
    var cur = "";
    words.forEach(function (w) {
      var next = cur ? cur + " " + w : w;
      if (cur && estimateTextWidthEmu(next, fontEmu) > maxW) {
        lines.push(cur);
        cur = w;
      } else {
        cur = next;
      }
    });
    if (cur) lines.push(cur);
    return lines;
  }

  function collectShapeRuns(pEl, fallbackSz) {
    var runs = [];
    Array.prototype.forEach.call(pEl.childNodes, function (ch) {
      if (ch.nodeType !== 1) return;
      if (ch.localName === "br") {
        runs.push({ text: "\n", szEmu: fallbackSz, bold: false, font: "Calibri", color: "#111111" });
        return;
      }
      if (ch.localName !== "r" && ch.localName !== "fld") return;
      var rPr = firstLocal(ch, "rPr");
      var szEmu = rPr && rPr.getAttribute("sz") ? szToEmu(rPr.getAttribute("sz"), 18) : fallbackSz;
      var text = "";
      xmlLocalAll(ch, "t").forEach(function (tEl) {
        text += tEl.textContent || "";
      });
      if (!text) return;
      runs.push({
        text: text,
        szEmu: szEmu,
        bold: runIsBold(rPr),
        font: runFontName(rPr),
        color: runFillColor(rPr),
      });
    });
    return runs;
  }

  function buildTextSlideFromDoc(slideDoc, slideCx, slideCy) {
    slideCx = slideCx || 12192000;
    slideCy = slideCy || 6858000;
    var svgW = 960;
    var svgH = Math.max(1, Math.round((960 * slideCy) / slideCx));
    var k = svgW / slideCx;
    function px(emu) {
      return emu * k;
    }
    var parts = [];
    var htmlParts = [];
    function textEl(x, y, fontPx, anchorSvg, color, bold, content) {
      var fs = Math.max(12, Math.round(fontPx));
      var fill = color || "#111111";
      var weight = bold ? "700" : "400";
      parts.push(
        '<text x="' +
          Math.round(x) +
          '" y="' +
          Math.round(y) +
          '" text-anchor="' +
          anchorSvg +
          '" fill="' +
          fill +
          '" font-size="' +
          fs +
          '" font-family="Calibri, Arial, sans-serif" font-weight="' +
          weight +
          '">' +
          escapeHtmlText(content) +
          "</text>"
      );
      var trans =
        anchorSvg === "middle"
          ? "translate(-50%,-80%)"
          : anchorSvg === "end"
            ? "translate(-100%,-80%)"
            : "translate(0,-80%)";
      htmlParts.push(
        '<div style="position:absolute;left:' +
          Math.round(x) +
          "px;top:" +
          Math.round(y) +
          "px;font-size:" +
          fs +
          "px;line-height:1;color:" +
          fill +
          ";font-family:Calibri,Arial,sans-serif;font-weight:" +
          weight +
          ";white-space:nowrap;transform:" +
          trans +
          '">' +
          escapeHtmlText(content) +
          "</div>"
      );
    }
    xmlLocalAll(slideDoc, "sp").forEach(function (sp) {
      if (isChromePlaceholderShape(sp)) return;
      var txBody = firstLocal(sp, "txBody") || xmlLocalAll(sp, "txBody")[0];
      if (!txBody) return;
      var box = shapeXfrm(sp);
      if (box.cx <= 0 || box.cy <= 0) return;
      var bodyPr = firstLocal(txBody, "bodyPr") || xmlLocalAll(txBody, "bodyPr")[0];
      var lIns = attrInt(bodyPr, "lIns", 91440);
      var tIns = attrInt(bodyPr, "tIns", 45720);
      var rIns = attrInt(bodyPr, "rIns", 91440);
      var bIns = attrInt(bodyPr, "bIns", 45720);
      var innerW = Math.max(1, box.cx - lIns - rIns);
      var innerH = Math.max(1, box.cy - tIns - bIns);
      var anchor = String((bodyPr && bodyPr.getAttribute("anchor")) || "t").toLowerCase();
      var wrap = String((bodyPr && bodyPr.getAttribute("wrap")) || "square").toLowerCase();
      var autoFit = !!(bodyPr && xmlLocalAll(bodyPr, "spAutoFit").length);

      var paras = [];
      xmlLocalAll(txBody, "p").forEach(function (pEl) {
        var pPr = firstLocal(pEl, "pPr");
        var algn = String((pPr && pPr.getAttribute("algn")) || "l").toLowerCase();
        var endPr = firstLocal(pEl, "endParaRPr");
        var fallbackSz =
          endPr && endPr.getAttribute("sz") ? szToEmu(endPr.getAttribute("sz"), 18) : ptToEmu(18);
        var runs = collectShapeRuns(pEl, fallbackSz);
        var text = runs
          .map(function (r) {
            return r.text;
          })
          .join("");
        if (!String(text).replace(/\s+/g, "").length) return;
        var maxSz = fallbackSz;
        runs.forEach(function (r) {
          if (r.szEmu > maxSz) maxSz = r.szEmu;
        });
        paras.push({ algn: algn, runs: runs, text: text, maxSz: maxSz });
      });
      if (!paras.length) return;

      var fit = 1;
      if (autoFit) {
        paras.forEach(function (para) {
          var w = estimateTextWidthEmu(para.text.replace(/\n/g, ""), para.maxSz);
          if (w > innerW) fit = Math.min(fit, innerW / w);
        });
        var lineCount = 0;
        paras.forEach(function (para) {
          lineCount += Math.max(1, para.text.split("\n").length);
        });
        var blockH = lineCount * paras[0].maxSz * 1.2 * fit;
        if (blockH > innerH) fit = Math.min(fit, innerH / (lineCount * paras[0].maxSz * 1.2));
        if (fit < 0.2) fit = 0.2;
        if (fit > 1) fit = 1;
      }

      var laid = [];
      paras.forEach(function (para) {
        var fontEmu = para.maxSz * fit;
        var chunks = para.text.split("\n");
        chunks.forEach(function (chunk) {
          var lines =
            wrap !== "none" ? wrapWordsToWidth(chunk, fontEmu, innerW) : [chunk || ""];
          lines.forEach(function (line) {
            if (!String(line).trim()) return;
            laid.push({
              text: line,
              algn: para.algn,
              fontEmu: fontEmu,
              bold: !!(para.runs[0] && para.runs[0].bold),
              color: (para.runs[0] && para.runs[0].color) || "#111111",
            });
          });
        });
      });
      if (!laid.length) return;

      var lineH = laid[0].fontEmu * 1.2;
      var blockH2 = laid.length * lineH;
      var startY;
      if (anchor === "ctr") startY = box.y + tIns + (innerH - blockH2) / 2;
      else if (anchor === "b") startY = box.y + tIns + innerH - blockH2;
      else startY = box.y + tIns;

      laid.forEach(function (line, i) {
        var baseline = startY + i * lineH + line.fontEmu * 0.8;
        var x;
        var anchorSvg = "start";
        if (line.algn === "ctr") {
          x = box.x + lIns + innerW / 2;
          anchorSvg = "middle";
        } else if (line.algn === "r") {
          x = box.x + lIns + innerW;
          anchorSvg = "end";
        } else {
          x = box.x + lIns;
        }
        parts.push(
          textEl(px(x), px(baseline), px(line.fontEmu), anchorSvg, line.color, line.bold, line.text)
        );
      });
    });

    if (!parts.length) {
      var lines = [];
      var maxEmu = ptToEmu(44);
      xmlLocalAll(slideDoc, "t").forEach(function (tEl) {
        var t = String(tEl.textContent || "").replace(/\s+/g, " ").trim();
        if (!t) return;
        var sp = tEl.parentNode;
        while (sp && sp.nodeType === 1 && sp.localName !== "sp") sp = sp.parentNode;
        if (sp && isChromePlaceholderShape(sp)) return;
        lines.push(t);
      });
      xmlLocalAll(slideDoc, "rPr").forEach(function (rPr) {
        if (rPr.getAttribute("sz")) {
          var e = szToEmu(rPr.getAttribute("sz"), 18);
          if (e > maxEmu) maxEmu = e;
        }
      });
      var fontPx = Math.max(28, px(maxEmu));
      var linePx = fontPx * 1.25;
      var top = svgH / 2 - ((lines.length - 1) * linePx) / 2 + fontPx * 0.3;
      lines.forEach(function (line, i) {
        parts.push(textEl(svgW / 2, top + i * linePx, fontPx, "middle", "#111111", false, line));
      });
    }

    var svg =
      '<svg xmlns="http://www.w3.org/2000/svg" width="' +
      svgW +
      '" height="' +
      svgH +
      '" viewBox="0 0 ' +
      svgW +
      " " +
      svgH +
      '" overflow="visible">' +
      '<rect width="100%" height="100%" fill="#ffffff"/>' +
      parts.join("") +
      "</svg>";
    var html =
      '<div class="text-slide-inner" data-w="' +
      svgW +
      '" data-h="' +
      svgH +
      '" style="width:' +
      svgW +
      "px;height:" +
      svgH +
      'px;background:#fff;position:relative;overflow:hidden">' +
      htmlParts.join("") +
      "</div>";
    return { img: svgToDataUrl(svg), html: html };
  }

  function safeDeckKey(name) {
    return (
      String(name || "deck")
        .replace(/[<>:"/\\|?*]+/g, "_")
        .replace(/\.pptx$/i, "")
        .replace(/\.pdf$/i, "")
        .trim() || "deck"
    );
  }

  /** Fingerprint so different files with the same name do not share notes cache. */
  function bufferFingerprint(buf) {
    if (!buf) return "0";
    var u8 =
      buf instanceof ArrayBuffer
        ? new Uint8Array(buf)
        : buf instanceof Uint8Array
          ? buf
          : new Uint8Array(0);
    var len = u8.byteLength;
    if (!len) return "0";
    var h = 2166136261;
    var step = Math.max(1, Math.floor(len / 160));
    for (var i = 0; i < len; i += step) {
      h ^= u8[i];
      h = Math.imul(h, 16777619);
    }
    h ^= u8[len - 1];
    h = Math.imul(h, 16777619);
    return len + "x" + (h >>> 0).toString(16);
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
    var sldSzEl = xmlLocalAll(presDoc, "sldSz")[0];
    var slideCx = attrInt(sldSzEl, "cx", 12192000);
    var slideCy = attrInt(sldSzEl, "cy", 6858000);
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
      var textHtml = "";
      if (!imgDataUrl) {
        var textSlide = buildTextSlideFromDoc(slideDoc, slideCx, slideCy);
        imgDataUrl = textSlide.img;
        textHtml = textSlide.html || "";
      }

      var notesHtml = "";
      var originalNotesXml = null;
      var notesRid = Object.keys(rels).find(function (id) {
        return (rels[id] || "").toLowerCase().indexOf("notesslide") >= 0;
      });
      if (notesRid) {
        var notesPath = joinPptPath(baseDir, rels[notesRid]);
        var notesFile = zip.file(notesPath);
        if (notesFile) {
          originalNotesXml = await notesFile.async("text");
          notesHtml = notesXmlToHtml(originalNotesXml);
        }
      }

      slides.push({
        n: i + 1,
        img: imgDataUrl,
        textHtml: textHtml,
        notesHtml: notesHtml,
        // Keep exact OOXML so SAVE can write the file back unchanged if not edited
        originalNotesXml: originalNotesXml,
        notesDirty: false,
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
    // Keep original bytes so SAVE can download a real .pptx with notes (no JSON step)
    var pptxBuffer = await file.arrayBuffer();
    var contentId = bufferFingerprint(pptxBuffer);

    try {
      await handoffPut(handoffId, {
        deckKey: deckKey,
        title: deckKey,
        slides: slides,
        pptxBuffer: pptxBuffer,
        fileName: file.name || deckKey + ".pptx",
        contentId: contentId,
        // Fresh open from Extract / Open — do not rehydrate old localStorage notes
        freshOpen: true,
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

  function escapeXml(s) {
    return String(s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  /**
   * notesHtml → OOXML paragraphs for a:txBody
   * Desktop-style: ONE visual line = ONE <a:p> (not soft a:br inside a paragraph).
   * That matches classic PowerPoint notes, not Google Slides / Docs paste.
   */
  function notesHtmlToOoxmlParagraphs(notesHtml) {
    if (!notesHtml || !String(notesHtml).replace(/<[^>]+>/g, "").trim()) {
      return '<a:p><a:endParaRPr lang="en-US" dirty="0"/></a:p>';
    }
    if (typeof document === "undefined") {
      // Fallback plain text (no DOM)
      return String(notesHtml)
        .replace(/<[^>]+>/g, "\n")
        .split(/\r\n|\n|\r/)
        .map(function (line) {
          line = line.replace(/\s+/g, " ").trim();
          if (!line) {
            return '<a:p><a:br><a:rPr lang="en-US" dirty="0"/></a:br><a:endParaRPr lang="en-US" dirty="0"/></a:p>';
          }
          return (
            '<a:p><a:r><a:rPr lang="en-US" dirty="0"/><a:t>' +
            escapeXml(line) +
            "</a:t></a:r></a:p>"
          );
        })
        .join("");
    }

    var wrap = document.createElement("div");
    wrap.innerHTML = notesHtml;

    // lines = array of run arrays: [{text, bold, italic, under}, ...]
    var lines = [];
    var current = [];

    function flushLine() {
      lines.push(current);
      current = [];
    }

    function pushText(t, bold, italic, under) {
      if (t == null || t === "") return;
      // Real newlines in text (e.g. ChatGPT paste) → new paragraphs
      var chunks = String(t).split(/\r\n|\n|\r/);
      for (var i = 0; i < chunks.length; i++) {
        if (i > 0) flushLine();
        if (chunks[i].length) {
          current.push({
            text: chunks[i],
            bold: !!bold,
            italic: !!italic,
            under: !!under,
          });
        }
      }
    }

    function walk(n, bold, italic, under) {
      if (n.nodeType === 3) {
        pushText(n.nodeValue, bold, italic, under);
        return;
      }
      if (n.nodeType !== 1) return;
      var tag = n.tagName.toLowerCase();
      if (tag === "br") {
        flushLine();
        return;
      }
      if (tag === "hr") {
        if (current.length) flushLine();
        current.push({
          text: "===============================================",
          bold: false,
          italic: false,
          under: false,
        });
        flushLine();
        return;
      }
      // Skip empty-notes placeholder
      if (tag === "p" && n.classList && n.classList.contains("empty-notes")) {
        return;
      }
      var isBlock =
        tag === "p" ||
        tag === "div" ||
        tag === "li" ||
        tag === "h1" ||
        tag === "h2" ||
        tag === "h3" ||
        tag === "h4" ||
        tag === "tr" ||
        tag === "blockquote";
      if (isBlock && current.length) flushLine();
      var b =
        bold ||
        tag === "strong" ||
        tag === "b" ||
        tag === "h1" ||
        tag === "h2" ||
        tag === "h3" ||
        tag === "h4";
      var it = italic || tag === "em" || tag === "i";
      var u = under || tag === "u";
      Array.prototype.forEach.call(n.childNodes, function (ch) {
        walk(ch, b, it, u);
      });
      if (isBlock) flushLine();
    }

    Array.prototype.forEach.call(wrap.childNodes, function (ch) {
      walk(ch, false, false, false);
    });
    if (current.length) flushLine();

    // Drop trailing blank lines
    while (lines.length && lines[lines.length - 1].length === 0) lines.pop();

    if (!lines.length) {
      return '<a:p><a:endParaRPr lang="en-US" dirty="0"/></a:p>';
    }

    return lines
      .map(function (runs) {
        // Blank line — same pattern as desktop PowerPoint notes
        if (!runs.length) {
          return (
            '<a:p><a:br><a:rPr lang="en-US" dirty="0"/></a:br>' +
            '<a:endParaRPr lang="en-US" dirty="0"/></a:p>'
          );
        }
        var inner = runs
          .map(function (r) {
            var rPr = ' lang="en-US"';
            if (r.bold) rPr += ' b="1"';
            if (r.italic) rPr += ' i="1"';
            if (r.under) rPr += ' u="sng"';
            rPr += ' dirty="0"';
            // Preserve leading/trailing spaces in a run (ChatGPT bold mid-sentence)
            var space =
              /^\s|\s$/.test(r.text) ? ' xml:space="preserve"' : "";
            return (
              "<a:r><a:rPr" +
              rPr +
              "/><a:t" +
              space +
              ">" +
              escapeXml(r.text) +
              "</a:t></a:r>"
            );
          })
          .join("");
        // No empty <a:pPr/> — keeps notes looking like desktop, not Google Slides
        return "<a:p>" + inner + "</a:p>";
      })
      .join("");
  }

  function findNotesTxBody(notesDoc) {
    var bodies = xmlLocalAll(notesDoc, "txBody");
    if (!bodies.length) return null;
    // Prefer placeholder body (notes text), not the slide thumbnail
    for (var i = 0; i < bodies.length; i++) {
      var sp = bodies[i].parentNode;
      while (sp && sp.localName !== "sp") sp = sp.parentNode;
      if (!sp) continue;
      var phs = xmlLocalAll(sp, "ph");
      for (var j = 0; j < phs.length; j++) {
        var typ = (phs[j].getAttribute("type") || "").toLowerCase();
        if (typ === "body" || typ.indexOf("notes") >= 0) return bodies[i];
      }
      var names = xmlLocalAll(sp, "cNvPr");
      for (var k = 0; k < names.length; k++) {
        var nm = (names[k].getAttribute("name") || "").toLowerCase();
        if (nm.indexOf("notes") >= 0) return bodies[i];
      }
    }
    // Fallback: last txBody (usually notes text)
    return bodies[bodies.length - 1];
  }

  function setTxBodyParagraphs(txBody, paragraphsXml) {
    // Keep a:bodyPr / a:lstStyle if present; replace a:p children
    var keep = [];
    Array.prototype.forEach.call(txBody.childNodes, function (ch) {
      if (ch.nodeType === 1 && (ch.localName === "bodyPr" || ch.localName === "lstStyle")) {
        keep.push(ch);
      }
    });
    while (txBody.firstChild) txBody.removeChild(txBody.firstChild);
    keep.forEach(function (ch) {
      txBody.appendChild(ch);
    });
    var frag = new DOMParser().parseFromString(
      "<root xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
        paragraphsXml +
        "</root>",
      "application/xml"
    );
    var root = frag.documentElement;
    Array.prototype.forEach.call(root.childNodes, function (ch) {
      if (ch.nodeType === 1) {
        txBody.appendChild(txBody.ownerDocument.importNode(ch, true));
      }
    });
  }

  /** Strip viewer placeholder / empty notes so we don't write junk into PPTX. */
  function normalizeNotesHtmlForSave(html) {
    if (!html) return "";
    var wrap = document.createElement("div");
    wrap.innerHTML = html;
    if (wrap.querySelector && wrap.querySelector("p.empty-notes")) return "";
    var text = (wrap.innerText || wrap.textContent || "").replace(/\u00a0/g, " ").trim();
    if (!text) return "";
    return html;
  }

  function nextNotesSlideNumber(zip) {
    var max = 0;
    Object.keys(zip.files).forEach(function (n) {
      var m = n.match(/^ppt\/notesSlides\/notesSlide(\d+)\.xml$/i);
      if (m) max = Math.max(max, parseInt(m[1], 10));
    });
    return max + 1;
  }

  /** Full notes part XML (string-built — no DOM xmlns corruption). */
  function buildNotesSlideXml(paragraphsXml) {
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
      'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
      "<p:cSld><p:spTree>" +
      '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
      '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>' +
      '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>' +
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/>' +
      '<p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr>' +
      '<p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>' +
      '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/>' +
      '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>' +
      '<p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/>' +
      "<p:txBody><a:bodyPr/><a:lstStyle/>" +
      paragraphsXml +
      "</p:txBody></p:sp>" +
      "</p:spTree></p:cSld>" +
      "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>"
    );
  }

  function buildNotesSlideRels(slideFileName) {
    // Link notes → parent slide (notesMaster optional; PPT opens without it)
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/' +
      slideFileName +
      '"/>' +
      "</Relationships>"
    );
  }

  function ensureNotesRelOnSlideRels(relsXml, notesFileName) {
    var target = "../notesSlides/" + notesFileName;
    if (/notesSlide/i.test(relsXml)) {
      // Point existing notes relationship at our notes part
      relsXml = relsXml.replace(
        /Target="[^"]*notesSlide[^"]*"/gi,
        'Target="' + target + '"'
      );
      return relsXml;
    }
    var maxRid = 0;
    relsXml.replace(/Id="rId(\d+)"/g, function (_, n) {
      maxRid = Math.max(maxRid, parseInt(n, 10));
      return _;
    });
    var rid = "rId" + (maxRid + 1);
    var rel =
      '<Relationship Id="' +
      rid +
      '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="' +
      target +
      '"/>';
    if (relsXml.indexOf("</Relationships>") < 0) {
      return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        rel +
        "</Relationships>"
      );
    }
    return relsXml.replace("</Relationships>", rel + "</Relationships>");
  }

  function minimalSlideRelsWithNotes(notesFileName) {
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/' +
      notesFileName +
      '"/>' +
      "</Relationships>"
    );
  }

  /**
   * Write notes into a copy of the original PPTX.
   * Unedited slides: leave notes parts alone (file defaults / spacing / bold).
   * Only rebuild notes for slides the user actually edited (notesDirty).
   */
  async function buildPptxBlobWithNotes(pptxBuffer, slides) {
    if (typeof JSZip === "undefined") throw new Error("JSZip failed to load");
    if (!pptxBuffer) {
      throw new Error("Original PowerPoint data missing — open the .pptx again, then SAVE.");
    }

    var zip = await JSZip.loadAsync(pptxBuffer);
    var presRelsFile = zip.file("ppt/_rels/presentation.xml.rels");
    var presFile = zip.file("ppt/presentation.xml");
    if (!presRelsFile || !presFile) throw new Error("Invalid PowerPoint file");

    var idToTarget = relsMap(await presRelsFile.async("text"));
    var presDoc = new DOMParser().parseFromString(
      await presFile.async("text"),
      "application/xml"
    );
    var R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    var slideTargets = [];
    xmlLocalAll(presDoc, "sldId").forEach(function (el) {
      var rid = el.getAttributeNS(R_NS, "id") || el.getAttribute("r:id");
      if (!rid || !idToTarget[rid]) return;
      var t = String(idToTarget[rid]).replace(/\\/g, "/").replace(/^\//, "");
      if (!t.startsWith("ppt/")) t = "ppt/" + t;
      slideTargets.push(t);
    });

    var nextNotes = nextNotesSlideNumber(zip);
    var notesPartsCreated = [];

    for (var i = 0; i < slideTargets.length; i++) {
      var slidePath = slideTargets[i];
      var slideFile = zip.file(slidePath);
      if (!slideFile) continue;

      var slideBase = slidePath.split("/").pop(); // slide1.xml
      var baseDir = slidePath.split("/").slice(0, -1).join("/");
      var relsPath = baseDir + "/_rels/" + slideBase + ".rels";
      var relsFile = zip.file(relsPath);
      var relsXml = relsFile
        ? await relsFile.async("text")
        : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
      var rels = relsMap(relsXml);

      var slideState = slides[i] || {};
      var dirty = !!slideState.notesDirty;

      // —— Unedited: leave this slide's notes exactly as in the uploaded file ——
      if (!dirty) {
        continue;
      }

      var notesHtml = normalizeNotesHtmlForSave(slideState.notesHtml || "");

      // Existing notes part?
      var notesRid = Object.keys(rels).find(function (id) {
        return (rels[id] || "").toLowerCase().indexOf("notesslide") >= 0;
      });
      var notesPath = null;
      var notesFileName = null;

      if (notesRid) {
        notesPath = joinPptPath(baseDir, rels[notesRid]);
        notesFileName = notesPath.split("/").pop();
      }

      // —— Edited (or new) notes only ——
      var paragraphsXml = notesHtmlToOoxmlParagraphs(notesHtml);

      if (!notesRid) {
        notesFileName = "notesSlide" + nextNotes + ".xml";
        nextNotes++;
        notesPath = "ppt/notesSlides/" + notesFileName;
        notesPartsCreated.push(notesPath);
        if (!relsFile) {
          relsXml = minimalSlideRelsWithNotes(notesFileName);
        } else {
          relsXml = ensureNotesRelOnSlideRels(relsXml, notesFileName);
        }
        zip.file(relsPath, relsXml);
        // New notes part: only link to parent slide
        zip.file(
          "ppt/notesSlides/_rels/" + notesFileName + ".rels",
          buildNotesSlideRels(slideBase)
        );
      } else {
        // Keep existing notes .rels (notesMaster etc.) — only replace notes XML body file
        notesPath = joinPptPath(baseDir, rels[notesRid]);
        notesFileName = notesPath.split("/").pop();
      }

      zip.file(notesPath, buildNotesSlideXml(paragraphsXml));
    }

    // Content_Types: notes slide overrides
    var ctFile = zip.file("[Content_Types].xml");
    if (ctFile) {
      var ctXml = await ctFile.async("string");
      var allNotes = Object.keys(zip.files).filter(function (n) {
        return /^ppt\/notesSlides\/notesSlide\d+\.xml$/i.test(n);
      });
      var ctAdds = [];
      allNotes.forEach(function (n) {
        var part = "/" + n;
        if (ctXml.indexOf('PartName="' + part + '"') < 0) {
          ctAdds.push(
            '<Override PartName="' +
              part +
              '" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>'
          );
        }
      });
      if (ctAdds.length) {
        ctXml = ctXml.replace("</Types>", ctAdds.join("") + "</Types>");
        zip.file("[Content_Types].xml", ctXml);
      }
    }

    return await zip.generateAsync({
      type: "blob",
      mimeType:
        "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      compression: "DEFLATE",
    });
  }

  async function downloadBlob(blob, fileName) {
    var name = fileName || "presentation_notes.pptx";
    if (window.showSaveFilePicker) {
      try {
        var handle = await window.showSaveFilePicker({
          suggestedName: name,
          types: [
            {
              description: "PowerPoint",
              accept: {
                "application/vnd.openxmlformats-officedocument.presentationml.presentation": [".pptx"],
              },
            },
          ],
        });
        var writable = await handle.createWritable();
        await writable.write(blob);
        await writable.close();
        return { method: "picker" };
      } catch (err) {
        if (err && err.name === "AbortError") return { cancelled: true };
      }
    }
    var a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.appendChild(a);
    a.click();
    setTimeout(function () {
      URL.revokeObjectURL(a.href);
      a.remove();
    }, 1500);
    return { method: "download" };
  }

  global.PptxNotesLoader = {
    pptxFileToSlides: pptxFileToSlides,
    openPptxInNewTab: openPptxInNewTab,
    openPlaceholderTab: openPlaceholderTab,
    handoffTake: handoffTake,
    handoffPut: handoffPut,
    safeDeckKey: safeDeckKey,
    bufferFingerprint: bufferFingerprint,
    isAppleMobile: isAppleMobile,
    buildPptxBlobWithNotes: buildPptxBlobWithNotes,
    normalizeNotesPasteHtml: normalizeNotesPasteHtml,
    downloadBlob: downloadBlob,
  };
})(typeof window !== "undefined" ? window : globalThis);

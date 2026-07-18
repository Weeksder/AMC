/* AMC Productivity Studio — browser port (JSZip + optional Tesseract.js) */
(function () {
  "use strict";

  function $(id) {
    return document.getElementById(id);
  }

  function setStatus(el, msg, kind) {
    if (!el) return;
    el.textContent = msg || "";
    el.className = "status" + (kind ? " " + kind : "");
  }

  function downloadBlob(blob, name) {
    var a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.appendChild(a);
    a.click();
    setTimeout(function () {
      URL.revokeObjectURL(a.href);
      a.remove();
    }, 1200);
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
    if (!window.showOpenFilePicker) return null;
    try {
      var handles = await window.showOpenFilePicker({
        multiple: false,
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
      return handles && handles[0] ? handles[0] : null;
    } catch (e) {
      if (e && e.name === "AbortError") return null;
      return null;
    }
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

  async function runOcrOnBlob(blob) {
    if (typeof Tesseract === "undefined") {
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
      ocrText.value = text || "[No text detected]";
      setStatus(ocrStatus, text ? "Done — you can Copy." : "No text found.", text ? "ok" : "err");
    } catch (e) {
      console.error(e);
      setStatus(ocrStatus, "OCR error: " + (e.message || e), "err");
    }
  }

  ocrFile.addEventListener("change", function () {
    var f = ocrFile.files && ocrFile.files[0];
    if (f) runOcrOnBlob(f);
  });

  document.addEventListener("paste", function (e) {
    if (!$("panel-ocr").classList.contains("active")) return;
    var items = e.clipboardData && e.clipboardData.items;
    if (!items) return;
    for (var i = 0; i < items.length; i++) {
      if (items[i].type.indexOf("image") === 0) {
        e.preventDefault();
        runOcrOnBlob(items[i].getAsFile());
        return;
      }
    }
  });

  $("ocrCopy").addEventListener("click", async function () {
    try {
      await navigator.clipboard.writeText(ocrText.value || "");
      setStatus(ocrStatus, "Copied.", "ok");
    } catch (e) {
      ocrText.select();
      document.execCommand("copy");
      setStatus(ocrStatus, "Copied.", "ok");
    }
  });
  $("ocrClear").addEventListener("click", function () {
    ocrText.value = "";
    setStatus(ocrStatus, "Ready");
  });

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
  var mergeTargetFile = null;
  var mergeTargetHandle = null; // FileSystemFileHandle for overwrite (Chrome/Edge)

  function setMergeSource(file) {
    mergeSourceFile = file;
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

  async function loadSourceFromFile(f) {
    if (!f || !/\.pptx$/i.test(f.name)) return;
    setMergeSource(f);
    var z = await JSZip.loadAsync(f);
    var c = slideNames(z).length;
    $("mergeSlides").value = "1-" + c;
    setStatus($("mergeStatus"), "Source: " + c + " slides", "ok");
  }
  async function loadTargetFromFile(f, handle) {
    if (!f || !/\.pptx$/i.test(f.name)) return;
    setMergeTarget(f, handle);
    var z = await JSZip.loadAsync(f);
    var c = slideNames(z).length;
    $("mergeInsert").value = String(c);
    var msg = "Target: " + c + " slides";
    if (handle) msg += " — Chrome/Edge can overwrite this file on Extract";
    else msg += " — result will download (browser cannot silently overwrite)";
    setStatus($("mergeStatus"), msg, "ok");
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
    var handle = await pickPptxWithHandle();
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
   * Same cleanup as desktop AMC Pdf's modification (extract path):
   * - Fix short "Source:" / "Note:" text boxes (margins, bullets, indent)
   * - Strip watermark-like text shapes
   * - Strip empty title placeholders and bare "Shape N" overlays when a picture exists
   */
  function cleanSlideXmlLikeDesktop(xml) {
    try {
      var doc = new DOMParser().parseFromString(xml, "application/xml");
      if (doc.getElementsByTagName("parsererror").length) return xml;

      function shapeName(sp) {
        var pr = localAll(sp, "cNvPr")[0];
        return pr ? pr.getAttribute("name") || "" : "";
      }
      function shapeText(sp) {
        return localAll(sp, "t")
          .map(function (t) {
            return t.textContent || "";
          })
          .join("");
      }
      function removeEl(el) {
        if (el && el.parentNode) el.parentNode.removeChild(el);
      }

      var pics = localAll(doc, "pic");
      var hasPic = pics.length > 0;

      // Remove watermark / citation text shapes (and empty title shells when a photo is present)
      localAll(doc, "sp").forEach(function (sp) {
        var text = shapeText(sp);
        var low = text.toLowerCase().replace(/\s+/g, " ").trim();
        var name = shapeName(sp);

        // Explicit watermark / footer / source lines
        if (
          /source\s*:|note\s*:|copyright|©|all rights reserved|confidential|watermark|amc\s*english|www\.|https?:\/\//i.test(
            low
          )
        ) {
          removeEl(sp);
          return;
        }

        // Empty title placeholder sitting on image slides (常見 AMC: 標題 1)
        if (hasPic && !low && /標題|title|placeholder/i.test(name)) {
          removeEl(sp);
          return;
        }

        // Bare "Shape 123" overlays with almost no text (often watermark chrome)
        if (hasPic && /^Shape\s*\d+$/i.test(name) && low.length < 8) {
          removeEl(sp);
          return;
        }
      });

      // Original desktop "fix" for short source/note-style boxes that remain
      localAll(doc, "txBody").forEach(function (tb) {
        var full = localAll(tb, "t")
          .map(function (t) {
            return t.textContent || "";
          })
          .join("")
          .toLowerCase();
        if (!(full.indexOf("source:") >= 0 || full.indexOf("note:") >= 0 || (full.length > 0 && full.length < 150))) {
          return;
        }
        var bodyPr = null;
        for (var i = 0; i < tb.childNodes.length; i++) {
          var ch = tb.childNodes[i];
          if (ch.nodeType === 1 && ch.localName === "bodyPr") bodyPr = ch;
        }
        if (bodyPr) {
          ["lIns", "tIns", "rIns", "bIns"].forEach(function (a) {
            if (bodyPr.getAttribute(a) != null) bodyPr.setAttribute(a, "0");
          });
        }
        localAll(tb, "pPr").forEach(function (ppr) {
          if (ppr.getAttribute("lvl")) ppr.setAttribute("lvl", "0");
          ppr.removeAttribute("marL");
          ppr.removeAttribute("indent");
          if (ppr.getAttribute("algn")) ppr.setAttribute("algn", "l");
          // remove bullet children
          Array.prototype.slice.call(ppr.childNodes).forEach(function (c) {
            if (c.nodeType !== 1) return;
            if (/^bu/i.test(c.localName || "")) removeEl(c);
          });
        });
      });

      var out = new XMLSerializer().serializeToString(doc);
      out = out.replace(/\sxmlns=""/g, "");
      // Keep a normal OOXML declaration
      if (out.indexOf("<?xml") !== 0) {
        out = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + out;
      }
      return out;
    } catch (e) {
      console.warn("cleanSlideXmlLikeDesktop failed", e);
      return xml;
    }
  }

  /**
   * LEGACY merge into another deck (optional). Prefer extractSlidesFromSource.
   */
  async function mergePresentations(srcFile, tgtFile, nums, insertAfter) {
    var srcZip = await JSZip.loadAsync(srcFile);
    var tgtZip = await JSZip.loadAsync(tgtFile);
    var srcSlides = slideNames(srcZip);
    var tgtSlides = slideNames(tgtZip);
    var tgtCount = tgtSlides.length;
    if (insertAfter > tgtCount) insertAfter = tgtCount;

    // Match target slide size to source (prevents stretch / repair quirks)
    try {
      var srcPrs = await srcZip.file("ppt/presentation.xml").async("string");
      var tgtPrs0 = await tgtZip.file("ppt/presentation.xml").async("string");
      var sz = srcPrs.match(/<p:sldSz\b[^/]*\/>/);
      if (sz) {
        tgtPrs0 = tgtPrs0.replace(/<p:sldSz\b[^/]*\/>/, sz[0]);
        tgtZip.file("ppt/presentation.xml", tgtPrs0);
      }
    } catch (e) {}

    // Media: prefix names so they never collide
    // IMPORTANT: rename carefully — image1.jpeg must not match inside image11.jpeg
    var mediaMap = {};
    var srcMedia = Object.keys(srcZip.files).filter(function (n) {
      return n.indexOf("ppt/media/") === 0 && !srcZip.files[n].dir;
    });
    for (var mi = 0; mi < srcMedia.length; mi++) {
      var mp = srcMedia[mi];
      var base = mp.split("/").pop();
      // Unique name that cannot be a prefix of another image name
      var newBase = "x" + mi + "_" + base;
      var newPath = "ppt/media/" + newBase;
      var c = 1;
      while (tgtZip.file(newPath)) {
        newBase = "x" + mi + "_" + c + "_" + base;
        newPath = "ppt/media/" + newBase;
        c++;
      }
      tgtZip.file(newPath, await srcZip.file(mp).async("uint8array"));
      mediaMap[base] = newBase;
    }

    function escapeRegExp(s) {
      return String(s).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    }

    function remapMediaInText(text) {
      if (!text) return text;
      // Longest names first so image11.jpeg is not partially matched as image1.jpeg
      var keys = Object.keys(mediaMap).sort(function (a, b) {
        return b.length - a.length;
      });
      keys.forEach(function (oldN) {
        var neu = mediaMap[oldN];
        var esc = escapeRegExp(oldN);
        // ../media/name or /media/name — only full filename before " ? # or end
        text = text.replace(
          new RegExp("(\\.\\./media/|/media/|media/)" + esc + '(?=["\'?#\\s>]|$)', "g"),
          "$1" + neu
        );
        // Target="image1.jpeg"
        text = text.replace(new RegExp('(Target=")' + esc + '(")', "g"), "$1" + neu + "$2");
      });
      return text;
    }

    // Ensure Content_Types has image defaults (blank shell may lack some)
    function ensureImageContentTypes(ctXml) {
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
        var re = new RegExp('Extension="' + ext + '"', "i");
        if (!re.test(ctXml)) {
          ctXml = ctXml.replace(
            "<Types",
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"'
          );
          // If xmlns already on Types, simple insert Default after <Types...>
          ctXml = ctXml.replace(
            /(<Types[^>]*>)/,
            '$1<Default Extension="' + ext + '" ContentType="' + defaults[ext] + '"/>'
          );
        }
      });
      return ctXml;
    }

    // Light cleanup — do NOT use DOM (breaks namespaces / images)
    function repairSlideXml(xml) {
      return xml;
    }

    var addedMeta = []; // {newNum, rid, id}
    var nextSlideNum = tgtCount + 1;
    var notesCounter = 1;
    // find max existing notesSlide number in target
    Object.keys(tgtZip.files).forEach(function (n) {
      var m = n.match(/^ppt\/notesSlides\/notesSlide(\d+)\.xml$/);
      if (m) notesCounter = Math.max(notesCounter, parseInt(m[1], 10) + 1);
    });

    for (var i = 0; i < nums.length; i++) {
      var sn = nums[i];
      var sp = "ppt/slides/slide" + sn + ".xml";
      var rp = "ppt/slides/_rels/slide" + sn + ".xml.rels";
      if (!srcZip.file(sp)) continue;

      var newNum = nextSlideNum++;
      var nsp = "ppt/slides/slide" + newNum + ".xml";
      var nrp = "ppt/slides/_rels/slide" + newNum + ".xml.rels";

      // Slide XML uses r:embed="rIdN" only — do not rename strings inside slide XML
      var xml = await srcZip.file(sp).async("string");
      xml = repairSlideXml(xml);
      if (xml.indexOf("<?xml") !== 0) {
        xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + xml;
      }
      tgtZip.file(nsp, xml);

      if (srcZip.file(rp)) {
        var rels = await srcZip.file(rp).async("string");
        // Relationships hold Target="../media/imageN.jpeg" — remap filenames here
        rels = remapMediaInText(rels);

        // Copy notes slide if linked
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
            var notesXml = await srcZip.file(notesTarget).async("string");
            notesXml = remapMediaInText(notesXml);
            if (notesXml.indexOf("<?xml") !== 0) {
              notesXml =
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + notesXml;
            }
            tgtZip.file(newNotesPath, notesXml);

            // notes rels
            var oldNotesRels =
              "ppt/notesSlides/_rels/" + notesTarget.split("/").pop() + ".rels";
            var newNotesRels = "ppt/notesSlides/_rels/" + newNotesName + ".rels";
            if (srcZip.file(oldNotesRels)) {
              var nr = await srcZip.file(oldNotesRels).async("string");
              nr = remapMediaInText(nr);
              // point slide target to new slide number
              nr = nr.replace(
                /Target="[^"]*slide\d+\.xml"/g,
                'Target="../slides/slide' + newNum + '.xml"'
              );
              tgtZip.file(newNotesRels, nr);
            }

            rels = rels.replace(
              /Target="[^"]*notesSlides\/notesSlide\d+\.xml"/gi,
              'Target="../notesSlides/' + newNotesName + '"'
            );
            notesCounter++;
          }
        }

        if (rels.indexOf("<?xml") !== 0) {
          rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + rels;
        }
        tgtZip.file(nrp, rels);
      } else {
        // minimal rels so PowerPoint is happy
        tgtZip.file(
          nrp,
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
        );
      }

      addedMeta.push({ newNum: newNum });
    }

    if (!addedMeta.length) throw new Error("No slides were copied (check slide numbers).");

    // --- presentation.xml.rels: append slide relationships (string) ---
    var prsRelsXml = await tgtZip.file("ppt/_rels/presentation.xml.rels").async("string");
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
      throw new Error("Invalid presentation.xml.rels");
    }
    prsRelsXml = prsRelsXml.replace(
      "</Relationships>",
      relChunks.join("") + "</Relationships>"
    );
    tgtZip.file("ppt/_rels/presentation.xml.rels", prsRelsXml);

    // --- presentation.xml sldIdLst (string) ---
    var prsXml = await tgtZip.file("ppt/presentation.xml").async("string");
    var maxId = 255;
    prsXml.replace(/<p:sldId\b[^>]*\bid="(\d+)"/g, function (_, n) {
      maxId = Math.max(maxId, parseInt(n, 10));
      return _;
    });
    var newSldTags = addedMeta.map(function (meta, idx) {
      meta.id = maxId + idx + 1;
      return '<p:sldId id="' + meta.id + '" r:id="' + meta.rid + '"/>';
    });

    // existing sldId tags in order
    var existingTags = [];
    prsXml.replace(/<p:sldId\b[^/]*\/>/g, function (tag) {
      existingTags.push(tag);
      return tag;
    });

    var ordered = existingTags
      .slice(0, insertAfter)
      .concat(newSldTags)
      .concat(existingTags.slice(insertAfter));
    var newList = "<p:sldIdLst>" + ordered.join("") + "</p:sldIdLst>";

    if (/<p:sldIdLst\/>/.test(prsXml)) {
      prsXml = prsXml.replace(/<p:sldIdLst\/>/, newList);
    } else if (/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/.test(prsXml)) {
      prsXml = prsXml.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, newList);
    } else {
      // insert after sldMasterIdLst
      prsXml = prsXml.replace(
        /<\/p:sldMasterIdLst>/,
        "</p:sldMasterIdLst>" + newList
      );
    }
    if (prsXml.indexOf("<?xml") !== 0) {
      prsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + prsXml;
    }
    tgtZip.file("ppt/presentation.xml", prsXml);

    // --- Content_Types: image defaults + slide overrides (no DOM xmlns="") ---
    var ctXml = await tgtZip.file("[Content_Types].xml").async("string");
    ctXml = ensureImageContentTypes(ctXml);
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
    Object.keys(tgtZip.files).forEach(function (n) {
      var m = n.match(/^ppt\/notesSlides\/(notesSlide\d+)\.xml$/);
      if (!m) return;
      var part = "/ppt/notesSlides/" + m[1] + ".xml";
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
    }
    tgtZip.file("[Content_Types].xml", ctXml);

    var blob = await tgtZip.generateAsync({
      type: "blob",
      mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      compression: "DEFLATE",
    });
    return { blob: blob, count: addedMeta.length };
  }

  // Load built-in blank target (must be deployed as blank-target.pptx next to this page)
  async function ensureBlankTarget() {
    var urls = ["blank-target.pptx", "./blank-target.pptx"];
    // If site is under a repo path, also try relative to current folder
    try {
      urls.push(new URL("blank-target.pptx", location.href).href);
    } catch (e) {}

    var lastErr = null;
    for (var i = 0; i < urls.length; i++) {
      try {
        var res = await fetch(urls[i], { cache: "no-cache" });
        if (!res.ok) throw new Error("HTTP " + res.status + " for " + urls[i]);
        var buf = await res.arrayBuffer();
        if (!buf || buf.byteLength < 1000) throw new Error("File too small / empty");
        var file = new File([buf], "blank-target.pptx", {
          type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        });
        await loadTargetFromFile(file, null);
        $("mergeTargetLabel").textContent = "blank-target.pptx (built-in)";
        $("mergeInsert").value = "0";
        setStatus(
          $("mergeStatus"),
          "Using built-in blank target. Choose a Source, then Extract.",
          "ok"
        );
        return true;
      } catch (e) {
        lastErr = e;
        console.warn("blank-target load failed:", urls[i], e);
      }
    }
    setStatus(
      $("mergeStatus"),
      "No target available — upload blank-target.pptx to your GitHub repo (same folder as index.html / amc-studio.html), then hard-refresh. Or click Target and pick any .pptx. (" +
        (lastErr && lastErr.message ? lastErr.message : "load failed") +
        ")",
      "err"
    );
    return false;
  }

  $("mergeGo").addEventListener("click", async function () {
    if (!mergeSourceFile) {
      setStatus($("mergeStatus"), "Choose a Source PPTX file first.", "err");
      return;
    }
    if (typeof JSZip === "undefined") {
      setStatus($("mergeStatus"), "JSZip missing.", "err");
      return;
    }
    var btn = $("mergeGo");
    btn.disabled = true;
    setStatus($("mergeStatus"), "Extracting slides + cleaning watermarks…");
    try {
      var probe = await JSZip.loadAsync(await readFileAsArrayBuffer(mergeSourceFile));
      var srcCount = slideNames(probe).length;
      if (!srcCount) throw new Error("No slides found in source file.");

      var rawList = ($("mergeSlides") && $("mergeSlides").value) || "";
      if (!String(rawList).trim()) {
        rawList = "1-" + srcCount;
        if ($("mergeSlides")) $("mergeSlides").value = rawList;
      }
      var nums = parseSlideList(rawList, srcCount);
      if (!nums.length) {
        throw new Error(
          "No valid slide numbers. This source has " + srcCount + " slides (try 1-" + srcCount + ")."
        );
      }

      var result = await extractSlidesFromSource(mergeSourceFile, nums);

      // Always force a real download with a NEW name (Brave/Edge: no silent overwrite)
      var stamp = new Date().toISOString().replace(/[:.]/g, "-").slice(0, 19);
      var base = (mergeSourceFile.name || "slides").replace(/\.pptx$/i, "");
      var outName = base + "_extracted_" + stamp + ".pptx";
      downloadBlob(result.blob, outName);
      setStatus(
        $("mergeStatus"),
        "Downloaded “" +
          outName +
          "” (" +
          result.count +
          " of " +
          result.totalInSource +
          " slides, cleaned). Check Downloads.",
        "ok"
      );
    } catch (e) {
      console.error(e);
      setStatus($("mergeStatus"), "Error: " + (e.message || e), "err");
    }
    btn.disabled = false;
  });

  if ($("mergeTargetLabel")) {
    $("mergeTargetLabel").textContent = "Not needed (extract from Source only)";
  }
  if ($("mergeStatus")) {
    setStatus(
      $("mergeStatus"),
      "1) Choose Source  2) Set slides (e.g. 1-10)  3) Extract — file goes to Downloads.",
      "ok"
    );
  }

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

  async function stripImagesFromFile(file) {
    var zip = await JSZip.loadAsync(file);
    var slides = slideNames(zip);
    for (var i = 0; i < slides.length; i++) {
      var path = slides[i];
      var xml = await zip.file(path).async("string");
      var doc = new DOMParser().parseFromString(xml, "application/xml");
      var pics = localAll(doc, "pic");
      pics.forEach(function (pic) {
        var parent = pic.parentNode;
        if (parent) parent.removeChild(pic);
      });
      zip.file(path, new XMLSerializer().serializeToString(doc));
    }
    return zip.generateAsync({
      type: "blob",
      mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
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

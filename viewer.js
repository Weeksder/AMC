(function () {
  let DECK_KEY = "deck";
  let SLIDES = [];
  let ORIGINAL_PPTX_BUFFER = null; // ArrayBuffer of the opened .pptx for SAVE
  let ORIGINAL_FILE_NAME = "";
  let STORAGE_PREFIX = "pptxNotesViewer:v3:" + DECK_KEY + ":";
  const NOTES_FORMAT = 3;

  let index = 0;
  let notesHeightPx = null;
  let wheelLock = false;
  let sidebarW = 176;
  let spacingLevel = 0;
  let fontSizeLevel = 2; // index into FONT_SIZES (14px default)

  const appEl = document.getElementById("app");
  const emptyState = document.getElementById("emptyState");
  const thumbsEl = document.getElementById("thumbs");
  const mainSlide = document.getElementById("mainSlide");
  const notesBody = document.getElementById("notesBody");
  const pageCurrent = document.getElementById("pageCurrent");
  const pageTotal = document.getElementById("pageTotal");
  const slideChrome = document.getElementById("slideChrome");
  const btnPrev = document.getElementById("btnPrev");
  const btnNext = document.getElementById("btnNext");
  const btnOpen = document.getElementById("btnOpen");
  const openFileInput = document.getElementById("openFileInput");
  const main = document.getElementById("main");
  const slidePane = document.getElementById("slidePane");
  const sidebar = document.getElementById("sidebar");
  const splitter = document.getElementById("splitter");
  const vSplitter = document.getElementById("vSplitter");
  const deckTitle = document.getElementById("deckTitle");
  const btnSpaceDown = document.getElementById("btnSpaceDown");
  const btnSpaceUp = document.getElementById("btnSpaceUp");
  const btnFontDown = document.getElementById("btnFontDown");
  const btnFontUp = document.getElementById("btnFontUp");
  const btnTogglePicture = document.getElementById("btnTogglePicture");
  const btnToggleTools = document.getElementById("btnToggleTools");
  const btnHome = document.getElementById("btnHome");
  const btnSaveJson = document.getElementById("btnSaveJson");
  const btnResetNotes = document.getElementById("btnResetNotes");
  const saveStatus = document.getElementById("saveStatus");

  function storageKey(i) {
    return STORAGE_PREFIX + "slide:" + SLIDES[i].n;
  }

  function loadSavedNotes(i) {
    try {
      const raw = localStorage.getItem(storageKey(i));
      if (raw === null) return null;
      if (raw.charAt(0) === "{") {
        const obj = JSON.parse(raw);
        if (obj && obj.format === NOTES_FORMAT && typeof obj.html === "string") return obj.html;
      }
      return null;
    } catch (e) {
      return null;
    }
  }

  function saveCurrentNotes() {
    if (!SLIDES.length || index < 0 || index >= SLIDES.length) return;
    try {
      // Only write editor HTML when the user actually edited this slide.
      // Browsers re-serialize contenteditable HTML on focus/navigation — that
      // must NOT mark notes dirty or SAVE would rebuild and change spacing.
      if (!SLIDES[index].notesDirty) return;

      let html = notesBody.innerHTML;
      // Don't persist the placeholder as real notes
      if (notesBody.querySelector && notesBody.querySelector("p.empty-notes")) {
        html = "";
      }
      localStorage.setItem(storageKey(index), JSON.stringify({ format: NOTES_FORMAT, html: html }));
      SLIDES[index].notesHtml = html;
      SLIDES[index].hasNotes = notesBody.innerText.trim().length > 0 &&
        !(notesBody.querySelector && notesBody.querySelector("p.empty-notes"));
      const thumb = thumbsEl.querySelector('.thumb[data-i="' + index + '"]');
      if (thumb) thumb.classList.toggle("has-notes", SLIDES[index].hasNotes);
    } catch (e) {
      console.warn(e);
    }
  }

  function anyNotesDirty() {
    return SLIDES.some(function (s) {
      return !!s.notesDirty;
    });
  }

  function emptyNotesHtml() {
    return '<p class="empty-notes">No notes for this slide. Click here to type.</p>';
  }

  var SKIP_LOCAL_NOTES = false;

  function setDeck(payload) {
    DECK_KEY = payload.deckKey || payload.title || "deck";
    SLIDES = payload.slides || [];
    ORIGINAL_PPTX_BUFFER = payload.pptxBuffer || null;
    ORIGINAL_FILE_NAME = payload.fileName || DECK_KEY + ".pptx";
    // Include content fingerprint so PDF→EDIT and an older same-name PPTX
    // do not share cached speaker notes in localStorage.
    var contentId = payload.contentId || "";
    if (
      !contentId &&
      ORIGINAL_PPTX_BUFFER &&
      window.PptxNotesLoader &&
      PptxNotesLoader.bufferFingerprint
    ) {
      contentId = PptxNotesLoader.bufferFingerprint(ORIGINAL_PPTX_BUFFER);
    }
    STORAGE_PREFIX =
      "pptxNotesViewer:v4:" + DECK_KEY + ":" + (contentId || "unknown") + ":";
    // Fresh open (Extract → Viewer / Open file): start from notes in the PPTX only
    SKIP_LOCAL_NOTES = payload.freshOpen === true;
    document.title = DECK_KEY + " — Notes Viewer";
    deckTitle.textContent = DECK_KEY;
    emptyState.classList.remove("show");
    appEl.style.visibility = "visible";
  }

  function buildThumbs() {
    thumbsEl.innerHTML = "";
    SLIDES.forEach(function (s, i) {
      s._originalHtml = s.notesHtml || "";
      // Only rehydrate browser drafts when not a brand-new open.
      // Previously, same filename reused old footnotes on a new PDF extract.
      if (s.notesDirty == null) s.notesDirty = false;
      if (!SKIP_LOCAL_NOTES) {
        const saved = loadSavedNotes(i);
        if (saved !== null) {
          const fromFile = (s.notesHtml || "").replace(/<[^>]+>/g, "").trim();
          const fromSaved = saved.replace(/<[^>]+>/g, "").trim();
          // Prefer notes already in the PPTX; only fill empty slides from drafts
          if (!fromFile && fromSaved) {
            s.notesHtml = saved;
            s.hasNotes = true;
            s.notesDirty = true; // draft is not original OOXML
          }
        }
      }
      s.hasNotes = !!(s.notesHtml || "").replace(/<[^>]+>/g, "").trim();
      const row = document.createElement("div");
      row.className = "thumb" + (s.hasNotes ? " has-notes" : "");
      row.dataset.i = String(i);
      row.innerHTML =
        '<div class="num">' +
        s.n +
        '</div><div class="frame"><img src="' +
        s.img +
        '" alt="Slide ' +
        s.n +
        '" loading="lazy"><span class="dot" title="Has notes"></span></div>';
      row.addEventListener("click", function () {
        show(i);
      });
      thumbsEl.appendChild(row);
    });
    // After first paint of a fresh open, allow saving drafts for this session
    SKIP_LOCAL_NOTES = false;
  }

  function applyNotesHeight() {
    if (appEl.classList.contains("hide-picture")) {
      main.style.removeProperty("--notes-h");
      main.style.gridTemplateRows = "minmax(0, 1fr)";
      return;
    }
    main.style.gridTemplateRows = "";
    const rect = main.getBoundingClientRect();
    let h = notesHeightPx;
    if (h == null) h = Math.round(rect.height * 0.34);
    h = Math.max(100, Math.min(rect.height - 128, h));
    notesHeightPx = h;
    main.style.setProperty("--notes-h", h + "px");
    try {
      localStorage.setItem(STORAGE_PREFIX + "notesHeight", String(h));
    } catch (e) {}
  }

  function show(i, opts) {
    opts = opts || {};
    if (!SLIDES.length || i < 0 || i >= SLIDES.length) return;
    if (!opts.skipSave && index !== i) saveCurrentNotes();
    index = i;
    const s = SLIDES[i];
    mainSlide.src = s.img;
    mainSlide.alt = "Slide " + s.n;
    let html = s.notesHtml;
    if (!html || !html.replace(/<[^>]+>/g, "").trim()) html = emptyNotesHtml();
    notesBody.innerHTML = html;
    pageCurrent.textContent = String(s.n);
    pageTotal.textContent = String(SLIDES.length);
    slideChrome.textContent = "Slide " + s.n + " of " + SLIDES.length;
    btnPrev.disabled = i === 0;
    btnNext.disabled = i === SLIDES.length - 1;
    thumbsEl.querySelectorAll(".thumb").forEach(function (el, j) {
      el.classList.toggle("active", j === i);
    });
    const active = thumbsEl.querySelector(".thumb.active");
    if (active) active.scrollIntoView({ block: "nearest", behavior: "smooth" });
    try {
      history.replaceState(null, "", "#" + s.n);
    } catch (e) {}
  }

  function nextSlide(dir) {
    show(index + dir);
  }

  // Sidebar width
  const MIN_SIDEBAR = 56;
  const MAX_SIDEBAR = 340;
  function applySidebarWidth(w, persist) {
    w = Math.max(MIN_SIDEBAR, Math.min(MAX_SIDEBAR, Math.round(w)));
    sidebarW = w;
    document.documentElement.style.setProperty("--sidebar-w", w + "px");
    const numW = w < 90 ? 0 : w < 120 ? 14 : 22;
    document.documentElement.style.setProperty("--thumb-num-w", numW + "px");
    appEl.classList.toggle("sidebar-compact", w < 130);
    appEl.classList.toggle("sidebar-mini", w < 90);
    if (persist !== false) {
      try {
        localStorage.setItem(STORAGE_PREFIX + "sidebarW", String(w));
      } catch (e) {}
    }
  }

  let vDragging = false;
  vSplitter.addEventListener("pointerdown", function (e) {
    vDragging = true;
    vSplitter.classList.add("dragging");
    vSplitter.setPointerCapture(e.pointerId);
    e.preventDefault();
    document.body.style.cursor = "ew-resize";
    document.body.style.userSelect = "none";
  });
  vSplitter.addEventListener("pointermove", function (e) {
    if (!vDragging) return;
    applySidebarWidth(e.clientX - appEl.getBoundingClientRect().left);
  });
  function endVDrag(e) {
    if (!vDragging) return;
    vDragging = false;
    vSplitter.classList.remove("dragging");
    document.body.style.cursor = "";
    document.body.style.userSelect = "";
    try {
      vSplitter.releasePointerCapture(e.pointerId);
    } catch (err) {}
  }
  vSplitter.addEventListener("pointerup", endVDrag);
  vSplitter.addEventListener("pointercancel", endVDrag);

  function onSlideWheel(e) {
    if (notesBody.contains(e.target)) return;
    e.preventDefault();
    if (wheelLock) return;
    if (Math.abs(e.deltaY) < 8) return;
    wheelLock = true;
    nextSlide(e.deltaY > 0 ? 1 : -1);
    setTimeout(function () {
      wheelLock = false;
    }, 140);
  }
  function onSidebarWheel(e) {
    if (notesBody.contains(e.target)) return;
    const dy = e.deltaY;
    const dx = e.deltaX;
    if (Math.abs(dx) > Math.abs(dy) || e.shiftKey) {
      e.preventDefault();
      const delta = e.shiftKey ? dy : dx;
      applySidebarWidth(sidebarW + (delta > 0 ? 12 : -12));
      return;
    }
    onSlideWheel(e);
  }
  slidePane.addEventListener("wheel", onSlideWheel, { passive: false });
  sidebar.addEventListener("wheel", onSidebarWheel, { passive: false });
  vSplitter.addEventListener("wheel", onSidebarWheel, { passive: false });

  btnPrev.addEventListener("click", function () {
    nextSlide(-1);
  });
  btnNext.addEventListener("click", function () {
    nextSlide(1);
  });
  btnHome.addEventListener("click", function () {
    location.href = "index.html";
  });

  document.addEventListener("keydown", function (e) {
    const inNotes = notesBody === document.activeElement || notesBody.contains(document.activeElement);
    if (e.key === "Escape") {
      if (appEl.classList.contains("hide-picture") || appEl.classList.contains("hide-tools")) {
        e.preventDefault();
        setHidePicture(false);
        setHideTools(false);
      }
      return;
    }
    if (inNotes) return;
    if (e.key === "ArrowLeft" || e.key === "ArrowUp" || e.key === "PageUp") {
      e.preventDefault();
      nextSlide(-1);
    } else if (e.key === "ArrowRight" || e.key === "ArrowDown" || e.key === "PageDown" || e.key === " ") {
      e.preventDefault();
      nextSlide(1);
    } else if (e.key === "Home") {
      e.preventDefault();
      show(0);
    } else if (e.key === "End") {
      e.preventDefault();
      show(SLIDES.length - 1);
    }
  });

  // Notes height splitter
  let dragging = false;
  splitter.addEventListener("pointerdown", function (e) {
    dragging = true;
    splitter.classList.add("dragging");
    splitter.setPointerCapture(e.pointerId);
    e.preventDefault();
    document.body.style.cursor = "ns-resize";
    document.body.style.userSelect = "none";
  });
  splitter.addEventListener("pointermove", function (e) {
    if (!dragging) return;
    notesHeightPx = main.getBoundingClientRect().bottom - e.clientY;
    applyNotesHeight();
  });
  function endDrag(e) {
    if (!dragging) return;
    dragging = false;
    splitter.classList.remove("dragging");
    document.body.style.cursor = "";
    document.body.style.userSelect = "";
    try {
      splitter.releasePointerCapture(e.pointerId);
    } catch (err) {}
  }
  splitter.addEventListener("pointerup", endDrag);
  splitter.addEventListener("pointercancel", endDrag);
  window.addEventListener("resize", applyNotesHeight);

  // Notes edit
  let saveTimer = null;
  notesBody.addEventListener("input", function () {
    if (SLIDES[index]) SLIDES[index].notesDirty = true;
    clearTimeout(saveTimer);
    saveTimer = setTimeout(saveCurrentNotes, 250);
  });
  notesBody.addEventListener("focus", function () {
    const empty = notesBody.querySelector("p.empty-notes");
    if (empty && notesBody.innerText.trim() === empty.innerText.trim()) {
      notesBody.innerHTML = "<p><br></p>";
    }
  });

  // Paste from ChatGPT / Docs / Slides → simple one-line-per-<p> HTML (keep bold)
  notesBody.addEventListener("paste", function (e) {
    e.preventDefault();
    if (SLIDES[index]) SLIDES[index].notesDirty = true;
    var html = "";
    var plain = "";
    try {
      html = (e.clipboardData && e.clipboardData.getData("text/html")) || "";
      plain = (e.clipboardData && e.clipboardData.getData("text/plain")) || "";
    } catch (err) {}
    var normalized = "<p><br></p>";
    if (
      window.PptxNotesLoader &&
      typeof PptxNotesLoader.normalizeNotesPasteHtml === "function"
    ) {
      normalized = PptxNotesLoader.normalizeNotesPasteHtml(html, plain);
    } else if (plain) {
      normalized = plain
        .replace(/\r\n/g, "\n")
        .split("\n")
        .map(function (line) {
          if (!line.length) return "<p><br></p>";
          return (
            "<p>" +
            String(line)
              .replace(/&/g, "&amp;")
              .replace(/</g, "&lt;")
              .replace(/>/g, "&gt;") +
            "</p>"
          );
        })
        .join("");
    }
    var empty = notesBody.querySelector("p.empty-notes");
    if (empty) {
      notesBody.innerHTML = "";
    }
    var ok = false;
    try {
      ok = document.execCommand("insertHTML", false, normalized);
    } catch (err) {
      ok = false;
    }
    if (!ok) {
      notesBody.innerHTML = normalized;
    }
    saveCurrentNotes();
  });

  // Enter → new paragraph (not Chrome's default <div>, which looks Google-like)
  notesBody.addEventListener("keydown", function (e) {
    if (e.key !== "Enter" || e.shiftKey) return;
    // Let list commands use default behavior
    try {
      if (document.queryCommandState("insertUnorderedList")) return;
    } catch (err) {}
    e.preventDefault();
    if (SLIDES[index]) SLIDES[index].notesDirty = true;
    try {
      document.execCommand("insertParagraph");
    } catch (err) {
      document.execCommand("insertHTML", false, "<p><br></p>");
    }
    saveCurrentNotes();
  });

  document.getElementById("fmtBar").addEventListener("click", function (e) {
    const btn = e.target.closest("button[data-cmd]");
    if (!btn) return;
    e.preventDefault();
    notesBody.focus();
    if (SLIDES[index]) SLIDES[index].notesDirty = true;
    document.execCommand(btn.dataset.cmd, false, null);
    saveCurrentNotes();
  });

  // Insert a horizontal split line (desktop notes separator style)
  var btnSplitLine = document.getElementById("btnSplitLine");
  if (btnSplitLine) {
    btnSplitLine.addEventListener("click", function (e) {
      e.preventDefault();
      notesBody.focus();
      if (SLIDES[index]) SLIDES[index].notesDirty = true;
      // Clear empty placeholder first
      var empty = notesBody.querySelector("p.empty-notes");
      if (empty) notesBody.innerHTML = "";
      var sep =
        '<p><br></p><hr class="notes-sep"><p><br></p>';
      var ok = false;
      try {
        ok = document.execCommand("insertHTML", false, sep);
      } catch (err) {
        ok = false;
      }
      if (!ok) {
        notesBody.innerHTML = (notesBody.innerHTML || "") + sep;
      }
      saveCurrentNotes();
    });
  }

  btnResetNotes.addEventListener("click", function () {
    if (!confirm("Reset notes for this slide to the original from the PowerPoint?")) return;
    try {
      localStorage.removeItem(storageKey(index));
    } catch (e) {}
    SLIDES[index].notesHtml = SLIDES[index]._originalHtml || "";
    SLIDES[index].notesDirty = false; // back to original OOXML on SAVE
    SLIDES[index].hasNotes = (SLIDES[index].notesHtml || "").replace(/<[^>]+>/g, "").trim().length > 0;
    show(index, { skipSave: true });
  });
  window.addEventListener("beforeunload", saveCurrentNotes);

  // Spacing (viewer display only)
  const SPACING = [
    { line: 1.12, para: "0px", blank: "0px" },
    { line: 1.2, para: "2px", blank: "0px" },
    { line: 1.3, para: "4px", blank: "2px" },
    { line: 1.45, para: "8px", blank: "4px" },
    { line: 1.6, para: "12px", blank: "6px" },
  ];
  function applySpacing(level) {
    const i = Math.max(0, Math.min(SPACING.length - 1, level | 0));
    spacingLevel = i;
    const s = SPACING[i];
    document.documentElement.style.setProperty("--notes-line-height", String(s.line));
    document.documentElement.style.setProperty("--notes-para-gap", s.para);
    document.documentElement.style.setProperty("--notes-blank-gap", s.blank);
    btnSpaceDown.disabled = i <= 0;
    btnSpaceUp.disabled = i >= SPACING.length - 1;
    try {
      localStorage.setItem(STORAGE_PREFIX + "spacing", String(i));
    } catch (e) {}
  }
  btnSpaceDown.addEventListener("click", function () {
    applySpacing(spacingLevel - 1);
  });
  btnSpaceUp.addEventListener("click", function () {
    applySpacing(spacingLevel + 1);
  });

  // Notes text size — viewer only (never written into the PPTX on SAVE)
  const FONT_SIZES = [11, 12, 14, 16, 18, 20, 24];
  const FONT_SIZE_PREF_KEY = "pptxNotesViewer:notesFontSize";
  function applyFontSize(level) {
    const i = Math.max(0, Math.min(FONT_SIZES.length - 1, level | 0));
    fontSizeLevel = i;
    var px = FONT_SIZES[i] + "px";
    document.documentElement.style.setProperty("--notes-font-size", px);
    if (btnFontDown) btnFontDown.disabled = i <= 0;
    if (btnFontUp) btnFontUp.disabled = i >= FONT_SIZES.length - 1;
    try {
      localStorage.setItem(FONT_SIZE_PREF_KEY, String(i));
    } catch (e) {}
  }
  if (btnFontDown) {
    btnFontDown.addEventListener("click", function () {
      applyFontSize(fontSizeLevel - 1);
    });
  }
  if (btnFontUp) {
    btnFontUp.addEventListener("click", function () {
      applyFontSize(fontSizeLevel + 1);
    });
  }

  function setHidePicture(hidden) {
    appEl.classList.toggle("hide-picture", !!hidden);
    btnTogglePicture.textContent = hidden ? "Show pic" : "Picture";
    btnTogglePicture.classList.toggle("on", hidden);
    try {
      localStorage.setItem(STORAGE_PREFIX + "hidePicture", hidden ? "1" : "0");
    } catch (e) {}
    applyNotesHeight();
  }
  function setHideTools(hidden) {
    appEl.classList.toggle("hide-tools", !!hidden);
    btnToggleTools.textContent = hidden ? "Show tools" : "Tools";
    btnToggleTools.classList.toggle("on", hidden);
    try {
      localStorage.setItem(STORAGE_PREFIX + "hideTools", hidden ? "1" : "0");
    } catch (e) {}
  }
  btnTogglePicture.addEventListener("click", function () {
    setHidePicture(!appEl.classList.contains("hide-picture"));
  });
  btnToggleTools.addEventListener("click", function () {
    setHideTools(!appEl.classList.contains("hide-tools"));
  });

  var btnViewerHelp = document.getElementById("btnViewerHelp");
  if (btnViewerHelp) {
    btnViewerHelp.addEventListener("click", function () {
      alert(
        "Notes Viewer help\n\n" +
          "• Browse slides in the left thumbnails or use ◀ ▶ / arrow keys.\n" +
          "• Type speaker notes in the bottom pane. Use B / I / U and lists to format.\n" +
          "• A− / A+ change notes text size in the viewer only (not saved to the file).\n" +
          "• Hide picture / Hide tools toggle the layout.\n" +
          "• SAVE downloads a .pptx with your notes written into the file.\n" +
          "• Open that file in PowerPoint and use View → Notes to see them.\n" +
          "• Open loads another PowerPoint. Home returns to the main page."
      );
    });
  }

  // SAVE → download a real .pptx with your notes (one step)
  if (btnSaveJson) {
    btnSaveJson.addEventListener("click", async function () {
      if (!SLIDES.length) {
        alert("Open a PowerPoint first.");
        return;
      }
      if (!ORIGINAL_PPTX_BUFFER) {
        alert("Open the .pptx file again, then press SAVE.\n(The original file is needed to build the download.)");
        return;
      }
      btnSaveJson.disabled = true;
      btnSaveJson.textContent = "…";
      try {
        // Flush editor only for slides the user actually edited
        saveCurrentNotes();
        var base = (ORIGINAL_FILE_NAME || DECK_KEY + ".pptx")
          .replace(/\.pptx$/i, "")
          .replace(/_notes$/i, "");
        var outName = base + "_notes.pptx";
        var blob;
        // No edits → download the original file as-is (spaces, bold, everything)
        if (!anyNotesDirty()) {
          blob = new Blob([ORIGINAL_PPTX_BUFFER], {
            type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
          });
        } else {
          blob = await PptxNotesLoader.buildPptxBlobWithNotes(ORIGINAL_PPTX_BUFFER, SLIDES);
        }
        var result = await PptxNotesLoader.downloadBlob(blob, outName);
        if (result && result.cancelled) {
          btnSaveJson.textContent = "SAVE";
          btnSaveJson.disabled = false;
          return;
        }
        // Keep buffer in sync so a second SAVE still has the full package
        try {
          ORIGINAL_PPTX_BUFFER = await blob.arrayBuffer();
          ORIGINAL_FILE_NAME = outName;
        } catch (e) {}
        if (saveStatus) {
          saveStatus.textContent = anyNotesDirty()
            ? "saved " + outName + " — open this file in PowerPoint (View → Notes)"
            : "saved " + outName + " (unchanged from upload)";
          saveStatus.classList.add("show");
        }
      } catch (err) {
        console.error(err);
        alert("Could not save: " + (err && err.message ? err.message : err));
      }
      btnSaveJson.textContent = "SAVE";
      btnSaveJson.disabled = false;
    });
  }

  // Open another file (local / iCloud / Files) — Safari-safe
  function triggerOpenPicker() {
    openFileInput.value = "";
    openFileInput.click();
  }
  btnOpen.addEventListener("click", triggerOpenPicker);
  var btnEmptyOpen = document.getElementById("btnEmptyOpen");
  if (btnEmptyOpen) btnEmptyOpen.addEventListener("click", triggerOpenPicker);

  openFileInput.addEventListener("change", async function () {
    const f = openFileInput.files && openFileInput.files[0];
    if (!f) return;
    if (btnOpen) btnOpen.disabled = true;
    var pre = null;
    var apple = PptxNotesLoader.isAppleMobile && PptxNotesLoader.isAppleMobile();
    if (!apple) pre = PptxNotesLoader.openPlaceholderTab && PptxNotesLoader.openPlaceholderTab();
    try {
      var result = await PptxNotesLoader.openPptxInNewTab(f, "viewer.html", {
        preOpenedWindow: pre,
        sameTabFallback: true,
      });
      if (saveStatus && result.mode === "new-tab") {
        saveStatus.textContent = "opened in new tab";
        saveStatus.classList.add("show");
      }
    } catch (err) {
      if (pre && !pre.closed) {
        try {
          pre.close();
        } catch (e) {}
      }
      alert(err && err.message ? err.message : String(err));
    }
    if (btnOpen) btnOpen.disabled = false;
  });

  function loadPrefs() {
    try {
      const h = parseInt(localStorage.getItem(STORAGE_PREFIX + "notesHeight") || "", 10);
      if (h > 0) notesHeightPx = h;
    } catch (e) {}
    try {
      const sp = localStorage.getItem(STORAGE_PREFIX + "spacing");
      if (sp !== null && sp !== "") {
        const n = parseInt(sp, 10);
        if (!isNaN(n)) spacingLevel = n;
      }
    } catch (e) {}
    applySpacing(spacingLevel);
    try {
      const fs = localStorage.getItem(FONT_SIZE_PREF_KEY);
      if (fs !== null && fs !== "") {
        const n = parseInt(fs, 10);
        if (!isNaN(n)) fontSizeLevel = n;
      }
    } catch (e) {}
    applyFontSize(fontSizeLevel);
    try {
      const sw = parseInt(localStorage.getItem(STORAGE_PREFIX + "sidebarW") || "", 10);
      if (sw > 0) sidebarW = sw;
    } catch (e) {}
    applySidebarWidth(sidebarW, false);
    // Picture/slide is always visible by default (do not restore "hidden")
    try {
      setHidePicture(false);
      setHideTools(localStorage.getItem(STORAGE_PREFIX + "hideTools") === "1");
    } catch (e) {
      setHidePicture(false);
      setHideTools(false);
    }
  }

  function bootWithSlides() {
    loadPrefs();
    buildThumbs();
    applyNotesHeight();
    const hash = parseInt((location.hash || "").replace("#", ""), 10);
    const start = hash >= 1 && hash <= SLIDES.length ? hash - 1 : 0;
    show(start, { skipSave: true });
  }

  (async function init() {
    appEl.style.visibility = "hidden";
    try {
      const params = new URLSearchParams(location.search);
      const id = params.get("open");
      if (id && window.PptxNotesLoader) {
        const payload = await PptxNotesLoader.handoffTake(id);
        if (payload && payload.slides && payload.slides.length) {
          setDeck(payload);
          try {
            const clean = new URL(location.href);
            clean.searchParams.delete("open");
            history.replaceState(null, "", clean.pathname + clean.search + (clean.hash || ""));
          } catch (e) {}
          bootWithSlides();
          return;
        }
      }
    } catch (e) {
      console.warn(e);
    }
    // No handoff — empty state
    emptyState.classList.add("show");
    appEl.style.visibility = "hidden";
  })();
})();

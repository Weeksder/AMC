# PPTX Notes Viewer (web)

Simple web app:

1. **Home** (`index.html`) — open or drop a `.pptx`
2. **Viewer** (`viewer.html`) — slides + speaker notes in a **new tab**

No Python required for this flow. Works on **GitHub Pages** or any static host.

## Try locally

From this `web` folder:

```powershell
cd "C:\Users\PaintPeel\Desktop\PPTX Notes Viewer\web"
python -m http.server 8080
```

Then open: http://localhost:8080/

> Opening `index.html` via `file://` may block pop-ups or modules in some browsers. A local server is more reliable.

## Deploy to GitHub Pages

1. Create a GitHub repository (public is fine for the app code).
2. Upload **everything inside this `web` folder** to the repo root  
   (or put it in a `/docs` folder if you prefer).
3. Repo → **Settings** → **Pages**:
   - Source: **Deploy from a branch**
   - Branch: `main` (or `master`), folder: `/ (root)` or `/docs`
4. After a minute, open:  
   `https://YOUR_USERNAME.github.io/YOUR_REPO/`

### Suggested files in the repo

```
index.html
viewer.html
viewer.css
viewer.js
pptx-loader.js
jszip.min.js
README.md
```

## Privacy

- Files are processed **in your browser**. They are not uploaded to GitHub.
- Notes edits are stored in **that browser** (localStorage) only.
- To continue on another device: open the same `.pptx` again (or export later if you add that).

## Pages in this folder

| File | What it is |
|------|------------|
| `index.html` | Home: Notes Viewer + AMC tools section below |
| `viewer.html` | Notes viewer tab |
| `amc-studio.html` | **Separate** full-page AMC Productivity Studio |
| `amc-tools.js` / `amc-tools.css` | AMC tools logic & styles |
| `jszip.min.js` | Shared PPTX zip library |
| `pptx-loader.js` / `viewer.js` / `viewer.css` | Notes viewer |

Upload **all** of these files to GitHub Pages (including **`blank-target.pptx`** — required for Slide Extractor).

## Features

### Notes Viewer
- Open `.pptx`, thumbs + notes, SAVE as `.pptx`
- Hide picture / tools, AMC Online Class link

### AMC Productivity Studio (below on home, or `amc-studio.html`)
- Text Extractor (image OCR)
- PPTX Slide Extractor (merge slides)
- Image Stripper
- Special Strip
- Text Formatter

## Apple devices

1. Open the site in **Safari**
2. Tap **Open file** for notes, or use AMC tools below
3. On iPhone/iPad the notes viewer usually stays **in this tab**

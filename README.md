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

## Features

- Thumbnails (resize by dragging the edge or scrolling sideways)
- Main slide + editable notes
- Hide picture / hide tools
- Open another `.pptx` in a new tab from the viewer
- Soft line breaks and bold from PowerPoint notes

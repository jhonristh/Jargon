# Jargon: Jhon Rey D. Oquendo — Journey to Become an Engineer

A static portfolio (HTML, CSS, vanilla JS). No npm, framework, or database.
All paths are relative, so it works on Vercel, VS Code Live Server, Python's server, and even by double-clicking `index.html`.

## 1. Test locally in VS Code
1. Open this folder (the one containing `index.html`) in VS Code: File → Open Folder.
2. Install the recommended extension **Live Server** when VS Code prompts you.
3. Right-click `index.html` → **Open with Live Server** (http://127.0.0.1:5500).

No Live Server? Run in the VS Code terminal, then open http://localhost:8000:

    python -m http.server 8000 --bind 127.0.0.1      (Windows: py -m http.server 8000 --bind 127.0.0.1)

Tip: Terminal → Run Task offers "Serve locally" and "Update site from Excel".

## 2. Fill in your content (Excel)
1. Open `tools/Portfolio_Content.xlsx`. Read the Instructions sheet.
2. Fill the yellow cells on Awards, Certificates, Seminars, Journey, Contact.
3. Copy files into `assets/certificates/` (certificates) or `assets/documents/` (award proof, résumé), and type the exact file names.
4. Close Excel, then run (once: `pip install openpyxl`):

       python tools/build_content.py        (Windows: py tools/build_content.py)

5. Refresh your browser. Only rows with **Show on site = Yes** and **Status = Verified** are published.

## 3. Deploy to Vercel
**Option A, GitHub (recommended):** push this folder to a GitHub repository → vercel.com → Add New Project → import it.
Framework Preset: **Other**. Leave Build Command and Output Directory empty. Deploy.
Every later `git push` redeploys automatically. Run the Excel build script before you commit.

**Option B, CLI:**

    npm i -g vercel
    vercel          # preview
    vercel --prod   # production

`vercel.json` adds basic security headers. `.vercelignore` keeps `tools/` (including the Excel file) off the public site.

## 4. Editing elsewhere
- `index.html`, `about.html`, `projects.html`, `skills.html`: edit the HTML directly.
- `cases.js`: project descriptions (BUSINA, AM Receiver, GabayNegosyo, RelAI).
- `style.css`: colors, typography, layout. `app.js`: menu, filters, tabs.
- Do not hand-edit text between `<!--NAME:START-->` and `<!--NAME:END-->` markers; the Excel build overwrites it.

Google Fonts loads online; system fonts are the fallback offline.

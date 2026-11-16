#!/usr/bin/env python3
"""Update the portfolio from tools/Portfolio_Content.xlsx.

Usage (from the folder that contains index.html):
    pip install openpyxl          # one time
    python tools/build_content.py # Windows: py tools/build_content.py

It rewrites the blocks between <!--NAME:START--> and <!--NAME:END--> in
achievements.html, journey.html, credentials.html and contact.html.
A row is published only when  Show on site = Yes  AND  Status = Verified.
Rows with Status = EXAMPLE or Draft are never published.
"""
import html
import re
import sys
from pathlib import Path
from urllib.parse import quote

try:
    from openpyxl import load_workbook
except ImportError:
    sys.exit("openpyxl is missing. Run:  pip install openpyxl")

TOOLS = Path(__file__).resolve().parent
ROOT = TOOLS.parent
XLSX = TOOLS / "Portfolio_Content.xlsx"
IMAGES = {".png", ".jpg", ".jpeg", ".webp", ".gif"}


def esc(v):
    return html.escape(str(v).strip()) if v is not None else ""


def read_sheet(wb, name):
    """Header is row 4; data starts at row 5. Returns published rows as dicts."""
    ws = wb[name]
    heads = [str(c.value).strip() if c.value else "" for c in ws[4]]
    out = []
    for r in ws.iter_rows(min_row=5, values_only=True):
        d = {h: (v if v is not None and str(v).strip() != "" else None) for h, v in zip(heads, r) if h}
        if not any(d.values()):
            continue
        status = str(d.get("Status") or "").strip().lower()
        show = str(d.get("Show on site") or "").strip().lower()
        if status != "verified" or show != "yes":
            continue
        out.append(d)
    return out


def local_file(folder, name):
    """Return a site-relative URL if the file exists, else warn and return None."""
    if not name:
        return None
    rel = f"assets/{folder}/{str(name).strip()}"
    if not (ROOT / rel).is_file():
        print(f"  ! File not found, link skipped: {rel}")
        return None
    return quote(rel)


def safe_url(v):
    v = str(v).strip()
    if re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", v):
        return "mailto:" + v
    if v.lower().startswith(("https://", "http://", "mailto:")):
        return v
    if re.match(r"^[\w.-]+\.[a-z]{2,}(/.*)?$", v, re.I):
        return "https://" + v
    return None


def empty(label, title, text, small, extra=""):
    lab = f"<span>{label}</span>" if label else ""
    return f'<div class="empty">{lab}<h2>{title}</h2><p>{text}</p>{small}{extra}</div>'


def row(left, title_html, body_html):
    return f'<div class="row"><span>{esc(left)}</span><h2>{title_html}</h2><div>{body_html}</div></div>'


def build_awards(wb):
    items = []
    for d in read_sheet(wb, "Awards"):
        if not d.get("Title"):
            continue
        body = f"<p>{esc(d.get('Description'))}</p>" if d.get("Description") else ""
        meta = []
        if d.get("Organization / Event"):
            meta.append(esc(d["Organization / Event"]))
        if d.get("Level"):
            meta.append(esc(d["Level"]))
        if d.get("My Role"):
            meta.append("My contribution: " + esc(d["My Role"]))
        if meta:
            body += f"<p><small>{' · '.join(meta)}</small></p>"
        proof = local_file("documents", d.get("Proof file"))
        if proof:
            body += f'<a class="inline" href="{proof}">View proof ↗</a>'
        items.append(row(d.get("Date / Year") or "", esc(d["Title"]), body))
    print(f"  Awards published: {len(items)}")
    return "".join(items) or empty("", "Milestones will appear here.", "Awards and recognition are being documented.", "")


def build_journey(wb):
    items = []
    for d in read_sheet(wb, "Journey"):
        if d.get("Title"):
            items.append(row(d.get("Period") or "", esc(d["Title"]), f"<p>{esc(d.get('Description'))}</p>"))
    print(f"  Journey published: {len(items)}")
    return "".join(items) or empty("", "Milestones will appear here.", "The journey is being documented.", "")


def build_seminars(wb):
    items = []
    for d in read_sheet(wb, "Seminars"):
        if not d.get("Title"):
            continue
        meta = " · ".join(esc(d[k]) for k in ("Organizer", "Format", "Participation") if d.get(k))
        body = f"<p><small>{meta}</small></p>" if meta else ""
        if d.get("Takeaway"):
            body += f"<p>{esc(d['Takeaway'])}</p>"
        if d.get("Topics"):
            tags = "".join(f"<span>{esc(t)}</span>" for t in str(d["Topics"]).split(",") if t.strip())
            body += f'<div class="tags">{tags}</div>'
        cert = local_file("certificates", d.get("Certificate file (optional)"))
        if cert:
            body += f'<p><a class="inline" href="{cert}">Related certificate ↗</a></p>'
        items.append(row(d.get("Date") or "", esc(d["Title"]), body))
    print(f"  Seminars published: {len(items)}")
    return "".join(items) or empty(
        "SEMINARS & WORKSHOPS", "More learning to document.",
        "Attendance records and reflections will appear here, including the event, organizer, date, and key takeaways.",
        "<small>The collection is being prepared.</small>")


def build_certs(wb):
    items = []
    for d in read_sheet(wb, "Certificates"):
        if not d.get("Title"):
            continue
        meta = " · ".join(esc(d[k]) for k in ("Issuer", "Type") if d.get(k))
        body = f"<p><small>{meta}</small></p>" if meta else ""
        f = local_file("certificates", d.get("File name"))
        links = ""
        if f:
            if Path(f).suffix.lower() in IMAGES:
                body += f'<a href="{f}"><img class="thumb" loading="lazy" src="{f}" alt="Certificate: {esc(d["Title"])}"></a>'
            links += f'<a class="inline" href="{f}">Open document ↗</a>'
        v = safe_url(d["Verification URL"]) if d.get("Verification URL") else None
        if v:
            links += f'<a class="inline" href="{esc(v)}" rel="noopener">Verify ↗</a>'
        if links:
            body += f"<p>{links}</p>"
        items.append(row(d.get("Date issued") or "", esc(d["Title"]), body))
    print(f"  Certificates published: {len(items)}")
    return "".join(items) or empty(
        "CERTIFICATE ARCHIVE", "Milestones, documented.",
        "Certificates of attendance, completion, and recognition will be displayed here with their issuer and date.",
        "<small>Certificate files have not yet been added.</small>")


def build_contact(wb):
    items = []
    for d in read_sheet(wb, "Contact"):
        url = safe_url(d["URL or email"]) if d.get("URL or email") else None
        if not (d.get("Label") and url):
            continue
        shown = esc(str(d["URL or email"]).replace("mailto:", ""))
        body = f"<p>{esc(d.get('Note'))}</p>" if d.get("Note") else ""
        items.append(row(d["Label"], f'<a href="{esc(url)}" rel="noopener">{shown} ↗</a>', body))
    print(f"  Contact links published: {len(items)}")
    if items:
        return "".join(items)
    return empty("", "Let’s connect.", "Professional contact details and profile links will be added here soon.", "",
                 '<a class="button" href="projects.html">Explore my work ↗</a>')


def inject(page, name, content):
    path = ROOT / page
    text = path.read_text(encoding="utf-8")
    pat = re.compile(rf"(<!--{name}:START-->).*?(<!--{name}:END-->)", re.S)
    if not pat.search(text):
        sys.exit(f"Marker <!--{name}:START--> not found in {page}")
    path.write_text(pat.sub(lambda m: m.group(1) + content + m.group(2), text), encoding="utf-8")


def main():
    if not XLSX.is_file():
        sys.exit(f"Not found: {XLSX}")
    wb = load_workbook(XLSX, data_only=True)
    print("Reading", XLSX.name)
    inject("achievements.html", "AWARDS", build_awards(wb))
    inject("journey.html", "JOURNEY", build_journey(wb))
    inject("credentials.html", "SEMINARS", build_seminars(wb))
    inject("credentials.html", "CERTS", build_certs(wb))
    inject("contact.html", "CONTACT", build_contact(wb))
    print("Done. Open index.html through a local server to preview.")


if __name__ == "__main__":
    main()

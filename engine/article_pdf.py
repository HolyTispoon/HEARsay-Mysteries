"""
The packet's newspaper article as a printable front page, docs/<Scenario ID>/newspaper-article.pdf.

The page is HTML (masthead, headline, deck, byline, the story in three columns, a paragraph quoting the Clause set
off in a box) printed to PDF by headless Chrome. The PDF's Subject records a fingerprint of that HTML; when the
article, masthead and layout are unchanged, the build keeps the PDF it already has instead of printing a new one,
so a rebuild changes nothing (Chrome stamps each print with the time).

Chrome: set CHROME=/path/to/chrome, or it is looked for in /Applications and on the PATH. Needs pypdf.
"""
import base64, hashlib, html, os, shutil, subprocess, tempfile

PDF_NAME = "newspaper-article.pdf"
CHROME_PATHS = ("/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
                "/Applications/Chromium.app/Contents/MacOS/Chromium")
CHROME_NAMES = ("google-chrome", "google-chrome-stable", "chromium", "chromium-browser", "chrome")

# The look of the game's Newspaper passage, laid out for one or two letter pages. Fonts are the Mac's own, so
# printing needs no network.
CSS = """
@page { size: letter; margin: 0.45in; }
:root { --paper: #f7f5ef; --ink: #1d2126; --soft: #4a5058; --navy: #062843; --gold-light: #d3a753; --rule: #c9c4b6; }
html { background: #fff; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
body { margin: 0; color: var(--ink); font: 9.6pt/1.42 Georgia, "Times New Roman", serif; }
.masthead { display: block; width: 100%; height: auto; padding-bottom: 6pt; border-bottom: 1pt solid var(--navy); }
h1 { margin: 14pt 0 0; font: 700 44pt/1 "Bodoni 72", Didot, "Bodoni MT", Georgia, serif; color: var(--navy); text-wrap: balance; }
.deck { margin: 10pt 0 0; max-width: 40em; font: italic 14pt/1.35 Georgia, serif; color: var(--soft); text-wrap: balance; }
.byline { margin: 12pt 0 0; padding: 6pt 0; border-top: 1pt solid var(--rule); border-bottom: 1pt solid var(--rule);
          font: 500 8pt/1.4 "Helvetica Neue", Arial, sans-serif; letter-spacing: .1em; text-transform: uppercase; }
.story { margin-top: 14pt; columns: 3; column-gap: 20pt; column-rule: 1pt solid var(--rule); text-align: justify; hyphens: auto; }
.story p { margin: 0 0 .7em; text-indent: 1.4em; }
.story p:first-child { text-indent: 0; }
.story p:first-child::first-letter { float: left; font: 700 4.2em/.8 "Bodoni 72", Didot, Georgia, serif; color: var(--navy); padding: .06em .08em 0 0; }
.story blockquote { break-inside: avoid; margin: .3em 0 1em; padding: 9pt 11pt; background: var(--navy); color: var(--paper);
                    text-align: left; hyphens: manual; font-style: italic; }
.end { display: inline-block; width: .55em; height: .55em; margin-left: .25em; background: var(--navy); }
"""


def page_html(article, masthead):
    """The front page: the article (from article_json: headline, deck, byline, paragraphs with quote flags) under the
    masthead image (a path, or None)."""
    e = html.escape
    parts = []
    last = len(article["paragraphs"]) - 1
    for i, p in enumerate(article["paragraphs"]):
        end = '<span class="end"></span>' if i == last and not p["quote"] else ""
        parts.append(f"<blockquote>{e(p['text'])}</blockquote>" if p["quote"] else f"<p>{e(p['text'])}{end}</p>")
    head = ""
    if masthead:
        with open(masthead, "rb") as f:
            head = f'<img class="masthead" alt="" src="data:image/png;base64,{base64.b64encode(f.read()).decode()}">'
    return (f'<!doctype html><html lang="en"><head><meta charset="utf-8"><title>{e(article["headline"])}</title>'
            f"<style>{CSS}</style></head><body>{head}<h1>{e(article['headline'])}</h1>"
            f'<p class="deck">{e(article["deck"])}</p><div class="byline">{e(article["byline"])}</div>'
            f'<article class="story">{"".join(parts)}</article></body></html>')


def fingerprint(page):
    return "HEARsay article " + hashlib.sha256(page.encode("utf-8")).hexdigest()


def recorded(pdf):
    """The fingerprint a PDF of ours records, or None."""
    from pypdf import PdfReader
    try:
        return (PdfReader(pdf).metadata or {}).get("/Subject")
    except Exception:
        return None


def find_chrome():
    found = os.environ.get("CHROME") or next((p for p in CHROME_PATHS if os.path.exists(p)), None)
    return found or next(filter(None, map(shutil.which, CHROME_NAMES)), None)


def write_pdf(article, masthead, out, previous=None):
    """Writes the article to the PDF `out`. `previous` is the PDF from the last build (kept if it matches).
    Returns None, or a problem."""
    from pypdf import PdfReader, PdfWriter
    page = page_html(article, masthead)
    if previous and recorded(previous) == fingerprint(page):
        shutil.copyfile(previous, out)
        return None
    chrome = find_chrome()
    if not chrome:
        return "Chrome was not found to print the newspaper article PDF. Install Google Chrome or set CHROME=/path/to/chrome."
    with tempfile.TemporaryDirectory() as tmp:
        src, raw = os.path.join(tmp, "article.html"), os.path.join(tmp, "article.pdf")
        with open(src, "w", encoding="utf-8") as f:
            f.write(page)
        r = subprocess.run([chrome, "--headless", "--disable-gpu", "--no-first-run", "--no-default-browser-check",
                            f"--user-data-dir={os.path.join(tmp, 'profile')}", "--no-pdf-header-footer",
                            f"--print-to-pdf={raw}", "file://" + src], capture_output=True, text=True, timeout=120)
        if not os.path.exists(raw):
            return "Chrome could not print the newspaper article PDF:\n" + (r.stderr or r.stdout).strip()
        w = PdfWriter(clone_from=PdfReader(raw))
        w.metadata = {"/Title": article["headline"], "/Author": article["byline"].removeprefix("By ").strip(),
                      "/Subject": fingerprint(page)}
        with open(out, "wb") as f:
            w.write(f)
    return None

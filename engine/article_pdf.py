"""
The packet's newspaper article as a printable front page, docs/<Scenario ID>/newspaper-article.pdf.

The page is HTML (masthead, headline, deck, byline, the illustration and its caption if the scenario has them, the
story in three columns, a paragraph quoting the Clause set off in a box, and the tab's box such as "Tonight's
panel"), on white, printed to PDF by headless Chrome. The PDF's Subject records a fingerprint of that HTML; when the
article, masthead and layout are unchanged, the build keeps the PDF it already has instead of printing a new one,
so a rebuild changes nothing (Chrome stamps each print with the time).

Chrome: set CHROME=/path/to/chrome, or it is looked for in /Applications and on the PATH. Needs pypdf.
"""
import base64, hashlib, html, mimetypes, os, re, shutil, subprocess, tempfile

PDF_NAME = "newspaper-article.pdf"
CHROME_PATHS = ("/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
                "/Applications/Chromium.app/Contents/MacOS/Chromium")
CHROME_NAMES = ("google-chrome", "google-chrome-stable", "chromium", "chromium-browser", "chrome")

# The look of the game's Newspaper passage, laid out for one or two letter pages. Fonts are the Mac's own, so
# printing needs no network.
CSS = """
@page { size: letter; margin: 0.45in; }
:root { --paper: #fff; --ink: #1d2126; --soft: #4a5058; --navy: #062843; --gold-light: #d3a753; --rule: #c9c4b6; }
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
figure { margin: 14pt 0 0; }
figure svg, figure img { display: block; width: 100%; height: auto; max-height: 2.1in; object-fit: contain; border: 1pt solid var(--rule); }
figcaption { margin-top: 5pt; font: 8.5pt/1.4 Georgia, serif; color: var(--soft); }
figcaption b { font: 700 7pt/1 "Helvetica Neue", Arial, sans-serif; letter-spacing: .1em; text-transform: uppercase; color: var(--navy); margin-right: 5pt; }
.panel { break-inside: avoid; margin: .3em 0 1em; padding: 8pt 10pt; border: 1pt solid var(--navy); text-align: left; hyphens: manual; }
.panel h2 { margin: 0 0 6pt; font: 700 7.5pt/1 "Helvetica Neue", Arial, sans-serif; letter-spacing: .14em; text-transform: uppercase; color: var(--navy); }
.panel ul { list-style: none; margin: 0; padding: 0; }
.panel li { margin: 0 0 5pt; font-size: 8.8pt; line-height: 1.3; color: var(--soft); }
.panel li b { display: block; font-size: 9.4pt; color: var(--ink); }
.end { display: inline-block; width: .55em; height: .55em; margin-left: .25em; background: var(--navy); }
"""


def image(path, cls=""):
    """An image file for the page: an SVG inline (so it draws with the page's fonts), anything else as a data URL."""
    with open(path, "rb") as f:
        data = f.read()
    if path.lower().endswith(".svg"):
        return re.sub(r"^\s*<\?xml[^>]*\?>\s*", "", data.decode("utf-8"))
    kind = mimetypes.guess_type(path)[0] or "image/png"
    return f'<img class="{cls}" alt="" src="data:{kind};base64,{base64.b64encode(data).decode()}">'


def page_html(article, masthead, art=None):
    """The front page: the article (from article_json: headline, deck, byline, paragraphs with quote flags, caption,
    panel) under the masthead image, with the illustration `art` (paths, or None)."""
    e = html.escape
    parts = []
    last = len(article["paragraphs"]) - 1
    for i, p in enumerate(article["paragraphs"]):
        end = '<span class="end"></span>' if i == last and not p["quote"] else ""
        parts.append(f"<blockquote>{e(p['text'])}</blockquote>" if p["quote"] else f"<p>{e(p['text'])}{end}</p>")
    panel = article.get("panel")
    if panel:
        items = "".join(f"<li><b>{e(i['name'])}</b>{e(i['text'])}</li>" for i in panel["items"])
        parts.insert(panel["after"], f'<aside class="panel"><h2>{e(panel["title"])}</h2><ul>{items}</ul></aside>')
    head = image(masthead, "masthead") if masthead else ""
    figure = ""
    if art:
        caption = f'<figcaption><b>Illustration</b>{e(article["caption"])}</figcaption>' if article.get("caption") else ""
        figure = f"<figure>{image(art)}{caption}</figure>"
    return (f'<!doctype html><html lang="en"><head><meta charset="utf-8"><title>{e(article["headline"])}</title>'
            f"<style>{CSS}</style></head><body>{head}<h1>{e(article['headline'])}</h1>"
            f'<p class="deck">{e(article["deck"])}</p><div class="byline">{e(article["byline"])}</div>'
            f'{figure}<article class="story">{"".join(parts)}</article></body></html>')


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


def write_pdf(article, masthead, art, out, previous=None):
    """Writes the article, with its masthead and illustration (paths, or None), to the PDF `out`. `previous` is the PDF from the last build (kept if it matches).
    Returns None, or a problem."""
    from pypdf import PdfReader, PdfWriter
    page = page_html(article, masthead, art)
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

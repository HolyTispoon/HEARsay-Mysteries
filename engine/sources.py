#!/usr/bin/env python3
"""
The sources of truth and the repository's copies of them.

The sources of truth live in Google Drive, outside this repository, and only the author edits them:
  - the framework Google Doc "HEARsay mysteries" (the build checks the game text against it, and takes the game's
    rules and HEAR guide from two of its tabs);
  - the game text Google Sheet "HEARsay game text";
  - for each scenario, its Response Matrix (a Google Sheet) and its Scenario Packet (a Google Doc).
Their addresses are in engine/sources.json and scenarios/<id>/sources.json.

The build reads the repository's copies of the Sheets and Docs: engine/HEARsay mysteries.docx, engine/HEARsay
game text.xlsx and, in each
scenario folder, the matrix .xlsx and the packet .docx. Nobody edits a copy. update_copies() replaces each copy with
a download of its Google file (File → Download → Microsoft Excel / Word, or the Drive connector's export) when the
download's content differs, after checking every download.

    python3 engine/sources.py [DOWNLOADS_FOLDER]    # lists the sources, and which downloads would be used
"""
import glob, json, os, shutil, sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SCENARIOS = os.path.join(ROOT, "scenarios")
sys.path.insert(0, HERE)


def _load(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def copies(ids):
    """(label, copy path, Google address) for the framework, the game text and each scenario's matrix and packet, in
    build order."""
    from game_text import DEFAULT as GAME_TEXT, FRAMEWORK
    urls = _load(os.path.join(HERE, "sources.json"))
    out = [("framework", FRAMEWORK, urls["framework"]), ("game text", GAME_TEXT, urls["game_text"])]
    for sid in ids:
        folder = os.path.join(SCENARIOS, sid)
        urls = _load(os.path.join(folder, "sources.json"))
        for kind, ext in (("matrix", ".xlsx"), ("packet", ".docx")):
            found = [p for p in glob.glob(os.path.join(folder, "*" + ext)) if not os.path.basename(p).startswith("~$")]
            if len(found) != 1:
                sys.exit(f"Expected one {kind} copy ({ext}) in scenarios/{sid}/, found {len(found)}")
            out.append((f"{sid} {kind}", found[0], urls.get(kind) or "(not in Google Drive yet)"))
    return out


def find_download(folder, copy_path):
    """The newest download in folder named like the copy ('The Commons' Comma – Response Matrix (1).xlsx' too)."""
    stem, ext = os.path.splitext(os.path.basename(copy_path))
    found = [p for p in glob.glob(os.path.join(glob.escape(folder), glob.escape(stem) + "*" + ext))
             if not os.path.basename(p).startswith("~$")]
    return max(found, key=os.path.getmtime) if found else None


def content(path):
    """What the build reads from a copy: every sheet's cell values, or every paragraph of a document."""
    if path.endswith(".docx"):
        from check_packet import packet_paragraphs
        return packet_paragraphs(path)
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True)
    out = {}
    for ws in wb:
        rows = [[None if v == "" else v for v in r] for r in ws.iter_rows(values_only=True)]
        rows = [r[:max([i + 1 for i, v in enumerate(r) if v is not None], default=0)] for r in rows]
        while rows and not rows[-1]:
            rows.pop()
        out[ws.title] = rows
    return out


def check_downloads(chosen):
    """Problems with the downloads, each checked with the other downloads (or the copies where there is none)."""
    from game_text import read_game_text
    from scenario_matrix import read_matrix
    from check_packet import check as check_packet
    use = {label: download or path for label, path, _, download in chosen}
    new = {label for label, *_, download in chosen if download}
    g, errors = read_game_text(use["game text"], use["framework"])
    problems = [f"game text: {e}" for e in errors]
    if errors:
        return problems
    for label in use:
        if not label.endswith(" matrix"):
            continue
        sid = label[:-len(" matrix")]
        if not new & {"framework", "game text", label, sid + " packet"}:
            continue
        _, merr = read_matrix(use[label], g.facts)
        problems += [f"{label}: {e}" for e in merr]
        if not merr:
            problems += [f"{sid} packet: {p}" for p in check_packet(use[label], use[sid + " packet"], g.facts)]
    return problems


def update_copies(folder, ids):
    """Replace each copy with its download from folder when the content differs. Stops the build, changing nothing,
    if a download has problems. Returns the labels of the copies that changed."""
    chosen = [(label, path, url, find_download(folder, path) if folder else None) for label, path, url in copies(ids)]
    if folder and not any(d for *_, d in chosen):
        sys.exit(f"No downloads in {folder}. Download the Google Sheets as .xlsx and the Docs as .docx, named as "
                 "they are in Drive, into that folder.")
    problems = check_downloads(chosen)
    if problems:
        sys.exit("The downloads have problems; fix them in the Google files, download again, and rebuild. "
                 "Nothing was copied.\n" + "\n".join(" - " + p for p in problems))
    changed = []
    for label, path, url, download in chosen:
        rel = os.path.relpath(path, ROOT)
        if not download:
            print(f"WARNING: {label}: no download, so the build uses the copy {rel}, which may be behind {url}"
                  if os.path.exists(path) else f"WARNING: {label}: no download and no copy {rel}; download {url}")
        elif os.path.exists(path) and content(download) == content(path):
            print(f"{label}: the copy already matches {os.path.basename(download)}")
        else:
            shutil.copyfile(download, path)
            changed.append(label)
            print(f"{label}: updated the copy {rel} from {os.path.basename(download)}")
    if changed:
        print("Commit the updated copies with the rebuilt docs/.")
    return changed


def main():
    folder = sys.argv[1] if len(sys.argv) > 1 else None
    ids = sorted(n for n in os.listdir(SCENARIOS) if os.path.isdir(os.path.join(SCENARIOS, n)) and not n.startswith("."))
    for label, path, url in copies(ids):
        d = find_download(folder, path) if folder else None
        print(f"{label}: {url}\n    copy: {os.path.relpath(path, ROOT)}" + (f"\n    download: {d}" if d else ""))


if __name__ == "__main__":
    main()

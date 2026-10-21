#!/usr/bin/env python3
"""
Check the game text against the framework Google Doc ("HEARsay mysteries").

    python3 check_framework.py ["HEARsay game text.xlsx" ["HEARsay mysteries.docx"]]

The Google Doc is the source of truth for the game's structure and rules. Its 'Rules of the game', 'HEAR
Framework introduction', 'Game facts' and 'Game text' tabs hold some of the game text and the framework's numbers
(see game_text.py); the rest of the game text is in the Sheet. The build uses the repository's copy,
engine/HEARsay mysteries.docx, and by default so does this script; another .docx export can be passed instead.

Checks:
  - every passage the framework lists (section 4, "Passages text: …") has a screen in the game, and every
    framework name on the Passages sheet appears in the framework;
  - every framework section the workbook cites exists;
  - rows marked verbatim appear word for word (and case for case) in the framework;
  - every fact's framework wording appears in the framework (outside the 'Game facts' and 'Game text' tabs) and
    contains the fact's value.
Facts with no framework wording, and passages with no framework name, are listed as notes: they are engine
choices the framework does not describe yet.
"""
import os, re, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from game_text import read_game_text, framework_blocks, DEFAULT, FRAMEWORK, FACTS_TAB, TEXT_TAB

NUMBER_WORDS = {"once": 1, "one": 1, "two": 2, "twice": 2, "three": 3, "four": 4, "five": 5, "six": 6, "seven": 7,
                "eight": 8, "nine": 9, "ten": 10, "sixteen": 16, "twenty": 20, "forty": 40, "half": 0.5}


def norm(s):
    s = (s.replace("\u2018", "'").replace("\u2019", "'").replace("\u201c", '"').replace("\u201d", '"')
          .replace("\u00a0", " ").replace("\u00ad", "").replace("\u2011", "-"))
    return " ".join(s.split())


def framework_text(path):
    """The framework's text to check against, and its paragraphs: every tab but the ones the game text comes from as
    tables, which would otherwise vouch for themselves."""
    if not path.lower().endswith(".docx"):
        sys.exit(f"Can't read {path}: export the framework as .docx")
    parts = []
    for tab, style, content in framework_blocks(path):
        if tab in (FACTS_TAB, TEXT_TAB):
            continue
        cells = [c for row in content for c in row] if style == "Table" else [content]
        parts += [c.replace("**", "") for c in cells]
    return norm("\n".join(parts)), parts


def passage_list(full, paragraphs):
    """The names in the framework's 'Passages text:' line."""
    source = None
    if paragraphs:
        source = next((norm(p) for p in paragraphs if "Passages text" in p), None)
    if source is None:
        source = full
    m = re.search(r"Passages text\s*:\s*(.*?)(?:[●•▪◦○]|$)", source)
    if not m:
        return None
    return [n.strip(" .") for n in m.group(1).split(",") if n.strip(" .")]


def section_exists(full, sec):
    return re.search(r"(?<![\d.])" + re.escape(sec) + r"\.?\s+[A-Z]", full) is not None


def numbers_in(s):
    out = {float(n) for n in re.findall(r"\d+(?:\.\d+)?", s)}
    out |= {float(v) for w, v in NUMBER_WORDS.items() if re.search(rf"\b{w}\b", s, re.I)}
    return out


def check(game_text_path, export_path=FRAMEWORK):
    g, errors = read_game_text(game_text_path, export_path)
    if errors:
        return [f"The game text workbook has problems; fix those first ({len(errors)}, run game_text.py)"], []
    full, paragraphs = framework_text(export_path)
    problems, notes = [], []

    # Passages
    listed = passage_list(full, paragraphs)
    names = {p["Framework name"].lower(): p for p in g.passages if p["Framework name"]}
    if listed is None:
        problems.append("Couldn't find the framework's 'Passages text:' list (section 4)")
        listed = []
    for n in listed:
        if n.lower() not in names:
            problems.append(f"Passages: the framework lists the passage '{n}', but no screen on the Passages sheet has that framework name")
    for p in g.passages:
        fn = p["Framework name"]
        if not fn:
            notes.append(f"Passage '{p['Engine passage']}' is an engine addition the framework does not describe")
        elif fn.lower() not in {x.lower() for x in listed} and norm(fn).lower() not in full.lower():
            problems.append(f"Passages: '{p['Engine passage']}' is given the framework name '{fn}', which the framework does not use")

    # Sections
    cited = [(f"Passages '{p['Engine passage']}'", p["Framework section"]) for p in g.passages] + \
            [(f"Text {r['Key']}", r["Framework section"]) for r in g.rows] + \
            [(f"Facts {r['Key']}", r["Framework section"]) for r in g.fact_rows]
    for where, sec in cited:
        if sec and not section_exists(full, sec):
            problems.append(f"{where}: the framework has no section {sec}")

    # Verbatim text
    for r in g.rows:
        if r["Check"].lower() == "verbatim" and norm(r["Text"]) not in full:  # case matters for verbatim text
            problems.append(f"Text {r['Key']}: marked verbatim, but the framework does not contain \"{r['Text']}\"")

    # Facts
    for r in g.fact_rows:
        w = r["Framework wording"]
        if not w:
            notes.append(f"Fact '{r['Fact']}' = {r['Value']} is an engine choice the framework does not state")
            continue
        if norm(w).lower() not in full.lower():
            problems.append(f"Facts {r['Key']}: the framework does not contain the wording \"{w}\"")
        elif float(g.facts.get(r["Key"], "nan")) not in numbers_in(w):
            problems.append(f"Facts {r['Key']}: the value {r['Value']} is not what the framework wording says (\"{w}\")")
    return problems, notes


def main():
    args = sys.argv[1:] or [DEFAULT]
    if len(args) > 2:
        sys.exit(__doc__)
    problems, notes = check(*args)
    print("GAME TEXT DISAGREES WITH THE FRAMEWORK:" if problems else "OK: the game text matches the framework")
    for p in problems:
        print(" -", p)
    for n in notes:
        print(" · note:", n)
    sys.exit(1 if problems else 0)


if __name__ == "__main__":
    main()

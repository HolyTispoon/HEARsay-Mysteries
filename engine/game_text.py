#!/usr/bin/env python3
"""
Read the game's own text and the framework's numbers from the framework's 'HEARsay mysteries.docx' and from
'HEARsay game text.xlsx'.

    python3 game_text.py ["HEARsay game text.xlsx" ["HEARsay mysteries.docx"]]

The sources of truth are the framework Google Doc "HEARsay mysteries" and the Google Sheet "HEARsay game text"
(their addresses are in sources.json). The .docx and .xlsx next to this script are the repository's copies of
them, which only the build updates (see sources.py).

The framework Doc's tabs (each exported as a Title paragraph with the tab's name, up to the next tab):
  - 'Rules of the game': a heading (rules.title), then one paragraph per rule (rules.1, rules.2, …);
  - 'HEAR Framework introduction': a heading (hear.title), a paragraph (hear.intro), then for H, E, A and R a
    heading (hear.H.move …) and a paragraph (hear.H.text …), then a last paragraph (hear.unreceptive);
  - 'Game facts': a table with the columns Key, Fact, Value, Framework wording and Framework section, one row per
    number of the framework (check_framework.py checks the wording against the rest of the Doc);
  - 'Game text': a table with the columns Key, Screen, Text and Placeholders, for the game's sentences that
    restate the framework.
Bold words in the Doc become **double asterisks**, which the game shows in bold.
The Sheet holds the rest of the game's text (headings, buttons, labels) on its Text sheet, and the list of screens
on its Passages sheet. A key may come from only one of them, and the Sheet's old Facts sheet must be gone or empty.

read_game_text() returns (game_text, errors). game_text.text maps each key to its text, game_text.facts maps each
fact key to its value, and game_text.rows / .passages / .fact_rows keep the rows for check_framework.py.
Needs openpyxl and python-docx (pip3 install openpyxl python-docx).
"""
import os, re, sys
from types import SimpleNamespace

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from scenario_matrix import _sheet, _keyed

try:
    from openpyxl import load_workbook
except ImportError:
    sys.exit("This needs openpyxl: pip3 install openpyxl")

DEFAULT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "HEARsay game text.xlsx")  # the copy of the Google Sheet
FRAMEWORK = os.path.join(os.path.dirname(os.path.abspath(__file__)), "HEARsay mysteries.docx")  # the copy of the Google Doc
# The framework tabs the game reads
RULES_TAB, HEAR_TAB, FACTS_TAB, TEXT_TAB = "Rules of the game", "HEAR Framework introduction", "Game facts", "Game text"
FACT_COLUMNS = ["Key", "Fact", "Value", "Framework wording", "Framework section"]
TEXT_COLUMNS = ["Key", "Screen", "Text", "Placeholders"]
# Keys that come from the rules and HEAR tabs, so not from the Sheet
DOC_KEYS = re.compile(r"^(rules\.(title|\d+)|hear\.(title|intro|unreceptive|[HEAR]\.(move|text)))$")
# Placeholders every text can use; the engine fills them from the scenario and the facts.
GLOBAL_PLACEHOLDERS = {"budget", "n_witnesses", "n_questions", "n_major", "bonus", "rank_exact", "rank_near", "ccc", "title"}
CHECKS = {"verbatim", "paraphrase", "label"}
# Facts the build and the engine rely on.
REQUIRED_FACTS = ["witnesses", "questions", "question_budget", "tripwires_per_witness", "clues_per_witness",
                  "red_herrings_per_witness", "clues", "givers_per_clue", "red_herrings", "givers_per_red_herring",
                  "rh_set_min", "rh_set_max", "major_clues", "clues_per_major", "followups_per_witness", "submissions",
                  "followup_bonus", "rank_exact", "rank_near"]


def _paragraph_text(p):
    """A paragraph's text with its bold runs in **double asterisks** (bold that comes from a heading style is left
    alone)."""
    out = []
    for r in p.runs:
        t = r.text
        out.append(f"**{t.strip()}**" if r.bold and t.strip() and "**" not in t else t)
    t = " ".join("".join(out).split())
    return re.sub(r"\*\*\s*\*\*", " ", t).strip()  # two bold runs side by side make one bold phrase


def framework_blocks(path):
    """The framework export in order, as (tab, style, text) for each non-empty paragraph and (tab, "Table", rows)
    for each table, rows being lists of cell texts. Paragraphs before the first tab title have tab None."""
    import docx
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    d = docx.Document(path)
    blocks, tab = [], None
    for el in d.element.body.iterchildren():
        if el.tag.endswith("}p"):
            p = Paragraph(el, d)
            t = " ".join(p.text.split())
            if p.style.name == "Title":
                tab = t
            elif t:
                blocks.append((tab, p.style.name, _paragraph_text(p)))
        elif el.tag.endswith("}tbl"):
            rows = [[" ".join(_paragraph_text(p) for p in c.paragraphs).strip() for c in r.cells] for r in Table(el, d).rows]
            blocks.append((tab, "Table", rows))
    return blocks


def _table(blocks, tab, columns, errors):
    """The rows of the one table in a tab, as dicts keyed by column, with their row numbers in '_row'."""
    tables = [rows for t, style, rows in blocks if t == tab and style == "Table"]
    if len(tables) != 1:
        errors.append(f"The framework's '{tab}' tab must hold one table, not {len(tables)}")
        return []
    head = [h.replace("*", "").strip() for h in tables[0][0]]
    if head[:len(columns)] != columns:
        errors.append(f"The table in the framework's '{tab}' tab must have the columns {', '.join(columns)}, "
                      f"not {', '.join(head)}")
        return []
    plain = {"Key", "Value", "Screen", "Placeholders", "Framework section"}  # cells that are names or numbers, never bold
    return [dict({c: (r[i].replace("*", "").strip() if c in plain else r[i]) for i, c in enumerate(columns)}, _row=n)
            for n, r in enumerate(tables[0][1:], 2) if any(x.strip() for x in r)]


def _placeholders(where, text, allowed, errors):
    for ph in re.findall(r"\{(\w+)\}", text):
        if ph not in allowed:
            errors.append(f"{where}: {{{ph}}} is not a placeholder this text can use (allowed: {', '.join(sorted(allowed))})")


def framework_text(path, engine_passages):
    """From the framework's tabs: ({key: text}, fact rows, errors)."""
    if not os.path.exists(path):
        return {}, [], [f"No copy of the framework Doc ({os.path.basename(path)}): download the Google Doc as .docx "
                        "into the downloads folder and build with --downloads"]
    blocks = framework_blocks(path)
    tab = lambda name: [(style, t) for b, style, t in blocks if b == name and style != "Table"]
    text, errors = {}, []
    heading = lambda style: style.startswith("Heading")

    rules = tab(RULES_TAB)
    if len(rules) < 2 or not heading(rules[0][0]) or any(heading(s) for s, _ in rules[1:]):
        errors.append(f"The framework's '{RULES_TAB}' tab must hold a heading, then one paragraph per rule")
    else:
        text["rules.title"] = rules[0][1].replace("*", "")
        text.update((f"rules.{i}", t) for i, (_, t) in enumerate(rules[1:], 1))

    hear = tab(HEAR_TAB)
    if [heading(s) for s, _ in hear] != [True, False] + [True, False] * 4 + [False]:
        errors.append(f"The framework's '{HEAR_TAB}' tab must hold a heading and a paragraph, then a heading and a "
                      "paragraph for each of H, E, A and R, then one last paragraph (unreceptive language)")
    else:
        text["hear.title"], text["hear.intro"], text["hear.unreceptive"] = hear[0][1].replace("*", ""), hear[1][1], hear[-1][1]
        for i, L in enumerate("HEAR"):
            move = hear[2 + 2 * i][1].replace("*", "")
            text[f"hear.{L}.move"], text[f"hear.{L}.text"] = move, hear[3 + 2 * i][1]
            if not move.upper().startswith(L):
                errors.append(f"The framework's '{HEAR_TAB}' tab: the {L} heading must start with {L}, not {move!r}")
    for k, t in text.items():
        _placeholders(f"The framework's {k}", t, GLOBAL_PLACEHOLDERS, errors)

    for r in _table(blocks, TEXT_TAB, TEXT_COLUMNS, errors):
        k, where = r["Key"], f"The framework's '{TEXT_TAB}' row {r['_row']} ({r['Key'] or 'no key'})"
        if not k:
            errors.append(f"{where}: Key is blank"); continue
        if k in text or DOC_KEYS.match(k):
            errors.append(f"{where}: key appears twice in the framework"); continue
        if r["Screen"] not in engine_passages:
            errors.append(f"{where}: screen {r['Screen']!r} is not on the Sheet's Passages sheet")
        _placeholders(where, r["Text"], GLOBAL_PLACEHOLDERS | set(r["Placeholders"].replace(",", " ").split()), errors)
        text[k] = r["Text"]

    fact_rows = _table(blocks, FACTS_TAB, FACT_COLUMNS, errors)
    return text, fact_rows, errors


def read_game_text(path=DEFAULT, framework=FRAMEWORK):
    errors = []
    wb = load_workbook(path, data_only=True)

    passages = _sheet(wb, "Passages", ["Engine passage", "Framework name", "Framework section", "Notes"], errors)
    engine_passages = {p["Engine passage"] for p in passages}

    rows = _sheet(wb, "Text", ["Key", "Engine passage", "Text", "Framework section", "Check", "Placeholders"], errors)
    text = {}
    for r in rows:
        k, where = r["Key"], f"Text row {r['_row']} ({r['Key'] or 'no key'})"
        if not k:
            errors.append(f"Text row {r['_row']}: Key is blank"); continue
        if k in text:
            errors.append(f"{where}: key appears twice"); continue
        if r["Engine passage"] not in engine_passages:
            errors.append(f"{where}: engine passage {r['Engine passage']!r} is not on the Passages sheet")
        check = r["Check"].lower()
        if check not in CHECKS:
            errors.append(f"{where}: Check must be verbatim, paraphrase or label, not {r['Check']!r}")
        if check in ("verbatim", "paraphrase") and not r["Framework section"]:
            errors.append(f"{where}: a {check} row needs a Framework section")
        _placeholders(where, r["Text"], GLOBAL_PLACEHOLDERS | set(r["Placeholders"].split()), errors)
        if check == "verbatim" and "{" in r["Text"]:
            errors.append(f"{where}: a verbatim row cannot contain placeholders")
        text[k] = r["Text"]

    doc_text, doc_facts, doc_errors = framework_text(framework, engine_passages)
    errors += doc_errors
    for r in rows:
        if r["Key"] in doc_text or DOC_KEYS.match(r["Key"]):
            errors.append(f"Text row {r['_row']} ({r['Key']}): this text now comes from the framework Doc; delete the row")
    text.update(doc_text)
    if "Facts" in wb.sheetnames and any(v not in (None, "") for row in wb["Facts"].iter_rows(min_row=2, values_only=True) for v in row):
        errors.append(f"The Sheet's Facts sheet: the facts now come from the framework Doc's '{FACTS_TAB}' tab; "
                      "delete the sheet")

    fact_rows = _keyed(doc_facts, "Key", f"The framework's '{FACTS_TAB}'", errors)
    facts = {}
    for k, r in fact_rows.items():
        try:
            v = float(r["Value"])
            facts[k] = int(v) if v.is_integer() else v
        except ValueError:
            errors.append(f"The framework's fact {k}: Value must be a number, not {r['Value']!r}")
    for k in REQUIRED_FACTS:
        if k not in facts:
            errors.append(f"The framework's '{FACTS_TAB}': '{k}' is missing")
    return SimpleNamespace(text=text, facts=facts, rows=rows, passages=passages, fact_rows=list(fact_rows.values())), errors


def main():
    path = sys.argv[1] if len(sys.argv) > 1 else DEFAULT
    g, errors = read_game_text(path, sys.argv[2] if len(sys.argv) > 2 else FRAMEWORK)
    print(f"{len(g.text)} texts, {len(g.passages)} passages, {len(g.facts)} facts")
    print("PROBLEMS:" if errors else "OK: the game text passes every check")
    for e in errors:
        print(" -", e)
    sys.exit(1 if errors else 0)


if __name__ == "__main__":
    main()

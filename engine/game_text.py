#!/usr/bin/env python3
"""
Read the game's own text and the framework's numbers from 'HEARsay game text.xlsx', and the rules and the HEAR
guide from the framework's 'HEARsay mysteries.docx'.

    python3 game_text.py ["HEARsay game text.xlsx" ["HEARsay mysteries.docx"]]

The sources of truth are the Google Sheet "HEARsay game text" and, for the rules and the HEAR guide, two tabs of
the framework Google Doc "HEARsay mysteries" (their addresses are in sources.json). The .xlsx and .docx next to
this script are the repository's copies of them, which only the build updates (see sources.py).

The framework's tabs (each exported as a Title paragraph with the tab's name, up to the next tab):
  - 'Rules of the game': a heading (rules.title), then one paragraph per rule (rules.1, rules.2, …);
  - 'HEAR Framework introduction': a heading (hear.title), a paragraph (hear.intro), then for H, E, A and R a
    heading (hear.H.move …) and a paragraph (hear.H.text …), then a last paragraph (hear.unreceptive).
Those keys must not also be on the Sheet's Text sheet.

read_game_text() returns (game_text, errors). game_text.text maps each key to its text, game_text.facts maps each
fact key to its value, and game_text.rows / .passages / .fact_rows keep the Sheet's rows for check_framework.py.
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
RULES_TAB, HEAR_TAB = "Rules of the game", "HEAR Framework introduction"  # the framework tabs the game shows
# Keys that come from those tabs, so not from the Sheet
DOC_KEYS = re.compile(r"^(rules\.(title|\d+)|hear\.(title|intro|unreceptive|[HEAR]\.(move|text)))$")
# Placeholders every text can use; the engine fills them from the scenario and the facts.
GLOBAL_PLACEHOLDERS = {"budget", "n_witnesses", "n_questions", "n_major", "bonus", "rank_exact", "rank_near", "ccc", "title"}
CHECKS = {"verbatim", "paraphrase", "label"}
# Facts the build and the engine rely on.
REQUIRED_FACTS = ["witnesses", "questions", "question_budget", "tripwires_per_witness", "clues_per_witness",
                  "red_herrings_per_witness", "clues", "givers_per_clue", "red_herrings", "givers_per_red_herring",
                  "rh_set_min", "rh_set_max", "major_clues", "clues_per_major", "followups_per_witness", "submissions",
                  "followup_bonus", "rank_exact", "rank_near"]


def framework_tabs(path):
    """{tab name: [(style, text), …]} for every tab of the framework export, empty paragraphs left out."""
    import docx
    tabs, tab = {}, None
    for p in docx.Document(path).paragraphs:
        t = " ".join(p.text.split())
        if p.style.name == "Title":
            tab = tabs.setdefault(t, [])
        elif tab is not None and t:
            tab.append((p.style.name, t))
    return tabs


def framework_text(path):
    """The rules and the HEAR guide from the framework's tabs: ({key: text}, errors)."""
    if not os.path.exists(path):
        return {}, [f"No copy of the framework Doc ({os.path.basename(path)}): download the Google Doc as .docx into "
                    "the downloads folder and build with --downloads"]
    tabs, text, errors = framework_tabs(path), {}, []
    heading = lambda style: style.startswith("Heading")
    rules = tabs.get(RULES_TAB) or []
    if len(rules) < 2 or not heading(rules[0][0]) or any(heading(s) for s, _ in rules[1:]):
        errors.append(f"The framework's '{RULES_TAB}' tab must hold a heading, then one paragraph per rule")
    else:
        text["rules.title"] = rules[0][1]
        text.update((f"rules.{i}", t) for i, (_, t) in enumerate(rules[1:], 1))
    hear = tabs.get(HEAR_TAB) or []
    shape = [heading(s) for s, _ in hear]
    if shape != [True, False] + [True, False] * 4 + [False]:
        errors.append(f"The framework's '{HEAR_TAB}' tab must hold a heading and a paragraph, then a heading and a "
                      "paragraph for each of H, E, A and R, then one last paragraph (unreceptive language)")
    else:
        text["hear.title"], text["hear.intro"], text["hear.unreceptive"] = hear[0][1], hear[1][1], hear[-1][1]
        for i, L in enumerate("HEAR"):
            text[f"hear.{L}.move"], text[f"hear.{L}.text"] = hear[2 + 2 * i][1], hear[3 + 2 * i][1]
            if not hear[2 + 2 * i][1].upper().startswith(L):
                errors.append(f"The framework's '{HEAR_TAB}' tab: the {L} heading must start with {L}, not {hear[2 + 2 * i][1]!r}")
    for k, t in text.items():
        for ph in re.findall(r"\{(\w+)\}", t):
            if ph not in GLOBAL_PLACEHOLDERS:
                errors.append(f"The framework's {k}: {{{ph}}} is not a placeholder this text can use "
                              f"(allowed: {', '.join(sorted(GLOBAL_PLACEHOLDERS))})")
    return text, errors


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
        allowed = GLOBAL_PLACEHOLDERS | set(r["Placeholders"].split())
        for ph in re.findall(r"\{(\w+)\}", r["Text"]):
            if ph not in allowed:
                errors.append(f"{where}: {{{ph}}} is not a placeholder this text can use "
                              f"(allowed: {', '.join(sorted(allowed))})")
        if check == "verbatim" and "{" in r["Text"]:
            errors.append(f"{where}: a verbatim row cannot contain placeholders")
        text[k] = r["Text"]

    doc_text, doc_errors = framework_text(framework)
    errors += doc_errors
    for r in rows:
        if DOC_KEYS.match(r["Key"]):
            errors.append(f"Text row {r['_row']} ({r['Key']}): this text now comes from the framework Doc "
                          f"('{RULES_TAB}' or '{HEAR_TAB}' tab); delete the row")
    text.update(doc_text)

    fact_rows = _keyed(_sheet(wb, "Facts", ["Key", "Fact", "Value", "Framework wording", "Framework section"], errors), "Key", "Facts", errors)
    facts = {}
    for k, r in fact_rows.items():
        try:
            v = float(r["Value"])
            facts[k] = int(v) if v.is_integer() else v
        except ValueError:
            errors.append(f"Facts {k}: Value must be a number, not {r['Value']!r}")
    for k in REQUIRED_FACTS:
        if k not in facts:
            errors.append(f"Facts: '{k}' is missing")
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

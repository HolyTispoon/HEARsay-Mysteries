#!/usr/bin/env python3
"""
Read a HEARsay scenario from its Scenario Packet (.docx) and write it out as the Response Matrix workbook (.xlsx).

    python3 scenario_packet.py [--with-suggestions] "SCENARIO – Scenario Packet.docx" ["SCENARIO – Response Matrix.xlsx"]

The packet (a Google Doc) is the scenario's one source of truth. The matrix is generated from it on every build
and never edited: the build writes it next to the packet, and reads the game from it (scenario_matrix.py, which
also checks the framework's rules). With one argument this lists what the packet holds and any problem.

What the build reads from the packet's main tab (the tab holding the "Details" heading); everything else in the
packet (the timeline, the positions, the aftermath, the other tabs) is prose for people and is not read. Headings
are found by name, with or without a number ("4. The witnesses"):
  - Details: lines "Title: …", "Scenario ID: …", "IFID: …", and a paragraph "Introduction blurb: …" (the Start
    summary).
  - The Hook: paragraphs "Situation. …", "Assignment (…). …" and "Reminder: …" (the whole paragraph is the
    reminder); then a Heading 2 ending in ": <the CCC's name>", followed by the CCC summary, the Clause (the one
    paragraph that is entirely a quotation, optional) and the Clause explainer (every paragraph after it).
  - The witnesses: the tables with the columns ID, Name, Background, Position on the CCC; ID, Name, Public Facing
    Card; and ID, Name, Location, Portrait, Receptiveness rank, Receptive on (a portrait's file name is the text
    after the last "/"); then "W01 Name: …" paragraphs (the withheld secrets).
  - The Mystery: paragraphs "The mystery. …" and "The solution. …"; the tables with the columns ID, Group, Clue,
    Given by and ID, Question, Tripwire for; the Major Clues, each a paragraph "MC1 (H) — Title. Synthesizes C01,
    C02, C03, C04." and then the paragraph shown to the player; and the final-answer form, the first table after
    the paragraph starting "Final-answer form": its two column headings are the two questions the player answers,
    and the correct person and place are each marked "(correct)".
  - The Red Herrings: "Set S1 — Title. Resolution: …" paragraphs, each followed by its "R01: … (given by Name
    (Q6))" paragraphs.
  - Follow-up pairs: a Heading 2 per witness name, then "Receptive option. …", "Unreceptive option. …",
    "Explanation shown on success. …" and "Walk-away on failure. …" (quotation marks around a whole option are
    dropped).
  - The Response Matrix: the overview grid (a checked copy), then a Heading 2 per witness name and one paragraph per
    question: "Q1 [C13, clue; opinion receptive]. <answer>" or "Q8 — TRIPWIRE. <walk-away text>".
  - The ending: paragraphs "Correct. …", "Right person, wrong place. …", "Wrong. …", "Out of questions. …", and
    "Map file: …", "Logo file: …", "Masthead file: …", "Start image file: …", "Newspaper illustration file: …"
    (the map is required, the others optional).
The text is the packet's accepted text: a pending suggestion is not read until the author accepts it
(--with-suggestions reads the packet as if they were all accepted, to check them first). Every part
the build needs must be there exactly once; anything it can't read is a problem, and the build stops.

Needs openpyxl and python-docx (pip3 install openpyxl python-docx).
"""
import os, re, sys

try:
    import docx
    from docx.oxml.ns import qn
    from docx.table import Table
    from docx.text.paragraph import Paragraph
except ImportError:
    sys.exit("This needs python-docx: pip3 install python-docx")

# The matrix workbook: each sheet's columns, in order (scenario_matrix.py finds them by these names)
SHEETS = {
    "Settings": ["Setting", "Value"],
    "Passages": ["Passage", "Part", "Text"],
    "Mystery": ["Field", "Text"],
    "Final-answer form": ["Kind", "Option", "Correct"],
    "Grid": None,  # Question, then one column per witness name
    "Long table": ["Witness ID", "Witness", "Question ID", "Question", "Item", "Item type", "Group / set", "Opinion receptive", "Answer text"],
    "Witnesses": ["ID", "Name", "Background", "Public background card", "Position on the CCC", "Withheld secret", "Tripwire",
                  "Tripwire question", "Walk-away text", "Receptiveness rank", "Receptive on", "Location", "Portrait"],
    "Questions": ["ID", "Question", "Tripwire for"],
    "Clues": ["ID", "Major clue group", "Clue", "Given by"],
    "Major Clues": ["ID", "Letter", "Title", "Clues", "Text shown to player"],
    "Red Herrings": ["ID", "Set", "Set title", "Red Herring", "Given by", "How the set is discounted"],
    "Follow-ups": ["Witness ID", "Witness", "Receptive option", "Unreceptive option", "Explanation on success", "Walk-away on failure"],
}
READ_ME = ("Generated from the Scenario Packet (the Google Doc) by the build; don't edit it. The packet is the "
           "scenario's source of truth: change the packet and rebuild. engine/scenario_packet.py says what the build "
           "reads from the packet; scenario_matrix.py reads this workbook and checks the framework's rules.")

# The packet's Heading 1 sections the build reads, by a pattern of their names
SECTIONS = {"details": r"details", "hook": r"the hook", "witnesses": r"the witnesses", "mystery": r"the mystery",
            "major": r"(the )?(\w+ )?major clues", "red herrings": r"(the )?(\d+ )?red herrings.*",
            "follow-ups": r"follow-up pairs", "answers": r"the response matrix", "ending": r"the ending.*"}
WITNESS_TABLES = {"Background": ["ID", "Name", "Background", "Position on the CCC"],
                  "Card": ["ID", "Name", "Public Facing Card"],
                  "Technical": ["ID", "Name", "Location", "Portrait", "Receptiveness rank", "Receptive on"]}
CLUE_TABLE = ["ID", "Group", "Clue", "Given by"]
QUESTION_TABLE = ["ID", "Question", "Tripwire for"]
FOLLOW_UP_LABELS = {"Receptive option": "good", "Unreceptive option": "bad", "Explanation shown on success": "why",
                    "Walk-away on failure": "walkout"}
ENDING_LABELS = [("Correct", "ending", "Correct"), ("Right person, wrong place", "ending", "Right person, wrong place"),
                 ("Wrong", "ending", "Wrong"), ("Out of questions", "Submission", "Out of questions")]
FILE_SETTINGS = [("Map", True), ("Logo", False), ("Masthead", False), ("Start image", False), ("Newspaper illustration", False)]
CORRECT = re.compile(r"\s*\(correct\)\s*$", re.I)


def _text(el, pending=False):
    """An element's accepted text: suggested insertions (inside w:ins) left out, suggested deletions (w:delText)
    kept; with pending=True, its text as if every suggestion were accepted. Line breaks as newlines, every run of
    spaces as one space."""
    out = []
    for x in el.iter(qn("w:t"), qn("w:delText"), qn("w:br"), qn("w:cr"), qn("w:tab"), qn("w:noBreakHyphen"), qn("w:p")):
        if x.tag == qn("w:p"):
            out.append("\n\n")  # a new paragraph inside a table cell
        elif x.tag == qn("w:t"):
            if pending or not any(a.tag == qn("w:ins") for a in x.iterancestors()):
                out.append(x.text or "")
        elif x.tag == qn("w:delText"):
            if not pending:
                out.append(x.text or "")
        else:
            out.append({qn("w:br"): "\n", qn("w:cr"): "\n", qn("w:tab"): " ", qn("w:noBreakHyphen"): "-"}[x.tag])
    lines = [" ".join(line.split()) for line in "".join(out).split("\n")]
    return re.sub(r"\n{3,}", "\n\n", "\n".join(lines)).strip()


def packet_blocks(path, pending=False):
    """The packet's main tab in order: ("p", style, text) for each paragraph with text and ("table", rows) for each
    table. Google's .docx export ends each tab with a section break; the main tab is the one with a "Details"
    heading."""
    d = docx.Document(path)
    tabs, blocks = [], []
    for el in d.element.body.iterchildren():
        if el.tag == qn("w:p"):
            t = _text(el, pending)
            if t:
                blocks.append(("p", Paragraph(el, d).style.name, t))
            if el.pPr is not None and el.pPr.find(qn("w:sectPr")) is not None:
                tabs.append(blocks)
                blocks = []
        elif el.tag == qn("w:tbl"):
            blocks.append(("table", [[_text(c._tc, pending) for c in r.cells] for r in Table(el, d).rows]))
    tabs.append(blocks)
    main = [b for b in tabs if any(x[0] == "p" and x[1] == "Heading 1" and _heading(x[2]) == "details" for x in b)]
    return main[0] if len(main) == 1 else None


def _heading(t):
    """The section key of a Heading 1, or None."""
    name = re.sub(r"^\d+(\.\d+)*\.?\s*", "", t).strip().lower()
    return next((k for k, rx in SECTIONS.items() if re.fullmatch(rx, name)), None)


def _unquote(t):
    return t[1:-1].strip() if len(t) > 1 and t[0] in "“\"" and t[-1] in "”\"" else t


def read_packet(path, pending=False):
    """Return (sheets, errors): sheets maps each matrix sheet name to its rows, the first row being the headings.
    pending=True reads the packet as if its pending suggestions were accepted (to check them before they are)."""
    errors = []
    blocks = packet_blocks(path, pending)
    if blocks is None:
        return {}, ["The packet must have exactly one tab with a 'Details' Heading 1 (its main tab)"]
    sections, key = {}, None
    for b in blocks:
        if b[0] == "p" and b[1] == "Heading 1":
            key = _heading(b[2])
            if key in sections:
                errors.append(f"The packet has two '{b[2]}' sections")
            sections.setdefault(key, [])
            continue
        if key:
            sections[key].append(b)
    for k in SECTIONS:
        if k not in sections:
            errors.append(f"The packet has no '{SECTIONS[k]}' Heading 1 section")
            sections[k] = []
    paras = lambda k: [b for b in sections[k] if b[0] == "p"]

    def one(k, rx, what, required=True, flags=re.S):
        """The one paragraph (or line, for Details) of section k that matches rx; its groups."""
        found = []
        for b in paras(k):
            for line in (b[2].split("\n") if k == "details" else [b[2]]):
                m = re.match(rx, line, flags)
                if m:
                    found.append(m)
        if len(found) > 1:
            errors.append(f"{what}: the packet has {len(found)} of them; keep one")
        elif not found and required:
            errors.append(f"{what}: not found in the packet")
        return found[0] if len(found) == 1 else None

    def table(columns, what):
        """The one table in the main tab with these columns, as dicts."""
        found = [b[1] for b in blocks if b[0] == "table" and b[1] and [c.strip() for c in b[1][0]] == columns]
        if len(found) != 1:
            errors.append(f"{what}: the packet needs one table with the columns {', '.join(columns)} (found {len(found)})")
            return []
        return [dict(zip(columns, r)) for r in found[0][1:] if any(x.strip() for x in r)]

    g = lambda m, i=1: m.group(i).strip() if m else ""
    # Details
    settings = [("Title", g(one("details", r"Title:\s*(.+)$", "Details: 'Title:'"))),
                ("Scenario ID", g(one("details", r"Scenario ID:\s*(.+)$", "Details: 'Scenario ID:'"))),
                ("IFID", g(one("details", r"IFID:\s*(.+)$", "Details: 'IFID:'")))]
    files = []
    for name, required in FILE_SETTINGS:
        found = [m.group(1) for b in paras("ending") for m in re.finditer(rf"\b{name} file:\s*([^\s()]+?)[.,;]?(?=\s|$)", b[2])]
        if len(found) > 1:
            errors.append(f"'{name} file:': the packet names it {len(found)} times; keep one")
        elif not found and required:
            errors.append(f"'{name} file:': not found in the packet's ending section")
        elif found:
            files.append((name, found[0]))
    settings += files

    # Passages
    passages = [("Start", "Summary", g(one("details", r"Introduction blurb\s*:\s*(.+)$", "Details: 'Introduction blurb:'")))]
    hook = paras("hook")
    h2 = [i for i, b in enumerate(hook) if b[1] == "Heading 2"]
    intro = hook[:h2[0]] if h2 else hook
    label = lambda rx, what: next((g(re.match(rx, b[2], re.S)) for b in intro if re.match(rx, b[2], re.S)), None) or \
        errors.append(f"the Hook: '{what}' not found in the packet") or ""
    passages += [("the Hook", "Situation", label(r"Situation\.\s*(.+)", "Situation.")),
                 ("the Hook", "Assignment", label(r"Assignment(?:\s*\([^)]*\))?\.\s*(.+)", "Assignment.")),
                 ("the Hook", "Reminder", next((b[2] for b in intro if b[2].startswith("Reminder:")), ""))]
    if len(h2) != 1 or ":" not in hook[h2[0]][2]:
        errors.append("the Hook: needs one Heading 2 ending in ': <the CCC's name>'")
    else:
        ccc = [b[2] for b in hook[h2[0] + 1:]]
        quoted = [i for i, t in enumerate(ccc) if t[0] in "“\"" and t[-1] in "”\""]
        if len(quoted) > 1:
            errors.append("the Hook: the CCC can have one quoted Clause, but the packet has several quoted paragraphs")
        cut = quoted[0] if len(quoted) == 1 else len(ccc)
        passages += [("the Hook", "CCC", hook[h2[0]][2].rsplit(":", 1)[1].strip()),
                     ("the Hook", "CCC summary", "\n\n".join(ccc[:cut])),
                     ("the Hook", "Clause", ccc[cut] if quoted else ""),
                     ("the Hook", "Clause explainer", "\n\n".join(ccc[cut + 1:]))]
    endings = []
    for lab, pas, part in ENDING_LABELS:
        endings.append((pas, part, g(one("ending", rf"{re.escape(lab)}\.\s*(.+)", f"Ending: '{lab}.'"))))
    passages += [endings[3]] + endings[:3]

    # Witnesses
    tw = {k: {r["ID"].strip(): r for r in table(cols, f"Witnesses table ({k})")} for k, cols in WITNESS_TABLES.items()}
    ids = list(tw["Background"])
    for k, rows in tw.items():
        if list(rows) != ids:
            errors.append(f"The witness tables must list the same IDs in the same order ({k} table: {', '.join(rows)})")
    names = {i: tw["Background"][i]["Name"].strip() for i in ids}
    for k, rows in tw.items():
        for i, r in rows.items():
            if i in names and r["Name"].strip() != names[i]:
                errors.append(f"Witness {i}: the {k} table names {r['Name']!r}, the first table {names[i]!r}")
    by_name = {n: i for i, n in names.items()}
    secrets = {}
    for b in paras("witnesses"):
        m = re.match(r"(W\d+)\s+([^:]+):\s*(.+)", b[2], re.S)
        if m and m.group(1) in names:
            if m.group(2).strip() != names[m.group(1)]:
                errors.append(f"Withheld secret {m.group(1)}: named {m.group(2)!r}, but {m.group(1)} is {names[m.group(1)]}")
            if m.group(1) in secrets:
                errors.append(f"Withheld secret {m.group(1)}: appears twice")
            secrets[m.group(1)] = m.group(3).strip()

    # The Mystery
    mystery = [("The mystery", g(one("mystery", r"The mystery\.\s*(.+)", "'The mystery.'"))),
               ("The solution", g(one("mystery", r"The solution\.\s*(.+)", "'The solution.'", flags=re.S | re.I)))]
    clues = table(CLUE_TABLE, "Clues table")
    questions = table(QUESTION_TABLE, "Questions table")
    Q = {q["ID"].strip(): q["Question"].strip() for q in questions}
    tripwire_of = {}
    for q in questions:
        who = q["Tripwire for"].strip()
        if who not in by_name:
            errors.append(f"Questions table {q['ID']}: 'Tripwire for' must be a witness's name, not {who!r}")
        elif by_name[who] in tripwire_of:
            errors.append(f"Questions table: {who} is the tripwire of two questions")
        else:
            tripwire_of[by_name[who]] = q["ID"].strip()
    start = [i for i, b in enumerate(blocks) if b[0] == "p" and b[2].startswith("Final-answer form")]
    form = next((b[1] for b in blocks[start[0]:] if b[0] == "table"), None) if len(start) == 1 else None
    final, labels = [], ["", ""]
    if not form or len(form[0]) != 2:
        errors.append("Final-answer form: the packet needs a two-column table after the 'Final-answer form' paragraph")
    else:
        labels = [h.strip() for h in form[0]]
        for col, kind in enumerate(("Person", "Place")):
            for r in form[1:]:
                o = r[col].strip()
                if o:
                    final.append((kind, CORRECT.sub("", o), "yes" if CORRECT.search(o) else ""))
    mystery += [("Person question", labels[0]), ("Place question", labels[1])]

    # Major Clues
    major = []
    mp = paras("mystery") + paras("major")
    for i, b in enumerate(mp):
        m = re.match(r"(MC\d+)\s*\((\w)\)\s*[—–-]\s*(.+?)\.\s*Synthesizes\s+(.+?)\.?$", b[2], re.S)
        if m:
            nxt = mp[i + 1][2] if i + 1 < len(mp) and not re.match(r"MC\d+\s*\(", mp[i + 1][2]) else ""
            if not nxt:
                errors.append(f"Major Clue {m.group(1)}: no paragraph after its heading line (the text shown to the player)")
            major.append([m.group(1), m.group(2), m.group(3).strip(), m.group(4).strip(), nxt])

    # Red Herrings
    herrings, current = [], None
    for b in paras("red herrings"):
        m = re.match(r"Set\s+(S\d+)\s*[—–-]\s*(.+?)\.\s*Resolution:\s*(.+)", b[2], re.S)
        if m:
            current = m.groups()
            continue
        m = re.match(r"(R\d+):\s*(.+?)\s*\(given by\s+(.+)\)\s*\.?$", b[2], re.S)
        if m:
            if not current:
                errors.append(f"Red Herring {m.group(1)}: comes before any 'Set S1 — Title. Resolution: …' paragraph")
                continue
            herrings.append([m.group(1), current[0], current[1].strip(), m.group(2), m.group(3).strip(), current[2].strip()])
        elif re.match(r"R\d+\b", b[2]):
            errors.append(f"Red Herring {b[2][:40]!r}…: must read 'R01: <text> (given by Name (Q1))'")

    # Follow-ups and answers: a Heading 2 per witness
    def by_witness(k):
        out, who = {}, None
        for b in paras(k):
            if b[1] == "Heading 2":
                who = by_name.get(b[2].strip())
                if not who:
                    errors.append(f"'{b[2]}' (a Heading 2 in the packet's {SECTIONS[k]} section) is not a witness's name")
                elif who in out:
                    errors.append(f"{b[2]} has two Heading 2s in the packet's {SECTIONS[k]} section")
                out.setdefault(who, []) if who else None
            elif who:
                out[who].append(b[2])
        return out
    follow_ups = []
    for wid, ps in by_witness("follow-ups").items():
        f = {}
        for p in ps:
            for lab, k in FOLLOW_UP_LABELS.items():
                if p.startswith(lab + "."):
                    if k in f:
                        errors.append(f"Follow-up {names[wid]}: '{lab}.' appears twice")
                    f[k] = p[len(lab) + 1:].strip()
        for lab, k in FOLLOW_UP_LABELS.items():
            if k not in f:
                errors.append(f"Follow-up {names[wid]}: '{lab}.' not found")
        follow_ups.append([wid, names[wid], _unquote(f.get("good", "")), _unquote(f.get("bad", "")), f.get("why", ""), f.get("walkout", "")])

    answers, walkaway = {}, {}
    for wid, ps in by_witness("answers").items():
        for p in ps:
            m = re.match(r"(Q\d+)\s*\[\s*([CR]\d+)\s*,\s*(clue|red herring)\s*;\s*opinion\s+(receptive|unreceptive)\s*\]\.?\s*(.+)", p, re.S | re.I)
            t = re.match(r"(Q\d+)\s*[—–-]\s*TRIPWIRE\.\s*(.+)", p, re.S)
            if m:
                if (wid, m.group(1)) in answers:
                    errors.append(f"{names[wid]} {m.group(1)}: answered twice in the packet")
                answers[(wid, m.group(1))] = [m.group(2), m.group(3).lower(), m.group(4).lower(), m.group(5)]
            elif t:
                if t.group(1) != tripwire_of.get(wid):
                    errors.append(f"{names[wid]}: the packet marks {t.group(1)} as the tripwire, but the Questions table "
                                  f"says {tripwire_of.get(wid) or 'none'}")
                walkaway[wid] = t.group(2)
            elif re.match(r"Q\d+\b", p):
                errors.append(f"{names[wid]}: can't read {p[:50]!r}…; an answer reads 'Q1 [C01, clue; opinion receptive]. <text>' "
                              "and a tripwire 'Q1 — TRIPWIRE. <text>'")

    # The workbook
    clue_group = {c["ID"].strip(): c["Group"].strip() for c in clues}
    rh_set = {r[0]: r[1] for r in herrings}
    long_rows, grid = [], {}
    for wid in ids:
        for qid in Q:
            if tripwire_of.get(wid) == qid:
                long_rows.append([wid, names[wid], qid, Q[qid], "TRIPWIRE", "tripwire", "", "", walkaway.get(wid, "")])
                grid[(wid, qid)] = "TW — " + walkaway.get(wid, "")
                continue
            a = answers.get((wid, qid))
            if not a:
                continue
            item, kind, flag, text = a
            group = clue_group.get(item) if kind == "clue" else rh_set.get(item)
            long_rows.append([wid, names[wid], qid, Q[qid], item, kind, group or "", "yes" if flag == "receptive" else "no", text])
            grid[(wid, qid)] = f"[{item} | {flag}] {text}"
    receptive = {wid: [r[2] for r in long_rows if r[0] == wid and r[7] == "yes"] for wid in ids}
    for wid in ids:
        listed = [x for x in re.split(r"[,\s;]+", tw["Technical"][wid]["Receptive on"]) if x and x.lower() != "none"] if wid in tw["Technical"] else []
        if sorted(listed) != sorted(receptive[wid]):
            errors.append(f"{names[wid]}: the witness table says receptive on {', '.join(listed) or 'none'}, but the answers "
                          f"are labelled receptive on {', '.join(receptive[wid]) or 'none'}")
    check_grid(sections["answers"], ids, names, Q, tripwire_of, answers, errors)

    T = lambda k, i, c: tw[k][i][c].strip() if i in tw[k] else ""
    sheets = {
        "Read me": [["Sheet", "What it holds"], ["", READ_ME]],
        "Settings": [SHEETS["Settings"]] + [list(s) for s in settings],
        "Passages": [SHEETS["Passages"]] + [list(p) for p in passages],
        "Mystery": [SHEETS["Mystery"]] + [list(m) for m in mystery],
        "Final-answer form": [SHEETS["Final-answer form"]] + [list(f) for f in final],
        "Grid": [["Question"] + [names[w] for w in ids]] + [[f"{q}. {Q[q]}"] + [grid.get((w, q), "") for w in ids] for q in Q],
        "Long table": [SHEETS["Long table"]] + long_rows,
        "Witnesses": [SHEETS["Witnesses"]] + [[w, names[w], T("Background", w, "Background"), T("Card", w, "Public Facing Card"),
                                               T("Background", w, "Position on the CCC"), secrets.get(w, ""), tripwire_of.get(w, ""),
                                               Q.get(tripwire_of.get(w), ""), walkaway.get(w, ""), T("Technical", w, "Receptiveness rank"),
                                               ", ".join(receptive[w]) or "none", T("Technical", w, "Location"),
                                               T("Technical", w, "Portrait").rsplit("/", 1)[-1].strip()] for w in ids],
        "Questions": [SHEETS["Questions"]] + [[q["ID"].strip(), q["Question"].strip(), q["Tripwire for"].strip()] for q in questions],
        "Clues": [SHEETS["Clues"]] + [[c["ID"].strip(), c["Group"].strip(), c["Clue"].strip(), c["Given by"].strip()] for c in clues],
        "Major Clues": [SHEETS["Major Clues"]] + major,
        "Red Herrings": [SHEETS["Red Herrings"]] + herrings,
        "Follow-ups": [SHEETS["Follow-ups"]] + follow_ups,
    }
    for w in ids:
        if w not in secrets:
            errors.append(f"Withheld secret {w}: no 'W01 Name: …' paragraph in the packet's witnesses section")
        if w not in walkaway:
            errors.append(f"{names[w]}: no 'Q… — TRIPWIRE. …' paragraph in the packet's Response Matrix section")
    return sheets, errors


def check_grid(blocks, ids, names, Q, tripwire_of, answers, errors):
    """The packet's overview grid (a table headed Question, then each witness's first name) must match the answers."""
    grids = [b[1] for b in blocks if b[0] == "table" and b[1] and b[1][0][0].strip() == "Question"]
    if len(grids) != 1:
        errors.append(f"The Response Matrix section needs one overview grid (a table whose first heading is 'Question'), found {len(grids)}")
        return
    head = [h.strip() for h in grids[0][0][1:]]
    cols = {}
    for i, h in enumerate(head):
        wid = next((w for w in ids if names[w] == h or names[w].split()[0] == h), None)
        if not wid:
            errors.append(f"Overview grid: column {h!r} is not a witness")
        cols[i] = wid
    for r in grids[0][1:]:
        qid = r[0].strip().split(".")[0].strip()
        if qid not in Q:
            errors.append(f"Overview grid: row {r[0]!r} is not a question ID")
            continue
        for i, cell in enumerate(r[1:]):
            wid = cols.get(i)
            if not wid:
                continue
            a = answers.get((wid, qid))
            want = "TW" if tripwire_of.get(wid) == qid else (f"{a[0]} {'R' if a[2] == 'receptive' else 'U'}" if a else "?")
            if " ".join(cell.split()) != want:
                errors.append(f"Overview grid {qid} × {names[wid]}: says {cell.strip()!r}, but the answer below is {want!r}")


def write_matrix(sheets, path):
    """Write the matrix workbook (the same packet always gives the same file)."""
    from scenario_matrix import write_workbook
    write_workbook(sheets, path)


def main():
    args = [a for a in sys.argv[1:] if a != "--with-suggestions"]
    if len(args) not in (1, 2):
        sys.exit(__doc__)
    sheets, errors = read_packet(args[0], pending="--with-suggestions" in sys.argv)
    if len(args) == 2 and not errors:
        write_matrix(sheets, args[1])
        print(f"Wrote {args[1]}")
    else:
        for name, rows in sheets.items():
            print(f"{name}: {len(rows) - 1} rows")
    print("PROBLEMS:" if errors else "OK: the packet can be read")
    for e in errors:
        print(" -", e)
    sys.exit(1 if errors else 0)


if __name__ == "__main__":
    main()

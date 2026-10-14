#!/usr/bin/env python3
"""
Build a HEARsay scenario into a playable Twine (SugarCube 2) game.

    python3 build_twine.py "SCENARIO – Response Matrix.xlsx" --assets ASSETS_DIR --out OUT_DIR [--repo]
                           [--packet "SCENARIO – Scenario Packet.docx" | --no-packet-check]
                           [--framework "HEARsay mysteries.docx" | --no-framework-check]

Three workbooks and two documents go into a game:
  - the framework Google Doc ("HEARsay mysteries"): the source of truth for the game's structure and rules;
  - 'HEARsay game text.xlsx' (next to this script): the game's own text and the framework's numbers,
    checked against an export of the framework;
  - the scenario's Response Matrix (.xlsx): the scenario, checked against the framework's numbers;
  - the scenario's packet (.docx): the scenario written up for people, checked against the matrix.

Steps:
  1. Reads the game text and checks it against the framework export (check_framework.py). The export is the
     newest 'HEARsay mysteries' .docx/.txt/.pdf in the folder above this script's folder, or --framework.
     If there is none (as in a published repository) the check is skipped with a warning.
  2. Reads the matrix and stops if it breaks a framework rule or disagrees with itself (scenario_matrix.py).
  3. Checks that the packet says the same things as the matrix (check_packet.py). The packet is the .docx
     next to the matrix whose name contains "Packet", or --packet.
  4. Writes the scenario and the game text into the engine as `setup.scenario` and `setup.gameText`,
     and checks that every text the engine asks for is in the game text workbook.
  5. Writes OUT_DIR/source/story.twee (Twee 3) and compiles it with Tweego into OUT_DIR/index.html.
  6. Copies the witness portraits, the campus map and the site icons (ICONS) into OUT_DIR/assets/; head.html
     links the icons from the page's <head>.
  With --repo it also copies the engine, the game text, the matrix, the packet and a README into OUT_DIR,
  so OUT_DIR is a complete, self-contained GitHub repository (GitHub Pages serves index.html).
Every check stops the build and lists each problem.

Needs openpyxl and python-docx (pip3 install openpyxl python-docx), and pypdf for a PDF framework export.

Tweego: https://www.motoslave.net/tweego/ . Put `tweego` on your PATH, or set TWEEGO=/path/to/tweego.
Tweego must have the SugarCube 2 story format installed (it ships with it).
"""
import argparse, glob, json, os, re, shutil, subprocess, sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from scenario_matrix import read_matrix, PASSAGE_PARTS
from check_packet import check as check_packet
from game_text import read_game_text, DEFAULT as GAME_TEXT
from check_framework import check as check_framework, find_export

ENGINE = os.path.join(HERE, "hearsay_engine.twee")
# Appended to the <head> of index.html; it links the site icons, which are the same for every scenario
HEAD = os.path.join(HERE, "head.html")
ICONS = ("favicon.svg", "favicon-32.png", "apple-touch-icon.png")
# Text keys the engine builds at run time rather than naming outright
DYNAMIC_KEYS = re.compile(r"^(rules\.\d+|hear\.[HEAR]\.(move|text)|result\.\w+|ending\.verdict\.\w+)$")


def load_game_text(framework, skip):
    g, errors = read_game_text(GAME_TEXT)
    if errors:
        sys.exit("The game text workbook has problems; fix them before building:\n" + "\n".join(" - " + e for e in errors))
    if skip:
        print("Skipping the framework check (--no-framework-check).")
        return g
    note = None
    if not framework:
        framework, note = find_export(os.path.dirname(HERE))
    if not framework:
        print("WARNING: no export of the framework ('HEARsay mysteries' .docx/.txt/.pdf) found, so the game text "
              "was not checked against it. Export the Google Doc next to the .gdoc file, or pass --framework.")
        return g
    problems, notes = check_framework(GAME_TEXT, framework)
    # the matrix's passage names must be framework names too
    from check_framework import framework_text, passage_list
    listed = [n.lower() for n in (passage_list(*framework_text(framework)) or [])]
    for pas in sorted({p for p, _ in PASSAGE_PARTS}):
        if pas.lower() not in listed:
            problems.append(f"The matrix files scenario text under '{pas}', which is not in the framework's passage list")
    if problems:
        sys.exit(f"The game text disagrees with the framework ({os.path.basename(framework)}); make them match before building:\n"
                 + "\n".join(" - " + p for p in problems))
    print(f"Framework OK: the game text matches {os.path.basename(framework)}"
          + (f" ({len(notes)} engine choices the framework does not describe; run check_framework.py to list them)" if notes else ""))
    if note:
        print("WARNING: " + note)
    return g


def check_engine_keys(g):
    with open(ENGINE, encoding="utf-8") as f:
        engine = f.read()
    # literal keys only; keys built at run time (T("hear." + L + ".move")) are covered by DYNAMIC_KEYS
    used = {k for k in re.findall(r"""\b[TR]\(\s*["'`]([\w.]+)["'`]""", engine) if not k.endswith(".")}
    missing = sorted(k for k in used if k not in g.text)
    if missing:
        sys.exit("The engine uses text that is not in the game text workbook: " + ", ".join(missing))
    unused = sorted(k for k in g.text if k not in used and not DYNAMIC_KEYS.match(k))
    if unused:
        print("Note: the game text workbook has text the engine never shows: " + ", ".join(unused))


def load(matrix_path, facts):
    d, errors = read_matrix(matrix_path, facts)
    if errors:
        sys.exit("The matrix has problems; fix them before building:\n" + "\n".join(" - " + e for e in errors))
    print(f"Matrix OK: {len(d.WITNESSES)} witnesses, {len(d.CLUES)} clues, {len(d.RED_HERRINGS)} red herrings, budget {d.QUESTION_BUDGET}")
    return d


def find_packet(matrix_path):
    found = [p for p in glob.glob(os.path.join(os.path.dirname(os.path.abspath(matrix_path)), "*.docx"))
             if "packet" in os.path.basename(p).lower() and not os.path.basename(p).startswith("~$")]
    if len(found) != 1:
        sys.exit(f"Expected one packet .docx next to the matrix, found {len(found)}. Name it with --packet, or pass --no-packet-check.")
    return found[0]


def scenario_json(d, asset_dir="assets"):
    p = d.PUZZLE
    for key in ("who", "where"):
        opts = p[f"{key}_options"]
        if p["final_answer"][key] not in opts:
            sys.exit(f"PUZZLE['final_answer']['{key}'] is not one of PUZZLE['{key}_options']")
    # Only what the game needs. Designer-only fields (secrets, receptive flags, tripwire logic) stay out.
    return {
        "id": d.SCENARIO_ID,
        "title": d.TITLE,
        "pitch": d.PITCH or p["mystery"],
        "budget": d.QUESTION_BUDGET,
        "followup_bonus": d.FOLLOWUP_BONUS,
        "asset_dir": asset_dir,
        "map": d.MAP,
        "rank_exact": d.FACTS["rank_exact"],
        "rank_near": d.FACTS["rank_near"],
        "puzzle": {
            "mystery": p["mystery"],
            "solution": p["solution"],
            "who_label": p["who_label"],
            "where_label": p["where_label"],
            "final_answer": p["final_answer"],
            "who_options": p["who_options"],
            "where_options": p["where_options"],
        },
        "controversy": {k: d.CONTROVERSY[k] for k in ("name", "summary", "clause", "clause_explainer")},
        "setting": {"player_role": d.SETTING["player_role"], "briefing": d.SETTING["briefing"]},
        "questions": [{"id": q["id"], "text": q["text"]} for q in d.QUESTIONS],
        "witnesses": [{k: w[k] for k in ("id", "name", "portrait", "role", "public_card", "location", "tripwire")}
                      for w in d.WITNESSES],
        "responses": {wid: {qid: {"item": a["item"], "text": a["text"]} for qid, a in qs.items()}
                      for wid, qs in d.RESPONSES.items()},
        "tripwire_responses": d.TRIPWIRE_RESPONSES,
        "follow_ups": d.FOLLOW_UPS,
        "clues": {k: v["text"] for k, v in d.CLUES.items()},
        "red_herrings": {k: v["text"] for k, v in d.RED_HERRINGS.items()},
        "red_herring_sets": d.RED_HERRING_SETS,
        "major_clues": [{k: m[k] for k in ("id", "letter", "title", "text")} for m in d.MAJOR_CLUES],
        "endings": d.ENDINGS,
        "ranking_answer": d.BONUS_RANKING["answer"],
    }


def find_file(root, name):
    for dp, _, files in os.walk(root):
        if name in files:
            return os.path.join(dp, name)
    sys.exit(f"Asset {name} not found under {root}")


def same(a, b):
    return os.path.exists(b) and os.path.samefile(a, b)


def copy(src, dst):
    if same(src, dst):
        return
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    shutil.copyfile(src, dst)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("matrix", help="the scenario's Response Matrix .xlsx")
    ap.add_argument("--assets", required=True, help="folder holding the campus map and the portrait PNGs (searched recursively)")
    ap.add_argument("--out", required=True, help="output folder")
    ap.add_argument("--repo", action="store_true", help="also copy engine, matrix, packet and README into --out")
    ap.add_argument("--packet", help="the scenario packet .docx (default: the .docx next to the matrix with 'Packet' in its name)")
    ap.add_argument("--no-packet-check", action="store_true", help="build without checking the packet against the matrix")
    ap.add_argument("--framework", help="an export of the framework Google Doc (.docx, .txt or .pdf)")
    ap.add_argument("--no-framework-check", action="store_true", help="build without checking the game text against the framework")
    a = ap.parse_args()

    g = load_game_text(a.framework, a.no_framework_check)
    check_engine_keys(g)
    d = load(a.matrix, g.facts)
    packet = None
    if a.no_packet_check:
        print("Skipping the packet check (--no-packet-check).")
    else:
        packet = a.packet or find_packet(a.matrix)
        problems = check_packet(a.matrix, packet)
        if problems:
            sys.exit(f"The packet ({os.path.basename(packet)}) disagrees with the matrix; make them match before building:\n"
                     + "\n".join(" - " + p for p in problems))
        print(f"Packet OK: {os.path.basename(packet)} matches the matrix")
    out = os.path.abspath(a.out)
    os.makedirs(out, exist_ok=True)

    # assets
    for w in d.WITNESSES:
        copy(find_file(a.assets, w["portrait"]), os.path.join(out, "assets", "portraits", w["portrait"]))
    copy(find_file(a.assets, d.MAP), os.path.join(out, "assets", d.MAP))
    for fn in ICONS:
        copy(os.path.join(HERE, fn), os.path.join(out, "assets", fn))

    # twee
    data = json.dumps(scenario_json(d), ensure_ascii=False).replace("</", "<\\/")
    header = (
        f":: StoryTitle\n{d.TITLE}\n\n\n"
        ":: StoryData\n" + json.dumps({"ifid": d.IFID, "format": "SugarCube", "format-version": "2.37.3",
                                       "start": "Start"}, indent=2) + "\n\n\n"
    )
    text = json.dumps(g.text, ensure_ascii=False).replace("</", "<\\/")
    scen = (f":: ScenarioData [script]\n/* Generated by build_twine.py from {os.path.basename(a.matrix)} and "
            f"{os.path.basename(GAME_TEXT)}. Do not edit by hand. */\nsetup.scenario = {data};\nsetup.gameText = {text};\n")
    with open(ENGINE, encoding="utf-8") as f:
        engine = f.read()
    # scenario data must load before the engine script, so it comes first
    twee = header + scen + "\n\n" + engine
    src_dir = os.path.join(out, "source")
    os.makedirs(src_dir, exist_ok=True)
    twee_path = os.path.join(src_dir, "story.twee")
    with open(twee_path, "w", encoding="utf-8") as f:
        f.write(twee)

    # compile
    tweego = os.environ.get("TWEEGO") or shutil.which("tweego")
    if not tweego:
        sys.exit(f"Wrote {twee_path}, but Tweego was not found. Install it or set TWEEGO=/path/to/tweego, then rerun.")
    r = subprocess.run([tweego, "-f", "sugarcube-2", "--head", HEAD, "-o", os.path.join(out, "index.html"), twee_path],
                       capture_output=True, text=True)
    if r.returncode != 0:
        sys.exit("Tweego failed:\n" + r.stderr)
    if r.stderr.strip():
        print(r.stderr.strip())

    if a.repo:
        for fn in ("hearsay_engine.twee", "build_twine.py", "scenario_matrix.py", "check_packet.py",
                   "game_text.py", "check_framework.py", "head.html", *ICONS, os.path.basename(GAME_TEXT)):
            copy(os.path.join(HERE, fn), os.path.join(out, "engine", fn))
        scen_out = os.path.join(out, "scenario")
        keep = {os.path.basename(a.matrix)} | ({os.path.basename(packet)} if packet else set())
        if os.path.isdir(scen_out):  # drop anything left from an older layout (e.g. the retired .py data files)
            for fn in os.listdir(scen_out):
                if fn not in keep and os.path.isfile(os.path.join(scen_out, fn)):
                    os.remove(os.path.join(scen_out, fn))
        copy(os.path.abspath(a.matrix), os.path.join(scen_out, os.path.basename(a.matrix)))
        if packet:
            copy(os.path.abspath(packet), os.path.join(scen_out, os.path.basename(packet)))
        readme = os.path.join(HERE, "README_template.md")
        if os.path.exists(readme):
            with open(readme, encoding="utf-8") as f:
                txt = (f.read().replace("{{TITLE}}", d.TITLE).replace("{{MATRIX}}", os.path.basename(a.matrix))
                       .replace("{{PACKET}}", os.path.basename(packet) if packet else "the scenario packet (.docx)"))
            play_url = d.PLAY_URL  # the scenario's GitHub Pages URL; the line is omitted if unset
            txt = txt.replace("{{PLAY_LINE}}", f"**Play online: {play_url}**\n\n" if play_url else "")
            with open(os.path.join(out, "README.md"), "w", encoding="utf-8") as f:
                f.write(txt)
        open(os.path.join(out, ".nojekyll"), "w").close()

    print(f"Built {os.path.join(out, 'index.html')}")


if __name__ == "__main__":
    main()

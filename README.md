# HEARsay Mysteries

**Play online: https://holytispoon.github.io/HEARsay-Mysteries/**

Short investigative games about conversational receptiveness. Each game is a mystery on a university campus: you interview student witnesses, and you get the most out of them by using the HEAR framework (**H**edge your claims, **E**mphasize agreement, **A**cknowledge the other perspective, **R**eframe to the positive), which comes from Julia Minson's research on conversational receptiveness.

The games are built in [Twine 2](https://twinery.org) with the [SugarCube 2](https://www.motoslave.net/sugarcube/2/) story format. All of them run on one engine; each scenario supplies its own story.

## Scenarios

| Scenario | Play | Folder |
|---|---|---|
| The Commons' Comma | https://holytispoon.github.io/HEARsay-Mysteries/commons-comma/ | `scenarios/commons-comma/` |

## Play

- **Online:** open the link above and pick a mystery.
- **Offline:** download the repository (Code → Download ZIP), unzip it, and open `docs/index.html` (or a game's `docs/<scenario>/index.html`) in a web browser. Keep each game's `index.html` and `assets` folder together.

A game takes about 45 to 75 minutes. Progress can be saved from the sidebar (Saves). Saves are kept in the browser you're playing in.

## What's in this repository

| Path | What it is |
|---|---|
| `engine/hearsay_engine.twee` | The game engine: SugarCube passages, JavaScript and CSS. The same for every scenario; it holds no player-facing text. |
| `engine/HEARsay game text.xlsx` | A read-only copy of the Google Sheet [HEARsay game text](https://docs.google.com/spreadsheets/d/111OiAL_txh5G7kzgSV70Lh6TEsS4dtQR0AXgkRBQUzE/edit), which is the source: the game's own headings, buttons and labels, filed under the framework's passage names. The same for every scenario. Only the build updates the copy. |
| `engine/HEARsay mysteries.docx` | A read-only copy of the framework Google Doc [HEARsay mysteries](https://docs.google.com/document/d/1q2QzVFRMtBDEl7KbaLdKgp_xrxrseyejLyNDNfUky3Y/edit). The build checks the game text against it, and takes from its tabs the game's rules ("Rules of the game"), the HEAR guide ("HEAR Framework introduction"), the framework's numbers ("Game facts": 8 witnesses, 16 Clues, the question budget …) and the sentences that restate the framework ("Game text"). The same for every scenario. Only the build updates the copy. |
| `engine/build.py` | Build script: runs the checks below, adds each scenario and the game text to the engine, compiles each game with Tweego into `docs/`, and writes the landing page. |
| `engine/game_text.py`, `engine/check_framework.py` | Read the game text and the framework's numbers (from the framework's tabs and the Sheet), and check them against the framework document. |
| `engine/scenario_matrix.py` | Reads a Response Matrix and checks the framework's rules (tripwires, clue counts, receptiveness ranks) and that the workbook agrees with itself. |
| `engine/check_packet.py` | Checks that a scenario packet says the same things as its Response Matrix. |
| `engine/sources.py`, `engine/sources.json`, `scenarios/<id>/sources.json` | Where the sources of truth are in Google Drive, and how the build updates the repository's copies from downloads of them. |
| `engine/landing.html` | The landing page. The build fills in the list of scenarios. |
| `engine/head.html`, `engine/favicon.svg`, `engine/favicon-32.png`, `engine/apple-touch-icon.png` | The site icon (a magnifying glass) and the `<head>` lines that link it. The SVG draws the glass in rust, or a lighter rust when the browser is in dark mode. The PNGs are renders of it: the 32 px one in rust; the 180 px iOS icon in the lighter rust with a margin, because iOS shows transparency as black and rounds the corners. |
| `engine/shared-assets/` | Images more than one scenario may use: `blank-university-logo.png` (Blank University's logo, with the white book pages made transparent so it sits on the game's paper colour) and `the-blank-masthead.png` (the masthead of *The Blank*, Blank University's student newspaper, made by `tools/make-the-blank-masthead.py`; it heads the newspaper article in the Newspaper passage). |
| `scenarios/<id>/… – Response Matrix.xlsx` | A read-only copy of the scenario's Response Matrix (a Google Sheet in Drive), which the game is built from: the scenario's text for each passage (the Hook, the endings …), the mystery, the final-answer form, witnesses, questions, every answer, Clues, Major Clues, Red Herrings and follow-ups. |
| `scenarios/<id>/… – Scenario Packet.docx` | A read-only copy of the designer's scenario packet (a Google Doc in Drive): the same scenario written up for people, plus the timeline, the positions and the hidden payoff. Spoilers throughout. |
| `scenarios/<id>/assets/` | The portraits and campus map the game shows, and the scenario's own logo if it has one. |
| `scenarios/<id>/source-art/` | Original artwork, kept for reference; the build doesn't use it. |
| `docs/` | The published site, generated by the build: the landing page and one folder per game (`index.html`, `assets/`, and `source/story.twee`, the Twee source the game is compiled from, which you can import into Twine). GitHub Pages serves this folder. |
| `ARCHITECTURE.md` | Why the repository is laid out this way. |

## Rebuild after editing

The sources of truth are Google Docs and Sheets in the Drive folder "HEARsay mystery game", not files in this repository, and only the author edits them: the framework Doc "HEARsay mysteries", the game text Sheet "HEARsay game text", and each scenario's Response Matrix (Sheet) and Scenario Packet (Doc). Edit a scenario's matrix and make the same change in its packet. To change the game in every scenario, edit the framework Doc's tabs or the game text Sheet:

- "Rules of the game": a heading, then one paragraph per rule.
- "HEAR Framework introduction": a heading and an introduction, a heading and a paragraph for each of H, E, A and R, then a paragraph on unreceptive language.
- "Game facts": a table (Key, Fact, Value, Framework wording, Framework section) of the framework's numbers. Each fact's wording must appear elsewhere in the Doc and contain its value; a fact with no wording is an engine choice the Doc doesn't state.
- "Game text": a table (Key, Screen, Text, Placeholders) of the game's sentences that restate the framework.
- The Sheet's Text sheet: everything else the game shows (headings, buttons, labels); its Passages sheet lists the screens.

`{curly}` words are filled in by the game, and bold words are shown in bold. A key can come from only one place.

To rebuild, download them (File → Download → Microsoft Excel for the Sheets, Microsoft Word for the Docs) into one folder, keeping their names, and pass that folder with `--downloads`. The build checks every download first and stops, copying nothing, if one has a problem. Then it replaces each copy in the repository whose content changed, and builds. Commit the changed copies with the rebuilt `docs/`.

Requirements: Python 3.8+ with `openpyxl` and `python-docx` (`pip3 install openpyxl python-docx`), and [Tweego](https://www.motoslave.net/tweego/) 2.1+ with SugarCube 2.37.3. Put `tweego` on your PATH, or set `TWEEGO=/path/to/tweego`; if its bundled SugarCube is older, set `TWEEGO_PATH` to a folder holding a `sugarcube-2` folder of version 2.37.3.

From the repository root:

```
python3 engine/build.py --downloads ~/Downloads/hearsay                  # every scenario
python3 engine/build.py commons-comma --downloads ~/Downloads/hearsay    # one scenario
python3 engine/sources.py ~/Downloads/hearsay                           # lists the sources and which downloads it finds
```

Without `--downloads` the build uses the copies as they are and warns that they may be behind the Google files.

The build stops if a matrix breaks a rule or disagrees with itself, if a packet disagrees with its matrix, if a matrix's Scenario ID differs from its folder name, or if two scenarios share an IFID, and it lists each problem. The game text is also checked against the framework document. Open `docs/index.html` to test, then commit the rebuilt `docs/` with your change.

## Add a scenario

1. Make a folder `scenarios/<id>/`, where `<id>` is a short lowercase ID such as `commons-comma`.
2. In Drive, copy a Response Matrix Sheet and a packet Doc into a new scenario folder and replace their content, keeping the same sheets and columns. On the matrix's Settings sheet, set the Scenario ID to `<id>` and put a new `IFID` (any UUID, in uppercase). Put their addresses in `scenarios/<id>/sources.json`, and their downloads in `scenarios/<id>/` as the first copies.
3. Put the portraits and the campus map in `scenarios/<id>/assets/`. For a university logo, add a `Logo` row to the matrix's Settings sheet with the file name, and the same file name under "Logo file" in the packet; the game then shows it on Start, as letterhead on the president's assignment in the Hook and the Case File, and on the Ending. The file is looked for in `scenarios/<id>/assets/`, then in `engine/shared-assets/`, so scenarios set at the same university can share one.
   For a newspaper article that sets up the story, add a tab named "Newspaper article" to the packet: a heading (the headline), one paragraph (the deck), the byline ("By …"), then the story. The game shows it between the Rules and the Hook, and in the sidebar for the rest of the game; a paragraph that quotes the Hook's Clause is set as a quote. For a masthead above it, add a `Masthead` row to the matrix's Settings sheet and "Masthead file" to the packet, as for the logo.
4. Run `python3 engine/build.py <id> --downloads <folder>`. The game appears at `docs/<id>/` and on the landing page. You don't need to change the engine.

## Credits

Game design: Tomer Perry. The HEAR framework is from Yeomans, Minson, Collins, Chen and Gino, "Conversational receptiveness: Improving engagement with opposing views", *Organizational Behavior and Human Decision Processes* 160 (2020). All characters and events are fictional.

# HEARsay Mysteries — project notes for Claude

A Twine 2 / SugarCube 2.37.3 game that teaches Julia Minson's HEAR framework (Hedge, Emphasize agreement, Acknowledge, Reframe to the positive). This repository holds every scenario (so far one, *The Commons' Comma*) and the engine they share; GitHub Pages serves the games from `docs/`. Owner: Tomer Perry (GitHub: HolyTispoon).

## Sources of truth: in Google Drive, edited by the author

Every source of truth is a Google Doc in the Drive folder "HEARsay mystery game", outside this repository, and **the author edits them**. There are two kinds: the framework Doc, and one packet Doc per scenario. Their addresses are in `engine/sources.json` and `scenarios/<id>/sources.json`. What you may change in each:

- **A scenario's packet Doc**: you may edit it whenever the task calls for it, as long as every change is a suggestion (`writeControl.writeMode: "SUGGEST"`) that the author accepts or rejects. Never make a direct edit to it; if the Docs API refuses suggestion mode, stop and ask.
- **The framework Doc**: by default, don't edit it: propose the exact change (cell by cell, or the wording) and let the author make it. Edit it yourself only when the author explicitly asks you to make that change in it, for that task ("you edit the files"); a request to "update", "fix" or "rework" something without that is a request for the proposed change. Without that request, the only thing you may add to it is a suggestion or comment pointing out a typo or a grammatical error.

Whenever you edit one:
- Change only what the task needs. Anything else you notice (a typo, an inconsistency, a related fix) is a proposal in your reply, or in the packet a separate suggestion, not a direct edit.
- Every change is a suggestion (`writeControl.writeMode: "SUGGEST"`); never fall back to a direct edit. (A new tab can't be a suggestion: create one only when the author agrees, and put its content in as suggestions.)
- Afterwards, download the file and check it. A Doc's .docx download includes its pending suggestions, as tracked changes (a suggested insertion is text inside `w:ins`, a suggested deletion is `w:delText`). The build reads only the accepted text, so it can't see your suggestions until the author accepts them: check a packet's with `python3 engine/scenario_packet.py --with-suggestions "<download>"`, which reads it as if they were accepted.

1. The framework Google Doc **"HEARsay mysteries"**: https://docs.google.com/document/d/1q2QzVFRMtBDEl7KbaLdKgp_xrxrseyejLyNDNfUky3Y/edit — the framework and rules. Five of its tabs are read by the build (game_text.py says how each is laid out): "Rules of the game" (`rules.*`), "HEAR Framework introduction" (`hear.*`), "Game facts" (the framework's numbers, a table), "Game text" (the sentences that restate the framework, a table) and "Screens and labels" (the list of screens, and the rest of the engine's player-facing text: headings, buttons, labels; two tables). `engine/hearsay_engine.twee` holds no wording. Read the Doc; don't rely on older copies (`Archive/Hearsay Mystery Game.docx` and the claude.ai project's framework.md are superseded).
2. Per scenario, in the scenario's Drive folder (for example `The Commons' Comma/`): the Scenario Packet (a Google Doc), the scenario's one source. `engine/scenario_packet.py` lists what the build reads from it (labelled paragraphs such as "Situation.", "Q1 [C13, clue; opinion receptive].", "Map file:", and tables found by their column headings), including the optional "Newspaper article" section (a Heading 1, such as "2. Newspaper article", in the packet's main tab), shown as the Newspaper passage. The rest of the packet (timeline, aftermath, other tabs) is prose for people.

The old Google Sheets ("HEARsay game text" and each scenario's Response Matrix) are no longer sources: the build generates both workbooks from the Docs.

The repository holds read-only copies of the Docs (`engine/HEARsay mysteries.docx`, `scenarios/<id>/… – Scenario Packet.docx`) and the workbooks generated from them (`engine/HEARsay game text.xlsx`, `scenarios/<id>/… – Response Matrix.xlsx`). Never edit a copy or a workbook: the build replaces each copy from a download of its Google Doc and regenerates the workbooks (see Build). If the documents and the live game disagree, ask the author which is right.

Use the doc's names for the game's parts: Introduction, the Hook, the CCC, Clues, Red Herrings, Major Clues, tripwire, follow-up, budget, final-answer form; passages Start, Rules, the Hook, Map, Interview, Answer, Follow-up, Evidence Board, Submission, Result, ending, ranking. Don't invent other names.

Structural changes to the framework are the user's decisions and belong in the Google Doc first. Ask before making them in the game.

## Where things live

- `engine/`: the same for every scenario. The engine (`hearsay_engine.twee`), the copy of the framework Doc and the game text workbook generated from it, `sources.json` and `sources.py` (the sources of truth and the copies), the build and check scripts, `landing.html` (the site's landing page; its wording is edited there), `head.html` and the site icons, and `shared-assets/` (images more than one scenario may use, such as Blank University's logo and The Blank's masthead; a scenario names them in its packet's optional "Logo file:" and "Masthead file:").
- `scenarios/<Scenario ID>/`: one folder per scenario, named by the packet's Scenario ID. The copy of the packet, the matrix generated from it, `sources.json` (the packet's Google address), and `assets/` (the portraits and map the game uses; the only place for the scenario's images, since the build reads no others. Originals the game doesn't use are in Drive `Archive/`).
- `docs/`: the published site, generated by the build. `index.html` is the landing page; `<id>/index.html` is each game, with its `assets/`, generated `source/story.twee` and, if the packet has a newspaper article, `newspaper-article.pdf` (printed by headless Chrome; kept as it is while the article is unchanged, since a cloud session has no Chrome). Never edit anything in `docs/` by hand; rebuild.
- `ARCHITECTURE.md`: why the repository is laid out this way, and the migration steps still open.
- Google Drive folder "HEARsay mystery game" (owner tperry@g.harvard.edu): the framework Google Doc, one folder per scenario with its packet Doc (`The Commons' Comma/`), the Drive `CLAUDE.md` (project notes, open items and the Tweego/SugarCube paths on the user's Mac) and `Archive/`. The repository is cloned into that folder as `HEARsay-Mysteries/`, synced by Google Drive for desktop. The user rebuilds and publishes with `bash tools/publish.sh` (browser downloads of the Google Docs, build in its own clone under `~/.cache/hearsay-publish`, never asks anything: it stops any earlier run still going, publishes a PR merged into `main` when the rebuild changed something, then resets the Drive clone to `main` with `update-drive.sh --yes`, listing what changed). Otherwise changes merged into `main` reach it only when the user runs `publish.sh`, or `bash tools/update-drive.sh` (which lists every file it would change and asks first); there is no scheduled pull, by the user's choice. Nothing is copied between the repository and Drive by hand. Don't leave that clone on another branch: do branch work in a separate clone. Drive still holds the old `Twine engine/` folder until it is archived (see `ARCHITECTURE.md`); don't update it.

## Build

Needs Python 3.8+ with `openpyxl`, `python-docx` (and `pypdf` with Google Chrome for a newspaper article's PDF) and Tweego 2.1+ pointed at SugarCube 2.37.3. From the repo root:

```
python3 engine/build.py --downloads "<folder of downloads>"                  # every scenario
python3 engine/build.py commons-comma --downloads "<folder of downloads>"    # one scenario
python3 engine/sources.py "<folder of downloads>"                           # lists the sources, copies and downloads found
```

The downloads folder holds the Google Docs downloaded as .docx, named as they are in Drive. The build first reads every download into the workbooks and checks them (game text, the packet's labels and tables, matrix rules, the packet's prose); if one has a problem it stops and copies nothing. Then it replaces each copy whose content differs, regenerates each workbook whose content would change, and says which ones changed; copies with no download are used as they are, with a warning. Always build with fresh downloads of everything, and commit the changed copies and workbooks with the rebuilt `docs/`.

Checks run in order: game text vs. the framework Doc, engine text keys; then per scenario the packet's structure, matrix vs. the Doc's Game facts and itself, Scenario ID vs. folder name, the packet's prose, then Twee, Tweego, assets; then no two scenarios share an IFID, and the landing page. It fails loudly: fix the source, don't loosen a check. An engine or game-text change needs every scenario rebuilt; commit the rebuilt `docs/` with it.

In a cloud session, `.claude/hooks/session-start.sh` installs Tweego 2.1.1, SugarCube 2.37.3 and the Python packages and sets `TWEEGO` and `TWEEGO_PATH`. To download the Docs, use `download_file_content` with `exportMimeType` `application/vnd.openxmlformats-officedocument.wordprocessingml.document`, decode the base64 `content` into a scratch folder under the file's Drive title plus .docx, and pass the folder with `--downloads`.

## Pull requests

Push every commit as soon as it is made, and keep every pushed branch in an open pull request against `main`. Open it with the first push, following `.github/pull_request_template.md` (summary, questions for the author, follow-up needed to activate the work); when later commits change the picture, update its description. Keep the template consistent with the one in HolyTispoon/Ethics-Sandbox-Virtual-Tabletop: the shared sections are word for word the same, and that repo adds a build-plan pointer and a Checks list.

## Gotchas

- Don't run the docx skill's `merge_runs.py` on the packet (it once turned "=" into "."). Edit the packet's XML directly.
- The packet repeats some facts, and the build checks each copy against the answers: the overview grid in "The Response Matrix", the Clues table's 'Given by', each red herring's "(given by …)", and the witness table's 'Receptive on'. Keep them in sync by hand.
- The final-answer form's two column headings are the questions the player answers, and "(correct)" marks the right person and place.
- Game text cells are trimmed. Rules display in the order of the framework's "Rules of the game" tab.
- Regression test: the build is deterministic, so for a change that should not alter a game, build `main` and the branch and compare `docs/<id>/index.html` byte for byte. For a change that should, walk the same screens of both builds in iframes with `Math.random` stubbed and diff `innerText`.

## Engine choices not yet in the Google Doc

History/undo disabled so tripwires stick; ranking 1 per exact place, ½ per place off; follow-up phrasings in random order; final-answer options in random order (kept for the playthrough); the debrief reveals the solution, Major Clues and each red-herring set's resolution; Follow-up Result, Case File, Newspaper and Sidebar are engine additions (Newspaper: the packet's newspaper article, shown only from the sidebar); the sidebar's Campus map, Evidence board, Case file and Newspaper stay locked until the Hook's button leads to the Map. Open with the user: the doc puts the big reveal before the ranking, but players who answer wrongly get the full story only in the debrief after ranking; the HEAR guide paraphrases the doc's table.

## Working with the user

Direct, factual, no flattery; criticism welcome.

Name each branch for the work it contains (for example `claude/add-site-icon`), not with a generated name, and start it from `main`; this is explicit permission to push to that branch instead of an assigned one.

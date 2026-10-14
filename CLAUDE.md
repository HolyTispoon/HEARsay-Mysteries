# HEARsay Mysteries — project notes for Claude

A Twine 2 / SugarCube 2.37.3 game that teaches Julia Minson's HEAR framework (Hedge, Emphasize agreement, Acknowledge, Reframe to the positive). This repository holds scenario 1, *The Commons' Comma*, served by GitHub Pages. Owner: Tomer Perry (GitHub: HolyTispoon).

## Author-only documents

Some documents are **Author only**. Never edit them directly. The only thing you may do is point out typos or grammatical errors, and only as suggestions (a comment or a suggested edit for the author to accept), never as a change to the text.

Author-only documents:
- The framework Google Doc **"HEARsay mysteries"**: https://docs.google.com/document/d/1q2QzVFRMtBDEl7KbaLdKgp_xrxrseyejLyNDNfUky3Y/edit

If a change seems needed in an Author-only document (for example to record an engine choice, or to paste the section 5 draft), propose the wording to the user and let them make it.

## Sources of truth (in order)

1. The Google Doc "HEARsay mysteries" (above): the framework and rules. Read it; don't rely on older copies (`Archive/Hearsay Mystery Game.docx` and the claude.ai project's framework.md are superseded; the old PDF export has been removed).
2. Per scenario: the Response Matrix (`scenario/… – Response Matrix.xlsx`), which the game is built from, and the Scenario Packet (`scenario/… – Scenario Packet.docx`), which must say the same things. Both are edited by hand.
3. `engine/HEARsay game text.xlsx`: all of the engine's player-facing text and the framework's numbers (Facts sheet). `engine/hearsay_engine.twee` holds no wording.

Use the doc's names for the game's parts: Introduction, the Hook, the CCC, Clues, Red Herrings, Major Clues, tripwire, follow-up, budget, final-answer form; passages Start, Rules, the Hook, Map, Interview, Answer, Follow-up, Evidence Board, Submission, Result, ending, ranking. Don't invent other names.

Structural changes to the framework are the user's decisions and belong in the Google Doc first. Ask before making them in the game.

## Where things live

- This repo: the built game (`index.html`, `assets/`, generated `source/story.twee`) plus copies of `engine/` and `scenario/`. Never edit `index.html` or `source/story.twee` by hand; rebuild.
- Google Drive folder "HEARsay mystery game" (owner tperry@g.harvard.edu): the working copy. its `CLAUDE.md` (the full project notes, including the old handoff, open items and the Tweego/SugarCube paths on the user's Mac), `Twine engine/`, `The Commons' Comma/` (matrix, packet, `Assets/`, `Game/` = what this repo contains), `Archive/`. Keep the repo and Drive in step.

## Build

Needs Python 3.8+ with `openpyxl`, `python-docx` (and `pypdf` for a PDF export), and Tweego 2.1+ pointed at SugarCube 2.37.3. From the repo root:

```
python3 engine/build_twine.py "scenario/The Commons' Comma – Response Matrix.xlsx" --assets assets --out . --framework "<export of the Google Doc>.docx"
```

Checks run in order: game text vs. framework export, engine text keys, matrix vs. Facts and itself, packet vs. matrix; then Twee, Tweego, assets. It fails loudly: fix the source, don't loosen a check. Without `--framework` the framework check is skipped with a warning.

## Gotchas

- Don't run the docx skill's `merge_runs.py` on the packet (it once turned "=" into "."). Edit the packet's XML directly.
- Matrix columns are found by header name. 'Given by', 'Tripwire question', Grid and repeated set titles are copies the build checks; keep them in sync by hand.
- Game text cells are trimmed. `rules.N` rows display in number order.
- Regression test: build old and new `index.html`, walk the same screens in iframes with `Math.random` stubbed, diff `innerText`.

## Engine choices not yet in the Google Doc

History/undo disabled so tripwires stick; follow-up bonus 1 point; ranking 1 per exact place, ½ per place off; follow-up phrasings in random order; the debrief reveals the solution, Major Clues and each red-herring set's resolution; Follow-up Result, Case File and Sidebar are engine additions. Open with the user: the doc puts the big reveal before the ranking, but players who answer wrongly get the full story only in the debrief after ranking; the HEAR guide paraphrases the doc's table.

## Working with the user

Direct, factual, no flattery; criticism welcome.

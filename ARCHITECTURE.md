# One repository for all scenarios

Status: the repository layout and the build are in place (2026-10-15). The steps still open are listed under Migration.

## Why

Until October 2026 each scenario was meant to get its own repository, and the files lived in up to four places: the repository, the Drive folder `Twine engine/`, the matrix and packet in `The Commons' Comma/`, and `The Commons' Comma/Game/` (a build output copied by hand into the repository). Every change had to be moved between them by hand, and a second scenario would have added a second copy of the engine.

The goal: one copy of every file, in one git repository, which is also the Drive working folder. The engine exists once and every scenario is built from it.

## Repository layout

```
engine/                         the same for every scenario (unchanged contents)
  hearsay_engine.twee
  HEARsay game text.xlsx
  build.py                      builds one scenario, or all of them, plus the landing page
  landing.html                  the landing page, with a placeholder for the list of scenarios
  scenario_matrix.py, check_packet.py, game_text.py, check_framework.py
  head.html, favicon.svg, favicon-32.png, apple-touch-icon.png
scenarios/
  commons-comma/                folder name = the matrix's Scenario ID
    The Commons' Comma – Response Matrix.xlsx
    The Commons' Comma – Scenario Packet.docx
    assets/                     game-ready images: portraits/, campus-map.jpg
    source-art/                 originals (campus-map.png, Witness portraits/); not used by the build
  <next-scenario>/
    …
docs/                           the published site (generated; GitHub Pages serves this folder)
  index.html                    landing page: one entry per scenario
  commons-comma/
    index.html                  the game
    assets/                     portraits, map and icons, copied by the build
    source/story.twee           the generated Twee source
  <next-scenario>/…
README.md, CLAUDE.md, ARCHITECTURE.md
.claude/hooks/session-start.sh   installs the build tools in cloud sessions
tools/setup-drive-clone.sh      sets up the Drive clone and the scheduled pull on the Mac
```

- The engine, the game text and the build scripts exist once. A fix to the engine reaches every scenario on the next build.
- A scenario is a folder holding its matrix, its packet and its images. Adding one means adding a folder; nothing in `engine/` changes.
- `docs/` is generated. Nobody edits it by hand. It stays committed (as `index.html` is today) so GitHub Pages can serve it without a build step on GitHub, and so every change to the game shows up in a pull request's diff.

## Build

```
python3 engine/build.py                     # every scenario, then the landing page
python3 engine/build.py commons-comma       # one scenario, then the landing page
```

The checks stay exactly as they are and run per scenario: game text against the framework export, engine text keys, matrix against the Facts sheet and itself, packet against matrix. A failure in one scenario stops the build.

What changed in `build_twine.py` (renamed `build.py`):

- Takes Scenario IDs instead of a matrix path, and finds the matrix, packet and `assets/` in `scenarios/<id>/`. `--assets`, `--out` and `--packet` are gone.
- Writes to `docs/<id>/`. The game's relative paths (`assets/…`, the icons in `head.html`) keep working because each game still has its own `assets/` next to its `index.html`.
- Writes `docs/index.html` from `engine/landing.html`: the project's introduction, then one entry per built scenario with its Title and summary (the Start passage's Summary).
- Drops `--repo` and `README_template.md`: there is no per-scenario repository any more.
- Looks for the framework export in the repository's parent folder too, which in Drive is the folder holding the Google Doc. Saving the export next to the Google Doc is then enough; `--framework` still works.
- No longer uses the matrix's Play URL setting (it filled in the per-scenario README).
- Checks that every scenario has its own IFID and Scenario ID, and that the folder name equals the Scenario ID.

The build is deterministic: rebuilding the game in a cloud session, with Tweego 2.1.1 and SugarCube 2.37.3, gives a byte-identical `index.html` (checked 2026-10-15). The restructure was verified that way: `docs/commons-comma/index.html` is byte-identical to the last build of the old layout, except that two code comments name `build.py` instead of `build_twine.py`.

## Published site

- Site address: `https://holytispoon.github.io/HEARsay-Mysteries/`, with each game at `…/<scenario-id>/`.
- GitHub Pages serves the `docs/` folder of `main` (Settings → Pages → Branch: main, folder: /docs).
- Saves are stored in the player's browser per site domain (`holytispoon.github.io`), which does not change. The game's title and IFID don't change either.
- The old link, `https://holytispoon.github.io/HEARsay-Mysteries-Commons-Comma/`, stopped working when the repository was renamed (GitHub redirects repository addresses after a rename, not Pages addresses). It is not redirected (decided 2026-10-18).

## Where builds run

- **Cloud sessions (Claude):** a SessionStart hook (`.claude/hooks/session-start.sh`) installs Tweego 2.1.1, SugarCube 2.37.3, `openpyxl` and `python-docx`, so every session can build and run the regression check. The framework check uses a text export of the Google Doc fetched through the Drive connector.
- **The Mac:** as today (Tweego path and SugarCube folder in the Drive `CLAUDE.md`), run from the clone.
- **GitHub (later, optional):** an Action that runs the checks on every pull request and fails if `docs/` is out of date. It can't run the framework check, because the Google Doc is not in the repository.

## Drive

```
HEARsay mystery game/
  HEARsay mysteries               the framework Google Doc (Author only)
  HEARsay mysteries.docx          its export, read by the build's framework check
  CLAUDE.md                       Drive notes; points to the repository for everything in it
  Section 5 draft – ….md
  HEARsay-Mysteries/              the git clone: engine/, scenarios/, docs/
  Archive/                        gains Twine engine/, the old The Commons' Comma/ folder, Claude outputs/
```

- The clone is the only working copy. You edit the matrix and packet in `scenarios/<id>/`, in Drive or Finder, as you do now in `The Commons' Comma/`.
- On the Mac, Google Drive for desktop syncs the clone, with the folder set to "Available offline" so git works on local files.
- A scheduled job on the Mac runs `git pull --ff-only` every 10 minutes while the clone is on `main`. Changes merged on GitHub (including Claude's) then appear in Drive without anyone moving files. It stops by itself (does nothing) if the clone has diverged, so it can't overwrite your edits.
- Your edits reach GitHub when you commit and push them (GitHub Desktop: one commit, one push). Nothing pushes automatically.
- Source art lives in the repository, in `scenarios/<id>/source-art/`.

## Migration, in order

1. **Restructure.** Done in this pull request: files moved with `git mv`, the build changed as above, the landing page, `README.md` and `CLAUDE.md` rewritten. Verified byte for byte (see Build).
2. **Publish.** Done: the repository is renamed `HEARsay-Mysteries`. Still to do, by the author, right after merging: Settings → Pages → Build and deployment → Branch `main`, folder `/docs` → Save. Until then the site shows the repository root, which no longer holds a game.
3. **Drive.** On the Mac, run `tools/setup-drive-clone.sh` (it clones into Drive and installs the scheduled pull). Then move `The Commons' Comma/Assets/` into the clone as `scenarios/commons-comma/source-art/` and commit it, and move `Twine engine/`, the rest of `The Commons' Comma/` and `Claude outputs/` to `Archive/`; Claude can do the Drive moves. As of 2026-10-15 every engine, matrix and packet file in those folders matches the repository in size. Update the Drive `CLAUDE.md`.
4. **Cloud builds.** Done: `.claude/hooks/session-start.sh`.
5. **Optional:** a check on GitHub that runs the build's checks on every pull request and fails if `docs/` is out of date.

## Decisions (2026-10-15)

1. Repository renamed to `HEARsay-Mysteries`; games at `https://holytispoon.github.io/HEARsay-Mysteries/<scenario-id>/`. The old Pages link is not redirected (2026-10-18). The matrix's Play URL setting holds the new address, for reference.
2. The landing page carries the project's introduction as well as the list of scenarios.
3. The site folder is `docs/`, served directly by GitHub Pages.
4. Source art is kept in the repository, in `scenarios/<id>/source-art/`.

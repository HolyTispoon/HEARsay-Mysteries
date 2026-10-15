# One repository for all scenarios

Status: proposal. Nothing below is built yet; the questions at the end need answers first.

## Why

Today each scenario is meant to get its own repository, and the files live in up to four places: this repository, the Drive folder `Twine engine/`, the matrix and packet in `The Commons' Comma/`, and `The Commons' Comma/Game/` (a build output that is copied by hand into this repository). Every change has to be moved between them by hand, and a second scenario would add a second copy of the engine.

The goal: one copy of every file, in one git repository, which is also the Drive working folder. The engine exists once and every scenario is built from it.

## Repository layout

```
engine/                         the same for every scenario (unchanged contents)
  hearsay_engine.twee
  HEARsay game text.xlsx
  build.py                      builds one scenario, or all of them, plus the landing page
  scenario_matrix.py, check_packet.py, game_text.py, check_framework.py
  head.html, favicon.svg, favicon-32.png, apple-touch-icon.png
scenarios/
  commons-comma/                folder name = the matrix's Scenario ID
    The Commons' Comma – Response Matrix.xlsx
    The Commons' Comma – Scenario Packet.docx
    assets/                     game-ready images: portraits/, campus-map.jpg
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

Changes to `build_twine.py` (renamed `build.py`):

- Takes Scenario IDs instead of a matrix path, and finds the matrix, packet and `assets/` in `scenarios/<id>/`.
- Writes to `docs/<id>/`. The game's relative paths (`assets/…`, the icons in `head.html`) keep working because each game still has its own `assets/` next to its `index.html`.
- Writes `docs/index.html`, the landing page, from each matrix's Title and summary (Start passage).
- Drops `--repo` and `README_template.md`: there is no per-scenario repository any more.
- Looks for the framework export in the repository's parent folder too, which in Drive is the folder holding the Google Doc. Saving the export next to the Google Doc is then enough; `--framework` still works.
- Derives each game's link from one site address plus the Scenario ID, instead of the matrix's Play URL setting.
- Checks that every scenario has its own IFID and Scenario ID, and that the folder name equals the Scenario ID.

The build is deterministic: rebuilding the current game in a cloud session, with Tweego 2.1.1 and SugarCube 2.37.3, gives a byte-identical `index.html` (checked 2026-10-15). So the restructure can be verified exactly: the moved game must come out byte-identical apart from where it is written.

## Published site

- Site address: `https://holytispoon.github.io/<repository name>/`, with each game at `…/<scenario-id>/`.
- GitHub Pages serves the `docs/` folder of `main` (Settings → Pages → Branch: main, folder: /docs).
- Saves are stored in the player's browser per site domain (`holytispoon.github.io`), which does not change. The game's title and IFID don't change either.
- The current link, `https://holytispoon.github.io/HEARsay-Mysteries-Commons-Comma/`, must keep working: people have it. See question 1.

## Where builds run

- **Cloud sessions (Claude):** a SessionStart hook installs Tweego 2.1.1, SugarCube 2.37.3, `openpyxl` and `python-docx`, so every session can build and run the regression check. The framework check uses a text export of the Google Doc fetched through the Drive connector.
- **The Mac:** as today (Tweego path and SugarCube folder in the Drive `CLAUDE.md`), run from the clone.
- **GitHub (later, optional):** an Action that runs the checks on every pull request and fails if `docs/` is out of date. It can't run the framework check, because the Google Doc is not in the repository.

## Drive

```
HEARsay mystery game/
  HEARsay mysteries               the framework Google Doc (Author only)
  HEARsay mysteries.docx          its export, read by the build's framework check
  CLAUDE.md                       Drive notes; points to the repository for everything in it
  Section 5 draft – ….md
  <repository name>/              the git clone: engine/, scenarios/, docs/
  Source art/
    The Commons' Comma/           originals: campus-map.png (3.5 MB), Witness portraits/
  Archive/                        gains Twine engine/, the old The Commons' Comma/ folder, Claude outputs/
```

- The clone is the only working copy. You edit the matrix and packet in `scenarios/<id>/`, in Drive or Finder, as you do now in `The Commons' Comma/`.
- On the Mac, Google Drive for desktop syncs the clone, with the folder set to "Available offline" so git works on local files.
- A scheduled job on the Mac runs `git pull --ff-only` every 10 minutes while the clone is on `main`. Changes merged on GitHub (including Claude's) then appear in Drive without anyone moving files. It stops by itself (does nothing) if the clone has diverged, so it can't overwrite your edits.
- Your edits reach GitHub when you commit and push them (GitHub Desktop: one commit, one push). Nothing pushes automatically.
- Source art stays in Drive only. The repository holds the game-ready images.

## Migration, in order

1. **Restructure (one pull request).** Move the files with `git mv` (history is kept), change the build as above, write the landing page, rewrite `README.md` for the whole collection, update `CLAUDE.md`. Verify: the moved game is byte-identical to the current one.
2. **Publish.** Rename the repository and switch Pages to `docs/` (you do this in the repository's settings; see question 1), and put the redirect for the old link in place. Check that the old link, the new link and the landing page all open the game.
3. **Drive.** Clone into Drive, install the scheduled pull, move `Twine engine/`, `The Commons' Comma/` and `Claude outputs/` to `Archive/`, move the source art, update the Drive `CLAUDE.md`. Before archiving, check that the Drive copies have no changes the repository lacks (as of 2026-10-15 their sizes all match).
4. **Cloud builds.** Add the SessionStart hook.
5. **Optional:** the pull-request check on GitHub.

Steps 1, 4 and 5 are pull requests Claude can do. Steps 2 and 3 need you for the repository settings and the Mac; Claude can do the Drive moves.

## Questions for the author

1. **Repository name and the old link.** Recommended: rename the repository to `HEARsay-Mysteries`, so games live at `https://holytispoon.github.io/HEARsay-Mysteries/commons-comma/`. GitHub redirects the old repository address after a rename, but not the old Pages address, so a small repository with the old name would hold one page that forwards to the new address. (Creating it also ends GitHub's redirect for the old repository address, which matters only to someone who cloned the old address.) The alternative is keeping the current name, which makes the address `…/HEARsay-Mysteries-Commons-Comma/commons-comma/`.
2. **Landing page.** The site root lists the scenarios. Should it be a plain list (title, one-line summary, link), or should it also carry the project's introduction?
3. **Name of the site folder.** `docs/` is the only folder name GitHub Pages can serve directly besides the root. If that name is confusing, the alternative is `site/` plus a short GitHub Action that publishes it; that adds a step that can fail.
4. **Source art.** Keep the originals in Drive only (recommended; the repository holds what the game uses), or add them to `scenarios/<id>/source-art/`?

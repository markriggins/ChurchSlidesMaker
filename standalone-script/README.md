# Eastern Gate slide-insert script

Standalone Google Apps Script that inserts song-file slides and builds scripture slides for an Eastern Gate service deck. It is **not** part of ChurchSlidesMaker and must not be pasted into that add-on. It never deletes or reorders existing slides, never modifies `<<...>>` marker slides (it only reads them to find a section or to copy the scripture layout), and never runs Reset.

## Deploy (one-time)

1. At [script.google.com](https://script.google.com), signed in as Mark, create a new standalone project (not bound to a deck, not the add-on). Name it something like `Eastern Gate Slide Insert`.
2. Replace the default `Code.gs` and add `appsscript.json` (View > Show manifest file if it is hidden) with the files in this folder.
3. Select `insertSongs` or `insertScriptures` and click Run. Approve the one-time Google authorization for Slides and external-request access when prompted. After that, the bot can run the script with no further prompts. If the project was already authorized for songs only, run it again so the new `script.external_request` scope can be approved.
4. Optional, for Apps Script API `scripts.run`: Deploy > New deployment > type **API executable**, access **Only myself**. Note the script ID (Project Settings) and the deployment ID.

There is no web app (`doGet` is not used).

For ESV only: in Project Settings, add a Script Property named `ESV_API_KEY`. Never put the key in this config, in `Code.gs`, or in logs.

## Songs

Call `insertSongs(config)` or `runInsertSongs(config)` (the latter also accepts a JSON string, which is convenient for `scripts.run`).

To avoid editing code: call `setSongConfig(config)` once, then `runFromProperties()`. That reads Script Properties key `SONG_CONFIG`.

```json
{
  "targetId": "<service deck id>",
  "insertAt": 0,
  "applyNotesFormat": false,
  "songs": [
    { "sourceId": "<song deck id>", "repeat": 3 }
  ]
}
```

- `insertAt` is a 0-based index. Use `null` (or omit it) only when the target still has a `<<songs section>>` marker slide: songs go right after the last existing song that follows that marker, or right after the marker if there are none. If there is no marker, `insertAt` is required.
- `repeat` defaults to `1`. Each song’s every slide is copied, in original order, that many times back-to-back.
- `applyNotesFormat` defaults to `false`. If `true`, the font and size from the `<<songs section>>` marker’s speaker notes (the lyric/chord format note) are applied to inserted slides’ notes. Marker text itself is never copied onto inserted notes. Source-file notes are copied as-is, including any `<<` the song file already had; those slides are listed in `notesContainingChevrons`.

### Example: 2026-11-15 deck

Today’s deck already has no marker, so `insertAt` must be set. To put 3 copies of Every Praise at the top, then You are Worthy of My Praise:

```json
{
  "targetId": "15epVCo1KnTY9HEFpYWaKCLZ1ahSVadX_aaxA-K0UuIc",
  "insertAt": 0,
  "applyNotesFormat": false,
  "songs": [
    { "sourceId": "1ItdDEy5gRoS_bV-7lu84Z4C3ehFh6r2kQLLosyFoMAQ", "repeat": 3 },
    { "sourceId": "1v-UgXJRMXuZ5x_mX5yi_r6DkXisnyCn228yECGkCOs4", "repeat": 1 }
  ]
}
```

On a fresh template copy that still has the `<<songs section>>` marker, use `"insertAt": null` instead and songs will land after that marker (or after songs already following it).

## Scriptures

Call `insertScriptures(config)` or `runInsertScriptures(config)`. Or `setScriptureConfig(config)` then `runScripturesFromProperties()` (`SCRIPTURE_CONFIG`).

```json
{
  "targetId": "15epVCo1KnTY9HEFpYWaKCLZ1ahSVadX_aaxA-K0UuIc",
  "insertAt": null,
  "translation": "web",
  "verseNumbers": true,
  "heHim": false,
  "passages": [
    { "ref": "John 1:1-2", "notes": "Optional speaker note" }
  ]
}
```

- Defaults match Mark’s form settings: `translation` is `web`, `verseNumbers` is `true`, `heHim` is `false`. `kjv` also uses bible-api.com. `esv` uses api.esv.org and only runs if `ESV_API_KEY` is already set; the key is never logged or printed.
- `insertAt` is a 0-based index. If `null` or omitted, scriptures are added at the end of the deck.
- Each new slide is a copy of the target’s `<<scriptures>>` marker slide (photo, font, and layout). If the target has no marker, the script copies ZZ_Template slide `SLIDES_API1704141085_0` (`1mdlmL7xdlQBAvHKc1H7tyIBYSuW1gcv6BKBh5Q6TysQ`). The marker itself is left unchanged.
- Text wraps to about 30 characters per line and 5–6 lines per slide, like the ZZ_Template John 3:16 example. The template font size is kept; long passages get more slides instead of smaller type.
- A passage that fits on one slide ends with `-- Ref` as the last line. A longer passage starts with a reference-only slide (just the ref), then text slides with no citation.
- When a passage has more than one verse and `verseNumbers` is true, each verse is prefixed with `N: ` (or `C:N: ` if the passage spans chapters).
- Inserted speaker notes start blank, so `<<scriptures>>` is never copied onto them. If `notes` is given on a passage, it is written on that passage’s first slide only, unless it contains a `<<...>>` marker (in that case it is skipped and reported).

## Return value

Both insert functions log and return a manifest of each new slide’s index, objectId, and source (song title and copy number, or scripture ref and role). Song manifests also list any inserted notes that contain `<<`.

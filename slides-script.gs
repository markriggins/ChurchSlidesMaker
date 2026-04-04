// ── Lyrics Wrapper — Google Slides Script ─────────────────────────────────────
//
// SETUP:
//   1. Open your Google Slides presentation
//   2. Extensions > Apps Script
//   3. Paste this entire file, replacing any existing code
//   4. Save (Ctrl+S), close the editor, reload the presentation
//   5. A "Lyrics" menu will appear in the menu bar
//
// USAGE:
//   1. Navigate to your template slide (the one with background + word art)
//   2. Lyrics > Create Verse Slides…
//   3. Paste the wrapped verse text from Lyrics Wrapper
//   4. Click Create Slides
//
// NOTE: The script finds the largest text box on your template slide and uses
//       that as the content area. Make sure your template has exactly one plain
//       text box (not word art) where you want the verse text to appear.

const LINES_PER_SLIDE = 5;

// ─── Menu ──────────────────────────────────────────────────────────────────────

function onOpen() {
  SlidesApp.getUi()
    .createMenu('Lyrics')
    .addItem('Create Verse Slides…', 'showDialog')
    .addToUi();
}

// ─── Dialog ────────────────────────────────────────────────────────────────────

function showDialog() {
  // Capture the template slide index now, while the selection is still valid
  const idx = getCurrentSlideIndex_();
  PropertiesService.getUserProperties().setProperty('templateSlideIdx', idx.toString());

  const html = HtmlService.createHtmlOutput(`
<!DOCTYPE html>
<html>
<head>
<style>
  body { font-family: sans-serif; margin: 14px; font-size: 13px; color: #222; }
  p    { margin: 0 0 8px; }
  textarea {
    width: 100%; height: 260px;
    font-family: 'Courier New', monospace; font-size: 11px;
    box-sizing: border-box; border: 1px solid #ccc; border-radius: 4px;
    padding: 8px; resize: vertical;
  }
  button {
    margin-top: 10px; padding: 7px 20px;
    background: #4a90e2; color: white;
    border: none; border-radius: 4px; cursor: pointer; font-size: 13px;
  }
  button:disabled { opacity: .5; cursor: default; }
  #status { margin-top: 8px; font-size: 12px; color: #666; min-height: 16px; }
  #status.err { color: #c00; }
</style>
</head>
<body>
<p>Paste wrapped verse text from <strong>Lyrics Wrapper</strong>:</p>
<textarea id="t" placeholder="11: For I long to see you,&#10;that I may impart to you&#10;some spiritual gift…&#10;-- Romans 1:11"></textarea>
<br>
<button id="btn" onclick="run()">Create Slides</button>
<div id="status"></div>
<script>
  function run() {
    var text = document.getElementById('t').value.trim();
    if (!text) return;
    var btn    = document.getElementById('btn');
    var status = document.getElementById('status');
    btn.disabled   = true;
    status.className = '';
    status.textContent = 'Creating slides…';
    google.script.run
      .withSuccessHandler(function(n) {
        status.textContent = 'Done — created ' + n + ' slide(s).';
        btn.disabled = false;
      })
      .withFailureHandler(function(e) {
        status.textContent = e.message;
        status.className   = 'err';
        btn.disabled = false;
      })
      .createVerseSlides(text);
  }
</script>
</body>
</html>`)
    .setWidth(440)
    .setHeight(420);

  SlidesApp.getUi().showModalDialog(html, 'Create Verse Slides');
}

// ─── Main slide-creation function (called from dialog) ────────────────────────

function createVerseSlides(text) {
  const pres  = SlidesApp.getActivePresentation();
  const idx   = parseInt(PropertiesService.getUserProperties().getProperty('templateSlideIdx') || '0');
  const templateSlide = pres.getSlides()[idx];
  if (!templateSlide) throw new Error('Template slide not found. Please re-open the dialog from the correct slide.');

  // Separate verse lines from the citation line (starts with "--")
  const allLines    = text.split('\n');
  const verseLines  = allLines.filter(l => l.trim() && !l.trim().startsWith('--'));
  const citation    = (allLines.find(l => l.trim().startsWith('--')) || '').trim();

  if (verseLines.length === 0) throw new Error('No verse text found — make sure you pasted the wrapped output.');

  // Build per-slide batches of verse lines
  const batches = [];
  for (let i = 0; i < verseLines.length; i += LINES_PER_SLIDE) {
    batches.push(verseLines.slice(i, i + LINES_PER_SLIDE));
  }

  const multiSlide = batches.length > 1;

  if (multiSlide) {
    // Multiple slides: citation goes on its own reference slide at the START,
    // not repeated at the end.
  } else {
    // Single slide: append citation at the bottom if there's room
    if (citation) {
      const last = batches[batches.length - 1];
      if (last.length < LINES_PER_SLIDE) {
        last.push(citation);
      } else {
        batches.push([citation]);
      }
    }
  }

  // Locate the content text box on the template
  const templateBox = findContentBox_(templateSlide);
  if (!templateBox) throw new Error('No text box found on the template slide. Add a text box where you want the verse text to appear.');

  let insertAt = idx + 1;
  let slidesCreated = 0;

  // For multi-slide passages: insert a reference-only slide first
  if (multiSlide && citation) {
    const refSlide = pres.duplicateSlide(templateSlide);
    pres.moveSlide(refSlide, insertAt);
    const box = findContentBox_(refSlide);
    if (box) {
      const refText = citation.replace(/^--\s*/, '');
      box.getText().setText(refText);
    }
    insertAt++;
    slidesCreated++;
  }

  // Create one content slide per batch
  for (let i = 0; i < batches.length; i++) {
    const newSlide = pres.duplicateSlide(templateSlide);
    pres.moveSlide(newSlide, insertAt + i);
    const box = findContentBox_(newSlide);
    if (box) {
      box.getText().setText(batches[i].join('\n'));
    }
    slidesCreated++;
  }

  return slidesCreated;
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

// Returns the index of the currently selected slide
function getCurrentSlideIndex_() {
  try {
    const pres  = SlidesApp.getActivePresentation();
    const pages = pres.getSlides();
    const curId = pres.getSelection().getCurrentPage().getObjectId();
    for (let i = 0; i < pages.length; i++) {
      if (pages[i].getObjectId() === curId) return i;
    }
  } catch (e) { /* fall through */ }
  return 0;
}

// Returns the largest TEXT_BOX shape on a slide (the content area)
function findContentBox_(slide) {
  const boxes = slide.getShapes().filter(
    s => s.getShapeType() === SlidesApp.ShapeType.TEXT_BOX
  );
  if (boxes.length === 0) return null;
  if (boxes.length === 1) return boxes[0];
  // Multiple text boxes → pick the one with the most area
  return boxes.sort(
    (a, b) => (b.getWidth() * b.getHeight()) - (a.getWidth() * a.getHeight())
  )[0];
}

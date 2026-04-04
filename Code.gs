// ── Lyrics Wrapper — Google Apps Script ───────────────────────────────────────
//
// SETUP:
//   1. Extensions > Apps Script in your Google Slides presentation
//   2. Rename the default "Code.gs" file (or replace its contents with this file)
//   3. Add a new HTML file named exactly "Sidebar" and paste Sidebar.html into it
//   4. Save both files (Ctrl+S), close the editor, reload the presentation
//   5. A "Lyrics" menu will appear — click "Open Lyrics Wrapper" to open the panel
//
// The panel handles scripture lookup, chord/lyric wrapping, and slide creation.

const LINES_PER_SLIDE = 5;

// ─── Menu ──────────────────────────────────────────────────────────────────────

function onOpen() {
  SlidesApp.getUi()
    .createMenu('Lyrics')
    .addItem('Open Lyrics Wrapper', 'showSidebar')
    .addToUi();
}

function showSidebar() {
  // Capture template slide index now, while the selection is valid
  const idx = getCurrentSlideIndex_();
  PropertiesService.getUserProperties().setProperty('templateSlideIdx', idx.toString());

  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Lyrics Wrapper');
  SlidesApp.getUi().showSidebar(html);
}

// ─── Scripture fetch (server-side — no CORS issues) ───────────────────────────

function fetchScripture(ref, translation, apiKey) {
  if (translation === 'esv') {
    if (!apiKey) throw new Error('Enter your ESV API key in the panel.');

    const params = [
      'include-headings=false',
      'include-footnotes=false',
      'include-verse-numbers=true',
      'include-short-copyright=false',
      'include-passage-references=false',
      'q=' + encodeURIComponent(ref)
    ].join('&');

    const resp = UrlFetchApp.fetch('https://api.esv.org/v3/passage/text/?' + params, {
      headers: { Authorization: 'Token ' + apiKey },
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200)
      throw new Error('ESV API error ' + resp.getResponseCode() + ': check your API key.');

    const data = JSON.parse(resp.getContentText());
    if (!data.passages || !data.passages[0]) throw new Error('Passage not found.');

    return { verses: parseESVText_(data.passages[0]), canonical: data.canonical || ref };

  } else {
    const resp = UrlFetchApp.fetch(
      'https://bible-api.com/' + encodeURIComponent(ref) + '?translation=' + translation,
      { muteHttpExceptions: true }
    );
    if (resp.getResponseCode() !== 200)
      throw new Error('Bible API error ' + resp.getResponseCode());

    const data = JSON.parse(resp.getContentText());
    if (data.error) throw new Error(data.error);

    const verses = data.verses.map(function(v) {
      return { verse: v.verse, text: v.text.replace(/[\s\u00a0]+/g, ' ').trim() };
    });
    return { verses: verses, canonical: data.reference || ref };
  }
}

function parseESVText_(raw) {
  var verses = [];
  var parts = raw.split(/\[(\d+)\]/);
  for (var i = 1; i < parts.length; i += 2) {
    var text = (parts[i + 1] || '').replace(/[\s\u00a0]+/g, ' ').trim();
    if (text) verses.push({ verse: parseInt(parts[i]), text: text });
  }
  return verses;
}

// ─── Slide creation ───────────────────────────────────────────────────────────

function createVerseSlides(text) {
  const pres = SlidesApp.getActivePresentation();
  const idx  = parseInt(PropertiesService.getUserProperties().getProperty('templateSlideIdx') || '0');
  const templateSlide = pres.getSlides()[idx];
  if (!templateSlide) throw new Error('Template slide not found. Close and reopen the panel from your template slide.');

  const allLines   = text.split('\n');
  const verseLines = allLines.filter(function(l) { return l.trim() && !l.trim().startsWith('--'); });
  const citation   = (allLines.find(function(l) { return l.trim().startsWith('--'); }) || '').trim();

  if (verseLines.length === 0) throw new Error('No verse text found.');

  // Group verse lines into batches
  const batches = [];
  for (let i = 0; i < verseLines.length; i += LINES_PER_SLIDE) {
    batches.push(verseLines.slice(i, i + LINES_PER_SLIDE));
  }

  const multiSlide = batches.length > 1;

  if (!multiSlide && citation) {
    // Short passage: citation at bottom of single slide
    const last = batches[batches.length - 1];
    if (last.length < LINES_PER_SLIDE) {
      last.push(citation);
    } else {
      batches.push([citation]);
    }
  }

  const templateBox = findContentBox_(templateSlide);
  if (!templateBox) throw new Error('No text box found on the template slide.');

  let insertAt = idx + 1;
  let created  = 0;

  // Multi-slide: reference-only slide first
  if (multiSlide && citation) {
    const refSlide = pres.duplicateSlide(templateSlide);
    pres.moveSlide(refSlide, insertAt);
    const box = findContentBox_(refSlide);
    if (box) box.getText().setText(citation.replace(/^--\s*/, ''));
    insertAt++;
    created++;
  }

  // Content slides
  for (let i = 0; i < batches.length; i++) {
    const newSlide = pres.duplicateSlide(templateSlide);
    pres.moveSlide(newSlide, insertAt + i);
    const box = findContentBox_(newSlide);
    if (box) box.getText().setText(batches[i].join('\n'));
    created++;
  }

  return created;
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function getCurrentSlideIndex_() {
  try {
    const pres  = SlidesApp.getActivePresentation();
    const pages = pres.getSlides();
    const curId = pres.getSelection().getCurrentPage().getObjectId();
    for (let i = 0; i < pages.length; i++) {
      if (pages[i].getObjectId() === curId) return i;
    }
  } catch(e) { /* no selection */ }
  return 0;
}

function findContentBox_(slide) {
  const boxes = slide.getShapes().filter(
    function(s) { return s.getShapeType() === SlidesApp.ShapeType.TEXT_BOX; }
  );
  if (boxes.length === 0) return null;
  if (boxes.length === 1) return boxes[0];
  return boxes.sort(function(a, b) {
    return (b.getWidth() * b.getHeight()) - (a.getWidth() * a.getHeight());
  })[0];
}

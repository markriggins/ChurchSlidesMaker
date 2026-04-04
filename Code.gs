// ── ChurchSlidesMaker — Google Apps Script ────────────────────────────────────
//
// SETUP:
//   1. Extensions > Apps Script in your Google Slides presentation
//   2. Rename the default "Code.gs" file (or replace its contents with this file)
//   3. Add a new HTML file named exactly "Sidebar" and paste Sidebar.html into it
//   4. Save both files (Ctrl+S), close the editor, reload the presentation
//   5. A "ChurchSlidesMaker" menu will appear — click "Open ChurchSlidesMaker" to open the panel
//
// The panel handles scripture lookup, chord/lyric wrapping, and slide creation.

const LINES_PER_SLIDE = 6; // default; overridden by sidebar input

// ─── Menu ──────────────────────────────────────────────────────────────────────

function onOpen(e) {
  SlidesApp.getUi()
    .createMenu('ChurchSlidesMaker')
    .addItem('Open ChurchSlidesMaker', 'showSidebar')
    .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

function showSidebar() {
  // Capture template slide index now, while the selection is valid
  const idx = getCurrentSlideIndex_();
  PropertiesService.getUserProperties().setProperty('templateSlideIdx', idx.toString());

  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('ChurchSlidesMaker');
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

// Rejoin character-wrapped lines back into one line per verse.
// A new logical line starts with a verse number ("16: ") or citation ("--").
// Used for scripture so Google Slides handles visual wrapping itself.
function unwrapScripture_(text) {
  var lines  = text.split('\n');
  var result = [];
  lines.forEach(function(line) {
    var t = line.trim();
    if (!t) return;
    if (/^\d+:/.test(t) || t.startsWith('--')) {
      result.push(t);                              // new verse / citation
    } else if (result.length > 0) {
      result[result.length - 1] += ' ' + t;       // continuation: rejoin
    } else {
      result.push(t);
    }
  });
  return result.join('\n');
}

function createVerseSlides(text, isScripture, linesPerSlide) {
  linesPerSlide = linesPerSlide || LINES_PER_SLIDE;
  const pres = SlidesApp.getActivePresentation();
  const idx  = parseInt(PropertiesService.getUserProperties().getProperty('templateSlideIdx') || '0');
  const templateSlide = pres.getSlides()[idx];
  if (!templateSlide) throw new Error('Template slide not found. Close and reopen the panel from your template slide.');

  const allLines = text.split('\n');
  const citation = (allLines.find(function(l) { return l.trim().startsWith('--'); }) || '').trim();

  // Get verse lines, unwrapping character-wrap newlines for scripture
  // so Google Slides handles visual line-breaking based on the text box size
  var verseText = allLines.filter(function(l) { return l.trim() && !l.trim().startsWith('--'); }).join('\n');
  const verseLines = verseText.split('\n').filter(Boolean);

  if (verseLines.length === 0) throw new Error('No verse text found.');

  // Group wrapped lines back into verse units (a new verse starts with "N: ").
  // If no verse numbers, each line is its own unit.
  var verseGroups = [];
  var curGroup = null;
  verseLines.forEach(function(line) {
    if (/^\d+:/.test(line.trim())) {
      if (curGroup) verseGroups.push(curGroup);
      curGroup = [line];
    } else if (curGroup) {
      curGroup.push(line);
    } else {
      verseGroups.push([line]); // no verse numbers — each line standalone
    }
  });
  if (curGroup) verseGroups.push(curGroup);

  // Pack verse groups into slides; never split a verse across slides.
  const batches = [];
  var cur = [];
  verseGroups.forEach(function(group) {
    if (cur.length > 0 && cur.length + group.length > linesPerSlide) {
      batches.push(cur);
      cur = group.slice();
    } else {
      cur = cur.concat(group);
    }
  });
  if (cur.length > 0) batches.push(cur);

  const multiSlide = batches.length > 1;

  if (!multiSlide && citation) {
    // Short passage: citation at bottom of single slide
    const last = batches[batches.length - 1];
    if (last.length < linesPerSlide) last.push(citation);
    else batches.push([citation]);
  }

  const templateBox = findContentBox_(templateSlide);
  if (!templateBox) throw new Error('No text box found on the template slide.');

  // slide.duplicate() places the copy immediately after the source slide,
  // so we chain duplications to keep everything in order.
  let lastSlide = templateSlide;
  let created   = 0;

  // Multi-slide: reference-only slide first
  if (multiSlide && citation) {
    const refSlide = lastSlide.duplicate();
    const box = findContentBox_(refSlide);
    if (box) box.getText().setText(citation.replace(/^--\s*/, ''));
    lastSlide = refSlide;
    created++;
  }

  // Content slides
  for (let i = 0; i < batches.length; i++) {
    const newSlide = lastSlide.duplicate();
    const box = findContentBox_(newSlide);
    if (box) {
      if (isScripture) {
        box.getText().setText(batches[i].join('\n'));
      } else {
        // Lyrics: body gets lyric lines only; notes get full chord+lyric text
        const lyricOnly = batches[i].filter(function(l) { return !isChordLine_(l); });
        box.getText().setText(lyricOnly.join('\n'));
        try {
          newSlide.getNotesPage().getSpeakerNotesShape().getText()
            .setText(batches[i].join('\n'));
        } catch(e) { /* notes unavailable */ }
      }
    }
    lastSlide = newSlide;
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

function isChordLine_(line) {
  var words = line.trim().split(/\s+/).filter(Boolean);
  if (!words.length) return false;
  var pat = /^[A-G][#b]?(m(aj)?|min|dim|aug|sus[24]?|add\d+)?\d*(\([^)]*\))?(\/[A-G][#b]?)?$/;
  return words.every(function(w) { return pat.test(w); });
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

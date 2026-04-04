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

// ─── Song library (mattgraham/worship on GitHub) ─────────────────────────────

const GITHUB_HEADERS = { 'User-Agent': 'ChurchSlidesMaker' };

function fetchSongList() {
  // Use the recursive git tree API — one request for the full file list.
  const resp = UrlFetchApp.fetch(
    'https://api.github.com/repos/mattgraham/worship/git/trees/HEAD?recursive=1',
    { headers: GITHUB_HEADERS, muteHttpExceptions: true }
  );
  if (resp.getResponseCode() !== 200)
    throw new Error('Could not reach song library (HTTP ' + resp.getResponseCode() + ').');
  const tree = JSON.parse(resp.getContentText()).tree || [];
  return tree
    .filter(function(f) { return f.type === 'blob' && f.path.endsWith('.onsong'); })
    .map(function(f) {
      return {
        name: f.path.replace(/\.onsong$/i, ''),
        url: 'https://raw.githubusercontent.com/mattgraham/worship/master/' + encodeURIComponent(f.path)
      };
    })
    .sort(function(a, b) { return a.name.localeCompare(b.name); });
}

function fetchSong(url) {
  const resp = UrlFetchApp.fetch(url, { headers: GITHUB_HEADERS, muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) throw new Error('Song not found.');
  return resp.getContentText();
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

function createVerseSlides(text, isScripture, linesPerSlide, chordFormat) {
  linesPerSlide = linesPerSlide || LINES_PER_SLIDE;
  chordFormat = chordFormat || 'above';
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

  // Group lines into atomic units that must never be split across slides:
  //   • Scripture: all wrapped lines of one verse stay together (grouped by "N: " prefix)
  //   • Lyrics: a chord line is always paired with its following lyric line
  var verseGroups = [];
  var curGroup = null;
  var vi = 0;
  while (vi < verseLines.length) {
    var line = verseLines[vi];
    var t = line.trim();
    if (/^\d+:/.test(t)) {
      // Start of a new scripture verse
      if (curGroup) verseGroups.push(curGroup);
      curGroup = [line];
      vi++;
    } else if (curGroup) {
      // Continuation line of a scripture verse
      curGroup.push(line);
      vi++;
    } else if (isChordLine_(t) && vi + 1 < verseLines.length && !isChordLine_(verseLines[vi + 1].trim())) {
      // Chord line paired with its lyric — keep together
      verseGroups.push([line, verseLines[vi + 1]]);
      vi += 2;
    } else {
      // Section label, standalone lyric, or orphan chord at end
      verseGroups.push([line]);
      vi++;
    }
  }
  if (curGroup) verseGroups.push(curGroup);

  // Count only lines that will appear in the slide text box.
  function displayCount(lines) {
    return lines.filter(function(l) { return !isChordLine_(l) && !isSectionLabel_(l); }).length;
  }

  // Pack verse groups into slides; never split a verse across slides.
  const batches = [];
  var cur = [];
  verseGroups.forEach(function(group) {
    if (cur.length > 0 && displayCount(cur) + displayCount(group) > linesPerSlide) {
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
        const lyricOnly = batches[i]
          .filter(function(l) { return !isChordLine_(l) && !isSectionLabel_(l); })
          .map(function(l) { return l.replace(/\[[A-G][^\]]*\]/g, ''); }); // strip any inline [Chord] markers
        box.getText().setText(lyricOnly.join('\n'));
        try {
          newSlide.getNotesPage().getSpeakerNotesShape().getText()
            .setText(chordFormat === 'inline' ? batches[i].join('\n') : convertInlineChordsForNotes_(batches[i]).join('\n'));
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

// Convert any remaining inline [Chord]lyric markers to chord-above format for notes.
function convertInlineChordsForNotes_(lines) {
  var out = [];
  lines.forEach(function(line) {
    if (line.indexOf('[') < 0 || !/\[[A-G]/.test(line)) { out.push(line); return; }
    var chords = '', lyrics = '', pos = 0, j = 0;
    while (j < line.length) {
      if (line[j] === '[') {
        var end = line.indexOf(']', j);
        if (end < 0) { lyrics += line[j]; pos++; j++; return; }
        var chord = line.slice(j + 1, end);
        if (/^[A-G]/.test(chord)) {
          while (chords.length < pos) chords += ' ';
          chords += chord;
        }
        j = end + 1;
      } else { lyrics += line[j]; pos++; j++; }
    }
    if (chords.trim()) out.push(chords.trimRight());
    if (lyrics.trim()) out.push(lyrics.trim());
  });
  return out;
}

function isSectionLabel_(line) {
  return /^\[?(Verse|Chorus|Bridge|Pre-?Chorus|Intro|Outro|Tag|Interlude|Vamp)\b/i.test(line.trim());
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

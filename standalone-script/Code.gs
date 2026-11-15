/**
 * Standalone helper for Eastern Gate service decks: insert song slides and
 * build scripture slides. Not part of ChurchSlidesMaker. Never runs Reset.
 * Never deletes or reorders existing slides. Never modifies slides whose
 * notes contain <<...>> markers, except to read them when locating a
 * section or to copy the scriptures template's look onto new slides.
 */

var SONGS_SECTION_RE = /<<\s*songs section\s*>>/i;
var SCRIPTURES_SECTION_RE = /<<\s*scriptures\s*>>/i;
var SECTION_MARKER_RE = /<<[^>]+>>/;
var CHEVRON_RE = /<</;

var ZZ_TEMPLATE_ID = '1mdlmL7xdlQBAvHKc1H7tyIBYSuW1gcv6BKBh5Q6TysQ';
var ZZ_SCRIPTURE_SLIDE_ID = 'SLIDES_API1704141085_0';

var WRAP_COLS = 30;
var LINES_PER_SLIDE = 6;

// ---------------------------------------------------------------------------
// Songs
// ---------------------------------------------------------------------------

/**
 * Insert song slides into a service deck.
 *
 * @param {Object} config
 * @param {string} config.targetId - Service deck ID.
 * @param {number|null} [config.insertAt] - 0-based insert index. If null or
 *     omitted, inserts after the last song slide following the
 *     <<songs section>> marker (or right after the marker if there are none).
 *     insertAt is required when the target has no <<songs section>> marker.
 * @param {Array<{sourceId: string, repeat?: number}>} config.songs
 * @param {boolean} [config.applyNotesFormat=false] - If true, apply the
 *     lyric/chord font and size from the <<songs section>> marker's speaker
 *     notes to each inserted slide's notes. Default false.
 * @return {Object} Manifest of inserted slides.
 */
function insertSongs(config) {
  config = config || {};
  var targetId = config.targetId;
  if (!targetId) {
    throw new Error('config.targetId is required');
  }
  var songs = config.songs;
  if (!songs || !songs.length) {
    throw new Error('config.songs must be a non-empty array');
  }

  var applyNotesFormat = config.applyNotesFormat === true;
  var target = SlidesApp.openById(targetId);
  var targetTitle = target.getName();
  var slides = target.getSlides();

  var markerInfo = findSongsMarker(slides);
  var notesFormat = null;
  if (applyNotesFormat && markerInfo.index !== -1) {
    notesFormat = getNotesFormat(slides[markerInfo.index]);
  }

  var insertAt = resolveSongInsertAt_(config.insertAt, slides, markerInfo);

  var manifest = {
    targetId: targetId,
    targetTitle: targetTitle,
    insertAt: insertAt,
    applyNotesFormat: applyNotesFormat,
    inserted: [],
    notesContainingChevrons: []
  };

  var cursor = insertAt;

  for (var s = 0; s < songs.length; s++) {
    var song = songs[s] || {};
    var sourceId = song.sourceId;
    if (!sourceId) {
      throw new Error('songs[' + s + '].sourceId is required');
    }
    var repeat = song.repeat == null ? 1 : song.repeat;
    if (typeof repeat !== 'number' || repeat < 1 || repeat !== Math.floor(repeat)) {
      throw new Error('songs[' + s + '].repeat must be a positive integer (got ' + song.repeat + ')');
    }
    if (sourceId === targetId) {
      throw new Error('songs[' + s + '].sourceId must not be the target deck');
    }

    var source = SlidesApp.openById(sourceId);
    var sourceTitle = source.getName();
    var sourceSlides = source.getSlides();

    for (var r = 1; r <= repeat; r++) {
      for (var i = 0; i < sourceSlides.length; i++) {
        var sourceSlide = sourceSlides[i];
        var sourceNotes = getNotesText(sourceSlide);
        var inserted = target.insertSlide(
          cursor,
          sourceSlide,
          SlidesApp.SlideLinkingMode.NOT_LINKED
        );
        var insertedNotes = getNotesText(inserted);
        if (insertedNotes !== sourceNotes) {
          setNotesText(inserted, sourceNotes);
        }
        if (notesFormat) {
          applyNotesFormatToSlide(inserted, notesFormat);
        }

        var entry = {
          index: cursor,
          objectId: inserted.getObjectId(),
          sourceId: sourceId,
          sourceTitle: sourceTitle,
          sourceSlideIndex: i,
          copyNumber: r
        };
        var notesNow = getNotesText(inserted);
        if (CHEVRON_RE.test(notesNow)) {
          entry.notesContainChevrons = true;
          manifest.notesContainingChevrons.push({
            index: cursor,
            objectId: entry.objectId,
            sourceTitle: sourceTitle,
            copyNumber: r,
            preview: notesNow.substring(0, 120)
          });
        }
        manifest.inserted.push(entry);
        cursor++;
      }
    }
  }

  Logger.log(JSON.stringify(manifest, null, 2));
  return manifest;
}

/**
 * Apps Script API / execution-API entry point.
 * Accepts a config object or a JSON string.
 */
function runInsertSongs(config) {
  if (typeof config === 'string') {
    config = JSON.parse(config);
  }
  return insertSongs(config);
}

/**
 * Read JSON config from Script Properties key SONG_CONFIG and insert.
 */
function runFromProperties() {
  var raw = PropertiesService.getScriptProperties().getProperty('SONG_CONFIG');
  if (!raw) {
    throw new Error('Script property SONG_CONFIG is missing or empty');
  }
  return insertSongs(JSON.parse(raw));
}

/**
 * Store a config object or JSON string in Script Properties as SONG_CONFIG
 * so a later runFromProperties() needs no code edit.
 */
function setSongConfig(config) {
  var json = typeof config === 'string' ? config : JSON.stringify(config);
  JSON.parse(json);
  PropertiesService.getScriptProperties().setProperty('SONG_CONFIG', json);
  return 'SONG_CONFIG saved';
}

function resolveSongInsertAt_(insertAt, slides, markerInfo) {
  if (insertAt !== null && insertAt !== undefined && insertAt !== '') {
    if (typeof insertAt !== 'number' || insertAt < 0 || insertAt !== Math.floor(insertAt)) {
      throw new Error('config.insertAt must be a non-negative integer or null');
    }
    if (insertAt > slides.length) {
      throw new Error('config.insertAt ' + insertAt + ' is past the end of the deck (' + slides.length + ' slides)');
    }
    return insertAt;
  }
  if (markerInfo.index === -1) {
    throw new Error('config.insertAt is required when the target has no <<songs section>> marker slide');
  }
  // After the last song following the marker, or right after the marker if none.
  // The song block is slides after the marker and before the next <<...>> marker
  // (or the end of the deck if there is no later marker).
  return markerInfo.songBlockEnd;
}

/**
 * Locate the first <<songs section>> marker and the end of its song block.
 * Read-only: does not modify any slide.
 */
function findSongsMarker(slides) {
  var songsMarkerIndex = -1;
  var songBlockEnd = -1;
  for (var i = 0; i < slides.length; i++) {
    var notes = getNotesText(slides[i]);
    if (songsMarkerIndex === -1) {
      if (SONGS_SECTION_RE.test(notes)) {
        songsMarkerIndex = i;
        songBlockEnd = i + 1;
      }
    } else if (SECTION_MARKER_RE.test(notes)) {
      songBlockEnd = i;
      break;
    } else {
      songBlockEnd = i + 1;
    }
  }
  return { index: songsMarkerIndex, songBlockEnd: songBlockEnd };
}

// ---------------------------------------------------------------------------
// Scriptures
// ---------------------------------------------------------------------------

/**
 * Build scripture slides and insert them into a service deck, using the
 * look of the target's <<scriptures>> template slide (or ZZ_Template if
 * the target has no marker).
 *
 * @param {Object} config
 * @param {string} config.targetId
 * @param {number|null} [config.insertAt] - 0-based index. If null or omitted,
 *     inserts at the end of the deck.
 * @param {string} [config.translation='web'] - 'web' (bible-api.com),
 *     'kjv' (bible-api.com), or 'esv' (api.esv.org; requires Script Property
 *     ESV_API_KEY, which is never logged or printed).
 * @param {boolean} [config.verseNumbers=true]
 * @param {boolean} [config.heHim=false] - If true, capitalize he/him/his
 *     (and he's/he'd/he'll). Default false.
 * @param {Array<{ref: string, notes?: string}>} config.passages
 * @return {Object} Manifest of inserted slides.
 */
function insertScriptures(config) {
  config = config || {};
  var targetId = config.targetId;
  if (!targetId) {
    throw new Error('config.targetId is required');
  }
  var passages = config.passages;
  if (!passages || !passages.length) {
    throw new Error('config.passages must be a non-empty array');
  }

  var translation = String(config.translation || 'web').toLowerCase();
  if (translation !== 'web' && translation !== 'kjv' && translation !== 'esv') {
    throw new Error("config.translation must be 'web', 'kjv', or 'esv'");
  }
  var verseNumbers = config.verseNumbers !== false;
  var heHim = config.heHim === true;

  var target = SlidesApp.openById(targetId);
  var targetTitle = target.getName();
  var slides = target.getSlides();
  var insertAt = resolveScriptureInsertAt_(config.insertAt, slides.length);

  var template = getScriptureTemplateSlide_(target);
  var templateStyle = captureShapeStyle_(findVerseShape_(template.slide));

  var manifest = {
    targetId: targetId,
    targetTitle: targetTitle,
    insertAt: insertAt,
    translation: translation,
    verseNumbers: verseNumbers,
    heHim: heHim,
    templateSource: template.source,
    inserted: [],
    passages: [],
    notesSkipped: []
  };

  var cursor = insertAt;

  for (var p = 0; p < passages.length; p++) {
    var passage = passages[p] || {};
    var rawRef = passage.ref;
    if (!rawRef) {
      throw new Error('passages[' + p + '].ref is required');
    }
    var ref = normalizeRef_(rawRef);
    // Pre-fetched text (e.g. ESV fetched elsewhere) skips the network call,
    // so no API key has to live in the script.
    var passageTranslation = String(passage.translation || translation).toLowerCase();
    var fetched = (passage.verses && passage.verses.length)
      ? { verses: passage.verses.map(function (x) { return { verse: x.verse, text: x.text }; }),
          displayRef: passage.displayRef || ref }
      : fetchPassage_(ref, passageTranslation);
    var verses = fetched.verses;
    if (heHim) {
      for (var v = 0; v < verses.length; v++) {
        verses[v].text = applyHeHim_(verses[v].text);
      }
    }
    var displayRef = fetched.displayRef || ref;
    var pages = paginatePassage_(verses, displayRef, verseNumbers);
    var notesDecision = scriptureNotesToApply_(passage.notes);

    var passageInfo = {
      ref: ref,
      displayRef: displayRef,
      verseCount: verses.length,
      slideCount: pages.length,
      notesApplied: false
    };

    for (var s = 0; s < pages.length; s++) {
      var inserted = target.insertSlide(
        cursor,
        template.slide,
        SlidesApp.SlideLinkingMode.NOT_LINKED
      );
      // Always clear copied notes first so <<scriptures>> (or any marker)
      // never lands on a generated slide.
      setNotesText(inserted, '');
      if (s === 0 && notesDecision.ok && notesDecision.text) {
        setNotesText(inserted, notesDecision.text);
        passageInfo.notesApplied = true;
      } else if (s === 0 && notesDecision.skipped) {
        manifest.notesSkipped.push({
          ref: displayRef,
          reason: notesDecision.reason
        });
      }

      var shape = findVerseShape_(inserted);
      shape.getText().setText(pages[s].text);
      applyShapeStyle_(shape, templateStyle);

      manifest.inserted.push({
        index: cursor,
        objectId: inserted.getObjectId(),
        ref: displayRef,
        role: pages[s].role,
        passageIndex: p,
        slideInPassage: s
      });
      cursor++;
    }

    manifest.passages.push(passageInfo);
  }

  Logger.log(JSON.stringify(manifest, null, 2));
  return manifest;
}

/**
 * Apps Script API / execution-API entry point for scriptures.
 * Accepts a config object or a JSON string.
 */
function runInsertScriptures(config) {
  if (typeof config === 'string') {
    config = JSON.parse(config);
  }
  return insertScriptures(config);
}

/**
 * Read JSON config from Script Properties key SCRIPTURE_CONFIG and insert.
 */
function runScripturesFromProperties() {
  var raw = PropertiesService.getScriptProperties().getProperty('SCRIPTURE_CONFIG');
  if (!raw) {
    throw new Error('Script property SCRIPTURE_CONFIG is missing or empty');
  }
  return insertScriptures(JSON.parse(raw));
}

/**
 * Store a config object or JSON string as SCRIPTURE_CONFIG.
 * Never put an ESV API key in this config; use the ESV_API_KEY property.
 */
function setScriptureConfig(config) {
  var json = typeof config === 'string' ? config : JSON.stringify(config);
  JSON.parse(json);
  PropertiesService.getScriptProperties().setProperty('SCRIPTURE_CONFIG', json);
  return 'SCRIPTURE_CONFIG saved';
}

function resolveScriptureInsertAt_(insertAt, slideCount) {
  if (insertAt === null || insertAt === undefined || insertAt === '') {
    return slideCount;
  }
  if (typeof insertAt !== 'number' || insertAt < 0 || insertAt !== Math.floor(insertAt)) {
    throw new Error('config.insertAt must be a non-negative integer or null');
  }
  if (insertAt > slideCount) {
    throw new Error('config.insertAt ' + insertAt + ' is past the end of the deck (' + slideCount + ' slides)');
  }
  return insertAt;
}

/**
 * Prefer the target's <<scriptures>> marker slide. If it has none, use
 * ZZ_Template slide SLIDES_API1704141085_0. Read-only on both sources.
 */
function getScriptureTemplateSlide_(target) {
  var slides = target.getSlides();
  for (var i = 0; i < slides.length; i++) {
    if (SCRIPTURES_SECTION_RE.test(getNotesText(slides[i]))) {
      return { slide: slides[i], source: 'target' };
    }
  }
  var zz = SlidesApp.openById(ZZ_TEMPLATE_ID);
  var zzSlides = zz.getSlides();
  for (var j = 0; j < zzSlides.length; j++) {
    if (zzSlides[j].getObjectId() === ZZ_SCRIPTURE_SLIDE_ID) {
      return { slide: zzSlides[j], source: 'zz_template' };
    }
  }
  for (var k = 0; k < zzSlides.length; k++) {
    if (SCRIPTURES_SECTION_RE.test(getNotesText(zzSlides[k]))) {
      return { slide: zzSlides[k], source: 'zz_template' };
    }
  }
  throw new Error('No <<scriptures>> template slide found in the target or ZZ_Template');
}

function findVerseShape_(slide) {
  var elements = slide.getPageElements();
  var best = null;
  var bestScore = -1;
  for (var i = 0; i < elements.length; i++) {
    var el = elements[i];
    if (el.getPageElementType() !== SlidesApp.PageElementType.SHAPE) {
      continue;
    }
    var shape = el.asShape();
    var text;
    try {
      text = shape.getText().asString();
    } catch (err) {
      continue;
    }
    if (text == null || text === '') {
      continue;
    }
    var score = text.length;
    if (/--/.test(text)) {
      score += 10000;
    }
    if (score > bestScore) {
      bestScore = score;
      best = shape;
    }
  }
  if (!best) {
    throw new Error('Scripture template slide has no text box to fill');
  }
  return best;
}

function captureShapeStyle_(shape) {
  var text = shape.getText();
  var style = text.getTextStyle();
  var captured = {
    fontFamily: style.getFontFamily(),
    fontSize: style.getFontSize(),
    bold: null,
    italic: null,
    foregroundColor: null,
    alignment: null
  };
  try {
    captured.bold = style.isBold();
  } catch (e1) {}
  try {
    captured.italic = style.isItalic();
  } catch (e2) {}
  try {
    captured.foregroundColor = style.getForegroundColor();
  } catch (e3) {}
  try {
    captured.alignment = text.getParagraphStyle().getParagraphAlignment();
  } catch (e4) {}
  if ((!captured.fontFamily || captured.fontSize == null) && text.asString().length > 0) {
    var first = text.getRange(0, 1).getTextStyle();
    captured.fontFamily = captured.fontFamily || first.getFontFamily();
    if (captured.fontSize == null) {
      captured.fontSize = first.getFontSize();
    }
    if (captured.bold == null) {
      try {
        captured.bold = first.isBold();
      } catch (e5) {}
    }
    if (!captured.foregroundColor) {
      try {
        captured.foregroundColor = first.getForegroundColor();
      } catch (e6) {}
    }
  }
  return captured;
}

function applyShapeStyle_(shape, captured) {
  if (!captured) {
    return;
  }
  var text = shape.getText();
  var style = text.getTextStyle();
  if (captured.fontFamily) {
    style.setFontFamily(captured.fontFamily);
  }
  if (captured.fontSize != null) {
    style.setFontSize(captured.fontSize);
  }
  if (captured.bold === true || captured.bold === false) {
    style.setBold(captured.bold);
  }
  if (captured.italic === true || captured.italic === false) {
    style.setItalic(captured.italic);
  }
  if (captured.foregroundColor) {
    try {
      style.setForegroundColor(captured.foregroundColor);
    } catch (e1) {}
  }
  if (captured.alignment) {
    try {
      text.getParagraphStyle().setParagraphAlignment(captured.alignment);
    } catch (e2) {}
  }
}

function scriptureNotesToApply_(notes) {
  if (notes == null || notes === '') {
    return { ok: true, text: '' };
  }
  var text = String(notes);
  if (SECTION_MARKER_RE.test(text)) {
    return { skipped: true, reason: 'notes contain a <<...>> marker and were not applied' };
  }
  return { ok: true, text: text };
}

function normalizeRef_(ref) {
  return String(ref).replace(/[\u2013\u2014]/g, '-').replace(/\s+/g, ' ').trim();
}

/**
 * Fetch a passage. The ESV key is read from Script Properties and is never
 * written to logs, the manifest, or error messages.
 */
function fetchPassage_(ref, translation) {
  if (translation === 'esv') {
    return fetchEsv_(ref);
  }
  return fetchBibleApi_(ref, translation);
}

function fetchBibleApi_(ref, translation) {
  var url = 'https://bible-api.com/' + encodeURIComponent(ref) +
    '?translation=' + encodeURIComponent(translation);
  var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
  var code = response.getResponseCode();
  var body = response.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error('bible-api.com failed for "' + ref + '" (HTTP ' + code + ')');
  }
  var data;
  try {
    data = JSON.parse(body);
  } catch (err) {
    throw new Error('bible-api.com returned invalid JSON for "' + ref + '"');
  }
  if (!data.verses || !data.verses.length) {
    throw new Error('bible-api.com returned no verses for "' + ref + '"');
  }
  var verses = [];
  for (var i = 0; i < data.verses.length; i++) {
    var v = data.verses[i];
    var text = String(v.text || '').replace(/\s+/g, ' ').trim();
    if (!text) {
      continue;
    }
    verses.push({
      chapter: v.chapter,
      verse: v.verse,
      text: text
    });
  }
  if (!verses.length) {
    throw new Error('bible-api.com returned empty verse text for "' + ref + '"');
  }
  return { displayRef: data.reference || ref, verses: verses };
}

function fetchEsv_(ref) {
  var key = PropertiesService.getScriptProperties().getProperty('ESV_API_KEY');
  if (!key) {
    throw new Error('ESV requested but Script Property ESV_API_KEY is not set');
  }
  var url = 'https://api.esv.org/v3/passage/text/?q=' + encodeURIComponent(ref) +
    '&include-headings=false&include-footnotes=false&include-footnote-body=false' +
    '&include-verse-numbers=true&include-first-verse-numbers=true' +
    '&include-short-copyright=false&include-copyright=false' +
    '&include-passage-references=false&include-selahs=false' +
    '&include-passage-horizontal-lines=false&include-heading-horizontal-lines=false' +
    '&line-length=0';
  var response = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    followRedirects: true,
    headers: { Authorization: 'Token ' + key }
  });
  var code = response.getResponseCode();
  var body = response.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error('api.esv.org failed for "' + ref + '" (HTTP ' + code + ')');
  }
  var data;
  try {
    data = JSON.parse(body);
  } catch (err) {
    throw new Error('api.esv.org returned invalid JSON for "' + ref + '"');
  }
  var raw = (data.passages && data.passages[0]) ? data.passages[0] : '';
  var verses = parseEsvText_(raw);
  if (!verses.length) {
    throw new Error('api.esv.org returned no verses for "' + ref + '"');
  }
  return { displayRef: data.canonical || ref, verses: verses };
}

function parseEsvText_(raw) {
  var text = String(raw || '')
    .replace(/\r/g, '')
    .replace(/\s*\(ESV\)\s*$/g, '')
    .replace(/^\s*\(?[1-9][A-Za-z0-9 .:;,\-–—]+\)?\s*\n/, '')
    .replace(/[\u00a0]/g, ' ')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n{2,}/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  var matches = [];
  var re = /\[(\d+)\]\s*/g;
  var m;
  while ((m = re.exec(text)) !== null) {
    matches.push({ verse: parseInt(m[1], 10), start: m.index, end: m.index + m[0].length });
  }
  var verses = [];
  if (!matches.length) {
    if (text) {
      verses.push({ verse: 1, text: text });
    }
    return verses;
  }
  for (var i = 0; i < matches.length; i++) {
    var from = matches[i].end;
    var to = i + 1 < matches.length ? matches[i + 1].start : text.length;
    var chunk = text.substring(from, to).trim();
    if (chunk) {
      verses.push({ verse: matches[i].verse, text: chunk });
    }
  }
  return verses;
}

function applyHeHim_(text) {
  return String(text).replace(/\b(he|him|his|he's|he'd|he'll)\b/gi, function (match) {
    return match.charAt(0).toUpperCase() + match.slice(1);
  });
}

/**
 * Wrap verses to ~30 characters per line and ~5-6 lines per slide.
 * A passage that fits in one slide (body + '-- Ref') stays on one slide
 * with the citation as the last line. A longer passage starts with a
 * reference-only slide, then text slides with no citation. Font size is
 * never reduced; extra slides are added instead.
 */
function paginatePassage_(verses, displayRef, verseNumbers) {
  var showNumbers = verseNumbers && verses.length > 1;
  var chapters = {};
  var chapterCount = 0;
  for (var i = 0; i < verses.length; i++) {
    if (verses[i].chapter != null && !chapters[verses[i].chapter]) {
      chapters[verses[i].chapter] = true;
      chapterCount++;
    }
  }
  var multiChapter = chapterCount > 1;

  var bodyLines = [];
  for (var v = 0; v < verses.length; v++) {
    var verse = verses[v];
    var prefix = '';
    if (showNumbers) {
      prefix = multiChapter
        ? verse.chapter + ':' + verse.verse + ': '
        : verse.verse + ': ';
    }
    var wrapped = wrapToColumns_(prefix + verse.text, WRAP_COLS);
    for (var w = 0; w < wrapped.length; w++) {
      bodyLines.push(wrapped[w]);
    }
  }

  var citation = '-- ' + displayRef;
  if (bodyLines.length + 1 <= LINES_PER_SLIDE) {
    return [{ role: 'text', text: bodyLines.concat([citation]).join('\n') }];
  }

  var pages = [{ role: 'reference', text: displayRef }];
  for (var start = 0; start < bodyLines.length; start += LINES_PER_SLIDE) {
    var chunk = bodyLines.slice(start, start + LINES_PER_SLIDE);
    pages.push({ role: 'text', text: chunk.join('\n') });
  }
  return pages;
}

function wrapToColumns_(text, maxChars) {
  var words = String(text || '').trim().split(/\s+/);
  var lines = [];
  var current = '';
  for (var i = 0; i < words.length; i++) {
    var word = words[i];
    if (!word) {
      continue;
    }
    var trial = current ? current + ' ' + word : word;
    if (trial.length <= maxChars) {
      current = trial;
    } else {
      if (current) {
        lines.push(current);
      }
      current = word;
    }
  }
  if (current) {
    lines.push(current);
  }
  return lines;
}

// ---------------------------------------------------------------------------
// Shared notes helpers
// ---------------------------------------------------------------------------

function getNotesText(slide) {
  try {
    var shape = slide.getNotesPage().getSpeakerNotesShape();
    if (!shape) {
      return '';
    }
    var text = shape.getText();
    if (!text) {
      return '';
    }
    return text.asString();
  } catch (err) {
    return '';
  }
}

function setNotesText(slide, notesText) {
  var shape = slide.getNotesPage().getSpeakerNotesShape();
  shape.getText().setText(notesText || '');
}

function getNotesFormat(slide) {
  try {
    var shape = slide.getNotesPage().getSpeakerNotesShape();
    if (!shape) {
      return null;
    }
    var text = shape.getText();
    var style = text.getTextStyle();
    var family = style.getFontFamily();
    var size = style.getFontSize();
    if ((!family || size == null) && text.asString().length > 0) {
      var first = text.getRange(0, 1).getTextStyle();
      family = family || first.getFontFamily();
      if (size == null) {
        size = first.getFontSize();
      }
    }
    if (!family && size == null) {
      return null;
    }
    return { fontFamily: family, fontSize: size };
  } catch (err) {
    return null;
  }
}

function applyNotesFormatToSlide(slide, format) {
  if (!format) {
    return;
  }
  var style = slide.getNotesPage().getSpeakerNotesShape().getText().getTextStyle();
  if (format.fontFamily) {
    style.setFontFamily(format.fontFamily);
  }
  if (format.fontSize != null) {
    style.setFontSize(format.fontSize);
  }
}

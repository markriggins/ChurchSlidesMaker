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

// ─── Song library ─────────────────────────────────────────────────────────────

const GITHUB_HEADERS = { 'User-Agent': 'ChurchSlidesMaker' };

// repoPath = "owner/repo" or "owner/repo/tree/branch/subdir" or a full GitHub URL
function fetchSongList(repoPath, githubToken) {
  repoPath = repoPath || 'mattgraham/worship';

  // Normalise full GitHub URLs → owner/repo[/tree/branch/subdir]
  var urlMatch = repoPath.match(/^(?:https?:\/\/)?github\.com\/(.+)/);
  if (urlMatch) repoPath = urlMatch[1].replace(/\/+$/, '');

  var parts = repoPath.split('/');
  if (parts.length < 2) throw new Error('Invalid repo: use owner/repo format.');
  var owner = parts[0], repo = parts[1];

  // Extract optional subdir, stripping /tree/<branch>/ if present
  var rest = parts.slice(2);
  if (rest[0] === 'tree' && rest.length > 1) rest = rest.slice(2);
  var subdir = rest.join('/');  // '' if none

  const cache = CacheService.getScriptCache();
  const cacheKey = 'songList_' + owner + '_' + repo + (subdir ? '_' + subdir.replace(/\//g, '_') : '');
  const cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  var ghHeaders = Object.assign({}, GITHUB_HEADERS);
  if (githubToken) ghHeaders['Authorization'] = 'token ' + githubToken;

  // Get default branch so raw URLs work for repos that use "main" instead of "master".
  var branch = 'master';
  try {
    const meta = UrlFetchApp.fetch('https://api.github.com/repos/' + owner + '/' + repo,
      { headers: ghHeaders, muteHttpExceptions: true });
    if (meta.getResponseCode() === 200)
      branch = JSON.parse(meta.getContentText()).default_branch || 'master';
  } catch(e) {}

  const resp = UrlFetchApp.fetch(
    'https://api.github.com/repos/' + owner + '/' + repo + '/git/trees/HEAD?recursive=1',
    { headers: ghHeaders, muteHttpExceptions: true }
  );
  if (resp.getResponseCode() !== 200) {
    var errBody = {}; try { errBody = JSON.parse(resp.getContentText()); } catch(e) {}
    if (resp.getResponseCode() === 403)
      throw new Error('GitHub rate limit or access error (403). Wait a minute and try again, or check that the repo is public. (' + (errBody.message || '') + ')');
    if (resp.getResponseCode() === 404)
      throw new Error('Repo not found (404): check the owner/repo name.');
    throw new Error('Could not reach song library (HTTP ' + resp.getResponseCode() + '). Try again in a minute.');
  }
  const tree = JSON.parse(resp.getContentText()).tree || [];
  const songs = tree
    .filter(function(f) {
      if (f.type !== 'blob' || !/\.(onsong|cho)$/i.test(f.path)) return false;
      return !subdir || f.path.startsWith(subdir + '/');
    })
    .map(function(f) {
      // Encode each path segment separately so slashes (subdirectory separators) survive.
      const encodedPath = f.path.split('/').map(encodeURIComponent).join('/');
      var displayPath = subdir ? f.path.slice(subdir.length + 1) : f.path;
      return {
        name: displayPath.replace(/\.(onsong|cho)$/i, ''),
        url: 'https://raw.githubusercontent.com/' + owner + '/' + repo + '/' + branch + '/' + encodedPath
      };
    })
    .sort(function(a, b) { return a.name.localeCompare(b.name); });

  try { cache.put(cacheKey, JSON.stringify(songs), 21600); } catch(e) {}
  return songs;
}

// Lists Google Slides files in a Drive folder.
// folderPath: "Parent/Child" path, a bare folder ID, or null (defaults to EasternGate/Songs).
// Returns [{name, url}] where url = "drive://<fileId>".
function getDriveSongsList(folderPath) {
  folderPath = folderPath || 'EasternGate/Songs';

  const cache = CacheService.getScriptCache();
  const cacheKey = 'driveSongs_' + folderPath.replace(/[^a-zA-Z0-9]/g, '_');
  const cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  var folder = null;

  // Try as a folder ID (26+ alphanumeric characters).
  if (/^[a-zA-Z0-9_-]{20,}$/.test(folderPath)) {
    try { folder = DriveApp.getFolderById(folderPath); } catch(e) {}
  }

  // Fall back to resolving as a slash-separated folder path.
  if (!folder) {
    var parts = folderPath.split('/').map(function(p) { return p.trim(); }).filter(Boolean);
    if (!parts.length) throw new Error('Invalid folder path: ' + folderPath);
    var tops = DriveApp.getFoldersByName(parts[0]);
    if (!tops.hasNext()) throw new Error('Folder "' + parts[0] + '" not found in Google Drive.');
    folder = tops.next();
    for (var i = 1; i < parts.length; i++) {
      var subs = folder.getFoldersByName(parts[i]);
      if (!subs.hasNext()) throw new Error('Subfolder "' + parts[i] + '" not found inside "' + parts[i-1] + '".');
      folder = subs.next();
    }
  }

  var songs = [];
  var files = folder.getFilesByType(MimeType.GOOGLE_SLIDES);
  while (files.hasNext()) {
    var f = files.next();
    songs.push({ name: f.getName(), url: 'drive://' + f.getId() });
  }
  songs.sort(function(a, b) { return a.name.localeCompare(b.name); });

  try { cache.put(cacheKey, JSON.stringify(songs), 3600); } catch(e) {}
  return songs;
}

// Appends all slides from a Drive presentation into the active presentation.
function insertDriveSlides(fileId) {
  var src  = SlidesApp.openById(fileId);
  var dest = SlidesApp.getActivePresentation();
  var section = findSectionMarker_(dest, '<<songs section>>');
  var insertAt = section ? findSectionEnd_(dest, section.index) : dest.getSlides().length;
  var count = 0;
  src.getSlides().forEach(function(slide) {
    dest.insertSlide(insertAt + count, slide);
    count++;
  });
  return count;
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
  // Prefer section-marker slides in the presentation over the legacy template-slide index.
  var templateSlide, useMarkerInsert = false, insertionIdx, legacyLastSlide;
  var marker  = isScripture ? '<<scriptures>>' : '<<songs section>>';
  var section = findSectionMarker_(pres, marker);
  if (section) {
    templateSlide    = section.slide;
    insertionIdx     = findSectionEnd_(pres, section.index);
    useMarkerInsert  = true;
  } else {
    const idx = parseInt(PropertiesService.getUserProperties().getProperty('templateSlideIdx') || '0');
    templateSlide = pres.getSlides()[idx];
    if (!templateSlide) throw new Error('Template slide not found. Add <<songs section>> or <<scriptures>> to a slide\'s speaker notes, or reopen the panel from your template slide.');
    legacyLastSlide = templateSlide;
  }

  const allLines = text.split('\n');
  const citation = (allLines.find(function(l) { return l.trim().startsWith('--'); }) || '').trim();

  // Get verse lines, preserving blank lines as stanza-break signals.
  var verseLines = allLines.filter(function(l) { return !l.trim().startsWith('--'); });

  if (!verseLines.some(function(l) { return l.trim(); })) throw new Error('No verse text found.');

  // Group lines into atomic units that must never be split across slides:
  //   • null  = blank-line stanza break (forces a new slide)
  //   • Scripture: verse-number line is its own group; each continuation line is also
  //     its own group so the batcher can split long verses across slides
  //   • Lyrics: a chord line is always paired with its following lyric line;
  //     wrapped continuations (U+200B prefix from client) stay with their originating group
  var verseGroups = [];
  var curGroup = null;
  var vi = 0;
  while (vi < verseLines.length) {
    var line = verseLines[vi];
    var isCont = line.charCodeAt(0) === 0x200B;
    var cleanLine = isCont ? line.slice(1) : line;
    var t = cleanLine.trim();
    if (!t) {
      // Blank line — stanza break signal
      if (curGroup) { verseGroups.push(curGroup); curGroup = null; }
      verseGroups.push(null);
      vi++;
    } else if (!isScripture && isCont && verseGroups.length > 0 && verseGroups[verseGroups.length - 1] !== null) {
      // Wrapped-line continuation — keep with the previous lyric group
      verseGroups[verseGroups.length - 1].push(cleanLine);
      vi++;
    } else if (/^\d+:/.test(t)) {
      // Start of a new scripture verse
      if (curGroup) verseGroups.push(curGroup);
      curGroup = [cleanLine];
      vi++;
    } else if (curGroup) {
      // Continuation line of a scripture verse — flush the verse-number group,
      // then treat this line as its own group so the batcher can split long verses
      verseGroups.push(curGroup);
      curGroup = null;
      verseGroups.push([cleanLine]);
      vi++;
    } else if (isChordLine_(t) && vi + 1 < verseLines.length) {
      var nextRaw = verseLines[vi + 1];
      var nextClean = nextRaw.charCodeAt(0) === 0x200B ? nextRaw.slice(1) : nextRaw;
      if (nextClean.trim() && !isChordLine_(nextClean.trim())) {
        // Chord line paired with its lyric — keep together
        verseGroups.push([cleanLine, nextClean]);
        vi += 2;
      } else {
        verseGroups.push([cleanLine]);
        vi++;
      }
    } else {
      // Section label, standalone lyric, or orphan chord at end
      verseGroups.push([cleanLine]);
      vi++;
    }
  }
  if (curGroup) verseGroups.push(curGroup);

  // Count only lines that will appear in the slide text box.
  function displayCount(lines) {
    return lines.filter(function(l) { return !isChordLine_(l) && !isSectionLabel_(l); }).length;
  }

  // Split verseGroups into sections separated by blank lines or section labels.
  // Each section gets its own evenly-distributed slide allocation.
  var sections = [];
  var curSection = [];
  verseGroups.forEach(function(group) {
    if (group === null) {
      if (curSection.length > 0) { sections.push(curSection); curSection = []; }
      return;
    }
    var isSection = !isScripture && group.length === 1 && isSectionLabel_(group[0]);
    if (isSection) {
      if (curSection.length > 0) sections.push(curSection);
      curSection = [group]; // section label starts its own section
    } else {
      curSection.push(group);
    }
  });
  if (curSection.length > 0) sections.push(curSection);

  // Batch each section with even distribution: ceil(total/linesPerSlide) slides,
  // each holding at most ceil(total/k) lines — so 8 lines / linesPerSlide=6 → 4+4 not 6+2.
  const batches = [];
  sections.forEach(function(section) {
    var total = section.reduce(function(sum, g) { return sum + displayCount(g); }, 0);
    var k = Math.max(1, Math.ceil(total / linesPerSlide));
    var target = Math.ceil(total / k);

    var cur = [];
    section.forEach(function(group) {
      if (cur.length > 0 && displayCount(cur) + displayCount(group) > target) {
        batches.push(cur);
        cur = group.slice();
      } else {
        cur = cur.concat(group);
      }
    });
    if (cur.length > 0) batches.push(cur);
  });

  const multiSlide = batches.length > 1;

  if (!multiSlide && citation) {
    // Short passage: citation at bottom of single slide
    const last = batches[batches.length - 1];
    if (last.length < linesPerSlide) last.push(citation);
    else batches.push([citation]);
  }

  const templateBox = findContentBox_(templateSlide);
  if (!templateBox) throw new Error('No text box found on the template slide.');

  // Read template font size and alignment (alignment as REST API string).
  var templateFontSize = null, templateAlignStr = null;
  try { templateFontSize = templateBox.getText().getTextStyle().getFontSize(); } catch(e) {}
  if (!templateFontSize) {
    // GAS returns null when font size is inherited from theme/layout — fall back to REST API.
    try {
      var _fd = Slides.Presentations.get(pres.getId(), {
        fields: 'slides(pageElements(objectId,shape/text/textElements/textRun/style/fontSize))'
      });
      var _ftid = templateBox.getObjectId();
      (_fd.slides || []).some(function(sd) {
        return (sd.pageElements || []).some(function(el) {
          if (el.objectId !== _ftid) return false;
          return (((el.shape || {}).text || {}).textElements || []).some(function(te) {
            if (te.textRun && te.textRun.style && te.textRun.style.fontSize) {
              templateFontSize = te.textRun.style.fontSize.magnitude;
              return true;
            }
            return false;
          });
        });
      });
    } catch(e) {}
  }
  try {
    var _ta = templateBox.getText().getParagraphStyle().getParagraphAlignment();
    if      (_ta === SlidesApp.ParagraphAlignment.CENTER)    templateAlignStr = 'CENTER';
    else if (_ta === SlidesApp.ParagraphAlignment.END)       templateAlignStr = 'END';
    else if (_ta === SlidesApp.ParagraphAlignment.JUSTIFIED) templateAlignStr = 'JUSTIFIED';
    else if (_ta)                                            templateAlignStr = 'START';
  } catch(e) {}

  // Helper: create the next content slide from the template.
  function nextSlide_() {
    if (useMarkerInsert) {
      pres.insertSlide(insertionIdx, templateSlide);
      var s = pres.getSlides()[insertionIdx];
      insertionIdx++;
      return s;
    } else {
      var s = legacyLastSlide.duplicate();
      legacyLastSlide = s;
      return s;
    }
  }

  // Collect REST requests so font size, alignment, and TEXT_AUTOFIT are applied
  // atomically in one batch — avoids race between GAS setText and REST autofit.
  var restRequests = [];
  function queueSlideStyle_(boxId, fontSizePt) {
    if (fontSizePt) {
      restRequests.push({
        updateTextStyle: {
          objectId: boxId,
          style: { fontSize: { magnitude: fontSizePt, unit: 'PT' } },
          textRange: { type: 'ALL' },
          fields: 'fontSize'
        }
      });
    }
    if (templateAlignStr) {
      restRequests.push({
        updateParagraphStyle: {
          objectId: boxId,
          style: { alignment: templateAlignStr },
          textRange: { type: 'ALL' },
          fields: 'alignment'
        }
      });
    }
    restRequests.push({
      updateShapeProperties: {
        objectId: boxId,
        shapeProperties: { autofit: { autofitType: 'TEXT_AUTOFIT' } },
        fields: 'autofit'
      }
    });
  }

  // Read template notes style so scripture slides can be blanked but keep the same look.
  var templateNotesStyle = null;
  if (isScripture) {
    try {
      var _ns = templateSlide.getNotesPage().getSpeakerNotesShape();
      var _nts = _ns.getText().getTextStyle();
      var _nps = _ns.getText().getParagraphStyle();
      var _na = _nps.getParagraphAlignment();
      var _naStr = null;
      if      (_na === SlidesApp.ParagraphAlignment.CENTER)    _naStr = 'CENTER';
      else if (_na === SlidesApp.ParagraphAlignment.END)       _naStr = 'END';
      else if (_na === SlidesApp.ParagraphAlignment.JUSTIFIED) _naStr = 'JUSTIFIED';
      else if (_na)                                            _naStr = 'START';
      templateNotesStyle = {
        fontSize: _nts.getFontSize(),
        bold:     _nts.isBold(),
        italic:   _nts.isItalic(),
        align:    _naStr
      };
    } catch(e) {}
  }

  // Blanks a scripture slide's notes to a single space (invisible) styled to match the template.
  function blankNotes_(slide) {
    try {
      var n = slide.getNotesPage().getSpeakerNotesShape();
      n.getText().setText(' ');
      if (templateNotesStyle) {
        var ts = n.getText().getTextStyle();
        try { if (templateNotesStyle.fontSize) ts.setFontSize(templateNotesStyle.fontSize); } catch(e2) {}
        try { ts.setBold(!!templateNotesStyle.bold);   } catch(e2) {}
        try { ts.setItalic(!!templateNotesStyle.italic); } catch(e2) {}
        if (templateNotesStyle.align) {
          try {
            n.getText().getParagraphStyle().setParagraphAlignment(
              SlidesApp.ParagraphAlignment[templateNotesStyle.align]
            );
          } catch(e2) {}
        }
      }
    } catch(e) {}
  }

  let created = 0;

  // Multi-slide: reference-only citation slide first
  if (multiSlide && citation) {
    const refSlide = nextSlide_();
    const box = findContentBox_(refSlide);
    if (box) {
      box.getText().setText(citation.replace(/^--\s*/, ''));
      queueSlideStyle_(box.getObjectId(), templateFontSize ? templateFontSize * 1.7 : null);
    }
    blankNotes_(refSlide);
    created++;
  }

  // Content slides
  for (let i = 0; i < batches.length; i++) {
    const newSlide = nextSlide_();
    const box = findContentBox_(newSlide);
    if (box) {
      if (isScripture) {
        const bodyText = batches[i].join('\n');
        box.getText().setText(bodyText);
        if (bodyText) queueSlideStyle_(box.getObjectId(), templateFontSize);
        blankNotes_(newSlide);
      } else {
        // Lyrics: body gets lyric lines only; notes get full chord+lyric text
        const lyricOnly = batches[i]
          .filter(function(l) { return !isChordLine_(l) && !isSectionLabel_(l); })
          .map(function(l) { return l.replace(/\[[A-G][^\]]*\]/g, ''); });
        const bodyText = lyricOnly.join('\n');
        if (!bodyText.trim()) {
          newSlide.remove();
          if (useMarkerInsert) insertionIdx--;
          continue;
        }
        box.getText().setText(bodyText);
        queueSlideStyle_(box.getObjectId(), templateFontSize);
        try {
          newSlide.getNotesPage().getSpeakerNotesShape().getText()
            .setText(chordFormat === 'inline' ? batches[i].join('\n') : convertInlineChordsForNotes_(batches[i]).join('\n'));
        } catch(e) {}
        clearSectionMarkers_(newSlide);
      }
    }
    created++;
  }

  // Apply font size, alignment, and TEXT_AUTOFIT together in one REST batch.
  if (restRequests.length > 0) {
    try {
      Slides.Presentations.batchUpdate({ requests: restRequests }, pres.getId());
    } catch(e) {
      Logger.log('Style batchUpdate error: ' + e.toString());
    }
  }

  return created;
}

// Returns {slide, index} for the first slide whose speaker notes contain marker (case-insensitive).
function findSectionMarker_(pres, marker) {
  var slides = pres.getSlides();
  var lc = marker.toLowerCase();
  for (var i = 0; i < slides.length; i++) {
    try {
      var notes = slides[i].getNotesPage().getSpeakerNotesShape().getText().asString().toLowerCase();
      if (notes.indexOf(lc) >= 0) return { slide: slides[i], index: i };
    } catch(e) {}
  }
  return null;
}

// Returns the index at which to insert (just before the next section marker, or end of deck).
function findSectionEnd_(pres, startIdx) {
  var slides = pres.getSlides();
  for (var i = startIdx + 1; i < slides.length; i++) {
    try {
      var notes = slides[i].getNotesPage().getSpeakerNotesShape().getText().asString();
      if (/<<[^>]+>>/.test(notes)) return i;
    } catch(e) {}
  }
  return slides.length;
}

// Removes <<marker>> tokens from a slide's speaker notes so it isn't treated as a section header.
function clearSectionMarkers_(slide) {
  try {
    var shape = slide.getNotesPage().getSpeakerNotesShape();
    var txt = shape.getText().asString();
    var cleaned = txt.replace(/<<[^>]+>>/g, '').trim();
    if (cleaned !== txt.trim()) shape.getText().setText(cleaned);
  } catch(e) {}
}

// ─── Convert Word Art to text boxes ──────────────────────────────────────────

// Shadow at 45° (down-right), 5pt offset, 0 blur, 80% black
// Transform values in EMU (1 PT = 12700 EMU); 5pt * cos(45°) * 12700 ≈ 44907
var DEFAULT_SHADOW = {
  type: 'OUTER',
  color: { rgbColor: { red: 0, green: 0, blue: 0 } },
  alpha: 0.8,
  transform: {
    scaleX: 1, scaleY: 1, shearX: 0, shearY: 0,
    translateX: 44907, translateY: 44907,
    unit: 'EMU'
  },
  alignment: 'BOTTOM_LEFT',
  propertyState: 'RENDERED'
};

function convertWordArt(fromSlide, toSlide) {
  const pres      = SlidesApp.getActivePresentation();
  const allSlides = pres.getSlides();
  const presId    = pres.getId();
  const start     = Math.max(0, (fromSlide || 1) - 1);
  const end       = toSlide ? Math.min(allSlides.length, toSlide) : allSlides.length;

  // ── Style defaults ──
  var fontFamily = 'Comic Sans MS';
  var fontSize   = 50;
  var bold = true, italic = false;
  var fgRed = 1.0, fgGreen = 1.0, fgBlue = 1.0;   // REST API uses 0–1 floats
  var restAlign  = 'CENTER';
  var shadow     = DEFAULT_SHADOW;

  // ── Fetch all REST data in one call: position/size, Word Art text, shadows ──
  const presData = Slides.Presentations.get(presId, {
    fields: 'slides(objectId,pageElements(objectId,size,transform,wordArt/renderedText,shape/shapeProperties/shadow))'
  });

  // ── Override style from first-slide text box if present ──
  const srcBox = findAnyTextShape_(allSlides[start]);
  if (srcBox) {
    const ts = srcBox.getText().getTextStyle();
    const ps = srcBox.getText().getParagraphStyle();
    fontFamily = ts.getFontFamily() || fontFamily;
    fontSize   = ts.getFontSize()   || fontSize;
    bold       = ts.isBold();
    italic     = ts.isItalic();
    try {
      const fg = ts.getForegroundColor();
      if (fg) {
        const rgb = fg.asRgbColor();
        fgRed   = rgb.getRed()   / 255;
        fgGreen = rgb.getGreen() / 255;
        fgBlue  = rgb.getBlue()  / 255;
      }
    } catch(e) {}
    try {
      const a = ps.getParagraphAlignment();
      const s = a ? a.toString() : '';
      if (['CENTER','START','END','JUSTIFIED'].indexOf(s) >= 0) restAlign = s;
    } catch(e) {}

    // Copy shadow from srcBox via REST response
    const srcId = srcBox.getObjectId();
    (presData.slides || []).forEach(function(sd) {
      (sd.pageElements || []).forEach(function(el) {
        if (el.objectId === srcId && el.shape &&
            el.shape.shapeProperties && el.shape.shapeProperties.shadow) {
          shadow = el.shape.shapeProperties.shadow;
        }
      });
    });
  }

  // ── Build set of slide objectIds in range ──
  const inRange = {};
  allSlides.slice(start, end).forEach(function(s) { inRange[s.getObjectId()] = true; });

  // ── Collect Word Art elements in range ──
  const wordArts = [];
  (presData.slides || []).forEach(function(sd) {
    if (!inRange[sd.objectId]) return;
    (sd.pageElements || []).forEach(function(el) {
      if (el.wordArt) {
        wordArts.push({
          pageObjectId: sd.objectId,
          deleteId:     el.objectId,
          text:         el.wordArt.renderedText || '',
          size:         el.size,
          transform:    el.transform
        });
      }
    });
  });

  if (wordArts.length === 0) return 0;

  // ── Build one batchUpdate: delete Word Art + create styled text boxes ──
  const requests = [];
  wordArts.forEach(function(wa, i) {
    const newId = 'wa_box_' + i + '_' + presId.slice(-8);

    requests.push({ deleteObject: { objectId: wa.deleteId } });

    requests.push({
      createShape: {
        objectId:          newId,
        shapeType:         'TEXT_BOX',
        elementProperties: { pageObjectId: wa.pageObjectId, size: wa.size, transform: wa.transform }
      }
    });

    requests.push({
      insertText: { objectId: newId, insertionIndex: 0, text: wa.text }
    });

    requests.push({
      updateTextStyle: {
        objectId:  newId,
        style: {
          fontFamily: fontFamily,
          fontSize:   { magnitude: fontSize, unit: 'PT' },
          bold:       bold,
          italic:     italic,
          foregroundColor: { opaqueColor: { rgbColor: { red: fgRed, green: fgGreen, blue: fgBlue } } }
        },
        textRange: { type: 'ALL' },
        fields: 'fontFamily,fontSize,bold,italic,foregroundColor'
      }
    });

    requests.push({
      updateParagraphStyle: {
        objectId:  newId,
        style:     { alignment: restAlign },
        textRange: { type: 'ALL' },
        fields:    'alignment'
      }
    });

    requests.push({
      updateShapeProperties: {
        objectId:        newId,
        shapeProperties: { shapeBackgroundFill: { propertyState: 'NOT_RENDERED' }, shadow: shadow },
        fields:          'shapeBackgroundFill,shadow'
      }
    });
  });

  Logger.log('convertWordArt shadow: ' + JSON.stringify(shadow));
  Slides.Presentations.batchUpdate({ requests: requests }, presId);

  // Verify: read back the shadow on the first new box
  if (wordArts.length > 0) {
    const checkId = 'wa_box_0_' + presId.slice(-8);
    const check = Slides.Presentations.get(presId, { fields: 'slides(pageElements(objectId,shape/shapeProperties/shadow))' });
    (check.slides || []).forEach(function(sd) {
      (sd.pageElements || []).forEach(function(el) {
        if (el.objectId === checkId)
          Logger.log('shadow read back: ' + JSON.stringify(el.shape && el.shape.shapeProperties && el.shape.shapeProperties.shadow));
      });
    });
  }

  return wordArts.length;
}

// ─── Normalize slides ─────────────────────────────────────────────────────────
//
// Copies the content box's text style and shape shadow from the first slide in
// the range to all subsequent slides. Also removes empty text boxes.

function normalizeSlides(fromSlide, toSlide) {
  const pres      = SlidesApp.getActivePresentation();
  const allSlides = pres.getSlides();
  const start = Math.max(0, (fromSlide || 1) - 1);
  const end   = toSlide ? Math.min(allSlides.length, toSlide) : allSlides.length;
  if (end <= start + 1) throw new Error('Select at least 2 slides.');

  const srcSlide = allSlides[start];
  const srcBox   = findAnyTextShape_(srcSlide);

  var fontFamily = null, fontSize = null, bold = null, italic = null, alignment = null;
  var fgRed = null, fgGreen = null, fgBlue = null;
  if (srcBox) {
    const srcTextStyle = srcBox.getText().getTextStyle();
    const srcParaStyle = srcBox.getText().getParagraphStyle();
    fontFamily = srcTextStyle.getFontFamily();
    fontSize   = srcTextStyle.getFontSize();
    bold       = srcTextStyle.isBold();
    italic     = srcTextStyle.isItalic();
    alignment  = srcParaStyle.getParagraphAlignment();
    try {
      const fg = srcTextStyle.getForegroundColor();
      if (fg) { const rgb = fg.asRgbColor(); fgRed = rgb.getRed(); fgGreen = rgb.getGreen(); fgBlue = rgb.getBlue(); }
    } catch(e) {}
  }

  // Copy background to all target slides, including Word Art slides.
  copyBackground_(pres, srcSlide, allSlides, start, end);

  let count = 0;
  for (let i = start + 1; i < end; i++) {
    const slide  = allSlides[i];
    // Skip text-style normalization for Word Art slides — background already copied above.
    if (slide.getShapes().length === 0 && slide.getPageElements().length > 0) continue;
    const shapes = slide.getShapes();

    shapes.forEach(function(shape) {
      var textRange;
      try { textRange = shape.getText(); } catch(e) { return; } // skip non-text shapes
      // Remove empty text boxes (TEXT_BOX type only — don't delete layout placeholders)
      const text = textRange.asString().replace(/[\n\s]/g, '');
      if (!text) {
        if (shape.getShapeType() === SlidesApp.ShapeType.TEXT_BOX) shape.remove();
        return;
      }
      // Apply source text style to content box
      const ts = shape.getText().getTextStyle();
      const ps = shape.getText().getParagraphStyle();
      try { if (fontFamily) ts.setFontFamily(fontFamily); } catch(e) {}
      try { if (fontSize)   ts.setFontSize(fontSize);     } catch(e) {}
      try { ts.setBold(bold);     } catch(e) {}
      try { ts.setItalic(italic); } catch(e) {}
      try { if (fgRed !== null) ts.setForegroundColor(fgRed, fgGreen, fgBlue); } catch(e) {}
      try { if (alignment) ps.setParagraphAlignment(alignment); } catch(e) {}
    });
    count++;
  }

  // Copy drop shadow via REST API (only if source has a text box)
  if (srcBox) copyShadow_(pres, srcBox, allSlides, start, end);

  return count;
}

function copyShadow_(pres, srcBox, allSlides, start, end) {
  const presId  = pres.getId();
  const srcId   = srcBox.getObjectId();
  const data    = Slides.Presentations.get(presId, { fields: 'slides(pageElements(objectId,shape/shapeProperties/shadow))' });
  var srcShadow = null;

  (data.slides || []).forEach(function(sd) {
    (sd.pageElements || []).forEach(function(el) {
      if (el.objectId === srcId && el.shape && el.shape.shapeProperties && el.shape.shapeProperties.shadow)
        srcShadow = el.shape.shapeProperties.shadow;
    });
  });

  Logger.log('copyShadow_ srcShadow: ' + JSON.stringify(srcShadow));
  // If srcBox has no explicit shadow (inherits from theme), fall back to DEFAULT_SHADOW
  if (!srcShadow) {
    Logger.log('copyShadow_: no shadow found on srcBox ' + srcId + ', using DEFAULT_SHADOW');
    srcShadow = DEFAULT_SHADOW;
  }

  const requests = [];
  for (let i = start + 1; i < end; i++) {
    const box = findAnyTextShape_(allSlides[i]);
    if (!box) continue;
    requests.push({ updateShapeProperties: { objectId: box.getObjectId(), shapeProperties: { shadow: srcShadow }, fields: 'shadow' } });
  }
  if (requests.length) Slides.Presentations.batchUpdate({ requests: requests }, presId);
}

function copyBackground_(pres, srcSlide, allSlides, start, end) {
  const presId = pres.getId();
  const data = Slides.Presentations.get(presId, {
    fields: 'slides(objectId,pageProperties/pageBackgroundFill)'
  });

  var srcBg = null;
  const srcId = srcSlide.getObjectId();
  (data.slides || []).forEach(function(sd) {
    if (sd.objectId === srcId && sd.pageProperties && sd.pageProperties.pageBackgroundFill)
      srcBg = sd.pageProperties.pageBackgroundFill;
  });
  if (!srcBg) return;

  const requests = [];
  for (let i = start + 1; i < end; i++) {
    requests.push({
      updatePageProperties: {
        objectId: allSlides[i].getObjectId(),
        pageProperties: { pageBackgroundFill: srcBg },
        fields: 'pageBackgroundFill'
      }
    });
  }
  if (requests.length) Slides.Presentations.batchUpdate({ requests: requests }, presId);
}

// ─── Transpose selected slides' speaker notes ─────────────────────────────────

function transposeSelectedNotes(steps, useFlats, fromSlide, toSlide) {
  const pres = SlidesApp.getActivePresentation();
  const allSlides = pres.getSlides();
  const start = Math.max(0, (fromSlide || 1) - 1);
  const end   = toSlide ? Math.min(allSlides.length, toSlide) : allSlides.length;
  const slides = allSlides.slice(start, end);

  let count = 0;
  slides.forEach(function(slide) {
    try {
      const notesShape = slide.getNotesPage().getSpeakerNotesShape();
      const original = notesShape.getText().asString().replace(/\n$/, '');
      if (!original.trim()) return;
      const transposed = original.split('\n').map(function(line) {
        return isChordLine_(line.trim()) ? transposeChordLine_(line, steps, useFlats) : line;
      }).join('\n');
      notesShape.getText().setText(transposed);
      count++;
    } catch(e) { /* slide has no notes */ }
  });
  return count;
}

function transposeNote_(note, steps, useFlats) {
  const SHARPS = ['C','C#','D','D#','E','F','F#','G','G#','A','A#','B'];
  const FLATS  = ['C','Db','D','Eb','E','F','Gb','G','Ab','A','Bb','B'];
  var i = SHARPS.indexOf(note);
  if (i < 0) i = FLATS.indexOf(note);
  if (i < 0) return note;
  return (useFlats ? FLATS : SHARPS)[((i + steps) % 12 + 12) % 12];
}

function transposeChordToken_(token, steps, useFlats) {
  // Strip outer parens: (G) → G; re-add after transposing
  var outer = token.length > 2 && token[0] === '(' && token[token.length - 1] === ')';
  if (outer) token = token.slice(1, -1);
  var slash = token.indexOf('/', 1);
  var result;
  if (slash > 0) {
    result = transposeChordToken_(token.slice(0, slash), steps, useFlats) +
             '/' + transposeChordToken_(token.slice(slash + 1), steps, useFlats);
  } else {
    var hasAcc = token.length > 1 && (token[1] === '#' || token[1] === 'b');
    var root = hasAcc ? token.slice(0, 2) : token[0];
    result = transposeNote_(root, steps, useFlats) + token.slice(root.length);
  }
  return outer ? '(' + result + ')' : result;
}

function transposeChordLine_(line, steps, useFlats) {
  return line.replace(/[^\s]+/g, function(t) { return transposeChordToken_(t, steps, useFlats); });
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function getSlideCount() {
  return SlidesApp.getActivePresentation().getSlides().length;
}

function getSelectedSlideRange() {
  try {
    const pres   = SlidesApp.getActivePresentation();
    const slides = pres.getSlides();
    const sel    = pres.getSelection();
    var from, to;
    const pageRange = sel.getPageRange();
    if (pageRange) {
      const pages = pageRange.getPages();
      const ids   = slides.map(function(s) { return s.getObjectId(); });
      const nums  = pages.map(function(p) { return ids.indexOf(p.getObjectId()) + 1; }).filter(function(n) { return n > 0; });
      if (nums.length) { from = Math.min.apply(null, nums); to = Math.max.apply(null, nums); }
    }
    if (!from) {
      const curId = sel.getCurrentPage().getObjectId();
      const idx   = slides.findIndex(function(s) { return s.getObjectId() === curId; });
      if (idx >= 0) { from = idx + 1; to = idx + 1; }
    }
    if (!from) return null;
    // Check for Word Art in range
    const hasWordArt = slides.slice(from - 1, to).some(function(slide) {
      return slide.getPageElements().some(function(el) {
        return el.getPageElementType() === SlidesApp.PageElementType.WORD_ART;
      });
    });
    return { from: from, to: to, hasWordArt: hasWordArt };
  } catch(e) {}
  return null;
}

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
          if (chords.length > 0 && chords[chords.length - 1] !== ' ') chords += ' ';
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
  var pat = /^[A-G][#b]?(m(?:aj)?|min|dim|aug|sus[24]?|add\d+)?\d*(?:\([^)]*\))?(?:\/[A-G][#b]?)?$/;
  return words.every(function(w) {
    if (w.length > 2 && w[0] === '(' && w[w.length - 1] === ')') w = w.slice(1, -1);
    return pat.test(w);
  });
}

// Like findContentBox_ but accepts any shape type that has text content.
// Used for normalizing pre-existing slides that may use rectangles or placeholders.
function diagnoseSlide2() {
  var pres   = SlidesApp.getActivePresentation();
  var slides = pres.getSlides();
  Logger.log('Total slides: ' + slides.length);

  var slide = slides[5]; // slide 6
  var allElements = slide.getPageElements();
  Logger.log('getPageElements() count: ' + allElements.length);
  for (var i = 0; i < allElements.length; i++) {
    Logger.log('  el[' + i + '] pageElementType=' + allElements[i].getPageElementType());
  }

  var shapes = slide.getShapes();
  Logger.log('getShapes() count: ' + shapes.length);
  for (var j = 0; j < shapes.length; j++) {
    var s = shapes[j];
    var txt = '';
    try { txt = s.getText().asString().slice(0, 50); } catch(e) { txt = '(error: ' + e.message + ')'; }
    Logger.log('  shape[' + j + '] shapeType=' + s.getShapeType() + ' text="' + txt + '"');
  }
}

function findAnyTextShape_(slide) {
  const shapes = slide.getShapes().filter(function(s) {
    try { return s.getText().asString().replace(/[\n\s]/g, '').length > 0; } catch(e) { return false; }
  });
  if (shapes.length === 0) return null;
  if (shapes.length === 1) return shapes[0];
  return shapes.sort(function(a, b) {
    return (b.getWidth() * b.getHeight()) - (a.getWidth() * a.getHeight());
  })[0];
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

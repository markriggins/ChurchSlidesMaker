// Harvest *background pictures only* from EasternGate/Songs Google Slides into
// EasternGate/Songs/Backgrounds.
//
// Typical EGA workflow: set a picture on slide 1 and duplicate. That copies the
// fill (or a full-bleed image sent-to-back) onto each slide — not the master.
// We read each slide's own picture fill first, then a full-bleed IMAGE element
// (covers ≥80% of the page). Never notes, WordArt, text, or small clipart.
// Duplicates hash to one file.
//
// Incremental: _harvest_state.json in Backgrounds stores each presentation's
// last-updated time. Unchanged decks are skipped. Re-run until the toast says done
// (6-minute Apps Script cap). Force = ignore prior timestamps.

var HARVEST_DEADLINE_MS_ = 4.5 * 60 * 1000;
var HARVEST_STATE_NAME_ = '_harvest_state.json';

function harvestSongBackgrounds() {
  harvestSongBackgrounds_(false);
}

function harvestSongBackgroundsForce() {
  harvestSongBackgrounds_(true);
}

function harvestSongBackgrounds_(force) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) {
    SlidesApp.getUi().alert('Harvest already running.');
    return;
  }
  try {
    var songs = findSongsFolder_();
    var dest = getOrCreateFolder_(songs, 'Backgrounds');
    var state = loadHarvestState_(dest);
    if (force) state.files = {};
    var hashes = {};
    (state.hashes || []).forEach(function (h) { hashes[h] = true; });

    var decks = listSlideDecks_(songs);
    var deadline = Date.now() + HARVEST_DEADLINE_MS_;
    var processed = 0;
    var skippedUnchanged = 0;
    var kept = 0;
    var skippedNoFill = 0;
    var remaining = 0;

    for (var i = 0; i < decks.length; i++) {
      var file = decks[i];
      var rec = (state.files || {})[file.getId()];
      var mtime = file.getLastUpdated().getTime();
      if (!force && rec && rec.mtime === mtime) {
        skippedUnchanged++;
        continue;
      }
      if (Date.now() > deadline) {
        remaining = decks.length - i;
        break;
      }
      var result = harvestPresentation_(file, dest, hashes);
      kept += result.kept;
      skippedNoFill += result.skippedNoFill;
      processed++;
      state.files = state.files || {};
      state.files[file.getId()] = {
        name: file.getName(),
        mtime: mtime,
        kept: result.kept,
        skippedNoFill: result.skippedNoFill,
        processedAt: new Date().toISOString()
      };
    }

    state.hashes = Object.keys(hashes);
    state.lastRun = new Date().toISOString();
    saveHarvestState_(dest, state);

    var msg =
      'Processed ' + processed + ' deck(s). Kept ' + kept +
      ' new background(s). Skipped ' + skippedNoFill +
      ' slide(s) with no picture fill. ' +
      skippedUnchanged + ' unchanged (skipped).';
    if (remaining > 0) {
      msg += ' Stopped with ' + remaining + ' left — run Harvest again.';
    } else {
      msg += ' Caught up.';
    }
    SlidesApp.getUi().alert(msg);
  } finally {
    lock.releaseLock();
  }
}

function findSongsFolder_() {
  var roots = DriveApp.getFoldersByName('EasternGate');
  while (roots.hasNext()) {
    var eg = roots.next();
    var songs = eg.getFoldersByName('Songs');
    if (songs.hasNext()) return songs.next();
  }
  throw new Error('Could not find Drive folder EasternGate/Songs');
}

function getOrCreateFolder_(parent, name) {
  var it = parent.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return parent.createFolder(name);
}

function listSlideDecks_(folder) {
  var out = [];
  var files = folder.getFilesByType(MimeType.GOOGLE_SLIDES);
  while (files.hasNext()) out.push(files.next());
  return out;
}

function harvestPresentation_(file, dest, hashes) {
  var kept = 0;
  var skippedNoFill = 0;
  var pres = Slides.Presentations.get(file.getId());
  var pageW = mag_(pres.pageSize && pres.pageSize.width);
  var pageH = mag_(pres.pageSize && pres.pageSize.height);
  var urlCache = {};
  var name = file.getName();

  function take(page, tag) {
    if (!page || isNotesPage_(page)) return;
    var url = pictureFillUrl_(page);
    if (!url) url = fullBleedImageUrl_(page, pageW, pageH);
    if (!url) {
      skippedNoFill++;
      return;
    }
    var blob = fetchImageUrl_(url, urlCache);
    if (!blob) {
      skippedNoFill++;
      return;
    }
    if (saveHarvestBlob_(blob, dest, hashes, name, tag)) kept++;
  }

  (pres.masters || []).forEach(function (p, i) { take(p, 'master' + (i + 1)); });
  (pres.layouts || []).forEach(function (p, i) { take(p, 'layout' + (i + 1)); });
  (pres.slides || []).forEach(function (p, i) { take(p, i + 1); });
  return { kept: kept, skippedNoFill: skippedNoFill };
}

function indexPages_(pages) {
  var out = {};
  (pages || []).forEach(function (p) {
    if (p.objectId) out[p.objectId] = p;
  });
  return out;
}

function mag_(dim) {
  if (dim == null) return 0;
  if (typeof dim === 'number') return dim;
  return dim.magnitude || 0;
}

function isNotesPage_(page) {
  return !!(page && page.notesProperties);
}

function pictureFillUrl_(page) {
  if (!page || isNotesPage_(page)) return null;
  var fill = page.pageProperties && page.pageProperties.pageBackgroundFill;
  if (!fill) return null;
  var pic = fill.stretchedPictureFill || fill.tiledPictureFill;
  return (pic && pic.contentUrl) || null;
}

function fullBleedImageUrl_(page, pageW, pageH) {
  // Image sent-to-back and stretched — common when slide 1 is duplicated.
  // Skip WordArt/shapes (not .image). Skip small photos/clipart.
  if (!page || !page.pageElements || pageW <= 0 || pageH <= 0) return null;
  var minW = pageW * 0.8;
  var minH = pageH * 0.8;
  var bestUrl = null;
  var bestArea = 0;
  var els = page.pageElements;
  for (var i = 0; i < els.length; i++) {
    var el = els[i];
    if (!el.image || !el.image.contentUrl) continue;
    var t = el.transform || {};
    var w = mag_(el.size && el.size.width) * (t.scaleX || 1);
    var h = mag_(el.size && el.size.height) * (t.scaleY || 1);
    if (w < minW || h < minH) continue;
    var area = w * h;
    if (area > bestArea) {
      bestArea = area;
      bestUrl = el.image.contentUrl;
    }
  }
  return bestUrl;
}

function fetchImageUrl_(url, cache) {
  if (cache[url]) return cache[url] === 'none' ? null : cache[url];
  try {
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) {
      cache[url] = 'none';
      return null;
    }
    var blob = resp.getBlob();
    cache[url] = blob;
    return blob;
  } catch (e) {
    cache[url] = 'none';
    return null;
  }
}

function saveHarvestBlob_(blob, dest, hashes, deckName, pageNum) {
  var digest = blobMd5_(blob);
  if (hashes[digest]) return false;
  hashes[digest] = true;
  var safe = String(deckName).replace(/[^\w.\- ]+/g, '').slice(0, 40).trim() || 'slide';
  blob.setName('slide_' + safe + '_p' + pageNum + '_' + digest.slice(0, 10) + '.png');
  dest.createFile(blob);
  return true;
}

function blobMd5_(blob) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, blob.getBytes());
  return raw.map(function (b) {
    return ('0' + (b & 255).toString(16)).slice(-2);
  }).join('');
}

function loadHarvestState_(dest) {
  var it = dest.getFilesByName(HARVEST_STATE_NAME_);
  if (!it.hasNext()) return { files: {}, hashes: [] };
  try {
    return JSON.parse(it.next().getBlob().getDataAsString());
  } catch (e) {
    return { files: {}, hashes: [] };
  }
}

function saveHarvestState_(dest, state) {
  var json = JSON.stringify(state);
  var it = dest.getFilesByName(HARVEST_STATE_NAME_);
  if (it.hasNext()) {
    it.next().setContent(json);
  } else {
    dest.createFile(HARVEST_STATE_NAME_, json, MimeType.PLAIN_TEXT);
  }
}

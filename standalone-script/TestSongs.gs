// Test runner for the song copy step. Select testSongs in the Apps Script
// editor and click Run. Defaults target the scratch deck
// ZZ-TEST-song-copy-2026-11-15 (a copy of the 2026-11-15 service deck), never
// the real deck. Inserts at the end of the deck; never deletes anything.
//
// Expected: To Him Who Sits on the Throne (C) has 1 slide, so repeat 2 is
// ignored and it is inserted once; Every Praise has 3 slides x 2 = 6.
// Total +7 slides (115 -> 122 if the scratch deck is untouched).
var TEST_SONGS_TARGET_ID = '1MS_XNxStFMk2OOr7_68N80pK6yHeE_L8ETQqRlu493g';
var TEST_SONGS_LIST = [
  // To Him Who Sits on the Throne (C)
  { fileId: '1yUFXlqHKPjPUAyNVDXKIALJOQP88k3P6szEOqQs1hNE', repeat: 2 },
  // Every Praise
  { fileId: '1ItdDEy5gRoS_bV-7lu84Z4C3ehFh6r2kQLLosyFoMAQ', repeat: 2 }
];

function testSongs(targetDeckId, songs) {
  var target = targetDeckId || TEST_SONGS_TARGET_ID;
  if (target === '15epVCo1KnTY9HEFpYWaKCLZ1ahSVadX_aaxA-K0UuIc') {
    throw new Error('testSongs must not run on the real 2026-11-15 deck');
  }
  var result = insertSongsAt(target, songs || TEST_SONGS_LIST, null);
  Logger.log('testSongs done: ' + result.insertedCount + ' slide(s) inserted, ' +
      result.slideCountBefore + ' -> ' + result.slideCountAfter);
  return result;
}

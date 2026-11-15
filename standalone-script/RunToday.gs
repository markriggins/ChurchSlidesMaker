// Test run for 2026-11-15 deck: Psalm 150 ESV, appended at end of deck.
function runToday() {
  var result = insertScriptures({
  "targetId": "15epVCo1KnTY9HEFpYWaKCLZ1ahSVadX_aaxA-K0UuIc",
  "translation": "esv",
  "verseNumbers": true,
  "heHim": false,
  "insertAt": null,
  "passages": [
    {
      "ref": "Psalm 150",
      "displayRef": "Psalm 150:1-6 (ESV)",
      "verses": [
        {
          "verse": 1,
          "text": "Praise the LORD! Praise God in his sanctuary; praise him in his mighty heavens!"
        },
        {
          "verse": 2,
          "text": "Praise him for his mighty deeds; praise him according to his excellent greatness!"
        },
        {
          "verse": 3,
          "text": "Praise him with trumpet sound; praise him with lute and harp!"
        },
        {
          "verse": 4,
          "text": "Praise him with tambourine and dance; praise him with strings and pipe!"
        },
        {
          "verse": 5,
          "text": "Praise him with sounding cymbals; praise him with loud clashing cymbals!"
        },
        {
          "verse": 6,
          "text": "Let everything that has breath praise the LORD! Praise the LORD!"
        }
      ]
    }
  ]
});
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

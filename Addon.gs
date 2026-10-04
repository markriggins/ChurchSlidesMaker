// Workspace add-on homepage (right-rail icon). Installed once; available on every
// Slides file, including File > Make a copy. The HTML sidebar is still the real UI.

function onSlidesHomepage() {
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('ChurchSlidesMaker'))
    .addSection(
      CardService.newCardSection()
        .addWidget(
          CardService.newTextParagraph().setText(
            'Authorize once. This add-on follows you to every sermon copy.'
          )
        )
        .addWidget(
          CardService.newTextButton()
            .setText('Open panel')
            .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
            .setOnClickAction(CardService.newAction().setFunctionName('openSidebarFromCard_'))
        )
        .addWidget(
          CardService.newTextButton()
            .setText('Reset')
            .setOnClickAction(CardService.newAction().setFunctionName('resetFromCard_'))
        )
        .addWidget(
          CardService.newTextButton()
            .setText('Harvest song backgrounds')
            .setOnClickAction(CardService.newAction().setFunctionName('harvestFromCard_'))
        )
    )
    .build();
}

function openSidebarFromCard_() {
  showSidebar();
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification().setText('Opened ChurchSlidesMaker'))
    .build();
}

function resetFromCard_() {
  var r = cleanupGeneratedSlides();
  var msg = 'Reset cancelled.';
  if (r && r.alreadyClean) msg = 'Already at starting point.';
  else if (r && !r.cancelled) msg = 'Removed ' + r.removed + ' slide(s).';
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification().setText(msg))
    .build();
}

function harvestFromCard_() {
  harvestSongBackgrounds();
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification().setText('Harvest finished (see alert).'))
    .build();
}

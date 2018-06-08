#target illustrator

(function () {

  var outPath = null;
  var docName = '';
  var actvDoc = null;
  var foilLayer = null;
  var foilImages = null;
  var foilImagesCount = null;
  var altGreetingsLayer = null;
  var altGreetingsLayers = null;
  var altGreetingsCount = null;
  var position = [0, 0];
  var progressbar = new ProgressBar('Foil Base - Foil Image AI Saver Script');
  var saveOptions = new IllustratorSaveOptions();

  saveOptions.embedICCProfile = true;

  if (!app.documents.length) return;

  outPath = prompt('Paste in SKU destination:', 'C:\\MIN-XXX-YYY - Title\\AI to JPGs');

  if (outPath === null || outPath === '') {
    alert('ERROR: Destination folder not defined.\n\nAborting process...');

    return;
  } else outPath = outPath.replace(/\\+$/, '') + '\\';

  for (var a = 0; a < app.documents.length; a++) {
    docName = app.documents[a].name;

    if (docName.match(/^MIN-[A-Z0-9]{3}-[A-Z0-9]{3}_[A-Z]_[1356]?_?FRT\.ai$/) === null) {
      alert('ERROR: Some file names are not in standard format.\n\nAborting...');

      return;
    }
  }

  actvDoc = app.activeDocument;

  try {
    foilLayer = !!actvDoc.layers['id:foil_artwork'] ? actvDoc.layers['id:foil_artwork'] : null;
  } catch (e) {
    alert('ERROR: "id:foil_artwork" layer is missing.\n\nAborting process...');

    return;
  }

  try {
    altGreetingsLayer = !!actvDoc.layers['alternate greetings'] ? actvDoc.layers['alternate greetings'] : null;
  } catch (e) {
    alert('ERROR: "alternate greetings" layer is missing.\n\nAborting process...');

    return;
  }

  actvDoc.artboards[0].rulerOrigin = [0, 0];
  position = foilLayer.groupItems[0].position;
  foilImagesCount = foilLayer.groupItems[0].rasterItems.length;
  altGreetingsCount = altGreetingsLayer.layers.length;

  if (foilImagesCount < 1) {
    alert('ERROR: No foil images.\n\nAborting process...');

    return;
  } else if (foilImagesCount < 3) if (!confirm('WARNING: Minimum of 3 foil images.\n\nDo you still want to continue?')) return;

  if (altGreetingsCount < 1) {
    alert('ERROR: No alternate greetings.\n\nAborting process...');

    return;
  }

  progressbar.reset('Processing Foil Base Magick...', app.documents.length * foilImagesCount * altGreetingsCount);

  if (!!actvDoc.selection.length)
    for (var b = actvDoc.selection.length - 1; b >= 0; b--)
      actvDoc.selection[b].selected = false;

  foilLayer.groupItems[0].selected = true;

  app.copy();

  actvDoc = null;
  foilLayer = null;
  altGreetingsLayer = null;

  for (var c = app.documents.length - 1; c >= 0; c--, progressbar.hit()) {
    app.documents[c].activate();

    actvDoc = app.activeDocument;
    docName = actvDoc.name.split('.')[0];
    foilLayer = actvDoc.layers['id:foil_artwork'];
    altGreetingsLayer = actvDoc.layers['alternate greetings'];

    for (var d = foilImagesCount - 1; d >= 0; d--, progressbar.hit()) {

      switch (d) {
        case 0:
          foilLayer.groupItems[0].remove();

          actvDoc.artboards[0].rulerOrigin = [0, 0];

          app.paste();

          actvDoc.selection[0].position = position;

          foilImages = foilLayer.groupItems[0].rasterItems;
          altGreetingsLayers = altGreetingsLayer.layers;

          for (var e = foilImagesCount - 1; e >= 0; e--)
            if (d !== e) foilImages[e].remove();

          for (var f = altGreetingsCount - 1; f >= 0; f--) {
            altGreetingsLayers[f].visible = false;
          }

          for (var g = 0; g < altGreetingsCount; g++, progressbar.hit()) {
            altGreetingsLayers[g].visible = true;

            switch (altGreetingsLayers[g].name) {
              case 'holiday greeting':
                actvDoc.saveAs(new File(outPath + docName + 'HOLIDAYGOLD.ai'), saveOptions);
                break;
              case 'christmas greeting':
                actvDoc.saveAs(new File(outPath + docName + 'CHRISTMASGOLD.ai'), saveOptions);
                break;
              case 'new year greeting':
                actvDoc.saveAs(new File(outPath + docName + 'NEWYEARGOLD.ai'), saveOptions);
                break;
              case 'religious greeting':
                actvDoc.saveAs(new File(outPath + docName + 'RELIGIOUSGOLD.ai'), saveOptions);
                break;
              default: break;
            }

            altGreetingsLayers[g].visible = false;
          }

          break;
        case 1:
          foilLayer.groupItems[0].remove();

          actvDoc.artboards[0].rulerOrigin = [0, 0];

          app.paste();

          actvDoc.selection[0].position = position;

          foilImages = foilLayer.groupItems[0].rasterItems;
          altGreetingsLayers = altGreetingsLayer.layers;

          for (var e = foilImagesCount - 1; e >= 0; e--)
            if (d !== e) foilImages[e].remove();

          for (var f = altGreetingsCount - 1; f >= 0; f--) {
            altGreetingsLayers[f].visible = false;
          }

          for (var g = 0; g < altGreetingsCount; g++, progressbar.hit()) {
            altGreetingsLayers[g].visible = true;

            switch (altGreetingsLayers[g].name) {
              case 'holiday greeting':
                actvDoc.saveAs(new File(outPath + docName + 'HOLIDAYROSEGOLD.ai'), saveOptions);
                break;
              case 'christmas greeting':
                actvDoc.saveAs(new File(outPath + docName + 'CHRISTMASROSEGOLD.ai'), saveOptions);
                break;
              case 'new year greeting':
                actvDoc.saveAs(new File(outPath + docName + 'NEWYEARROSEGOLD.ai'), saveOptions);
                break;
              case 'religious greeting':
                actvDoc.saveAs(new File(outPath + docName + 'RELIGIOUSROSEGOLD.ai'), saveOptions);
                break;
              default: break;
            }

            altGreetingsLayers[g].visible = false;
          }

          break;
        case 2:
          foilLayer.groupItems[0].remove();

          actvDoc.artboards[0].rulerOrigin = [0, 0];

          app.paste();

          actvDoc.selection[0].position = position;

          foilImages = foilLayer.groupItems[0].rasterItems;
          altGreetingsLayers = altGreetingsLayer.layers;

          for (var e = foilImagesCount - 1; e >= 0; e--)
            if (d !== e) foilImages[e].remove();

          for (var f = altGreetingsCount - 1; f >= 0; f--) {
            altGreetingsLayers[f].visible = false;
          }

          for (var g = 0; g < altGreetingsCount; g++, progressbar.hit()) {
            altGreetingsLayers[g].visible = true;

            switch (altGreetingsLayers[g].name) {
              case 'holiday greeting':
                actvDoc.saveAs(new File(outPath + docName + 'HOLIDAYSILVER.ai'), saveOptions);
                break;
              case 'christmas greeting':
                actvDoc.saveAs(new File(outPath + docName + 'CHRISTMASSILVER.ai'), saveOptions);
                break;
              case 'new year greeting':
                actvDoc.saveAs(new File(outPath + docName + 'NEWYEARSILVER.ai'), saveOptions);
                break;
              case 'religious greeting':
                actvDoc.saveAs(new File(outPath + docName + 'RELIGIOUSSILVER.ai'), saveOptions);
                break;
              default: break;
            }

            altGreetingsLayers[g].visible = false;
          }

          break;
        case 3:
          foilLayer.groupItems[0].remove();

          actvDoc.artboards[0].rulerOrigin = [0, 0];

          app.paste();

          actvDoc.selection[0].position = position;

          foilImages = foilLayer.groupItems[0].rasterItems;
          altGreetingsLayers = altGreetingsLayer.layers;

          for (var e = foilImagesCount - 1; e >= 0; e--)
            if (d !== e) foilImages[e].remove();

          for (var f = altGreetingsCount - 1; f >= 0; f--) {
            altGreetingsLayers[f].visible = false;
          }

          for (var g = 0; g < altGreetingsCount; g++, progressbar.hit()) {
            altGreetingsLayers[g].visible = true;

            switch (altGreetingsLayers[g].name) {
              case 'holiday greeting':
                actvDoc.saveAs(new File(outPath + docName + 'HOLIDAYGOLDGLITTER.ai'), saveOptions);
                break;
              case 'christmas greeting':
                actvDoc.saveAs(new File(outPath + docName + 'CHRISTMASGOLDGLITTER.ai'), saveOptions);
                break;
              case 'new year greeting':
                actvDoc.saveAs(new File(outPath + docName + 'NEWYEARGOLDGLITTER.ai'), saveOptions);
                break;
              case 'religious greeting':
                actvDoc.saveAs(new File(outPath + docName + 'RELIGIOUSGOLDGLITTER.ai'), saveOptions);
                break;
              default: break;
            }

            altGreetingsLayers[g].visible = false;
          }

          break;
        case 4:
          foilLayer.groupItems[0].remove();

          actvDoc.artboards[0].rulerOrigin = [0, 0];

          app.paste();

          actvDoc.selection[0].position = position;

          foilImages = foilLayer.groupItems[0].rasterItems;
          altGreetingsLayers = altGreetingsLayer.layers;

          for (var e = foilImagesCount - 1; e >= 0; e--)
            if (d !== e) foilImages[e].remove();

          for (var f = altGreetingsCount - 1; f >= 0; f--) {
            altGreetingsLayers[f].visible = false;
          }

          for (var g = 0; g < altGreetingsCount; g++, progressbar.hit()) {
            altGreetingsLayers[g].visible = true;

            switch (altGreetingsLayers[g].name) {
              case 'holiday greeting':
                actvDoc.saveAs(new File(outPath + docName + 'HOLIDAYSILVERGLITTER.ai'), saveOptions);
                break;
              case 'christmas greeting':
                actvDoc.saveAs(new File(outPath + docName + 'CHRISTMASSILVERGLITTER.ai'), saveOptions);
                break;
              case 'new year greeting':
                actvDoc.saveAs(new File(outPath + docName + 'NEWYEARSILVERGLITTER.ai'), saveOptions);
                break;
              case 'religious greeting':
                actvDoc.saveAs(new File(outPath + docName + 'RELIGIOUSSILVERGLITTER.ai'), saveOptions);
                break;
              default: break;
            }

            altGreetingsLayers[g].visible = false;
          }

          break;
        default: break;
      }
    }

    actvDoc.close(SaveOptions.DONOTSAVECHANGES);
  }

  progressbar.close();

  outPath = null;
  docName = '';
  actvDoc = null;
  foilLayer = null;
  foilImages = null;
  foilImagesCount = null;
  altGreetingsLayer = null;
  altGreetingsLayers = null;
  altGreetingsCount = null;
  position = [0, 0];
  progressbar = null;

  return;
})();

// progress bar by Marc Autret
function ProgressBar (title) {
  var palette = new Window('palette', title, {x: 0, y: 0, width: 340, height: 60});
  var bar = palette.add('progressbar', {x: 20, y: 12, width: 300, height: 12}, 0, 100);
  var text = palette.add('statictext', {x: 10, y: 36, width: 320, height: 20}, '');

  text.justify = 'center';
  palette.center();

  this.reset = function(msg, maxVal) {
    text.text = msg;
    bar.value = 0;
    bar.maxvalue = maxVal || 0;
    bar.visible = !!maxVal;

    palette.show();
  };

  this.hit = function () {
    ++bar.value;

    palette.update();
  };

  this.show = function () {
    palette.show();
  };

  this.hide = function () {
    palette.hide();
  };

  this.close = function () {
    palette.close();
  };
}

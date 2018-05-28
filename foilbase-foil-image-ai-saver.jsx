#target illustrator

(function () {

  var outPath = null;
  var docName = '';
  var actvDoc = null;
  var foilLayer = null;
  var position = [0, 0];
  var foilImages = null;
  var foilImagesCount = null;
  var progressbar = new ProgressBar('Foil Base - Foil Image AI Saver Script');
  var saveOptions = new IllustratorSaveOptions();

  saveOptions.embedICCProfile = true;

  if (!app.documents.length) return;

  outPath = prompt('Paste in SKU destination:', 'C:\\MIN-XXX-YYY - Title\\AI to JPGs');

  if (outPath === null || outPath === '') {
    alert('ERROR: Destination folder not defined.\n\nAborting process...');

    return;
  } else outPath += '\\';

  progressbar.reset('Processing Foil Base Magick...', app.documents.length * 4);

  for (var a = 0; a < app.documents.length; a++, progressbar.hit()) {
    docName = app.documents[a].name;

    if (docName.match(/^MIN-[A-Z0-9]{3}-[A-Z0-9]{3}_[A-Z]_[1356]?_?FRT\.ai$/) === null) {
      progressbar.close();

      alert('ERROR: Some file names are not in standard format.\n\nAborting...');

      return;
    }
  }

  actvDoc = app.activeDocument;

  try {
    foilLayer = !!actvDoc.layers['id:foil_artwork'] ? actvDoc.layers['id:foil_artwork'] : null;
  } catch (e) {
    progressbar.close();

    alert('ERROR: "id:foil_artwork" layer is missing.\n\nAborting process...');

    return;
  }

  if (foilLayer.groupItems[0].rasterItems.length < 1) {
    progressbar.close();

    alert('ERROR: No foil images.\n\nAborting process...');

    return;
  } else if (foilLayer.groupItems[0].rasterItems.length < 3) {
    progressbar.hide();

    if (confirm('WARNING: Minimum of 3 foil images.\n\nDo you still want to continue?')) progressbar.show();
    else {
      progressbar.close();

      return;
    }

  }

  actvDoc.artboards[0].rulerOrigin = [0, 0];
  position = foilLayer.groupItems[0].position;
  foilImagesCount = foilLayer.groupItems[0].rasterItems.length;

  if (!!actvDoc.selection.length) {
    for (var b = actvDoc.selection.length - 1; b >= 0; b--) {
      actvDoc.selection[b].selected = false;
    }
  }

  foilLayer.groupItems[0].selected = true;

  app.copy();

  actvDoc = null;
  foilLayer = null;

  for (var c = app.documents.length - 1; c >= 0; c--, progressbar.hit()) {
    app.documents[c].activate();

    actvDoc = app.activeDocument;
    docName = actvDoc.name.split('.')[0];
    foilLayer = actvDoc.layers['id:foil_artwork'];

    for (var d = foilImagesCount - 1; d >= 0; d--, progressbar.hit()) {

      switch (d) {
        case 0:
          foilLayer.groupItems[0].remove();

          actvDoc.artboards[0].rulerOrigin = [0, 0];

          app.paste();

          actvDoc.selection[0].position = position;

          foilImages = foilLayer.groupItems[0].rasterItems;

          for (var e = foilImagesCount - 1; e >= 0; e--) {
            if (d !== e) foilImages[e].remove();
          }

          actvDoc.saveAs(new File(outPath + docName + 'GOLD.ai'), saveOptions);

          break;
        case 1:
          foilLayer.groupItems[0].remove();

          actvDoc.artboards[0].rulerOrigin = [0, 0];

          app.paste();

          actvDoc.selection[0].position = position;

          foilImages = foilLayer.groupItems[0].rasterItems;

          for (var e = foilImagesCount - 1; e >= 0; e--) {
            if (d !== e) foilImages[e].remove();
          }

          actvDoc.saveAs(new File(outPath + docName + 'ROSEGOLD.ai'), saveOptions);

          break;
        case 2:
          foilLayer.groupItems[0].remove();

          actvDoc.artboards[0].rulerOrigin = [0, 0];

          app.paste();

          actvDoc.selection[0].position = position;

          foilImages = foilLayer.groupItems[0].rasterItems;

          for (var e = foilImagesCount - 1; e >= 0; e--) {
            if (d !== e) foilImages[e].remove();
          }

          actvDoc.saveAs(new File(outPath + docName + 'SILVER.ai'), saveOptions);

          break;
        case 3:
          foilLayer.groupItems[0].remove();

          actvDoc.artboards[0].rulerOrigin = [0, 0];

          app.paste();

          actvDoc.selection[0].position = position;

          foilImages = foilLayer.groupItems[0].rasterItems;

          for (var e = foilImagesCount - 1; e >= 0; e--) {
            if (d !== e) foilImages[e].remove();
          }

          actvDoc.saveAs(new File(outPath + docName + 'GOLDGLITTER.ai'), saveOptions);

          break;
        case 4:
          foilLayer.groupItems[0].remove();

          actvDoc.artboards[0].rulerOrigin = [0, 0];

          app.paste();

          actvDoc.selection[0].position = position;

          foilImages = foilLayer.groupItems[0].rasterItems;

          for (var e = foilImagesCount - 1; e >= 0; e--) {
            if (d !== e) foilImages[e].remove();
          }

          actvDoc.saveAs(new File(outPath + docName + 'SILVERGLITTER.ai'), saveOptions);

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
  position = [0, 0];
  foilImages = null;
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

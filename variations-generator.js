function getColumnByName(row, name) {
  var nameNormalised = name.toLowerCase().trim();
  for (var i = 0; i < row.length; i++) {
    if (nameNormalised === row[i].toLowerCase().trim().replace(/\s*\(.*\)/g, '')) {
      return i;
    }
  }

  return -1;
}

function getDefiningProperties(row, columns, properties) {
  var values = row[getColumnByName(columns, properties.shift())].split(',');
  var result = [];

  if (properties.length === 0) {
    for (var i = 0; i < values.length; i++) {
      result.push([values[i].trim()]);
    }

    return result;
  }

  var nextValues = getDefiningProperties(row, columns, properties);
  var variation = null;

  for(var i = 0; i < values.length; i++) {
    for (var j = 0; j < nextValues.length; j++) {
      variation = nextValues[j].slice();
      variation.unshift(values[i].trim());
      result.push(variation);
    }
  }

  return result;
}

function generateName(values) {
  return values.join(', ')
    .replace(/\s*\(.*\)/g, '');
}

function generateSku(base, values) {
  return (base + '-' + values.join('-').toLowerCase())
    .replace(/\s*\(.*\)/g, '')
    .replace(/\s+/g, '-')           // Replace spaces with -
    .replace(/[^\w\-]+/g, '')       // Remove all non-word chars
    .replace(/\-\-+/g, '-')         // Replace multiple - with single -
    .replace(/^-+/, '')             // Trim - from start of text
    .replace(/-+$/, '');
}

function colorHeaders(sheet) {
  var data = sheet.getDataRange().getValues();
  var row = null;

  for (var i = 0; i < data.length; i++) {
    row = data[i];
    if (row[0].toLowerCase() == 'sku') {
      sheet.getRange(i+1, 1, 1, row.length).setBackgroundRGB(182, 225, 205);
    }
  }
}

function generateVariations() {
  var spreadsheet = SpreadsheetApp.getActive();
  var activeSheet = spreadsheet.getActiveSheet();
  var data = activeSheet.getDataRange().getValues();

  var row = null;
  var columns = null;
  var variation = null;
  var result = [];
  var maxColumns = 0;

  for (var i = 0; i < data.length; i++) {
    row = data[i];

    if (row[0].toLowerCase().trim() == 'sku') {
      columns = row;
      result.push(columns);
      maxColumns = Math.max(maxColumns, columns.length);

      continue;
    }

    var definingProperties = row[getColumnByName(columns, 'defining properties')].split(',');
    var values = getDefiningProperties(row, columns, definingProperties.slice());

    for (var j = 0; j < values.length; j++) {
      variation = row.slice();
      variation[getColumnByName(columns, 'variation name')] = generateName(values[j]);
      variation[getColumnByName(columns, 'sku')] = generateSku(variation[getColumnByName(columns, 'sku')], values[j]);

      for (var k = 0; k < definingProperties.length; k++) {
        variation[getColumnByName(columns, definingProperties[k].trim())] = values[j][k];
      }

      result.push(variation);
    }
  }

  var d = new Date();
  var sheetName = 'Generated variations for ' + activeSheet.getName() + ' ' + d.toLocaleTimeString();
  var variationsSheet = spreadsheet.insertSheet(sheetName, spreadsheet.getNumSheets());

  variationsSheet.activate();
  variationsSheet.getRange(1, 1, result.length, maxColumns).setValues(result);

  colorHeaders(variationsSheet);
}

function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Generate variations', functionName: 'generateVariations'}
  ];
  spreadsheet.addMenu('Products', menuItems);
}

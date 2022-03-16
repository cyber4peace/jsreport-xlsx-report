module.exports = (reporter) => async (context, req) => {
  try {
    reporter.logger.debug('Parsing xlsx report content begin', req)

    const workbookPath = 'xl/workbook.xml';
    const stringsPath = 'xl/sharedStrings.xml';
    const calcChainPath = 'xl/calcChain.xml';
    const worksheetPath = 'xl/worksheets/sheet{idx}.xml';

    const workbook = context.$xlsxReport[workbookPath]?.workbook;

    const strings = context.$xlsxReport[stringsPath];
    const stringList = strings.sst?.si ?? [];
    const stringIdxToClear = [];

    const calcList = (context.$xlsxReport[calcChainPath]?.calcChain?.c ?? []).map(item => ({
      calcRef: item,
    }));

    let currentTable = undefined;
    const tables = [];

    const worksheetFiles = workbook.sheets[0].sheet
      .map(item => {
        const fileName = worksheetPath.replace('{idx}', item.$.sheetId);
        return {
          sheetId: item.$.sheetId,
          worksheet: context.$xlsxReport[fileName].worksheet,
        }
      })
      .reduce((prev, worksheetFile) => {
        worksheetFile.worksheet.sheetData[0]?.row?.forEach(elementRow => {
          let lastCell = false;
          let str = '';
          elementRow.c?.forEach(elementCell => {
            if (elementCell.f?.[0]) {
              const calc = calcList
                .find(item => item.calcRef.$.i === worksheetFile.sheetId && item.calcRef.$.r === elementCell.$.r);
              if (calc) {
                calc.cell = elementCell;
              }
            }
            else if (elementCell.v?.[0] && stringList[elementCell.v[0]].t?.[0]) {
              lastCell = false;
              str = String(stringList[elementCell.v[0]].t[0]);
              if (str.startsWith('{{#xlsxTable')) {
                currentTable = {
                  name: str.slice(0, str.indexOf('}}') + 2),
                  sheetId: worksheetFile.sheetId,
                  rowIndex: elementRow.$.r,
                  rows: [
                    elementRow
                  ]
                };
                if (currentTable.name.split(' ').length !== 2) {
                  throw reporter.createError(`Unable to generate xlsx report (maybe you are missing an attribute name at {{#xlsxTable}}`, {
                    weak: true,
                    statusCode: 400
                  })
                }
                str = str.replaceAll(currentTable.name, '');
              }
              if (str.endsWith('{{/xlsxTable}}')) {
                str = str.replaceAll('{{/xlsxTable}}', '');
                tables.push(currentTable);
                currentTable = undefined;
                lastCell = true;
              }
              if (currentTable || lastCell) {
                elementCell.$ = {
                  ...elementCell.$,
                  t: "inlineStr"
                }
                elementCell.is = {
                  t: str
                };
                stringIdxToClear.push(elementCell.v[0]);
                delete elementCell.v;
              }
            }
          });
          if (currentTable) {
            currentTable.rows.push(elementRow);
          }
        });
        return {
          ...prev,
          [worksheetFile.sheetId]: worksheetFile.worksheet,
        }
      }, {});

    stringIdxToClear.forEach(i => stringList[i].t[0] = '');

    if (tables.length && currentTable) {
      throw reporter.createError(`Unable to generate xlsx report (maybe you are missing closing tag {{/xlsxTable}}`, {
        weak: true,
        statusCode: 400
      })
    }

    const contentToRender = JSON.stringify(strings) + '###xlsxFile###' + tables
      .map(table => `${table.name}${JSON.stringify({
        sheetId: table.sheetId,
        rowIndex: table.rowIndex,
        rows: table.rows,
      })}###xlsxRow###{{/xlsxTable}}`)
      .join('');
      
    reporter.logger.debug('Starting child request to render docx dynamic parts', req)

    const { content: newContent } = await reporter.render({
      template: {
        content: contentToRender,
        engine: req.template.engine,
        recipe: 'html',
        helpers: req.template.helpers
      }
    }, req)

    const newContentFiles = newContent.toString().split('###xlsxFile###');
    context.$xlsxReport[stringsPath] = JSON.parse(newContentFiles[0]);

    const firstRow = [];
    if (newContentFiles[1]) {
      const rowsRendered = newContentFiles[1].split('###xlsxRow###').filter(r => r).map(r => JSON.parse(r));
      rowsRendered.reverse();
      rowsRendered.forEach(rowRendered => {
        rowRendered.rows.forEach(row => {
          insertRow(worksheetFiles[rowRendered.sheetId], rowRendered.sheetId, calcList, row, !firstRow.includes(`${rowRendered.sheetId}/${rowRendered.rowIndex}`));
        });
        firstRow.push(`${rowRendered.sheetId}/${rowRendered.rowIndex}`);
      });
    }

    if (context.$xlsxReport[calcChainPath] && calcList.length) {
      context.$xlsxReport[calcChainPath].calcChain.c = calcList.map(calc => calc.calcRef);
    }

    reporter.logger.debug('Parsing xlsx report content finished', req)
    return context;
  } catch (e) {
    throw reporter.createError('Error while executing xlsx report recipe', {
      original: e,
      weak: true
    })
  }
}

function insertRow (worksheet, sheetId, calcList, row, isFirst) {
  const sheetData = Array.isArray(worksheet.sheetData) ? worksheet.sheetData[0] : worksheet.sheetData;
  const rowList = sheetData.row ?? [];
  const mergeList = worksheet.mergeCells[0].mergeCell ?? [];
  const index = row.$.r;
  
  const rowInsert = rowList.find(row => parseInt(row.$?.r ?? '0') >= index);
  const rowInsertIndex = rowInsert ? rowList.indexOf(rowInsert) : rowList.length;
  
  if (isFirst && rowInsert.$.r === row.$.r) {
    row.c?.forEach(cell => {
      if (cell.f) {
        const calc = calcList.find(calc => calc.calcRef.$.i === sheetId && calc.cell.$.r === cell.$.r);
        if (calc) {
          calc.cell = cell;
        }
      }
    });
    rowList.splice(rowInsertIndex, 1, row);
    return;
  }
  rowList.splice(rowInsertIndex, 0, row);
  rowList.forEach((row, idx) => {
    if (idx > rowInsertIndex) {
      const newIndex = parseInt(row.$.r) + 1;
      row.$.r = String(newIndex);
      row.c?.forEach(cell => {
        if (cell.$?.r) {
          cell.$.r = cell.$.r[0] + newIndex;
        }
      })
    }
  });

  mergeList.forEach(merge => {
    if (~merge.$?.ref?.indexOf(':')) {
      const ref = merge.$.ref.split(':');
      const startRange = parseInt(ref[0].slice(1)) + 1;
      const endRange = parseInt(ref[1].slice(1)) + 1;
      
      merge.$.ref = (startRange > index ? ref[0][0] + startRange : ref[0]) + ':' + (endRange > index ? ref[1][0] + endRange : ref[1]);
    }
  });

  row.c?.filter(cell => cell.f).forEach(cell => {
    calcList.push({
      calcRef: {
        "$": {
          "r": cell.$.r,
          "i": sheetId,
        }
      },
      cell,
    });
  });

  calcList.forEach(calc => {
    const oldAddress = regionFromString(calc.calcRef.$.r);
    const newAddress = regionFromString(calc.cell.$.r);
    const offset = newAddress[0].row - oldAddress[0].row;
    if (typeof calc.cell.f[0] === 'object') {
      if (calc.cell.f[0].$?.ref && calc.calcRef.$.r !== calc.cell.$.r) {
        const region = regionFromString(calc.cell.f[0].$.ref);
        if (region[1] && region[1].row >= index) {
          region[1].row = region[1].row + offset;
        }
        region[0].row = newAddress[0].row;
        calc.cell.f[0].$.ref = regionToString(region);
      }
      if (typeof calc.cell.f[0]._ === 'string') {
        calc.cell.f[0]._ = correctFormula(calc.cell.f[0]._, index, offset);
      }      
    } else if (typeof calc.cell.f[0] === 'string') {
      calc.cell.f[0] = correctFormula(calc.cell.f[0], index, offset);
    }
    calc.calcRef.$.r = calc.cell.$.r;
    if (calc.cell.v) {
      delete calc.cell.v;
    }
  });
}

// "SUM(C7:C7)"
function correctFormula (value, insertIndex, offset) {
  const ret = [];
  let cell = '';
  for (let i = 0; i < value.length; i++) {
    if ((/[A-Z]/).test(value[i]) && (cell.length === 0 || (/[:A-Z]/).test(cell.slice(-1)))) {
      cell = cell + value[i];
    }
    else if ((/[0-9]/).test(value[i]) && (cell.length > 0)) {
      cell = cell + value[i];
    }
    else if (value[i] === ':' && (cell.length > 0 || (/[0-9]/).test(cell.slice(-1)))) {
      cell = cell + value[i];
    }
    else {
      if ((/[0-9]/).test(cell.slice(-1))) {
        ret.push(correctRegion(cell, insertIndex, offset));
      }
      else {
        ret.push(cell);
      }
      ret.push(value[i]);
      cell = '';
    }
  }
  if ((/[0-9]/).test(cell.slice(-1))) {
    ret.push(correctRegion(cell, insertIndex, offset));
  }

  return ret.join('');
}

function correctRegion (cell, insertIndex, offset) {
  const region = regionFromString(cell);
  if (region[1]) {
    if (region[1].row >= insertIndex) {
      region[1].row = region[1].row + offset;
    }
  }
  if (!region[1] || (region[0].col != region[1].col)) {
    if (region[0].row >= insertIndex) {
      region[0].row = region[0].row + offset;
    }
  }
  return regionToString(region);
}

function regionFromString (value) {
  return value.split(':').map(item => ({
    col: item[0],
    row: parseInt(item.slice(1)),
  }));
}

function regionToString (value) {
  return value.map(item => `${item.col}${item.row}`).join(':');
}
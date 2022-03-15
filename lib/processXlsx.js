module.exports = (reporter) => async (context, req) => {
  try {
    reporter.logger.debug('Parsing xlsx report content begin', req)

    const stringsPath = 'xl/sharedStrings.xml';
    const worksheetsPath = 'xl/worksheets';

    const strings = context.$xlsxReport[stringsPath];
    if (!strings) {
      throw new Error('Not found strings file');
    }
    const stringList = strings.sst?.si ?? [];

    let currentTable = '';
    const tables = {};

    const files = Object.entries(context.$xlsxReport)
      .filter(([ key, value ]) => key.startsWith(worksheetsPath))
      .reduce((prev, [ key, value ]) => {
        const sheetData = Array.isArray(value.worksheet?.sheetData) ? value.worksheet?.sheetData[0] : value.worksheet?.sheetData;
        sheetData?.row?.forEach(elementRow => {
          let lastCell = false;
          let str = '';
          elementRow.c?.filter(elementCell => elementCell.v?.[0] && stringList[elementCell.v[0]].t?.[0]).forEach(elementCell => {
            lastCell = false;
            str = stringList[elementCell.v[0]].t[0];
            if (str.startsWith('{{#xlsxTable')) {
              currentTable = str.slice(0, str.indexOf('}}') + 2);
              if (currentTable.split(' ').length !== 2) {
                throw reporter.createError(`Unable to generate xlsx report (maybe you are missing an attribute name at {{#xlsxTable}}`, {
                  weak: true,
                  statusCode: 400
                })
              }
              tables[currentTable] = {
                fileName: key,
                rows: [
                  elementRow
                ]
              }
              str = str.replaceAll(currentTable, '');
            }
            if (str.endsWith('{{/xlsxTable}}')) {
              str = str.replaceAll('{{/xlsxTable}}', '');
              currentTable = '';
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
              stringList[elementCell.v[0]].t[0] = '';
              delete elementCell.v;
            }
          });
          if (currentTable) {
            tables[currentTable].rows.push(elementRow);
          }
        });
        return {
          ...prev,
          [key]: value,
        }
      }, {});

    if (Object.keys(tables).length && currentTable) {
      throw reporter.createError(`Unable to generate xlsx report (maybe you are missing closing tag {{/xlsxTable}}`, {
        weak: true,
        statusCode: 400
      })
    }

    const contentToRender = JSON.stringify(strings) + '###xlsxFile###' + Object.entries(tables)
      .map(([ key, value ]) => `${key}${JSON.stringify(value)}###xlsxRow###{{/xlsxTable}}`)
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

    Object.values(tables).forEach(value => {
      value.rows.forEach(row => {
        removeRow(files[value.fileName].worksheet, row);
      })
    });
    if (newContentFiles[1]) {
      const rowsRendered = newContentFiles[1].split('###xlsxRow###').filter(r => r).map(r => JSON.parse(r));
      rowsRendered.reverse();
      rowsRendered.forEach(rowRendered => {
        rowRendered.rows.forEach(row => {
          insertRow(files[rowRendered.fileName].worksheet, row);
        })
      });
    }

    console.log('*** *** ***');
    console.log(JSON.stringify(files));

    reporter.logger.debug('Parsing xlsx report content finished', req)
    return context;
  } catch (e) {
    throw reporter.createError('Error while executing xlsx report recipe', {
      original: e,
      weak: true
    })
  }
}

function removeRow (worksheet, row) {
  const sheetData = Array.isArray(worksheet.sheetData) ? worksheet.sheetData[0] : worksheet.sheetData;
  const rowList = sheetData.row ?? [];
  const index = rowList.indexOf(rowList.find(r => r.$.r == row.$.r));
  rowList.splice(index, 1)
}

function insertRow(worksheet, row) {
  const sheetData = Array.isArray(worksheet.sheetData) ? worksheet.sheetData[0] : worksheet.sheetData;
  const rowList = sheetData.row ?? [];
  const mergeList = sheetData.mergeCell ?? [];
  const index = row.$.r;
  
  const rowInsert = rowList.find(row => parseInt(row.$?.r ?? '0') >= index);
  const rowInsertIndex = rowInsert ? rowList.indexOf(rowInsert) : rowList.length;
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
  })
}
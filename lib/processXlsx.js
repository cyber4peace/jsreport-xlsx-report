module.exports = (reporter) => async (context, req) => {
  try {
    reporter.logger.debug('Parsing xlsx report content begin', req)

    const strings = context.$xlsxReport['xl/sharedStrings.xml'];
    if (!strings) {
      throw new Error('Not found strings file');
    }
    const stringList = strings.sst?.si?.map(item => item.t?.[0]) ?? [];
    const contentToRender = JSON.stringify(strings);

    const worksheets = Object.entries(context.$xlsxReport)
      .filter(([ key, value ]) => key.startsWith('xl/worksheets'))
      .reduce((prev, [ key, value ]) => {
        value.worksheet?.sheetData?.forEach(elementData => {
          elementData.row?.forEach(elementRow => {
            elementRow.c?.forEach(elementCell => {
              if (elementCell.v?.[0]) {
                const isOpenTable = stringList[elementCell.v[0]]?.startsWith('{{#xlsxTable}}');
                const isCloseTable = stringList[elementCell.v[0]]?.startsWith('{{/xlsxTable}}');
                const str = stringList[elementCell.v[0]]?.replaceAll('{{#xlsxTable}}', '').replaceAll('{{/xlsxTable}}', '');
                elementCell.$ = {
                  ...elementCell.$,
                  t: "inlineStr"
                }
                elementCell.is = {
                  t: str
                };
                delete elementCell.v;
              }
            });
          })
        });
        return {
          ...prev,
          [key]: value,
        }
      }, {});

    console.log('******')
    console.log(JSON.stringify(worksheets))

    reporter.logger.debug('Starting child request to render docx dynamic parts', req)
/*
    const { content: newContent } = await reporter.render({
      template: {
        content: contentToRender,
        engine: req.template.engine,
        recipe: 'html',
        helpers: req.template.helpers
      }
    }, req)

    context.$xlsxReport[stringsPath] = JSON.parse(newContent)
*/
    reporter.logger.debug('Parsing xlsx report content finished', req)
    return context;
  } catch (e) {
    throw reporter.createError('Error while executing xlsx report recipe', {
      original: e,
      weak: true
    })
  }
}
module.exports = (reporter) => async (context, req) => {
  try {
    reporter.logger.debug('Parsing xlsx report content begin', req)
    
    const stringsPath = Object.keys(context.$xlsxReport).find(k => k.endsWith('sharedStrings.xml'));
    if (!stringsPath) {
      throw new Error('Not found strings file');
    }
    const strings = context.$xlsxReport[stringsPath];
    const contentToRender = JSON.stringify(strings);

    reporter.logger.debug('Starting child request to render docx dynamic parts', req)

    const { content: newContent } = await reporter.render({
      template: {
        content: contentToRender,
        engine: req.template.engine,
        recipe: 'html',
        helpers: req.template.helpers
      }
    }, req)

    context.$xlsxReport[stringsPath] = JSON.parse(newContent)

    reporter.logger.debug('Parsing xlsx report content finished', req)
    return context;
  } catch (e) {
    throw reporter.createError('Error while executing xlsx report recipe', {
      original: e,
      weak: true
    })
  }
}
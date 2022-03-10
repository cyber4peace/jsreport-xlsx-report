const fs = require('fs').promises
const path = require('path')

module.exports = (reporter, definition) => {
  reporter.options.sandbox.modules.push({
    alias: 'fsproxy.js',
    path: path.join(__dirname, '../lib/fsproxy.js')
  })

  reporter.options.sandbox.modules.push({
    alias: 'lodash',
    path: require.resolve('lodash')
  })

  reporter.options.sandbox.modules.push({
    alias: 'xml2js-preserve-spaces',
    path: require.resolve('xml2js-preserve-spaces')
  })

  if (reporter.options.sandbox.allowedModules !== '*') {
    reporter.options.sandbox.allowedModules.push('path')
  }

  reporter.extensionsManager.recipes.push({
    name: 'xlsx-report',
    execute: (req, res) => require('./recipe')(reporter, definition, req, res)
  })

  reporter.beforeRenderListeners.insert({ after: 'data' }, definition.name, async (req) => {
    if (req.template.recipe !== 'xlsx-report') {
      return
    }

    const serialize = require('./serialize.js')
    const parse = serialize.parse

    const findTemplate = async () => {
      if (
        (!req.template.xlsxReport || (!req.template.xlsxReport.shortid && !req.template.xlsxReport.templateAsset))
      ) {
        return fs.readFile(path.join(__dirname, '../static/defaultXlsxReport.json')).then((content) => JSON.parse(content))
      }

      if (req.template.xlsxReport && req.template.xlsxReport.templateAsset && req.template.xlsxReport.templateAsset.content) {
        return parse(Buffer.from(req.template.xlsxReport.templateAsset.content, req.template.xlsxReport.templateAsset.encoding || 'utf8'))
      }

      let docs = []
      let xlsxReportShortid

      if (req.template.xlsxReport && req.template.xlsxReport.shortid) {
        xlsxReportShortid = req.template.xlsxReport.shortid
        docs = await reporter.documentStore.collection('assets').find({ shortid: xlsxReportShortid }, req)
      }

      if (!docs.length) {
        if (!xlsxReportShortid) {
          throw reporter.createError('Unable to find xlsx template. xlsx template not specified', {
            statusCode: 404
          })
        }

        throw reporter.createError(`Unable to find xlsx template with shortid ${xlsxReportShortid}`, {
          statusCode: 404
        })
      }

      return parse(docs[0].content)
    }

    const template = await findTemplate()

    req.data = req.data || {}
    req.data.$xlsxReport = template
    req.data.$xlsxModuleDirname = path.join(__dirname, '../')
    req.data.$tempAutoCleanupDirectory = reporter.options.tempAutoCleanupDirectory
    req.data.$addBufferSize = definition.options.addBufferSize || 50000000
    req.data.$escapeAmp = definition.options.escapeAmp
    req.data.$numberOfParsedAddIterations = definition.options.numberOfParsedAddIterations == null ? 50 : definition.options.numberOfParsedAddIterations
  })
}

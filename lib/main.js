const extend = require('node.extend.without.arrays')
const serialize = require('./serialize')

module.exports = (reporter, definition) => {
  definition.options = extend(true, { preview: {} }, reporter.options.xlsxReport, reporter.options.office, definition.options)
  reporter.options.xlsxReport = definition.options

  if (definition.options.previewInExcelOnline != null) {
    definition.options.preview.enabled = definition.options.previewInExcelOnline
  }

  if (definition.options.showExcelOnlineWarning != null) {
    definition.options.preview.showWarning = definition.options.showExcelOnlineWarning
  }

  if (definition.options.publicUriForPreview != null) {
    definition.options.preview.publicUri = definition.options.publicUriForPreview
  }

  reporter.extensionsManager.recipes.push({
    name: 'xlsx-report',
  })

  reporter.documentStore.registerEntityType('XlsxReportType', {
    name: { type: 'Edm.String' },
    contentRaw: { type: 'Edm.Binary', document: { extension: 'xlsx' } },
    content: { type: 'Edm.String', document: { extension: 'txt' } }
  })

  // NOTE: xlsxTemplates are deprecated, we will remove it in jsreport v4
  reporter.documentStore.registerEntitySet('xlsxReports', {
    entityType: 'jsreport.XlsxReportType',
    splitIntoDirectories: true,
    // since it is deprecated we don't want that imports process xlsxTemplates
    exportable: false
  })

  reporter.documentStore.registerComplexType('XlsxReportRefType', {
    templateAssetShortid: { type: 'Edm.String', referenceTo: 'assets', schema: { type: 'null' } }
  })

  if (reporter.documentStore.model.entityTypes.TemplateType) {
    reporter.documentStore.model.entityTypes.TemplateType.xlsxReport = { type: 'jsreport.XlsxReportRefType', schema: { type: 'null' } }
  }

  reporter.documentStore.on('after-init', () => {
    reporter.documentStore.collection('xlsxReports').beforeInsertListeners.add('xlsxReports', (doc) => {
      return serialize(doc.contentRaw).then((serialized) => (doc.content = serialized))
    })

    reporter.documentStore.collection('xlsxReports').beforeUpdateListeners.add('xlsxReports', (query, update, req) => {
      if (update.$set && update.$set.contentRaw) {
        return serialize(update.$set.contentRaw).then((serialized) => (update.$set.content = serialized))
      }
    })
  })

  reporter.initializeListeners.add('xlsxReport', () => {
    if (reporter.express) {
      reporter.express.exposeOptionsToApi(definition.name, {
        preview: {
          enabled: definition.options.preview.enabled,
          showWarning: definition.options.preview.showWarning
        }
      })
    }
  })
}

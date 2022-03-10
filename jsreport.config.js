const office = require('@jsreport/office')

const schema = {
  type: 'object',
  properties: {
    previewInExcelOnline: { type: 'boolean' },
    publicUriForPreview: { type: 'string' },
    escapeAmp: { type: 'boolean' },
    addBufferSize: { type: 'number' },
    numberOfParsedAddIterations: { type: 'number' },
    showExcelOnlineWarning: { type: 'boolean', default: true }
  }
}

module.exports = {
  name: 'xlsx-report',
  main: 'lib/main.js',
  worker: 'lib/worker.js',
  optionsSchema: office.extendSchema('xlsx-report', {
    xlsxReport: { ...schema },
    extensions: {
      xlsxReport: { ...schema }
    }
  }),
  dependencies: ['data'],
  requires: {
    core: '3.x.x',
    studio: '3.x.x'
  }
}

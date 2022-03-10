const fs = require('fs')
const stream = require('stream')
const zlib = require('zlib')
const merge = require('merge2')

const fallback = require('./fallback.js')
const jsonToXml = require('./jsonToXml.js')
const fsproxy = require('./fsproxy.js')
const processXlsx = require('./processXlsx')

const { response, serializeOfficeXmls } = require('@jsreport/office')

const stringToStream = (str) => {
  const s = new stream.Readable()
  s.push(str)
  s.push(null)
  return s
}

const worksheetOrder = {
  sheetPr: -2,
  dimension: -1,
  sheetViews: 1,
  sheetFormatPr: 2,
  cols: 3,
  sheetData: 4,
  sheetCalcPr: 5,
  sheetProtection: 6,
  protectedRanges: 7,
  scenarios: 8,
  autoFilter: 9,
  sortState: 10,
  dataConsolidate: 11,
  customSheetViews: 12,
  mergeCells: 13,
  phoneticPr: 14,
  conditionalFormatting: 15,
  dataValidations: 16,
  hyperlinks: 17,
  printOptions: 18,
  pageMargins: 19,
  pageSetup: 20,
  headerFooter: 21,
  rowBreaks: 22,
  colBreaks: 23,
  customProperties: 24,
  cellWatches: 25,
  ignoredErrors: 26,
  smartTags: 27,
  drawing: 28,
  legacyDrawing: 29,
  legacyDrawingHF: 30,
  legacyDrawingHeaderFooter: 31,
  drawingHeaderFooter: 32,
  picture: 33,
  oleObjects: 34,
  controls: 35,
  webPublishItems: 36,
  tableParts: 37,
  extLst: 38
}

function print (root) {
  ensureWorksheetOrder(root.$xlsxReport)
  bufferedFlush(root)
  return {
    $xlsxReport: root.$xlsxReport,
    $files: root.$files || []
  }
}

function ensureWorksheetOrder (data) {
  // eslint-disable-next-line no-unused-vars
  for (const key in data) {
    if (key.indexOf('xl/worksheets/') !== 0) {
      continue
    }

    if (!data[key] || !data[key].worksheet) {
      continue
    }

    const worksheet = data[key].worksheet
    const sortedWorksheet = {}
    Object.keys(worksheet).sort(function (a, b) {
      if (!worksheetOrder[a]) return -1 // undefined in worksheetOrder goes at top of list
      if (!worksheetOrder[b]) return 1
      if (worksheetOrder[a] === worksheetOrder[b]) return 0
      return worksheetOrder[a] > worksheetOrder[b] ? 1 : -1
    }).forEach(function (a) {
      sortedWorksheet[a] = worksheet[a]
    })
    data[key].worksheet = sortedWorksheet
  }
}

function bufferedFlush (root) {
  const buffers = root.$buffers || {}

  Object.keys(buffers).forEach(function (f) {
    Object.keys(buffers[f]).forEach(function (k) {
      if (buffers[f][k].data.length) {
        flush(buffers[f][k], root)
      }
    })
  })
}

function flush (buf, root) {
  root.$files.push(fsproxy.write(root.$tempAutoCleanupDirectory, buf.data))
  buf.collection.push({ $$: root.$files.length - 1 })
  buf.data = ''
}

module.exports = async (reporter, definition, req, res) => {
  let $xlsxReport
  let $files

  try {
    const context = await processXlsx(reporter)(req.data, req)
    const content = print(context)
    $xlsxReport = content.$xlsxReport
    $files = content.$files
  } catch (e) {
    reporter.logger.warn('Ошибка создания отчета' + '\n' + e.stack, req)
    return fallback(reporter, definition, req, res)
  }

  const files = Object.keys($xlsxReport).map((k) => {
    if (k.includes('xl/media/') || k.includes('.bin')) {
      return {
        path: k,
        data: Buffer.from($xlsxReport[k], 'base64')
      }
    }

    if (k.includes('.xml')) {
      const xmlAndFiles = jsonToXml($xlsxReport[k])
      const fullXml = xmlAndFiles.xml

      if (fullXml.indexOf('&&') < 0) {
        return {
          path: k,
          data: Buffer.from(fullXml, 'utf8')
        }
      }

      const xmlStream = merge()

      if (fullXml.indexOf('&&') < 0) {
        return {
          path: k,
          data: Buffer.from(fullXml, 'utf8')
        }
      }

      let xml = fullXml

      while (xml) {
        const separatorIndex = xml.indexOf('&&')

        if (separatorIndex < 0) {
          xmlStream.add(stringToStream(xml))
          xml = ''
          continue
        }

        xmlStream.add(stringToStream(xml.substring(0, separatorIndex)))
        xmlStream.add(fs.createReadStream($files[xmlAndFiles.files.shift()]).pipe(zlib.createInflate()))
        xml = xml.substring(separatorIndex + '&&'.length)
      }

      return {
        path: k,
        data: xmlStream
      }
    }

    return {
      path: k,
      data: Buffer.from($xlsxReport[k], 'utf8')
    }
  })

  await serializeOfficeXmls({ reporter, files, officeDocumentType: 'xlsx' }, req, res)
  return response({
    previewOptions: definition.options.preview,
    officeDocumentType: 'xlsx',
    stream: res.stream,
    logger: reporter.logger
  }, req, res)
}

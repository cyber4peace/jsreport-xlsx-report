import XlsxReportProperties from './XlsxReportProperties.js'
import Studio from 'jsreport-studio'

Studio.addPropertiesComponent(XlsxReportProperties.title, XlsxReportProperties, (entity) => entity.__entitySet === 'templates' && entity.recipe === 'xlsx-report')

Studio.entityEditorComponentKeyResolvers.push((entity) => {
  if (entity.__entitySet === 'templates' && entity.recipe === 'xlsx-report') {
    let officeAsset

    if (entity.xlsxReport != null && entity.xlsxReport.shortid != null) {
      officeAsset = Studio.getEntityByShortid(entity.xlsxReport.shortid, false)
    }

    return {
      key: 'assets',
      entity: officeAsset,
      props: {
        icon: 'fa-link',
        embeddingCode: '',
        helpersEntity: entity,
        displayName: `xlsx report asset: ${officeAsset != null ? officeAsset.name : '<none>'}`,
        emptyMessage: 'No xlsx report asset assigned, please add a reference to a xlsx report asset in the properties'
      }
    }
  }
})

Studio.runListeners.push((request, entities) => {
  if (request.template.recipe !== 'xlsx-report') {
    return
  }

  if (Studio.extensions["xlsx-report"].options.preview.enabled === false) {
    return
  }

  if (Studio.extensions["xlsx-report"].options.preview.showWarning === false) {
    return
  }

  if (Studio.getSettingValueByKey('office-preview-informed', false) === true) {
    return
  }

  Studio.setSetting('office-preview-informed', true)

  Studio.openModal(() => (
    <div>
      Ваш отчет будет загружен на общедоступный сервер, чтобы была возможность использовать онлайн Excel для предварительного просмотра.
    </div>
  ))
})

import React, { Component } from 'react'
import Studio from 'jsreport-studio'

const EntityRefSelect = Studio.EntityRefSelect
const sharedComponents = Studio.sharedComponents

class XlsxReportProperties extends Component {
  static selectAssets (entities) {
    return Object.keys(entities).filter((k) => entities[k].__entitySet === 'assets').map((k) => entities[k])
  }

  static title (entity, entities) {
    if (
      (!entity.xlsxReport || !entity.xlsxReport.shortid)
    ) {
      return 'xlsx report'
    }

    const foundAssets = XlsxReportProperties.selectAssets(entities).filter((e) => entity.xlsxReport != null && entity.xlsxReport.shortid === e.shortid)

    if (!foundAssets.length) {
      return 'xlsx report'
    }

    const name = foundAssets[0].name
    return 'xlsx report: ' + name
  }

  componentDidMount () {
    this.removeInvalidXlsxReportReferences()
  }

  componentDidUpdate () {
    this.removeInvalidXlsxReportReferences()
  }

  removeInvalidXlsxReportReferences () {
    const { entity, entities, onChange } = this.props

    if (!entity.xlsxReport) {
      return
    }

    const updatedXlsxReportAssets = Object.keys(entities)
      .filter((k) => entities[k].__entitySet === 'assets' && entity.xlsxReport != null && entities[k].shortid === entity.xlsxReport.shortid)

    if (entity.xlsxReport && entity.xlsxReport.shortid && updatedXlsxReportAssets.length === 0) {
      onChange({ _id: entity._id, xlsxReport: null })
    }
  }

  changeXlsxReport (oldXlsxReport, prop, value) {
    let newValue

    if (value == null) {
      newValue = { ...oldXlsxReport }
      newValue[prop] = null
    } else {
      return { ...oldXlsxReport, [prop]: value }
    }

    newValue = Object.keys(newValue).length ? newValue : null

    return newValue
  }

  render () {
    const { entity, onChange } = this.props

    return (
      <div className='properties-section'>
        <div className='form-group'>
          <label>xlsx report asset</label>
          <EntityRefSelect
            headingLabel='Select xlsx report'
            newLabel='New xlsx asset for report'
            value={entity.xlsxReport ? entity.xlsxReport.shortid : ''}
            onChange={(selected) => onChange({
              _id: entity._id,
              xlsxReport: selected != null && selected.length > 0 ? { shortid: selected[0].shortid } : null
            })}
            filter={(references) => ({ assets: references.assets })}
            renderNew={(modalProps) => <sharedComponents.NewAssetModal {...modalProps} options={{ ...modalProps.options, defaults: { folder: entity.folder }, activateNewTab: false }} />}
          />
        </div>
      </div>
    )
  }
}

export default XlsxReportProperties

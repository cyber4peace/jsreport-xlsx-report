function xlsxTable (context, options) {
  const Handlebars = require('handlebars')
  return Handlebars.helpers.each(context, options)
}
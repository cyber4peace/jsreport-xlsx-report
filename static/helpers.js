function xlsxTable (context, options) {
  const Handlebars = require('handlebars');

  const ret = [];
  let data;
  if (options.data) {
    data = Handlebars.createFrame(options.data);
  }
  if (data && Array.isArray(context)) {
    for (let i = 0; i < context.length; i++) {
      if (i in context) {
        data.key = i;
        data.index = i;
        data.first = i === 0;
        data.last = (i === (data.length - 1));
      }
      ret.push(options.fn(context[i], { data: data }));
    }
  }

  return new Handlebars.SafeString(ret.join('$$$xlsxRow$$$'));
}
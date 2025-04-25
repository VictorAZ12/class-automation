function doGet(request) {
    // initiate html service
    return HtmlService.createTemplateFromFile('Index')
        .evaluate();
  }
  function include(filename) {
    // include css and js into html template (printing scriptlets)
    return HtmlService.createHtmlOutputFromFile(filename)
        .getContent();
  }
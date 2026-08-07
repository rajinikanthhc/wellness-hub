function doGet() {
  return HtmlService.createTemplateFromFile("index")
    .evaluate()
    .setTitle("Wellness Hub")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

function include(file) {
  return HtmlService.createHtmlOutputFromFile(file).getContent();
}

function getYoga() {
  return getYogaList();
}
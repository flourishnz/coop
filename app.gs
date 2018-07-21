
// how to launch app 


function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .getContent();
}

function doGet() {
  var t = HtmlService.createTemplateFromFile('Index');
  t.data = getPureFresh()
  return t.evaluate();
}


//function doGet() {
//  return HtmlService.createTemplateFromFile('html')
//     .evaluate()
//     .setTitle("Tweaking assistant")
//     .setSandboxMode(HtmlService.SandboxMode.IFRAME);
//}
//
//
//

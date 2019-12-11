/**
 * @OnlyCurrentDoc
 */

function onOpen(e) {
  DocumentApp.getUi().createAddonMenu().addItem('Open', 'showSidebar').addToUi();
}

function onInstall(e) {
  setUserSettingsToDefault();
  onOpen(e);
}

function showSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('sidebar').setTitle('Sentence Structure Visualizer');
  DocumentApp.getUi().showSidebar(ui);
}

function binarySearchImplementation(boundaries,n) {
  for(;boundaries.length!=1;){
    midpoint = Math.floor(boundaries.length/2);
    if(boundaries[midpoint] == n) {
      return n;
    } else if (boundaries[midpoint] > n) {
      boundaries = boundaries.slice(0,midpoint);
    } else {
      boundaries = boundaries.slice(midpoint,boundaries.length);
    };
  };
  return boundaries[0];
}

function getUserSettingsFromStorage() {
  var userProperties = PropertiesService.getUserProperties();
  Logger.log({
    "ColorBoundaries" : JSON.parse(userProperties.getProperty("ColorBoundaries")),
    "ColorCodes"      : JSON.parse(userProperties.getProperty("ColorCodes")),
    "Unit"            : JSON.parse(userProperties.getProperty("Unit"))
  });
  return {
    "ColorBoundaries" : JSON.parse(userProperties.getProperty("ColorBoundaries")),
    "ColorCodes"      : JSON.parse(userProperties.getProperty("ColorCodes")),
    "Unit"            : JSON.parse(userProperties.getProperty("Unit"))
  };
}

function setUserSettings(settings) {
  Logger.log(settings);
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperties(settings);
}

function getColorBySentence(sentence, settings, body) {
  var n =0;
  if(settings["Unit"]=="Words") {
    var words = sentence.match(/\b([\w'’]+)\b/g);
    if(!(words == null)) {
      n = words.length;
    }
  } else {
    n = sentence.length;
  };
  var index = settings["ColorBoundaries"].indexOf(binarySearchImplementation(settings["ColorBoundaries"], n));
  return settings["ColorCodes"][index];
}

function highlightText(settings) {
  const body = DocumentApp.getActiveDocument().getBody();
  var paragraphs = body.getParagraphs();
  for(i=0;i<paragraphs.length;i++) {
    const unicodeRanges = "\u0020\u0022-\u002D\u002F-\u003E\u0040-\u007E\u00A0-\u03FF\u1E00-\u2027\u202F-\u2046\u204A-\u205E\u2070-\u20BF\u20D0-\u20F0\u2100-\u2BFF\u2C60-\u2C7F\uA722-\uA7FF"
    const regexp = new RegExp('[' + unicodeRanges + ']+[.?!\v]', 'g');
    var sentences = paragraphs[i].getText().match(regexp);

    if(sentences == null) {
      const newRe = new RegExp(".*?[0-9a-zA-Z]+.*");
      sentences = paragraphs[i].getText().match(newRe);
    };
    if(!(sentences == null)){
      var prevOffset = 0;  
      for(j=0;j<sentences.length;j++){
        if(!(sentences[j] == null) && sentences[j].length) {
          var colorCode = getColorBySentence(sentences[j], settings, body);
          if(prevOffset){paragraphs[i].editAsText().setBackgroundColor(prevOffset+1,prevOffset+sentences[j].length-1,colorCode);}
          else{paragraphs[i].editAsText().setBackgroundColor(prevOffset,prevOffset+sentences[j].length-1,colorCode);}
          prevOffset=prevOffset+sentences[j].length;
        };
      };
    };
  };
}

function removeHighlight() {
  Logger.log("Remove Highlight Called");
  var text = DocumentApp.getActiveDocument().getBody().getParagraphs();
  for(var i=0;i<text.length;i++) {
    text[i].editAsText().setBackgroundColor(null);
  }
  Logger.log("Remove Highlight Completed");
}

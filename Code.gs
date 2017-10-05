function onInstall(e) {
  onOpen(e);
}

function onOpen() {
  DocumentApp.getUi()
    .createMenu('AutoVocab')
    .addItem('Create', 'createVocab')
    .addToUi();
}

function createVocab() {
  
  var ui = DocumentApp.getUi();
  
  var num = getStartNum();
  
  do {
    var result = ui.prompt('AutoVocab', 'Enter your word:',
                           ui.ButtonSet.OK_CANCEL);
    
    var button = result.getSelectedButton();
    var text = result.getResponseText();
    
    if(button == ui.Button.OK) {
      createWordEntry(text, num);
      
    }
    
    num++;
  }while(button == ui.Button.OK);
}

function getStartNum() {
  var ui = DocumentApp.getUi();

  var result = ui.prompt('AutoVocab', 'Enter starting number:', ui.ButtonSet.OK);
  return result.getResponseText();
}
  

function createWordEntry(word, num) {
  word = fixWord(word);
  
  var json = getJSON(word);
  
  if(json === false) {
    error();
    return;
  }
  
  var definition = getDefinition(json);
  var partOfSpeech = getPartOfSpeech(json);
  
  addWord(word, definition, partOfSpeech, num);
}

function getJSON(word) {
  var data = {
    "Accept": "application/json",
    "app_id": "631930eb",
    "app_key": "837087a1c825b937864d29449337645a"
  }
  var options = {
    'method': 'GET',
    'headers': data,
    'followRedirects': true,
  }
  
  try {
    var response = UrlFetchApp.fetch('https://od-api.oxforddictionaries.com:443/api/v1/entries/en/'+word, options).getContentText();
  
    var responseJSON = JSON.parse(response)
    return responseJSON;
  } catch (e) {
    error();
    return false;
  }
}

function error() {
  var ui = DocumentApp.getUi();
  
  ui.alert('Error', 'Uh oh, something went wrong!', ui.ButtonSet.OK);
}

function fixWord(word) {
  //Checks to see if word is changed in dictionary (Lacerated to Lacerate)
  try {
    var response = UrlFetchApp.fetch('https://en.oxforddictionaries.com/definition/'+word).getContentText();
    var regex = new RegExp("<meta content=\"Definition of (.+?) -");
    word = regex.exec(response)[1];
  } catch (e) {
    
  }
  
  word = word.charAt(0).toUpperCase() + word.slice(1);
  
  return word;
}

function getDefinition(json) {
  var definition;
  try {
    definition = json['results'][0]['lexicalEntries'][0]['entries'][0]['senses'][0]['definitions'][0];
  } catch(e) {
    error();
    return "Error!";
  }
  
  definition = definition.charAt(0).toUpperCase() + definition.slice(1);
  
  if(definition.charAt(definition.length-1) != '.') {
    definition += '.';
  }
  
  return definition;
}

function getPartOfSpeech(json) {
  try {
  var longPOS = json['results'][0]['lexicalEntries'][0]['lexicalCategory'];
  } catch (e) {
    error();
    return "Error!";
  }
  var shortPOS;
  
  switch(longPOS) {
    case 'Noun':
      shortPOS = 'N';
      break;
    case 'Adjective':
      shortPOS = 'ADJ';
      break;
    case 'Verb':
      shortPOS = 'V';
      break;
    case 'Adverb':
      shortPOS = 'ADV';
      break;
    default:
      shortPOS = longPOS;
      break;
  }
  
  return shortPOS;
}

function addWord(word, definition, partOfSpeech, num) {
  var body = DocumentApp.getActiveDocument().getBody();
  
  var text = body.editAsText();
  text.appendText(num + '. (' + partOfSpeech + ') ' + word + ' – ' + definition + '\n');
  
  var style = {};
  style[DocumentApp.Attribute.BOLD] = true;
  
  var wordLocation = text.findText(word);
  wordLocation.getElement().setAttributes(wordLocation.getStartOffset(), wordLocation.getEndOffsetInclusive(), style);
}

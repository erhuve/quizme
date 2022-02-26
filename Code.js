const DEFAULT_INPUT_TEXT = '';
const DEFAULT_OUTPUT_TEXT = '';
const NUM_LETTER_MAP = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']

/**
 * Callback for rendering the main card.
 * @return {CardService.Card} The card to show the user.
 */
function onHomepage(e) {
  return createSelectionCard(e, DEFAULT_INPUT_TEXT, DEFAULT_OUTPUT_TEXT);
}

/**
 * Main function to generate the main card.
 * @param {String} inputText The text used for generating questions.
 * @return {CardService.Card} The card to show to the user.
 */
function createSelectionCard(e, outputText) {
  var hostApp = e['hostApp'];
  var builder = CardService.newCardBuilder();
  
  var userInputSection = CardService.newCardSection();
  //   .addWidget(CardService.newTextInput()
  //     .setFieldName('input')
  //     .setValue(inputText)
  //     .setTitle('Enter text to generate questions from...')
  //     .setMultiline(true));

  // userInputSection.addWidget(CardService.newButtonSet()
  //   .addButton(CardService.newTextButton()
  //     .setText('Copy Selected Text')
  //     .setOnClickAction(CardService.newAction().setFunctionName('getDocsSelection'))
  //     .setDisabled(false))
  //   .addButton(CardService.newTextButton()
  //     .setText('Clear')
  //     .setOnClickAction(CardService.newAction().setFunctionName('clearText'))
  //     .setDisabled(false)));

  userInputSection.addWidget(CardService.newDecoratedText()
    .setText("Show answers")
    .setWrapText(true)
    .setSwitchControl(CardService.newSwitch()
      .setFieldName("shouldShowAnswers")
      .setValue("true")));

  userInputSection.addWidget(generateNumberDropdown('numberOfQuestions', 5).setTitle("Max number of questions"));


  builder.addSection(userInputSection);
  builder.addSection(CardService.newCardSection()
    .addWidget(CardService.newTextParagraph()
    .setText("Highlight select text in the document, then click the button!")));
  
  builder.addSection(CardService.newCardSection()
    .addWidget(CardService.newButtonSet()
      .addButton(CardService.newTextButton()
        .setText('Quiz me!')
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setOnClickAction(CardService.newAction().setFunctionName('generateQuestions'))
        .setDisabled(false))));

  builder.addSection(CardService.newCardSection()
    .addWidget(CardService.newTextInput()
      .setFieldName('output')
      .setValue(outputText)
      .setTitle('Questions...')
      .setMultiline(true)
    )
  );

  return builder.build();

}

/**
 * Helper function to clear the text.
 * @return {CardService.Card} The card to show to the user.
 */
function clearText(e) {
  return createSelectionCard(e, DEFAULT_INPUT_TEXT);
}

/**
 * Helper function to append to the document
 * @return void, simply appends to the document
 */
function generateQuestions(e) {
  // var inputText = e.formInput.input ?? '';
  var inputText = getDocsSelection();
  Logger.log(inputText);
  var numberOfQuestions = e.formInput.numberOfQuestions ?? 10;
  var shouldShowAnswers = e.formInput.shouldShowAnswers ?? 'false';
  var shouldGenerateNewDoc = e.formInput.shouldGenerateNewDoc ?? 'false'; // TODO: generate new doc w/ questions
  // var fontColor = e.formInput.fontColor ?? 'red'; // TODO: write questions in fontColor

  // var query = "joke"
  // var url = "https://dad-jokes.p.rapidapi.com/random/" + query; // TODO: Call http://104.198.92.212/ instead
  // var options = {
  //   headers: {
  //     'x-rapidapi-host': 'dad-jokes.p.rapidapi.com',
  //     'x-rapidapi-key': '2aad18d4aamshe379f00afc91688p10252ajsn2974d3aeb21d'
  //   }
  // }
  var url = 'http://104.198.92.212';
  var data = {
    'text': inputText,
    'num_questions': numberOfQuestions,
    'answer_style': 'multiple_choice'
  };
  var options = {
    'method' : 'post',
    'contentType': 'application/json',
    'muteHttpExceptions': true,
    // Convert the JavaScript object to a JSON string.
    'payload' : JSON.stringify(data)
  };

  try {
    Logger.log("I am trying to fetch now");
    var response = UrlFetchApp.fetch(url, options);
    Logger.log("Response received");
    var json = response.getContentText();
    var data = JSON.parse(json);
  
    var qaPairs = data.qa_pairs;
    var questions = qaPairs.map(qa => qa.question);
    var answers = qaPairs.map(qa => qa.answer);
    if (inputText !== undefined) {
      return createSelectionCard(e, inputText, generateQuestionsText(questions, answers));
    }
  }
  catch {
    return createSelectionCard(e, inputText, 'Error generating questions. Try with less text or fewer questions.');
  }
  


  // var setup = data.body[0].setup
  // var punchline = data.body[0].punchline

  // if (inputText !== undefined) {
  //   var outputText = inputText + '\n' + setup + '\n' + punchline + '\n' + numberOfQuestions + '\n' + shouldShowAnswers;
  //   return createSelectionCard(e, inputText, outputText);
  // }
}

function generateQuestionsText(questions, answers) {
  Logger.log("In Generate q text");
  var outputText = '';
  for (var i = 0; i < questions.length; i++) {
    outputText += `${i+1}. ` + questions[i] + '\n';
    if (getObjType(answers[i]) === 'array') {
      // MC Question
      for (var j = 0; j < answers[i].length; j++) {
        outputText += NUM_LETTER_MAP[j] + '. ' + answers[i][j].answer + '\n';
      }
    }
    outputText += '\n';
  }
  Logger.log(outputText);
  return outputText;
}

// Helper function for us to check if it is an MC or open question
// https://stackoverflow.com/questions/20664528/how-to-get-the-object-type
function getObjType(obj) {
  
  // If the object is an array, this will return the stored value,
  // if the object is a string, this will return only one letter of the string
  var type = obj[0];
  if (type.length == 1) {
    return 'string';
  }
  try {
    type = obj.foobar();
  } catch (error) {
    
    // TypeError no longer contains object type, just return 'array'
    Logger.log(error);
    return 'array';
  }
}

/**
 * Helper function to generate the color menu.
 * @param {String} fieldName
 * @param {String} previousSelected The color the user previously had selected.
 * @return {CardService.SelectionInput} The card to show to the user.
 */
function generateColorDropdown(fieldName, previousSelected) {
  var selectionInput = CardService.newSelectionInput().setTitle("Font for questions")
    .setFieldName(fieldName)
    .setType(CardService.SelectionInputType.DROPDOWN);

  COLOR_MAP.forEach((color, index, array) => {
    selectionInput.addItem(color.text, color.val, color.val == previousSelected);
  })

  return selectionInput;
}

/**
 * Helper function to generate the # of questions menu.
 * @param {String} fieldName
 * @param {String} previousSelected The # of questions the user previously had selected.
 * @return {CardService.SelectionInput} The card to show to the user.
 */
function generateNumberDropdown(fieldName, previousSelected) {
  var selectionInput = CardService.newSelectionInput().setTitle("Number of questions")
    .setFieldName(fieldName)
    .setType(CardService.SelectionInputType.DROPDOWN);

  for (i = 1; i <= 10; i = i + 1) { //starts loop
    selectionInput.addItem(`${i}`, i, i == previousSelected);
  };

  return selectionInput;
}

/**
 * Helper function to get the text selected.
 * @return {CardService.Card} The selected text.
 */
function getDocsSelection() {
  var text = '';
  var selection = DocumentApp.getActiveDocument().getSelection();
  Logger.log(selection)
  if (selection) {
    var elements = selection.getRangeElements();
    for (var i = 0; i < elements.length; i++) {
      Logger.log(elements[i]);
      var element = elements[i];
      // Only modify elements that can be edited as text; skip images and other non-text elements.
      if (element.getElement().asText() && element.getElement().asText().getText() !== '') {
        text += element.getElement().asText().getText() + '\n';
      }
    }
  }

  if (text !== '') {
    return text;
  }
}

/**
 * Helper function to get the text of the selected cells.
 * @return {CardService.Card} The selected text.
 */
function getSheetsSelection(e) {
  var text = '';
  var ranges = SpreadsheetApp.getActive().getSelection().getActiveRangeList().getRanges();
  for (var i = 0; i < ranges.length; i++) {
    const range = ranges[i];
    const numRows = range.getNumRows();
    const numCols = range.getNumColumns();
    for (let i = 1; i <= numCols; i++) {
      for (let j = 1; j <= numRows; j++) {
        const cell = range.getCell(j, i);
        if (cell.getValue()) {
          text += cell.getValue() + '\n';
        }
      }
    }
  }
  if (text !== '') {
    var transformation = "Not really a transformation, just the same text: " + text; 
    return createSelectionCard(e, text, transformation);
  }
}

/**
 * Helper function to get the selected text of the active slide.
 * @return {CardService.Card} The selected text.
 */
function getSlidesSelection(e) {
  var text = '';
  var selection = SlidesApp.getActivePresentation().getSelection();
  var selectionType = selection.getSelectionType();
  if (selectionType === SlidesApp.SelectionType.TEXT) {
    var textRange = selection.getTextRange();
    if (textRange.asString() !== '') {
      text += textRange.asString() + '\n';
    }
  }
  if (text !== '') {
    var transformation = "Not really a transformation, just the same text: " + text; 
    return createSelectionCard(e, text, transformation);
  }
}

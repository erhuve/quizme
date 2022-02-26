const DEFAULT_INPUT_TEXT = '';

const COLOR_MAP =
  [
    { text: 'Red', val: 'red' },
    { text: 'Orange', val: 'orange' },
    { text: 'Yellow', val: 'yellow' },
    { text: 'Green', val: 'green' },
    { text: 'Blue', val: 'blue' },
    { text: 'Purple', val: 'purple' },
    { text: 'Pink', val: 'pink' }
  ]

/**
 * Callback for rendering the main card.
 * @return {CardService.Card} The card to show the user.
 */
function onHomepage(e) {
  return createSelectionCard(e, DEFAULT_INPUT_TEXT);
}

/**
 * Main function to generate the main card.
 * @param {String} inputText The text used for generating questions.
 * @return {CardService.Card} The card to show to the user.
 */
function createSelectionCard(e, inputText) {
  var hostApp = e['hostApp'];
  var builder = CardService.newCardBuilder();
  
  var userInputSection = CardService.newCardSection()
    .addWidget(CardService.newTextInput()
      .setFieldName('input')
      .setValue(inputText)
      .setTitle('Enter text to generate questions from...')
      .setMultiline(true));

  userInputSection.addWidget(CardService.newButtonSet()
    .addButton(CardService.newTextButton()
      .setText('Copy Selected Text')
      .setOnClickAction(CardService.newAction().setFunctionName('getDocsSelection'))
      .setDisabled(false))
    .addButton(CardService.newTextButton()
      .setText('Clear')
      .setOnClickAction(CardService.newAction().setFunctionName('clearText'))
      .setDisabled(false)));

  userInputSection.addWidget(CardService.newDecoratedText()
    .setText("Show answers")
    .setWrapText(true)
    .setSwitchControl(CardService.newSwitch()
      .setFieldName("shouldShowAnswers")
      .setValue("true")));

  userInputSection.addWidget(
    CardService.newDecoratedText()
      .setText("Generate new Google Doc with questions")
      .setWrapText(true)
      .setSwitchControl(CardService.newSwitch()
        .setFieldName("shouldGenerateNewDoc")
        .setValue("true")));

  userInputSection.addWidget(generateNumberDropdown('numberOfQuestions', 10).setTitle("Number of questions"));
  userInputSection.addWidget(generateColorDropdown('fontColor', "red").setTitle("Font color for questions"));


  builder.addSection(userInputSection);
  
  builder.addSection(CardService.newCardSection()
    .addWidget(CardService.newButtonSet()
      .addButton(CardService.newTextButton()
        .setText('Add Text & Call an API')
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setOnClickAction(CardService.newAction().setFunctionName('appendToDocWithApiCall'))
        .setDisabled(false))));

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
function appendToDocWithApiCall(e) {
  var inputText = e.formInput.input ?? '';
  var numberOfQuestions = e.formInput.numberOfQuestions ?? 10;
  var shouldShowAnswers = e.formInput.shouldShowAnswers ?? 'false';
  var shouldGenerateNewDoc = e.formInput.shouldGenerateNewDoc ?? 'false'; // TODO: generate new doc w/ questions
  var fontColor = e.formInput.fontColor ?? 'red'; // TODO: write questions in fontColor

  var query = "joke"
  var url = "https://dad-jokes.p.rapidapi.com/random/" + query; // TODO: Call http://104.198.92.212/ instead
  var options = {
    headers: {
      'x-rapidapi-host': 'dad-jokes.p.rapidapi.com',
      'x-rapidapi-key': '2aad18d4aamshe379f00afc91688p10252ajsn2974d3aeb21d'
    }
  }
  
  // var data = {
  //   'inputText': inputText,
  //   'numberOfQuestions': numberOfQuestions,
  //   'shouldShowAnswers': shouldShowAnswers
  // }
  // var options = {
  //   'method' : 'post',
  //   'contentType': 'application/json',
  //   // Convert the JavaScript object to a JSON string.
  //   'payload' : JSON.stringify(data)
  // };

  var response = UrlFetchApp.fetch(url, options)
  var json = response.getContentText();
  var data = JSON.parse(json);
  var setup = data.body[0].setup
  var punchline = data.body[0].punchline

  if (inputText !== undefined) {
    var body = DocumentApp.getActiveDocument().getBody();
    body.appendParagraph('{newText}');
    body.appendParagraph('{setup}');
    body.appendParagraph('{punchline}');
    body.replaceText('{newText}', inputText);
    body.replaceText('{setup}', setup);
    body.replaceText('{punchline}', punchline);

    // test parameters are set correctly
    body.appendParagraph('{numberOfQuestions}');
    body.appendParagraph('{shouldGenerateNewDoc}');
    body.appendParagraph('{shouldShowAnswers}');
    body.appendParagraph('{fontColor}');
    body.replaceText('{numberOfQuestions}', numberOfQuestions);
    body.replaceText('{shouldGenerateNewDoc}', shouldGenerateNewDoc);
    body.replaceText('{shouldShowAnswers}', shouldShowAnswers);
    body.replaceText('{fontColor}', fontColor);
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

  for (i = 5; i < 25; i = i + 5) { //starts loop
    selectionInput.addItem(`${i}`, i, i == previousSelected);
  };

  return selectionInput;
}

/**
 * Helper function to get the text selected.
 * @return {CardService.Card} The selected text.
 */
function getDocsSelection(e) {
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
    var transformation = "Not really a transformation, just the same text: " + text; 
    return createSelectionCard(e, text, transformation);
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

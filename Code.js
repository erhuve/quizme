const DEFAULT_TEXT = '';
const NUM_LETTER_MAP = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']

/**
 * Callback for rendering the main card.
 * @return {CardService.Card} The card to show the user.
 */
function onHomepage(e) {
  return createSelectionCard(e);
}

/**
 * Main function to generate the main card.
 * @param {String} inputText The text used for generating questions.
 * @return {CardService.Card} The card to show to the user.
 */
function createSelectionCard(e, questionsText=DEFAULT_TEXT, answersText=DEFAULT_TEXT) {
  var hostApp = e['hostApp'];
  var builder = CardService.newCardBuilder();
  
  if (questionsText != DEFAULT_TEXT) {
    builder.addSection(CardService.newCardSection()
      .setHeader("Questions")
      .addWidget(CardService.newTextParagraph()
        .setText(questionsText)));

    builder.addSection(CardService.newCardSection()
      .setCollapsible(true)
      .setHeader("Answers")
      .addWidget(CardService.newTextParagraph()
      .setText(answersText)));
  }

  else {
    var userInputSection = CardService.newCardSection();
    userInputSection.addWidget(CardService.newTextParagraph()
      .setText("Select text in the doc you want to generate questions for, and fill out the fields below."));
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

    userInputSection.addWidget(generateNumberDropdown('numberOfQuestions', 5).setTitle("Max number of questions"));
    
    userInputSection
      .addWidget(CardService.newDecoratedText()
        .setText("Create Google Form")
        .setWrapText(true)
        .setSwitchControl(CardService.newSwitch()
          .setFieldName("shouldCreateNewForm")
          .setValue("false")))
      .addWidget(CardService.newTextInput()
        .setFieldName('formTitle')
        .setTitle('Enter title for form.'));

    userInputSection.addWidget(CardService.newTextButton()
          .setText('Quiz me!')
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
          .setOnClickAction(CardService.newAction().setFunctionName('generateQuestions'))
          .setDisabled(false))
    
    builder.addSection(userInputSection);

    builder.build()
  }
  return builder.build();
}

/**
 * Helper function to clear the text.
 * @return {CardService.Card} The card to show to the user.
 */
function clearText(e) {
  return createSelectionCard(e);
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
      if (e.formInput.shouldCreateNewForm) { 
        createNewForm(e, questions, answers)
      }
      return createSelectionCard(e, generateQuestionsText(questions, answers), generateAnswersText(answers));
    }
  }
  catch {
    return createSelectionCard(e, 'Error generating questions. Try with less text or fewer questions.');
  }
  
}

/**
 * Helper function to create a new form with the generated qa pairs.
 * @param {String} title
 * @param {[String]} generated questions
 * @param {[String | [Object]]} answers to questions
 * @return 
 */
function createNewForm(e, questions, answers) {
  var form = FormApp.create(e.formInput.formTitle ?? "New Form")
  .setProgressBar(true)
  .setPublishingSummary(true)
  .setIsQuiz(true)
  form.setDescription('Auto-generated quiz by QuizMe')

  for (var i = 0; i < questions.length; i++) {
    if (getObjType(answers[i]) === 'array') { // MC Question
      var item = form.addMultipleChoiceItem().setPoints(1)
      var choices = answers[i].map(ans => 
      item.createChoice(ans.answer, ans.correct)
      )
      item
        .setTitle(questions[i])  
        .setChoices(choices)
        .setRequired(true); 
    }
    else { // Sentences Question
      var item = form.addParagraphTextItem();
      var feedback = FormApp.createFeedback().setText(answers[i])
      item.setGeneralFeedback(feedback.build());
    }
  }
}

function generateQuestionsText(questions, answers) {
  Logger.log("In generateQuestionsText");
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

function generateAnswersText(answers) {
  Logger.log("In generateAnswersText");
  var outputText = '';
  for (var i = 0; i < answers.length; i++) {
    outputText += `${i+1}. `;
    if (getObjType(answers[i]) === 'array') {
      // MC Question
      correctIndex = answers[i].findIndex(ans => ans.correct)
      outputText += NUM_LETTER_MAP[correctIndex] + '. ' + answers[i][correctIndex].answer + '\n';
    }
  }
  Logger.log(outputText);
  return outputText;
}

/**
 * Helper function for us to check if it is an MC or open question
 * @param {Object} obj
 * https://stackoverflow.com/questions/20664528/how-to-get-the-object-type
*/
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
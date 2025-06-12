
//  Replace 'YOUR_OPENAI_API_KEY' with your actual OpenAI API key






/**
 * Adds a custom menu to Google Sheets when opened.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('AI Assistant')
    .addItem('GPT', 'callGPT')
    .addToUi();
}

/**
 * callGPT
 * -------
 * This function orchestrates the process of generating a spreadsheet formula using ChatGPT.
 */
function callGPT() {
  var ui = SpreadsheetApp.getUi();
 
  // 1) Prompt for the formula description
  var formulaPrompt = ui.prompt(
    'GPT',
    'What formula do you want me to generate?',
    ui.ButtonSet.OK_CANCEL
  );
 
  if (formulaPrompt.getSelectedButton() !== ui.Button.OK) {
    ui.alert('Action canceled.');
    return;
  }

  var myprompt = formulaPrompt.getResponseText();
  // Call GPT to generate the formula based on the user's prompt.
  var formula = GPT(myprompt, 450);
 
  // Split the output based on '+' to check if it's a graph command.
  var parts = formula.split('+');  // Example output: ["graph", "A1:B10"]

  var cellPrompt;
  // If the output indicates a graph, ask the user for the graph type.
  if (parts[0] === "graph"){
    cellPrompt = ui.prompt(
      'Graph Type',
      'Enter the graph type (BAR, LINE, PIE, COLUMN):',
      ui.ButtonSet.OK_CANCEL
    );
  } else {
    // Otherwise, prompt the user for a destination cell in A1 notation.
    cellPrompt = ui.prompt(
      'Set Destination Cell',
      'Enter the cell location in A1 notation (e.g., B2):',
      ui.ButtonSet.OK_CANCEL
    );
  }
 
  if (cellPrompt.getSelectedButton() !== ui.Button.OK) {
    ui.alert('Action canceled.');
    return;
  }
  var cellLocation = cellPrompt.getResponseText().toUpperCase();
 
  if (parts[0] === "graph"){
    // Create a chart using the cell range and graph type provided.
    createChart(parts[1], cellLocation);
  } else {
    // Insert the user prompt (as a descriptive formula) into the specified cell.
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var formulaText = '="' + myprompt + '"';
    sheet.getRange(cellLocation).setFormula(formulaText);

    // Insert the generated formula into the cell directly below the target cell.
    var currentRange = sheet.getRange(cellLocation);
    var nextRange = currentRange.offset(1, 0);
    var belowCellRef = nextRange.getA1Notation();
    sheet.getRange(belowCellRef).setFormula(formula);
  }
}

/**
 * GPT
 * ---
 * This function sends the user’s prompt to the OpenAI API and returns the generated output.
 */
function GPT(prompt, maxTokens = 450) {
  if (!prompt) {
    return "Error: Please provide a prompt.";
  }
  const apiUrl = 'https://api.openai.com/v1/chat/completions';
  const payload = {
    'model': 'gpt-4o-mini',
    'messages': [
      {
        'role': 'system',
        'content': 'You are an expert in using Google Sheets. Provide ONLY the correct spreadsheet formula for the question asked. ' +
        'For example, if asked to calculate the sum of numbers between cells A3 and A7, output: =SUM(A3:A7). Do not include quotes. ' +
        '** If the query is about graphing, the output should start with "graph" followed by the cell range(s), e.g., graph+A1:B10. ** ' +
        'If multiple separate ranges are given for a graph, format as graph+A2:A21,C2:C21. ** If values from different ranges require operations, use ARRAYFORMULA.'
      },
      {'role': 'user', 'content': prompt}
    ],
    'max_tokens': maxTokens
  };
  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'headers': {
      'Authorization': 'Bearer ' + OPENAI_API_KEY
    },
    'payload': JSON.stringify(payload)
  };
  try {
    const response = UrlFetchApp.fetch(apiUrl, options);
    const json = JSON.parse(response.getContentText());
    var formula = json.choices[0].message.content.trim();
    return formula;
  } catch (error) {
    return "Error: " + error.toString();
  }
}

/**
 * createChart
 * -----------
 * This function creates a chart on the active spreadsheet.
 */
function createChart(cellrange, chartTypeString) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Convert user input to lowercase
  chartTypeString = chartTypeString.toLowerCase();
  
  // Check for a valid chart type
  var chartTypeEnum = getChartType(chartTypeString);
  if (!chartTypeEnum) {
    SpreadsheetApp.getUi().alert("❌ Error: Unknown chart type - " + chartTypeString);
    return;
  }
  
  try {
    var rangeArray = cellrange.split(",");
    var rangeList = sheet.getRangeList(rangeArray);
    
    if (rangeList.getRanges().length === 0) {
      SpreadsheetApp.getUi().alert("⚠️ Error: The specified range does not exist.");
      return;
    }
    
    var chartBuilder = sheet.newChart()
      .setChartType(chartTypeEnum)
      .setOption('title', 'Generated Chart') // Adds a title
      .setPosition(5, 5, 0, 0);
    
    // Add data ranges to chart
    rangeList.getRanges().forEach(function (range) {
      chartBuilder.addRange(range);
    });

    var chart = chartBuilder.build();
    sheet.insertChart(chart);

    SpreadsheetApp.getUi().alert('✅ Chart successfully created: ' + chartTypeString);
  } catch (error) {
    Logger.log("❌ Chart creation error: " + error);
    SpreadsheetApp.getUi().alert("❌ Error: Failed to create the chart. Please check the range and chart type.");
  }
}

/**
 * getChartType
 * ------------
 * Maps user-friendly chart type names to Google Sheets' ChartType enums.
 */
function getChartType(chartType) {
  var types = {
    'column': Charts.ChartType.COLUMN,
    'line': Charts.ChartType.LINE,
    'pie': Charts.ChartType.PIE,
    'bar': Charts.ChartType.BAR
  };
  return types[chartType] || null;
}


/**
 * Shows the sidebar UI in Google Sheets.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('AI Summarizer')
    .addItem('Open Sidebar', 'showSidebar')
    .addToUi();
}

/**
 * Displays the sidebar with the configuration form.
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('sidebar')
      .setTitle('AI Summarizer');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Processes the spreadsheet data based on sidebar configuration.
 * @param {Object} config Configuration object from the sidebar.
 * @return {Object} Result object with status and message.
 */
function processRows(config) {
  try {
    if (!config.prompt || !config.prompt.includes('{{value}}')) {
      throw new Error('Prompt is required and must include {{value}}');
    }
    if (!config.resultColumn && !config.resultCell) {
      throw new Error('Result column or cell is required');
    }
    if (!config.headerRow || config.headerRow < 0) {
      config.headerRow = 0; 
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    let values = [];
    let startRow, endRow, inputColIndex, resultColIndex, resultCell;

    // Handle input type
    if (config.inputType === 'cell') {
      const cellMatch = config.inputCell.match(/^([A-Z]+)([0-9]+)$/);
      if (!cellMatch) {
        throw new Error('Invalid cell reference (e.g., use A2)');
      }
      const col = columnToIndex(cellMatch[1]);
      const row = parseInt(cellMatch[2]);
      if (row <= config.headerRow || row > lastRow || col > lastCol) {
        throw new Error('Cell is out of bounds or in header rows');
      }
      const cellValue = sheet.getRange(row, col).getValue();
      if (typeof cellValue !== 'string' || cellValue.trim() === '') {
        throw new Error('Cell is empty or not text');
      }
      values = [[cellValue]];
      startRow = row;
      inputColIndex = col;

      if (config.resultCell) {
        const resultMatch = config.resultCell.match(/^([A-Z]+)([0-9]+)$/);
        if (!resultMatch) {
          throw new Error('Invalid result cell reference (e.g., use B2)');
        }
        resultCell = config.resultCell;
      } else {
        resultColIndex = columnToIndex(config.resultColumn.toUpperCase());
      }
    } else if (config.inputType === 'column') {
      // Validate input column
      inputColIndex = columnToIndex(config.inputColumn.toUpperCase());
      if (inputColIndex < 1 || inputColIndex > lastCol) {
        throw new Error('Invalid input column');
      }
      startRow = config.headerRow + 1;
      endRow = lastRow;
      const range = sheet.getRange(startRow, inputColIndex, endRow - startRow + 1, 1);
      values = range.getValues();
      resultColIndex = columnToIndex(config.resultColumn.toUpperCase());
    } else if (config.inputType === 'range') {
      // Validate input column
      inputColIndex = columnToIndex(config.inputColumn.toUpperCase());
      if (inputColIndex < 1 || inputColIndex > lastCol) {
        throw new Error('Invalid input column');
      }
      if (config.rowSelection.type === 'auto') {
        startRow = config.headerRow + 1;
        endRow = Math.min(startRow + (config.rowSelection.numRows || 3) - 1, lastRow);
      } else if (config.rowSelection.type === 'fixed') {
        startRow = Math.max(config.headerRow + 1, config.rowSelection.startRow || 1);
        endRow = Math.min(config.rowSelection.endRow || lastRow, lastRow);
      } else {
        throw new Error('Invalid row selection type');
      }
      if (startRow > lastRow || endRow < startRow) {
        throw new Error('No valid rows to process');
      }
      const range = sheet.getRange(startRow, inputColIndex, endRow - startRow + 1, 1);
      values = range.getValues();
      resultColIndex = columnToIndex(config.resultColumn.toUpperCase());
    } else {
      throw new Error('Invalid input type');
    }

    if (resultColIndex && (resultColIndex < 1 || resultColIndex > lastCol)) {
      throw new Error('Invalid result column');
    }

    // Process each value
    const results = [];
    for (let i = 0; i < values.length; i++) {
      const text = values[i][0];
      if (typeof text === 'string' && text.trim() !== '') {
        try {
          const summary = GPT_SUMMARIZE(
            text,
            config.prompt,
            config.temperature || 0.7,
            config.model || 'gemini-2.0-flash'
          );
          results.push([summary]);
        } catch (error) {
          results.push(['Error: ' + error.message]);
        }
      } else {
        results.push(['']);
      }
    }

    if (config.inputType === 'cell' && resultCell) {
      sheet.getRange(resultCell).setValue(results[0][0]);
    } else {
      sheet.getRange(startRow, resultColIndex, results.length, 1).setValues(results);
    }

    return { status: 'success', message: `Processed ${results.length} ${config.inputType === 'cell' ? 'cell' : 'rows'}` };
  } catch (error) {
    return { status: 'error', message: 'Processing failed: ' + error.message };
  }
}

/**
 * Converts a column letter to a 1-based index (e.g., 'A' -> 1, 'B' -> 2).
 * @param {string} column Column letter (e.g., 'B').
 * @return {number} 1-based column index.
 */
function columnToIndex(column) {
  let index = 0;
  for (let i = 0; i < column.length; i++) {
    index = index * 26 + (column.charCodeAt(i) - 64);
  }
  return index;
}

/**
 * Summarizes text using the Gemini API.
 * @param {string} text The text to summarize or a cell reference.
 * @param {string} format The prompt format for summarization (e.g., "Summarize {{value}}").
 * @param {number} temperature The temperature for the model (0 to 1).
 * @param {string} model The Gemini model to use.
 * @return {string} The summarized text.
 * @customfunction
 */
function GPT_SUMMARIZE(text, format, temperature, model) {
  // Validate inputs with specific error messages
  if (typeof text !== 'string' || text.trim() === '') {
    throw new Error('Text is required and must not be empty');
  }
  if (typeof format !== 'string' || format.trim() === '') {
    throw new Error('Format is required and must not be empty');
  }
  if (typeof temperature !== 'number' || temperature < 0 || temperature > 1) {
    throw new Error('Temperature must be a number between 0 and 1');
  }
  if (!model) {
    model = 'gemini-2.0-flash'; 
  }

  const prompt = format.replace('{{value}}', text);

  try {
    return callGeminiAPI(prompt, temperature, model);
  } catch (error) {
    throw new Error('Gemini API error: ' + error.message);
  }
}

/**
 * Calls the Gemini API to process a prompt.
 * @param {string} prompt The prompt to send to the API.
 * @param {number} temperature The temperature for the model.
 * @param {string} model The Gemini model to use.
 * @return {string} The API response text.
 */
function callGeminiAPI(prompt, temperature, model) {
  const apiKey = ''; // insert Gemini API key
  if (!apiKey) {
    throw new Error('Gemini API key not set. Run setApiKey() to configure.');
  }

  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;
  const options = {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify({
      contents: [{
        parts: [{
          text: prompt
        }]
      }],
      generationConfig: {
        temperature: temperature
      }
    })
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());
    if (data.candidates && data.candidates[0].content) {
      return data.candidates[0].content.parts[0].text;
    } else {
      throw new Error('No valid response from Gemini API');
    }
  } catch (error) {
    throw new Error('API request failed: ' + error.message);
  }
}

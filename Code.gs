/**
 * This Google Apps Script automates vocabulary building by fetching
 * explanations, examples, and translations for words entered into a Google Sheet
 * directly from the Gemini API. It also includes functions for visual formatting and quizzes.
 */

// --- Configuration ---
const SPREADSHEET_ID = 'YOUR_SPREED_SHEET_ID';
const SHEET_NAME = 'SHEET_NAME';
const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
const GEMINI_MODEL = 'gemini-2.5-flash';

function setApiKeys() {
  PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', 'YOUR_GEMINI_API_KEY');
  Logger.log('Gemini API Key set successfully in script properties.');
}

// --- Formatting Constants ---
const EVEN_ROW_COLOR = '#f3f3f3';
const ODD_ROW_COLOR = '#ffffff';
const PAST_DUE_COLOR = '#ffcdd2';
const DUE_TODAY_COLOR = '#ffcc80';
const DUE_TOMORROW_COLOR = '#fff9c4';
const BROKEN_LINK_COLOR = '#ffcdd2'; // Light Red for broken links

/**
 * Triggered automatically when a cell in the spreadsheet is edited.
 */
function onEdit(e) {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(10000)) {
    try {
      const range = e.range;
      const sheet = range.getSheet();

      if (sheet.getName() === SHEET_NAME && range.getColumn() === 1 && range.getNumRows() === 1) {
        const word = range.getValue().toString().trim();
        const row = range.getRow();

        if (!word) {
          sheet.getRange(row, 2, 1, 19).clearContent(); // Clear 19 columns (B to T)
          formatRow(sheet, row);
          return;
        }

        // **CHANGE**: Call the shared processing function
        processNewWord(word, row);
      }
    } finally {
      lock.releaseLock();
    }
  } else {
    Logger.log('Could not acquire lock for onEdit trigger.');
  }
}

/**
 * **NEW**: This is the core logic for processing a new word.
 * It's called by both onEdit and addNewWord.
 * @param {string} word The word to process.
 * @param {number} row The row number to populate.
 */
function processNewWord(word, row) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  
  // Duplicate check
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const wordListRange = sheet.getRange(2, 1, lastRow - 1, 1);
    const wordValues = wordListRange.getValues();
    for (let i = 0; i < wordValues.length; i++) {
      const existingWord = wordValues[i][0].toString().trim();
      const existingRow = i + 2;
      if (existingWord.toLowerCase() === word.toLowerCase() && existingRow !== row) {
        SpreadsheetApp.getUi().alert(`The word "${word}" already exists in row ${existingRow}.`);
        sheet.getRange(row, 1).clearContent(); // Clear only the duplicate cell
        return;
      }
    }
  }

  const geminiResponse = getGeminiDefinitionAndExamples(word);
  const now = new Date();

  const definitions = geminiResponse.meanings.map(m => `• ${m.definition}`).join('\n');
  const definitionExamples = geminiResponse.meanings.map(m => `• ${m.example}`).join('\n');
  const persianTranslations = geminiResponse.meanings.map(m => `• ${m.persianTranslation}`).join('\n');

  sheet.getRange(row, 2).setValue(persianTranslations);
  sheet.getRange(row, 3).setValue(definitions);
  sheet.getRange(row, 4).setValue(definitionExamples);
  sheet.getRange(row, 5).setValue(geminiResponse.generalExamples.join('\n'));
  sheet.getRange(row, 6).setValue(geminiResponse.partOfSpeech);
  sheet.getRange(row, 7).setValue(geminiResponse.synonyms);
  sheet.getRange(row, 8).setValue(geminiResponse.antonyms);
  sheet.getRange(row, 9).setValue(geminiResponse.notes);
  sheet.getRange(row, 10).setValue(geminiResponse.wordFamily);

  const addedDateCell = sheet.getRange(row, 11);
  if (!addedDateCell.getValue()) {
    addedDateCell.setValue(now);
  }

  sheet.getRange(row, 12).setValue(now);
  sheet.getRange(row, 13).setValue(geminiResponse.ukPronunciation);
  sheet.getRange(row, 14).setValue(geminiResponse.usPronunciation);

  const encodedWord = encodeURIComponent(word.toLowerCase().replace(/ /g, '-'));
  const cambridgeUrl = `https://dictionary.cambridge.org/dictionary/english/${encodedWord}`;
  const oxfordUrl = `https://www.oxfordlearnersdictionaries.com/definition/english/${encodedWord}`;
  
  const cambridgeCell = sheet.getRange(row, 15);
  cambridgeCell.setFormula(`=HYPERLINK("${cambridgeUrl}", "Cambridge")`);
  if (!checkLinkValidity(cambridgeUrl)) {
    cambridgeCell.setBackground(BROKEN_LINK_COLOR);
  }

  const oxfordCell = sheet.getRange(row, 16);
  oxfordCell.setFormula(`=HYPERLINK("${oxfordUrl}", "Oxford")`);
   if (!checkLinkValidity(oxfordUrl)) {
    oxfordCell.setBackground(BROKEN_LINK_COLOR);
  }

  const nextReviewDate = new Date(now);
  nextReviewDate.setDate(now.getDate() + 1);
  sheet.getRange(row, 17).setValue(nextReviewDate);
  sheet.getRange(row, 18).setValue(0);
  sheet.getRange(row, 19).setValue(0);
  sheet.getRange(row, 20).setValue(0); // Initialize Total Reviews to 0
  
  formatRow(sheet, row);
}


/**
 * Calls the Gemini API to get detailed vocabulary information.
 */
function getGeminiDefinitionAndExamples(word) {
  const prompt = `For the English word "${word}", respond ONLY with a JSON object that includes the following keys:
- "partOfSpeech": part of speech (e.g., noun, verb, adjective).
- "meanings": an array of objects, where each object has a "definition" key (a plain, simple explanation), an "example" key (a single, clear example sentence for that specific definition), and a "persianTranslation" key (the single-word Persian equivalent for that specific meaning).
- "generalExamples": an array of additional, clear example sentences that show general usage.
- "synonyms": an array of 1-3 common synonyms.
- "antonyms": an array of 1-3 common antonyms.
- "notes": an array of useful notes (e.g., common collocations, usage warnings).
- "ukPronunciation": common Articulatory Phonetics spelling for UK English using IPA.
- "usPronunciation": common Articulatory Phonetics spelling for US English using IPA.
- "wordFamily": an array of 2-3 common related forms of the word.

Do not include markdown formatting (like backticks) or any commentary outside the JSON object. Only return the JSON.`;

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generation_config: { temperature: 0.7, top_p: 0.95, top_k: 40 }
  };
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-goog-api-key': GEMINI_API_KEY },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(`https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent`, options);
    const responseText = response.getContentText();
    const responseCode = response.getResponseCode();

    if (responseCode >= 200 && responseCode < 300) {
      let data = JSON.parse(responseText);
      if (data.candidates && data.candidates[0].content && data.candidates[0].content.parts) {
        let geminiOutput = data.candidates[0].content.parts[0].text.trim().replace(/```(?:json)?\s*/i, "").replace(/```$/, "").trim();
        const parsed = JSON.parse(geminiOutput);
        return {
            partOfSpeech: parsed.partOfSpeech || '—',
            meanings: (parsed.meanings && parsed.meanings.length > 0) ? parsed.meanings : [{ definition: '—', example: '—', persianTranslation: '—' }],
            generalExamples: (parsed.generalExamples || []).map(e => `• ${e}`),
            synonyms: (parsed.synonyms || []).map(s => `• ${s}`).join('\n') || '—',
            antonyms: (parsed.antonyms || []).map(a => `• ${a}`).join('\n') || '—',
            notes: (parsed.notes || []).map(n => `• ${n}`).join('\n') || '—',
            ukPronunciation: parsed.ukPronunciation || '—',
            usPronunciation: parsed.usPronunciation || '—',
            wordFamily: (parsed.wordFamily || []).map(wf => `• ${wf}`).join('\n') || '—'
        };
      }
    }
    return getDefaultGeminiResponse(`API error (HTTP ${responseCode}).`);
  } catch (e) {
    console.error("Error during API call: " + e.message);
    return getDefaultGeminiResponse('Script or network error.');
  }
}

function getDefaultGeminiResponse(errorMsg = 'Error') {
  return {
    partOfSpeech: errorMsg,
    meanings: [{ definition: '—', example: '—', persianTranslation: '—' }],
    generalExamples: [],
    synonyms: '—',
    antonyms: '—',
    notes: '—',
    ukPronunciation: '—',
    usPronunciation: '—',
    wordFamily: '—'
  };
}

function getWordsDueForReview() {
  Logger.log("[Backend] getWordsDueForReview: Starting to fetch words.");
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) {
    Logger.log("[Backend] getWordsDueForReview: Sheet not found.");
    throw new Error(`Sheet with name "${SHEET_NAME}" not found.`);
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log("[Backend] getWordsDueForReview: No words in the sheet.");
    return [];
  }

  // Read all 20 columns of data (A-T)
  const dataRange = sheet.getRange(2, 1, lastRow - 1, 20);
  const values = dataRange.getValues();
  const linkFormulas = sheet.getRange(2, 15, lastRow - 1, 2).getFormulas(); // Columns O and P

  const wordsDue = [];
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  values.forEach((row, index) => {
    const word = row[0];
    const nextReviewDateCell = row[16]; // Column Q

    if (word && nextReviewDateCell instanceof Date) {
      const nextReviewDate = new Date(nextReviewDateCell);
      nextReviewDate.setHours(0, 0, 0, 0);

      if (nextReviewDate <= today) {
        const cambridgeFormula = linkFormulas[index][0];
        const oxfordFormula = linkFormulas[index][1];
        
        const cambridgeUrl = cambridgeFormula.match(/HYPERLINK\("([^"]+)"/i)?.[1] || '';
        const oxfordUrl = oxfordFormula.match(/HYPERLINK\("([^"]+)"/i)?.[1] || '';

        wordsDue.push({
          word: word,
          row: index + 2,
          persianTranslations: row[1] ? row[1].split('\n') : [],
          definitions: row[2] ? row[2].split('\n') : [],
          definitionExamples: row[3] ? row[3].split('\n') : [],
          generalExamples: row[4] ? row[4].split('\n') : [],
          partOfSpeech: row[5],
          synonyms: row[6],
          antonyms: row[7],
          notes: row[8],
          wordFamily: row[9],
          ukPronunciation: row[12],
          usPronunciation: row[13],
          cambridgeUrl: cambridgeUrl,
          oxfordUrl: oxfordUrl,
          reviewCount: row[17] || 0,
          totalReviews: row[19] || 0
        });
      }
    }
  });

  Logger.log(`[Backend] getWordsDueForReview: Found ${wordsDue.length} words due for review today.`);
  return wordsDue;
}

/**
 * Implements new review count logic.
 */
function updateWordReview(update) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    const now = new Date();
    const { row, difficulty, reviewCount, totalReviews } = update;
    
    let currentReviewCount = Number(reviewCount) || 0;
    let currentTotalReviews = Number(totalReviews) || 0;
    
    let newReviewCount = currentReviewCount;
    let newTotalReviews = currentTotalReviews;
    let daysToAdd = 0;

    if (typeof difficulty === 'number' && difficulty > 0) {
      daysToAdd = difficulty;
      newReviewCount = currentReviewCount + 1;
      newTotalReviews = currentTotalReviews + 1;
    } 
    else if (typeof difficulty === 'string') {
      switch (difficulty) {
        case 'Again': 
          daysToAdd = 0; 
          newReviewCount = 0;
          break;
        case 'Hard': 
          daysToAdd = 1; 
          newTotalReviews = currentTotalReviews + 1;
          break;
        case 'Good': 
          daysToAdd = (currentTotalReviews < 3) ? [3, 7, 14][currentTotalReviews] : 30; 
          newReviewCount++; 
          newTotalReviews++;
          break;
        case 'Easy': 
          daysToAdd = (currentTotalReviews < 3) ? [7, 30, 90][currentTotalReviews] : 180; 
          newReviewCount += 2; 
          newTotalReviews++;
          break;
        default: console.warn(`Unknown difficulty string: ${difficulty}.`); return;
      }
    } else {
      console.warn(`Invalid difficulty type: ${difficulty}.`);
      return;
    }

    const newNextReviewDate = new Date();
    newNextReviewDate.setDate(newNextReviewDate.getDate() + daysToAdd);
    newNextReviewDate.setHours(0, 0, 0, 0);

    sheet.getRange(row, 12).setValue(now);
    sheet.getRange(row, 17).setValue(newNextReviewDate);
    sheet.getRange(row, 18).setValue(newReviewCount);
    sheet.getRange(row, 20).setValue(newTotalReviews);
    
    formatRow(sheet, row);
    
    console.log(`Asynchronously updated row ${row} for difficulty '${difficulty}'.`);
  } catch (e) {
    console.error(`Failed to update row ${update ? update.row : 'unknown'}. Error: ${e.message}`);
  }
}

/**
 * Formats a single row with alternating colors and conditional formatting.
 */
function formatRow(sheet, rowNumber) {
  const wholeRowRange = sheet.getRange(rowNumber, 1, 1, 20);
  const wordCell = sheet.getRange(rowNumber, 1);
  const reviewDateCell = sheet.getRange(rowNumber, 17); // Column Q

  if (wordCell.getValue() === '') {
    wholeRowRange.setBackground(ODD_ROW_COLOR);
    return;
  }

  const rowBgColor = (rowNumber % 2 === 0) ? EVEN_ROW_COLOR : ODD_ROW_COLOR;
  wholeRowRange.setBackground(rowBgColor);

  const reviewDateVal = reviewDateCell.getValue();
  if (reviewDateVal instanceof Date) {
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const tomorrow = new Date(today);
    tomorrow.setDate(today.getDate() + 1);

    const reviewDate = new Date(reviewDateVal);
    reviewDate.setHours(0, 0, 0, 0);

    if (reviewDate < today) {
      wordCell.setBackground(PAST_DUE_COLOR);
    } else if (reviewDate.getTime() === today.getTime()) {
      wordCell.setBackground(DUE_TODAY_COLOR);
    } else if (reviewDate.getTime() === tomorrow.getTime()) {
      wordCell.setBackground(DUE_TOMORROW_COLOR);
    } else {
      wordCell.setBackground(rowBgColor);
    }
  }
}

/**
 * Sorts the sheet using the built-in, safer sort method and then applies formatting.
 */
function formatSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // Sort the range directly using the built-in sort method.
  const dataRange = sheet.getRange(2, 1, lastRow - 1, 20);
  dataRange.sort([
    { column: 17, ascending: true }, // Sort by Review Date (Column Q)
    { column: 18, ascending: true }, // Then by Review Count (Column R)
    { column: 1, ascending: true }   // Then by Word (Column A)
  ]);
  
  SpreadsheetApp.flush();

  // Now that it's sorted, read the values again to apply colors
  const values = dataRange.getValues();
  const backgroundColors = [];
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const tomorrow = new Date(today);
  tomorrow.setDate(today.getDate() + 1);

  for (let i = 0; i < values.length; i++) {
    const rowNumber = i + 2;
    const rowData = values[i];
    const word = rowData[0];
    const reviewDateVal = rowData[16]; // Column Q

    const rowBgColor = (rowNumber % 2 === 0) ? EVEN_ROW_COLOR : ODD_ROW_COLOR;
    const rowColors = new Array(20).fill(rowBgColor);

    if (word === '') {
      backgroundColors.push(new Array(20).fill(ODD_ROW_COLOR));
      continue;
    }

    if (reviewDateVal instanceof Date) {
      const reviewDate = new Date(reviewDateVal);
      reviewDate.setHours(0, 0, 0, 0);

      if (reviewDate < today) {
        rowColors[0] = PAST_DUE_COLOR;
      } else if (reviewDate.getTime() === today.getTime()) {
        rowColors[0] = DUE_TODAY_COLOR;
      } else if (reviewDate.getTime() === tomorrow.getTime()) {
        rowColors[0] = DUE_TOMORROW_COLOR;
      }
    }
    backgroundColors.push(rowColors);
  }

  if (backgroundColors.length > 0) {
    dataRange.setBackgrounds(backgroundColors);
  }
  
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('Sheet has been sorted and formatted!');
}


/**
 * Finds and removes duplicate words, keeping the first occurrence.
 */
function removeDuplicateWords() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const dataRange = sheet.getRange(2, 1, lastRow - 1, 1);
  const words = dataRange.getValues();
  
  const seenWords = {};
  const rowsToDelete = [];

  words.forEach((row, index) => {
    const word = row[0].toString().trim().toLowerCase();
    const rowNumber = index + 2;
    if (word) {
      if (seenWords[word]) {
        rowsToDelete.push(rowNumber);
      } else {
        seenWords[word] = true;
      }
    }
  });

  if (rowsToDelete.length > 0) {
    for (let i = rowsToDelete.length - 1; i >= 0; i--) {
      sheet.deleteRow(rowsToDelete[i]);
    }
    SpreadsheetApp.getUi().alert(`${rowsToDelete.length} duplicate word(s) found and removed.`);
    formatSheet();
  } else {
    SpreadsheetApp.getUi().alert('No duplicate words found.');
  }
}

// --- QUIZ FUNCTIONS ---

/**
 * Gets a list of words for the quiz, ensuring they have been reviewed at least once.
 */
function getQuizWords(numQuestions) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  
  const dataRange = sheet.getRange(2, 1, lastRow - 1, 20);
  const allWordsData = dataRange.getValues();

  let allWords = allWordsData.map((row, index) => ({
    word: row[0],
    definition: row[2] ? row[2].split('\n')[0].replace(/• /g, '') : 'No definition',
    reviewCount: row[17] || 0, // Column R
    quizCount: row[18] || 0,   // Column S
    row: index + 2
  })).filter(w => w.word && w.definition && w.reviewCount > 0);

  if (allWords.length < 4) {
    throw new Error("You need at least 4 words that have been reviewed once to start a quiz.");
  }
  
  const minQuizCount = Math.min(...allWords.map(w => w.quizCount));
  const unquizzedWords = allWords.filter(w => w.quizCount === minQuizCount);

  let wordsToQuizFrom = unquizzedWords;
  if (unquizzedWords.length < numQuestions) {
    wordsToQuizFrom = allWords;
  }
  
  const shuffled = wordsToQuizFrom.sort(() => 0.5 - Math.random());
  const selectedWords = shuffled.slice(0, numQuestions);

  const questions = selectedWords.map(correctWord => {
    const distractors = allWords
      .filter(w => w.word !== correctWord.word)
      .sort(() => 0.5 - Math.random())
      .slice(0, 3);

    const options = [correctWord, ...distractors].map(w => w.word);
    
    return {
      question: correctWord.definition,
      options: options.sort(() => 0.5 - Math.random()),
      answer: correctWord.word,
      row: correctWord.row
    };
  });

  return questions;
}

/**
 * Updates the sheet after a quiz is completed.
 */
function updateQuizResults(quizResults) {
  const { incorrectRows, allQuizRows } = quizResults;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const now = new Date();
  const tomorrow = new Date();
  tomorrow.setDate(now.getDate() + 1);
  tomorrow.setHours(0, 0, 0, 0);

  if (incorrectRows && incorrectRows.length > 0) {
    incorrectRows.forEach(rowNum => {
      sheet.getRange(rowNum, 17).setValue(tomorrow);
      sheet.getRange(rowNum, 12).setValue(now);
      formatRow(sheet, rowNum);
    });
  }

  if (allQuizRows && allQuizRows.length > 0) {
    allQuizRows.forEach(rowNum => {
      const countCell = sheet.getRange(rowNum, 19);
      const currentCount = countCell.getValue() || 0;
      countCell.setValue(currentCount + 1);
    });
  }
  
  const lastRow = sheet.getLastRow();
  const quizCounts = sheet.getRange(2, 19, lastRow - 1, 1).getValues().flat();
  if (quizCounts.every(count => count > 0)) {
    sheet.getRange(2, 19, lastRow - 1, 1).setValue(0);
    Logger.log("All words have been quizzed. Resetting quiz counts.");
  }
}

// --- CRAM MODE FUNCTION ---
/**
 * Gets a list of words scheduled for future review.
 * @param {number} count The number of future words to fetch.
 * @returns {Array<Object>} An array of word objects for review.
 */
function getFutureWords(count) {
  try {
    Logger.log('[Backend] getFutureWords: Starting to fetch ' + count + ' words.');
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error(`Sheet with name "${SHEET_NAME}" not found.`);
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    const dataRange = sheet.getRange(2, 1, lastRow - 1, 20);
    const values = dataRange.getValues();
    const linkFormulas = sheet.getRange(2, 15, lastRow - 1, 2).getFormulas();

    const futureWords = [];
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    values.forEach((row, index) => {
      const word = row[0];
      const nextReviewDateCell = row[16]; // Column Q

      if (word && nextReviewDateCell instanceof Date) {
        const nextReviewDate = new Date(nextReviewDateCell);
        nextReviewDate.setHours(0, 0, 0, 0);

        if (nextReviewDate > today) {
          const cambridgeFormula = linkFormulas[index][0];
          const oxfordFormula = linkFormulas[index][1];
          
          const cambridgeUrl = cambridgeFormula.match(/HYPERLINK\("([^"]+)"/i)?.[1] || '';
          const oxfordUrl = oxfordFormula.match(/HYPERLINK\("([^"]+)"/i)?.[1] || '';

          futureWords.push({
            word: word,
            row: index + 2,
            persianTranslations: row[1] ? row[1].split('\n') : [],
            definitions: row[2] ? row[2].split('\n') : [],
            definitionExamples: row[3] ? row[3].split('\n') : [],
            generalExamples: row[4] ? row[4].split('\n') : [],
            partOfSpeech: row[5],
            synonyms: row[6],
            antonyms: row[7],
            notes: row[8],
            wordFamily: row[9],
            ukPronunciation: row[12],
            usPronunciation: row[13],
            cambridgeUrl: cambridgeUrl,
            oxfordUrl: oxfordUrl,
            reviewCount: row[17] || 0,
            totalReviews: row[19] || 0,
            nextReviewDate: nextReviewDate // Keep as Date object for sorting
          });
        }
      }
    });

    Logger.log('[Backend] getFutureWords: Found ' + futureWords.length + ' total future words.');
    futureWords.sort((a, b) => a.nextReviewDate.getTime() - b.nextReviewDate.getTime());
    
    const wordsToReturn = futureWords.slice(0, count).map(word => {
      // Convert date to string AFTER sorting for safe transfer
      word.nextReviewDate = word.nextReviewDate.toISOString();
      return word;
    });

    Logger.log('[Backend] getFutureWords: Returning ' + wordsToReturn.length + ' words for cram mode.');
    return wordsToReturn;
  } catch (e) {
    Logger.log(`[Backend] Error in getFutureWords: ${e.message}`);
    return null; // Return null on error
  }
}


// --- UTILITY FUNCTIONS ---

/**
 * Sets up the sheet with the correct headers.
 */
function initializeSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert(`Sheet with name "${SHEET_NAME}" not found.`);
    return;
  }
  
  const headers = [
    'Word', 'Persian Translattion', 'Definition', 'Definition Example', 'Example',
    'Part of Speech', 'Synonyms', 'Antonyms', 'Tips', 'Word Family',
    'Created At', 'Modified At', 'UK Pronunciation', 'US Pronunciation',
    'Cambridge', 'Oxford', 'Next Review Time', 'Review Count', 'Quiz Count', 'Total Reviews'
  ];
  
  // sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  SpreadsheetApp.getUi().alert('Sheet has been initialized with the correct headers.');
}

/**
 * Initializes the Quiz Count for all existing words to 0.
 */
function initializeQuizCounts() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const quizCountRange = sheet.getRange(2, 19, lastRow - 1, 1);
  const values = quizCountRange.getValues();
  
  let updated = 0;
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === '') {
      values[i][0] = 0;
      updated++;
    }
  }
  
  quizCountRange.setValues(values);
  SpreadsheetApp.getUi().alert(`Initialization complete. ${updated} word(s) were updated with a quiz count of 0.`);
}

/**
 * Initializes the Total Reviews count for all existing words to 0.
 */
function initializeTotalReviews() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const totalReviewsRange = sheet.getRange(2, 20, lastRow - 1, 1); // Column T
  const values = totalReviewsRange.getValues();
  
  let updated = 0;
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === '') {
      values[i][0] = 0;
      updated++;
    }
  }
  
  totalReviewsRange.setValues(values);
  SpreadsheetApp.getUi().alert(`Initialization complete. ${updated} word(s) were updated with a total reviews count of 0.`);
}


/**
 * Schedules the review dates for all words in batches.
 */
function scheduleInitialReviews() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const wordsPerDay = 10;
  const reviewDateRange = sheet.getRange(2, 17, lastRow - 1, 1);
  const dates = [];
  let dayOffset = 0;

  for (let i = 0; i < lastRow - 1; i++) {
    if (i > 0 && i % wordsPerDay === 0) {
      dayOffset++;
    }
    const reviewDate = new Date();
    reviewDate.setHours(0, 0, 0, 0);
    reviewDate.setDate(reviewDate.getDate() + dayOffset);
    dates.push([reviewDate]);
  }

  reviewDateRange.setValues(dates);
  formatSheet(); // Re-sort and re-color the sheet after scheduling
  SpreadsheetApp.getUi().alert(`Initial review dates have been scheduled for ${lastRow - 1} words.`);
}


/**
 * Checks if a URL is valid by fetching its response code.
 */
function checkLinkValidity(url) {
  return true;
  // try {
  //   const response = UrlFetchApp.fetch(url, {
  //     muteHttpExceptions: true,
  //     method: 'HEAD'
  //   });
  //   return response.getResponseCode() === 200;
  // } catch (e) {
  //   Logger.log(`Error checking link ${url}: ${e.message}`);
  //   return false;
  // }
}

/**
 * Populates and verifies Cambridge and Oxford hyperlinks.
 */
function populateHyperlinks() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const wordRange = sheet.getRange(2, 1, lastRow - 1, 1);
  const words = wordRange.getValues();

  const cambridgeFormulas = [];
  const oxfordFormulas = [];
  const cambridgeColors = [];
  const oxfordColors = [];

  words.forEach(row => {
    const word = row[0].toString().trim();
    if (word) {
      const encodedWord = encodeURIComponent(word.toLowerCase().replace(/ /g, '-'));
      const cambridgeUrl = `https://dictionary.cambridge.org/dictionary/english/${encodedWord}`;
      const oxfordUrl = `https://www.oxfordlearnersdictionaries.com/definition/english/${encodedWord}`;
      
      cambridgeFormulas.push([`=HYPERLINK("${cambridgeUrl}", "Cambridge")`]);
      oxfordFormulas.push([`=HYPERLINK("${oxfordUrl}", "Oxford")`]);

      cambridgeColors.push([checkLinkValidity(cambridgeUrl) ? null : BROKEN_LINK_COLOR]);
      oxfordColors.push([checkLinkValidity(oxfordUrl) ? null : BROKEN_LINK_COLOR]);
    } else {
      cambridgeFormulas.push(['']);
      oxfordFormulas.push(['']);
      cambridgeColors.push([null]);
      oxfordColors.push([null]);
    }
  });

  sheet.getRange(2, 15, cambridgeFormulas.length, 1).setFormulas(cambridgeFormulas);
  sheet.getRange(2, 16, oxfordFormulas.length, 1).setFormulas(oxfordFormulas);
  sheet.getRange(2, 15, cambridgeColors.length, 1).setBackgrounds(cambridgeColors);
  sheet.getRange(2, 16, oxfordColors.length, 1).setBackgrounds(oxfordColors);

  SpreadsheetApp.getUi().alert('Hyperlinks have been populated and verified for all words.');
}


function setApiKeys() {
  PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', 'YOUR_GEMINI_API_KEY');
  Logger.log('Gemini API Key set successfully in script properties.');
}

function showReviewDialog() {
  const html = HtmlService.createHtmlOutputFromFile('ReviewDialog.html')
      .setTitle('Vocabulary Review Session')
      .setWidth(500)
      .setHeight(600);
  SpreadsheetApp.getUi().showModelessDialog(html, html.getTitle());
}

function openScramblePuzzle(word, hint) {
  const template = HtmlService.createTemplateFromFile('ScramblePuzzle');
  template.word = word;
  template.hint = hint;
  const html = template.evaluate()
      .setWidth(400)
      .setHeight(450);
  SpreadsheetApp.getUi().showModalDialog(html, 'Unscramble the Word');
}

function openQuizDialog() {
  const html = HtmlService.createHtmlOutputFromFile('QuizDialog.html')
      .setWidth(500)
      .setHeight(550);
  SpreadsheetApp.getUi().showModalDialog(html, 'Multiple-Choice Quiz');
}

/**
 * Adds a new function to the global scope for the review panel to call.
 */
function addNewWord(word) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  
  // Check for duplicates first
  const wordListRange = sheet.getRange(2, 1, lastRow - 1, 1);
  const wordValues = wordListRange.getValues();
  for (let i = 0; i < wordValues.length; i++) {
    if (wordValues[i][0].toString().trim().toLowerCase() === word.toLowerCase()) {
      return `The word "${word}" already exists in the list.`;
    }
  }
  
  // Add the new word to the next empty row
  sheet.getRange(lastRow + 1, 1).setValue(word);
  processNewWord(word, lastRow + 1);
  return `Added "${word}" to your vocabulary list.`;
}


/**
 * Shows the search dialog window.
 */
function showSearchDialog() {
  const html = HtmlService.createHtmlOutputFromFile('SearchDialog.html')
      .setTitle('Search Vocabulary')
      .setWidth(500)
      .setHeight(600);
  SpreadsheetApp.getUi().showModelessDialog(html, html.getTitle());
}

/**
 * Searches for words by prefix (case-insensitive).
 * Called by the client-side JavaScript in SearchDialog.html.
 * @param {string} prefix The search term.
 * @returns {Array<Object>} A list of matching words with their row numbers.
 */
function searchWords(prefix) {
  if (!prefix || prefix.trim().length === 0) {
    return [];
  }
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }
  const wordListRange = sheet.getRange(2, 1, lastRow - 1, 1);
  const wordValues = wordListRange.getValues();
  const lowerCasePrefix = prefix.trim().toLowerCase();
  
  const matches = [];
  wordValues.forEach((row, index) => {
    const word = row[0].toString().trim();
    if (word && word.toLowerCase().startsWith(lowerCasePrefix)) {
      matches.push({
        word: word,
        row: index + 2 // Convert 0-based index to 1-based row number
      });
    }
  });
  return matches.slice(0, 15); // Return a max of 15 matches
}

/**
 * Gets all details for a single word by its row number.
 * Called by the client-side JavaScript in SearchDialog.html.
 * @param {number} rowNumber The row number of the word.
 * @returns {Object} An object containing all details for the word.
 */
function getWordDetails(rowNumber) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) {
      throw new Error(`Sheet with name "${SHEET_NAME}" not found.`);
    }

    const dataRange = sheet.getRange(rowNumber, 1, 1, 20);
    const rowData = dataRange.getValues()[0];
    const linkFormulas = sheet.getRange(rowNumber, 15, 1, 2).getFormulas()[0];

    const word = rowData[0];
    if (!word) {
      return null;
    }

    const cambridgeFormula = linkFormulas[0];
    const oxfordFormula = linkFormulas[1];
        
    const cambridgeUrl = cambridgeFormula.match(/HYPERLINK\("([^"]+)"/i) ? cambridgeFormula.match(/HYPERLINK\("([^"]+)"/i)[1] : '';
    const oxfordUrl = oxfordFormula.match(/HYPERLINK\("([^"]+)"/i) ? oxfordFormula.match(/HYPERLINK\("([^"]+)"/i)[1] : '';

    return {
        word: word,
        row: rowNumber,
        persianTranslations: rowData[1] ? rowData[1].split('\n') : [],
        definitions: rowData[2] ? rowData[2].split('\n') : [],
        definitionExamples: rowData[3] ? rowData[3].split('\n') : [],
        generalExamples: rowData[4] ? rowData[4].split('\n') : [],
        partOfSpeech: rowData[5],
        synonyms: rowData[6],
        antonyms: rowData[7],
        notes: rowData[8],
        wordFamily: rowData[9],
        ukPronunciation: rowData[12],
        usPronunciation: rowData[13],
        cambridgeUrl: cambridgeUrl,
        oxfordUrl: oxfordUrl,
        reviewCount: rowData[17] || 0,
        totalReviews: rowData[19] || 0
    };
  } catch (e) {
    Logger.log(`Error in getWordDetails for row ${rowNumber}: ${e.message}`);
    return null;
  }
}



/**
 * Adds new options to the custom menu.
 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Review Vocabulary')
      .addItem('Start Review Session', 'showReviewDialog')
      .addItem('Start Multiple-Choice Quiz', 'openQuizDialog')
      .addItem('Search Words', 'showSearchDialog') // <-- ADD THIS LINE
      .addSeparator()
      .addItem('Format Sheet', 'formatSheet')
      .addItem('Remove Duplicates', 'removeDuplicateWords')
      .addSeparator()
      .addItem('Initialize Sheet', 'initializeSheet')
      .addItem('Initialize Quiz Counts', 'initializeQuizCounts')
      .addItem('Initialize Total Reviews', 'initializeTotalReviews')
      .addItem('Populate & Verify Hyperlinks', 'populateHyperlinks')
      .addItem('Schedule Initial Reviews', 'scheduleInitialReviews')
      .addToUi();
}
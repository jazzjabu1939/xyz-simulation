// XYZ Simulation Results - Google Apps Script
// Deploy as Web App: Execute as "Me", Access "Anyone"

const SHEET_ID = '135j2j9ZOYLFmDT7toXZVt331wHo7G6YpF4F-_1MDs1c';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    
    const sheet = SpreadsheetApp.openById(SHEET_ID).getActiveSheet();
    
    // Build row: Session ID, Timestamp, Score, Percentage, Time Used, Q1-Q20
    const row = [
      data.sessionId,
      new Date().toISOString(),
      data.score,
      data.percentage + '%',
      data.timeUsed,
    ];
    
    // Add answers Q1-Q20
    for (let i = 0; i < 20; i++) {
      const answer = data.answers[i] || 'N/A';
      const correct = data.correctAnswers[i];
      const isCorrect = answer === correct;
      row.push(isCorrect ? '✓ ' + answer.toUpperCase() : '✗ ' + answer.toUpperCase());
    }
    
    sheet.appendRow(row);
    
    return ContentService
      .createTextOutput(JSON.stringify({ success: true, sessionId: data.sessionId }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'XYZ Simulation Results API is running' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Test function
function testPost() {
  const testData = {
    postData: {
      contents: JSON.stringify({
        sessionId: 'TEST01',
        score: 15,
        percentage: 75,
        timeUsed: '4:32',
        answers: {0: 'yes', 1: 'no', 2: 'yes', 3: 'yes', 4: 'no', 5: 'yes', 6: 'no', 7: 'yes', 8: 'no', 9: 'yes', 10: 'yes', 11: 'no', 12: 'yes', 13: 'no', 14: 'yes', 15: 'yes', 16: 'no', 17: 'yes', 18: 'no', 19: 'yes'},
        correctAnswers: ['yes', 'no', 'yes', 'no', 'no', 'yes', 'yes', 'yes', 'no', 'yes', 'yes', 'no', 'no', 'no', 'yes', 'yes', 'no', 'yes', 'yes', 'yes']
      })
    }
  };
  
  const result = doPost(testData);
  Logger.log(result.getContent());
}

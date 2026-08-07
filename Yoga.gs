function getYogaList() {

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.YOGA);

  const data = sheet.getDataRange().getValues();

  data.shift(); // Remove header

  return data.map(row => ({
    yogaId: row[0],
    name: row[1],
    category: row[2],
    difficulty: row[3],
    duration: row[4],
    benefits: row[5],
    instructions: row[6],
    image: row[7],
    video: row[8],
    favorite: row[9]
  }));

}
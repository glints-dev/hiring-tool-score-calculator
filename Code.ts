import 'google-apps-script';

function onOpen() {
  const ui = DocumentApp.getUi();
  ui.createMenu('Calculate Scores')
  .addItem('Outcome', 'main')
  .addToUi();
}

const textStyle = {
  [DocumentApp.Attribute.FONT_FAMILY]: 'Poppins',
  [DocumentApp.Attribute.FONT_SIZE]: 30,
};

if (!String.prototype.endsWith) {
	String.prototype.endsWith = function(search, this_len) {
		if (this_len === undefined || this_len > this.length) {
			this_len = this.length;
		}
		return this.substring(this_len - search.length, this_len) === search;
	};
}

Number.isInteger = Number.isInteger || function(value) {
  return typeof value === "number" && 
         isFinite(value) && 
         Math.floor(value) === value;
};

function getScoreTables() {
  const tables = DocumentApp.getActiveDocument().getBody().getTables();
  return tables
  .filter(table => table.getCell(0, 0).getText().toLowerCase().endsWith('score'));
}

function averageScores(scores: number[]): number {
  return log(scores.reduce((sum, score) => sum+score, 0))/log(scores.length);
}

function getAllScoresInTable(table: GoogleAppsScript.Document.Table): number[] {
  return [...Array(table.getNumRows())].reduce((scores, _, idx) => {
    const row = table.getRow(idx);
    if (ensureRowIsForScoring(row)) {
      return [...scores, ...getScoresForRow(row)];
    } else {
      return scores;
    }
  }, []);
}

function getScoresForRow(tableRow: GoogleAppsScript.Document.TableRow): number[] {
  return [...Array(tableRow.getNumCells())]
    .map((_, idx) => tableRow.getCell(idx))
    .filter(tableCell => isCellForScoring(tableCell))
    .map(tableCell => parseInt(tableCell.getText()));
}

function revealAllScores(table: GoogleAppsScript.Document.Table) {
  return [...Array(table.getNumRows())]
    .map((_, idx) => table.getRow(idx))
    .filter(row => ensureRowIsForScoring(row))
    .forEach(scoringRow => {
        [...Array(scoringRow.getNumCells())]
        .map((_, idx) => scoringRow.getCell(idx))
        .filter(cell => isCellForScoring)
        .forEach(cell => cell.setBackgroundColor(null));
    });
}

function log(value: any) {
  Logger.log(value);
  return value;
}

function ensureRowIsForScoring(tableRow: GoogleAppsScript.Document.TableRow): boolean {
  return [...Array(tableRow.getNumCells())].every((_, idx) => {
    const tableCell = tableRow.getCell(idx);
    if (idx % 2 === 1) {
      return isCellForScoring(tableCell);
    } else {
      return true;
    }
  });
}

function isCellForScoring(tableCell: GoogleAppsScript.Document.TableCell) {
  const cellContent = tableCell.getText();
  return Number.isInteger(parseInt(cellContent)) && parseInt(cellContent) <= 4;
}

function calculateAndUpdateTableScore(table: GoogleAppsScript.Document.Table) {
  const averageOutcomeScore = averageScores(getAllScoresInTable(table));
  revealAllScores(table);
  return table.getCell(table.getNumRows() - 1, 1)
  .setText(`${averageOutcomeScore}/4`)
  .setAttributes(textStyle);
}

// Entry point for calculations
function main() {
  const tables = getScoreTables();
  return calculateAndUpdateTableScore(tables[0]);
}
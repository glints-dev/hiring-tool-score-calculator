import 'google-apps-script';

function onOpen() {
  const ui = DocumentApp.getUi();
  ui.createMenu('Calculate Scores')
  .addItem('Outcome', 'calculateOutcomeScore')
  .addToUi();
}

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
  return scores.reduce((sum, score) => sum+score, 0)/scores.length;
}

function getAllScoresInTable(table: GoogleAppsScript.Document.Table): number[] {
  return new Array(table.getNumRows()).reduce((scores, _, idx) => {
    const row = table.getRow(idx);
    if (ensureRowIsForScoring(row)) {
      return [...scores, ...getScoresForRow(row)];
    }
  }, []);
}

function getScoresForRow(tableRow: GoogleAppsScript.Document.TableRow): number[] {
  return new Array(tableRow.getNumCells())
    .map((_, idx) => tableRow.getCell(idx))
    .filter(tableCell => isCellForScoring(tableCell))
    .map(tableCell => parseInt(tableCell.getText()));
}

function ensureRowIsForScoring(tableRow: GoogleAppsScript.Document.TableRow): boolean {
  return new Array(tableRow.getNumCells()).every((_, idx) => {
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
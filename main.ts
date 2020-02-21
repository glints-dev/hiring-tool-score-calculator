import 'google-apps-script';

const maxScore = 4;

function onOpen() {
  const ui = DocumentApp.getUi();
  const menu = ui.createMenu('Scoring Calculator')
  .addItem('Calculate all scores', 'calculateAllTables');

  menu.addToUi();
}

function getScoreTables() {
  const tables = DocumentApp.getActiveDocument().getBody().getTables();
  return tables
  .filter(table => table.getCell(0, 0).getText().toLowerCase().endsWith('score'));
}

function getAverageScores(scores: number[]): number {
  return scores.reduce((sum, score) => sum+score, 0)/scores.length || 0;
}

function getAllScoresInTable(table: GoogleAppsScript.Document.Table): number[] {
  return Array(table.getNumRows()).fill().reduce((scores, _, idx) => {
    const row = table.getRow(idx);
    if (ensureRowIsForScoring(row)) {
      return [...scores, ...getScoresForRow(row)];
    } else {
      return scores;
    }
  }, []);
}

function getScoresForRow(tableRow: GoogleAppsScript.Document.TableRow): number[] {
  return Array(tableRow.getNumCells()).fill(0)
    .map((_, idx) => tableRow.getCell(idx))
    .filter(tableCell => isCellForScoring(tableCell))
    .map(tableCell => parseInt(tableCell.getText()));
}

function revealTableScores(table: GoogleAppsScript.Document.Table) {
  return Array(table.getNumRows()).fill(0)
    .map((_, idx) => table.getRow(idx))
    .filter(row => ensureRowIsForScoring(row))
    .forEach(scoringRow => {
        Array(scoringRow.getNumCells()).fill(0)
        .map((_, idx) => scoringRow.getCell(idx))
        .filter(cell => isCellForScoring)
        .forEach(cell => cell.setBackgroundColor(null));
    });
}

function ensureRowIsForScoring(tableRow: GoogleAppsScript.Document.TableRow): boolean {
  return Array(tableRow.getNumCells()).fill(0).every((_, idx) => {
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
  return Number.isInteger(parseInt(cellContent)) && parseInt(cellContent) <= maxScore;
}

function updateTableScore(table: GoogleAppsScript.Document.Table, averageScore: number) {
  return table.getCell(table.getNumRows() - 1, 1)
  .setText(`${averageScore.toFixed(2)}/${maxScore}`);
}

function getTableTitle(table: GoogleAppsScript.Document.Table): string {
  return table.getCell(0, 0).getText();
}

// Calculate final average score
function updateFinalAverageScore(averageScoreTable: GoogleAppsScript.Document.Table,
   averageScores: number[]): void {
  const finalAverageScore = getAverageScores(averageScores);
  // Update score cell
  averageScoreTable.getCell(1,1).setText(`${finalAverageScore.toFixed(2)}/${maxScore}`);
 
  // Compute and update result
  averageScoreTable.getCell(averageScoreTable.getNumRows() - 1, 1)
  .setText(finalAverageScore >= 3 ? 'Hire' : 'Reject');
}

// Entry point for calculations
function calculateAllTables() {
  const tables = getScoreTables();
  const averageScoreTable = tables.splice(
    tables.findIndex(table => getTableTitle(table).toUpperCase() === 'AVERAGE SCORE'), 1
  );

  const averageScores = [];
  // Calculate average score for every sccoring table
  tables.forEach(table => {
    const averageScore = getAverageScores(getAllScoresInTable(table));
    updateTableScore(table, averageScore);
    revealTableScores(table);
    averageScores.push(averageScore);
  });

  updateFinalAverageScore(averageScoreTable[0], averageScores);
}

const { exit } = require("process");
const XLSX = require("xlsx");

const INFO_SHEETS = "info";
const DATA_IGNORE_SHEETS = [INFO_SHEETS, "def"];

interface JsonData {
  game_type?: string;
  title?: string;
  data: { [key: string]: any };
}

function getMaxRow(sheet, col, startRow) {
  let maxRow = startRow;
  while (sheet[`${col}${maxRow}`] != null) {
    maxRow++;
  }
  return maxRow;
}

function getMaxCol(sheet, row, startCol) {
  let maxCol = startCol;
  while (sheet[`${maxCol}${row}`] != null) {
    maxCol = nextChar(maxCol);
  }
  return maxCol;
}

function prevChar(c) {
  return String.fromCharCode(c.charCodeAt(0) - 1);
}

function nextChar(c) {
  return String.fromCharCode(c.charCodeAt(0) + 1);
}

async function run(filepath: string, json: JsonData) {
  const xlsx = XLSX.readFile(filepath);
  const sheetNames = xlsx.SheetNames;

  for (let i = 0; i < sheetNames.length; i++) {
    const sheetName = sheetNames[i];
    const sheet = xlsx.Sheets[sheetName];

    if (DATA_IGNORE_SHEETS.indexOf(sheetName) >= 0) {
      if (INFO_SHEETS === sheetName) {
        const startRow = 1;
        const endRow = getMaxRow(sheet, "A", startRow);
        for (let r = startRow; r < endRow; r++) {
          const key = sheet[`A${r}`];
          const value = sheet[`B${r}`];
          if (key && value) {
            json[key.v] = value.v;
          }
        }
      }
      continue;
    }

    const tableName = sheet["B1"].v;
    const startRow = sheet["B4"].v;
    const startCol = sheet["B3"].v;

    const endRow = getMaxRow(sheet, startCol, startRow);
    const endCol = getMaxCol(sheet, startRow, startCol);

    const table = json.data[tableName] ? json.data[tableName] : [];

    for (let r = startRow + 1; r < endRow; r++) {
      const cell = sheet[`${prevChar(startCol)}${r}`];
      if (cell && cell.v === "#") {
        continue;
      }

      const row = {};
      for (let c = startCol; c < endCol; c = nextChar(c)) {
        const key = sheet[`${c}${startRow}`].v;
        let value = null;
        const cell = sheet[`${c}${r}`];
        if (cell != null) {
          value = cell.v;
          if (cell.t === "s") {
            value = value.replace("\\n", "\n");
          }
        }
        row[key] = value;
      }

      for (let i = 0; i < table.length; i++) {
        const key = sheet[`${startCol}${startRow}`].v;
        if (row[key] === table[i][key]) {
          return `Duplicated key!! sheet=${sheetName}, table=${tableName}, ${key}=${row[key]}`;
        }
      }
      table.push(row);
    }
    json.data[tableName] = table;
  }

  return null;
}

export const xlsxParser = {
  run,
};

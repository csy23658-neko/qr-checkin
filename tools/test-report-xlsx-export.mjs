import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';

const source = fs.readFileSync(new URL('../app.js', import.meta.url), 'utf8');

let capturedRows;
let writtenWorkbook;
let writtenFilename;

function encodeCell({ r, c }) {
  let column = '';
  for (let value = c + 1; value > 0; value = Math.floor((value - 1) / 26)) {
    column = String.fromCharCode(65 + ((value - 1) % 26)) + column;
  }
  return `${column}${r + 1}`;
}

const XLSX = {
  utils: {
    aoa_to_sheet(rows) {
      capturedRows = JSON.parse(JSON.stringify(rows));
      const sheet = {};
      rows.forEach((row, r) => row.forEach((value, c) => {
        if (value !== undefined && value !== null && value !== '') {
          sheet[encodeCell({ r, c })] = { v: value };
        }
      }));
      return sheet;
    },
    encode_cell: encodeCell,
    book_new() {
      return { sheets: [] };
    },
    book_append_sheet(book, sheet, name) {
      book.sheets.push({ name, sheet });
    },
  },
  writeFile(book, filename) {
    writtenWorkbook = book;
    writtenFilename = filename;
  },
};

const context = vm.createContext({
  alert(message) {
    throw new Error(`Unexpected alert: ${message}`);
  },
  Blob,
  clearInterval,
  console,
  CONFIG: { GOOGLE_CLIENT_ID: '', API_BASE_URL: '' },
  document: {},
  fetch,
  navigator: {},
  setInterval,
  setTimeout,
  URL,
  window: { addEventListener() {} },
  XLSX,
});

vm.runInContext(source, context);
vm.runInContext(`
  state.sheetId = 'event-sheet-id';
  state.spreadsheetTitle = '測試活動';
  state.rows = [
    { id: '001', name: '王小明', address: '測試地址', notes: '備註', checkinStamp: '2026/09/19 09:30:00', checkinBy: '志工甲' },
    { id: '002', name: '陳小華', address: '', notes: '', checkinStamp: '', checkinBy: '' },
  ];
  downloadReportXlsx();
`, context);

assert.deepEqual(capturedRows, [
  ['編號', '姓名', '通訊地址', '備註', '報到時間', '報到人員'],
  ['001', '王小明', '測試地址', '備註', '2026/09/19 09:30:00', '志工甲'],
  ['002', '陳小華', '', '', '', ''],
]);
assert.equal(capturedRows.every(row => row.length === 6), true);
assert.equal(capturedRows[1][0], '001');
assert.equal(writtenWorkbook.sheets[0].name, '出席報告');
assert.equal(writtenWorkbook.sheets[0].sheet['!cols'].length, 6);
assert.equal(writtenWorkbook.sheets[0].sheet.A2.z, '@');
assert.equal(writtenWorkbook.sheets[0].sheet.F2.z, '@');
assert.match(writtenFilename, /^測試活動_\d{8}\.xlsx$/);

console.log('XLSX attendance report export: PASS');

import { SpreadsheetFile, Workbook } from '@oai/artifact-tool';

const PARTICIPANT_ROWS = 1000;
const LAST_PARTICIPANT_ROW = PARTICIPANT_ROWS + 1;

const workbook = Workbook.create();
const list = workbook.worksheets.add('名單');
list.showGridLines = false;
list.getRange('A1:D1').values = [['編號', '姓名', '通訊地址', '備註']];
list.getRange(`A2:D${LAST_PARTICIPANT_ROW}`).values = Array.from({ length: PARTICIPANT_ROWS }, () => ['', '', '', '']);
list.getRange('A1:D1').format = {
  fill: '#3860B2',
  font: { bold: true, color: '#FFFFFF' },
  horizontalAlignment: 'center',
  verticalAlignment: 'center',
};
list.getRange(`A1:D${LAST_PARTICIPANT_ROW}`).format.borders = { preset: 'all', style: 'thin', color: '#D9E1F2' };
list.getRange(`A2:D${LAST_PARTICIPANT_ROW}`).format.numberFormat = '@';
list.getRange(`A2:D${LAST_PARTICIPANT_ROW}`).format.verticalAlignment = 'center';
list.getRange('A:A').format.columnWidth = 18;
list.getRange('B:B').format.columnWidth = 16;
list.getRange('C:C').format.columnWidth = 34;
list.getRange('D:D').format.columnWidth = 26;
list.getRange('1:1').format.rowHeight = 25;
list.freezePanes.freezeRows(1);

const guide = workbook.worksheets.add('填寫說明');
guide.showGridLines = false;
guide.getRange('A1:D1').merge();
guide.getRange('A1').values = [['活動名單範本｜填寫說明']];
guide.getRange('A1:D1').format = {
  fill: '#3860B2',
  font: { bold: true, color: '#FFFFFF', size: 16 },
  horizontalAlignment: 'center',
  verticalAlignment: 'center',
};
guide.getRange('A3:D7').values = [
  ['欄位', '填寫規則', '範例', '備註'],
  ['編號', '必填、不可重複，維持文字格式。編入後不可修改。', '1150326162', 'QR Code 內容即為編號。'],
  ['姓名', '必填，維持文字格式。', '王小明', '用於通知單與現場辨識。'],
  ['通訊地址', '依活動需求填寫，維持文字格式。', '臺北市...', '用於合併列印。'],
  ['備註', '可留空，維持文字格式。', '素食', '特殊需求或行政註記。'],
];
guide.getRange('A3:D3').format = {
  fill: '#DCE6F7', font: { bold: true, color: '#1F3E72' }, horizontalAlignment: 'center',
};
guide.getRange('A3:D7').format.borders = { preset: 'all', style: 'thin', color: '#D9E1F2' };
guide.getRange('A3:D7').format.wrapText = true;
guide.getRange('A:A').format.columnWidth = 15;
guide.getRange('B:B').format.columnWidth = 42;
guide.getRange('C:C').format.columnWidth = 20;
guide.getRange('D:D').format.columnWidth = 28;
guide.getRange('1:1').format.rowHeight = 30;
guide.getRange('3:7').format.rowHeight = 38;

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save('活動名單範本.xlsx');

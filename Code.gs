// ================================================================
// 완제품 입출고 관리 시스템 - Google Apps Script
// 시트 구성: 1) 입출고 기록  2) 재고 현황  3) 품목 관리
// ================================================================

const SHEET_RECORD = '입출고 기록';
const SHEET_STOCK  = '재고 현황';
const SHEET_ITEMS  = '품목 관리';

// ── 메뉴 등록 ──────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📦 입출고 관리')
    .addItem('🏗 시스템 초기화 (최초 1회)', 'initSystem')
    .addSeparator()
    .addItem('➕ 입고 등록', 'showInboundDialog')
    .addItem('➖ 출고 등록', 'showOutboundDialog')
    .addSeparator()
    .addItem('🔄 재고 현황 갱신', 'refreshStock')
    .addItem('📊 품목 목록 갱신', 'refreshItems')
    .addToUi();
}

// ================================================================
// 시스템 초기화
// ================================================================
function initSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const res = ui.alert('초기화 확인', '시트를 새로 생성합니다. 기존 데이터가 있으면 삭제됩니다. 계속하시겠습니까?', ui.ButtonSet.YES_NO);
  if (res !== ui.Button.YES) return;

  _deleteSheet(ss, SHEET_RECORD);
  _deleteSheet(ss, SHEET_STOCK);
  _deleteSheet(ss, SHEET_ITEMS);

  _createRecordSheet(ss);
  _createStockSheet(ss);
  _createItemsSheet(ss);

  _styleHeaders(ss);

  // 기본 시트 삭제
  const defaultSheet = ss.getSheetByName('시트1') || ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1) ss.deleteSheet(defaultSheet);

  ss.setActiveSheet(ss.getSheetByName(SHEET_RECORD));
  ui.alert('✅ 초기화 완료!', '입출고 기록 / 재고 현황 / 품목 관리 시트가 생성되었습니다.\n\n메뉴 [📦 입출고 관리]에서 입고·출고를 등록하세요.', ui.ButtonSet.OK);
}

function _deleteSheet(ss, name) {
  const s = ss.getSheetByName(name);
  if (s) ss.deleteSheet(s);
}

// ================================================================
// 1) 입출고 기록 시트
// ================================================================
function _createRecordSheet(ss) {
  const sheet = ss.insertSheet(SHEET_RECORD);

  // 제목
  sheet.getRange('A1:K1').merge()
    .setValue('완  제  품  입  출  고  관  리  대  장')
    .setFontSize(18).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBackground('#1a237e').setFontColor('#ffffff');
  sheet.setRowHeight(1, 50);

  // 회사 정보
  _setLabelCell(sheet, 'A2', '회사명 :');
  sheet.getRange('B2:D2').merge().setBackground('#ffffff').setBorder(true,true,true,true,false,false);
  _setLabelCell(sheet, 'E2', '부  서 :');
  sheet.getRange('F2:G2').merge().setBackground('#ffffff').setBorder(true,true,true,true,false,false);
  _setLabelCell(sheet, 'H2', '담당자 :');
  sheet.getRange('I2:K2').merge().setBackground('#ffffff').setBorder(true,true,true,true,false,false);
  sheet.setRowHeight(2, 28);

  // 기간/문서번호
  _setLabelCell(sheet, 'A3', '관리 기간 :');
  sheet.getRange('B3:D3').merge().setBackground('#ffffff').setBorder(true,true,true,true,false,false);
  _setLabelCell(sheet, 'E3', '문서번호 :');
  sheet.getRange('F3:K3').merge().setBackground('#ffffff').setBorder(true,true,true,true,false,false);
  sheet.setRowHeight(3, 28);

  // 헤더 그룹 (4행)
  const hdr4 = [['No','날짜','구분','품  목  정  보','','','수  량','','','담당자','비고']];
  sheet.getRange('A4:K4').setValues(hdr4);
  sheet.getRange('D4:F4').merge();
  sheet.getRange('G4:I4').merge();
  sheet.getRange('A4:K4')
    .setBackground('#1a237e').setFontColor('#ffffff').setFontWeight('bold').setFontSize(11)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(4, 26);

  // 헤더 서브 (5행)
  const hdr5 = [['','','','품  목  명','규격/사양','단위','입  고','출  고','재고누계','','']];
  sheet.getRange('A5:K5').setValues(hdr5);
  [['A5','#283593'],['B5','#283593'],['C5','#283593'],
   ['D5','#2e7d32'],['E5','#2e7d32'],['F5','#2e7d32'],
   ['G5','#1565c0'],['H5','#1565c0'],['I5','#1565c0'],
   ['J5','#283593'],['K5','#283593']].forEach(([cell, color]) => {
    sheet.getRange(cell).setBackground(color).setFontColor('#ffffff').setFontWeight('bold');
  });
  sheet.getRange('A5:K5').setHorizontalAlignment('center').setVerticalAlignment('middle').setFontSize(11);
  sheet.setRowHeight(5, 26);

  // 데이터 행 (6~55 = 50행)
  for (let i = 0; i < 50; i++) {
    const row = 6 + i;
    const bg = (i % 2 === 0) ? '#ffffff' : '#f8f9ff';
    sheet.setRowHeight(row, 22);
    sheet.getRange(row, 1).setValue(i + 1).setHorizontalAlignment('center').setFontColor('#aaaaaa').setBackground(bg);
    sheet.getRange(row, 2).setBackground(bg).setHorizontalAlignment('center').setNumberFormat('yyyy-mm-dd');
    sheet.getRange(row, 3).setBackground(bg).setHorizontalAlignment('center');
    sheet.getRange(row, 4).setBackground(bg);
    sheet.getRange(row, 5).setBackground(bg).setHorizontalAlignment('center');
    sheet.getRange(row, 6).setBackground(bg).setHorizontalAlignment('center');
    sheet.getRange(row, 7).setBackground(bg).setHorizontalAlignment('right').setNumberFormat('#,##0').setFontColor('#1565c0');
    sheet.getRange(row, 8).setBackground(bg).setHorizontalAlignment('right').setNumberFormat('#,##0').setFontColor('#c62828');
    if (row === 6) {
      sheet.getRange(row, 9).setFormula('=IF(G6+H6=0,"",G6-H6)');
    } else {
      sheet.getRange(row, 9).setFormula(`=IF(G${row}+H${row}=0,IF(I${row-1}="","",I${row-1}),IF(I${row-1}="",0,I${row-1})+G${row}-H${row})`);
    }
    sheet.getRange(row, 9).setBackground(bg).setHorizontalAlignment('right').setNumberFormat('#,##0').setFontWeight('bold');
    sheet.getRange(row, 10).setBackground(bg).setHorizontalAlignment('center');
    sheet.getRange(row, 11).setBackground(bg);
  }

  // 합계행
  const sumRow = 56;
  sheet.setRowHeight(sumRow, 26);
  sheet.getRange(sumRow, 1, 1, 6).merge().setValue('합   계   (TOTAL)')
    .setFontWeight('bold').setHorizontalAlignment('center').setBackground('#e8eaf6').setFontSize(11);
  sheet.getRange(sumRow, 7).setFormula('=SUM(G6:G55)').setBackground('#e8eaf6').setFontWeight('bold')
    .setHorizontalAlignment('right').setNumberFormat('#,##0').setFontColor('#1565c0');
  sheet.getRange(sumRow, 8).setFormula('=SUM(H6:H55)').setBackground('#e8eaf6').setFontWeight('bold')
    .setHorizontalAlignment('right').setNumberFormat('#,##0').setFontColor('#c62828');
  sheet.getRange(sumRow, 9).setFormula('=IFERROR(I55,"")').setBackground('#e8eaf6').setFontWeight('bold')
    .setHorizontalAlignment('right').setNumberFormat('#,##0');
  sheet.getRange(sumRow, 10, 1, 2).merge().setBackground('#e8eaf6');

  // 비고란
  sheet.setRowHeight(57, 28);
  sheet.getRange('A57:K57').merge()
    .setValue('※ 비고: 본 대장은 완제품의 입출고 현황을 기록·관리하기 위한 문서입니다. 모든 입출고는 증빙서류와 대조 확인 후 기재하시기 바랍니다.')
    .setFontSize(10).setFontColor('#555555').setBackground('#fafafa').setVerticalAlignment('middle');

  // 서명란
  sheet.setRowHeight(58, 24); sheet.setRowHeight(59, 55);
  sheet.getRange('A58:H58').merge().setBackground('#ffffff');
  ['I58','J58','K58'].forEach((c,i) => {
    sheet.getRange(c).setValue(['작  성','검  토','승  인'][i])
      .setFontWeight('bold').setHorizontalAlignment('center')
      .setBackground('#e8eaf6').setBorder(true,true,true,true,false,false);
  });
  ['I59','J59','K59'].forEach(c => {
    sheet.getRange(c).setBackground('#ffffff').setBorder(true,true,true,true,false,false);
  });

  // 테두리 & 고정
  sheet.getRange('A4:K56').setBorder(true,true,true,true,true,true,'#cccccc', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange('A4:K5').setBorder(true,true,true,true,null,null,'#1a237e', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange('A56:K56').setBorder(true,null,null,null,null,null,'#1a237e', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.setFrozenRows(5);

  // 열 너비
  sheet.setColumnWidth(1,40); sheet.setColumnWidth(2,100); sheet.setColumnWidth(3,55);
  sheet.setColumnWidth(4,160); sheet.setColumnWidth(5,90); sheet.setColumnWidth(6,50);
  sheet.setColumnWidth(7,75); sheet.setColumnWidth(8,75); sheet.setColumnWidth(9,80);
  sheet.setColumnWidth(10,70); sheet.setColumnWidth(11,120);
}

// ================================================================
// 2) 재고 현황 시트
// ================================================================
function _createStockSheet(ss) {
  const sheet = ss.insertSheet(SHEET_STOCK);

  sheet.getRange('A1:F1').merge()
    .setValue('완제품 재고 현황')
    .setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBackground('#1b5e20').setFontColor('#ffffff');
  sheet.setRowHeight(1, 45);

  const headers = [['품목명','규격/사양','단위','총 입고','총 출고','현재 재고']];
  sheet.getRange('A2:F2').setValues(headers)
    .setBackground('#2e7d32').setFontColor('#ffffff').setFontWeight('bold')
    .setHorizontalAlignment('center').setFontSize(11);
  sheet.setRowHeight(2, 26);

  // 안내 문구
  sheet.getRange('A3:F3').merge()
    .setValue('※ [📦 입출고 관리] → [재고 현황 갱신] 버튼을 누르면 자동으로 업데이트됩니다.')
    .setFontSize(10).setFontColor('#777777').setBackground('#f1f8e9').setHorizontalAlignment('center');
  sheet.setRowHeight(3, 24);

  sheet.setColumnWidth(1,160); sheet.setColumnWidth(2,100);
  sheet.setColumnWidth(3,55);  sheet.setColumnWidth(4,90);
  sheet.setColumnWidth(5,90);  sheet.setColumnWidth(6,90);
  sheet.setFrozenRows(2);
}

// ================================================================
// 3) 품목 관리 시트
// ================================================================
function _createItemsSheet(ss) {
  const sheet = ss.insertSheet(SHEET_ITEMS);

  sheet.getRange('A1:D1').merge()
    .setValue('품목 관리')
    .setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBackground('#4a148c').setFontColor('#ffffff');
  sheet.setRowHeight(1, 45);

  const headers = [['품목명','규격/사양','단위','비고']];
  sheet.getRange('A2:D2').setValues(headers)
    .setBackground('#6a1b9a').setFontColor('#ffffff').setFontWeight('bold')
    .setHorizontalAlignment('center').setFontSize(11);
  sheet.setRowHeight(2, 26);

  // 샘플 품목
  const samples = [
    ['완제품 A', '100×200mm', 'EA', ''],
    ['완제품 B', '50×100mm',  'BOX', ''],
    ['완제품 C', '200×300mm', 'SET', ''],
  ];
  sheet.getRange(3, 1, samples.length, 4).setValues(samples);

  for (let r = 3; r <= 52; r++) {
    const bg = (r % 2 === 0) ? '#f3e5f5' : '#ffffff';
    sheet.getRange(r, 1, 1, 4).setBackground(bg);
    sheet.setRowHeight(r, 22);
  }

  sheet.getRange('A2:D52').setBorder(true,true,true,true,true,true,'#ce93d8', SpreadsheetApp.BorderStyle.SOLID);
  sheet.setColumnWidth(1,160); sheet.setColumnWidth(2,100);
  sheet.setColumnWidth(3,60);  sheet.setColumnWidth(4,150);
  sheet.setFrozenRows(2);
}

// ================================================================
// 헤더 스타일 (공통)
// ================================================================
function _styleHeaders(ss) {
  const configs = [
    { name: SHEET_RECORD, color: '#1565C0' },
    { name: SHEET_STOCK,  color: '#2E7D32' },
    { name: SHEET_ITEMS,  color: '#6A1B9A' },
  ];
  configs.forEach(cfg => {
    const sheet = ss.getSheetByName(cfg.name);
    if (!sheet) return;
    const header = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    header.setBackground(cfg.color)
          .setFontColor('#FFFFFF')
          .setFontWeight('bold')
          .setFontSize(11);
    sheet.setFrozenRows(1);
  });
}

function _setLabelCell(sheet, cell, value) {
  sheet.getRange(cell).setValue(value)
    .setFontWeight('bold').setBackground('#e8eaf6')
    .setBorder(true,true,true,true,false,false);
}

// ================================================================
// 재고 현황 갱신
// ================================================================
function refreshStock() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const recordSheet = ss.getSheetByName(SHEET_RECORD);
  const stockSheet  = ss.getSheetByName(SHEET_STOCK);
  if (!recordSheet || !stockSheet) {
    SpreadsheetApp.getUi().alert('먼저 시스템을 초기화해주세요.');
    return;
  }

  const lastRow = recordSheet.getLastRow();
  if (lastRow < 6) {
    SpreadsheetApp.getUi().alert('입출고 기록 데이터가 없습니다.');
    return;
  }

  const data = recordSheet.getRange(6, 1, lastRow - 5, 9).getValues();
  const map = {};

  data.forEach(row => {
    const type = row[2]; const name = row[3]; const spec = row[4];
    const unit = row[5]; const inQty = Number(row[6])||0; const outQty = Number(row[7])||0;
    if (!name) return;
    if (!map[name]) map[name] = { spec, unit, inQty: 0, outQty: 0 };
    map[name].inQty  += inQty;
    map[name].outQty += outQty;
  });

  // 기존 데이터 지우기 (4행부터)
  const existingRows = stockSheet.getLastRow();
  if (existingRows >= 4) stockSheet.getRange(4, 1, existingRows - 3, 6).clearContent().clearFormat();

  const rows = Object.entries(map).map(([name, v], i) => {
    const bg = (i % 2 === 0) ? '#ffffff' : '#f1f8e9';
    return [name, v.spec, v.unit, v.inQty, v.outQty, v.inQty - v.outQty];
  });

  if (rows.length > 0) {
    stockSheet.getRange(4, 1, rows.length, 6).setValues(rows);
    rows.forEach((_, i) => {
      const r = 4 + i;
      const bg = (i % 2 === 0) ? '#ffffff' : '#f1f8e9';
      stockSheet.getRange(r, 1, 1, 6).setBackground(bg);
      stockSheet.getRange(r, 4, 1, 3).setNumberFormat('#,##0').setHorizontalAlignment('right');
      // 재고 0 이하 강조
      const stock = rows[i][5];
      if (stock <= 0) stockSheet.getRange(r, 6).setBackground('#ffcdd2').setFontColor('#c62828').setFontWeight('bold');
      stockSheet.setRowHeight(r, 22);
    });
    stockSheet.getRange(4, 1, rows.length, 6)
      .setBorder(true,true,true,true,true,true,'#a5d6a7', SpreadsheetApp.BorderStyle.SOLID);
  }

  SpreadsheetApp.getUi().alert(`✅ 재고 현황 갱신 완료!\n총 ${rows.length}개 품목`);
}

// ================================================================
// 품목 목록 갱신
// ================================================================
function refreshItems() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const itemsSheet = ss.getSheetByName(SHEET_ITEMS);
  if (!itemsSheet) {
    SpreadsheetApp.getUi().alert('먼저 시스템을 초기화해주세요.');
    return;
  }
  SpreadsheetApp.getUi().alert('품목 관리 시트에서 직접 품목을 추가/수정하세요.\n(A열: 품목명, B열: 규격, C열: 단위)');
  ss.setActiveSheet(itemsSheet);
}

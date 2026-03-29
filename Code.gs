// ================================================================
// 완제품 입출고 관리 시스템 - Google Apps Script
// 시트 구성: 1) 입출고 기록  2) 재고 현황  3) 품목 관리
// ================================================================

const SHEET_RECORD      = '입출고 기록';
const SHEET_STOCK       = '재고 현황';
const SHEET_ITEMS       = '품목 관리';
const SHEET_STATS       = '통계';
const SHEET_LEADERBOARD = '리드보드';

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
    .addItem('📊 통계 갱신', 'refreshStats')
    .addItem('🏆 리드보드 갱신', 'refreshLeaderboard')
    .addItem('📋 품목 목록 갱신', 'refreshItems')
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
  _deleteSheet(ss, SHEET_STATS);
  _deleteSheet(ss, SHEET_LEADERBOARD);

  _createRecordSheet(ss);
  _createStockSheet(ss);
  _createItemsSheet(ss);
  _createStatsSheet(ss);
  _createLeaderboardSheet(ss);

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
// ================================================================
// 4) 통계 시트 생성
// ================================================================
function _createStatsSheet(ss) {
  const sheet = ss.insertSheet(SHEET_STATS);

  sheet.getRange('A1:G1').merge()
    .setValue('📊 입출고 통계')
    .setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBackground('#E65100').setFontColor('#ffffff');
  sheet.setRowHeight(1, 45);

  sheet.getRange('A2:G2').merge()
    .setValue('※ [📦 입출고 관리] → [통계 갱신] 버튼을 누르면 자동으로 업데이트됩니다.')
    .setFontSize(10).setFontColor('#777777').setBackground('#FFF3E0').setHorizontalAlignment('center');
  sheet.setRowHeight(2, 24);

  // 월별 집계 헤더
  sheet.getRange('A3:G3').merge()
    .setValue('▶ 월별 입출고 집계')
    .setFontSize(12).setFontWeight('bold').setBackground('#BF360C').setFontColor('#ffffff')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(3, 26);

  const monthHeaders = [['연월', '입고 건수', '입고 수량', '출고 건수', '출고 수량', '순증감', '비고']];
  sheet.getRange('A4:G4').setValues(monthHeaders)
    .setBackground('#FF8A65').setFontColor('#ffffff').setFontWeight('bold')
    .setHorizontalAlignment('center').setFontSize(11);
  sheet.setRowHeight(4, 24);

  // 일별 집계 헤더 (월별 데이터 아래에 동적으로 추가됨 — 갱신 시 위치 결정)
  sheet.setColumnWidth(1, 90); sheet.setColumnWidth(2, 80); sheet.setColumnWidth(3, 90);
  sheet.setColumnWidth(4, 80); sheet.setColumnWidth(5, 90); sheet.setColumnWidth(6, 80);
  sheet.setColumnWidth(7, 100);
  sheet.setFrozenRows(4);
}

// ================================================================
// 5) 리드보드 시트 생성
// ================================================================
function _createLeaderboardSheet(ss) {
  const sheet = ss.insertSheet(SHEET_LEADERBOARD);

  sheet.getRange('A1:F1').merge()
    .setValue('🏆 리드보드')
    .setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBackground('#1A237E').setFontColor('#FFD700');
  sheet.setRowHeight(1, 45);

  sheet.getRange('A2:F2').merge()
    .setValue('※ [📦 입출고 관리] → [리드보드 갱신] 버튼을 누르면 자동으로 업데이트됩니다.')
    .setFontSize(10).setFontColor('#777777').setBackground('#E8EAF6').setHorizontalAlignment('center');
  sheet.setRowHeight(2, 24);

  sheet.setColumnWidth(1, 50);  sheet.setColumnWidth(2, 170);
  sheet.setColumnWidth(3, 100); sheet.setColumnWidth(4, 90);
  sheet.setColumnWidth(5, 90);  sheet.setColumnWidth(6, 90);
  sheet.setFrozenRows(2);
}

// ================================================================
// 통계 갱신
// ================================================================
function refreshStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const recordSheet = ss.getSheetByName(SHEET_RECORD);
  const statsSheet  = ss.getSheetByName(SHEET_STATS);
  if (!recordSheet || !statsSheet) {
    SpreadsheetApp.getUi().alert('먼저 시스템을 초기화해주세요.');
    return;
  }

  const lastRow = recordSheet.getLastRow();
  if (lastRow < 6) {
    SpreadsheetApp.getUi().alert('입출고 기록 데이터가 없습니다.');
    return;
  }

  const data = recordSheet.getRange(6, 1, lastRow - 5, 10).getValues();

  // 월별 집계
  const monthMap = {};
  data.forEach(row => {
    const date = row[1]; if (!date) return;
    const d = (date instanceof Date) ? date : new Date(date);
    if (isNaN(d)) return;
    const ym = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM');
    const inQty  = Number(row[6]) || 0;
    const outQty = Number(row[7]) || 0;
    if (!monthMap[ym]) monthMap[ym] = { inCnt: 0, inQty: 0, outCnt: 0, outQty: 0 };
    if (inQty  > 0) { monthMap[ym].inCnt++;  monthMap[ym].inQty  += inQty;  }
    if (outQty > 0) { monthMap[ym].outCnt++; monthMap[ym].outQty += outQty; }
  });

  // 기존 데이터 삭제 (5행부터)
  const existLast = statsSheet.getLastRow();
  if (existLast >= 5) statsSheet.getRange(5, 1, existLast - 4, 7).clearContent().clearFormat();

  const months = Object.keys(monthMap).sort();
  const monthRows = months.map((ym, i) => {
    const v = monthMap[ym];
    return [ym, v.inCnt, v.inQty, v.outCnt, v.outQty, v.inQty - v.outQty, ''];
  });

  if (monthRows.length > 0) {
    const r = statsSheet.getRange(5, 1, monthRows.length, 7);
    r.setValues(monthRows);
    monthRows.forEach((row, i) => {
      const rowNum = 5 + i;
      const bg = (i % 2 === 0) ? '#ffffff' : '#FFF3E0';
      statsSheet.getRange(rowNum, 1, 1, 7).setBackground(bg).setHorizontalAlignment('center');
      statsSheet.getRange(rowNum, 3).setNumberFormat('#,##0').setFontColor('#1565C0').setFontWeight('bold');
      statsSheet.getRange(rowNum, 5).setNumberFormat('#,##0').setFontColor('#C62828').setFontWeight('bold');
      const net = row[5];
      statsSheet.getRange(rowNum, 6).setNumberFormat('+#,##0;-#,##0;0')
        .setFontColor(net >= 0 ? '#2E7D32' : '#C62828').setFontWeight('bold');
      statsSheet.setRowHeight(rowNum, 22);
    });
    statsSheet.getRange(5, 1, monthRows.length, 7)
      .setBorder(true, true, true, true, true, true, '#FFCCBC', SpreadsheetApp.BorderStyle.SOLID);

    // 합계행
    const sumRow = 5 + monthRows.length;
    statsSheet.setRowHeight(sumRow, 24);
    const totalIn  = monthRows.reduce((s, r) => s + r[2], 0);
    const totalOut = monthRows.reduce((s, r) => s + r[4], 0);
    statsSheet.getRange(sumRow, 1, 1, 2).merge().setValue('합계').setFontWeight('bold')
      .setBackground('#BF360C').setFontColor('#ffffff').setHorizontalAlignment('center');
    statsSheet.getRange(sumRow, 3).setValue(totalIn).setNumberFormat('#,##0')
      .setFontColor('#1565C0').setFontWeight('bold').setBackground('#FFF3E0').setHorizontalAlignment('center');
    statsSheet.getRange(sumRow, 4).setValue(monthRows.reduce((s,r)=>s+r[3],0))
      .setBackground('#FFF3E0').setHorizontalAlignment('center');
    statsSheet.getRange(sumRow, 5).setValue(totalOut).setNumberFormat('#,##0')
      .setFontColor('#C62828').setFontWeight('bold').setBackground('#FFF3E0').setHorizontalAlignment('center');
    statsSheet.getRange(sumRow, 6).setValue(totalIn - totalOut).setNumberFormat('+#,##0;-#,##0;0')
      .setFontColor((totalIn-totalOut)>=0?'#2E7D32':'#C62828').setFontWeight('bold')
      .setBackground('#FFF3E0').setHorizontalAlignment('center');
    statsSheet.getRange(sumRow, 7).setBackground('#FFF3E0');
  }

  SpreadsheetApp.getUi().alert(`✅ 통계 갱신 완료!\n총 ${months.length}개월 집계`);
}

// ================================================================
// 리드보드 갱신
// ================================================================
function refreshLeaderboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const recordSheet = ss.getSheetByName(SHEET_RECORD);
  const lbSheet     = ss.getSheetByName(SHEET_LEADERBOARD);
  if (!recordSheet || !lbSheet) {
    SpreadsheetApp.getUi().alert('먼저 시스템을 초기화해주세요.');
    return;
  }

  const lastRow = recordSheet.getLastRow();
  if (lastRow < 6) {
    SpreadsheetApp.getUi().alert('입출고 기록 데이터가 없습니다.');
    return;
  }

  const data = recordSheet.getRange(6, 1, lastRow - 5, 10).getValues();

  // 품목별 집계
  const itemMap = {};
  // 담당자별 집계
  const personMap = {};

  data.forEach(row => {
    const name   = row[3]; if (!name) return;
    const inQty  = Number(row[6]) || 0;
    const outQty = Number(row[7]) || 0;
    const person = row[9] || '미지정';

    if (!itemMap[name]) itemMap[name] = { inQty: 0, outQty: 0, cnt: 0 };
    itemMap[name].inQty  += inQty;
    itemMap[name].outQty += outQty;
    if (inQty > 0 || outQty > 0) itemMap[name].cnt++;

    if (!personMap[person]) personMap[person] = { inQty: 0, outQty: 0, cnt: 0 };
    personMap[person].inQty  += inQty;
    personMap[person].outQty += outQty;
    if (inQty > 0 || outQty > 0) personMap[person].cnt++;
  });

  // 기존 내용 삭제
  const existLast = lbSheet.getLastRow();
  if (existLast >= 3) lbSheet.getRange(3, 1, existLast - 2, 6).clearContent().clearFormat();

  let currentRow = 3;

  // ── 섹션 헬퍼 ──
  function writeSection(title, bgColor, headers, rows, highlight) {
    // 섹션 제목
    lbSheet.getRange(currentRow, 1, 1, 6).merge().setValue(title)
      .setBackground(bgColor).setFontColor('#ffffff').setFontSize(12).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    lbSheet.setRowHeight(currentRow, 28);
    currentRow++;

    // 헤더
    lbSheet.getRange(currentRow, 1, 1, headers.length).setValues([headers])
      .setBackground(_darken(bgColor)).setFontColor('#ffffff').setFontWeight('bold')
      .setHorizontalAlignment('center').setFontSize(11);
    lbSheet.setRowHeight(currentRow, 24);
    currentRow++;

    // 데이터
    rows.forEach((row, i) => {
      const medal = i === 0 ? '🥇' : i === 1 ? '🥈' : i === 2 ? '🥉' : `${i+1}위`;
      const fullRow = [medal, ...row];
      lbSheet.getRange(currentRow, 1, 1, fullRow.length).setValues([fullRow]);
      const bg = i < 3 ? highlight[i] : (i % 2 === 0 ? '#ffffff' : '#F5F5F5');
      lbSheet.getRange(currentRow, 1, 1, 6).setBackground(bg).setHorizontalAlignment('center');
      lbSheet.getRange(currentRow, 2).setHorizontalAlignment('left');
      lbSheet.getRange(currentRow, 3, 1, 4).setNumberFormat('#,##0');
      lbSheet.setRowHeight(currentRow, 22);
      currentRow++;
    });
    lbSheet.getRange(currentRow - rows.length - 2, 1, rows.length + 2, 6)
      .setBorder(true, true, true, true, true, true, '#BDBDBD', SpreadsheetApp.BorderStyle.SOLID);
    currentRow++; // 빈 행
    lbSheet.setRowHeight(currentRow - 1, 10);
  }

  const goldHighlight   = ['#FFF8E1', '#F5F5F5', '#F9F9F9'];

  // ── TOP 5: 입고량 순위 ──
  const inTop = Object.entries(itemMap)
    .map(([n, v]) => [n, v.inQty, v.outQty, v.inQty - v.outQty, v.cnt])
    .sort((a, b) => b[1] - a[1]).slice(0, 5);
  writeSection('📥 입고량 TOP 5', '#1565C0',
    ['순위', '품목명', '총 입고', '총 출고', '현재 재고', '거래 건수'],
    inTop, goldHighlight);

  // ── TOP 5: 출고량 순위 ──
  const outTop = Object.entries(itemMap)
    .map(([n, v]) => [n, v.inQty, v.outQty, v.inQty - v.outQty, v.cnt])
    .sort((a, b) => b[2] - a[2]).slice(0, 5);
  writeSection('📤 출고량 TOP 5', '#C62828',
    ['순위', '품목명', '총 입고', '총 출고', '현재 재고', '거래 건수'],
    outTop, goldHighlight);

  // ── TOP 5: 재고 순위 ──
  const stockTop = Object.entries(itemMap)
    .map(([n, v]) => [n, v.inQty, v.outQty, v.inQty - v.outQty, v.cnt])
    .sort((a, b) => (b[1]-b[2]) - (a[1]-a[2])).slice(0, 5);
  writeSection('📦 현재 재고 TOP 5', '#2E7D32',
    ['순위', '품목명', '총 입고', '총 출고', '현재 재고', '거래 건수'],
    stockTop, goldHighlight);

  // ── 담당자별 처리량 순위 ──
  const personTop = Object.entries(personMap)
    .map(([n, v]) => [n, v.inQty, v.outQty, v.inQty + v.outQty, v.cnt])
    .sort((a, b) => b[3] - a[3]);
  writeSection('👤 담당자별 처리량 순위', '#4A148C',
    ['순위', '담당자', '입고 수량', '출고 수량', '총 처리량', '건수'],
    personTop, goldHighlight);

  SpreadsheetApp.getUi().alert('✅ 리드보드 갱신 완료!');
}

function _darken(hex) {
  // 간단한 색상 어둡게 (헤더용)
  const map = {
    '#1565C0': '#0D47A1', '#C62828': '#B71C1C', '#2E7D32': '#1B5E20',
    '#4A148C': '#38006B', '#BF360C': '#870000', '#1A237E': '#0D1B6E',
    '#E65100': '#BF360C'
  };
  return map[hex] || hex;
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

// ================================================================
// 완제품 입출고 관리 시스템 - AppSheet 최적화 버전
// 시트 구성:
//   1) 입출고기록  - 메인 트랜잭션 (AppSheet 메인 테이블)
//   2) 품목        - 품목 마스터 (AppSheet 참조 테이블)
//   3) 재고현황    - 품목별 재고 요약 (AppSheet 읽기 전용)
// ================================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📦 입출고 관리')
    .addItem('🏗 시스템 초기화 (최초 1회)', 'initAppSheet')
    .addSeparator()
    .addItem('🔄 재고현황 갱신', 'refreshStock')
    .addToUi();
}

// ================================================================
// 초기화
// ================================================================
function initAppSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const res = ui.alert('초기화', '기존 시트를 삭제하고 새로 생성합니다. 계속하시겠습니까?', ui.ButtonSet.YES_NO);
  if (res !== ui.Button.YES) return;

  ['입출고기록','품목','재고현황'].forEach(name => {
    const s = ss.getSheetByName(name);
    if (s) ss.deleteSheet(s);
  });

  _createItemSheet(ss);
  _createRecordSheet(ss);
  _createStockSheet(ss);

  // 기본 시트 제거
  ['시트1','Sheet1'].forEach(name => {
    const s = ss.getSheetByName(name);
    if (s && ss.getSheets().length > 1) ss.deleteSheet(s);
  });

  ss.setActiveSheet(ss.getSheetByName('입출고기록'));

  ui.alert('✅ 완료!',
    '시트 3개가 생성되었습니다.\n\n' +
    '▶ AppSheet 연결 방법:\n' +
    '1. appsheet.com 접속\n' +
    '2. Create → App → Start with existing data\n' +
    '3. 이 구글 시트 선택\n' +
    '4. 자동으로 앱 생성!',
    ui.ButtonSet.OK);
}

// ================================================================
// 1) 품목 시트 (마스터 데이터)
// ================================================================
function _createItemSheet(ss) {
  const sheet = ss.insertSheet('품목');

  // AppSheet용 헤더 - 영문 포함하면 타입 인식률 높아짐
  const headers = [
    ['품목ID', '품목명', '규격_사양', '단위', '안전재고', '비고']
  ];
  sheet.getRange('A1:F1').setValues(headers);

  // 샘플 데이터
  const samples = [
    ['ITEM-001', '완제품 A', '100×200mm', 'EA',  10, ''],
    ['ITEM-002', '완제품 B', '50×100mm',  'BOX', 5,  ''],
    ['ITEM-003', '완제품 C', '200×300mm', 'SET', 3,  ''],
  ];
  sheet.getRange(2, 1, samples.length, 6).setValues(samples);

  // 스타일
  sheet.getRange('A1:F1')
    .setBackground('#4a148c').setFontColor('#ffffff')
    .setFontWeight('bold').setHorizontalAlignment('center');

  sheet.setColumnWidth(1, 90);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 110);
  sheet.setColumnWidth(4, 60);
  sheet.setColumnWidth(5, 80);
  sheet.setColumnWidth(6, 150);
  sheet.setFrozenRows(1);

  // 데이터 유효성 - 단위 드롭다운
  const unitRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['EA','BOX','SET','KG','TON','M','L'], true).build();
  sheet.getRange('D2:D200').setDataValidation(unitRule);
}

// ================================================================
// 2) 입출고기록 시트 (메인)
// ================================================================
function _createRecordSheet(ss) {
  const sheet = ss.insertSheet('입출고기록');

  // ★ AppSheet 핵심: 첫 행이 컬럼명, ID 컬럼 필수
  const headers = [
    ['기록ID', '날짜', '구분', '품목명', '규격_사양', '단위', '입고수량', '출고수량', '담당자', '비고', '등록시각']
  ];
  sheet.getRange('A1:K1').setValues(headers);

  // 샘플 데이터 2건
  const today = new Date();
  const samples = [
    ['REC-001', today, '입고', '완제품 A', '100×200mm', 'EA', 100, 0, '홍길동', '초도입고', today],
    ['REC-002', today, '출고', '완제품 A', '100×200mm', 'EA', 0,  30, '홍길동', '납품',     today],
  ];
  sheet.getRange(2, 1, samples.length, 11).setValues(samples);

  // 스타일
  sheet.getRange('A1:K1')
    .setBackground('#1a237e').setFontColor('#ffffff')
    .setFontWeight('bold').setHorizontalAlignment('center');

  // 날짜 서식
  sheet.getRange('B2:B1000').setNumberFormat('yyyy-mm-dd');
  sheet.getRange('K2:K1000').setNumberFormat('yyyy-mm-dd hh:mm');

  // 수량 서식
  sheet.getRange('G2:H1000').setNumberFormat('#,##0');

  // 구분 드롭다운
  const typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['입고','출고'], true).build();
  sheet.getRange('C2:C1000').setDataValidation(typeRule);

  // 품목명 드롭다운 (품목 시트 연동)
  const itemRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getSheetByName('품목').getRange('B2:B200'), true).build();
  sheet.getRange('D2:D1000').setDataValidation(itemRule);

  // 열 너비
  sheet.setColumnWidth(1, 90);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 60);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 110);
  sheet.setColumnWidth(6, 55);
  sheet.setColumnWidth(7, 75);
  sheet.setColumnWidth(8, 75);
  sheet.setColumnWidth(9, 75);
  sheet.setColumnWidth(10, 130);
  sheet.setColumnWidth(11, 140);
  sheet.setFrozenRows(1);

  // 짝수행 배경색
  for (let r = 2; r <= 101; r++) {
    sheet.setRowHeight(r, 22);
    if (r % 2 === 0) sheet.getRange(r, 1, 1, 11).setBackground('#f8f9ff');
  }
}

// ================================================================
// 3) 재고현황 시트 (자동 계산)
// ================================================================
function _createStockSheet(ss) {
  const sheet = ss.insertSheet('재고현황');

  const headers = [['품목명', '규격_사양', '단위', '총입고', '총출고', '현재재고', '안전재고', '상태']];
  sheet.getRange('A1:H1').setValues(headers);
  sheet.getRange('A1:H1')
    .setBackground('#1b5e20').setFontColor('#ffffff')
    .setFontWeight('bold').setHorizontalAlignment('center');

  sheet.getRange('A2').setValue('※ [📦 입출고 관리] → [재고현황 갱신] 버튼으로 업데이트하세요.')
    .setFontColor('#888888').setFontSize(10);

  sheet.setColumnWidth(1, 150); sheet.setColumnWidth(2, 110);
  sheet.setColumnWidth(3, 55);  sheet.setColumnWidth(4, 80);
  sheet.setColumnWidth(5, 80);  sheet.setColumnWidth(6, 90);
  sheet.setColumnWidth(7, 80);  sheet.setColumnWidth(8, 80);
  sheet.setFrozenRows(1);
}

// ================================================================
// 재고현황 갱신
// ================================================================
function refreshStock() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const recSh  = ss.getSheetByName('입출고기록');
  const itemSh = ss.getSheetByName('품목');
  const stSh   = ss.getSheetByName('재고현황');
  if (!recSh || !stSh) { SpreadsheetApp.getUi().alert('먼저 시스템을 초기화해주세요.'); return; }

  const lastRow = recSh.getLastRow();
  const map = {};

  if (lastRow >= 2) {
    const data = recSh.getRange(2, 1, lastRow - 1, 11).getValues();
    data.forEach(row => {
      const name = row[3]; const spec = row[4]; const unit = row[5];
      const inQ  = Number(row[6])||0; const outQ = Number(row[7])||0;
      if (!name) return;
      if (!map[name]) map[name] = { spec, unit, inQ: 0, outQ: 0 };
      map[name].inQ  += inQ;
      map[name].outQ += outQ;
    });
  }

  // 안전재고 가져오기
  const safetyMap = {};
  if (itemSh) {
    const iLast = itemSh.getLastRow();
    if (iLast >= 2) {
      itemSh.getRange(2, 1, iLast - 1, 6).getValues().forEach(row => {
        if (row[1]) safetyMap[row[1]] = Number(row[4])||0;
      });
    }
  }

  // 기존 데이터 삭제
  const existing = stSh.getLastRow();
  if (existing >= 2) stSh.getRange(2, 1, existing - 1, 8).clearContent().clearFormat();

  const rows = Object.entries(map).map(([name, v]) => {
    const stock   = v.inQ - v.outQ;
    const safety  = safetyMap[name] || 0;
    const status  = stock <= 0 ? '⛔ 재고없음' : stock <= safety ? '⚠️ 부족' : '✅ 정상';
    return [name, v.spec, v.unit, v.inQ, v.outQ, stock, safety, status];
  });

  if (rows.length > 0) {
    stSh.getRange(2, 1, rows.length, 8).setValues(rows);
    rows.forEach((row, i) => {
      const r  = 2 + i;
      const bg = row[5] <= 0 ? '#ffcdd2' : row[5] <= (row[6]||0) ? '#fff9c4' : (i%2===0?'#ffffff':'#f1f8e9');
      stSh.getRange(r, 1, 1, 8).setBackground(bg);
      stSh.getRange(r, 4, 1, 3).setNumberFormat('#,##0').setHorizontalAlignment('right');
      stSh.setRowHeight(r, 22);
    });
    stSh.getRange(2, 1, rows.length, 8)
      .setBorder(true,true,true,true,true,true,'#a5d6a7', SpreadsheetApp.BorderStyle.SOLID);
  }

  SpreadsheetApp.getUi().alert(`✅ 재고현황 갱신 완료! (${rows.length}개 품목)`);
  ss.setActiveSheet(stSh);
}

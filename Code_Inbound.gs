// ================================================================
// 완제품 입고 관리대장 — Google Apps Script Web App
// 배포: 웹 앱 → "나를 대신하여 실행" → 액세스: "모든 사용자"
// 시트 컬럼: 기록ID | 날짜 | 품목명 | 규격_사양 | 수량 | 단위 | 입고처 | 담당자 | 비고 | 등록시각
// ================================================================

const INBOUND_SHEET = '완제품입고';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📥 입고 관리')
    .addItem('🏗 시트 초기화 (최초 1회)', 'initInboundSheet')
    .addToUi();
}

function initInboundSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const ok = ui.alert('초기화', '"완제품입고" 시트를 새로 생성합니다. 계속?', ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;

  let sheet = ss.getSheetByName(INBOUND_SHEET);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(INBOUND_SHEET);

  sheet.getRange('A1:J1').setValues([[
    '기록ID','날짜','품목명','규격_사양','수량','단위','입고처','담당자','비고','등록시각'
  ]]).setBackground('#0F4C81').setFontColor('#ffffff')
     .setFontWeight('bold').setHorizontalAlignment('center');

  sheet.getRange('B2:B1000').setNumberFormat('yyyy-mm-dd');
  sheet.getRange('E2:E1000').setNumberFormat('#,##0');
  sheet.getRange('J2:J1000').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  [110,100,160,110,70,55,120,80,150,140].forEach((w,i) => sheet.setColumnWidth(i+1,w));
  sheet.setFrozenRows(1);
  ss.setActiveSheet(sheet);

  ui.alert('✅ 완료!',
    '"완제품입고" 시트가 생성되었습니다.\n\n' +
    '▶ 웹 앱 배포 방법:\n' +
    '1. [배포] → [새 배포]\n' +
    '2. 종류: 웹 앱\n' +
    '3. 다음 사용자로 실행: 나\n' +
    '4. 액세스 권한: 모든 사용자\n' +
    '5. 배포 URL → inbound.html 설정 화면에 붙여넣기',
    ui.ButtonSet.OK);
}

// ================================================================
// doGet — 목록 조회
// ================================================================
function doGet(e) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(INBOUND_SHEET);
    if (!sheet) return _json({ status: 'ok', rows: [] });

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return _json({ status: 'ok', rows: [] });

    const p      = e.parameter || {};
    const from   = p.from   || '';
    const to     = p.to     || '';
    const item   = (p.item   || '').toLowerCase();
    const worker = (p.worker || '').toLowerCase();

    const rows = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;
      const dateStr = row[1] instanceof Date
        ? Utilities.formatDate(row[1], Session.getScriptTimeZone(), 'yyyy-MM-dd')
        : String(row[1]).substring(0, 10);
      if (from && dateStr < from) continue;
      if (to   && dateStr > to)   continue;
      if (item   && !String(row[2]).toLowerCase().includes(item))   continue;
      if (worker && !String(row[7]).toLowerCase().includes(worker)) continue;
      rows.push({
        id:       String(row[0]),
        date:     dateStr,
        item:     String(row[2] || ''),
        spec:     String(row[3] || ''),
        qty:      Number(row[4]) || 0,
        unit:     String(row[5] || ''),
        supplier: String(row[6] || ''),
        worker:   String(row[7] || ''),
        note:     String(row[8] || '')
      });
    }
    return _json({ status: 'ok', rows: rows });
  } catch (err) {
    return _json({ status: 'error', message: err.toString() });
  }
}

// ================================================================
// doPost — 등록 / 수정 / 삭제
// ================================================================
function doPost(e) {
  try {
    const d     = JSON.parse(e.postData.contents);
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    let   sheet = ss.getSheetByName(INBOUND_SHEET);

    if (!sheet) {
      sheet = ss.insertSheet(INBOUND_SHEET);
      sheet.appendRow(['기록ID','날짜','품목명','규격_사양','수량','단위','입고처','담당자','비고','등록시각']);
    }

    if (d.action === 'ping') {
      return _json({ status: 'ok', message: '완제품 입고 API 연결 성공! 🎉' });
    }

    if (d.action === 'add') {
      sheet.appendRow([
        d.id,
        d.date,
        d.item,
        d.spec     || '',
        Number(d.qty) || 0,
        d.unit     || '',
        d.supplier || '',
        d.worker   || '',
        d.note     || '',
        Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
      ]);
      return _json({ status: 'ok' });
    }

    if (d.action === 'update') {
      const vals = sheet.getDataRange().getValues();
      for (let i = 1; i < vals.length; i++) {
        if (String(vals[i][0]) === String(d.id)) {
          sheet.getRange(i + 1, 2, 1, 8).setValues([[
            d.date,
            d.item,
            d.spec     || '',
            Number(d.qty) || 0,
            d.unit     || '',
            d.supplier || '',
            d.worker   || '',
            d.note     || ''
          ]]);
          break;
        }
      }
      return _json({ status: 'ok' });
    }

    if (d.action === 'delete') {
      const vals = sheet.getDataRange().getValues();
      for (let i = vals.length - 1; i >= 1; i--) {
        if (String(vals[i][0]) === String(d.id)) {
          sheet.deleteRow(i + 1);
          break;
        }
      }
      return _json({ status: 'ok' });
    }

    return _json({ status: 'error', message: '알 수 없는 action: ' + d.action });
  } catch (err) {
    return _json({ status: 'error', message: err.toString() });
  }
}

function _json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

const SPREADSHEET_ID = '1sAvWHksA3hhlDRvGsXqHdbSvO07MFTCrebuVoRk9LSc';
const REQUEST_SHEET_NAME = '권한관리';

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);

    if (body.action === 'appendRows')    return handleAppendRows(body);
    if (body.action === 'requestAccess') return handleRequestAccess(body);
    if (body.action === 'approveAccess') return handleReviewAccess(body, '승인');
    if (body.action === 'rejectAccess')  return handleReviewAccess(body, '거절');

    const { sheet, keys } = body;

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const ws = ss.getSheetByName(sheet);
    if (!ws) return respond({ status: 'error', msg: 'sheet not found: ' + sheet });

    const all = ws.getDataRange().getValues();
    const headers = all[0].map(h => String(h).trim());

    const doneCol = headers.indexOf('완료여부');
    if (doneCol === -1) return respond({ status: 'error', msg: '완료여부 column not found' });

    const keyDefs = keys.map(k => ({
      idx: headers.indexOf(String(k.col).trim()),
      val: String(k.val || '').trim()
    }));

    for (let i = 1; i < all.length; i++) {
      const row = all[i];
      const match = keyDefs.every(k => k.idx >= 0 && String(row[k.idx] || '').trim() === k.val);
      if (match) {
        ws.getRange(i + 1, doneCol + 1).setValue('완료');
        SpreadsheetApp.flush();
        return respond({ status: 'ok', row: i + 1 });
      }
    }

    return respond({ status: 'error', msg: 'row not found' });
  } catch (err) {
    return respond({ status: 'error', msg: err.message });
  }
}

// ── 품목추가 (사양변경 / 단종(단품) 페이지의 '품목추가' 팝업) ──────────────────
// 허용된 시트·열에만 기록한다. 나머지 열(재고 등)은 update_sheets_inventory.py 가 채운다.
const APPEND_RULES = {
  '사양변경': {
    fields:   ['구사양코드', '신사양코드', '색상', '사양변경내용', '단품명칭', '관련문서'],
    required: ['구사양코드', '신사양코드', '색상', '사양변경내용', '단품명칭', '관련문서'],
    codes:    ['구사양코드', '신사양코드', '색상'],
    dupKeys:  ['구사양코드', '신사양코드', '색상'],
    common:   ['사양변경내용', '관련문서'],
  },
  '단종(단품)': {
    fields:   ['구분', '단품코드', '색상', '발의년도'],
    required: ['구분', '단품코드', '색상'],
    codes:    ['단품코드', '색상'],
    dupKeys:  ['단품코드', '색상'],
    common:   ['구분'],
    auto:     { '발의년도': function () { return new Date().getFullYear(); } },
  },
};
const APPEND_MAX_ROWS = 50;
const APPEND_MAX_LEN = 500;

function handleAppendRows(body) {
  const rule = APPEND_RULES[body.sheet];
  if (!rule) return respond({ status: 'error', msg: '품목추가를 지원하지 않는 시트: ' + body.sheet });
  const common = body.common || {};
  const input = Array.isArray(body.rows) ? body.rows : [];
  if (!input.length) return respond({ status: 'error', msg: '등록할 행이 없습니다' });
  if (input.length > APPEND_MAX_ROWS) return respond({ status: 'error', msg: '한 번에 최대 ' + APPEND_MAX_ROWS + '행까지 등록할 수 있습니다' });

  // 정규화 + 필수값/형식 검사
  const rows = [];
  for (let i = 0; i < input.length; i++) {
    const r = {};
    for (const f of rule.fields) {
      const src = rule.common.indexOf(f) >= 0 ? common[f] : (input[i] || {})[f];
      let v = String(src == null ? '' : src).trim().slice(0, APPEND_MAX_LEN);
      if (rule.codes.indexOf(f) >= 0) {
        v = v.toUpperCase();
        if (v && !/^[A-Z0-9\-_.]+$/.test(v)) return respond({ status: 'error', msg: (i + 1) + '행 ' + f + ' 형식 오류: ' + v });
      }
      if (rule.auto && rule.auto[f]) v = rule.auto[f]();
      r[f] = v;
    }
    for (const f of rule.required) {
      if (!r[f]) return respond({ status: 'error', msg: (i + 1) + '행 ' + f + ' 값이 비어 있습니다' });
    }
    if (r['관련문서'] && !/^https?:\/\/\S+$/i.test(r['관련문서'])) {
      return respond({ status: 'error', msg: '문서링크는 http(s):// 주소여야 합니다' });
    }
    rows.push(r);
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return respond({ status: 'error', msg: '다른 등록이 진행 중입니다. 잠시 후 다시 시도하세요' });
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const ws = ss.getSheetByName(body.sheet);
    if (!ws) return respond({ status: 'error', msg: 'sheet not found: ' + body.sheet });

    const all = ws.getDataRange().getValues();
    const headers = all[0].map(function (h) { return String(h).trim(); });
    const col = {};
    for (const f of rule.fields) {
      col[f] = headers.indexOf(f);
      if (col[f] < 0) return respond({ status: 'error', msg: '시트에 열이 없습니다: ' + f });
    }

    // 중복 검사 (시트 기존 행 + 요청 내 중복)
    const keyOf = function (get) { return rule.dupKeys.map(function (k) { return String(get(k)).trim().toUpperCase(); }).join('|'); };
    const existing = {};
    for (let i = 1; i < all.length; i++) existing[keyOf(function (k) { return all[i][col[k]]; })] = true;
    const dups = [];
    const seen = {};
    for (const r of rows) {
      const k = keyOf(function (f) { return r[f]; });
      if (existing[k] || seen[k]) dups.push(rule.dupKeys.map(function (f) { return r[f]; }).join(' / '));
      seen[k] = true;
    }
    if (dups.length) return respond({ status: 'error', msg: '이미 등록된 항목이 있어 등록하지 않았습니다', dups: dups });

    // 마지막 데이터 행 (기록 대상 열 기준) 다음부터 기록
    let last = all.length;
    while (last > 1 && rule.fields.every(function (f) { return String(all[last - 1][col[f]] || '').trim() === ''; })) last--;

    const written = [];
    rows.forEach(function (r, i) {
      const rowNum = last + 1 + i;
      for (const f of rule.fields) {
        const cell = ws.getRange(rowNum, col[f] + 1);
        if (rule.codes.indexOf(f) >= 0) cell.setNumberFormat('@');   // 색상 '000' 등 숫자 변환 방지
        cell.setValue(r[f]);
      }
      written.push(rowNum);
    });
    SpreadsheetApp.flush();
    console.log('appendRows', body.sheet, 'user=' + (body.user || '-'), 'rows=' + written.join(','));
    return respond({ status: 'ok', added: written.length, rows: written });
  } finally {
    lock.releaseLock();
  }
}

function getRequestSheet(ss) {
  let ws = ss.getSheetByName(REQUEST_SHEET_NAME);
  if (!ws) {
    ws = ss.insertSheet(REQUEST_SHEET_NAME);
    ws.appendRow(['EP ID', '신청일시', '상태', '처리일시']);
  }
  return ws;
}

function handleRequestAccess(body) {
  const epId = String(body.epId || '').trim();
  if (!epId) return respond({ status: 'error', msg: 'epId required' });

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const ws = getRequestSheet(ss);
  const all = ws.getDataRange().getValues();

  for (let i = 1; i < all.length; i++) {
    if (String(all[i][0] || '').trim() === epId && String(all[i][2] || '').trim() === '대기') {
      return respond({ status: 'ok', msg: 'already pending' });
    }
  }

  ws.appendRow([epId, new Date(), '대기', '']);
  SpreadsheetApp.flush();
  return respond({ status: 'ok' });
}

function handleReviewAccess(body, decision) {
  const epId = String(body.epId || '').trim();
  if (!epId) return respond({ status: 'error', msg: 'epId required' });

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const ws = getRequestSheet(ss);
  const all = ws.getDataRange().getValues();

  for (let i = all.length - 1; i >= 1; i--) {
    if (String(all[i][0] || '').trim() === epId && String(all[i][2] || '').trim() === '대기') {
      ws.getRange(i + 1, 3).setValue(decision);
      ws.getRange(i + 1, 4).setValue(new Date());
      SpreadsheetApp.flush();
      return respond({ status: 'ok', row: i + 1 });
    }
  }
  return respond({ status: 'error', msg: 'pending request not found' });
}

function respond(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

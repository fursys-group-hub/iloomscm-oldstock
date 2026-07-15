const SPREADSHEET_ID = '1sAvWHksA3hhlDRvGsXqHdbSvO07MFTCrebuVoRk9LSc';
const REQUEST_SHEET_NAME = '권한관리';

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);

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

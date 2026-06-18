const SPREADSHEET_ID = '1sAvWHksA3hhlDRvGsXqHdbSvO07MFTCrebuVoRk9LSc';

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
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

function respond(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

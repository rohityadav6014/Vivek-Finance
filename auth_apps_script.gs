// Vivek Finance — Auth proxy
// Deploy as Web App. Execute as: Me. Who has access: Anyone.
// Reads the first sheet of the spreadsheet referenced by SHEET_ID.
// Matches columns flexibly by header (case-insensitive). It will prefer
// a user-id column when present, and fall back to an email column.
//   user id  / userid / user_id / username / login / id
//   email    / mail   / e-mail  / emailid  / email id
//   password / passwords / pwd / pass
//   active   / enabled / status   (optional; blank/TRUE/yes = active)

const SHEET_ID = '1taIJ5YyrJAbEJ8vLGqGScyrzWbtEeeJ_Zx_nktAn85Y';

function doPost(e) { return handle(e); }
function doGet(e)  { return handle(e); }

function handle(e) {
  try {
    const params = (e && e.parameter) || {};
    // accept either userid or legacy email param
    const userId = String(params.userid || params.userId || params.user_id || params.email || '').trim().toLowerCase();
    const password = String(params.password || '');
    if (!userId || !password) return json({ ok: false, error: 'missing_credentials' });

    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return json({ ok: false, error: 'no_users' });

    const headers = data[0].map(function (h) { return String(h).trim().toLowerCase(); });
    const userIdCol   = findCol(headers, ['user id', 'userid', 'user_id', 'username', 'login', 'id']);
    const emailCol    = findCol(headers, ['email', 'mail', 'e-mail', 'emailid', 'email id']);
    const passwordCol = findCol(headers, ['password', 'passwords', 'pwd', 'pass']);
    const activeCol   = findCol(headers, ['active', 'enabled', 'status']);

    // Prefer user-id column when present; otherwise authenticate against email column.
    const identityCol = userIdCol >= 0 ? userIdCol : emailCol;
    if (identityCol < 0 || passwordCol < 0) return json({ ok: false, error: 'sheet_misconfigured' });

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var rowIdentity = String(row[identityCol] || '').trim().toLowerCase();
      var rowPassword = String(row[passwordCol] || '');
      if (rowIdentity !== userId || rowPassword !== password) continue;

      if (activeCol >= 0) {
        var raw = row[activeCol];
        var v = String(raw).trim().toLowerCase();
        var isActive = raw === true || raw === 1 || v === '' || v === 'true' || v === 'yes' || v === 'y' || v === '1' || v === 'active';
        if (!isActive) return json({ ok: false, error: 'inactive' });
      }
      return json({ ok: true, userId: rowIdentity });
    }
    return json({ ok: false, error: 'invalid_credentials' });
  } catch (err) {
    return json({ ok: false, error: 'server_error', detail: String(err) });
  }
}

function findCol(headers, names) {
  for (var i = 0; i < names.length; i++) {
    var idx = headers.indexOf(names[i]);
    if (idx >= 0) return idx;
  }
  return -1;
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Google Apps Script — เชื่อม POS กับ Google Sheets + Google Drive
// วิธีใช้: เปิด Google Sheet > ส่วนขยาย > Apps Script > วางโค้ดนี้ > Deploy > New deployment
// ประเภท Web app | Execute as: Me | Who has access: Anyone > คัดลอก URL ไปใส่ที่ API_URL ใน pos.html
const SS = '11txDJTb_523YyQbqg_IpUYp2b16E-PGOg9DR9xacaB0';
const FOLDER = '1simptlAElXaWscjR9g9F1fxRPWfB7CDr';
const PH = ['id','name','cat','price','cost','stock','unit','img'];
const SH = ['id','date','items','total','pay','received'];

function sh(n, h) {
  const s = SpreadsheetApp.openById(SS);
  const x = s.getSheetByName(n) || s.insertSheet(n);
  if (!x.getLastRow()) x.appendRow(h);
  return x;
}
function rows(x, h) {
  const v = x.getDataRange().getValues(); v.shift();
  return v.map(r => Object.fromEntries(h.map((k, i) => [k, r[i]])));
}
function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function doGet() {
  const p = rows(sh('Products', PH), PH);
  const s = rows(sh('Sales', SH), SH).map(r => {
    try { r.items = JSON.parse(r.items); } catch (e) { r.items = []; }
    return r;
  });
  return out({ products: p, sales: s });
}
function writeProducts(list) {
  const x = sh('Products', PH);
  x.clearContents(); x.appendRow(PH);
  if (list.length) x.getRange(2, 1, list.length, PH.length)
    .setValues(list.map(p => PH.map(k => p[k] === undefined ? '' : p[k])));
}
function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  if (d.action == 'products') writeProducts(d.products);
  if (d.action == 'sale') {
    const s = d.sale;
    sh('Sales', SH).appendRow([s.id, s.date, JSON.stringify(s.items), s.total, s.pay, s.received]);
    writeProducts(d.products);
  }
  if (d.action == 'ping') {
    let f = '';
    try { f = DriveApp.getFolderById(FOLDER).getName(); } catch (err) {}
    return out({ ok: 1, sheet: SpreadsheetApp.openById(SS).getName(), folder: f });
  }
  if (d.action == 'image') {
    const f = DriveApp.getFolderById(FOLDER).createFile(Utilities.newBlob(Utilities.base64Decode(d.data), d.mime, d.name));
    f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return out({ url: 'https://drive.google.com/thumbnail?id=' + f.getId() + '&sz=w400' });
  }
  return out({ ok: 1 });
}

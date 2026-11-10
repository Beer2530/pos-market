// Google Apps Script — POS: Google Sheets + Google Drive (รองรับหลายเครื่องพร้อมกัน)
// วิธีใช้: วางโค้ดนี้ทับของเดิมใน Apps Script > Deploy > Manage deployments > Edit > New version > Deploy
const SS = '11txDJTb_523YyQbqg_IpUYp2b16E-PGOg9DR9xacaB0';
const FOLDER = '1simptlAElXaWscjR9g9F1fxRPWfB7CDr';
const PH = ['id','name','cat','price','cost','stock','unit','img'];
const SH = ['id','date','items','total','pay','received'];
const MH = ['id','date','pid','name','type','qty','after','note'];

function sh(n, h) {
  const s = SpreadsheetApp.openById(SS);
  const x = s.getSheetByName(n) || s.insertSheet(n);
  if (!x.getLastRow()) { x.getRange('B:B').setNumberFormat(h[1] == 'date' ? '@' : 'General'); x.appendRow(h); }
  return x;
}
function rows(x, h) {
  const v = x.getDataRange().getValues(); v.shift();
  return v.map(r => Object.fromEntries(h.map((k, i) => [k, r[i]])));
}
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function ids(x) {
  const n = x.getLastRow();
  return n < 2 ? [] : x.getRange(2, 1, n - 1, 1).getValues().map(r => String(r[0]));
}
function find(x, id) { const i = ids(x).indexOf(String(id)); return i < 0 ? -1 : i + 2; }
function stock(x, id, delta) {
  const r = find(x, id); if (r < 0) return '';
  const c = x.getRange(r, 6), n = (+c.getValue() || 0) + delta; c.setValue(n); return n;
}
function applyDiff(px, oldI, newI, type) {
  const m = {}, names = {};
  oldI.forEach(i => { m[i.id] = (m[i.id] || 0) + i.qty; names[i.id] = i.name; });
  newI.forEach(i => { m[i.id] = (m[i.id] || 0) - i.qty; names[i.id] = i.name; });
  const mx = sh('StockLog', MH);
  Object.keys(m).forEach(id => {
    if (!m[id]) return;
    const after = stock(px, id, m[id]);
    mx.appendRow(['M' + Date.now() + id, new Date().toISOString(), id, names[id], type, m[id], after, '']);
  });
}
function doGet() {
  const s = rows(sh('Sales', SH), SH).map(r => { try { r.items = JSON.parse(r.items); } catch (e) { r.items = []; } return r; });
  return out({ products: rows(sh('Products', PH), PH), sales: s, moves: rows(sh('StockLog', MH), MH).slice(-200) });
}
function doPost(e) {
  const d = JSON.parse(e.postData.contents);
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
  const lock = LockService.getScriptLock(); lock.waitLock(25000);
  try {
    const px = sh('Products', PH);
    if (d.action == 'seed' && px.getLastRow() < 2 && d.products.length)
      px.getRange(2, 1, d.products.length, PH.length).setValues(d.products.map(p => PH.map(k => p[k] === undefined ? '' : p[k])));
    if (d.action == 'sale') {            // ตัดสต็อกฝั่งเซิร์ฟเวอร์ ป้องกันหลายเครื่องเขียนทับกัน
      const sx = sh('Sales', SH), s = d.sale;
      if (ids(sx).indexOf(String(s.id)) < 0) {
        sx.appendRow([s.id, s.date, JSON.stringify(s.items), s.total, s.pay, s.received]);
        s.items.forEach(i => stock(px, i.id, -i.qty));
      }
    }
    if (d.action == 'editSale') {        // แก้ไขบิล: คำนวณส่วนต่างสต็อกจากบิลเดิมในชีต (ส่งซ้ำก็ไม่ซ้ำซ้อน)
      const sx = sh('Sales', SH), s = d.sale, r = find(sx, s.id);
      if (r > 0) {
        let old = []; try { old = JSON.parse(sx.getRange(r, 3).getValue() || '[]'); } catch (e) {}
        applyDiff(px, old, s.items, 'แก้ไขบิล');
        sx.getRange(r, 1, 1, 6).setValues([[s.id, s.date, JSON.stringify(s.items), s.total, s.pay, s.received]]);
      }
    }
    if (d.action == 'deleteSale') {      // ยกเลิกบิล: คืนสต็อกแล้วลบแถว
      const sx = sh('Sales', SH), r = find(sx, d.id);
      if (r > 0) {
        let old = []; try { old = JSON.parse(sx.getRange(r, 3).getValue() || '[]'); } catch (e) {}
        applyDiff(px, old, [], 'ยกเลิกบิล');
        sx.deleteRow(r);
      }
    }
    if (d.action == 'adjust') {          // เพิ่ม/ตัดสต็อกด้วยตนเอง + บันทึกประวัติ
      const mx = sh('StockLog', MH), m = d.move;
      if (ids(mx).indexOf(String(m.id)) < 0) {
        const after = stock(px, m.pid, m.qty);
        mx.appendRow([m.id, m.date, m.pid, m.name, m.type, m.qty, after, m.note]);
      }
    }
    if (d.action == 'upsert') {          // แก้ไขข้อมูลสินค้า (ไม่แตะสต็อกของสินค้าเดิม)
      const p = d.product, r = find(px, p.id), v = PH.map(k => p[k] === undefined ? '' : p[k]);
      if (r < 0) px.appendRow(v);
      else { v[5] = px.getRange(r, 6).getValue(); px.getRange(r, 1, 1, PH.length).setValues([v]); }
    }
    if (d.action == 'delete') { const r = find(px, d.id); if (r > 0) px.deleteRow(r); }
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }
  return out({ ok: 1 });
}

// Google Apps Script — POS: Google Sheets + Google Drive (รองรับหลายเครื่อง และเร็วขึ้น)
// วิธีใช้: วางโค้ดนี้ทับของเดิมใน Apps Script > Deploy > Manage deployments > Edit > New version > Deploy
const SS = '11txDJTb_523YyQbqg_IpUYp2b16E-PGOg9DR9xacaB0';
const FOLDER = '1simptlAElXaWscjR9g9F1fxRPWfB7CDr';
const PH = ['id','name','cat','price','cost','stock','unit','img','dec'];   // dec = ขายเป็นทศนิยมได้ (1/0)
const SH = ['id','date','items','total','pay','received'];
const CH = ['name','img'];                // หมวดหมู่ + รูป (แท็บ Categories)
const MH = ['id','date','pid','name','type','qty','after','note'];

let _ss;
function ss() { return _ss || (_ss = SpreadsheetApp.openById(SS)); }
function sh(n, h) {
  let x = ss().getSheetByName(n);
  if (!x) x = ss().insertSheet(n);
  if (!x.getLastRow()) { x.getRange('B:B').setNumberFormat(h[1] == 'date' ? '@' : 'General'); x.appendRow(h); }
  else if (x.getLastColumn() < h.length) x.getRange(1, 1, 1, h.length).setValues([h]);   // เติมหัวคอลัมน์ใหม่ (dec)
  return x;
}
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

// เลขเวอร์ชันข้อมูล: เพิ่มทุกครั้งที่มีการเขียน ให้แอปถามเลขนี้ (เบามาก) ก่อน แล้วค่อยโหลดข้อมูลเต็มเมื่อเปลี่ยน
function ver() { return +PropertiesService.getScriptProperties().getProperty('ver') || 0; }
function bump() { const p = PropertiesService.getScriptProperties(), v = ver() + 1; p.setProperty('ver', String(v)); return v; }

function rows(x, h, from) {
  const n = x.getLastRow(), a = Math.max(2, from || 2);
  if (n < 2 || a > n) return [];
  return x.getRange(a, 1, n - a + 1, h.length).getValues().map(r => { const o = {}; h.forEach((k, i) => o[k] = r[i]); return o; });
}
function ids(x, k) {                      // อ่านเฉพาะคอลัมน์ id (k = เอาเฉพาะแถวล่าสุด k แถว)
  const n = x.getLastRow(), a = k ? Math.max(2, n - k + 1) : 2;
  return n < 2 ? [] : x.getRange(a, 1, n - a + 1, 1).getValues().map(r => String(r[0]));
}
function find(x, id) { const i = ids(x).indexOf(String(id)); return i < 0 ? -1 : i + 2; }

// ---- สต็อก: อ่านครั้งเดียว แก้ในหน่วยความจำ เขียนกลับครั้งเดียว ----
function loadStock(px) {
  const n = px.getLastRow(), v = n < 2 ? [] : px.getRange(2, 1, n - 1, 6).getValues(), idx = {};
  v.forEach((r, i) => idx[String(r[0])] = i);
  return { v, idx, dirty: false };
}
function bumpStock(c, id, delta) {
  const i = c.idx[String(id)]; if (i === undefined) return '';
  c.v[i][5] = Math.round(((+c.v[i][5] || 0) + delta) * 1000) / 1000; c.dirty = true; return c.v[i][5];
}
function saveStock(px, c) { if (c.dirty) px.getRange(2, 6, c.v.length, 1).setValues(c.v.map(r => [r[5]])); }
function applyDiff(px, oldI, newI, type) {
  const m = {}, nm = {};
  oldI.forEach(i => { m[i.id] = (m[i.id] || 0) + i.qty; nm[i.id] = i.name; });
  newI.forEach(i => { m[i.id] = (m[i.id] || 0) - i.qty; nm[i.id] = i.name; });
  const c = loadStock(px), mx = sh('StockLog', MH), now = new Date().toISOString();
  const log = [];
  Object.keys(m).forEach(id => { m[id] = Math.round(m[id] * 1000) / 1000; if (m[id]) log.push(['M' + Date.now() + id, now, id, nm[id], type, m[id], bumpStock(c, id, m[id]), '']); });
  saveStock(px, c);
  log.forEach(r => mx.appendRow(r));
}

function doGet(e) {
  if (e && e.parameter && e.parameter.v) return out({ ver: ver() });
  const v = ver();
  const sales = rows(sh('Sales', SH), SH).map(r => { try { r.items = JSON.parse(r.items); } catch (x) { r.items = []; } return r; });
  const mx = sh('StockLog', MH);
  return out({ ver: v, products: rows(sh('Products', PH), PH), sales: sales, moves: rows(mx, MH, mx.getLastRow() - 49), cats: rows(sh('Categories', CH), CH) });
}

function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  if (d.action == 'ping') {
    let f = '';
    try { f = DriveApp.getFolderById(FOLDER).getName(); } catch (err) {}
    return out({ ok: 1, sheet: ss().getName(), folder: f, ver: ver() });
  }
  if (d.action == 'image') {
    const f = DriveApp.getFolderById(FOLDER).createFile(Utilities.newBlob(Utilities.base64Decode(d.data), d.mime, d.name));
    f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return out({ url: 'https://drive.google.com/thumbnail?id=' + f.getId() + '&sz=w400' });
  }
  const lock = LockService.getScriptLock(); lock.waitLock(25000);
  let v = ver();
  try {
    const px = sh('Products', PH);
    if (d.action == 'seed' && px.getLastRow() < 2 && d.products.length) {
      px.getRange(2, 1, d.products.length, PH.length).setValues(d.products.map(p => PH.map(k => p[k] === undefined ? '' : p[k])));
      v = bump();
    }
    if (d.action == 'sale') {              // ตัดสต็อกฝั่งเซิร์ฟเวอร์ ป้องกันหลายเครื่องเขียนทับกัน
      const sx = sh('Sales', SH), s = d.sale;
      if (ids(sx, 300).indexOf(String(s.id)) < 0) {
        sx.appendRow([s.id, s.date, JSON.stringify(s.items), s.total, s.pay, s.received]);
        const c = loadStock(px); s.items.forEach(i => bumpStock(c, i.id, -i.qty)); saveStock(px, c);
        v = bump();
      }
    }
    if (d.action == 'editSale') {          // คำนวณส่วนต่างสต็อกจากบิลเดิมในชีต (ส่งซ้ำก็ไม่ซ้ำซ้อน)
      const sx = sh('Sales', SH), s = d.sale, r = find(sx, s.id);
      if (r > 0) {
        let old = []; try { old = JSON.parse(sx.getRange(r, 3).getValue() || '[]'); } catch (x) {}
        applyDiff(px, old, s.items, 'แก้ไขบิล');
        sx.getRange(r, 1, 1, 6).setValues([[s.id, s.date, JSON.stringify(s.items), s.total, s.pay, s.received]]);
        v = bump();
      }
    }
    if (d.action == 'deleteSale') {        // ยกเลิกบิล: คืนสต็อกแล้วลบแถว
      const sx = sh('Sales', SH), r = find(sx, d.id);
      if (r > 0) {
        let old = []; try { old = JSON.parse(sx.getRange(r, 3).getValue() || '[]'); } catch (x) {}
        applyDiff(px, old, [], 'ยกเลิกบิล');
        sx.deleteRow(r);
        v = bump();
      }
    }
    if (d.action == 'adjust') {            // เพิ่ม/ตัดสต็อกด้วยตนเอง + บันทึกประวัติ
      const mx = sh('StockLog', MH), m = d.move;
      if (ids(mx, 300).indexOf(String(m.id)) < 0) {
        const c = loadStock(px), after = bumpStock(c, m.pid, m.qty); saveStock(px, c);
        mx.appendRow([m.id, m.date, m.pid, m.name, m.type, m.qty, after, m.note]);
        v = bump();
      }
    }
    if (d.action == 'upsert') {            // แก้ไขข้อมูลสินค้า (ไม่แตะสต็อกของสินค้าเดิม)
      const p = d.product, r = find(px, p.id), vals = PH.map(k => p[k] === undefined ? '' : p[k]);
      if (r < 0) px.appendRow(vals);
      else { vals[5] = px.getRange(r, 6).getValue(); px.getRange(r, 1, 1, PH.length).setValues([vals]); }
      v = bump();
    }
    if (d.action == 'cat') {               // บันทึกรูปหมวดหมู่
      const cx = sh('Categories', CH), c = d.cat, r = find(cx, c.name), vals = [c.name, c.img || ''];
      if (r < 0) cx.appendRow(vals); else cx.getRange(r, 1, 1, 2).setValues([vals]);
      v = bump();
    }
    if (d.action == 'delete') { const r = find(px, d.id); if (r > 0) { px.deleteRow(r); v = bump(); } }
  } finally { lock.releaseLock(); }
  return out({ ok: 1, ver: v });
}

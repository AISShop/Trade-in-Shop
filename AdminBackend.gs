/*************************************************************
 * AdminBackend.gs — Staff Registry API for AIS Admin Dashboard
 * Sheet: 1mMnLp3Kh4nQ-fGpQblWdUwCJM_MEhI4B4S-F0HYNSeQ
 * Tabs : Staff_AIS (10 cols) , Staff_PC (11 cols)
 *
 * Staff_AIS (A-J):
 *   A ID | B LineID | C Username | D PIN | E Branch |
 *   F Active | G CreatedAt | H UpdatedAt | I PictureUrl | J UID
 *
 * Staff_PC (A-K):
 *   A ID | B LineID | C Name | D PIN | E Brand | F Branch |
 *   G Active | H CreatedAt | I UpdatedAt | J PictureUrl | K UID
 *
 * Notes:
 *   - LineID  = LINE displayName of the registrant
 *   - UID     = real LINE userId (Uxxxxxxxx..., long string) -> dedupe key
 *   - AIS uses "Username", PC uses "Name" (same logical name field)
 *   - PC has extra "Brand" column
 *   - SOFT DELETE: remove() sets Active=FALSE and marks the row deleted
 *     (status='deleted'). Row is NEVER physically removed. restore()
 *     brings it back. list() hides deleted rows unless includeDeleted=1.
 *     Deletion state is tracked in an optional "Deleted" column, which
 *     is auto-detected from the header row; if the column does not exist
 *     the code adds it automatically on first soft-delete.
 *
 * Transport: JSONP via doGet (?action=...&cb=...) to match the
 * existing Dashboard pattern (avoids CORS).
 *************************************************************/

var SHEET_ID = '1mMnLp3Kh4nQ-fGpQblWdUwCJM_MEhI4B4S-F0HYNSeQ';

// Per-tab column maps (0-based). "name" is the logical name field
// (Username for AIS / Name for PC). "brand" only exists on PC.
var SCHEMA = {
  ais: {
    tab: 'Staff_AIS',
    ncols: 10,
    col: { ID:0, LINEID:1, NAME:2, PIN:3, BRANCH:4,
           ACTIVE:5, CREATED:6, UPDATED:7, PICTURE:8, UID:9 }
  },
  pc: {
    tab: 'Staff_PC',
    ncols: 11,
    col: { ID:0, LINEID:1, NAME:2, PIN:3, BRAND:4, BRANCH:5,
           ACTIVE:6, CREATED:7, UPDATED:8, PICTURE:9, UID:10 }
  }
};

// ---- AIS Pro lookup (separate tab, matched by UserName + PIN) ----
// Tab AIS_Pro: A Location | B UserName | C PINEmployee | D Brand Pro
var AISPRO_TAB = 'AIS_Pro';
var AISPRO_COL = { LOCATION:0, USERNAME:1, PIN:2, BRANDPRO:3 };

// ---- Mgr/Sup exclusion list (Tab Mgr_Sup) — พนักงานที่ไม่นับใน Performance ----
var MGRSUP_TAB = 'Mgr_Sup';

// return 'Apple' | 'Samsung' | '' by matching username + pin
function aisProType(username, pin) {
  var u = S(username).replace(/\s+/g,' ').trim().toUpperCase();
  var p = S(pin).trim();
  if (!u || !p) return '';
  try {
    var data = getAisProData_();  // cached
    for (var i = 0; i < data.length; i++) {
      if (data[i].u === u && data[i].p === p) return data[i].pro;
    }
  } catch (e) {}
  return '';
}

// อ่านชีท AIS_Pro ครั้งเดียว แล้ว cache 6 ชม. (เร็วขึ้นมากสำหรับ whoami)
function getAisProData_() {
  var cache = CacheService.getScriptCache();
  var hit = cache.get('aispro_v1');
  if (hit) { try { return JSON.parse(hit); } catch(e) {} }
  var out = [];
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(AISPRO_TAB);
  if (sh) {
    var v = sh.getDataRange().getValues();
    for (var i = 1; i < v.length; i++) {
      var ru = S(v[i][AISPRO_COL.USERNAME]).replace(/\s+/g,' ').trim().toUpperCase();
      var rp = S(v[i][AISPRO_COL.PIN]).trim();
      if (!ru) continue;
      var bp = S(v[i][AISPRO_COL.BRANDPRO]).toLowerCase();
      var pro = (bp.indexOf('apple') >= 0) ? 'Apple'
              : ((bp.indexOf('samsung')>=0||bp.indexOf('sasmung')>=0||bp.indexOf('sam')>=0||bp.indexOf('sas')>=0) ? 'Samsung'
              : S(v[i][AISPRO_COL.BRANDPRO]));
      out.push({ u:ru, p:rp, pro:pro });
    }
  }
  try { cache.put('aispro_v1', JSON.stringify(out), 21600); } catch(e) {} // 6 ชม.
  return out;
}

// ล้าง cache ที่เกี่ยวกับพนักงาน (เรียกหลังแก้ทะเบียน)
function clearStaffCache_(p) {
  try {
    var cache = CacheService.getScriptCache();
    cache.remove('aispro_v1');
    var uid = S(p && p.uid);
    if (uid) cache.remove('who_' + uid);
  } catch (e) {}
}

// ---- admin auth (server-side only) ----
var ADMIN_USER = 'admin';
var ADMIN_PASS = 'ais2026';   // <- change before go-live

// ====================== entry =========================
function doGet(e) {
  var p  = (e && e.parameter) || {};
  var cb = p.cb || '';
  try {
    var out;
    switch (p.action) {
      case 'ping':     out = { ok:true, msg:'AdminBackend alive', ts:nowISO() }; break;
      case 'login':    out = login(p); break;
      case 'list':     out = listStaff(p); break;
      case 'whoami':   out = whoami(p); break;
      case 'proList':  out = proList(); break;
      case 'mgrSupList': out = mgrSupList(); break;
      case 'updatePhoto': out = updatePhotoStaff(p); clearStaffCache_(p); break;
      case 'register': out = registerStaff(p); clearStaffCache_(p); break;
      case 'update':   out = updateStaff(p); clearStaffCache_(p); break;
      case 'toggle':   out = toggleActive(p); clearStaffCache_(p); break;
      case 'remove':   out = removeStaff(p); clearStaffCache_(p); break;
      case 'restore':  out = restoreStaff(p); clearStaffCache_(p); break;
      case 'purge':    out = purgeStaff(p); clearStaffCache_(p); break;
      case 'rosUpload':  out = rosUpload(p); break;
      case 'rosUploadFile': out = rosUploadFile(p); break;
      case 'rosSummary': out = rosSummary(p); break;
      case 'rosBatches': out = rosBatches(); break;
      case 'rosClear':   out = rosClear(p); break;
      case 'rosRebuild': out = rosRebuild(p); break;
      case 'rosFix':     out = rosFix(); break;
      case 'rosModels':    out = rosModels(); break;
      case 'rosSetModels': out = rosSetModels(p); break;
      case 'linkupData':   out = luData(p.month || ''); break;
      case 'linkupBundle': out = luBundle(p.month || ''); break;
      case 'quickpayData':   out = qpData(p.month || ''); break;
      case 'quickpayBundle': out = qpBundle(p.month || ''); break;
      case 'careplusData':   out = cpData(p.month || ''); break;
      case 'careplusBundle': out = cpBundle(p.month || ''); break;
      case 'tabletData':     out = tbData(p.month || ''); break;
      case 'tabletBundle':   out = tbBundle(p.month || ''); break;
      case 'tradeinTeams':   out = tradeinTeams(); break;
      case 'careplusTeams':  out = careplusTeams(); break;
      case 'linkupTeams':    out = linkupTeams(); break;
      case 'financeTeams':   out = financeTeams(); break;
      case 'linkupRebuild': out = luRebuildCache_(); break;
      case 'finUpload':    out = finUpload(p); break;
      case 'finUploadFile':out = finUploadFile(p); break;
      case 'finSummary':   out = finSummary(p); break;
      case 'finBatches':   out = finBatches(); break;
      case 'finUsers':     out = finUsers(p); break;
      case 'finClear':     out = finClear(p); break;
      case 'finRebuild':   out = finRebuild(p); break;
      case 'finPeek':      out = finPeek(p); break;
      default:         out = { ok:false, error:'unknown action: '+(p.action||'(none)') };
    }
    return respond(out, cb);
  } catch (err) {
    return respond({ ok:false, error:String(err && err.message || err) }, cb);
  }
}

// POST: merge JSON body into parameters (for bulk ROS upload)
function doPost(e) {
  try {
    if (e && e.postData && e.postData.contents) {
      var body = JSON.parse(e.postData.contents);
      e.parameter = e.parameter || {};
      for (var k in body) e.parameter[k] = body[k];
    }
  } catch (err) {}
  return doGet(e);
}

// ====================== actions =======================
function login(p) {
  if (S(p.user) === ADMIN_USER && S(p.pass) === ADMIN_PASS)
    return { ok:true, token:'ok' };
  return { ok:false, error:'invalid credentials' };
}

// mgrSupList -> รายชื่อ Mgr/Sup ที่ไม่นับใน Performance
// ชีทจริง: A=PIN, B=UserName
function mgrSupList() {
  var out = [];
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = ss.getSheetByName(MGRSUP_TAB);
    if (!sh) return { ok:true, data:[] };
    var v = sh.getDataRange().getValues();
    for (var i = 1; i < v.length; i++) {
      var pin  = S(v[i][0]).trim();                                   // A = PIN
      var name = S(v[i][1]).replace(/\s+/g,' ').trim().toUpperCase(); // B = UserName
      if (!name && !pin) continue;
      out.push({ name:name, pin:pin });
    }
  } catch (e) {}
  return { ok:true, data:out };
}

// proList -> full AIS_Pro roster for dashboard cross-reference
function proList() {
  var out = [];
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = ss.getSheetByName(AISPRO_TAB);
    if (!sh) return { ok:true, data:[] };
    var v = sh.getDataRange().getValues();
    for (var i = 1; i < v.length; i++) {
      var name = S(v[i][AISPRO_COL.USERNAME]).replace(/\s+/g,' ').trim().toUpperCase();
      if (!name) continue;
      var bp = S(v[i][AISPRO_COL.BRANDPRO]).toLowerCase();
      var pro = bp.indexOf('apple') >= 0 ? 'Apple' : ((bp.indexOf('samsung')>=0||bp.indexOf('sasmung')>=0||bp.indexOf('sam')>=0||bp.indexOf('sas')>=0) ? 'Samsung' : S(v[i][AISPRO_COL.BRANDPRO]));
      out.push({ name:name, pin:S(v[i][AISPRO_COL.PIN]), branch:S(v[i][AISPRO_COL.LOCATION]), pro:pro });
    }
  } catch (e) {}
  return { ok:true, data:out };
}

// whoami?uid=...  -> หาเฉพาะ uid เดียว (cache + อ่านดิบ หา uid ตรงๆ)
function whoami(p) {
  var uid = S(p.uid);
  if (!uid) return { ok:true, found:null };
  var cache = CacheService.getScriptCache();
  var ck = 'who_' + uid;
  var hit = cache.get(ck);
  if (hit) { try { return JSON.parse(hit); } catch(e) {} }
  var res = whoamiLookup_(uid);
  try { cache.put(ck, JSON.stringify(res), 600); } catch(e) {} // 10 นาที
  return res;
}
function whoamiLookup_(uid) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var types = ['ais','pc'];
  for (var t = 0; t < types.length; t++) {
    var sc = SCHEMA[types[t]];
    var sh = ss.getSheetByName(sc.tab);
    if (!sh) continue;
    var v = sh.getDataRange().getValues();
    var c = sc.col;
    var dIdx = deletedColIndex(v[0]);
    for (var i = 1; i < v.length; i++) {
      if (S(v[i][c.UID]) !== uid) continue;
      if (S(v[i][c.ID]) === '') continue;
      if (dIdx >= 0 && S(v[i][dIdx]).toUpperCase() === 'TRUE') continue;
      if (S(v[i][c.ACTIVE]).toUpperCase() !== 'TRUE') continue;
      var rec = {
        type: types[t].toUpperCase(),
        name: S(v[i][c.NAME]),
        branch: S(v[i][c.BRANCH]),
        brand: (c.BRAND !== undefined) ? S(v[i][c.BRAND]) : ''
      };
      if (types[t] === 'ais') rec.aisPro = aisProType(rec.name, S(v[i][c.PIN]));
      return { ok:true, found:rec };
    }
  }
  return { ok:true, found:null };
}

// updatePhoto?uid=...&picture=...  -> อัปเดตรูปเบื้องหลัง (แยกจาก whoami เพื่อไม่ให้ช้า)
function updatePhotoStaff(p) {
  var uid = S(p.uid), pic = S(p.picture);
  if (!uid || !pic) return { ok:true, updated:false };
  var types = ['ais','pc'];
  for (var t = 0; t < types.length; t++) {
    var rows = readTab(types[t], false);
    for (var i = 0; i < rows.length; i++) {
      if (S(rows[i].uid) === uid) {
        if (pic !== S(rows[i].picture)) {
          try { updatePhotoCell_(types[t], uid, pic); return { ok:true, updated:true }; } catch(e) {}
        }
        return { ok:true, updated:false };
      }
    }
  }
  return { ok:true, updated:false };
}

// อัปเดตเซลล์ PictureUrl ของพนักงานที่ตรง uid
function updatePhotoCell_(type, uid, picture) {
  var sc = SCHEMA[type];
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(sc.tab);
  if (!sh) return;
  var v = sh.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    if (S(v[i][sc.col.UID]) === uid) {
      sh.getRange(i + 1, sc.col.PICTURE + 1).setValue(picture);
      return;
    }
  }
}

// list?type=ais|pc[&includeDeleted=1]   (omit type = both)
function listStaff(p) {
  var type = S(p.type).toLowerCase();
  var inc  = S(p.includeDeleted) === '1' || S(p.includeDeleted).toLowerCase() === 'true';
  var data = {};
  if (type === 'ais' || type === 'pc') {
    data[type] = readTab(type, inc);
  } else {
    data.ais = readTab('ais', inc);
    data.pc  = readTab('pc', inc);
  }
  return { ok:true, data:data };
}

// register?type=ais&name=...&pin=...&branch=...&uid=...&lineId=...&picture=...[&brand=...]
function registerStaff(p) {
  var type = reqType(p);
  var sc   = SCHEMA[type];
  var sh   = sheet(sc.tab);
  var rows = sh.getDataRange().getValues();

  var uid = S(p.uid);
  for (var i = 1; i < rows.length; i++) {
    if (uid && S(rows[i][sc.col.UID]) === uid)
      return { ok:false, error:'already registered', id:S(rows[i][sc.col.ID]) };
    if (!uid && S(rows[i][sc.col.NAME]) && S(rows[i][sc.col.NAME]) === S(p.name))
      return { ok:false, error:'name exists', id:S(rows[i][sc.col.ID]) };
  }

  var id  = nextId(rows, type, sc);
  var now = nowISO();
  var row = new Array(sc.ncols).fill('');
  row[sc.col.ID]      = id;
  row[sc.col.LINEID]  = S(p.lineId);
  row[sc.col.NAME]    = S(p.name);
  row[sc.col.PIN]     = S(p.pin);
  row[sc.col.BRANCH]  = S(p.branch);
  if (sc.col.BRAND !== undefined) row[sc.col.BRAND] = S(p.brand);
  row[sc.col.ACTIVE]  = 'TRUE';
  row[sc.col.CREATED] = now;
  row[sc.col.UPDATED] = now;
  row[sc.col.PICTURE] = S(p.picture);
  row[sc.col.UID]     = uid;
  sh.appendRow(row);
  return { ok:true, id:id };
}

// update?type=ais&id=...&[name|pin|branch|brand|picture|lineId|uid]=...
function updateStaff(p) {
  var type = reqType(p);
  var sc   = SCHEMA[type];
  var sh   = sheet(sc.tab);
  var r    = findRow(sh, sc, S(p.id));
  if (r < 0) return { ok:false, error:'id not found' };

  var fields = ['name','pin','branch','brand','picture','lineId','uid'];
  var keyToCol = { name:sc.col.NAME, pin:sc.col.PIN, branch:sc.col.BRANCH,
                   brand:sc.col.BRAND, picture:sc.col.PICTURE,
                   lineId:sc.col.LINEID, uid:sc.col.UID };
  var changed = 0;
  for (var k = 0; k < fields.length; k++) {
    var key = fields[k];
    var c   = keyToCol[key];
    if (c !== undefined && p[key] !== undefined) {
      sh.getRange(r, c+1).setValue(S(p[key])); changed++;
    }
  }
  if (changed) sh.getRange(r, sc.col.UPDATED+1).setValue(nowISO());
  return { ok:true, changed:changed };
}

// toggle?type=ais&id=...
function toggleActive(p) {
  var type = reqType(p);
  var sc   = SCHEMA[type];
  var sh   = sheet(sc.tab);
  var r    = findRow(sh, sc, S(p.id));
  if (r < 0) return { ok:false, error:'id not found' };
  var cur = S(sh.getRange(r, sc.col.ACTIVE+1).getValue()).toUpperCase() === 'TRUE';
  sh.getRange(r, sc.col.ACTIVE+1).setValue(cur ? 'FALSE' : 'TRUE');
  sh.getRange(r, sc.col.UPDATED+1).setValue(nowISO());
  return { ok:true, active:!cur };
}

// remove?type=ais&id=...   -> SOFT DELETE (Active=FALSE, Deleted=TRUE)
function removeStaff(p) {
  var type = reqType(p);
  var sc   = SCHEMA[type];
  var sh   = sheet(sc.tab);
  var r    = findRow(sh, sc, S(p.id));
  if (r < 0) return { ok:false, error:'id not found' };
  var dcol = deletedCol(sh, sc);           // 1-based; creates header if missing
  sh.getRange(r, sc.col.ACTIVE+1).setValue('FALSE');
  sh.getRange(r, dcol).setValue('TRUE');
  sh.getRange(r, sc.col.UPDATED+1).setValue(nowISO());
  return { ok:true, softDeleted:true };
}

// restore?type=ais&id=...  -> bring a soft-deleted row back (Active=TRUE)
function restoreStaff(p) {
  var type = reqType(p);
  var sc   = SCHEMA[type];
  var sh   = sheet(sc.tab);
  var r    = findRow(sh, sc, S(p.id));
  if (r < 0) return { ok:false, error:'id not found' };
  var dcol = deletedCol(sh, sc);
  sh.getRange(r, dcol).setValue('FALSE');
  sh.getRange(r, sc.col.ACTIVE+1).setValue('TRUE');
  sh.getRange(r, sc.col.UPDATED+1).setValue(nowISO());
  return { ok:true, restored:true };
}

// purge?type=ais&id=...  -> HARD delete (physically remove row). Use with care.
function purgeStaff(p) {
  var type = reqType(p);
  var sc   = SCHEMA[type];
  var sh   = sheet(sc.tab);
  var r    = findRow(sh, sc, S(p.id));
  if (r < 0) return { ok:false, error:'id not found' };
  sh.deleteRow(r);
  return { ok:true, purged:true };
}

// ====================== helpers =======================
function readTab(type, includeDeleted) {
  var sc = SCHEMA[type];
  var sh = sheet(sc.tab);
  var v  = sh.getDataRange().getValues();
  var c  = sc.col;
  var dIdx = deletedColIndex(v[0]);   // 0-based index of "Deleted" col, or -1
  var out = [];
  for (var i = 1; i < v.length; i++) {
    if (S(v[i][c.ID]) === '') continue;
    var deleted = (dIdx >= 0) && (S(v[i][dIdx]).toUpperCase() === 'TRUE');
    if (deleted && !includeDeleted) continue;
    var o = {
      id:      S(v[i][c.ID]),
      lineId:  S(v[i][c.LINEID]),
      name:    S(v[i][c.NAME]),
      pin:     S(v[i][c.PIN]),
      branch:  S(v[i][c.BRANCH]),
      active:  S(v[i][c.ACTIVE]).toUpperCase() === 'TRUE',
      created: S(v[i][c.CREATED]),
      updated: S(v[i][c.UPDATED]),
      picture: S(v[i][c.PICTURE]),
      uid:     S(v[i][c.UID]),
      deleted: deleted
    };
    if (c.BRAND !== undefined) o.brand = S(v[i][c.BRAND]);
    out.push(o);
  }
  return out;
}

// 0-based index of a header literally named "Deleted" (case-insensitive)
function deletedColIndex(headerRow) {
  if (!headerRow) return -1;
  for (var i = 0; i < headerRow.length; i++)
    if (S(headerRow[i]).trim().toLowerCase() === 'deleted') return i;
  return -1;
}

// 1-based column for "Deleted"; creates the header at first empty col if absent
function deletedCol(sh, sc) {
  var header = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), sc.ncols)).getValues()[0];
  var idx = deletedColIndex(header);
  if (idx >= 0) return idx + 1;
  var newCol = sc.ncols + 1;            // place right after the known schema
  sh.getRange(1, newCol).setValue('Deleted');
  return newCol;
}

function nextId(rows, type, sc) {
  var prefix = (type === 'ais') ? 'A' : 'P';
  var max = 0;
  for (var i = 1; i < rows.length; i++) {
    var id = S(rows[i][sc.col.ID]);
    var n  = parseInt(id.replace(/\D/g, ''), 10);
    if (!isNaN(n) && n > max) max = n;
  }
  return prefix + ('000' + (max + 1)).slice(-4);
}

function findRow(sh, sc, id) {
  var v = sh.getRange(1, sc.col.ID+1, sh.getLastRow(), 1).getValues();
  for (var i = 1; i < v.length; i++) if (S(v[i][0]) === id) return i + 1;
  return -1;
}

function reqType(p) {
  var t = S(p.type).toLowerCase();
  if (t !== 'ais' && t !== 'pc') throw new Error('type must be ais or pc');
  return t;
}

function sheet(name) {
  var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(name);
  if (!sh) throw new Error('tab not found: ' + name);
  return sh;
}

/* ================= ROS Sale (accumulating store) =================
   Tab: ROS_Sale
   A=DATE(yyyy-MM-dd) B=PRODUCT_TYPE C=MODEL D=LOCATION_CODE E=LOCATION_NAME
   F=QTY G=SRC_FILE H=UPLOADED_AT
   กลยุทธ์: replace-by-date — ลบข้อมูลเดิมของ "วันที่" ที่อยู่ในไฟล์ใหม่ แล้วใส่ของใหม่แทน
   => import ซ้ำได้ไม่เบิ้ล / วันที่ไม่อยู่ในไฟล์ไม่ถูกแตะ / ขายซ้ำจริงยังนับครบ
================================================================== */
var ROS_TAB = 'ROS_Sale';
var ROS_CACHE_TAB = 'ROS_CACHE';   // freeze สรุป — dashboard อ่านตัวนี้
var ROS_HEAD = ['DATE','PRODUCT_TYPE','MODEL','LOCATION_CODE','LOCATION_NAME','QTY','SRC_FILE','UPLOADED_AT','USER','PIN','MODELS_USED'];
// เฉพาะ 9 สาขา South
var ROS_ALLOW_LOC = {'1204':1,'1205':1,'1213':1,'1215':1,'1216':1,'1245':1,'1252':1,'1255':1,'1259':1};
// product type ที่ไม่นับ (SIM/Service)
var ROS_SKIP_PTYPE = {'SIM':1,'ESIM':1,'SERVICE':1};

function rosSheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(ROS_TAB);
  if (!sh) {
    sh = ss.insertSheet(ROS_TAB);
    sh.getRange(1, 1, 1, ROS_HEAD.length).setValues([ROS_HEAD]);
    sh.setFrozenRows(1);
  }
  // บังคับคอลัมน์ A (DATE) เป็น "ข้อความ" กัน Google Sheets แปลงเป็น Date object
  sh.getRange('A:A').setNumberFormat('@');
  return sh;
}

// อ่านค่าวันที่จากชีตให้ได้ 'yyyy-MM-dd' เสมอ (เผื่อแถวเก่าที่ถูกแปลงเป็น Date ไปแล้ว)
var ROS_EN_MON = {Jan:1,Feb:2,Mar:3,Apr:4,May:5,Jun:6,Jul:7,Aug:8,Sep:9,Oct:10,Nov:11,Dec:12};
function rosDayStr_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Bangkok', 'yyyy-MM-dd');
  var s = S(v).trim();
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[0];
  var m2 = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m2) {
    var y = Number(m2[3]); if (y > 2400) y -= 543;
    return y + '-' + ('0' + m2[2]).slice(-2) + '-' + ('0' + m2[1]).slice(-2);
  }
  // 'Sun Jul 05 2026 ...'
  var m3 = s.match(/^[A-Za-z]{3}\s+([A-Za-z]{3})\s+(\d{1,2})\s+(\d{4})/);
  if (m3 && ROS_EN_MON[m3[1]]) return m3[3] + '-' + ('0' + ROS_EN_MON[m3[1]]).slice(-2) + '-' + ('0' + m3[2]).slice(-2);
  return s;
}

// rows = [[date,ptype,model,locCode,locName,qty], ...] (ส่งมาแบบ JSON ผ่าน POST)
// ===== Server-side ROS parser (สำหรับ auto-import จากเมล) =====
var ROS_IMPORT_TOKEN = 'AIS-South-2026-SecretImport';
var ROS_PARSE_ALLOW_LOC = {'1204':1,'1205':1,'1213':1,'1215':1,'1216':1,'1245':1,'1252':1,'1255':1,'1259':1};
// map code -> ชื่อสาขา (fallback เมื่อไฟล์ ROS ไม่มีคอลัมน์ LOCATION NAME)
var ROS_LOC_NAME = {
  '1204':'Shop Central Surat Thani',
  '1205':'Shop Central Hat Yai',
  '1213':'Shop Central Phuket Floresta',
  '1215':'Shop Central Nakhon Si Thammarat',
  '1216':'Shop Central Samui',
  '1245':'Shop Serenade Club Central Festival Hat Yai',
  '1252':'Shop Hatyai Village',
  '1255':'Shop Central Phuket Festival',
  '1259':'Shop Robinson Trang'
};

function rosDateStr_(v) {
  var s = S(v).trim();
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) { var y = Number(m[3]); if (y > 2400) y -= 543; return y + '-' + ('0'+m[2]).slice(-2) + '-' + ('0'+m[1]).slice(-2); }
  var m2 = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m2) return m2[0];
  return s;
}

// parse CSV -> array of row objects (auto-detect delimiter + strip BOM)
function rosParseCsv_(text) {
  if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1);
  var lines = text.split(/\r\n|\n|\r/);
  if (!lines.length) return [];
  var hLine = lines[0];
  var cands = [
    { d: ',',  n: (hLine.match(/,/g)  || []).length },
    { d: '\t', n: (hLine.match(/\t/g) || []).length },
    { d: '|',  n: (hLine.match(/\|/g) || []).length },
    { d: ';',  n: (hLine.match(/;/g)  || []).length }
  ];
  cands.sort(function(a,b){ return b.n - a.n; });
  var delim = cands[0].n > 0 ? cands[0].d : ',';
  var header = rosSplitCsvLine_(lines[0], delim).map(function(x){ return S(x).replace(/\s+/g,' ').trim().toUpperCase(); });
  var rows = [];
  for (var i = 1; i < lines.length; i++) {
    if (!lines[i]) continue;
    var cols = rosSplitCsvLine_(lines[i], delim);
    var obj = {};
    for (var j = 0; j < header.length; j++) obj[header[j]] = S(cols[j]).trim();
    rows.push(obj);
  }
  return rows;
}
function rosSplitCsvLine_(line, delim) {
  delim = delim || ',';
  var out = [], cur = '', inQ = false;
  for (var i = 0; i < line.length; i++) {
    var c = line[i];
    if (c === '"') { if (inQ && line[i+1] === '"') { cur += '"'; i++; } else inQ = !inQ; }
    else if (c === delim && !inQ) { out.push(cur); cur = ''; }
    else cur += c;
  }
  out.push(cur);
  return out;
}

// แปลง row objects -> rows array [d,ptype,model,code,name,qty,user,pin] (กรอง DEVICE + รุ่น + 9 สาขา)
function rosParseRows_(data) {
  var MODEL_SET = rosModelSet_();
  var out = [];
  data.forEach(function(r){
    var model = S(r['MODEL']).trim();
    if (!model) return;
    var ptype = S(r['PRODUCT TYPE']).trim().toUpperCase() || 'DEVICE';
    if (ptype !== 'DEVICE') return;
    if (!MODEL_SET[model.replace(/\s+/g,'').toUpperCase()]) return;
    var code = S(r['LOCATION CODE']).trim();
    if (!ROS_PARSE_ALLOW_LOC[code]) return;
    var d = rosDateStr_(r['RECPT/INV DATE'] || r['PROCESS DATE']);
    var name = S(r['LOCATION NAME']).replace(/\s+/g,' ').trim() || ROS_LOC_NAME[code] || code;
    var qty = Number(S(r['QTY']).replace(/[^0-9.\-]/g,'')) || 1;
    var user = (S(r['SALE_EMPLOYEE_ID']) || S(r['COLUMN1'])).replace(/\s+/g,'').toLowerCase();
    var pin = S(r['SALE EMPLOYEE CODE']).replace(/\s+/g,'').trim();
    out.push([d, ptype.toUpperCase(), model.toUpperCase(), code, name, qty, user, pin]);
  });
  return out;
}

// ===== Endpoint: รับไฟล์ CSV แล้ว import ROS อัตโนมัติ =====
function rosUploadFile(p) {
  if (typeof ROS_IMPORT_TOKEN !== 'undefined' && ROS_IMPORT_TOKEN) {
    if (S(p.token) !== ROS_IMPORT_TOKEN) return { ok:false, error:'unauthorized' };
  }

  // กรณีส่ง Excel (.xlsx) มาเป็น base64 -> แปลงเป็น rows ผ่าน Drive
  if (p.xlsx) {
    try {
      var blob = Utilities.newBlob(Utilities.base64Decode(S(p.xlsx)), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', S(p.src)||'ros.xlsx');
      var rowsArr = rosXlsxBlobToRows_(blob);
      var objs = rosRowsArrToObjs_(rowsArr);
      var rowsX = rosParseRows_(objs);
      if (!rowsX.length) return { ok:false, error:'no valid rows จาก Excel (เช็ค DEVICE/รุ่น/9 สาขา)' };
      return rosUpload({ rows: rowsX, src: S(p.src) || 'auto-mail-xlsx' });
    } catch (ex) { return { ok:false, error:'xlsx: ' + ex.message }; }
  }

  var text = S(p.csv);
  if (!text) return { ok:false, error:'no csv data' };
  if (p.b64 === true || S(p.b64) === 'true') {
    try { text = Utilities.newBlob(Utilities.base64Decode(text)).getDataAsString('UTF-8'); }
    catch (e) { return { ok:false, error:'base64 decode failed: ' + e.message }; }
  }
  var data = rosParseCsv_(text);
  if (!data.length) return { ok:false, error:'parse: no rows' };
  var rows = rosParseRows_(data);
  if (!rows.length) return { ok:false, error:'no valid rows (เช็ค DEVICE/รุ่น/9 สาขา)' };
  return rosUpload({ rows: rows, src: S(p.src) || 'auto-mail' });
}

// แปลง Excel blob -> rows (array of arrays) ผ่าน Drive (รองรับ v2 insert + v3 create)
function rosXlsxBlobToRows_(blob) {
  var tempId = null;
  try {
    if (Drive.Files && typeof Drive.Files.insert === 'function') {
      tempId = Drive.Files.insert({ title: 'TEMP_ROS_'+Date.now(), mimeType: MimeType.GOOGLE_SHEETS }, blob, { convert: true }).id;
    } else if (Drive.Files && typeof Drive.Files.create === 'function') {
      tempId = Drive.Files.create({ name: 'TEMP_ROS_'+Date.now(), mimeType: MimeType.GOOGLE_SHEETS }, blob).id;
    } else {
      throw new Error('Drive API ไม่พร้อม — เปิด Advanced Drive Service');
    }
    var ssX = SpreadsheetApp.openById(tempId);
    var sh = ssX.getSheetByName('Data_Sale') || ssX.getSheets()[0];
    return sh.getDataRange().getValues();
  } finally {
    if (tempId) {
      try {
        if (Drive.Files && typeof Drive.Files.remove === 'function') Drive.Files.remove(tempId);
        else DriveApp.getFileById(tempId).setTrashed(true);
      } catch(e){}
    }
  }
}
// แปลง rows (array of arrays) -> objs (หา header row ที่มี MODEL + LOCATION NAME)
function rosRowsArrToObjs_(rows) {
  if (!rows || !rows.length) return [];
  var headerIdx = -1, header = null;
  for (var i = 0; i < Math.min(rows.length, 15); i++) {
    var up = rows[i].map(function(x){ return S(x).replace(/\s+/g,' ').trim().toUpperCase(); });
    if (up.indexOf('MODEL') >= 0 && (up.indexOf('LOCATION CODE') >= 0 || up.indexOf('LOCATION NAME') >= 0)) { headerIdx = i; header = up; break; }
  }
  if (headerIdx < 0) return [];
  var out = [];
  for (var r = headerIdx+1; r < rows.length; r++) {
    var obj = {};
    for (var c = 0; c < header.length; c++) obj[header[c]] = rows[r][c];
    out.push(obj);
  }
  return out;
}

function rosUpload(p) {
  var rows = p.rows;
  if (typeof rows === 'string') rows = JSON.parse(rows);
  if (!rows || !rows.length) return { ok:false, error:'no rows' };
  var src = S(p.src) || 'unknown.xlsx';
  var stamp = nowISO();

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sh = rosSheet_();

    // วันที่ที่อยู่ในไฟล์ใหม่ (normalize ให้เป็น yyyy-MM-dd)
    var dates = {};
    rows.forEach(function(r){ dates[rosDayStr_(r[0])] = 1; });

    // เก็บเฉพาะแถวเดิมที่ "ไม่ใช่" วันที่อยู่ในไฟล์ใหม่ (เขียนทับทีเดียว เร็วกว่า deleteRow ทีละแถว)
    var last = sh.getLastRow();
    var keep = [], removed = 0;
    if (last > 1) {
      var cur = sh.getRange(2, 1, last - 1, ROS_HEAD.length).getValues();
      for (var i = 0; i < cur.length; i++) {
        var cd = rosDayStr_(cur[i][0]);
        if (dates[cd]) removed++;
        else { cur[i][0] = cd; keep.push(cur[i]); }  // เขียนกลับเป็น text เสมอ
      }
    }

    var mu = (rosModels().models || []).join(',');
    var out = rows.map(function(r){
      return [rosDayStr_(r[0]), S(r[1]), S(r[2]), S(r[3]), S(r[4]), Number(r[5]) || 1, src, stamp, S(r[6]), S(r[7]), mu];
    });
    var final = keep.concat(out);

    // ล้างพื้นที่ข้อมูลเดิมแล้วเขียนใหม่ทั้งก้อน
    if (last > 1) sh.getRange(2, 1, last - 1, ROS_HEAD.length).clearContent();
    if (final.length) {
      sh.getRange(2, 1, final.length, 1).setNumberFormat('@');   // DATE = text
      sh.getRange(2, 1, final.length, ROS_HEAD.length).setValues(final);
    }

    CacheService.getScriptCache().remove('ros_sum_v1');
    rosRebuildCache_();   // อัปเดต cache หลัง import
    return { ok:true, inserted:out.length, removed:removed, dates:Object.keys(dates).sort(), src:src };
  } finally {
    lock.releaseLock();
  }
}

// สรุปยอดขายต่อสาขา (นับจาก ROS_Sale ที่สะสมไว้) — กรองตามเดือน (yyyy-MM) หรือช่วงวัน
function rosSummary(p) {
  var month = S(p.month);
  var from = S(p.from), to = S(p.to);
  if (from || to) return rosComputeFromSheet_(month, from, to);
  var key = month || '__ALL__';
  var cached = rosCacheRead_(key);
  if (cached) { cached.cached = true; return cached; }
  var res = rosComputeFromSheet_(month, '', '');
  rosCacheWrite_(key, res);
  return res;
}

function rosComputeFromSheet_(month, from, to) {
  var sh = rosSheet_();
  var last = sh.getLastRow();
  if (last < 2) return { ok:true, byLoc:{}, byUser:{}, total:0, dates:[], months:[] };
  var v = sh.getRange(2, 1, last - 1, ROS_HEAD.length).getValues();
  return rosCompute_(v, month, from, to);
}

// อ่านรายการรุ่นที่ตั้งไว้ (เป็น set สำหรับกรอง)
function rosModelSet_() {
  var raw = PropertiesService.getScriptProperties().getProperty('ROS_MODELS');
  var list;
  if (raw) { try { list = JSON.parse(raw); } catch (e) { list = null; } }
  if (!list || !list.length) list = ROS_MODELS_DEFAULT;
  var set = {};
  list.forEach(function(m){ set[String(m).replace(/\s+/g,'').toUpperCase()] = 1; });
  return set;
}

function rosCompute_(v, month, from, to) {
  var byLoc = {}, byUser = {}, total = 0, dset = {}, mset = {}, minD = '', maxD = '', muset = {};
  var MODEL_SET = rosModelSet_();
  for (var i = 0; i < v.length; i++) {
    var d = rosDayStr_(v[i][0]);
    if (d) mset[d.slice(0, 7)] = 1;
    if (month) { if (d.slice(0, 7) !== month) continue; }
    else {
      if (from && d < from) continue;
      if (to && d > to) continue;
    }
    var code = S(v[i][3]), name = S(v[i][4]), q = Number(v[i][5]) || 0;
    // กรองเฉพาะ 9 สาขา South
    if (!ROS_ALLOW_LOC[code]) continue;
    // นับเฉพาะ DEVICE
    var ptype = S(v[i][1]).toUpperCase();
    if (ptype !== 'DEVICE') continue;
    // นับเฉพาะรุ่นที่ตั้งไว้
    var model = S(v[i][2]).replace(/\s+/g,'').toUpperCase();
    if (!MODEL_SET[model]) continue;
    var user = S(v[i][8]).replace(/\s+/g,'').toLowerCase();
    var pin = S(v[i][9]).replace(/\s+/g,'').trim();
    var k = code + '|' + name;
    if (!byLoc[k]) byLoc[k] = { code:code, name:name, qty:0 };
    byLoc[k].qty += q;
    var uk = user || ('#'+pin);
    if (uk && uk!=='#') {
      if (!byUser[uk]) byUser[uk] = { user:user, pin:pin, code:code, name:name, qty:0 };
      byUser[uk].qty += q;
      if (!byUser[uk].pin && pin) byUser[uk].pin = pin;
    }
    total += q; dset[d] = 1;
    var mu = S(v[i][10]); if (mu) muset[mu] = 1;
    if (!minD || d < minD) minD = d;
    if (!maxD || d > maxD) maxD = d;
  }
  var muArr = Object.keys(muset);
  var modelsUsed = null;
  if (muArr.length) {
    var all = {};
    muArr.forEach(function(x){ x.split(',').forEach(function(m){ if(m) all[m] = 1; }); });
    modelsUsed = Object.keys(all).sort();
  }
  return { ok:true, byLoc:byLoc, byUser:byUser, total:total,
           dates:Object.keys(dset).sort(),
           months:Object.keys(mset).sort(),
           minD:minD, maxD:maxD, month:month || '',
           modelsUsed:modelsUsed, mixed:(muArr.length > 1) };
}

// ===== ROS Cache helpers =====
function rosCacheSheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(ROS_CACHE_TAB);
  if (!sh) {
    sh = ss.insertSheet(ROS_CACHE_TAB);
    sh.getRange(1,1,1,3).setValues([['KEY','UPDATED_AT','JSON']]);
    sh.setFrozenRows(1);
  }
  return sh;
}
function rosCacheRead_(key) {
  var sh = rosCacheSheet_();
  var last = sh.getLastRow();
  if (last < 2) return null;
  var v = sh.getRange(2,1,last-1,3).getValues();
  for (var i=0;i<v.length;i++){
    if (S(v[i][0]) === key) { try { return JSON.parse(v[i][2]); } catch(e){ return null; } }
  }
  return null;
}
function rosCacheWrite_(key, obj) {
  var sh = rosCacheSheet_();
  var json = JSON.stringify(obj);
  var last = sh.getLastRow();
  var rowIdx = -1;
  if (last >= 2) {
    var keys = sh.getRange(2,1,last-1,1).getValues();
    for (var i=0;i<keys.length;i++){ if (S(keys[i][0])===key){ rowIdx = i+2; break; } }
  }
  if (rowIdx < 0) rowIdx = last + 1;
  sh.getRange(rowIdx,1,1,3).setValues([[key, nowISO(), json]]);
}
function rosCacheClear_() {
  var sh = rosCacheSheet_();
  var last = sh.getLastRow();
  if (last > 1) sh.getRange(2,1,last-1,3).clearContent();
}
function rosRebuildCache_() {
  rosCacheClear_();
  var sh = rosSheet_();
  var last = sh.getLastRow();
  if (last < 2) return { months:[], rows:0 };
  var v = sh.getRange(2, 1, last - 1, ROS_HEAD.length).getValues();
  var mset = {};
  for (var i=0;i<v.length;i++){ var d=rosDayStr_(v[i][0]); if(d) mset[d.slice(0,7)]=1; }
  var months = Object.keys(mset).sort();
  months.forEach(function(m){ rosCacheWrite_(m, rosCompute_(v, m, '', '')); });
  rosCacheWrite_('__ALL__', rosCompute_(v, '', '', ''));
  return { months:months, rows:v.length };
}
function rosRebuild(p) {
  var r = rosRebuildCache_();
  return { ok:true, rebuilt:true, months:r.months, rows:r.rows };
}

// รายการไฟล์ที่เคย import + ช่วงวันที่มีข้อมูล
function rosBatches() {
  var sh = rosSheet_();
  var last = sh.getLastRow();
  if (last < 2) return { ok:true, batches:[], dates:[] };
  var v = sh.getRange(2, 1, last - 1, ROS_HEAD.length).getValues();
  var b = {}, dset = {};
  for (var i = 0; i < v.length; i++) {
    var src = S(v[i][6]), up = S(v[i][7]), d = rosDayStr_(v[i][0]);
    dset[d] = 1;
    if (!b[src]) b[src] = { src:src, rows:0, uploadedAt:up, minD:d, maxD:d };
    b[src].rows++;
    if (up > b[src].uploadedAt) b[src].uploadedAt = up;
    if (d < b[src].minD) b[src].minD = d;
    if (d > b[src].maxD) b[src].maxD = d;
  }
  var arr = Object.keys(b).map(function(k){ return b[k]; });
  arr.sort(function(x,y){ return x.uploadedAt < y.uploadedAt ? 1 : -1; });
  return { ok:true, batches:arr, dates:Object.keys(dset).sort() };
}

// ลบข้อมูลทั้งหมด (ใช้ตอน reset) — ต้องส่ง confirm=YES
// ซ่อมข้อมูลเดิม: แปลงคอลัมน์ DATE ที่ถูก Sheets เปลี่ยนเป็น Date object กลับเป็น text yyyy-MM-dd
/* ---- รายการรุ่นที่ใช้กรอง (แก้ได้จากหน้าเว็บ ไม่ต้องแก้โค้ด) ---- */
var ROS_MODELS_DEFAULT = ['IP17_256GB','IP17_512GB','IP17E_256GB','IP17P_256GB','IP17P_512GB','IP17PM_1TB','IP17PM_256GB','IP17PM_512GB'];

function rosModels() {
  var raw = PropertiesService.getScriptProperties().getProperty('ROS_MODELS');
  var list;
  if (raw) { try { list = JSON.parse(raw); } catch (e) { list = null; } }
  if (!list || !list.length) list = ROS_MODELS_DEFAULT;
  return { ok:true, models:list };
}

function rosSetModels(p) {
  var v = p.models;
  if (typeof v === 'string') { try { v = JSON.parse(v); } catch (e) { v = S(v).split(/[\s,]+/); } }
  if (!v || !v.length) return { ok:false, error:'no models' };
  var clean = [], seen = {};
  for (var i = 0; i < v.length; i++) {
    var m = S(v[i]).replace(/\s+/g, '').toUpperCase();
    if (m && !seen[m]) { seen[m] = 1; clean.push(m); }
  }
  if (!clean.length) return { ok:false, error:'no valid models' };
  PropertiesService.getScriptProperties().setProperty('ROS_MODELS', JSON.stringify(clean));
  return { ok:true, models:clean };
}

function rosFix() {
  var sh = rosSheet_();
  var last = sh.getLastRow();
  if (last < 2) return { ok:true, fixed:0 };
  var rng = sh.getRange(2, 1, last - 1, 1);
  var v = rng.getValues();
  var out = [], fixed = 0;
  for (var i = 0; i < v.length; i++) {
    var before = v[i][0];
    var after = rosDayStr_(before);
    if (String(before) !== after) fixed++;
    out.push([after]);
  }
  rng.setNumberFormat('@');
  rng.setValues(out);
  return { ok:true, fixed:fixed, total:v.length };
}

function rosClear(p) {
  if (S(p.confirm) !== 'YES') return { ok:false, error:'need confirm=YES' };
  var sh = rosSheet_();
  var last = sh.getLastRow();
  var removed = Math.max(0, last - 1);
  if (last > 1) {
    var lastCol = Math.max(sh.getLastColumn(), ROS_HEAD.length);
    sh.getRange(2, 1, last - 1, lastCol).clearContent();
    try { sh.deleteRows(2, last - 1); } catch (e) {}
  }
  sh.getRange(1, 1, 1, ROS_HEAD.length).setValues([ROS_HEAD]);
  CacheService.getScriptCache().remove('ros_sum_v1');
  rosCacheClear_();
  return { ok:true, cleared:true, removed:removed };
}

function nowISO() {
  return Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
}

function S(x) { return (x === null || x === undefined) ? '' : String(x); }

function respond(obj, cb) {
  var json = JSON.stringify(obj);
  if (cb) return ContentService.createTextOutput(cb + '(' + json + ')')
                  .setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(json)
                  .setMimeType(ContentService.MimeType.JSON);
}

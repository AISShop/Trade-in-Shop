/* =====================================================================
 * sb.js — ชั้นข้อมูลกลาง AIS Area South (Supabase)
 * ทุกหน้าใช้ไฟล์นี้ไฟล์เดียว: StaffForm / AdminDashboard / Trade-in Dashboard
 * แก้ค่าเชื่อมต่อที่ CFG ด้านล่างที่เดียว
 *
 * ต้องโหลด supabase-js ก่อนไฟล์นี้:
 *   <script src="https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2"></script>
 *   <script src="https://aisshop.github.io/Trade-in-Shop/sb.js"></script>
 * ===================================================================== */
(function () {
  var CFG = {
    URL:          'https://fpbkfkeydplyjhhehldf.supabase.co',   // <- Project URL
    ANON_KEY:     'sb_publishable__EF_KySrg3aW_DjjWvarwg_H7jyj3wx',                   // <- anon public key (เปิดเผยได้ ปลอดภัยด้วย RLS)
    LOGIN_DOMAIN: 'aisshop.local',  // พิมพ์ username เฉยๆ ระบบเติม @aisshop.local ให้
    PHOTO_BUCKET: 'tradein-photos',
    TRADE_NEWEST_FIRST: true        // Admin: เรียงรายการใหม่สุดก่อน
  };

  if (!window.supabase || !window.supabase.createClient) {
    console.error('[sb.js] ต้องโหลด supabase-js ก่อน sb.js');
    return;
  }
  var sb = window.supabase.createClient(CFG.URL, CFG.ANON_KEY, {
    auth: { persistSession: true, autoRefreshToken: true, storageKey: 'ais-south-auth' }
  });

  var TRADE_COLS = 'id,ts,branch,name,staff_type,brand,ais_pro,accept,series,model,price,buy,' +
                   'reason_no_accept,reason_no_trade,shop_price,other,pic_acc,pic_no_acc';

  /* ---------- helpers ---------- */
  function toEmail(u) {
    u = String(u || '').trim();
    return u.indexOf('@') >= 0 ? u : (u.toLowerCase() + '@' + CFG.LOGIN_DOMAIN);
  }
  function p2(n) { return (n < 10 ? '0' : '') + n; }
  // ISO/Date -> 'yyyy-mm-dd HH:MM:SS' เวลาไทย (ไม่ขึ้นกับเครื่องผู้ใช้)
  function bkkStr(v) {
    var d = new Date(v); if (isNaN(d)) return '';
    var t = new Date(d.getTime() + 7 * 3600e3);
    return t.getUTCFullYear() + '-' + p2(t.getUTCMonth() + 1) + '-' + p2(t.getUTCDate()) + ' ' +
           p2(t.getUTCHours()) + ':' + p2(t.getUTCMinutes()) + ':' + p2(t.getUTCSeconds());
  }
  function numStr(v) { return (v === null || v === undefined) ? '' : String(v); }
  function S(v) { return (v === null || v === undefined) ? '' : String(v); }
  function isAuthErr(e) {
    var m = String((e && (e.message || e.code)) || e || '').toLowerCase();
    return /jwt|auth|401|403|permission|not logged/.test(m);
  }

  // ดึงทุกแถว (แบ่งหน้า 1000 แถว ยิงขนานทีละ 6 หน้า — เร็ว ไม่ค้าง)
  async function fetchAll(table, cols, apply) {
    var PAGE = 1000;
    var q0 = sb.from(table).select(cols, { count: 'exact' });
    if (apply) q0 = apply(q0);
    var r0 = await q0.order('id', { ascending: true }).range(0, PAGE - 1);
    if (r0.error) throw r0.error;
    var out = r0.data || [];
    var total = r0.count || out.length;
    var jobs = [];
    for (var from = PAGE; from < total; from += PAGE) {
      var q = sb.from(table).select(cols);
      if (apply) q = apply(q);
      jobs.push(q.order('id', { ascending: true }).range(from, from + PAGE - 1));
    }
    for (var i = 0; i < jobs.length; i += 6) {
      var res = await Promise.all(jobs.slice(i, i + 6));
      for (var j = 0; j < res.length; j++) {
        if (res[j].error) throw res[j].error;
        out = out.concat(res[j].data || []);
      }
    }
    return out;
  }

  /* ---------- config (เป้า / รุ่น / สาขา) ---------- */
  var _cfgPromise = null;
  function loadConfig(force) {
    if (_cfgPromise && !force) return _cfgPromise;
    _cfgPromise = Promise.all([
      sb.from('series').select('code,label,target_pct,ros_models,sort,active').order('sort'),
      sb.from('branches').select('code,name,area,aliases,sort,active').order('sort')
    ]).then(function (r) {
      var series = (r[0].data || []).filter(function (x) { return x.active; });
      var branches = (r[1].data || []).filter(function (x) { return x.active; });
      var targets = {};
      series.forEach(function (s) { targets[s.code] = Number(s.target_pct); });
      window.AIS_TARGETS = targets;       // หน้า dashboard อ่านเป้าจากตรงนี้
      window.AIS_SERIES = series;
      window.AIS_BRANCHES = branches;
      return { series: series, branches: branches, targets: targets };
    }).catch(function (e) { console.warn('[sb.js] config', e); _cfgPromise = null; return null; });
    return _cfgPromise;
  }

  /* ---------- auth ---------- */
  async function login(user, pass) {
    var r = await sb.auth.signInWithPassword({ email: toEmail(user), password: String(pass || '') });
    if (r.error) return { ok: false, error: r.error.message };
    return { ok: true, token: 'sb', user: r.data.user };
  }
  async function isAdmin() {
    var r = await sb.rpc('is_admin');
    return !r.error && r.data === true;
  }
  async function adminLogin(user, pass) {
    var r = await login(user, pass);
    if (!r.ok) return r;
    if (!(await isAdmin())) { await sb.auth.signOut(); return { ok: false, error: 'ไม่มีสิทธิ์ admin' }; }
    return r;
  }
  async function hasSession() {
    var r = await sb.auth.getSession();
    return !!(r && r.data && r.data.session);
  }
  function logout() { return sb.auth.signOut().catch(function () {}); }

  /* ---------- Trade-in: Admin (รูปแบบฟิลด์เดียวกับ TradeInBackend เดิม) ---------- */
  function toAdminRow(r) {
    return {
      id: r.id, ts: bkkStr(r.ts), branch: S(r.branch), name: S(r.name),
      type: S(r.staff_type), brand: S(r.brand), aisPro: S(r.ais_pro),
      accept: r.accept ? 'ตอบรับ' : 'ไม่ตอบรับ',
      series: S(r.series), model: S(r.model), price: numStr(r.price), buy: S(r.buy),
      reasonNoAcc: S(r.reason_no_accept), reasonNoTrade: S(r.reason_no_trade),
      shopPrice: S(r.shop_price), other: S(r.other),
      picAcc: S(r.pic_acc), picNoAcc: S(r.pic_no_acc)
    };
  }
  async function tradeList() {
    try {
      var rows = await fetchAll('tradein', TRADE_COLS, function (q) { return q.eq('deleted', false); });
      rows.sort(function (a, b) { return CFG.TRADE_NEWEST_FIRST ? (b.ts < a.ts ? -1 : 1) : (a.ts < b.ts ? -1 : 1); });
      return { ok: true, data: rows.map(toAdminRow) };
    } catch (e) { return { ok: false, error: e.message || String(e), auth: isAuthErr(e) }; }
  }

  /* ---------- Trade-in: Dashboard หน้า 1 (รูปแบบคอลัมน์เดียวกับชีทเดิม) ---------- */
  function toLegacyRow(r) {
    return {
      'ประทับเวลา': bkkStr(r.ts),
      'สาขา': S(r.branch),
      'ชื่อ / (Staff,PC)': S(r.name),
      'Staff Type (AIS or PC)': S(r.staff_type),
      'Brand': S(r.brand),
      'AIS Pro': S(r.ais_pro),
      'Series': S(r.series),
      'เสนอขาย Trade in ลูกค้าตอบรับหรือไม่': r.accept ? 'ตอบรับ' : 'ไม่ตอบรับ',
      'สาเหตุที่ไม่ตอบรับ': S(r.reason_no_accept),
      'สาเหตุที่ไม่เทรด': S(r.reason_no_trade),
      'ภาพประกอบราคาเทรด (ถ้ามี)': [r.pic_acc, r.pic_no_acc].filter(Boolean).join(', ')
    };
  }
  // sinceYmd = 'yyyy-mm-dd' -> ดึงเฉพาะตั้งแต่วันนั้น (เวลาไทย) / ว่าง = ทั้งหมด
  async function legacyRows(sinceYmd) {
    try {
      if (!(await hasSession())) return { ok: false, error: 'not logged in', auth: true };
      await loadConfig();
      var rows = await fetchAll('tradein', TRADE_COLS, function (q) {
        q = q.eq('deleted', false);
        return sinceYmd ? q.gte('ts', sinceYmd + 'T00:00:00+07:00') : q;
      });
      return { ok: true, rows: rows.map(toLegacyRow) };
    } catch (e) { return { ok: false, error: e.message || String(e), auth: isAuthErr(e) }; }
  }

  /* ---------- ตัวเลือก dropdown ---------- */
  async function getOptions() {
    var r = await sb.from('form_options').select('category,value,sort').order('sort');
    if (r.error) return { ok: false, error: r.error.message };
    var d = { noAccept: [], noTrade: [] };
    (r.data || []).forEach(function (x) { if (d[x.category]) d[x.category].push(x.value); });
    return { ok: true, data: d };
  }
  async function setOptions(category, values) {
    var r = await sb.rpc('set_form_options', { p_category: category, p_values: values });
    if (r.error) return { ok: false, error: /unauthorized/i.test(r.error.message) ? 'unauthorized' : r.error.message };
    return { ok: true, count: r.data };
  }

  /* ---------- ฟอร์มพนักงาน ---------- */
  async function searchModel(q) {
    var key = String(q || '').toLowerCase().replace(/\s+/g, '').replace(/[%_,()]/g, '');
    if (key.length < 2) return { ok: true, data: [] };
    var r = await sb.from('trade_models').select('name').eq('active', true)
      .ilike('search_key', '%' + key + '%').order('name').limit(25);
    if (r.error) return { ok: false, error: r.error.message };
    return { ok: true, data: (r.data || []).map(function (x) { return x.name; }) };
  }
  async function priceHint(model) {
    var r = await sb.rpc('price_hint', { p_model: model });
    if (r.error) return { ok: false, error: r.error.message };
    return { ok: true, data: r.data };
  }
  function uuid() {
    if (window.crypto && crypto.randomUUID) return crypto.randomUUID();
    return Date.now().toString(36) + Math.random().toString(36).slice(2, 12);
  }
  async function uploadPhoto(dataUrl) {
    if (!dataUrl) return '';
    var blob = await (await fetch(dataUrl)).blob();
    var now = new Date(Date.now() + 7 * 3600e3);
    var path = 'form/' + now.getUTCFullYear() + '-' + p2(now.getUTCMonth() + 1) + '/' + uuid() + '.jpg';
    var up = await sb.storage.from(CFG.PHOTO_BUCKET).upload(path, blob, { contentType: 'image/jpeg', upsert: false });
    if (up.error) throw new Error('อัปโหลดรูปไม่สำเร็จ: ' + up.error.message);
    return sb.storage.from(CFG.PHOTO_BUCKET).getPublicUrl(path).data.publicUrl;
  }
  // payload = รูปแบบเดิมของ StaffForm (picAccB64 / picNoAccB64 เป็น dataURL)
  async function submitTradein(payload) {
    try {
      var p = Object.assign({}, payload);
      p.picAcc = await uploadPhoto(p.picAccB64);
      p.picNoAcc = await uploadPhoto(p.picNoAccB64);
      delete p.picAccB64; delete p.picNoAccB64; delete p.action;
      var r = await sb.rpc('submit_tradein', { p: p });
      if (r.error) return { ok: false, error: r.error.message };
      return r.data;
    } catch (e) { return { ok: false, error: e.message || String(e) }; }
  }

  window.AISDB = {
    CFG: CFG, client: sb,
    login: login, adminLogin: adminLogin, isAdmin: isAdmin, hasSession: hasSession, logout: logout,
    loadConfig: loadConfig, fetchAll: fetchAll, bkkStr: bkkStr,
    tradeList: tradeList, legacyRows: legacyRows,
    getOptions: getOptions, setOptions: setOptions,
    searchModel: searchModel, priceHint: priceHint, submitTradein: submitTradein
  };
})();

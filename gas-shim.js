/* gas-shim.js — แก้ "network / ไม่สามารถเปิดไฟล์ได้" เมื่อเบราว์เซอร์ล็อกอิน Google หลายบัญชี
 * ดักทุก <script src="https://script.google.com/macros/..."> (JSONP) แล้วเรียกแทนด้วย fetch แบบ "ไม่ส่งบัญชี Google"
 * (credentials:'omit' = เหมือนหน้าต่างไม่ระบุตัวตน) · ลองซ้ำ 2 ครั้ง · สุดท้ายค่อยใช้ JSONP แบบเดิม
 * ใช้: ใส่ <script src="gas-shim.js"></script> เป็นบรรทัดแรกใน <head> ของแต่ละหน้า — ไม่ต้องแก้โค้ดเดิม
 */
(function(){
  if (window.__gasShim) return; window.__gasShim = true;
  var RX = /^https:\/\/script\.google(usercontent)?\.com\/macros\//i;
  var origAppend = Node.prototype.appendChild, origInsert = Node.prototype.insertBefore;

  function fire(el, type){ try{ var h = el['on' + type]; if (typeof h === 'function') h.call(el, { type:type, target:el }); }catch(e){} try{ el.dispatchEvent(new Event(type)); }catch(e){} }

  function load(el, src){
    var tries = 0;
    function attempt(){
      tries++;
      fetch(src, { method:'GET', credentials:'omit', redirect:'follow', cache:'no-store' })
        .then(function(r){ if (!r.ok) throw new Error('HTTP ' + r.status); return r.text(); })
        .then(function(t){
          if (!/^\s*[\w$.]+\s*\(/.test(t)) throw new Error('not JSONP');     // ได้หน้า error ของ Google แทนข้อมูล
          (0, eval)(t);                                                       // เรียก callback ของหน้าเดิม
          fire(el, 'load');
        })
        .catch(function(){
          if (tries < 3) return setTimeout(attempt, 800 * tries);              // ลองซ้ำ
          // สุดท้าย: ใช้ JSONP แบบเดิม (ส่งบัญชี Google)
          var s = document.createElement('script'); s.__gasShimSkip = true; s.src = src;
          s.onload = function(){ fire(el, 'load'); }; s.onerror = function(){ fire(el, 'error'); };
          origAppend.call(document.head || document.documentElement, s);
        });
    }
    attempt();
  }

  function intercept(node){
    if (node && node.tagName === 'SCRIPT' && !node.__gasShimSkip) {
      var src = node.getAttribute('src') || '';
      if (RX.test(src)) {
        node.removeAttribute('src');
        node.type = 'text/x-gas-shim';                                        // ใส่ลงหน้าได้ แต่เบราว์เซอร์ไม่โหลดเอง
        node.setAttribute('data-gas-src', src);
        setTimeout(function(){ load(node, src); }, 0);
      }
    }
    return node;
  }
  Node.prototype.appendChild = function(n){ return origAppend.call(this, intercept(n)); };
  Node.prototype.insertBefore = function(n, ref){ return origInsert.call(this, intercept(n), ref); };
})();

/* ═══════════════════════════════════════════════════════════════
   THP-GHANA SMART ATTENDANCE SYSTEM — Application Logic v6
   SERVER-FIRST ARCHITECTURE
   ─────────────────────────────────────────────
   KEY CHANGE: Google Sheets is the single source of truth.
   • Login calls the server — no local password checking
   • Session restore calls validateSession on the server
   • All writes (clock, leave, staff) go to server FIRST
   • localStorage is a READ CACHE only (speeds up renders)
   • No hardcoded DEF_STAFF — staff data lives on server
   ─────────────────────────────────────────────
   TABLE OF CONTENTS:
   1. Utility Helpers
   2. Security (hashing for display only, session storage)
   3. UI Helpers (toast, theme, navigation)
   4. API Module (server-first fetch wrapper)
   5. Ghana Public Holidays
   6. Leave Configuration & Helpers
   7. App Class
      a. Constructor & Hydration
      b. Clock / Time Display
      c. QR Landing
      d. Login (SERVER-SIDE) & Session Restore
      e. Clock In / Out (SERVER-FIRST)
      f. Staff Logs & Filters
      g. Leave Management
      h. Leave Review
      i. Notification Badges
      j. Manager Dashboard & Reports
      k. Admin Dashboard & Records
      l. Staff Management (CRUD)
      m. Reports & Exports
      n. QR Code Generation
      o. Password Change (SERVER-FIRST)
   8. Session Restore (calls server validateSession)
═══════════════════════════════════════════════════════════════ */

'use strict';

/* ═══════════════════════════════════════════════
   1. UTILITY HELPERS
═══════════════════════════════════════════════ */
const $=id=>document.getElementById(id);
const fx=(n,d=2)=>parseFloat(n||0).toFixed(d);
const fmtT=iso=>{if(!iso)return'--';const d=new Date(iso);return isNaN(d)?iso:d.toLocaleTimeString([],{hour:'2-digit',minute:'2-digit'});};
const fmtD=iso=>{if(!iso)return'--';const d=new Date(iso);if(isNaN(d))return iso;return d.toLocaleDateString('en-GB',{day:'2-digit',month:'short',year:'numeric'});};
const fmtISO=iso=>{if(!iso)return'--';
  if(typeof iso==='string'&&iso.match(/^\d{1,2}\s\w{3}\s\d{4}$/))return iso;
  const[y,m,dd]=(iso+'').split('-');if(!y||!m||!dd)return iso;
  const d=new Date(iso);return isNaN(d)?iso:d.toLocaleDateString('en-GB',{day:'2-digit',month:'short',year:'numeric'});};
const fmtDT=iso=>{if(!iso)return'--';const d=new Date(iso);return isNaN(d)?iso:d.toLocaleDateString('en-GB',{day:'2-digit',month:'short',year:'numeric'})+' '+d.toLocaleTimeString([],{hour:'2-digit',minute:'2-digit'});};

/* ═══════════════════════════════════════════════
   2. SECURITY
═══════════════════════════════════════════════ */
async function hashPass(id,pass){
  const raw=id+':'+pass;
  const buf=await crypto.subtle.digest('SHA-256',new TextEncoder().encode(raw));
  return Array.from(new Uint8Array(buf)).map(b=>b.toString(16).padStart(2,'0')).join('');
}
function isHashed(p){return p&&p.length===64&&/^[0-9a-f]+$/.test(p);}

/* Session stored in localStorage: {id, token, expiresAt}
   Token is validated against the server Sessions sheet */
const SESSION_HOURS=12;
function saveSession(id,token){
  localStorage.setItem('thp_session',JSON.stringify({id,token,expiresAt:Date.now()+SESSION_HOURS*3600000}));
}
function getSession(){
  try{
    const s=JSON.parse(localStorage.getItem('thp_session')||'null');
    if(!s||!s.id||!s.token)return null;
    if(s.expiresAt&&Date.now()>s.expiresAt){localStorage.removeItem('thp_session');return null;}
    return s;
  }catch(e){return null;}
}
function clearSession(){localStorage.removeItem('thp_session');}

const today=()=>fmtD(new Date().toISOString());
const todayISO=()=>new Date().toISOString().slice(0,10);
const sameDay=(dateStr)=>{if(!dateStr)return false;const d=new Date(dateStr);return!isNaN(d)&&d.toISOString().slice(0,10)===todayISO();};

/* ═══════════════════════════════════════════════
   3. UI HELPERS
═══════════════════════════════════════════════ */
const AV_COLORS=['#2D3592','#3DBFB8','#F5A623','#22c55e','#ef4444','#818cf8','#06b6d4','#f97316','#a855f7','#ec4899'];
function avColor(name){ return AV_COLORS[name.charCodeAt(0)%AV_COLORS.length]; }
function ini(s){ return s.split(' ').map(w=>w[0]).join('').toUpperCase().slice(0,2); }
function roleLabel(r){const m={'country_leader':'Country Leader','manager':'Manager','staff':'Staff','admin':'Admin'};return m[r]||r||'Staff';}

function toast(msg,type='ok'){
  const el=document.createElement('div'); el.className='toast '+type;
  el.innerHTML=(type==='ok'?'✅ ':type==='info'?'ℹ️ ':'❌ ')+msg;
  $('toasts').appendChild(el); setTimeout(()=>el.remove(),3600);
}
function togglePass(){
  const inp=$('uni-pass'),btn=$('eye-btn');
  if(!inp)return;
  if(inp.type==='password'){inp.type='text';if(btn)btn.textContent='🙈';}
  else{inp.type='password';if(btn)btn.textContent='👁';}
}
function switchTab(t){} // legacy no-op
function showView(id){
  document.querySelectorAll('.view').forEach(v=>v.classList.remove('active'));
  $(id).classList.add('active');
  ['mob-nav-staff','mob-nav-mgr','mob-nav-admin'].forEach(n=>{const el=$(n);if(el)el.style.display='none';});
  if(id==='staff-view'   &&$('mob-nav-staff')) $('mob-nav-staff').style.display='block';
  if(id==='manager-view' &&$('mob-nav-mgr'))   $('mob-nav-mgr').style.display='block';
  if(id==='admin-view'   &&$('mob-nav-admin')) $('mob-nav-admin').style.display='block';
}
/* Belt-and-braces: the mobile tab strip must never show on wide screens,
   even if a cached stylesheet is being served. */
function hideSplash(){
  const sp=document.getElementById('thp-splash');
  if(sp&&!sp.classList.contains('hide')){
    sp.classList.add('hide');
    setTimeout(()=>{if(sp.parentNode)sp.remove();},600);
  }
}
window.addEventListener('load',()=>setTimeout(hideSplash,650));
setTimeout(hideSplash,6000);   // never let the splash trap the user

function _syncNavForWidth(){
  const wide=window.innerWidth>768;
  document.querySelectorAll('.mob-nav').forEach(n=>{n.style.display=wide?'none':'';});
  document.querySelectorAll('.mob-menu-btn').forEach(b=>{b.style.display=wide?'none':'';});
  if(wide){
    document.querySelectorAll('.sidebar.open').forEach(e=>e.classList.remove('open'));
    const bd=document.getElementById('sb-backdrop');if(bd)bd.classList.remove('on');
  }
}
window.addEventListener('resize',_syncNavForWidth);
window.addEventListener('orientationchange',()=>setTimeout(_syncNavForWidth,120));
document.addEventListener('DOMContentLoaded',_syncNavForWidth);
setTimeout(_syncNavForWidth,300);

function toggleNavGroup(id){
  const g=document.getElementById(id);if(!g)return;
  const h=g.querySelector('.nav-grp-hdr'),b=g.querySelector('.nav-grp-body');
  const willCollapse=!h.classList.contains('collapsed');
  h.classList.toggle('collapsed',willCollapse);
  b.classList.toggle('hidden',willCollapse);
  try{localStorage.setItem('thp_nav_'+id,willCollapse?'0':'1');}catch(e){}
}
function _restoreNavGroups(){
  document.querySelectorAll('.nav-group').forEach(g=>{
    let v=null;try{v=localStorage.getItem('thp_nav_'+g.id);}catch(e){}
    const h=g.querySelector('.nav-grp-hdr'),b=g.querySelector('.nav-grp-body');
    if(!h||!b)return;
    const collapsed=v===null?h.classList.contains('collapsed'):v==='0';
    h.classList.toggle('collapsed',collapsed);
    b.classList.toggle('hidden',collapsed);
  });
}
function _hideEmptyNavGroups(){
  document.querySelectorAll('.nav-group').forEach(g=>{
    const vis=[...g.querySelectorAll('.nav-item')].filter(n=>getComputedStyle(n).display!=='none');
    g.classList.toggle('empty',vis.length===0);
  });
}
function showPanel(id,sbId,e){
  _syncNavForWidth();
  if(window.innerWidth<=768)setTimeout(closeAllSB,80);
  $(sbId).nextElementSibling.querySelectorAll('.panel').forEach(p=>p.classList.remove('active'));
  $(id).classList.add('active');
  document.querySelectorAll('#'+sbId+' .nav-item').forEach(n=>n.classList.remove('active'));
  if(e)e.currentTarget.classList.add('active');
  if(window.innerWidth<=768)$(sbId).classList.remove('open');
}
function toggleSB(id){
  const sb=$(id);if(!sb)return;
  const open=sb.classList.toggle('open');
  const bd=$('sb-backdrop');if(bd)bd.classList.toggle('on',open);
  if(open){_restoreNavGroups();_hideEmptyNavGroups();}
}
function closeAllSB(){
  document.querySelectorAll('.sidebar.open').forEach(e=>e.classList.remove('open'));
  const bd=$('sb-backdrop');if(bd)bd.classList.remove('on');
}
function closeModal(id){$(id).classList.remove('open');}
function selectLeaveType(el){APP.selectLeave(el);}

/* ── THEME — light only ── */
function toggleTheme(){ /* dark mode retired; kept so old handlers do not error */ }
(function initTheme(){
  document.documentElement.setAttribute('data-theme','light');
  try{localStorage.removeItem('thp_theme');}catch(e){}
})();

/* ── LOADING OVERLAY ── */
function showLoader(msg){
  const el=$('loading-overlay');if(!el)return;
  if(msg){const t=$('lo-text');if(t)t.textContent=msg;}
  el.classList.remove('fade-out');
  el.classList.add('active');
  // Never let a failed network call leave someone staring at this screen.
  clearTimeout(window._loaderGuard);
  window._loaderGuard=setTimeout(()=>{
    if(el.classList.contains('active')){
      hideLoader();
      const online=navigator.onLine;
      toast(online
        ? 'That took too long. Some data may not have loaded — try refreshing.'
        : 'No internet connection. Check your network and refresh.','err');
    }
  },20000);
}
function hideLoader(){
  clearTimeout(window._loaderGuard);
  const el=$('loading-overlay');if(!el)return;
  el.classList.add('fade-out');
  setTimeout(()=>{el.classList.remove('active','fade-out');},450);
}

/* ═══════════════════════════════════════════════
   4. API MODULE — Supabase Primary + GAS Mirror
   ─────────────────────────────────────────────
   Supabase (PostgreSQL) is the primary database.
   Google Sheets is synced every 6 hours as backup.
   GAS is still used for email notifications only.
   localStorage is a read cache.
═══════════════════════════════════════════════ */

/* ── SUPABASE CONFIG — UPDATE THESE ── */
/* ── SUPABASE CONFIG ── */
const SUPABASE={
  URL:'https://jhpqzkwzxprsnaczkyjq.supabase.co',
  KEY:'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImpocHF6a3d6eHByc25hY3preWpxIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzQxOTE4NTMsImV4cCI6MjA4OTc2Nzg1M30.GKJz9EhxGP1wTQBiufLoVLxWOstx-9Z0MPWHxj2c8VM',
};

/* ── GAS config (kept for email notifications only) ── */
const GAS_URL_KEY='thp_script_url';
const GAS_DEFAULT_URL='https://script.google.com/macros/s/AKfycbxYjyPS7HHfVCheKSUi-gYm_a02tpxhz4aleReROhkvE8Zv3dFxdkKAJzH16gHcIsD77g/exec';

const API={
  /* ── Supabase REST helpers ── */
  _headers(){
    return {
      'apikey':SUPABASE.KEY,
      'Authorization':'Bearer '+SUPABASE.KEY,
      'Content-Type':'application/json',
      'Prefer':'return=representation'
    };
  },
  lastError:'',
  async _supa(path,opts={}){
    try{
      const r=await fetch(SUPABASE.URL+'/rest/v1/'+path,{headers:this._headers(),...opts});
      if(!r.ok){
        const body=await r.text();
        let msg=body;
        try{const j=JSON.parse(body);msg=j.message||j.hint||j.details||body;}catch(e){}
        if(r.status===404||/does not exist/i.test(msg))msg='Database table/column missing — run the pending SQL migration in Supabase. ('+msg+')';
        this.lastError=msg;
        console.warn('Supabase error:',r.status,path,msg);
        return null;
      }
      this.lastError='';
      const text=await r.text();
      return text?JSON.parse(text):[];
    }catch(e){this.lastError=e.message||'Network error';console.warn('Supabase fetch:',e);return null;}
  },
  async _get(table,query=''){return this._supa(table+(query?'?'+query:''));},
  async _insert(table,data){return this._supa(table,{method:'POST',body:JSON.stringify(data)});},
  async _update(table,query,data){return this._supa(table+'?'+query,{method:'PATCH',body:JSON.stringify(data)});},
  async _delete(table,query){return this._supa(table+'?'+query,{method:'DELETE'});},
  async _upsert(table,data){return this._supa(table,{method:'POST',body:JSON.stringify(data),headers:{...this._headers(),'Prefer':'resolution=merge-duplicates,return=representation'}});},

  /* ── GAS helper (emails only) ── */
  getGasUrl(){return localStorage.getItem(GAS_URL_KEY)||GAS_DEFAULT_URL;},
  setGasUrl(url){localStorage.setItem(GAS_URL_KEY,url);},
  async gasPost(payload){
    const url=this.getGasUrl();if(!url)return null;
    try{
      const r=await fetch(url,{method:'POST',headers:{'Content-Type':'text/plain'},body:JSON.stringify(payload),redirect:'follow'});
      if(!r.ok)return null;return JSON.parse(await r.text());
    }catch(e){console.warn('GAS POST:',e);return null;}
  },

  showBar(state,msg){
    /* Small non-blocking mini-toast for save/update operations */
    if(state==='syncing')return; /* skip "syncing" — only show result */
    const type=state==='synced'?'ok':state==='error'?'err':'info';
    toast(msg,type);
  },

  /* ═══════════════════════════════════════════
     AUTH — Supabase sessions table
  ═══════════════════════════════════════════ */
  async login(id,pass){
    if(!id||!pass)return{success:false,error:'Missing credentials'};

    // Admin login
    if(id==='ADMIN01'){
      const settings=await this._get('settings','key=eq.admin_password');
      const adminPass=(settings&&settings[0])?settings[0].value:'admin123';
      if(String(pass)!==String(adminPass)){this.logLogin('ADMIN01','Administrator','failed','Incorrect password');return{success:false,error:'Incorrect password'};}
      this.logLogin('ADMIN01','Administrator','success','');
      const token=this._genToken();
      await this._cleanSessions(id);
      await this._insert('sessions',{staff_id:id,token,expires_at:new Date(Date.now()+12*3600000).toISOString()});
      return{success:true,user:{id:'ADMIN01',name:'Administrator',role:'admin'},token};
    }

    // Staff login
    const rows=await this._get('staff','id=eq.'+encodeURIComponent(id));
    if(!rows||!rows.length){this.logLogin(id,'','failed','Staff ID not found');return{success:false,error:'Staff ID not found'};}
    const s=rows[0];
    if(s.active===false){
      this.logLogin(id,s.name,'locked','Account deactivated');
      return{success:false,error:'This account has been deactivated. Please contact the Administrator.'};
    }
    if(String(s.password)!==String(pass)){this.logLogin(id,s.name,'failed','Incorrect password');return{success:false,error:'Incorrect password'};}
    this.logLogin(id,s.name,'success','');
    const token=this._genToken();
    await this._cleanSessions(id);
    await this._insert('sessions',{staff_id:id,token,expires_at:new Date(Date.now()+12*3600000).toISOString()});
    return{
      success:true,token,
      user:{id:s.id,name:s.name,unit:(s.unit||'').trim(),role:s.role||'staff',
        color:s.avatar_color||'',email:s.email||'',gender:s.gender||'male',
        supervisor:s.supervisor||'',phone:s.phone||'',emergencyContact:s.emergency_contact||''}
    };
  },

  async logLogin(staffId,name,outcome,reason){
    try{
      await this._insert('login_log',{id:'LG'+Date.now().toString(36)+Math.random().toString(36).slice(2,6),
        staff_id:staffId||'',name:name||'',outcome:outcome||'success',reason:reason||'',
        created_at:new Date().toISOString()});
    }catch(e){}
  },
  async validateSession(id,token){
    if(!id||!token)return{success:false,error:'No session'};
    const rows=await this._get('sessions','staff_id=eq.'+encodeURIComponent(id)+'&token=eq.'+encodeURIComponent(token));
    if(!rows||!rows.length)return{success:false,error:'Invalid session'};
    const sess=rows[0];
    if(new Date(sess.expires_at)<new Date())return{success:false,error:'Session expired'};
    if(id==='ADMIN01')return{success:true,user:{id:'ADMIN01',name:'Administrator',role:'admin'}};
    const staff=await this._get('staff','id=eq.'+encodeURIComponent(id));
    if(!staff||!staff.length)return{success:false,error:'Staff not found'};
    const s=staff[0];
    return{success:true,user:{id:s.id,name:s.name,unit:(s.unit||'').trim(),role:s.role||'staff',
      color:s.avatar_color||'',email:s.email||'',gender:s.gender||'male',
      supervisor:s.supervisor||'',phone:s.phone||'',emergencyContact:s.emergency_contact||''}};
  },

  async logout(id,token){
    if(id)await this._cleanSessions(id);
    clearSession();return{success:true};
  },

  async _cleanSessions(id){await this._delete('sessions','staff_id=eq.'+encodeURIComponent(id));},
  _genToken(){let t='';for(let i=0;i<32;i++)t+=Math.floor(Math.random()*256).toString(16);return t+Date.now().toString(36);},

  /* ═══════════════════════════════════════════
     ATTENDANCE — Supabase attendance table
  ═══════════════════════════════════════════ */
  async saveRecord(rec){
    this.showBar('syncing','Saving…');
    // Server-side duplicate check — prevent multiple clock-ins on same day
    const todayCheck=await this._get('attendance','staff_id=eq.'+encodeURIComponent(rec.id)+'&date=eq.'+encodeURIComponent(rec.date)+'&limit=1');
    if(todayCheck&&todayCheck.length){
      this.showBar('error','Already clocked in today');return{success:false,duplicate:true};
    }
    const r=await this._insert('attendance',{
      date:rec.date,staff_id:rec.id,name:rec.name,unit:(rec.unit||'').trim(),
      clock_in:rec.in,clock_out:rec.out||null,hours:rec.hours||null,status:rec.status||'Active',
      work_mode:rec.work_mode||'Office'
    });
    if(r){this.showBar('synced','Saved ✓');return{success:true};}
    this.showBar('error','Save failed');return{success:false};
  },

  async updateRecord(rec){
    this.showBar('syncing','Updating…');
    const rows=await this._get('attendance','staff_id=eq.'+encodeURIComponent(rec.id)+'&clock_in=eq.'+encodeURIComponent(rec.in)+'&limit=1');
    if(rows&&rows.length){
      await this._update('attendance','id=eq.'+rows[0].id,{
        clock_out:rec.out||null,hours:rec.hours||null,status:rec.status||''
      });
      this.showBar('synced','Updated ✓');return{success:true};
    }
    this.showBar('error','Update failed');return{success:false};
  },

  /* ═══════════════════════════════════════════
     STAFF — Supabase staff table
  ═══════════════════════════════════════════ */
  async saveStaff(id,data){
    const r=await this._upsert('staff',[{
      id,name:data.name,unit:(data.unit||'').trim(),role:data.role||'staff',
      password:data.pass,avatar_color:data.color||'',email:data.email||'',
      gender:data.gender||'male',supervisor:data.supervisor||'',
      phone:data.phone||'',emergency_contact:data.emergencyContact||'',
      contract_start:data.contractStart||null,contract_end:data.contractEnd||null
    }]);
    return r?{success:true}:{success:false};
  },
  /* ── Update only contract dates ── */
  async updateContract(id,start,end){
    this.showBar('syncing','Saving contract…');
    const r=await this._update('staff','id=eq.'+encodeURIComponent(id),{
      contract_start:start||null,contract_end:end||null
    });
    if(r!==null){this.showBar('synced','Contract saved ✓');return{success:true};}
    this.showBar('error','Save failed');return{success:false};
  },
  /* ── Secure access to RLS-locked tables (Phase 1) ──
     payroll_staff, payroll_runs, payroll_settings and hr_cases are
     locked at the database level, so the browser key cannot reach
     them. These go through Apps Script, which holds the service key
     server-side and verifies the caller's session before responding. ── */
  _sessTok(){
    try{const sess=JSON.parse(localStorage.getItem('thp_session')||'null');return sess?.token||'';}catch(e){return '';}
  },
  async secureGet(table,query){
    const u=APP?.user;
    if(!u){this.lastError='Not signed in';return null;}
    if(!this._sessTok()){this.lastError='Session expired — sign out and sign in again';return null;}
    const r=await this.gasPost({action:'secureGet',table,query:query||'',staffId:u.id,token:this._sessTok()});
    if(r===null){this.lastError='No response from Apps Script — check the deployment';return null;}
    if(!r.success){this.lastError=r.error||'The server refused the request';return null;}
    return r.rows||[];
  },
  async secureSave(table,rows){
    const u=APP?.user;
    if(!u){this.lastError='You are not signed in. Refresh and sign in again.';return null;}
    if(!this._sessTok()){this.lastError='Your session has expired. Sign out and sign in again.';return null;}
    if(!this.getGasUrl()){this.lastError='The Apps Script URL is not configured — see Google Sheets in the admin panel.';return null;}
    const r=await this.gasPost({action:'secureSave',table,rows,staffId:u.id,token:this._sessTok()});
    if(r===null){this.lastError='No response from Apps Script. Check the connection, and that the script is deployed.';return null;}
    if(!r.success){this.lastError=r.error||'The server rejected the save.';return null;}
    return r.rows||[];
  },
  async secureDelete(table,query){
    const u=APP?.user;if(!u)return null;
    const r=await this.gasPost({action:'secureDelete',table,query,staffId:u.id,token:this._sessTok()});
    if(!r||!r.success){this.lastError=(r&&r.error)||'Secure delete failed';return null;}
    return true;
  },

  /* ── HR Staff Files ── */
  async getHRFile(id){
    const r=await this._get('hr_staff_files','staff_id=eq.'+encodeURIComponent(id));
    return (r&&r.length)?r[0]:null;
  },
  async getAllHRFiles(){
    const r=await this._get('hr_staff_files','select=staff_id,phone,dob,next_of_kin,ssnit_number,bank_account');
    return r||[];
  },
  async saveHRFile(id,data){
    this.showBar('syncing','Saving file…');
    const r=await this._upsert('hr_staff_files',[{staff_id:id,...data,updated_at:new Date().toISOString()}]);
    if(r){this.showBar('synced','File saved ✓');return{success:true};}
    this.showBar('error','Save failed');return{success:false};
  },
  async deleteStaff(id){
    await this._delete('staff','id=eq.'+encodeURIComponent(id));
    return{success:true};
  },

  /* ── Self-service profile update ── */
  async updateProfile(id,data){
    this.showBar('syncing','Updating profile…');
    const r=await this._update('staff','id=eq.'+encodeURIComponent(id),{
      email:data.email||'',phone:data.phone||'',emergency_contact:data.emergencyContact||''
    });
    if(r!==null){this.showBar('synced','Profile updated ✓');return{success:true};}
    this.showBar('error','Update failed');return{success:false};
  },

  /* ═══════════════════════════════════════════
     LEAVE — Supabase leave_requests table
  ═══════════════════════════════════════════ */
  async applyLeave(leave){
    const r=await this._insert('leave_requests',{
      id:leave.id,staff_id:leave.staffId,name:leave.name,unit:(leave.unit||'').trim(),
      type:leave.type,start_date:leave.startDate,end_date:leave.endDate,days:leave.days,
      reason:leave.reason||'',sick_note:leave.sickNote||'',staff_email:leave.staffEmail||'',
      supervisor_id:leave.supervisorId||'',supervisor_status:leave.supervisorStatus||'Pending',
      final_approver_id:leave.finalApproverId||'',
      final_approver_status:leave.finalApproverStatus||'Waiting',
      overall_status:leave.status||'Pending',
      handover_note:leave.handoverNote||'',comp_ref:leave.compRef||''
    });
    if(!r)return{success:false};
    // Trigger email notification via GAS — include all recipient emails
    const emailPayload={action:'applyLeave',leave:{...leave,
      supervisorEmail:leave._supervisorEmail||'',
      finalApproverEmail:leave._finalApproverEmail||''
    }};
    this.gasPost(emailPayload).catch(()=>{});
    return{success:true,leaveId:leave.id};
  },

  async updateLeave(id,status,note,stage,extraEmailData){
    const isFinal=(stage==='final'||stage==='hr');
    const update=isFinal
      ?{final_approver_status:status,final_approver_note:note||'',overall_status:status,updated_at:new Date().toISOString()}
      :status==='Rejected'
        ?{supervisor_status:status,supervisor_note:note||'',final_approver_status:'N/A',overall_status:'Rejected',updated_at:new Date().toISOString()}
        :{supervisor_status:status,supervisor_note:note||'',final_approver_status:'Pending',overall_status:'Pending',updated_at:new Date().toISOString()};
    const r=await this._update('leave_requests','id=eq.'+encodeURIComponent(id),update);
    if(r===null)return{success:false};
    // Trigger email via GAS — include recipient emails
    const emailPayload={action:'updateLeave',id,status,note,stage,...(extraEmailData||{})};
    this.gasPost(emailPayload).catch(()=>{});
    return{success:true};
  },

  /* ═══════════════════════════════════════════
     PASSWORD — Supabase staff.password
  ═══════════════════════════════════════════ */
  async changePassword(id,oldPass,newPass,token){
    if(id==='ADMIN01'){
      const settings=await this._get('settings','key=eq.admin_password');
      const adminPass=(settings&&settings[0])?settings[0].value:'admin123';
      if(String(oldPass)!==String(adminPass))return{success:false,error:'Incorrect current password'};
      await this._upsert('settings',[{key:'admin_password',value:newPass,updated_at:new Date().toISOString()}]);
      return{success:true};
    }
    const rows=await this._get('staff','id=eq.'+encodeURIComponent(id));
    if(!rows||!rows.length)return{success:false,error:'Staff not found'};
    if(String(rows[0].password)!==String(oldPass))return{success:false,error:'Incorrect current password'};
    await this._update('staff','id=eq.'+encodeURIComponent(id),{password:newPass});
    return{success:true};
  },

  /* ── Forgot Password — generate temp pass, save to Supabase, email via GAS ── */
  async resetPassword(staffId){
    if(!staffId)return{success:false,error:'Staff ID required'};
    if(staffId==='ADMIN01')return{success:false,error:'Admin password cannot be reset this way. Contact the system administrator.'};
    const rows=await this._get('staff','id=eq.'+encodeURIComponent(staffId));
    if(!rows||!rows.length)return{success:false,error:'Staff ID not found in the system.'};
    const staff=rows[0];
    const email=(staff.email||'').trim();
    if(!email)return{success:false,error:'No email registered for this account. Please contact the Administrator to reset your password.'};
    // Generate a 6-character temporary password
    const chars='ABCDEFGHJKMNPQRSTUVWXYZ23456789';
    let tempPass='';for(let i=0;i<6;i++)tempPass+=chars[Math.floor(Math.random()*chars.length)];
    // Save temp password to Supabase (plain text — user will be forced to change on login)
    await this._update('staff','id=eq.'+encodeURIComponent(staffId),{password:tempPass});
    // Send email via GAS
    const emailResult=await this.gasPost({
      action:'resetPassword',
      staffId,
      staffName:staff.name,
      staffEmail:email,
      tempPassword:tempPass
    }).catch(()=>null);
    return{success:true,email:email.replace(/(.{2})(.*)(@.*)/, '$1***$3'),emailSent:!!emailResult};
  },

  /* ═══════════════════════════════════════════
     HOLIDAYS — Supabase holidays table
  ═══════════════════════════════════════════ */
  async getHolidays(){
    const r=await this._get('holidays','order=date');
    return r?{success:true,holidays:r.map(h=>({id:h.id,name:h.name,date:h.date,type:h.type,recurring:h.recurring,year:h.year,createdAt:h.created_at}))}:{success:false};
  },
  async saveHoliday(holiday){
    const r=await this._upsert('holidays',[{
      id:holiday.id||('HOL'+Date.now()),name:holiday.name,date:holiday.date,
      type:holiday.type||'custom',recurring:holiday.recurring||'no',year:holiday.year||''
    }]);
    return r?{success:true,holidayId:(r[0]||{}).id}:{success:false};
  },
  async deleteHoliday(holidayId){
    await this._delete('holidays','id=eq.'+encodeURIComponent(holidayId));
    return{success:true};
  },

  /* ── Seed Ghana holidays (client-side, inserts into Supabase) ── */
  async seedGhanaHolidays(year){
    if(!year)year=new Date().getFullYear();
    const pad=n=>String(n).padStart(2,'0');
    const iso=d=>`${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`;
    const easter=easterSunday(year);
    const gf=new Date(easter);gf.setDate(easter.getDate()-2);
    const em=new Date(easter);em.setDate(easter.getDate()+1);
    const eids=estimateEidDates(year);
    const fd=farmersDayISO(year);
    const holidays=[
      {name:"New Year's Day",date:year+'-01-01',type:'fixed',recurring:'yes'},
      {name:'Constitution Day',date:year+'-01-07',type:'fixed',recurring:'yes'},
      {name:'Independence Day',date:year+'-03-06',type:'fixed',recurring:'yes'},
      {name:'Good Friday',date:iso(gf),type:'fixed',recurring:'yes'},
      {name:'Easter Monday',date:iso(em),type:'fixed',recurring:'yes'},
      {name:'May Day',date:year+'-05-01',type:'fixed',recurring:'yes'},
      {name:'Republic Day',date:year+'-07-01',type:'fixed',recurring:'yes'},
      {name:"Founders' Day",date:year+'-08-04',type:'fixed',recurring:'yes'},
      {name:'Kwame Nkrumah Memorial Day',date:year+'-09-21',type:'fixed',recurring:'yes'},
      {name:"Farmer's Day",date:fd,type:'fixed',recurring:'no'},
      {name:'Christmas Day',date:year+'-12-25',type:'fixed',recurring:'yes'},
      {name:'Boxing Day',date:year+'-12-26',type:'fixed',recurring:'yes'},
      {name:'Eid al-Fitr (estimated)',date:eids.eidFitr,type:'custom',recurring:'no'},
      {name:'Eid al-Adha (estimated)',date:eids.eidAdha,type:'custom',recurring:'no'},
    ];
    const existing=await this._get('holidays','year=eq.'+year);
    const existingDates=new Set((existing||[]).map(h=>h.date+'_'+h.name));
    let added=0,skipped=0;
    for(const h of holidays){
      if(existingDates.has(h.date+'_'+h.name)){skipped++;continue;}
      await this._insert('holidays',{id:'GH'+year+'_'+(added+skipped+1),name:h.name,date:h.date,type:h.type,recurring:h.recurring,year:String(year)});
      added++;
    }
    return{success:true,added,skipped,year};
  },

  /* ═══════════════════════════════════════════
     SICK NOTE UPLOAD — Supabase Storage
  ═══════════════════════════════════════════ */
  async uploadSickNote(leaveId,fileName,fileData,mimeType){
    try{
      // Decode base64 to blob
      const byteChars=atob(fileData);
      const byteArr=new Uint8Array(byteChars.length);
      for(let i=0;i<byteChars.length;i++)byteArr[i]=byteChars.charCodeAt(i);
      const blob=new Blob([byteArr],{type:mimeType||'application/octet-stream'});

      // Sanitize filename
      const safeName=fileName.replace(/[^a-zA-Z0-9._-]/g,'_');
      const path=leaveId+'/'+safeName;
      const r=await fetch(SUPABASE.URL+'/storage/v1/object/sick-notes/'+path,{
        method:'POST',
        headers:{
          'apikey':SUPABASE.KEY,
          'Authorization':'Bearer '+SUPABASE.KEY,
          'Content-Type':mimeType||'application/octet-stream',
          'x-upsert':'true'
        },
        body:blob
      });
      if(!r.ok){
        const errText=await r.text().catch(()=>'');
        console.warn('Sick note upload failed:',r.status,errText);
        // Still save the filename in the leave record
        await this._update('leave_requests','id=eq.'+encodeURIComponent(leaveId),{sick_note:fileName});
        toast('File reference saved, but storage upload failed','err');
        return{success:false};
      }

      const fileUrl=SUPABASE.URL+'/storage/v1/object/public/sick-notes/'+path;
      // Update leave record with file URL
      await this._update('leave_requests','id=eq.'+encodeURIComponent(leaveId),{sick_note:fileName+' | '+fileUrl});
      toast('Document uploaded ✓');
      return{success:true,fileUrl,downloadUrl:fileUrl,fileName,leaveId};
    }catch(e){
      console.warn('Upload error:',e);
      // Fallback — save filename only
      await this._update('leave_requests','id=eq.'+encodeURIComponent(leaveId),{sick_note:fileName}).catch(()=>{});
      toast('Upload failed — file reference saved','err');
      return{success:false};
    }
  },

  /* ═══════════════════════════════════════════
     HYDRATE — single call, loads all data
  ═══════════════════════════════════════════ */
  async hydrate(){
    const[staffRows,attRows,leaveRows,holRows,setRows]=await Promise.all([
      this._get('staff','order=name'),
      this._get('attendance','order=id.desc&limit=5000'),
      this._get('leave_requests','order=applied_at.desc&limit=2000'),
      this._get('holidays','order=date'),
      this._get('settings','order=key')
    ]);
    if(!staffRows)return{success:false};

    // Transform staff rows to {id: {name,unit,...}} format
    const staff={};
    (staffRows||[]).forEach(s=>{
      staff[s.id]={name:s.name,unit:(s.unit||'').trim(),role:s.role||'staff',pass:s.password,
        color:s.avatar_color||'',email:s.email||'',gender:s.gender||'male',
        supervisor:s.supervisor||'',phone:s.phone||'',emergencyContact:s.emergency_contact||'',
        contractStart:s.contract_start||'',contractEnd:s.contract_end||'',
        position:s.position||'',employmentType:s.employment_type||'Staff',dateJoined:s.date_joined||'',
        active:s.active!==false,exitDate:s.exit_date||'',exitReason:s.exit_reason||''};
    });

    // Transform attendance rows
    const records=(attRows||[]).map(r=>({
      date:r.date,id:r.staff_id,name:r.name,unit:r.unit,
      in:r.clock_in,out:r.clock_out||null,hours:r.hours||null,status:r.status||'Active',
      work_mode:r.work_mode||'Office'
    }));

    // Transform leave rows
    const leave=(leaveRows||[]).map(r=>({
      id:r.id,staffId:r.staff_id,name:r.name,unit:(r.unit||'').trim(),type:r.type,
      startDate:r.start_date,endDate:r.end_date,days:r.days,reason:r.reason,sickNote:r.sick_note,
      staffEmail:r.staff_email||'',
      supervisorId:r.supervisor_id||'',supervisorStatus:r.supervisor_status||'Pending',supervisorNote:r.supervisor_note||'',
      finalApproverId:r.final_approver_id||'',finalApproverStatus:r.final_approver_status||'Pending',finalApproverNote:r.final_approver_note||'',
      status:r.overall_status||'Pending',hrStatus:r.final_approver_status||r.overall_status||'Pending',hrNote:r.final_approver_note||'',
      appliedAt:r.applied_at||'',updatedAt:r.updated_at||'',
      handoverNote:r.handover_note||'',compRef:r.comp_ref||''
    }));

    // Transform holidays
    const holidays=(holRows||[]).map(h=>({id:h.id,name:h.name,date:h.date,type:h.type,recurring:h.recurring,year:h.year,createdAt:h.created_at}));

    // Transform settings
    const settings={};
    (setRows||[]).forEach(r=>{settings[r.key]=r.value;});

    // Cache locally
    localStorage.setItem('thp_staff',JSON.stringify(staff));
    localStorage.setItem('thp_recs',JSON.stringify(records));
    localStorage.setItem('thp_leave',JSON.stringify(leave));
    localStorage.setItem('thp_holidays',JSON.stringify(holidays));

    return{success:true,staff,records,leave,holidays,settings};
  },

  /* ── Connection status ── */
  updateChips(){
    const ok=!!SUPABASE.URL&&SUPABASE.URL!=='https://YOUR_PROJECT_ID.supabase.co';
    ['st-sync-chip','mgr-sync-chip','ad-sync-chip'].forEach(id=>{
      const el=$(id);if(!el)return;
      el.className='sync-pill '+(ok?'live':'no-url');
      el.textContent=ok?'⬤ Supabase connected':'⬤ Not configured';
    });
    if($('conn-badge')){$('conn-badge').className='badge '+(ok?'b-ok':'b-warn');$('conn-badge').textContent=ok?'✓ Supabase':'⚠ Not Connected';}
  },

  /* ── GAS URL management (for admin Google Sheets panel) ── */
  saveUrl(inputId){
    const url=$(inputId).value.trim();
    if(!url)return toast('Please enter a URL','err');
    if(!url.includes('script.google.com'))return toast('Not a valid Apps Script URL','err');
    this.setGasUrl(url);
    if($('script-url-input'))$('script-url-input').value=url;
    toast('GAS URL saved (used for email notifications)');
  },
  dismissBanner(){$('setup-banner').style.display='none';localStorage.setItem('thp_banner_dismissed','1');},

  async testConnection(){
    const el=$('sync-result');if(el)el.textContent='Testing Supabase…';
    const r=await this._get('staff','limit=1');
    if(r!==null){
      if(el)el.innerHTML='<span style="color:var(--green)">✅ Supabase connected! ('+((r||[]).length?'data found':'empty')+')</span>';
      toast('Supabase connection successful!');
    } else {
      if(el)el.innerHTML='<span style="color:var(--red)">❌ Failed. Check Supabase URL and key in app.js.</span>';
      toast('Connection failed','err');
    }
  },

  async pullFromSheets(){
    toast('Data is now served from Supabase. Use the Supabase dashboard to manage data.','info');
  },
  async pushAllToSheets(){
    toast('GAS sync runs automatically every 6 hours. Run syncAllFromSupabase() manually in Apps Script if needed.','info');
  }
};

// Legacy alias so HTML onclick="SYNC.xxx" still works
const SYNC=API;

/* ═══════════════════════════════════════════════
   5. GHANA PUBLIC HOLIDAYS (Enhanced)
   ─────────────────────────────────────────────
   Merges:
   a) Built-in fixed holidays (always available offline)
   b) Admin-managed holidays from the Holidays sheet
   c) Estimated Eid dates & Farmer's Day
   d) Government-declared extensions/one-offs
═══════════════════════════════════════════════ */
function easterSunday(year){
  const a=year%19,b=Math.floor(year/100),c=year%100;
  const d=Math.floor(b/4),e=b%4,f=Math.floor((b+8)/25);
  const g=Math.floor((b-f+1)/3),h=(19*a+b-d-g+15)%30;
  const i=Math.floor(c/4),k=c%4,l=(32+2*e+2*i-h-k)%7;
  const m=Math.floor((a+11*h+22*l)/451);
  const month=Math.floor((h+l-7*m+114)/31);
  const day=((h+l-7*m+114)%31)+1;
  return new Date(year, month-1, day);
}

// Farmer's Day: first Friday of December
function farmersDayISO(year){
  const dec1=new Date(year,11,1);
  const dow=dec1.getDay();
  let fridayDate;
  if(dow===5) fridayDate=1;
  else if(dow<5) fridayDate=1+(5-dow);
  else fridayDate=1+(5+7-dow);
  return `${year}-12-${String(fridayDate).padStart(2,'0')}`;
}

// Estimated Eid dates (approximate — shifts ~10.6 days earlier/year)
// Reference: Eid al-Fitr 2024 ≈ Apr 10, Eid al-Adha 2024 ≈ Jun 17
function estimateEidDates(year){
  const pad=n=>String(n).padStart(2,'0');
  const iso=d=>`${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`;
  const refFitr=new Date(2024,3,10),refAdha=new Date(2024,5,17);
  const shift=Math.round((year-2024)*-10.6);
  const estFitr=new Date(year,refFitr.getMonth(),refFitr.getDate()+shift);
  const estAdha=new Date(year,refAdha.getMonth(),refAdha.getDate()+shift);
  return{eidFitr:iso(estFitr),eidAdha:iso(estAdha)};
}

// Built-in Ghana holidays (always available even without server)
function ghBuiltinHolidayISOs(year){
  const easter=easterSunday(year);
  const gf=new Date(easter);gf.setDate(easter.getDate()-2);
  const em=new Date(easter);em.setDate(easter.getDate()+1);
  const pad=n=>String(n).padStart(2,'0');
  const iso=d=>`${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`;
  const eids=estimateEidDates(year);
  const holidays=new Set([
    `${year}-01-01`,               // New Year
    `${year}-01-07`,               // Constitution Day
    `${year}-03-06`,               // Independence Day
    iso(gf),                       // Good Friday
    iso(em),                       // Easter Monday
    `${year}-05-01`,               // May Day
    `${year}-07-01`,               // Republic Day
    `${year}-08-04`,               // Founders Day
    `${year}-09-21`,               // Kwame Nkrumah Memorial Day
    farmersDayISO(year),           // Farmer's Day (1st Friday Dec)
    `${year}-12-25`,               // Christmas
    `${year}-12-26`,               // Boxing Day
    eids.eidFitr,                  // Eid al-Fitr (estimated)
    eids.eidAdha,                  // Eid al-Adha (estimated)
  ]);
  // Postponed / cancelled holidays for a specific year — removed from the block list.
  // Add the ORIGINAL date here; put the new observed date in the Holidays admin panel.
  const HOLIDAY_EXCEPTIONS=[
    '2026-07-01',                  // Republic Day 2026 postponed to 2026-07-03
    '2026-08-04',                  // Founders' Day 2026 — THP-Ghana working day
  ];
  HOLIDAY_EXCEPTIONS.forEach(d=>holidays.delete(d));
  return holidays;
}

// Named holiday lookup for display purposes
function ghHolidayNames(year){
  const easter=easterSunday(year);
  const gf=new Date(easter);gf.setDate(easter.getDate()-2);
  const em=new Date(easter);em.setDate(easter.getDate()+1);
  const pad=n=>String(n).padStart(2,'0');
  const iso=d=>`${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`;
  const eids=estimateEidDates(year);
  return {
    [`${year}-01-01`]:"New Year's Day",
    [`${year}-01-07`]:'Constitution Day',
    [`${year}-03-06`]:'Independence Day',
    [iso(gf)]:'Good Friday',
    [iso(em)]:'Easter Monday',
    [`${year}-05-01`]:'May Day',
    [`${year}-07-01`]:'Republic Day',
    [`${year}-08-04`]:"Founders' Day",
    [`${year}-09-21`]:'Kwame Nkrumah Memorial Day',
    [farmersDayISO(year)]:"Farmer's Day",
    [`${year}-12-25`]:'Christmas Day',
    [`${year}-12-26`]:'Boxing Day',
    [eids.eidFitr]:'Eid al-Fitr (est.)',
    [eids.eidAdha]:'Eid al-Adha (est.)',
  };
}

// Merge built-in + admin-managed holidays
function getAllHolidayISOs(year,adminHolidays){
  const builtIn=ghBuiltinHolidayISOs(year);
  const all=new Set(builtIn);
  if(adminHolidays&&adminHolidays.length){
    adminHolidays.forEach(h=>{
      if(!h.date)return;
      const d=h.date.slice(0,10); // YYYY-MM-DD
      const hYear=parseInt(d.slice(0,4));
      if(h.recurring==='yes'||hYear===year) all.add(d);
    });
  }
  return all;
}

// Get all holiday names (built-in + admin) for display
function getAllHolidayNamesMap(year,adminHolidays){
  const names=ghHolidayNames(year);
  if(adminHolidays&&adminHolidays.length){
    adminHolidays.forEach(h=>{
      if(!h.date)return;
      const d=h.date.slice(0,10);
      const hYear=parseInt(d.slice(0,4));
      if(h.recurring==='yes'||hYear===year) names[d]=h.name;
    });
  }
  return names;
}

// Legacy compatibility — these now use admin holidays from APP.holidays
function ghHolidayISOs(year){
  return getAllHolidayISOs(year,(typeof APP!=='undefined')?APP.holidays:[]);
}
function isHoliday(dateObj){
  const year=dateObj.getFullYear();
  const iso=`${year}-${String(dateObj.getMonth()+1).padStart(2,'0')}-${String(dateObj.getDate()).padStart(2,'0')}`;
  return ghHolidayISOs(year).has(iso);
}
function getHolidayName(dateObj){
  const year=dateObj.getFullYear();
  const iso=`${year}-${String(dateObj.getMonth()+1).padStart(2,'0')}-${String(dateObj.getDate()).padStart(2,'0')}`;
  const names=getAllHolidayNamesMap(year,(typeof APP!=='undefined')?APP.holidays:[]);
  return names[iso]||null;
}
function isWeekend(dateObj){const d=dateObj.getDay();return d===0||d===6;}
function isWorkingDay(dateObj){return !isWeekend(dateObj)&&!isHoliday(dateObj);}
function workingDaysBetween(startStr,endStr){
  const s=new Date(startStr),e=new Date(endStr);
  let count=0,cur=new Date(s);
  while(cur<=e){if(isWorkingDay(cur))count++;cur.setDate(cur.getDate()+1);}
  return count;
}
function leaveOnDate(leaveArr,staffId,dateStr){
  const dt=new Date(dateStr);if(isNaN(dt))return null;
  return leaveArr.find(l=>{
    if(l.staffId!==staffId)return false;
    if(l.status!=='Approved')return false;
    const s=new Date(l.startDate),e=new Date(l.endDate);
    return dt>=s&&dt<=e;
  })||null;
}

/* ═══════════════════════════════════════════════
   6. LEAVE CONFIGURATION & HELPERS
═══════════════════════════════════════════════ */
const LEAVE_LIMITS={'Annual Leave':24,'Sick Leave':null,'Maternity Leave':65,'Paternity Leave':5,'Compassionate Leave':5,'Compensatory Leave':null};
function _leaveProgress(lv){
  let score=0;
  if(lv.supervisorStatus==='Approved')score+=2;
  else if(lv.supervisorStatus==='Rejected')score+=2;
  else if(lv.supervisorStatus==='N/A')score+=1;
  if(lv.finalApproverStatus==='Approved')score+=4;
  else if(lv.finalApproverStatus==='Rejected')score+=4;
  else if(lv.finalApproverStatus==='Pending')score+=2;
  if(lv.status==='Approved'||lv.status==='Rejected')score+=8;
  return score;
}
const HR_MANAGER_ID='THPG/03/2008';
const COUNTRY_LEADER_ID='THPG/12/2024';
const DIRECT_TO_CL=['THPG/08/2025','THPG/03/2008','THPG/05/2010','THPG/05/2025','THPG/09/2010','THPG/12/2024'];
const SUPERVISOR_ROLES=['manager','country_leader'];
const ADMIN_ID='ADMIN01';
function isManagerRole(role){return role==='manager'||role==='country_leader';}
function getAdminPass(){return localStorage.getItem('thp_admin_pass')||'admin123';}

/* ═══════════════════════════════════════════════
   7. APP CLASS — Server-First Architecture
═══════════════════════════════════════════════ */
class App{
  constructor(){
    /* Load from cache (server will overwrite on login/hydrate) */
    this.records=JSON.parse(localStorage.getItem('thp_recs'))||[];
    this.staff=JSON.parse(localStorage.getItem('thp_staff')||'{}');
    this.leave=JSON.parse(localStorage.getItem('thp_leave'))||[];
    this.holidays=JSON.parse(localStorage.getItem('thp_holidays'))||[];
    this.user=null;this.qrSid=null;this.HOURS=8;
    this._adFilter={status:''};
    this._mgrFilter={status:''};
    this._stFilter={status:''};
    this._sort={ad:{col:'date',dir:'desc'},mgr:{col:'date',dir:'desc'},st:{col:'date',dir:'desc'}};
    this._clock();this._qrParam();this._initBanner();API.updateChips();
  }
  /* Cache writes — these update localStorage (read cache) */
  _cacheR(){localStorage.setItem('thp_recs',JSON.stringify(this.records));}
  _cacheS(){localStorage.setItem('thp_staff',JSON.stringify(this.staff));}
  _cacheL(){localStorage.setItem('thp_leave',JSON.stringify(this.leave));}
  _cacheH(){localStorage.setItem('thp_holidays',JSON.stringify(this.holidays));}
  /* Legacy aliases */
  _saveR(){this._cacheR();}
  _saveS(){this._cacheS();}
  _saveL(){this._cacheL();}

  _initBanner(){
    if(!API.getGasUrl()) API.setGasUrl(GAS_DEFAULT_URL);
    if($('script-url-input')) $('script-url-input').value=API.getGasUrl();
    if($('banner-url')) $('banner-url').value=API.getGasUrl();
    $('setup-banner').style.display='none';
    localStorage.setItem('thp_banner_dismissed','1');
    API.updateChips();
  }
  _clock(){
    const t=()=>{
      const n=new Date(),ts=n.toLocaleTimeString([],{hour:'2-digit',minute:'2-digit',second:'2-digit'}),ds=n.toLocaleDateString('en-GB',{weekday:'long',year:'numeric',month:'long',day:'numeric'});
      ['st-time','m-time','qr-clock'].forEach(id=>{const e=$(id);if(e)e.textContent=ts;});
      ['st-date-hdr','st-date-sub','m-date-hdr','m-date-sub','ad-date','mgr-date','qr-date'].forEach(id=>{const e=$(id);if(e)e.textContent=ds;});
    };t();setInterval(t,1000);
  }
  _qrParam(){
    const sid=new URLSearchParams(window.location.search).get('staff');
    if(sid){this.qrSid=sid;
      // Hydrate staff data for QR landing
      API.get('getStaff').then(r=>{
        if(r&&r.staff&&r.staff[sid]){
          this.staff=r.staff;this._cacheS();
          $('qr-greet').textContent='Hello, '+r.staff[sid].name+'!';
          showView('qr-landing-view');
        }
      });
    }
  }

  /* ── QR clock ── */
  async qrIn(){
    const now=new Date();
    if(isWeekend(now))return toast('Not allowed on weekends.','err');
    const dept=$('qr-dept').value;if(!dept)return toast('Select unit','err');
    const s=this.staff[this.qrSid];if(!s)return toast('Staff not found','err');
    const rec={date:fmtD(now.toISOString()),id:this.qrSid,name:s.name,unit:s.unit||dept,in:now.toISOString(),out:null,hours:null,status:'Active'};
    /* Server first */
    const r=await API.saveRecord(rec);
    if(r&&r.success){
      this.records.push(rec);this._cacheR();
      $('qr-msg').innerHTML='<span style="color:var(--green)">✅ Clocked in at '+fmtT(now.toISOString())+'</span>';
    } else {
      toast('Failed to clock in — server error','err');
    }
  }
  async qrOut(){
    const rec=this.records.find(r=>r.id===this.qrSid&&!r.out);if(!rec)return toast('No active session','err');
    const now=new Date(),hrs=(now-new Date(rec.in))/3600000;
    rec.out=now.toISOString();rec.hours=fx(hrs);rec.status=hrs>=this.HOURS?'Completed':'Early Exit';
    const r=await API.updateRecord(rec);
    if(r&&r.success){
      this._cacheR();
      $('qr-msg').innerHTML='<span style="color:var(--teal)">✅ Clocked out — '+fx(hrs)+' hrs</span>';
    }
  }

  /* ═══════════════════════════════════════════
     LOGIN — SERVER-FIRST
     The server validates credentials and returns
     a session token + user object. No local
     password checking at all.
  ═══════════════════════════════════════════ */
  async loginAuto(){
    const id=$('uni-id').value.trim().toUpperCase();
    const pass=$('uni-pass').value;
    const errEl=$('lc-err');
    const btn=document.querySelector('.lc-btn');
    const setErr=(msg)=>{if(errEl){errEl.textContent=msg;errEl.style.animation='none';void errEl.offsetWidth;errEl.style.animation='errShake .35s ease';}};
    if(!id||!pass){setErr('Please enter your Staff ID and password.');return;}

    // Rate limiting (client-side courtesy — real security is server-side)
    if(!this._loginAttempts)this._loginAttempts={};
    const now=Date.now();
    if(!this._loginAttempts[id])this._loginAttempts[id]=[];
    this._loginAttempts[id]=this._loginAttempts[id].filter(t=>now-t<120000);
    if(this._loginAttempts[id].length>=5){
      const secsLeft=Math.ceil((120000-(now-this._loginAttempts[id][0]))/1000);
      setErr(`Too many attempts. Try again in ${secsLeft}s.`);return;
    }

    if(btn){btn.classList.add('loading');btn.querySelector('span').textContent='Signing in…';}

    /* ── Hash the password before sending (server stores hashed passwords) ── */
    const hashed=await hashPass(id,pass);

    /* ── Call server login ── */
    const result=await API.login(id, hashed);

    if(btn){btn.classList.remove('loading');btn.querySelector('span').textContent='Sign In';}

    if(!result){
      /* Network error — try plain password as fallback for first-time/default passwords */
      const fallback=await API.login(id, pass);
      if(!fallback||!fallback.success){
        this._loginAttempts[id].push(Date.now());
        setErr('Could not reach server. Check your connection.');return;
      }
      // Server accepted plain password — hash and update
      this._afterLogin(fallback, id, pass);
      return;
    }

    if(!result.success){
      /* Server rejected — try with plain password (legacy/default passwords) */
      const fallback=await API.login(id, pass);
      if(fallback&&fallback.success){
        this._afterLogin(fallback, id, pass);
        return;
      }
      this._loginAttempts[id].push(Date.now());
      setErr(result.error||'Incorrect password.');return;
    }

    this._afterLogin(result, id, pass);
  }

  async _afterLogin(result, id, rawPass){
    /* Save session token from server */
    saveSession(id, result.token);
    this.user=result.user;
    /* Remember the raw password used to log in — needed for first-login password change */
    this._loginRawPass=rawPass;

    /* Show loading overlay while hydrating */
    showLoader('Loading your data…');

    /* Hydrate ALL data from server */
    const data=await API.hydrate();
    if(data&&data.success){
      this.staff=data.staff||{};
      this.records=data.records||[];
      this.leave=data.leave||[];
      this.holidays=data.holidays||[];
      this._cacheH();
    }

    const loT=$('lo-text');if(loT)loT.textContent='Setting up your dashboard…';

    /* DO NOT auto-migrate passwords here.
       The change password form handles migration properly. */

    const role=this.user.role;
    const isDefault=(rawPass==='1234'||rawPass==='admin123');
    const isTempPass=(/^[A-Z0-9]{6}$/.test(rawPass)&&!isDefault);

    if(role==='admin'){
      showView('admin-view');
      setTimeout(()=>{
        this.renderAdmin();this._renderDash();this._renderStaffGrid();this._renderReports();this.renderAdminLeave();this._updateNotifBadges();
        this._populateSupervisorDropdown();this._initEntQR();this.renderAdminHolidays();
        this._checkContractReminders();
        if($('script-url-input')&&API.getGasUrl())$('script-url-input').value=API.getGasUrl();
        hideLoader();
      },100);
      API.updateChips();
      return toast('Welcome, Administrator! 👋');
    }

    if(isManagerRole(role)){
      showView('manager-view');
      setTimeout(()=>{
        if($('m-unit-display'))$('m-unit-display').textContent=this.user.unit;
        this._toggleMgrReports(id);this._setLeaveTabLabel(id);
        if($('mgr-name'))$('mgr-name').textContent=this.user.name;
        const av=$('mgr-av');if(av){av.textContent=ini(this.user.name);av.style.background=this.user.color||avColor(this.user.name);}
        const mav=$('mob-mgr-av');if(mav){mav.textContent=ini(this.user.name);mav.style.background=this.user.color||avColor(this.user.name);}
        const mn=$('mob-mgr-name');if(mn)mn.textContent=this.user.name;
        this._sessCheck();this._initWorkModeListeners();this._stats();this._renderMgrDash();this.renderMgrRecs();this.loadLeave();this._updateNotifBadges();
        if($('m-chpw-name'))$('m-chpw-name').textContent=this.user.name;
        if($('mgr-role'))$('mgr-role').textContent=roleLabel(this.user.role);
        this._checkDefaultPass('mgr');this._renderProfileForm('m-');this._renderMgrLeaveBal();
        if(id===COUNTRY_LEADER_ID){const dn=$('nav-mgr-deleg');if(dn)dn.classList.remove('cl-only-tab');const dm=$('mob-mgr-deleg');if(dm)dm.classList.remove('cl-only-tab');}
        this._applyPrivileges(id);this._checkContractReminders();
        this._startAutoClockOut();this._checkClockInReminder();
        if(isDefault||isTempPass){setTimeout(()=>showPanel('m-chpw','sb-mgr',null),400);if(isTempPass)setTimeout(()=>toast('🔐 You logged in with a temporary password. Please set a new one now.','info'),1500);}
        hideLoader();
      },100);
    } else {
      showView('staff-view');
      setTimeout(()=>{
        $('st-name').textContent=this.user.name;
        const av=$('st-av');if(av){av.textContent=ini(this.user.name);av.style.background=this.user.color||avColor(this.user.name);}
        const mav=$('mob-st-av');if(mav){mav.textContent=ini(this.user.name);mav.style.background=this.user.color||avColor(this.user.name);}
        const mn=$('mob-st-name');if(mn)mn.textContent=this.user.name;
        this._stats();this.renderStaffLogs();this._staffQR();this._sessCheck();this._initWorkModeListeners();this._renderLeaveBal();this.renderStaffLeave();this._initLeaveForm();this._updateNotifBadges();
        this.renderStaffFeed();this.checkBirthdayWish();
        this.renderMyPayslips();
        (this._applyPrivileges?this:APP)._applyPrivileges(id);
        if($('unit-display'))$('unit-display').textContent=this.user.unit;
        this._filterLeaveByGender();this._checkDefaultPass('');this._renderProfileForm('');
        this._startAutoClockOut();this._checkClockInReminder();
        if(isDefault||isTempPass){setTimeout(()=>showPanel('p-chpw','sb-staff',null),400);setTimeout(()=>toast(isTempPass?'🔐 You logged in with a temporary password. Please set a new one now.':'⚠️ Please change your default password.','info'),1500);}
        hideLoader();
      },100);
    }
    API.updateChips();
    toast('Welcome back, '+this.user.name+'! 👋');
  }

  /* ── Forgot Password ── */
  showForgotPass(){
    // Hide login fields, show forgot panel
    ['uni-id','uni-pass'].forEach(id=>{const el=$(id);if(el)el.closest('.lc-field').style.display='none';});
    const err=$('lc-err');if(err)err.style.display='none';
    const btn=document.querySelector('.lc-btn');if(btn)btn.style.display='none';
    const forgotLink=document.querySelector('.lc-forgot');if(forgotLink)forgotLink.style.display='none';
    $('forgot-panel').style.display='block';
    $('forgot-id')?.focus();
  }
  showLoginForm(){
    ['uni-id','uni-pass'].forEach(id=>{const el=$(id);if(el)el.closest('.lc-field').style.display='';});
    const err=$('lc-err');if(err){err.style.display='';err.textContent='';}
    const btn=document.querySelector('.lc-btn');if(btn)btn.style.display='';
    const forgotLink=document.querySelector('.lc-forgot');if(forgotLink)forgotLink.style.display='';
    $('forgot-panel').style.display='none';
    const msg=$('forgot-msg');if(msg)msg.textContent='';
    $('uni-id')?.focus();
  }
  async forgotPassword(){
    const id=$('forgot-id')?.value.trim().toUpperCase();
    const msg=$('forgot-msg');
    if(!id){if(msg)msg.innerHTML='<span style="color:var(--red)">Please enter your Staff ID.</span>';return;}

    // Show loading state
    const btn=$('forgot-panel')?.querySelector('.lc-btn');
    if(btn){btn.classList.add('loading');btn.querySelector('span').textContent='Sending…';}
    if(msg)msg.innerHTML='<span style="color:var(--teal)">⏳ Looking up your account…</span>';

    const result=await API.resetPassword(id);

    if(btn){btn.classList.remove('loading');btn.querySelector('span').textContent='Send Temporary Password';}

    if(!result||!result.success){
      if(msg)msg.innerHTML=`<span style="color:var(--red)">${result?.error||'Something went wrong. Try again.'}</span>`;
      return;
    }

    if(msg)msg.innerHTML=`<span style="color:var(--green)">✓ Temporary password sent to <strong>${result.email}</strong>.<br>Check your inbox (and spam folder), then come back and sign in.</span>`;
    // Clear input and disable button briefly
    if($('forgot-id'))$('forgot-id').value='';
    if(btn){btn.disabled=true;setTimeout(()=>{btn.disabled=false;},10000);}
  }

  /* ── Admin password change ── */
  async changeAdminPass(){
    const old=$('a-chpw-old').value.trim();
    const np=$('a-chpw-new').value.trim();
    const conf=$('a-chpw-confirm').value.trim();
    const msg=$('a-chpw-msg');msg.textContent='';
    if(!old||!np||!conf){msg.innerHTML='<span style="color:var(--red)">Fill all fields.</span>';return;}
    if(np.length<4){msg.innerHTML='<span style="color:var(--red)">Min 4 characters.</span>';return;}
    if(np!==conf){msg.innerHTML='<span style="color:var(--red)">Passwords don\'t match.</span>';return;}
    if(np===old){msg.innerHTML='<span style="color:var(--red)">Must be different.</span>';return;}
    msg.innerHTML='<span style="color:var(--teal)">⏳ Saving…</span>';
    const session=getSession();
    /* Try plain text first, then hashed — server may store either */
    let r=await API.changePassword(ADMIN_ID,old,np,session?.token);
    if(!r||!r.success){
      const oldHashed=await hashPass(ADMIN_ID,old);
      r=await API.changePassword(ADMIN_ID,oldHashed,np,session?.token);
    }
    if(r&&r.success){
      msg.innerHTML='<span style="color:var(--green)">✓ Admin password updated.</span>';
      $('a-chpw-old').value='';$('a-chpw-new').value='';$('a-chpw-confirm').value='';
      toast('Admin password changed.');
    } else {
      msg.innerHTML=`<span style="color:var(--red)">${r?.error||'Failed — check current password.'}</span>`;
    }
  }

  /* ── Logout — invalidate server session ── */
  async logout(){
    if(!confirm('Sign out?'))return;
    const session=getSession();
    if(session) await API.logout(session.id, session.token);
    clearSession();
    this.user=null;
    const loginEl=$('login-view');if(loginEl)loginEl.style.display='';
    showView('login-view');
  }

  _sessCheck(){
    const pfx=isManagerRole(this.user.role)?'m-':'';
    const rec=this.records.find(r=>r.id===this.user.id&&!r.out);
    if(rec){$(pfx+'btn-ci').disabled=true;$(pfx+'btn-co').disabled=false;this._sess(true);}
  }

  /* ── Clock in/out — SERVER FIRST ── */
  _pfx(){return isManagerRole(this.user?.role)?'m-':'';}

  /* ── Work mode change handler — show/hide trip panel ── */
  _initWorkModeListeners(){
    const p=this._pfx();
    const sel=$(p+'work-mode');if(!sel)return;
    sel.addEventListener('change',()=>{
      const tp=$(p+'trip-panel');
      if(tp)tp.style.display=sel.value==='Work Trip'?'block':'none';
    });
  }

  async clockIn(){
    const now=new Date();
    const p=this._pfx();
    const ciBtn=$(p+'btn-ci');
    if(ciBtn)ciBtn.disabled=true; // Prevent double-tap
    const _bail=(msg,type)=>{if(ciBtn)ciBtn.disabled=false;return toast(msg,type||'err');};
    const workMode=$(p+'work-mode')?.value||'Office';

    // Work Trip mode — redirect to trip registration
    if(workMode==='Work Trip'){
      if(ciBtn)ciBtn.disabled=false;
      const tp=$(p+'trip-panel');if(tp)tp.style.display='block';
      return toast('Fill in your trip dates below and register.','info');
    }

    if(isWeekend(now))return _bail('Not allowed on weekends.');
    if(isHoliday(now)){const hName=getHolidayName(now);return _bail(`Today is a public holiday${hName?' — '+hName:''}.`,'info');}
    if(this.records.find(r=>r.id===this.user.id&&!r.out))return _bail('Already clocked in');
    const todayStr=todayISO();
    const alreadyToday=this.records.find(r=>r.id===this.user.id&&((r.date||r.in||'').slice(0,10)===todayStr||(r.in&&new Date(r.in).toISOString().slice(0,10)===todayStr)));
    if(alreadyToday)return _bail('Already clocked in today.');
    // Double-check server for duplicates (handles multi-tab / stale cache)
    const serverCheck=await API._get('attendance','staff_id=eq.'+encodeURIComponent(this.user.id)+'&date=eq.'+encodeURIComponent(fmtD(now.toISOString()))+'&limit=1');
    if(serverCheck&&serverCheck.length)return _bail('Already clocked in today (server confirmed).');
    const onLeave=leaveOnDate(this.leave,this.user.id,todayStr);
    if(onLeave)return _bail(`On approved ${onLeave.type} today.`,'info');

    const unit=(this.user.unit||'').trim();
    const rec={date:fmtD(now.toISOString()),id:this.user.id,name:this.user.name,unit,in:now.toISOString(),out:null,hours:null,status:'Active',work_mode:workMode};

    /* SERVER FIRST */
    const result=await API.saveRecord(rec);
    if(!result||!result.success){if(ciBtn)ciBtn.disabled=false;toast('Server error — try again','err');return;}

    this.records.push(rec);this._cacheR();
    $(p+'btn-ci').disabled=true;$(p+'btn-co').disabled=false;this._sess(true);this._stats();
    const modeLabel=workMode==='Office'?'':'('+workMode+') ';
    toast('Clocked in '+modeLabel+'at '+fmtT(now.toISOString()));
  }

  /* ── Register Work Trip — auto-marks attendance for entire trip duration ── */
  async registerWorkTrip(){
    const p=this._pfx();
    const startDate=$(p+'trip-start')?.value;
    const endDate=$(p+'trip-end')?.value;
    const dest=$(p+'trip-dest')?.value.trim()||'Work Trip';
    if(!startDate||!endDate)return toast('Select trip start and end dates.','err');
    if(new Date(endDate)<new Date(startDate))return toast('End date before start date.','err');

    const unit=(this.user.unit||'').trim();
    const days=[];
    const cur=new Date(startDate);
    const end=new Date(endDate);
    while(cur<=end){days.push(new Date(cur));cur.setDate(cur.getDate()+1);}
    if(!days.length)return toast('No days in range.','err');

    toast(`Registering ${days.length} trip day(s)…`,'info');
    let added=0;
    for(const day of days){
      const dayISO=day.toISOString().slice(0,10);
      const already=this.records.find(r=>r.id===this.user.id&&((r.date||r.in||'').slice(0,10)===dayISO||(r.in&&new Date(r.in).toISOString().slice(0,10)===dayISO)));
      if(already)continue;
      const clockIn=new Date(day);clockIn.setHours(8,0,0,0);
      const clockOut=new Date(day);clockOut.setHours(17,0,0,0);
      const rec={date:fmtD(clockIn.toISOString()),id:this.user.id,name:this.user.name,unit,
        in:clockIn.toISOString(),out:clockOut.toISOString(),hours:'9.00',
        status:'Completed (Work Trip — '+dest+')',work_mode:'Work Trip'};
      const r=await API.saveRecord(rec);
      if(r&&r.success){this.records.push(rec);added++;}
    }
    this._cacheR();this._stats();
    if(isManagerRole(this.user.role))this.renderMgrRecs();else this.renderStaffLogs();
    $(p+'trip-panel').style.display='none';
    $(p+'trip-start').value='';$(p+'trip-end').value='';$(p+'trip-dest').value='';
    $(p+'work-mode').value='Office';
    toast(`✈️ Work trip registered! ${added} day(s) auto-marked as present.`);
  }
  clockOut(){
    const rec=this.records.find(r=>r.id===this.user.id&&!r.out);if(!rec)return;
    const hrs=(new Date()-new Date(rec.in))/3600000;
    const p=this._pfx();
    if(hrs<this.HOURS)$(p+'early-panel').style.display='block';else this._fin(rec,hrs,'Completed');
  }
  toggleOther(sel){const p=this._pfx();$(p+'other-reason').style.display=sel.value==='Other'?'block':'none';}
  confirmExit(){
    const p=this._pfx();
    const reason=$(p+'exit-reason').value;if(!reason)return toast('Select a reason','err');
    const rec=this.records.find(r=>r.id===this.user.id&&!r.out);if(!rec)return;
    const hrs=(new Date()-new Date(rec.in))/3600000;
    this._fin(rec,hrs,'Early Exit ('+($(p+'exit-reason').value==='Other'?($(p+'other-reason').value||'Other'):reason)+')');
    $(p+'early-panel').style.display='none';$(p+'exit-reason').value='';$(p+'other-reason').style.display='none';
  }
  async _fin(rec,hrs,status){
    const p=this._pfx();
    rec.out=new Date().toISOString();rec.hours=fx(hrs);rec.status=status;
    /* SERVER FIRST */
    await API.updateRecord(rec);
    this._cacheR();$(p+'btn-co').disabled=true;
    this._sess(false);this._stats();
    if(isManagerRole(this.user.role))this.renderMgrReport();else this.renderStaffLogs();
    toast(status.includes('Early')?'Early exit recorded.':'Shift complete — '+fx(hrs)+' hrs');
  }
  _sess(on){
    const p=this._pfx();
    const badge=$(p+'sess-badge'),txt=$(p+'sess-txt');
    if(badge)badge.className='sess-badge '+(on?'sess-on':'sess-off');
    if(txt)txt.textContent=on?'At Post':'Signed Out';
  }
  _stats(){
    const p=this._pfx();
    const n=new Date(),mm=n.getMonth(),yy=n.getFullYear();
    const mo=this.records.filter(r=>r.id===this.user.id&&r.out).filter(r=>{const d=new Date(r.in);return d.getMonth()===mm&&d.getFullYear()===yy;});
    const hrs=mo.reduce((a,r)=>a+parseFloat(r.hours||0),0);
    if($(p+'s-days'))$(p+'s-days').textContent=mo.length;
    if($(p+'s-avg'))$(p+'s-avg').textContent=mo.length?fx(hrs/mo.length):'0.00';
    if($(p+'s-early'))$(p+'s-early').textContent=mo.filter(r=>r.status.includes('Early')).length;
    if($(p+'s-hrs'))$(p+'s-hrs').textContent=fx(mo.reduce((a,r)=>a+parseFloat(r.hours||0),0));
  }

  /* ── Staff logs ── */
  _wmBadge(r){return r.work_mode&&r.work_mode!=='Office'?`<span style="font-size:.66rem;display:inline-block;padding:1px 5px;border-radius:4px;background:rgba(61,191,184,.15);color:var(--teal);margin-left:3px">${r.work_mode}</span>`:'';}
  renderStaffLogs(){
    const mv=$('st-mth')?.value;
    let recs=this.records.filter(r=>r.id===this.user.id);
    if(mv){const[y,m]=mv.split('-').map(Number);recs=recs.filter(r=>{const d=new Date(r.in);return d.getFullYear()===y&&d.getMonth()===m-1;});}
    if(this._stFilter.status)recs=recs.filter(r=>r.status&&r.status.includes(this._stFilter.status));
    recs=this._applySort('st',recs);
    const cnt=$('st-count');if(cnt)cnt.textContent=recs.length;
    this._updateSortHeaders('st-table',this._sort.st);
    const body=$('st-logs');
    if(!recs.length){body.innerHTML='<tr><td colspan="6"><div class="empty"><div class="empty-ico">📭</div>No records found</div></td></tr>';return;}
    body.innerHTML=recs.map(r=>`<tr><td>${fmtD(r.date||r.in)}</td><td>${r.unit}${this._wmBadge(r)}</td><td>${fmtT(r.in)}</td><td>${r.out?fmtT(r.out):'<span style="color:var(--teal)">Active</span>'}</td><td>${r.hours||'--'}</td><td>${this._bdg(r.status)}</td></tr>`).join('');
  }
  setStFilter(key,val,el){
    this._stFilter[key]=val;
    el.closest('.filter-chips').querySelectorAll('.chip').forEach(c=>c.classList.remove('active'));
    el.classList.add('active');
    this.renderStaffLogs();
  }

  /* ── Manager My Logs (personal attendance) ── */
  renderMgrMyLogs(){
    const mv=$('mgr-my-mth')?.value;
    let recs=this.records.filter(r=>r.id===this.user.id);
    if(mv){const[y,m]=mv.split('-').map(Number);recs=recs.filter(r=>{const d=new Date(r.date||r.in);return d.getFullYear()===y&&d.getMonth()===m-1;});}
    recs.sort((a,b)=>new Date(b.date||b.in)-new Date(a.date||a.in));
    const cnt=$('mgr-my-count');if(cnt)cnt.textContent=recs.length;
    const body=$('mgr-my-logs-body');if(!body)return;
    if(!recs.length){body.innerHTML='<tr><td colspan="6"><div class="empty"><div class="empty-ico">📭</div>No records found</div></td></tr>';return;}
    body.innerHTML=recs.map(r=>`<tr><td>${fmtD(r.date||r.in)}</td><td>${r.unit}${this._wmBadge(r)}</td><td>${fmtT(r.in)}</td><td>${r.out?fmtT(r.out):'<span style="color:var(--teal)">Active</span>'}</td><td>${r.hours||'--'}</td><td>${this._bdg(r.status)}</td></tr>`).join('');
  }

  /* ── Leave balances ── */
  _leaveDaysUsed(staffId,type){
    const yr=new Date().getFullYear();
    return this.leave.filter(l=>l.staffId===staffId&&l.type===type&&l.status==='Approved'&&new Date(l.startDate).getFullYear()===yr).reduce((a,l)=>a+parseInt(l.days||0),0);
  }
  _renderLeaveBal(){
    const gender=this.staff[this.user.id]?.gender||'';
    let types=['Annual Leave','Sick Leave','Paternity Leave','Compassionate Leave','Compensatory Leave'];
    if(gender==='female')types=['Annual Leave','Sick Leave','Maternity Leave','Compassionate Leave','Compensatory Leave'];
    const icons={'Annual Leave':'🌴','Sick Leave':'🏥','Maternity Leave':'👶','Paternity Leave':'👨‍👶','Compassionate Leave':'🕊','Compensatory Leave':'⏰'};
    $('st-leave-bal').innerHTML=`<h4>Leave Balances (${new Date().getFullYear()})</h4>`+
      types.map(t=>{
        const limit=LEAVE_LIMITS[t];const used=this._leaveDaysUsed(this.user.id,t);
        if(limit===null){const sub=t==='Compensatory Leave'?'As certified by Country Leader':'As certified by medical professional';return`<div class="bal-row"><div class="bal-icon">${icons[t]||'📋'}</div><div class="bal-info"><div class="bal-lbl">${t}</div><div style="font-size:.72rem;color:var(--text2)">${sub}</div></div><div class="bal-num">${used} days used</div></div>`;}
        const rem=Math.max(0,limit-used);const pct=Math.round((used/limit)*100);
        return`<div class="bal-row"><div class="bal-icon">${icons[t]||'📋'}</div><div class="bal-info"><div class="bal-lbl">${t}</div><div class="bal-trk"><div class="bal-fill" style="width:${pct}%;background:${pct>85?'var(--red)':pct>60?'var(--gold)':'var(--green)'}"></div></div></div><div class="bal-num">${rem}/${limit} left</div></div>`;
      }).join('');
  }
  _setMobTab(navId,idx){const nav=$(navId);if(!nav)return;nav.querySelectorAll('.mob-tab').forEach((t,i)=>t.classList.toggle('active',i===idx));}

  _initLeaveForm(){
    const supSel=$('lv-supervisor-sel'),finalSel=$('lv-final-sel');
    if(!supSel||!finalSel)return;
    const uid=this.user?.id||'';
    const isDirectToCL=DIRECT_TO_CL.includes(uid);
    const routingBlock=$('lv-routing-block'),directBlock=$('lv-direct-block');
    if(isDirectToCL){
      if(routingBlock)routingBlock.style.display='none';
      if(directBlock)directBlock.style.display='block';
    } else {
      if(routingBlock)routingBlock.style.display='block';
      if(directBlock)directBlock.style.display='none';
      const managers=Object.entries(this.staff)
        .filter(([id,s])=>SUPERVISOR_ROLES.includes(s.role)&&id!==uid&&id!==COUNTRY_LEADER_ID)
        .sort((a,b)=>a[1].name.localeCompare(b[1].name));
      supSel.innerHTML='<option value="">— Select supervisor —</option>'+
        managers.map(([id,s])=>`<option value="${id}">${s.name} (${s.unit})</option>`).join('');
      const agathaName=this.staff[COUNTRY_LEADER_ID]?.name||'Agatha Quayson';
      finalSel.innerHTML=`<option value="${COUNTRY_LEADER_ID}">${agathaName} — Country Leader</option>`;
      finalSel.value=COUNTRY_LEADER_ID;finalSel.disabled=true;
    }
  }
  _onSupChange(){
    const supId=$('lv-supervisor-sel')?.value;
    const info=$('lv-routing-info'),path=$('lv-routing-path');
    if(!info||!path)return;
    if(supId){
      const supName=this.staff[supId]?.name||supId;
      const finalName=this.staff[COUNTRY_LEADER_ID]?.name||'Agatha Quayson';
      path.textContent=`${supName} → ${finalName} (Country Leader)`;
      info.style.display='block';
    } else {info.style.display='none';}
  }

  /* ── Notification badges ── */
  _updateNotifBadges(){
    if(!this.user)return;
    const role=this.user.role;
    const setBadge=(sidebarId,mobileId,count)=>{
      const n=count>0?String(count>99?'99+':count):'';
      const show=count>0;
      const sb=$(sidebarId);if(sb){sb.textContent=n;sb.classList.toggle('show',show);}
      const mb=$(mobileId);if(mb){mb.textContent=n;mb.classList.toggle('show',show);}
    };
    if(isManagerRole(role)){
      const uid=this.user.id;
      const isFinalApprover=uid===COUNTRY_LEADER_ID||this._isActiveDelegate(uid);
      const pending=isFinalApprover
        ? this.leave.filter(l=>(l.finalApproverId===COUNTRY_LEADER_ID||l.finalApproverId===uid)&&(l.finalApproverStatus==='Pending'||l.hrStatus==='Pending')&&(l.supervisorStatus==='Approved'||l.supervisorStatus==='N/A')).length
        : this.leave.filter(l=>l.supervisorId===uid&&l.supervisorStatus==='Pending').length;
      setBadge('badge-mgr-leave','mob-badge-mgr-leave',pending);
    }
    if(role==='staff'){
      const seen=this._getSeenLeaveIds();
      const updated=this.leave.filter(l=>l.staffId===this.user.id&&(l.status==='Approved'||l.status==='Rejected')&&!seen[l.id]).length;
      setBadge('badge-staff-leave','mob-badge-staff-leave',updated);
    }
    if(role==='admin'){
      const pending=this.leave.filter(l=>l.status==='Pending').length;
      setBadge('badge-admin-leave','mob-badge-admin-leave',pending);
    }
  }
  _getSeenLeaveIds(){
    try{return JSON.parse(localStorage.getItem('thp_seen_leave')||'{}');}catch(e){return{};}
  }
  _markLeaveDecisionsSeen(){
    if(this.user?.role!=='staff')return;
    const seen=this._getSeenLeaveIds();
    let changed=false;
    this.leave.forEach(l=>{
      if(l.staffId===this.user.id&&(l.status==='Approved'||l.status==='Rejected')&&!seen[l.id]){
        seen[l.id]=true;changed=true;
      }
    });
    if(changed){
      localStorage.setItem('thp_seen_leave',JSON.stringify(seen));
      this._updateNotifBadges();
    }
  }

  _renderMgrLeaveBal(){
    const gender=this.staff[this.user.id]?.gender||'';
    let types=['Annual Leave','Sick Leave','Paternity Leave','Compassionate Leave','Compensatory Leave'];
    if(gender==='female')types=['Annual Leave','Sick Leave','Maternity Leave','Compassionate Leave','Compensatory Leave'];
    const icons={'Annual Leave':'🌴','Sick Leave':'🏥','Maternity Leave':'👶','Paternity Leave':'👨‍👶','Compassionate Leave':'🕊','Compensatory Leave':'⏰'};
    const el=$('mgr-leave-bal');if(!el)return;
    el.innerHTML=`<h4>Leave Balances (${new Date().getFullYear()})</h4>`+
      types.map(t=>{
        const limit=LEAVE_LIMITS[t];const used=this._leaveDaysUsed(this.user.id,t);
        if(limit===null){const sub=t==='Compensatory Leave'?'As certified by Country Leader':'As certified by medical professional';return`<div class="bal-row"><div class="bal-icon">${icons[t]||'📋'}</div><div class="bal-info"><div class="bal-lbl">${t}</div><div style="font-size:.72rem;color:var(--text2)">${sub}</div></div><div class="bal-num">${used} used</div></div>`;}
        const rem=Math.max(0,limit-used);const pct=Math.round((used/limit)*100);
        return`<div class="bal-row"><div class="bal-icon">${icons[t]||'📋'}</div><div class="bal-info"><div class="bal-lbl">${t}</div><div class="bal-trk"><div class="bal-fill" style="width:${pct}%;background:${pct>85?'var(--red)':pct>60?'var(--gold)':'var(--green)'}"></div></div></div><div class="bal-num">${rem}/${limit} left</div></div>`;
      }).join('');
    this._filterLeaveByGender();
  }

  renderMgrMyLeave(){
    const mine=this.leave.filter(l=>l.staffId===this.user.id);
    const body=$('mgr-myleave-body');if(!body)return;
    body.innerHTML=mine.length?mine.slice().reverse().map(l=>{
      const fa=l.finalApproverStatus||l.hrStatus||'Pending';
      const faName=this.staff[l.finalApproverId]?.name||'Final Approver';
      const faLabel=fa==='Approved'?'✓ Approved':fa==='Rejected'?'✗ Rejected':fa==='Waiting'?'⏳ Awaiting supervisor':'⏳ Pending';
      const faBdg=`<span class="stage-badge ${fa==='Approved'?'stage-ok':fa==='Rejected'?'stage-rej':'stage-pend'}"><div style="font-size:.68rem;opacity:.7">${faName}</div>${faLabel}</span>`;
      const note=l.finalApproverNote||l.supervisorNote||'—';
      return`<tr><td>${l.type}</td><td>${fmtISO(l.startDate)}</td><td>${fmtISO(l.endDate)}</td><td>${l.days}</td><td>${faBdg}</td><td style="font-size:.74rem">${note}</td></tr>`;
    }).join(''):'<tr><td colspan="6"><div class="empty"><div class="empty-ico">🌴</div>No leave requests</div></td></tr>';
  }

  _toggleMgrReports(uid){
    const REPORT_MANAGERS=['THPG/05/2025','THPG/03/2008'];
    const show=REPORT_MANAGERS.includes(uid);
    const sidebar=$('nav-mgr-report'),mobile=$('mob-mgr-report');
    if(sidebar)sidebar.style.display=show?'':'none';
    if(mobile)mobile.style.display=show?'':'none';
  }
  _setLeaveTabLabel(uid){
    const isAgatha=uid===COUNTRY_LEADER_ID;
    const sideText=$('nav-mgr-leave-text'),mobText=$('mob-mgr-leave-text');
    const title=$('mgr-leave-title'),subtitle=$('mgr-leave-subtitle');
    if(sideText)sideText.textContent=isAgatha?'Leave Approval':'Leave Review';
    if(mobText)mobText.textContent=isAgatha?'Approval':'Review';
    if(title)title.textContent=isAgatha?'Leave Approval':'Leave Review';
    if(subtitle)subtitle.textContent=isAgatha?'Your decision is final':'Forward to Country Leader for final sign-off';
    const brand=$('mgr-brand-title'),mobRole=$('mob-mgr-role');
    const rl=this.staff[uid]?.role||'manager';
    if(brand)brand.textContent=roleLabel(rl);
    if(mobRole)mobRole.textContent=roleLabel(rl)+' · THP-Ghana';
  }

  _renderMgrDash(){
    const td=today(),teamStaff=Object.entries(this.staff);
    const todayRecs=this.records.filter(r=>sameDay(r.date||r.in));
    const active=this.records.filter(r=>!r.out).length;
    const todayISOStr=todayISO();
    const onLeaveToday=teamStaff.filter(([id])=>{
      const alreadyClockedIn=todayRecs.some(r=>r.id===id);
      return !alreadyClockedIn&&leaveOnDate(this.leave,id,todayISOStr);
    });
    $('mgr-stats').innerHTML=`
      <div class="stat stat-teamsize"><div class="stat-lbl">Team Size</div><div class="stat-val">${teamStaff.length}</div></div>
      <div class="stat"><div class="stat-lbl">Present Today</div><div class="stat-val g">${todayRecs.length}</div></div>
      <div class="stat"><div class="stat-lbl">Active Now</div><div class="stat-val a">${active}</div></div>
      <div class="stat"><div class="stat-lbl">On Leave</div><div class="stat-val" style="color:var(--gold)">${onLeaveToday.length}</div></div>
      <div class="stat"><div class="stat-lbl">Pending Leave</div><div class="stat-val p">${this.leave.filter(l=>l.status==='Pending').length}</div></div>`;
    const body=$('mgr-today');
    const tr=todayRecs.slice().reverse();
    const leaveRows=onLeaveToday.map(([id,s])=>{
      const lv=leaveOnDate(this.leave,id,todayISOStr);
      return`<tr style="opacity:.8"><td><strong>${s.name}</strong></td><td>${s.unit}</td><td colspan="3" style="color:var(--text2);font-style:italic">On leave</td><td><span class="badge" style="background:rgba(99,102,241,.15);color:#4338ca">🌴 ${lv.type}</span></td></tr>`;
    }).join('');
    if(!tr.length&&!leaveRows){body.innerHTML='<tr><td colspan="6"><div class="empty"><div class="empty-ico">📭</div>No attendance today</div></td></tr>';return;}
    body.innerHTML=tr.map(r=>`<tr><td><strong>${r.name}</strong></td><td>${r.unit}</td><td>${fmtT(r.in)}</td><td>${r.out?fmtT(r.out):'<span style="color:var(--teal)">Active</span>'}</td><td>${r.hours||'--'}</td><td>${this._bdg(r.status)}</td></tr>`).join('')+leaveRows;
  }
  renderMgrRecs(){
    const mv=$('mgr-mth')?.value,srch=($('mgr-srch')?.value||'').toLowerCase();
    let recs=this.records.slice();
    if(mv){const[y,m]=mv.split('-').map(Number);recs=recs.filter(r=>{const d=new Date(r.in);return d.getFullYear()===y&&d.getMonth()===m-1;});}
    if(srch)recs=recs.filter(r=>r.name.toLowerCase().includes(srch));
    if(this._mgrFilter.status)recs=recs.filter(r=>r.status&&r.status.includes(this._mgrFilter.status));
    recs=this._applySort('mgr',recs);
    const cnt=$('mgr-count');if(cnt)cnt.textContent=recs.length;
    this._updateSortHeaders('mgr-table',this._sort.mgr);
    const body=$('mgr-recs-body');
    if(!recs.length){body.innerHTML='<tr><td colspan="6"><div class="empty"><div class="empty-ico">📭</div>No records</div></td></tr>';return;}
    body.innerHTML=recs.map(r=>`<tr><td>${fmtD(r.date||r.in)}</td><td><strong>${r.name}</strong></td><td>${fmtT(r.in)}</td><td>${r.out?fmtT(r.out):'<span style="color:var(--teal)">Active</span>'}</td><td>${r.hours||'--'}</td><td>${this._bdg(r.status)}</td></tr>`).join('');
  }
  setMgrFilter(key,val,el){this._mgrFilter[key]=val;el.closest('.filter-chips').querySelectorAll('.chip').forEach(c=>c.classList.remove('active'));el.classList.add('active');this.renderMgrRecs();}
  clearMgrFilters(){this._mgrFilter={status:''};if($('mgr-srch'))$('mgr-srch').value='';if($('mgr-mth'))$('mgr-mth').value='';document.querySelectorAll('#m-recs .chip').forEach(c=>c.classList.remove('active'));document.querySelector('#m-recs .chip-all')?.classList.add('active');this.renderMgrRecs();}

  /* ── Leave type selection ── */
  _lvPfx(){return isManagerRole(this.user?.role)?'mlv-':'lv-';}
  selectLeave(el){
    document.querySelectorAll('.ltype-card').forEach(c=>c.classList.remove('sel'));
    el.classList.add('sel');
    const type=el.dataset.type;
    const p=this._lvPfx();
    const sickUpload=$(p+'sick-upload');
    if(sickUpload)sickUpload.style.display=type==='Sick Leave'?'block':'none';
    const compDates=$(p+'comp-dates');
    if(compDates)compDates.style.display=type==='Compensatory Leave'?'block':'none';
    this.calcLeaveDays();
  }
  calcLeaveDays(){
    const p=this._lvPfx();
    const s=$(p+'start')?.value,e=$(p+'end')?.value;
    const preview=$(p+'days-preview');
    if(!s||!e||!preview)return;
    const days=workingDaysBetween(s,e);
    $(p+'days-count').textContent=days;
    preview.style.display=days>0?'block':'none';
  }
  handleSickFile(inp){
    const p=this._lvPfx();
    const file=inp.files[0];if(!file)return;
    if(file.size>5*1024*1024){toast('File too large — max 5MB','err');inp.value='';return;}
    const el=$(p+'file-name');if(el)el.textContent='📎 '+file.name;
    toast('Document attached: '+file.name,'info');
  }
  _filterLeaveByGender(){
    const gender=this.staff[this.user.id]?.gender||'';
    document.querySelectorAll('[data-type="Maternity Leave"]').forEach(el=>el.classList.toggle('ltype-hidden',gender==='male'));
    document.querySelectorAll('[data-type="Paternity Leave"]').forEach(el=>el.classList.toggle('ltype-hidden',gender==='female'));
  }

  /* ── Apply leave — SERVER FIRST ── */
  async applyLeave(){
    const p=this._lvPfx();
    const selCard=document.querySelector('.ltype-card.sel');
    const type=selCard?.dataset?.type;
    const start=$(p+'start')?.value,end=$(p+'end')?.value;
    const reason=$(p+'reason')?.value.trim();
    const errEl=$(p+'err');
    const setErr=m=>{if(errEl)errEl.textContent=m;};
    setErr('');
    if(!type)return setErr('Select a leave type.');
    if(!start||!end)return setErr('Select start and end dates.');
    if(new Date(end)<new Date(start))return setErr('End date before start date.');
    const gender=this.staff[this.user.id]?.gender||'';
    if(type==='Maternity Leave'&&gender!=='female')return setErr('Maternity: female staff only.');
    if(type==='Paternity Leave'&&gender!=='male')return setErr('Paternity: male staff only.');
    if(type==='Sick Leave'){const fi=$(p+'sick-file');if(fi&&!fi.files.length)return setErr('Upload a medical certificate.');}
    if(type==='Compensatory Leave'){const cr=$(p+'comp-ref')?.value.trim();if(!cr)return setErr('Specify the dates you worked (weekends/holidays).');}
    const days=workingDaysBetween(start,end);
    if(days===0)return setErr('Dates fall on weekends/holidays.');
    const limit=LEAVE_LIMITS[type];
    if(limit!==null&&limit!==undefined){
      const used=this.leave.filter(l=>l.staffId===this.user.id&&l.type===type&&l.status!=='Rejected').reduce((a,l)=>a+(parseInt(l.days)||0),0);
      if(used+days>limit)return setErr(`${limit-used} days left for ${type}.`);
    }
    const overlap=this.leave.find(l=>l.staffId===this.user.id&&l.type===type&&l.status!=='Rejected'&&new Date(l.startDate)<=new Date(end)&&new Date(l.endDate)>=new Date(start));
    if(overlap)return setErr('Overlapping request exists.');

    const handoverNote=$(p+'handover')?.value.trim()||'';
    const compRef=type==='Compensatory Leave'?($(p+'comp-ref')?.value.trim()||''):'';

    const uid=this.user.id;
    const isDirectToCL=DIRECT_TO_CL.includes(uid);
    let supervisorId,supervisorStatus,finalApproverId;
    if(uid===COUNTRY_LEADER_ID){supervisorId=COUNTRY_LEADER_ID;supervisorStatus='N/A';finalApproverId=COUNTRY_LEADER_ID;}
    else if(isDirectToCL){supervisorId=COUNTRY_LEADER_ID;supervisorStatus='N/A';finalApproverId=COUNTRY_LEADER_ID;}
    else{
      const pickedSup=$('lv-supervisor-sel')?.value||'';
      if(!pickedSup)return setErr('Select a supervisor.');
      supervisorId=pickedSup;supervisorStatus='Pending';finalApproverId=COUNTRY_LEADER_ID;
    }

    const id='LV'+Date.now();
    const lv={id,staffId:uid,name:this.user.name,unit:this.user.unit,type,startDate:start,endDate:end,days,reason,
      staffEmail:this.staff[uid]?.email||'',
      supervisorId,supervisorStatus,supervisorNote:'',
      finalApproverId,finalApproverStatus:uid===COUNTRY_LEADER_ID?'Approved':supervisorStatus==='N/A'?'Pending':'Waiting',finalApproverNote:'',
      status:uid===COUNTRY_LEADER_ID?'Approved':'Pending',hrStatus:uid===COUNTRY_LEADER_ID?'Approved':'Pending',hrNote:'',
      sickNote:type==='Sick Leave'?($(p+'sick-file')?.files[0]?.name||''):'',
      handoverNote,compRef,
      _supervisorEmail:this.staff[supervisorId]?.email||'',
      _finalApproverEmail:this.staff[finalApproverId]?.email||''
    };

    /* SERVER FIRST */
    const result=await API.applyLeave(lv);
    if(!result||!result.success){toast('Server error — try again','err');return;}

    /* Upload sick note to Supabase Storage if present */
    if(type==='Sick Leave'){
      const fileInput=$(p+'sick-file');
      if(fileInput&&fileInput.files.length){
        const file=fileInput.files[0];
        try{
          toast('Uploading medical document…','info');
          const base64=await this._fileToBase64(file);
          const uploadResult=await API.uploadSickNote(result.leaveId||id,file.name,base64,file.type);
          if(uploadResult&&uploadResult.success){
            lv.sickNote=file.name+' | '+uploadResult.fileUrl;
            lv.sickNoteUrl=uploadResult.fileUrl;
            lv.sickNoteDownload=uploadResult.downloadUrl;
            toast('Medical document uploaded ✓');
          } else {
            toast('Document saved locally but upload failed','err');
          }
        }catch(e){console.warn('Sick note upload:',e);toast('Document upload error — leave still submitted','err');}
      }
    }

    this.leave.push(lv);this._cacheL();this._updateNotifBadges();
    if(isManagerRole(this.user.role))this.renderMgrMyLeave();else this.renderStaffLeave();
    // Clear form
    $(p+'start').value='';$(p+'end').value='';$(p+'reason').value='';
    if($(p+'handover'))$(p+'handover').value='';
    if($(p+'comp-ref'))$(p+'comp-ref').value='';
    if($(p+'comp-dates'))$(p+'comp-dates').style.display='none';
    if($('lv-supervisor-sel'))$('lv-supervisor-sel').value='';
    const preview=$(p+'days-preview');if(preview)preview.style.display='none';
    setErr('');
    toast(uid===COUNTRY_LEADER_ID?'Leave auto-approved.':isDirectToCL?'Submitted — awaiting Country Leader.':'Submitted — awaiting supervisor.','info');
  }

  renderStaffLeave(){
    const body=$('st-leave-body');if(!body)return;
    const mine=this.leave.filter(l=>l.staffId===this.user.id);
    if(!mine.length){body.innerHTML='<tr><td colspan="8"><div class="empty"><div class="empty-ico">🏖</div>No leave requests</div></td></tr>';return;}
    const _bdg=(status,na)=>{
      if(na&&status==='N/A')return '<span class="stage-badge" style="background:rgba(148,163,184,.15);color:var(--text3)">— Skipped</span>';
      if(status==='Approved')return '<span class="stage-badge stage-ok">✓ Approved</span>';
      if(status==='Rejected')return '<span class="stage-badge stage-rej">✗ Rejected</span>';
      if(status==='Waiting')return '<span class="stage-badge stage-pend">⏳ Waiting</span>';
      return '<span class="stage-badge stage-pend">⏳ Pending</span>';
    };
    body.innerHTML=mine.slice().reverse().map(l=>{
      const supName=this.staff[l.supervisorId]?.name||l.supervisorId||'—';
      const finName=this.staff[l.finalApproverId]?.name||l.finalApproverId||'—';
      const note=l.finalApproverNote||l.supervisorNote||'—';
      const editBtn=l.status==='Pending'?`<br><button class="bsm" style="margin-top:4px;font-size:.68rem;background:var(--surf2);border:1px solid var(--border);color:var(--text2)" onclick="APP.openLeaveEditModal('${l.id}')">✏ Edit dates</button>`:'';
      return`<tr><td>${l.type}</td><td>${fmtISO(l.startDate)}</td><td>${fmtISO(l.endDate)}</td><td>${l.days}</td><td><div style="font-size:.7rem;color:var(--text3)">${supName}</div>${_bdg(l.supervisorStatus,true)}</td><td><div style="font-size:.7rem;color:var(--text3)">${finName}</div>${_bdg(l.finalApproverStatus||l.hrStatus)}</td><td>${_bdg(l.status)}${editBtn}</td><td style="font-size:.74rem;color:var(--text2)">${note}</td></tr>`;
    }).join('');
  }

  /* ── Staff: edit dates on a pending leave request ── */
  openLeaveEditModal(id){
    const l=this.leave.find(x=>x.id===id&&x.staffId===this.user.id);
    if(!l)return;
    if(l.status!=='Pending')return toast('Only pending requests can be edited','err');
    $('le-id').value=id;
    $('le-type').textContent=l.type;
    $('le-start').value=String(l.startDate).slice(0,10);
    $('le-end').value=String(l.endDate).slice(0,10);
    this.updateLeaveEditDays();
    $('le-msg').textContent='';
    $('leave-edit-modal').classList.add('open');
  }
  updateLeaveEditDays(){
    const s=$('le-start')?.value,e=$('le-end')?.value;
    const el=$('le-days');if(!el)return;
    if(s&&e&&e>=s){el.textContent=workingDaysBetween(s,e)+' working day(s)';}
    else el.textContent='—';
  }
  async saveLeaveEdit(){
    const id=$('le-id').value;
    const l=this.leave.find(x=>x.id===id);if(!l)return;
    const s=$('le-start').value,e=$('le-end').value,msg=$('le-msg');
    if(!s||!e)return msg.innerHTML='<span style="color:var(--red)">Both dates are required.</span>';
    if(e<s)return msg.innerHTML='<span style="color:var(--red)">End date is before start date.</span>';
    const days=workingDaysBetween(s,e);
    if(days<1)return msg.innerHTML='<span style="color:var(--red)">Selected range has no working days.</span>';
    // Changed dates need fresh review — reset approvals to their initial state
    const supNA=l.supervisorStatus==='N/A';
    const upd={start_date:s,end_date:e,days,
      supervisor_status:supNA?'N/A':'Pending',supervisor_note:'',
      final_approver_status:supNA?'Pending':'Waiting',final_approver_note:'',
      overall_status:'Pending',updated_at:new Date().toISOString()};
    msg.innerHTML='<span style="color:var(--teal)">⏳ Saving…</span>';
    const r=await API._update('leave_requests','id=eq.'+encodeURIComponent(id),upd);
    if(r===null)return msg.innerHTML='<span style="color:var(--red)">Save failed — '+(API.lastError||'reason unknown')+'</span>';
    Object.assign(l,{startDate:s,endDate:e,days,
      supervisorStatus:upd.supervisor_status,supervisorNote:'',
      finalApproverStatus:upd.final_approver_status,finalApproverNote:'',
      status:'Pending',hrStatus:upd.final_approver_status,hrNote:''});
    this._cacheL();
    closeModal('leave-edit-modal');
    this.renderStaffLeave();this._renderLeaveBal();
    toast('Leave dates updated ✓ — request re-submitted for review');
  }

  /* ── Load leave from server ── */
  async loadLeave(){
    this.renderMgrLeave();this.renderAdminLeave();
    try{
      const rows=await API._get('leave_requests','order=applied_at.desc&limit=2000');
      if(rows){
        this.leave=rows.map(r=>({id:r.id,staffId:r.staff_id,name:r.name,unit:(r.unit||'').trim(),type:r.type,
          startDate:r.start_date,endDate:r.end_date,days:r.days,reason:r.reason,sickNote:r.sick_note,
          staffEmail:r.staff_email||'',supervisorId:r.supervisor_id||'',supervisorStatus:r.supervisor_status||'Pending',
          supervisorNote:r.supervisor_note||'',finalApproverId:r.final_approver_id||'',
          finalApproverStatus:r.final_approver_status||'Pending',finalApproverNote:r.final_approver_note||'',
          status:r.overall_status||'Pending',hrStatus:r.final_approver_status||r.overall_status||'Pending',
          hrNote:r.final_approver_note||'',appliedAt:r.applied_at||'',updatedAt:r.updated_at||'',
          handoverNote:r.handover_note||'',compRef:r.comp_ref||''}));
        this._cacheL();
        this.renderMgrLeave();this.renderAdminLeave();
        if(this.renderStaffLeave)this.renderStaffLeave();
        this._updateNotifBadges();
      }
    }catch(e){console.warn('loadLeave:',e);}
  }

  renderMgrLeave(){
    const body=$('mgr-leave-body');if(!body)return;
    const uid=this.user.id;
    const isFinalApprover=uid===COUNTRY_LEADER_ID||this._isActiveDelegate(uid);
    const items=isFinalApprover
      ? this.leave.filter(l=>(l.finalApproverId===COUNTRY_LEADER_ID||l.finalApproverId===uid)&&(l.finalApproverStatus==='Pending'||l.hrStatus==='Pending')&&(l.supervisorStatus==='Approved'||l.supervisorStatus==='N/A'))
      : this.leave.filter(l=>l.supervisorId===uid&&l.supervisorStatus==='Pending');
    if(!items.length){body.innerHTML='<tr><td colspan="7"><div class="empty"><div class="empty-ico">🏖</div>No pending requests</div></td></tr>';return;}
    body.innerHTML=items.slice().reverse().map(l=>`<tr>
      <td><strong>${l.name}</strong><div style="font-size:.68rem;color:var(--text2)">${l.unit}</div></td>
      <td>${l.type}</td>
      <td style="font-size:.76rem">${fmtISO(l.startDate)} → ${fmtISO(l.endDate)}</td>
      <td>${l.days}</td>
      <td style="font-size:.75rem;color:var(--text2)">${l.reason||'—'}</td>
      <td>${l.sickNote?this._renderSickNoteLink(l.sickNote):'—'}</td>
      <td><button class="bsm bsm-navy" onclick="APP.openLeaveModal('${l.id}')">👁 Review</button></td>
    </tr>`).join('');
  }
  renderAdminLeave(){
    const body=$('ad-leave-body');if(!body)return;
    const f=($('ad-leave-filter')?.value)||'';
    // Deduplicate leave entries by ID
    const seen=new Set();
    const unique=this.leave.filter(l=>{if(seen.has(l.id))return false;seen.add(l.id);return true;});
    const items=f?unique.filter(l=>l.status===f):unique;
    if(!items.length){body.innerHTML='<tr><td colspan="9"><div class="empty"><div class="empty-ico">🏖</div>No leave requests</div></td></tr>';return;}
    body.innerHTML=items.slice().reverse().map(l=>`<tr>
      <td><strong>${l.name}</strong></td><td style="color:var(--text2);font-size:.76rem">${l.unit}</td><td>${l.type}</td>
      <td style="font-size:.76rem">${fmtISO(l.startDate)} → ${fmtISO(l.endDate)}</td><td>${l.days}</td>
      <td style="font-size:.76rem;color:var(--text2)">${l.reason||'—'}</td>
      <td><span class="stage-badge ${l.supervisorStatus==='Approved'?'stage-ok':l.supervisorStatus==='Rejected'?'stage-rej':'stage-pend'}">${l.supervisorStatus==='N/A'?'Skipped':l.supervisorStatus}</span></td>
      <td><span class="stage-badge ${(l.finalApproverStatus||l.hrStatus)==='Approved'?'stage-ok':(l.finalApproverStatus||l.hrStatus)==='Rejected'?'stage-rej':'stage-pend'}">${l.finalApproverStatus||l.hrStatus||'Pending'}</span></td>
      <td><button class="bsm bsm-navy" onclick="APP.openLeaveModal('${l.id}')">Review</button></td>
    </tr>`).join('');
  }

  /* ═══════════════════════════════════════════
     LEAVE HISTORY — Agatha's decision log
     Shows all leave requests she has acted on
  ═══════════════════════════════════════════ */
  renderLeaveHistory(){
    const body=$('mgr-hist-body');if(!body)return;
    const uid=this.user?.id;
    const filter=$('mgr-hist-filter')?.value||'';
    const isFinalApprover=uid===COUNTRY_LEADER_ID;

    // Get all leave where this manager made a decision
    let items=this.leave.filter(l=>{
      if(isFinalApprover){
        return l.finalApproverId===uid&&(l.finalApproverStatus==='Approved'||l.finalApproverStatus==='Rejected');
      } else {
        return l.supervisorId===uid&&(l.supervisorStatus==='Approved'||l.supervisorStatus==='Rejected');
      }
    });

    if(filter){
      items=items.filter(l=>isFinalApprover?(l.finalApproverStatus===filter):(l.supervisorStatus===filter));
    }

    const cnt=$('mgr-hist-count');if(cnt)cnt.textContent=items.length;

    if(!items.length){body.innerHTML='<tr><td colspan="8"><div class="empty"><div class="empty-ico">📒</div>No leave decisions yet</div></td></tr>';return;}

    body.innerHTML=items.slice().reverse().map(l=>{
      const decision=isFinalApprover?(l.finalApproverStatus||'—'):(l.supervisorStatus||'—');
      const note=isFinalApprover?(l.finalApproverNote||'—'):(l.supervisorNote||'—');
      const decBadge=decision==='Approved'
        ?'<span class="stage-badge stage-ok">✓ Approved</span>'
        :'<span class="stage-badge stage-rej">✗ Rejected</span>';
      return`<tr>
        <td><strong>${l.name}</strong></td>
        <td style="font-size:.76rem;color:var(--text2)">${l.unit}</td>
        <td>${l.type}</td>
        <td style="font-size:.76rem">${fmtISO(l.startDate)} → ${fmtISO(l.endDate)}</td>
        <td>${l.days}</td>
        <td>${decBadge}</td>
        <td style="font-size:.76rem;color:var(--text2)">${note}</td>
        <td style="font-size:.74rem;color:var(--text3)">${l.updatedAt?fmtDT(l.updatedAt):(l.appliedAt?fmtDT(l.appliedAt):'—')}</td>
      </tr>`;
    }).join('');
  }

  /* ═══════════════════════════════════════════
     LEAVE REGISTER — HR record (Admin/Edna)
     Official record of all finalized leave
  ═══════════════════════════════════════════ */
  renderLeaveRegister(){
    const body=$('ad-reg-body');if(!body)return;
    const filter=$('ad-reg-filter')?.value||'';
    const unitFilter=$('ad-reg-unit')?.value||'';

    let items=this.leave.filter(l=>l.status==='Approved'||l.status==='Rejected');
    if(filter)items=items.filter(l=>l.status===filter);
    if(unitFilter)items=items.filter(l=>l.unit===unitFilter);

    const cnt=$('ad-reg-count');if(cnt)cnt.textContent=items.length;

    if(!items.length){body.innerHTML='<tr><td colspan="11"><div class="empty"><div class="empty-ico">📒</div>No finalized leave records</div></td></tr>';return;}

    body.innerHTML=items.slice().reverse().map(l=>{
      const supName=this.staff[l.supervisorId]?.name||l.supervisorId||'—';
      const finalName=this.staff[l.finalApproverId]?.name||l.finalApproverId||'—';
      const statusBadge=l.status==='Approved'
        ?'<span class="stage-badge stage-ok">✓ Approved</span>'
        :'<span class="stage-badge stage-rej">✗ Rejected</span>';
      const notes=[l.supervisorNote,l.finalApproverNote].filter(n=>n).join(' · ')||'—';
      return`<tr>
        <td><strong>${l.name}</strong><div style="font-size:.68rem;color:var(--text3)">${l.staffId}</div></td>
        <td style="font-size:.76rem">${l.unit}</td>
        <td>${l.type}</td>
        <td style="font-size:.76rem">${fmtISO(l.startDate)}</td>
        <td style="font-size:.76rem">${fmtISO(l.endDate)}</td>
        <td>${l.days}</td>
        <td style="font-size:.74rem">${supName}<div style="font-size:.66rem">${l.supervisorStatus==='N/A'?'Skipped':l.supervisorStatus}</div></td>
        <td style="font-size:.74rem">${finalName}<div style="font-size:.66rem">${l.finalApproverStatus||'—'}</div></td>
        <td>${statusBadge}</td>
        <td style="font-size:.74rem;color:var(--text2);max-width:120px">${notes}</td>
        <td style="font-size:.72rem;color:var(--text3)">${l.updatedAt?fmtDT(l.updatedAt):(l.appliedAt?fmtDT(l.appliedAt):'—')}</td>
      </tr>`;
    }).join('');
  }

  exportLeaveRegister(){
    const filter=$('ad-reg-filter')?.value||'';
    const unitFilter=$('ad-reg-unit')?.value||'';
    let items=this.leave.filter(l=>l.status==='Approved'||l.status==='Rejected');
    if(filter)items=items.filter(l=>l.status===filter);
    if(unitFilter)items=items.filter(l=>l.unit===unitFilter);
    let csv='Staff ID,Name,Unit,Type,Start Date,End Date,Days,Supervisor,Supervisor Decision,Final Approver,Final Decision,Overall Status,Notes,Date\n';
    items.slice().reverse().forEach(l=>{
      const supName=this.staff[l.supervisorId]?.name||l.supervisorId||'';
      const finalName=this.staff[l.finalApproverId]?.name||l.finalApproverId||'';
      const notes=[l.supervisorNote,l.finalApproverNote].filter(n=>n).join(' | ')||'';
      csv+=`"${l.staffId}","${l.name}","${l.unit}","${l.type}","${l.startDate}","${l.endDate}","${l.days}","${supName}","${l.supervisorStatus}","${finalName}","${l.finalApproverStatus||''}","${l.status}","${notes}","${l.updatedAt||l.appliedAt||''}"\n`;
    });
    this._dl(csv,'THP_Leave_Register_'+Date.now()+'.csv','text/csv');
  }

  openLeaveModal(id){
    const lv=this.leave.find(l=>l.id===id);if(!lv)return;
    const uid=this.user?.id;
    const isFinalApprover=uid===COUNTRY_LEADER_ID||this._isActiveDelegate(uid);
    const supName=this.staff[lv.supervisorId]?.name||'—';
    const finalName=this.staff[lv.finalApproverId]?.name||'—';
    const _bs=(s)=>s==='Approved'?'stage-ok':s==='Rejected'?'stage-rej':s==='N/A'?'stage-ok':'stage-pend';
    $('lm-title').textContent=(isFinalApprover?'✅ Final Approval — ':'👤 Supervisor Review — ')+lv.name;
    let infoHTML=`<strong>Type:</strong> ${lv.type} &nbsp; <strong>Days:</strong> ${lv.days}<br>
      <strong>Dates:</strong> ${fmtISO(lv.startDate)} → ${fmtISO(lv.endDate)}<br>
      <strong>Reason:</strong> ${lv.reason||'—'}<br>`;
    if(lv.compRef)infoHTML+=`<strong>Compensatory Dates Worked:</strong> ${lv.compRef}<br>`;
    if(lv.sickNote)infoHTML+=`<strong>Medical Doc:</strong> ${this._renderSickNoteLink(lv.sickNote)}<br>`;
    if(lv.handoverNote)infoHTML+=`<strong>Handover Note:</strong> <span style="color:var(--text)">${lv.handoverNote}</span> <button class="bsm bsm-navy" style="margin-left:6px;font-size:.7rem" onclick="APP._dlHandover('${id}')">⬇ Download</button><br>`;
    infoHTML+=`<strong>Supervisor (${supName}):</strong> <span class="stage-badge ${_bs(lv.supervisorStatus)}">${lv.supervisorStatus}</span><br>
      <strong>Final (${finalName}):</strong> <span class="stage-badge ${_bs(lv.finalApproverStatus||lv.hrStatus)}">${lv.finalApproverStatus||lv.hrStatus||'Pending'}</span>`;
    $('lm-info').innerHTML=infoHTML;
    $('lm-note').value='';$('lm-id').value=id;$('leave-modal').classList.add('open');
  }

  /* ── Download handover note as text file ── */
  _dlHandover(leaveId){
    const lv=this.leave.find(l=>l.id===leaveId);if(!lv||!lv.handoverNote)return toast('No handover note','err');
    const content='HANDOVER NOTE\n'+('═'.repeat(40))+'\nStaff: '+lv.name+'\nType: '+lv.type+'\nDates: '+lv.startDate+' to '+lv.endDate+'\n'+('═'.repeat(40))+'\n\n'+lv.handoverNote;
    this._dl(content,'Handover_'+lv.name.replace(/\s/g,'_')+'_'+leaveId+'.txt','text/plain');
  }

  /* ── Decide leave — SERVER FIRST ── */
  async decideLeave(status){
    const id=$('lm-id').value,note=$('lm-note').value.trim();
    const lv=this.leave.find(l=>l.id===id);if(!lv)return;
    const uid=this.user?.id;
    const isFinalApprover=uid===COUNTRY_LEADER_ID||this._isActiveDelegate(uid);
    const stage=isFinalApprover?'final':'supervisor';

    /* SERVER FIRST */
    const extraEmailData={
      staffName:lv.name,staffEmail:lv.staffEmail||this.staff[lv.staffId]?.email||'',
      supervisorEmail:this.staff[lv.supervisorId]?.email||'',
      finalApproverEmail:this.staff[lv.finalApproverId]?.email||'',
      leaveType:lv.type,leaveDays:lv.days,
      startDate:lv.startDate,endDate:lv.endDate,
      decidedBy:this.user.name
    };
    const result=await API.updateLeave(id,status,note,stage,extraEmailData);
    if(!result||!result.success){toast('Server error — try again','err');return;}

    /* Update local cache */
    if(isFinalApprover){
      lv.finalApproverStatus=status;lv.finalApproverNote=note;lv.hrStatus=status;lv.hrNote=note;lv.status=status;
    } else {
      lv.supervisorStatus=status;lv.supervisorNote=note;
      if(status==='Rejected'){lv.finalApproverStatus='N/A';lv.hrStatus='N/A';lv.status='Rejected';}
      else{lv.finalApproverStatus='Pending';lv.status='Pending';}
    }
    this._cacheL();closeModal('leave-modal');
    this.renderMgrLeave();this.renderAdminLeave();
    if(this.renderStaffLeave)this.renderStaffLeave();
    this._updateNotifBadges();
    toast(`${status} — ${lv.name}`);
  }

  /* ── Admin records ── */
  renderAdmin(){
    const srch=($('ad-srch')?.value||'').toLowerCase(),mv=$('ad-mth')?.value,unit=$('ad-unit')?.value;
    let recs=this.records.slice();
    if(srch)recs=recs.filter(r=>r.name.toLowerCase().includes(srch)||r.id.toLowerCase().includes(srch));
    if(unit)recs=recs.filter(r=>r.unit===unit);
    if(mv){const[y,m]=mv.split('-').map(Number);recs=recs.filter(r=>{const d=new Date(r.in);return d.getFullYear()===y&&d.getMonth()===m-1;});}
    if(this._adFilter.status)recs=recs.filter(r=>r.status&&r.status.includes(this._adFilter.status));
    recs=this._applySort('ad',recs);
    const cnt=$('ad-count');if(cnt)cnt.textContent=recs.length;
    this._updateSortHeaders('ad-table',this._sort.ad);
    const body=$('ad-body');
    if(!recs.length){body.innerHTML='<tr><td colspan="8"><div class="empty"><div class="empty-ico">📭</div>No records</div></td></tr>';return;}
    body.innerHTML=recs.map(r=>`<tr><td>${fmtD(r.date||r.in)}</td><td><strong>${r.name}</strong></td><td>${r.unit}</td><td>${fmtT(r.in)}</td><td>${r.out?fmtT(r.out):'<span style="color:var(--teal)">Active</span>'}</td><td>${r.hours||'--'}</td><td>${this._bdg(r.status)}</td><td><button class="bsm" style="background:rgba(239,68,68,.15);color:#ef4444;border:1px solid rgba(239,68,68,.25)" onclick="APP.deleteRecord('${r.id}','${r.in}')">🗑</button></td></tr>`).join('');
  }
  setAdFilter(key,val,el){this._adFilter[key]=val;el.closest('.filter-chips').querySelectorAll('.chip').forEach(c=>c.classList.remove('active'));el.classList.add('active');this.renderAdmin();}
  clearAdFilters(){this._adFilter={status:''};if($('ad-srch'))$('ad-srch').value='';if($('ad-mth'))$('ad-mth').value='';if($('ad-unit'))$('ad-unit').value='';document.querySelectorAll('#a-recs .chip').forEach(c=>c.classList.remove('active'));document.querySelector('#a-recs .chip-all')?.classList.add('active');this.renderAdmin();}

  sortTable(tbl,col){
    const s=this._sort[tbl];
    if(s.col===col)s.dir=s.dir==='asc'?'desc':'asc';else{s.col=col;s.dir='asc';}
    if(tbl==='ad')this.renderAdmin();else if(tbl==='mgr')this.renderMgrRecs();else if(tbl==='st')this.renderStaffLogs();
  }
  _applySort(tbl,recs){
    const{col,dir}=this._sort[tbl]||{col:'date',dir:'desc'};
    const mul=dir==='asc'?1:-1;
    return recs.slice().sort((a,b)=>{
      let av,bv;
      if(col==='date'){av=new Date(a.in).getTime();bv=new Date(b.in).getTime();}
      else if(col==='name'){av=(a.name||'').toLowerCase();bv=(b.name||'').toLowerCase();return av<bv?-mul:av>bv?mul:0;}
      else if(col==='hours'){av=parseFloat(a.hours)||0;bv=parseFloat(b.hours)||0;}
      else if(col==='status'){av=(a.status||'').toLowerCase();bv=(b.status||'').toLowerCase();return av<bv?-mul:av>bv?mul:0;}
      else{av=new Date(a.in).getTime();bv=new Date(b.in).getTime();}
      return(av-bv)*mul;
    });
  }
  _updateSortHeaders(tableId,{col,dir}){
    const tbl=document.getElementById(tableId);if(!tbl)return;
    tbl.querySelectorAll('th.sortable').forEach(th=>{
      th.classList.remove('sort-asc','sort-desc');
      const onclick=th.getAttribute('onclick')||'';
      const m=onclick.match(/'([^']+)'\)$/);
      if(m&&m[1]===col)th.classList.add(dir==='asc'?'sort-asc':'sort-desc');
    });
  }

  async deleteRecord(staffId,inTime){
    if(!confirm('Delete this record?'))return;
    /* SERVER FIRST — find by staff_id + clock_in, then delete by Supabase id */
    try{
      const rows=await API._get('attendance','staff_id=eq.'+encodeURIComponent(staffId)+'&clock_in=eq.'+encodeURIComponent(inTime)+'&limit=1');
      if(rows&&rows.length){
        await API._delete('attendance','id=eq.'+rows[0].id);
      }
    }catch(e){console.warn('Delete error:',e);}
    this.records=this.records.filter(r=>!(r.id===staffId&&r.in===inTime));
    this._cacheR();this.renderAdmin();this._renderDash();this._renderReports();
    toast('Record deleted','info');
  }

  _renderDash(){
    const tR=this.records.filter(r=>sameDay(r.date||r.in)),act=this.records.filter(r=>!r.out).length,tot=Object.keys(this.staff).length,pend=this.leave.filter(l=>l.status==='Pending').length;
    $('ad-stats').innerHTML=`
      <div class="stat"><div class="stat-lbl">Total Staff</div><div class="stat-val">${tot}</div></div>
      <div class="stat"><div class="stat-lbl">Present Today</div><div class="stat-val g">${tR.length}</div></div>
      <div class="stat"><div class="stat-lbl">Active Now</div><div class="stat-val a">${act}</div></div>
      <div class="stat"><div class="stat-lbl">Pending Leave</div><div class="stat-val p">${pend}</div></div>
      <div class="stat"><div class="stat-lbl">All Records</div><div class="stat-val t">${this.records.length}</div></div>`;
    const units=['Finance & Grant','Monitoring & Evaluation (M&E)','Partnership','Communication','Programs','Transport & Logistics','HR & Operations','Procurement','National Service','Intern','Security'];
    const mx=Math.max(...units.map(u=>this.records.filter(r=>r.unit===u).length),1);
    $('unit-bars').innerHTML=units.map(u=>{const c=this.records.filter(r=>r.unit===u).length;return`<div class="bar-row"><div class="bar-lbl">${u.split(' ')[0]}</div><div class="bar-trk"><div class="bar-fill" style="width:${Math.round(c/mx*100)}%"></div></div><div class="bar-n">${c}</div></div>`;}).join('');
    const comp=this.records.filter(r=>r.status==='Completed').length,early=this.records.filter(r=>r.status&&r.status.includes('Early')).length,active=this.records.filter(r=>r.status==='Active').length,total=comp+early+active||1;
    const cv=$('donut'),ctx=cv.getContext('2d');let ang=-Math.PI/2;ctx.clearRect(0,0,118,118);
    [{v:comp,c:'#22c55e'},{v:early,c:'#F5A623'},{v:active,c:'#3DBFB8'}].forEach(s=>{const sl=(s.v/total)*2*Math.PI;ctx.beginPath();ctx.moveTo(59,59);ctx.arc(59,59,48,ang,ang+sl);ctx.closePath();ctx.fillStyle=s.c;ctx.fill();ang+=sl;});
    const surfColor=getComputedStyle(document.documentElement).getPropertyValue('--surf').trim()||'#1a1f2e';
    const textColor=getComputedStyle(document.documentElement).getPropertyValue('--text').trim()||'#f1f5f9';
    ctx.beginPath();ctx.arc(59,59,25,0,2*Math.PI);ctx.fillStyle=surfColor;ctx.fill();
    ctx.fillStyle=textColor;ctx.font='bold 10px serif';ctx.textAlign='center';ctx.textBaseline='middle';ctx.fillText(this.records.length,59,59);
    $('donut-lgd').innerHTML=[{l:'Completed',c:'#22c55e',v:comp},{l:'Early',c:'#F5A623',v:early},{l:'Active',c:'#3DBFB8',v:active}].map(d=>`<div class="lgd-item"><div class="lgd-dot" style="background:${d.c}"></div>${d.l}: <strong>${d.v}</strong></div>`).join('');
  }

  /* ── Staff grid (admin) ── */
  _renderStaffGrid(){
    const grid=$('staff-grid'),ent=Object.entries(this.staff);
    if(!ent.length){grid.innerHTML='<div class="empty"><div class="empty-ico">👥</div>No staff</div>';return;}
    grid.innerHTML=ent.map(([id,s])=>{
      const col=s.color||avColor(s.name);
      return`<div class="scard"${s.active===false?' style="opacity:.55"':''}><div class="scard-top"><div class="av" style="background:${col}">${ini(s.name)}</div><div class="s-info"><div class="s-name">${s.name}</div><div class="s-id">${id}</div></div></div><div class="s-meta"><span class="s-unit">${s.unit||'—'}</span><span class="s-role-badge role-${s.role||'staff'}">${roleLabel(s.role)}</span></div><div class="scard-btns"><button class="btn-edit" onclick="APP.openEdit('${id}')">✏ Edit</button><button class="btn-edit" style="background:rgba(245,166,35,.1);border-color:rgba(245,166,35,.2);color:var(--gold)" onclick="APP.adminResetPass('${id}')">🔑 Reset</button><button class="btn-edit" style="background:${s.active===false?'rgba(22,163,74,.12);color:#16a34a':'rgba(100,116,139,.12);color:var(--text2)'}" onclick="APP.toggleStaffActive('${id}')">${s.active===false?'↩ Reactivate':'⏸ Deactivate'}</button><button class="btn-del" onclick="APP.delStaff('${id}')">🗑</button></div></div>`;
    }).join('');
  }
  async addStaff(){
    const id=$('ns-id').value.trim().toUpperCase(),name=$('ns-nm').value.trim(),unit=$('ns-unit').value.trim(),role=$('ns-role').value;
    const pass=$('ns-pw').value,email=$('ns-email').value.trim();
    const gender=$('ns-gender')?.value||'male';
    const supervisor=$('ns-supervisor')?.value||'';
    if(!id||!name||!unit||!pass)return toast('Fill all required fields','err');
    if(this.staff[id])return toast('Staff ID exists','err');
    if(!/^THPG\/\d{2}\/\d{4}(-\d+)?$/i.test(id))return toast('Format: THPG/MM/YYYY','err');
    if(pass.length<4)return toast('Min 4 char password','err');
    const color=avColor(name);
    const phoneVal=$('ns-phone')?.value.trim()||'';
    const staffData={name,unit,role,pass,color,email,gender,supervisor,phone:phoneVal};
    /* SERVER FIRST */
    const r=await API.saveStaff(id,staffData);
    if(!r||!r.success){toast('Server error','err');return;}
    this.staff[id]=staffData;this._cacheS();this._renderStaffGrid();this._populateSupervisorDropdown();
    ['ns-id','ns-nm','ns-pw','ns-email','ns-phone'].forEach(i=>{if($(i))$(i).value='';});
    this.audit('Staff created','Staff',name,id+' · '+unit);
    // Welcome message with first-time login instructions (SMS first, email fallback)
    const phone=$('ns-phone')?.value.trim()||'';
    if(email||phone){
      API.gasPost({action:'welcomeStaff',staffId:id,name,email,phone,
        unit,role,tempPass:pass,supervisor:this._sName(supervisor)||''})
        .then(r=>{
          if(r&&r.success)toast(name+' added — welcome '+(r.channel==='sms'?'SMS':'email')+' sent ✓');
          else toast(name+' added, but the welcome message failed to send','info');
        }).catch(()=>toast(name+' added (welcome message not sent)','info'));
    } else {
      toast(name+' added — no email or phone on file, so no welcome message sent','info');
    }
    toast(name+' added!');
  }
  _populateSupervisorDropdown(){
    const sel=$('ns-supervisor');if(!sel)return;
    const managers=Object.entries(this.staff).filter(([,s])=>s.role==='manager'||s.role==='country_leader');
    sel.innerHTML='<option value="">-- None --</option>'+managers.map(([id,s])=>`<option value="${id}">${s.name}</option>`).join('');
  }
  async delStaff(id){
    if(!confirm('Remove '+this.staff[id]?.name+'?'))return;
    await API.deleteStaff(id);
    delete this.staff[id];this._cacheS();this._renderStaffGrid();toast('Staff removed.');
  }
  openEdit(id){
    $('em-id').value=id;$('em-name').value=this.staff[id].name;$('em-unit').value=this.staff[id].unit;
    $('em-role').value=this.staff[id].role||'staff';$('em-email').value=this.staff[id].email||'';
    $('edit-modal').classList.add('open');
  }
  async saveEdit(){
    const id=$('em-id').value;
    this.staff[id].name=$('em-name').value.trim();this.staff[id].unit=$('em-unit').value;
    this.staff[id].role=$('em-role').value;this.staff[id].email=$('em-email').value.trim();
    this.staff[id].color=avColor(this.staff[id].name);
    await API.saveStaff(id,this.staff[id]);
    this._cacheS();closeModal('edit-modal');this._renderStaffGrid();toast('Updated.');
  }
  async adminResetPass(id){
    const s=this.staff[id];if(!s)return;
    const newPass=prompt(`Reset password for ${s.name}?\nEnter new (min 4) or blank for "1234".`);
    if(newPass===null)return;
    const plainPass=newPass.trim()||'1234';
    if(plainPass.length<4)return toast('Min 4 characters','err');
    const hashed=await hashPass(id,plainPass);
    const oldStored=s.pass;
    const r=await API.changePassword(id,oldStored,hashed);
    if(r&&r.success){
      s.pass=hashed;this._cacheS();this._renderStaffGrid();
      toast(`Password reset for ${s.name} ☁️`);
    } else {toast('Reset failed','err');}
    toast(`Tell ${s.name.split(' ')[0]}: new password is ${plainPass}`,'info');
  }

  _checkDefaultPass(prefix){
    const notice=$(prefix==='mgr'?'m-chpw-first-notice':'chpw-first-notice');
    if(!notice)return;
    const stored=this.staff[this.user.id]?.pass||'';
    // Show notice if password is still default plain text (not yet changed to a hash)
    const isDefault=!isHashed(stored)||stored==='1234';
    notice.style.display=isDefault?'flex':'none';
  }

  /* ═══════════════════════════════════════════
     SELF-SERVICE PROFILE
  ═══════════════════════════════════════════ */
  _renderProfileForm(prefix){
    const uid=this.user?.id;if(!uid)return;
    const s=this.staff[uid];if(!s)return;
    const p=prefix||'';
    const emailEl=$(p+'prof-email');if(emailEl)emailEl.value=s.email||'';
    const phoneEl=$(p+'prof-phone');if(phoneEl)phoneEl.value=s.phone||'';
    const ecEl=$(p+'prof-emergency');if(ecEl)ecEl.value=s.emergencyContact||'';
    const nameEl=$(p+'prof-name');if(nameEl)nameEl.textContent=s.name;
    const unitEl=$(p+'prof-unit');if(unitEl)unitEl.textContent=s.unit;
    const roleEl=$(p+'prof-role');if(roleEl)roleEl.textContent=roleLabel(s.role);
    const dobEl=$(p+'prof-dob');
    if(dobEl){dobEl.value='';API.getHRFile(uid).then(f=>{if(f?.dob)dobEl.value=String(f.dob).slice(0,10);});}
  }

  async saveProfile(prefix){
    const uid=this.user?.id;if(!uid)return;
    const p=prefix||'';
    const email=$(p+'prof-email')?.value.trim()||'';
    const phone=$(p+'prof-phone')?.value.trim()||'';
    const emergencyContact=$(p+'prof-emergency')?.value.trim()||'';
    const msgEl=$(p+'prof-msg');if(msgEl)msgEl.textContent='';

    if(msgEl)msgEl.innerHTML='<span style="color:var(--teal)">⏳ Saving…</span>';
    const dob=$(p+'prof-dob')?.value||'';
    const r=await API.updateProfile(uid,{email,phone,emergencyContact});
    if(dob)API._upsert('hr_staff_files',[{staff_id:uid,dob,phone}]).catch(()=>{});
    if(r&&r.success){
      // Update local cache
      this.staff[uid].email=email;
      this.staff[uid].phone=phone;
      this.staff[uid].emergencyContact=emergencyContact;
      this.user.email=email;
      this._cacheS();
      if(msgEl)msgEl.innerHTML='<span style="color:var(--green)">✓ Profile updated!</span>';
      toast('Profile saved!');
    } else {
      if(msgEl)msgEl.innerHTML='<span style="color:var(--red)">Failed to save. Try again.</span>';
    }
  }

  async changePassword(ctx=''){
    const pfx=ctx==='mgr'?'m-chpw-':'chpw-';
    const oldPass=$(pfx+'old').value.trim(),newPass=$(pfx+'new').value.trim(),confirmVal=$(pfx+'confirm').value.trim();
    const msgEl=$(pfx+'msg');msgEl.textContent='';
    if(!oldPass||!newPass||!confirmVal){msgEl.innerHTML='<span style="color:var(--red)">Fill all fields.</span>';return;}
    if(newPass.length<4){msgEl.innerHTML='<span style="color:var(--red)">Min 4 characters.</span>';return;}
    if(newPass!==confirmVal){msgEl.innerHTML='<span style="color:var(--red)">Don\'t match.</span>';return;}
    if(newPass===oldPass){msgEl.innerHTML='<span style="color:var(--red)">Must be different.</span>';return;}

    const uid=this.user.id;
    const newHashed=await hashPass(uid,newPass);
    msgEl.innerHTML='<span style="color:var(--teal)">⏳ Saving…</span>';

    /* Send the plain-text old password directly to the server.
       The server stores passwords as plain text (e.g. "1234") and
       compares with String() coercion, so this always works. */
    const session=getSession();
    let r=await API.changePassword(uid,oldPass,newHashed,session?.token);

    /* If that failed, try with the hashed version of old password
       (in case password was previously migrated to a hash) */
    if(!r||!r.success){
      const oldHashed=await hashPass(uid,oldPass);
      r=await API.changePassword(uid,oldHashed,newHashed,session?.token);
    }

    if(r&&r.success){
      this.staff[uid].pass=newHashed;this._cacheS();
      this._loginRawPass=null; // clear
      $(pfx+'old').value='';$(pfx+'new').value='';$(pfx+'confirm').value='';
      this._checkDefaultPass(ctx);
      msgEl.innerHTML='<span style="color:var(--green)">✅ Password changed — synced ☁️</span>';
      toast('Password updated!');
      setTimeout(()=>{
        if(ctx==='mgr')showPanel('m-dash','sb-mgr',null);else showPanel('p-clock','sb-staff',null);
      },2000);
    } else {
      msgEl.innerHTML=`<span style="color:var(--red)">${r?.error||'Failed. Try again.'}</span>`;
    }
  }

  /* ── Reports ── */
  _renderReports(){
    const body=$('rep-body');if(!body)return;
    body.innerHTML=Object.entries(this.staff).map(([id,s])=>{
      const recs=this.records.filter(r=>r.id===id&&r.out),hrs=recs.reduce((a,r)=>a+parseFloat(r.hours||0),0);
      const early=recs.filter(r=>r.status.includes('Early')).length,avg=recs.length?fx(hrs/recs.length):'0.00';
      const rate=recs.length?Math.min(100,Math.round((hrs/(recs.length*8))*100)):0;
      const col=s.color||avColor(s.name);
      return`<tr><td style="color:var(--text2);font-size:.74rem">${id}</td><td><div style="display:flex;align-items:center;gap:7px"><div class="av av-sm" style="background:${col}">${ini(s.name)}</div><strong>${s.name}</strong></div></td><td>${s.unit}</td><td><span class="s-role-badge role-${s.role||'staff'}">${roleLabel(s.role)}</span></td><td>${recs.length}</td><td>${fx(hrs)}</td><td>${avg}</td><td>${early>0?`<span style="color:var(--gold)">${early}</span>`:early}</td><td><div style="display:flex;align-items:center;gap:6px"><div style="flex:1;height:5px;background:var(--surf2);border-radius:3px"><div style="width:${rate}%;height:100%;background:var(--green);border-radius:3px"></div></div><span style="font-size:.7rem">${rate}%</span></div></td></tr>`;
    }).join('');
  }

  /* ── Manager reports ── */
  _mgrRepStaffSearch(){return ($('mgr-rep-staff')?.value||'').trim().toLowerCase();}
  /* Search box must re-run whichever report is selected */
  onRepSearch(){
    clearTimeout(this._repT);
    this._repT=setTimeout(()=>this.renderMgrReport(),250);
  }
  _mgrRepFilter(recs){
    const from=$('mgr-rep-from')?.value,to=$('mgr-rep-to')?.value;
    if(from)recs=recs.filter(r=>new Date(r.date||r.in)>=new Date(from));
    if(to)recs=recs.filter(r=>new Date(r.date||r.in)<=new Date(to+'T23:59:59'));
    const q=this._mgrRepStaffSearch();
    if(q)recs=recs.filter(r=>(r.id||'').toLowerCase().includes(q)||(r.name||'').toLowerCase().includes(q));
    return recs;
  }
  _mgrRepDays(){
    const from=$('mgr-rep-from')?.value,to=$('mgr-rep-to')?.value;
    const now=new Date(),y=now.getFullYear(),m=now.getMonth();
    const start=from?new Date(from):new Date(y,m,1);
    const end=to?new Date(to+'T23:59:59'):now;
    const days=[];
    for(let d=new Date(start);d<=end;d.setDate(d.getDate()+1)){if(!isWeekend(d))days.push(new Date(d));}
    return days;
  }
  clearMgrRepDates(){if($('mgr-rep-from'))$('mgr-rep-from').value='';if($('mgr-rep-to'))$('mgr-rep-to').value='';if($('mgr-rep-staff'))$('mgr-rep-staff').value='';this.renderMgrReport();}
  /* ── Multi-area reporting hub ── */
  async _renderOtherReport(type){
    const hdr=$('m-report-hdr'),body=$('m-report-body'),sub=$('m-report-sub');
    if(!hdr||!body)return;
    body.innerHTML='<tr><td colspan="9" style="color:var(--text3)">Loading…</td></tr>';
    const q=this._mgrRepStaffSearch();
    const staffList=Object.entries(this.staff).filter(([i,st])=>st.role!=='admin')
      .filter(([i,st])=>!q||i.toLowerCase().includes(q)||(st.name||'').toLowerCase().includes(q))
      .sort((a,b)=>a[1].name.localeCompare(b[1].name));
    const H=cols=>hdr.innerHTML=cols.map(c=>`<th>${c}</th>`).join('');
    const none=n=>body.innerHTML=`<tr><td colspan="${n}"><div class="empty"><div class="empty-ico">📭</div>No records</div></td></tr>`;
    const titles={leave:'Leave Register',contracts:'Contract Status',staffdir:'Staff Directory',
      hrfiles:'Staff File Completeness',training:'Training & Capacity Building',appraisal:'Performance Appraisals',
      recruit:'Recruitment Pipeline',demographics:'Headcount & Demographics'};
    if(sub)sub.textContent=titles[type]||'';

    if(type==='leave'){
      H(['Staff','Unit','Type','Start','End','Days','Supervisor','Final','Status']);
      let rows=this.leave.slice();
      const from=$('mgr-rep-from')?.value,to=$('mgr-rep-to')?.value;
      if(from)rows=rows.filter(l=>String(l.endDate).slice(0,10)>=from);
      if(to)rows=rows.filter(l=>String(l.startDate).slice(0,10)<=to);
      if(q)rows=rows.filter(l=>(l.name||'').toLowerCase().includes(q)||(l.staffId||'').toLowerCase().includes(q));
      if(!rows.length)return none(9);
      rows.sort((a,b)=>String(b.startDate).localeCompare(String(a.startDate)));
      body.innerHTML=rows.map(l=>`<tr><td><strong>${l.name}</strong></td><td>${l.unit||''}</td><td>${l.type}</td>
        <td>${fmtISO(l.startDate)}</td><td>${fmtISO(l.endDate)}</td><td>${l.days}</td>
        <td style="font-size:.74rem">${l.supervisorStatus||''}</td><td style="font-size:.74rem">${l.finalApproverStatus||l.hrStatus||''}</td>
        <td>${_bdg(l.status)}</td></tr>`).join('');
      return;
    }
    if(type==='contracts'){
      H(['Staff','Unit','Start','End','Days Left','Status']);
      const rows=staffList.map(([i,st])=>({i,st,f:this._contractFlag(st.contractEnd)}))
        .sort((a,b)=>(a.f.days??99999)-(b.f.days??99999));
      if(!rows.length)return none(6);
      body.innerHTML=rows.map(r=>`<tr><td><strong>${r.st.name}</strong><br><span style="font-size:.7rem;color:var(--text3)">${r.i}</span></td>
        <td>${r.st.unit||'—'}</td><td>${r.st.contractStart?fmtISO(r.st.contractStart):'—'}</td>
        <td>${r.st.contractEnd?fmtISO(r.st.contractEnd):'—'}</td><td>${r.f.days??'—'}</td>
        <td><span class="c-flag ${r.f.cls}">${r.f.label}</span></td></tr>`).join('');
      return;
    }
    if(type==='staffdir'){
      H(['Staff ID','Name','Unit','Role','Email','Phone','Supervisor']);
      if(!staffList.length)return none(7);
      body.innerHTML=staffList.map(([i,st])=>`<tr><td style="font-size:.76rem">${i}</td><td><strong>${st.name}</strong></td>
        <td>${st.unit||'—'}</td><td>${roleLabel(st.role)}</td><td style="font-size:.74rem">${st.email||'—'}</td>
        <td style="font-size:.74rem">${st.phone||'—'}</td><td style="font-size:.74rem">${this._sName(st.supervisor)||'—'}</td></tr>`).join('');
      return;
    }
    if(type==='hrfiles'){
      H(['Staff','Unit','DOB','Phone','Next of Kin','SSNIT','File Status']);
      const files=await API.getAllHRFiles();const fm={};files.forEach(f=>fm[f.staff_id]=f);
      if(!staffList.length)return none(7);
      body.innerHTML=staffList.map(([i,st])=>{
        const f=fm[i]||{};const n=[f.dob,f.phone,f.next_of_kin,f.ssnit_number].filter(v=>v&&String(v).trim()).length;
        const flag=!fm[i]?'<span class="c-flag none">No file</span>':n>=4?'<span class="c-flag green">Complete</span>':n>=2?'<span class="c-flag amber">Partial ('+n+'/4)</span>':'<span class="c-flag red">Started</span>';
        const tick=v=>v&&String(v).trim()?'✓':'—';
        return `<tr><td><strong>${st.name}</strong></td><td>${st.unit||'—'}</td><td>${tick(f.dob)}</td>
          <td>${tick(f.phone)}</td><td>${tick(f.next_of_kin)}</td><td>${tick(f.ssnit_number)}</td><td>${flag}</td></tr>`;}).join('');
      return;
    }
    if(type==='training'){
      H(['Staff','Course','Provider','Completed','Expiry','Status']);
      const rows=await API._get('training_records','order=completed_date.desc.nullslast&limit=400')||[];
      const f=q?rows.filter(r=>(this._sName(r.staff_id)||'').toLowerCase().includes(q)):rows;
      if(!f.length)return none(6);
      const today=new Date().toISOString().slice(0,10);
      body.innerHTML=f.map(r=>{
        const exp=r.expiry_date?String(r.expiry_date).slice(0,10):'';
        const st=exp?(exp<today?'<span class="c-flag red">Expired</span>':'<span class="c-flag green">Valid</span>'):'<span class="c-flag none">—</span>';
        return `<tr><td><strong>${this._sName(r.staff_id)}</strong></td><td>${r.course||''}</td><td>${r.provider||'—'}</td>
          <td>${r.completed_date?String(r.completed_date).slice(0,10):'—'}</td><td>${exp||'—'}</td><td>${st}</td></tr>`;}).join('');
      return;
    }
    if(type==='appraisal'){
      H(['Staff','Period','Type','Supervisor','Score','Rating','Status']);
      const rows=await API._get('performance_appraisals','order=review_date.desc.nullslast&limit=300')||[];
      const f=q?rows.filter(r=>(this._sName(r.staff_id)||'').toLowerCase().includes(q)):rows;
      if(!f.length)return none(7);
      body.innerHTML=f.map(r=>{const sc=(+r.final_score||0);
        return `<tr><td><strong>${this._sName(r.staff_id)}</strong></td><td>${r.period||'—'}</td><td>${r.review_type||''}</td>
          <td style="font-size:.76rem">${this._sName(r.line_manager)||'—'}</td><td><strong>${sc.toFixed(2)}</strong>/5</td>
          <td>${this._ratingWord(sc)}</td><td><span class="c-flag ${r.status==='Closed'||r.status==='Acknowledged'?'green':'amber'}">${r.status||'Draft'}</span></td></tr>`;}).join('');
      return;
    }
    if(type==='recruit'){
      H(['Candidate','Position','Contact','Stage','Rating','Applied']);
      const vacs=await API._get('recruitment_vacancies','select=id,position')||[];
      const vm={};vacs.forEach(v=>vm[v.id]=v.position);
      const rows=await API._get('recruitment_applicants','order=applied_date.desc.nullslast&limit=400')||[];
      const f=q?rows.filter(r=>(r.name||'').toLowerCase().includes(q)):rows;
      if(!f.length)return none(6);
      body.innerHTML=f.map(r=>`<tr><td><strong>${r.name}</strong></td><td>${vm[r.vacancy_id]||'—'}</td>
        <td style="font-size:.74rem">${r.email||''}<br>${r.phone||''}</td><td>${this._rcStageFlag(r.stage)}</td>
        <td>${this._stars(r.rating)}</td><td style="font-size:.76rem">${r.applied_date?String(r.applied_date).slice(0,10):'—'}</td></tr>`).join('');
      return;
    }
    if(type==='individual'){
      const q2=this._mgrRepStaffSearch();
      hdr.innerHTML='<th style="width:32%">Item</th><th>Detail</th>';
      if(!q2){body.innerHTML='<tr><td colspan="2"><div class="empty"><div class="empty-ico">👤</div>Type a staff ID or name in the search box above to pull that person\'s full record</div></td></tr>';return;}
      const hit=Object.entries(this.staff).find(([i,st])=>st.role!=='admin'&&(i.toLowerCase().includes(q2)||(st.name||'').toLowerCase().includes(q2)));
      if(!hit)return none(2);
      const [sid,st]=hit;
      const f=(await API._get('hr_staff_files','staff_id=eq.'+encodeURIComponent(sid))||[])[0]||{};
      const lv=this.leave.filter(l=>l.staffId===sid);
      const tr=await API._get('training_records','staff_id=eq.'+encodeURIComponent(sid))||[];
      const ap=await API._get('performance_appraisals','staff_id=eq.'+encodeURIComponent(sid)+'&order=review_date.desc')||[];
      const asg=await API._get('assets','assigned_to=eq.'+encodeURIComponent(sid))||[];
      const ck=await API._get('staff_checklists','staff_id=eq.'+encodeURIComponent(sid))||[];
      const days=this._mgrRepDays().filter(d=>!isHoliday(d));
      let pres=0,onl=0;
      days.forEach(d=>{const ds=fmtD(d.toISOString());
        if(leaveOnDate(this.leave,sid,d.toISOString().slice(0,10)))onl++;
        else if(this.records.some(r=>r.id===sid&&fmtD(r.date||r.in)===ds))pres++;});
      const cf=this._contractFlag(st.contractEnd);
      const row=(k,v)=>`<tr><td style="color:var(--text2)">${k}</td><td>${v||'—'}</td></tr>`;
      const sec=t=>`<tr><td colspan="2" style="background:var(--surf2);font-weight:700;color:var(--teal);font-size:.76rem;letter-spacing:.4px">${t}</td></tr>`;
      let h=sec('EMPLOYMENT')
        +row('Employee ID',sid)+row('Name','<strong>'+st.name+'</strong>')
        +row('Department / Unit',st.unit)+row('Position',st.position)
        +row('Employment Type',st.employmentType||'Staff')
        +row('Date Joined',st.dateJoined?fmtISO(st.dateJoined):'')
        +row('Reporting Manager',this._sName(st.supervisor))
        +row('Role',roleLabel(st.role));
      h+=sec('CONTRACT &amp; PROBATION')
        +row('Contract Start',st.contractStart?fmtISO(st.contractStart):'')
        +row('Contract End',st.contractEnd?fmtISO(st.contractEnd):'')
        +row('Contract Status','<span class="c-flag '+cf.cls+'">'+cf.label+'</span>')
        +row('Probation Status',f.probation_status||'N/A')
        +row('Probation Ends',f.probation_end?fmtISO(f.probation_end):'')
        +row('Date Confirmed',f.confirmed_date?fmtISO(f.confirmed_date):'');
      h+=sec('PERSONAL &amp; STATUTORY')
        +row('Date of Birth',f.dob?fmtISO(f.dob):'')
        +row('Phone',f.phone||st.phone)+row('Email',st.email)
        +row('Residential Address',f.residential_address)
        +row('Emergency Contact',f.emergency_contact)
        +row('Next of Kin',(f.next_of_kin||'')+(f.next_of_kin_phone?' · '+f.next_of_kin_phone:''))
        +row('SSNIT Number',f.ssnit_number)+row('Qualifications',f.qualifications);
      h+=sec('ATTENDANCE (selected period)')
        +row('Working Days',days.length)+row('Days Present','<strong>'+pres+'</strong>')
        +row('Days On Leave',onl)+row('Days Absent',Math.max(0,days.length-pres-onl));
      h+=sec('LEAVE HISTORY')
        +(lv.length?lv.map(l=>row(l.type,fmtISO(l.startDate)+' → '+fmtISO(l.endDate)+' · '+l.days+'d · '+l.status)).join(''):row('','No leave recorded'));
      h+=sec('APPRAISALS')
        +(ap.length?ap.map(x=>row(x.period||'',(+x.final_score||0).toFixed(2)+'/5 · '+this._ratingWord(x.final_score)+' · '+(x.status||''))).join(''):row('','No appraisal on record'));
      h+=sec('TRAINING')
        +(tr.length?tr.map(t=>row(t.course,(t.provider||'')+(t.completed_date?' · completed '+String(t.completed_date).slice(0,10):'')+(t.expiry_date?' · expires '+String(t.expiry_date).slice(0,10):''))).join(''):row('','No training recorded'));
      h+=sec('ASSETS HELD')
        +(asg.length?asg.map(x=>row(x.id,(x.name||'')+' · '+(x.condition||''))).join(''):row('','No assets assigned'));
      h+=sec('ON/OFFBOARDING');
      if(ck.length){ck.forEach(c=>{let it=[];try{it=JSON.parse(c.items||'[]');}catch(e){}
        h+=row(c.kind==='offboarding'?'Offboarding':'Onboarding',it.filter(x=>x.done).length+'/'+it.length+' complete · '+(c.status||''));});}
      else h+=row('','No checklist started');
      let dc=[];try{dc=JSON.parse(f.documents||'[]');}catch(e){}
      h+=sec('DOCUMENTS ON FILE')
        +(dc.length?dc.map(d=>row(d.name,d.at||'')).join(''):row('','No documents uploaded'));
      body.innerHTML=h;
      return;
    }
    if(type==='demographics'){
      H(['Category','Item','Count','Share']);
      const tot=staffList.length;
      const rows=[];
      const add=(cat,item,n)=>rows.push(`<tr><td style="color:var(--text3);font-size:.74rem">${cat}</td><td><strong>${item}</strong></td><td>${n}</td><td>${tot?Math.round(n/tot*100):0}%</td></tr>`);
      add('Overall','Total Staff',tot);
      add('Gender','Female',staffList.filter(([i,s])=>s.gender==='female').length);
      add('Gender','Male',staffList.filter(([i,s])=>(s.gender||'male')==='male').length);
      const uc={};staffList.forEach(([i,s])=>{const u=(s.unit||'Unassigned').trim();uc[u]=(uc[u]||0)+1;});
      Object.entries(uc).sort((a,b)=>b[1]-a[1]).forEach(([u,n])=>add('Unit',u,n));
      const ec={};staffList.forEach(([i,s])=>{const t=this._empType(s.unit);ec[t]=(ec[t]||0)+1;});
      Object.entries(ec).forEach(([t,n])=>add('Employment',t,n));
      const rc={};staffList.forEach(([i,s])=>{const r=roleLabel(s.role);rc[r]=(rc[r]||0)+1;});
      Object.entries(rc).forEach(([r,n])=>add('Role',r,n));
      add('Contracts','Expiring / expired (≤30d)',staffList.filter(([i,s])=>this._contractFlag(s.contractEnd).cls==='red').length);
      add('Contracts','No end date on file',staffList.filter(([i,s])=>!s.contractEnd).length);
      body.innerHTML=rows.join('');
      return;
    }
  }
  exportReportCSV(){
    const tbl=$('m-report-table');if(!tbl)return toast('Nothing to export','err');
    let csv='';
    tbl.querySelectorAll('tr').forEach(tr=>{
      const cells=[...tr.querySelectorAll('th,td')].map(td=>'"'+td.innerText.replace(/\s+/g,' ').trim().replace(/"/g,'""')+'"');
      if(cells.length)csv+=cells.join(',')+'\n';
    });
    const t=$('mgr-rep-type')?.value||'report';
    this._dl(csv,'THP_'+t+'_'+Date.now()+'.csv','text/csv');
  }
  renderMgrReport(){
    const _rt=$('mgr-rep-type')?.value||'attendance';
    if(_rt!=='attendance'){this._renderOtherReport(_rt);return;}
    const isHR=this.user.id===HR_MANAGER_ID;
    const isFinance=this.user.id==='THPG/05/2025';
    const hdr=$('m-report-hdr'),body=$('m-report-body');if(!hdr||!body)return;
    const sub=$('m-report-sub');
    if(sub)sub.textContent=isHR?'All staff attendance & leave records':'Staff attendance summary — Present / Absent / On Leave / Holiday';
    if(isHR){
      let recs=this._mgrRepFilter(this.records.slice());
      hdr.innerHTML='<th>Date</th><th>Staff ID</th><th>Name</th><th>Unit</th><th>Clock In</th><th>Clock Out</th><th>Hours</th><th>Status</th>';
      let rows=recs.length?recs.slice().reverse().map(r=>`<tr><td>${fmtD(r.date||r.in)}</td><td style="color:var(--text2);font-size:.76rem">${r.id}</td><td><strong>${r.name}</strong></td><td>${r.unit}</td><td>${fmtT(r.in)}</td><td>${r.out?fmtT(r.out):'Active'}</td><td>${r.hours||'--'}</td><td>${this._bdg(r.status)}</td></tr>`).join(''):'<tr><td colspan="8"><div class="empty"><div class="empty-ico">📭</div>No records</div></td></tr>';

      // Add holiday section for HR report
      const allDays=this._mgrRepDays();
      const holDays=allDays.filter(dt=>isHoliday(dt));
      if(holDays.length){
        rows+=`<tr><td colspan="8" style="padding:1.2rem .5rem .5rem;border:none"><h4 style="margin:0;color:var(--gold)">📅 Public Holidays in Period</h4></td></tr>`;
        rows+=`<tr style="background:var(--surf2)"><th colspan="3">Date</th><th colspan="5">Holiday Name</th></tr>`;
        holDays.forEach(dt=>{
          const hName=getHolidayName(dt)||'Public Holiday';
          rows+=`<tr><td colspan="3">${fmtD(dt.toISOString())}</td><td colspan="5" style="color:var(--gold)">${hName}</td></tr>`;
        });
      }

      // Add leave register section for HR
      const leaveItems=this.leave.filter(l=>l.status==='Approved'||l.status==='Pending');
      if(leaveItems.length){
        rows+=`<tr><td colspan="8" style="padding:1.2rem .5rem .5rem;border:none"><h4 style="margin:0;color:var(--teal)">🏖 Leave Requests (Approved &amp; Pending)</h4></td></tr>`;
        rows+=`<tr style="background:var(--surf2)"><th>Staff</th><th>Unit</th><th>Type</th><th>Dates</th><th>Days</th><th>Supervisor</th><th>Final Status</th><th>Attachment</th></tr>`;
        leaveItems.slice().reverse().forEach(l=>{
          const supName=this.staff[l.supervisorId]?.name||'—';
          const faBdg=l.status==='Approved'?'<span class="stage-badge stage-ok">✓ Approved</span>':'<span class="stage-badge stage-pend">⏳ Pending</span>';
          const attach=l.sickNote?this._renderSickNoteLink(l.sickNote):'—';
          rows+=`<tr><td><strong>${l.name}</strong></td><td style="font-size:.76rem">${l.unit}</td><td>${l.type}</td><td style="font-size:.76rem">${fmtISO(l.startDate)} → ${fmtISO(l.endDate)}</td><td>${l.days}</td><td style="font-size:.76rem">${supName}</td><td>${faBdg}</td><td>${attach}</td></tr>`;
        });
      }
      body.innerHTML=rows;
    } else {
      const EXCLUDED_UNITS=['National Service','Intern'];
      let staffList=Object.entries(this.staff).filter(([,s])=>s.active!==false&&!EXCLUDED_UNITS.includes((s.unit||'').trim()));
      const q=this._mgrRepStaffSearch();
      if(q)staffList=staffList.filter(([id,s])=>id.toLowerCase().includes(q)||s.name.toLowerCase().includes(q));
      const allDays=this._mgrRepDays();
      const mode=$('mgr-rep-mode')?.value||'daily';
      if(mode==='summary'){
        const workDays=allDays.filter(dt=>!isHoliday(dt));
        hdr.innerHTML='<th>Staff ID</th><th>Name</th><th>Unit</th><th>Days Present</th><th>On Leave</th><th>Absent</th><th>Working Days</th>';
        const srows=staffList.map(([id,s])=>{
          let present=0,leave=0,absent=0;
          workDays.forEach(dt=>{
            const dateStr=fmtD(dt.toISOString());
            const onLv=leaveOnDate(this.leave,id,dt.toISOString().slice(0,10));
            if(onLv)leave++;
            else if(this.records.some(r=>r.id===id&&fmtD(r.date||r.in)===dateStr))present++;
            else absent++;
          });
          return{id,name:s.name,unit:s.unit,present,leave,absent};
        }).sort((a,b)=>b.present-a.present||a.name.localeCompare(b.name));
        body.innerHTML=srows.length?srows.map(r=>`<tr><td style="color:var(--text2);font-size:.76rem">${r.id}</td><td><strong>${r.name}</strong></td><td>${r.unit}</td><td><span class="badge b-ok">${r.present}</span></td><td><span class="badge" style="background:rgba(99,102,241,.15);color:#4338ca">${r.leave}</span></td><td><span class="badge b-err">${r.absent}</span></td><td>${workDays.length}</td></tr>`).join(''):'<tr><td colspan="7"><div class="empty"><div class="empty-ico">📭</div>No data</div></td></tr>';
        return;
      }
      const rows=[];
      allDays.forEach(dt=>{
        const dateStr=fmtD(dt.toISOString());
        const hol=isHoliday(dt);
        const holName=hol?getHolidayName(dt):null;
        if(hol){
          // Holiday row — show once for all staff
          rows.push({id:'—',name:'ALL STAFF',unit:'—',date:dateStr,dt,present:false,onLeave:null,holiday:true,holidayName:holName||'Public Holiday'});
        } else {
          staffList.forEach(([id,s])=>{
            const onLeave=leaveOnDate(this.leave,id,dt.toISOString().slice(0,10));
            const present=onLeave?false:this.records.some(r=>r.id===id&&fmtD(r.date||r.in)===dateStr);
            rows.push({id,name:s.name,unit:s.unit,date:dateStr,dt,present,onLeave,holiday:false});
          });
        }
      });
      rows.sort((a,b)=>b.dt-a.dt);
      hdr.innerHTML='<th>Staff ID</th><th>Date</th><th>Name</th><th>Unit</th><th>Status</th>';
      body.innerHTML=rows.length?rows.map(r=>{
        if(r.holiday)return`<tr style="background:rgba(245,166,35,.08)"><td style="color:var(--gold)">📅</td><td style="color:var(--gold);font-weight:600">${r.date}</td><td colspan="2" style="color:var(--gold);font-weight:600">${r.holidayName}</td><td><span class="badge" style="background:rgba(245,166,35,.15);color:#d97706">📅 Holiday</span></td></tr>`;
        let badge;if(r.present)badge='<span class="badge b-ok">✓ Present</span>';else if(r.onLeave)badge=`<span class="badge" style="background:rgba(99,102,241,.15);color:#4338ca">🌴 ${r.onLeave.type}</span>`;else badge='<span class="badge b-err">✗ Absent</span>';
        return`<tr><td style="color:var(--text2);font-size:.76rem">${r.id}</td><td>${r.date}</td><td><strong>${r.name}</strong></td><td>${r.unit}</td><td>${badge}</td></tr>`;
      }).join(''):'<tr><td colspan="5"><div class="empty"><div class="empty-ico">📭</div>No data</div></td></tr>';
    }
  }
  exportMgrReport(){
    const isHR=this.user.id===HR_MANAGER_ID;
    if(isHR){
      let recs=this._mgrRepFilter(this.records.slice()).reverse();
      let csv='ATTENDANCE RECORDS\nDate,Staff ID,Name,Unit,Clock In,Clock Out,Hours,Status\n';
      recs.forEach(r=>{csv+=`"${fmtD(r.date||r.in)}","${r.id}","${r.name}","${r.unit}","${fmtT(r.in)}","${r.out?fmtT(r.out):'Active'}","${r.hours||'--'}","${r.status}"\n`;});
      // Add leave section
      csv+='\nLEAVE REQUESTS (Approved & Pending)\nStaff ID,Name,Unit,Type,Start Date,End Date,Days,Supervisor,Status,Attachment\n';
      const leaveItems=this.leave.filter(l=>l.status==='Approved'||l.status==='Pending');
      leaveItems.slice().reverse().forEach(l=>{
        const supName=this.staff[l.supervisorId]?.name||'';
        csv+=`"${l.staffId}","${l.name}","${l.unit}","${l.type}","${l.startDate}","${l.endDate}","${l.days}","${supName}","${l.status}","${l.sickNote||''}"\n`;
      });
      this._dl(csv,'THP_HR_Report_'+Date.now()+'.csv','text/csv');
    }
    else{const EXCLUDED_UNITS=['National Service','Intern'];const staffList=Object.entries(this.staff).filter(([,s])=>s.active!==false&&!EXCLUDED_UNITS.includes((s.unit||'').trim()));const allDays=this._mgrRepDays();let csv='Staff ID,Date,Name,Unit,Status\n';allDays.forEach(dt=>{const dateStr=fmtD(dt.toISOString());const hol=isHoliday(dt);if(hol){const holName=getHolidayName(dt)||'Public Holiday';csv+=`"—","${dateStr}","ALL STAFF","—","Holiday — ${holName}"\n`;}else{staffList.forEach(([id,s])=>{const present=this.records.some(r=>r.id===id&&fmtD(r.date||r.in)===dateStr);const onLeave=present?null:leaveOnDate(this.leave,id,dt.toISOString().slice(0,10));csv+=`"${id}","${dateStr}","${s.name}","${s.unit}","${present?'Present':onLeave?'On Leave':'Absent'}"\n`;});}});this._dl(csv,'THP_Report_'+Date.now()+'.csv','text/csv');}
  }
  printMgrReport(){
    const html=this._buildReportHTML(false);
    if(!html)return;
    const w=window.open('','_blank');
    w.document.write(html);
    w.document.close();
  }

  /* ── QR & misc ── */
  _initEntQR(){const box=$('ent-qr-box');if(!box)return;box.innerHTML='';const url=window.location.href.split('?')[0];try{new QRCode(box,{text:url,width:195,height:195,colorDark:'#000',colorLight:'#fff',correctLevel:QRCode.CorrectLevel.H});}catch(e){}if($('ent-url-txt'))$('ent-url-txt').textContent=url;if($('hosted-url'))$('hosted-url').placeholder=url;}
  genEntrance(){const url=$('hosted-url').value.trim();if(!url)return toast('Enter a URL','err');$('ent-qr-box').innerHTML='';new QRCode($('ent-qr-box'),{text:url,width:195,height:195,colorDark:'#000',colorLight:'#fff',correctLevel:QRCode.CorrectLevel.H});$('ent-url-txt').textContent=url;toast('QR updated!');}
  _staffQR(){const box=$('st-qr-box');if(!box)return;box.innerHTML='';const url=window.location.href.split('?')[0]+'?staff='+this.user.id;$('st-qr-url').textContent=url;try{new QRCode(box,{text:url,width:148,height:148,colorDark:'#000',colorLight:'#fff',correctLevel:QRCode.CorrectLevel.H});}catch(e){}}

  async resetAllData(){
    if(!confirm('⚠️ Delete ALL records, leave, and reset staff?'))return;
    if(!confirm('FINAL: This cannot be undone. Proceed?'))return;
    this.records=[];this.leave=[];
    const r=await API.hydrate();
    if(r&&r.success)this.staff=r.staff||{};
    this._cacheR();this._cacheL();this._cacheS();
    this.renderAdmin();this._renderDash();this._renderStaffGrid();this._renderReports();this.renderAdminLeave();
    this._updateNotifBadges();
    toast('Data reset. Note: clear Google Sheets manually if needed.','info');
  }

  dlQR(boxId,fn){const c=document.querySelector('#'+boxId+' canvas');if(!c){toast('QR not ready','err');return;}const a=document.createElement('a');a.href=c.toDataURL('image/png');a.download=fn+'_'+Date.now()+'.png';a.click();}
  _bdg(s){if(!s)return'';if(s==='Active')return'<span class="badge b-active">● Active</span>';if(s.includes('Early'))return`<span class="badge b-early">⚠ Early</span>`;return'<span class="badge b-ok">✓ Done</span>';}

  /* ── Report HTML builder (shared by Word, PDF, Print) ── */
  _buildReportHTML(forExport){
    const tbl=$('m-report-table');if(!tbl)return'';
    const isHR=this.user.id===HR_MANAGER_ID;
    const now=new Date();
    const dateStr=now.toLocaleDateString('en-GB',{day:'2-digit',month:'long',year:'numeric'});
    const timeStr=now.toLocaleTimeString([],{hour:'2-digit',minute:'2-digit'});
    const fromVal=$('mgr-rep-from')?.value,toVal=$('mgr-rep-to')?.value;
    const periodLabel=fromVal&&toVal?`${fmtISO(fromVal)} — ${fmtISO(toVal)}`:`${now.toLocaleDateString('en-GB',{month:'long',year:'numeric'})} (Month to Date)`;
    const reportTitle=isHR?'Staff Attendance & Leave Report':'Staff Attendance Summary Report';
    const reportSub=isHR?'Includes all clock-in/out records, leave requests, and public holidays':'Present / Absent / On Leave / Holiday status per working day';
    const generatedBy=this.user.name+' ('+roleLabel(this.user.role)+')';
    const staffFilter=$('mgr-rep-staff')?.value.trim();
    const filterNote=staffFilter?`<br><strong>Filter:</strong> "${staffFilter}"`:'';

    const logoSrc=document.querySelector('.lo-logo')?.src||document.querySelector('img[alt="THP"]')?.src||'';

    return`<!DOCTYPE html><html><head><meta charset="UTF-8"><title>${reportTitle} — THP-Ghana</title>
<style>
  @page{size:A4 landscape;margin:15mm 12mm;}
  *{box-sizing:border-box;margin:0;padding:0;}
  body{font-family:'Segoe UI',Arial,sans-serif;color:#1e293b;font-size:11px;line-height:1.5;padding:0;}
  .page{padding:8mm;}
  .rpt-header{display:flex;align-items:center;justify-content:space-between;border-bottom:3px solid #2D3592;padding-bottom:12px;margin-bottom:6px;}
  .rpt-logo-block{display:flex;align-items:center;gap:14px;}
  .rpt-logo{height:52px;width:auto;}
  .rpt-org{font-size:15px;font-weight:700;color:#2D3592;line-height:1.3;}
  .rpt-org small{display:block;font-size:10px;font-weight:400;color:#64748b;letter-spacing:.5px;text-transform:uppercase;}
  .rpt-meta{text-align:right;font-size:9.5px;color:#64748b;line-height:1.6;}
  .rpt-meta strong{color:#1e293b;}
  .rpt-title-strip{background:#2D3592;color:#fff;padding:10px 16px;border-radius:6px;margin-bottom:12px;display:flex;justify-content:space-between;align-items:center;}
  .rpt-title-strip h1{font-size:14px;font-weight:700;margin:0;}
  .rpt-title-strip .rpt-period{font-size:10px;opacity:.9;}
  .rpt-title-strip .rpt-sub{font-size:9px;opacity:.75;margin-top:2px;}
  table{width:100%;border-collapse:collapse;font-size:10px;margin-bottom:14px;}
  th{background:#2D3592;color:#fff;padding:7px 6px;text-align:left;font-weight:600;font-size:9.5px;text-transform:uppercase;letter-spacing:.3px;border:1px solid #2D3592;}
  td{padding:6px;border:1px solid #e2e8f0;vertical-align:top;}
  tr:nth-child(even) td{background:#f8fafc;}
  .b-present,.b-ok,.b-done,.b-approved,.badge.b-ok,.stage-badge.stage-ok{background:#dcfce7;color:#166534;padding:2px 8px;border-radius:10px;font-size:9px;font-weight:600;display:inline-block;}
  .b-absent,.b-err,.badge.b-err{background:#fee2e2;color:#991b1b;padding:2px 8px;border-radius:10px;font-size:9px;font-weight:600;display:inline-block;}
  .b-leave,.badge{background:#e0e7ff;color:#3730a3;padding:2px 8px;border-radius:10px;font-size:9px;font-weight:600;display:inline-block;}
  .b-holiday{background:#fef3c7;color:#92400e;padding:2px 8px;border-radius:10px;font-size:9px;font-weight:600;display:inline-block;}
  .b-active,.badge.b-active{background:#ccfbf1;color:#0f766e;padding:2px 8px;border-radius:10px;font-size:9px;font-weight:600;display:inline-block;}
  .b-early,.badge.b-early{background:#fef9c3;color:#854d0e;padding:2px 8px;border-radius:10px;font-size:9px;font-weight:600;display:inline-block;}
  .stage-badge.stage-pend,.b-pending{background:#fef9c3;color:#854d0e;padding:2px 8px;border-radius:10px;font-size:9px;font-weight:600;display:inline-block;}
  .stage-badge.stage-rej{background:#fee2e2;color:#991b1b;padding:2px 8px;border-radius:10px;font-size:9px;font-weight:600;display:inline-block;}
  .hol-row td{background:#fffbeb !important;border-left:3px solid #F5A623;}
  .sig-block{display:flex;gap:60px;margin-top:30px;padding-top:8px;}
  .sig-line{flex:1;border-top:1px solid #94a3b8;padding-top:6px;font-size:9px;color:#64748b;}
  .sig-line strong{color:#1e293b;display:block;margin-bottom:2px;}
  .rpt-footer{border-top:2px solid #e2e8f0;padding-top:10px;margin-top:16px;display:flex;justify-content:space-between;align-items:center;font-size:8.5px;color:#94a3b8;}
  @media print{.no-print{display:none!important;}}
</style></head><body>
<div class="page">
  <div class="rpt-header">
    <div class="rpt-logo-block">
      ${logoSrc?`<img src="${logoSrc}" class="rpt-logo" alt="THP">`:''}
      <div class="rpt-org">The Hunger Project — Ghana<small>Staff Attendance & Leave Management System</small></div>
    </div>
    <div class="rpt-meta">
      <strong>Generated:</strong> ${dateStr} at ${timeStr}<br>
      <strong>By:</strong> ${generatedBy}<br>
      <strong>Report ID:</strong> RPT-${Date.now().toString(36).toUpperCase()}${filterNote}
    </div>
  </div>
  <div class="rpt-title-strip">
    <div><h1>${reportTitle}</h1><div class="rpt-sub">${reportSub}</div></div>
    <div class="rpt-period">${periodLabel}</div>
  </div>
  ${tbl.outerHTML
    .replace(/style="background:rgba\(245,166,35,\.08\)"/g,'class="hol-row"')
    .replace(/style="background:rgba\(245,166,35,\.15\);color:#d97706"/g,'class="b-holiday"')
    .replace(/style="background:rgba\(99,102,241,\.15\);color:#4338ca"/g,'class="b-leave"')
  }
  <div class="sig-block">
    <div class="sig-line"><strong>Prepared by:</strong>${generatedBy}</div>
    <div class="sig-line"><strong>Reviewed by:</strong>______________________</div>
    <div class="sig-line"><strong>Date:</strong>${dateStr}</div>
  </div>
  <div class="rpt-footer">
    <div>CONFIDENTIAL — For internal use only. The Hunger Project — Ghana.</div>
    <div>Page 1</div>
  </div>
</div>
${forExport?'':`<div class="no-print" style="text-align:center;padding:16px">
  <button onclick="window.print()" style="padding:10px 28px;font-size:14px;background:#2D3592;color:#fff;border:none;border-radius:8px;cursor:pointer;font-weight:600">🖨 Print Report</button>
  <button onclick="window.close()" style="padding:10px 28px;font-size:14px;background:#ef4444;color:#fff;border:none;border-radius:8px;cursor:pointer;font-weight:600;margin-left:10px">✕ Close</button>
</div>`}
</body></html>`;
  }

  /* ── Export as Word (.doc) ── */
  exportMgrWord(){
    const html=this._buildReportHTML(true);
    if(!html)return toast('No report to export','err');
    const wordContent='<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40"><head><meta charset="UTF-8"><!--[if gte mso 9]><xml><w:WordDocument><w:View>Print</w:View><w:Zoom>100</w:Zoom></w:WordDocument></xml><![endif]--></head><body>'+html+'</body></html>';
    const blob=new Blob(['\ufeff',wordContent],{type:'application/msword'});
    const url=URL.createObjectURL(blob);
    const a=document.createElement('a');a.href=url;a.download='THP_Report_'+Date.now()+'.doc';a.click();
    URL.revokeObjectURL(url);
    toast('Word document downloaded ✓');
  }

  /* ── Export as PDF (via print dialog) ── */
  exportMgrPDF(){
    const html=this._buildReportHTML(true);
    if(!html)return toast('No report to export','err');
    const w=window.open('','_blank');
    w.document.write(html+'<script>setTimeout(()=>{window.print();},500);<\/script>');
    w.document.close();
    toast('Print dialog opened — select "Save as PDF" to download.','info');
  }
  exportCSV(mode){
    let recs=this.records.slice();
    if(mode==='staff'||mode==='mgr-my')recs=recs.filter(r=>r.id===this.user.id);
    let csv='Date,Staff ID,Name,Unit,Clock In,Clock Out,Hours,Status\n';
    recs.forEach(r=>{csv+=`"${r.date}","${r.id}","${r.name}","${r.unit}","${new Date(r.in).toLocaleString()}","${r.out?new Date(r.out).toLocaleString():'--'}","${r.hours||'--'}","${r.status}"\n`;});
    this._dl(csv,'THP_Attendance_'+Date.now()+'.csv','text/csv');
  }
  exportSummary(){
    let csv='Staff ID,Name,Unit,Role,Days Present,Total Hours,Avg Hours,Early Exits\n';
    Object.entries(this.staff).forEach(([id,s])=>{const recs=this.records.filter(r=>r.id===id&&r.out),hrs=recs.reduce((a,r)=>a+parseFloat(r.hours||0),0);csv+=`"${id}","${s.name}","${s.unit}","${s.role||'staff'}","${recs.length}","${fx(hrs)}","${recs.length?fx(hrs/recs.length):'0.00'}","${recs.filter(r=>r.status.includes('Early')).length}"\n`;});
    this._dl(csv,'THP_Summary_'+Date.now()+'.csv','text/csv');
  }
  _dl(c,n,t){const b=new Blob([c],{type:t}),u=URL.createObjectURL(b);const a=document.createElement('a');a.href=u;a.download=n;a.click();URL.revokeObjectURL(u);}

  /* ── File to Base64 helper ── */
  _fileToBase64(file){
    return new Promise((resolve,reject)=>{
      const reader=new FileReader();
      reader.onload=()=>{
        const base64=reader.result.split(',')[1];
        resolve(base64);
      };
      reader.onerror=reject;
      reader.readAsDataURL(file);
    });
  }

  /* ── Sick note link renderer ── */
  _renderSickNoteLink(sickNote){
    if(!sickNote)return'—';
    // Format: "filename | https://drive.google.com/..."
    if(sickNote.includes('|')){
      const parts=sickNote.split('|').map(s=>s.trim());
      const fileName=parts[0];
      const url=parts[1];
      if(url&&url.startsWith('http')){
        return `<a href="${url}" target="_blank" rel="noopener" style="color:var(--teal);text-decoration:underline;font-size:.8rem">📎 ${fileName} ↗</a>`;
      }
    }
    return `<span style="color:var(--teal);font-size:.8rem">📎 ${sickNote}</span>`;
  }

  /* ═══════════════════════════════════════════
     ADMIN HOLIDAY MANAGEMENT PANEL
  ═══════════════════════════════════════════ */
  renderAdminHolidays(){
    const body=$('ad-holidays-body');if(!body)return;
    const yearInput=$('ad-hol-year');
    if(yearInput&&!yearInput._initialized){yearInput.value=new Date().getFullYear();yearInput._initialized=true;}
    const year=parseInt(yearInput?.value)||new Date().getFullYear();

    // Get all holidays: built-in + admin-managed for this year
    const builtInNames=ghHolidayNames(year);
    const builtInDates=Object.keys(builtInNames);
    const adminHols=(this.holidays||[]).filter(h=>{
      if(!h.date)return false;
      const hYear=parseInt(h.date.slice(0,4));
      return h.recurring==='yes'||hYear===year;
    });

    // Merge into one list
    const allRows=[];
    // Built-in
    builtInDates.forEach(d=>{
      allRows.push({date:d,name:builtInNames[d],type:'auto',id:null,recurring:'yes'});
    });
    // Admin-managed (skip dupes)
    const builtInSet=new Set(builtInDates);
    adminHols.forEach(h=>{
      if(!builtInSet.has(h.date)){
        allRows.push({date:h.date,name:h.name,type:h.type||'custom',id:h.id,recurring:h.recurring||'no'});
      } else {
        // Admin override of built-in — show admin version
        const idx=allRows.findIndex(r=>r.date===h.date);
        if(idx>=0){allRows[idx].name=h.name;allRows[idx].id=h.id;allRows[idx].type='override';}
      }
    });

    // Sort by date
    allRows.sort((a,b)=>a.date.localeCompare(b.date));

    const typeBadge=t=>{
      if(t==='auto')return'<span class="stage-badge" style="background:rgba(34,197,94,.15);color:#16a34a;font-size:.68rem">Built-in</span>';
      if(t==='fixed')return'<span class="stage-badge" style="background:rgba(59,130,246,.15);color:#2563eb;font-size:.68rem">Fixed</span>';
      if(t==='custom')return'<span class="stage-badge" style="background:rgba(245,166,35,.15);color:#d97706;font-size:.68rem">Custom</span>';
      if(t==='override')return'<span class="stage-badge" style="background:rgba(168,85,247,.15);color:#7c3aed;font-size:.68rem">Override</span>';
      return'<span class="stage-badge stage-pend" style="font-size:.68rem">'+t+'</span>';
    };

    const cnt=$('ad-hol-count');if(cnt)cnt.textContent=allRows.length;

    if(!allRows.length){body.innerHTML='<tr><td colspan="5"><div class="empty"><div class="empty-ico">📅</div>No holidays for '+year+'</div></td></tr>';return;}

    body.innerHTML=allRows.map(r=>{
      const dateObj=new Date(r.date+'T00:00:00');
      const dayName=dateObj.toLocaleDateString('en-GB',{weekday:'short'});
      const dateDisplay=dateObj.toLocaleDateString('en-GB',{day:'2-digit',month:'short',year:'numeric'});
      const isPast=dateObj<new Date(new Date().toISOString().slice(0,10)+'T00:00:00');
      const rowStyle=isPast?'opacity:.6':'';
      const actions=r.id
        ?`<button class="bsm" style="background:rgba(59,130,246,.1);color:var(--blue);border:1px solid rgba(59,130,246,.2)" onclick="APP.editHoliday('${r.id}')">✏</button>
           <button class="bsm" style="background:rgba(239,68,68,.1);color:#ef4444;border:1px solid rgba(239,68,68,.2)" onclick="APP.removeHoliday('${r.id}')">🗑</button>`
        :'<span style="color:var(--text3);font-size:.7rem">System</span>';
      return`<tr style="${rowStyle}"><td style="font-size:.76rem">${dateDisplay}<div style="font-size:.66rem;color:var(--text3)">${dayName}</div></td><td><strong>${r.name}</strong></td><td>${typeBadge(r.type)}</td><td style="font-size:.72rem;color:var(--text2)">${r.recurring==='yes'?'Every year':year+' only'}</td><td>${actions}</td></tr>`;
    }).join('');
  }

  async addHoliday(){
    const name=$('hol-name')?.value.trim();
    const date=$('hol-date')?.value;
    const type=$('hol-type')?.value||'custom';
    const recurring=$('hol-recurring')?.checked?'yes':'no';
    const msg=$('hol-msg');if(msg)msg.textContent='';

    if(!name||!date){if(msg)msg.innerHTML='<span style="color:var(--red)">Name and date required.</span>';return;}

    const year=parseInt(date.slice(0,4));
    const holiday={name,date,type,recurring,year:String(year)};

    // Check for editing
    const editId=$('hol-edit-id')?.value;
    if(editId){holiday.id=editId;}

    if(msg)msg.innerHTML='<span style="color:var(--teal)">⏳ Saving…</span>';
    const r=await API.saveHoliday(holiday);
    if(r&&r.success){
      // Refresh holidays
      const hr=await API.getHolidays();
      if(hr&&hr.holidays){this.holidays=hr.holidays;this._cacheH();}
      this.renderAdminHolidays();
      // Clear form
      if($('hol-name'))$('hol-name').value='';
      if($('hol-date'))$('hol-date').value='';
      if($('hol-recurring'))$('hol-recurring').checked=false;
      if($('hol-edit-id'))$('hol-edit-id').value='';
      if($('hol-form-title'))$('hol-form-title').textContent='Add Holiday';
      if(msg)msg.innerHTML='<span style="color:var(--green)">✓ Holiday saved!</span>';
      toast(editId?'Holiday updated!':'Holiday added!');
    } else {
      if(msg)msg.innerHTML=`<span style="color:var(--red)">${r?.error||'Failed to save.'}</span>`;
    }
  }

  editHoliday(id){
    const h=this.holidays.find(hol=>hol.id===id);if(!h)return;
    if($('hol-name'))$('hol-name').value=h.name;
    if($('hol-date'))$('hol-date').value=h.date;
    if($('hol-type'))$('hol-type').value=h.type||'custom';
    if($('hol-recurring'))$('hol-recurring').checked=h.recurring==='yes';
    if($('hol-edit-id'))$('hol-edit-id').value=id;
    if($('hol-form-title'))$('hol-form-title').textContent='Edit Holiday';
    if($('hol-msg'))$('hol-msg').textContent='';
    // Scroll to form
    $('hol-name')?.focus();
  }

  async removeHoliday(id){
    const h=this.holidays.find(hol=>hol.id===id);
    if(!h)return;
    if(!confirm('Remove "'+h.name+'" ('+h.date+')?'))return;
    const r=await API.deleteHoliday(id);
    if(r&&r.success){
      this.holidays=this.holidays.filter(hol=>hol.id!==id);
      this._cacheH();this.renderAdminHolidays();
      toast('Holiday removed.');
    } else {toast('Failed to remove','err');}
  }

  async seedGhanaHolidays(){
    const year=parseInt($('ad-hol-year')?.value)||new Date().getFullYear();
    if(!confirm('Seed all Ghana public holidays for '+year+'?\nThis adds standard holidays, estimated Eid dates, and Farmer\'s Day.'))return;
    toast('Seeding holidays for '+year+'…','info');
    const r=await API.seedGhanaHolidays(year);
    if(r&&r.success){
      const hr=await API.getHolidays();
      if(hr&&hr.holidays){this.holidays=hr.holidays;this._cacheH();}
      this.renderAdminHolidays();
      toast(`Holidays seeded for ${year}! Added: ${r.added}, Skipped: ${r.skipped}`);
    } else {
      toast('Seed failed','err');
    }
  }

  /* ═══════════════════════════════════════════
     HR STAFF FILES (Phase 1)
     Visible to Admin, Edna (HR), Agatha (CL)
  ═══════════════════════════════════════════ */
  async renderHRFiles(prefix){
    const p=prefix||'a-';
    const body=$(p+'hrfiles-body');if(!body)return;
    const q=($(p+'hr-search')?.value||'').trim().toLowerCase();
    const fUnit=$(p+'hr-unit')?.value||'';
    const fStat=$(p+'hr-status')?.value||'';
    body.innerHTML='<tr><td colspan="5" style="color:var(--text3)">Loading…</td></tr>';
    const files=await API.getAllHRFiles();
    const fileMap={};files.forEach(f=>fileMap[f.staff_id]=f);
    let list=Object.entries(this.staff).filter(([id,s])=>(s.role||'')!=='admin');
    const us=$(p+'hr-unit');
    if(us&&us.options.length<=1){
      [...new Set(list.map(([i,s])=>(s.unit||'').trim()).filter(Boolean))].sort().forEach(u=>{const o=document.createElement('option');o.value=u;o.textContent=u;us.appendChild(o);});
    }
    const statOf=([id,s])=>{const f=fileMap[id];if(!f)return'none';const core=[f.dob,f.phone,f.next_of_kin,f.ssnit_number].filter(v=>v&&String(v).trim()).length;return core>=4?'Complete':core>=2?'Partial':'Started';};
    const sm=$(p+'hr-summary');
    if(sm){
      const units=new Set(list.map(([i,s])=>(s.unit||'').trim()).filter(Boolean));
      const c={Complete:0,Partial:0,Started:0,none:0};list.forEach(e=>c[statOf(e)]++);
      sm.innerHTML=`<div class="cs-box"><div class="cs-num">${list.length}</div><div class="cs-lbl">Total Staff</div></div>
        <div class="cs-box"><div class="cs-num">${units.size}</div><div class="cs-lbl">Units</div></div>
        <div class="cs-box"><div class="cs-num" style="color:#16a34a">${c.Complete}</div><div class="cs-lbl">Complete</div></div>
        <div class="cs-box"><div class="cs-num" style="color:#d97706">${c.Partial}</div><div class="cs-lbl">Partial</div></div>
        <div class="cs-box"><div class="cs-num" style="color:var(--text3)">${c.none}</div><div class="cs-lbl">No File</div></div>`;
    }
    if(q)list=list.filter(([id,s])=>id.toLowerCase().includes(q)||(s.name||'').toLowerCase().includes(q));
    if(fUnit)list=list.filter(([i,s])=>(s.unit||'').trim()===fUnit);
    if(fStat)list=list.filter(e=>statOf(e)===fStat);
    list.sort((a,b)=>a[1].name.localeCompare(b[1].name));
    if(!list.length){body.innerHTML='<tr><td colspan="5"><div class="empty"><div class="empty-ico">📭</div>No staff found</div></td></tr>';return;}
    body.innerHTML=list.map(([id,s])=>{
      const f=fileMap[id];const st=statOf([id,s]);
      const badge=st==='Complete'?'<span class="c-flag green">✓ Complete</span>':st==='Partial'?'<span class="c-flag amber">◐ Partial</span>':st==='Started'?'<span class="c-flag red">◌ Started</span>':'<span class="c-flag none">No file</span>';
      return `<tr><td><strong>${s.name}</strong><br><span style="font-size:.72rem;color:var(--text3)">${id}</span></td>`+
        `<td style="font-size:.8rem">${s.unit||'—'}</td>`+
        `<td style="font-size:.8rem">${f?.phone||s.phone||'—'}</td>`+
        `<td>${badge}</td>`+
        `<td><button class="bsm bsm-navy" onclick="APP.openHRFileModal('${id}')">🗂 Open</button></td></tr>`;
    }).join('');
  }
  async openHRFileModal(id){
    const st=this.staff[id];if(!st)return;
    $('hf-id').value=id;
    $('hf-staff-name').textContent=st.name+' ('+id+')';
    $('hf-msg').textContent='Loading…';
    $('hrfile-modal').classList.add('open');
    $('hf-empid').value=id;
    $('hf-dept').value=st.unit||'';
    $('hf-position').value=st.position||'';
    $('hf-emptype').value=st.employmentType||'Staff';
    $('hf-joined').value=st.dateJoined?String(st.dateJoined).slice(0,10):'';
    const mgrs=Object.entries(this.staff)
      .filter(([i2,s2])=>(s2.role==='manager'||s2.role==='country_leader')&&i2!==id)
      .sort((a,b)=>a[1].name.localeCompare(b[1].name));
    $('hf-manager').innerHTML='<option value="">— Not set —</option>'
      +mgrs.map(([i2,s2])=>`<option value="${i2}" ${st.supervisor===i2?'selected':''}>${s2.name}</option>`).join('');
    const f=await API.getHRFile(id)||{};
    $('hf-probst').value=f.probation_status||'N/A';
    $('hf-probend').value=f.probation_end?String(f.probation_end).slice(0,10):'';
    $('hf-confirmed').value=f.confirmed_date?String(f.confirmed_date).slice(0,10):'';
    $('hf-dob').value=f.dob?String(f.dob).slice(0,10):'';
    $('hf-phone').value=f.phone||st.phone||'';
    $('hf-address').value=f.residential_address||'';
    $('hf-emergency').value=f.emergency_contact||st.emergencyContact||'';
    $('hf-nok').value=f.next_of_kin||'';
    $('hf-nok-phone').value=f.next_of_kin_phone||'';
    $('hf-ssnit').value=f.ssnit_number||'';
    $('hf-quals').value=f.qualifications||'';
    $('hf-photo-url').value=f.photo_url||'';
    const _pv=$('hf-photo-prev'),_ph=$('hf-photo-ph');
    const _t=this._drivePhoto(f.photo_url);
    if(_pv){if(_t){_pv.src=_t;_pv.style.display='block';if(_ph)_ph.style.display='none';}
            else{_pv.style.display='none';if(_ph)_ph.style.display='block';}}
    const _sub=$('hf-staff-sub');
    if(_sub)_sub.textContent=[st.position||'',st.unit||'',st.employmentType||''].filter(Boolean).join(' · ');
    let docs=[];try{docs=JSON.parse(f.documents||'[]');}catch(e){}
    $('hf-docs-json').value=JSON.stringify(docs);
    this._renderHRDocs(docs);
    $('hf-msg').textContent='';
  }
  _renderHRDocs(docs){
    $('hf-docs').innerHTML=docs.length
      ? docs.map((d,i)=>`<span class="hr-doc-chip">📎 <a href="${d.url}" target="_blank">${d.name}</a> <a href="#" onclick="APP.removeHRDoc(${i});return false" style="color:var(--red)">✕</a></span>`).join('')
      : '<span style="font-size:.76rem;color:var(--text3)">No documents uploaded yet.</span>';
  }
  removeHRDoc(i){
    let docs=[];try{docs=JSON.parse($('hf-docs-json').value||'[]');}catch(e){}
    const d=docs[i];if(!d)return;
    if(!confirm('Remove this document from the staff file?\n\n'+d.name))return;
    const typed=prompt('This cannot be undone.\n\nType DELETE to confirm removal of:\n'+d.name);
    if(typed===null)return;
    if(String(typed).trim().toUpperCase()!=='DELETE')return toast('Not deleted — confirmation did not match','info');
    this.audit('Staff document removed','Document',this.staff[$('hf-id').value]?.name||'',d.name);
    docs.splice(i,1);
    $('hf-docs-json').value=JSON.stringify(docs);
    this._renderHRDocs(docs);
    toast('Removed — click Save File to confirm.','info');
  }
  async uploadHRDoc(){
    const fileInput=$('hf-doc-file');const id=$('hf-id').value;
    if(!fileInput?.files?.length)return toast('Choose a file first','err');
    const file=fileInput.files[0];
    if(file.size>5*1024*1024)return toast('File too large — maximum 5 MB. Scan at lower resolution or split the document.','err');
    if(this._hfBusy)return toast('An upload is already in progress — please wait','info');
    this._hfBusy=true;
    $('hf-msg').innerHTML='<span style="color:var(--teal)">⏳ Uploading '+(file.size/1048576).toFixed(1)+' MB — large scans take 30–60 seconds. Please do not close this window.</span>';
    try{
      const b64=await this._fileToBase64(file);
      const r=await API.gasPost({action:'uploadHRDoc',staffId:id,fileName:file.name,fileData:b64,mimeType:file.type});
      if(r&&r.success&&r.fileUrl){
        let docs=[];try{docs=JSON.parse($('hf-docs-json').value||'[]');}catch(e){}
        docs.push({name:file.name,url:r.fileUrl,at:new Date().toISOString().slice(0,10)});
        $('hf-docs-json').value=JSON.stringify(docs);
        this._renderHRDocs(docs);
        fileInput.value='';
        $('hf-msg').innerHTML='<span style="color:var(--green)">✓ Uploaded — click Save File to confirm.</span>';
      }else $('hf-msg').innerHTML='<span style="color:var(--red)">Upload failed'+((r&&r.error)?': '+r.error:' — no response from Drive. Check your connection and try again.')+'</span>';
    }catch(e){$('hf-msg').innerHTML='<span style="color:var(--red)">Upload error: '+(e.message||e)+'</span>';}
    finally{this._hfBusy=false;}
  }
  async saveHRFile(){
    const id=$('hf-id').value;if(!id)return;
    $('hf-msg').innerHTML='<span style="color:var(--teal)">⏳ Saving…</span>';
    const empUpd={unit:$('hf-dept').value.trim(),position:$('hf-position').value.trim(),
      employment_type:$('hf-emptype').value,date_joined:$('hf-joined').value||null,
      supervisor:$('hf-manager').value,phone:$('hf-phone').value.trim()};
    const e1=await API._update('staff','id=eq.'+encodeURIComponent(id),empUpd);
    const data={
      dob:$('hf-dob').value||null,
      phone:$('hf-phone').value.trim(),
      residential_address:$('hf-address').value.trim(),
      emergency_contact:$('hf-emergency').value.trim(),
      next_of_kin:$('hf-nok').value.trim(),
      next_of_kin_phone:$('hf-nok-phone').value.trim(),
      ssnit_number:$('hf-ssnit').value.trim(),
      qualifications:$('hf-quals').value.trim(),
      probation_status:$('hf-probst').value,
      probation_end:$('hf-probend').value||null,
      confirmed_date:$('hf-confirmed').value||null,
      photo_url:$('hf-photo-url').value||'',
      documents:$('hf-docs-json').value||'[]'
    };
    const r=await API.saveHRFile(id,data);
    if(r&&r.success&&e1!==null){
      const st=this.staff[id];
      if(st){st.unit=empUpd.unit;st.position=empUpd.position;st.employmentType=empUpd.employment_type;
        st.dateJoined=empUpd.date_joined;st.supervisor=empUpd.supervisor;st.phone=empUpd.phone;}
      this._cacheS();
      this.audit('Staff file updated','HR',this.staff[id]?.name||id,
        $('hf-position').value.trim()+' · '+$('hf-emptype').value+' · probation '+$('hf-probst').value);
      closeModal('hrfile-modal');
      this.renderHRFiles('a-');this.renderHRFiles('m-');
      toast('Staff file saved ✓');
    } else $('hf-msg').innerHTML='<span style="color:var(--red)">Save failed: '+(API.lastError||'unknown')+'</span>';
  }

  /* ── Privileges (admin-assigned access) ── */
  _privDefaults(){return{hr:[COUNTRY_LEADER_ID,HR_MANAGER_ID],cases:[HR_MANAGER_ID,COUNTRY_LEADER_ID],payroll:['THPG/05/2025','THPG/01/2026-3',COUNTRY_LEADER_ID]};}
  async _fetchPriv(){
    let p={};
    try{const r=await API._get('settings','key=eq.privileges');if(r&&r.length&&r[0].value)p=JSON.parse(r[0].value);}catch(e){}
    const d=this._privDefaults();
    return{hr:p.hr||d.hr,cases:p.cases||d.cases,payroll:p.payroll||d.payroll};
  }
  async _applyPrivileges(id){
    const p=await this._fetchPriv();this._priv=p;
    const scope=this.user?.role==='admin'?'#sb-admin'
      :(this.user?.role==='manager'||this.user?.role==='country_leader')?'#sb-mgr':'#sb-staff';

    // Re-assert the hidden state first. Anything not explicitly granted
    // below stays hidden, whatever earlier code may have done.
    const GATED={'hr-tab':'hr','contract-tab':'hr','cases-tab':'cases','payroll-tab':'payroll'};
    const GATED_IDS={
      'hr-tab':['nav-mgr-hrfiles'],'contract-tab':['nav-mgr-contract'],
      'cases-tab':['nav-mgr-cases'],'payroll-tab':['nav-mgr-payroll','nav-st-payroll']
    };
    // Re-hide by walking every nav item / mobile tab whose panel is gated
    const GATED_PANELS={
      payroll:['m-payroll','m-grants','st-payroll'],
      cases:['m-cases'],
      hr:['m-hrfiles','m-birthdays','m-announce','m-filemgr','m-performance','m-recruit',
          'm-training','m-orgchart','m-assistant','m-audit','m-onboard','m-assets'],
      contract:['m-contracts']
    };
    const clsFor=k=>k==='payroll'?'payroll-tab':k==='cases'?'cases-tab':k==='contract'?'contract-tab':'hr-tab';
    Object.entries(GATED_PANELS).forEach(([k,panels])=>{
      const cls=clsFor(k);
      panels.forEach(pid=>{
        document.querySelectorAll("[onclick*=\"'"+pid+"'\"]").forEach(el=>{
          if(el.classList.contains('nav-item')||el.classList.contains('mob-tab'))el.classList.add(cls);
        });
      });
    });

    const show=(cls)=>{
      document.querySelectorAll(scope+' .'+cls).forEach(e=>e.classList.remove(cls));
      document.querySelectorAll('.mob-tab.'+cls).forEach(e=>e.classList.remove(cls));
    };
    if(p.hr.includes(id)){
      show('contract-tab');show('hr-tab');
      document.body.classList.add('hr-mode');
      this.renderHRDash('m-');
    } else {
      document.body.classList.remove('hr-mode');
    }
    if(p.cases.includes(id))show('cases-tab');
    if(p.payroll.includes(id))show('payroll-tab');

    _restoreNavGroups();_hideEmptyNavGroups();
    setTimeout(()=>{_restoreNavGroups();_hideEmptyNavGroups();},200);
    this.refreshActionBadges();
    if(document.getElementById('m-slip-body'))this.renderMyPayslips('m-');
    if(document.getElementById('m-torate-body')&&this.user?.role!=='staff')this.renderToRate();
    setTimeout(()=>this.showFirstLoginHint(),900);
  }
  async renderPrivileges(){
    const body=$('a-priv-body');if(!body)return;
    body.innerHTML='<tr><td colspan="4" style="color:var(--text3)">Loading…</td></tr>';
    const p=await this._fetchPriv();
    const list=Object.entries(this.staff).filter(([i,s])=>s.role!=='admin').sort((a,b)=>a[1].name.localeCompare(b[1].name));
    body.innerHTML=list.map(([i,s])=>`<tr><td><strong>${s.name}</strong><br><span style="font-size:.72rem;color:var(--text3)">${i} · ${s.role}</span></td>
      <td style="text-align:center"><input type="checkbox" class="pv-hr" value="${i}" ${p.hr.includes(i)?'checked':''}></td>
      <td style="text-align:center"><input type="checkbox" class="pv-cases" value="${i}" ${p.cases.includes(i)?'checked':''}></td>
      <td style="text-align:center"><input type="checkbox" class="pv-pay" value="${i}" ${p.payroll.includes(i)?'checked':''}></td></tr>`).join('');
  }
  async savePrivileges(){
    const grab=c=>[...document.querySelectorAll('.'+c+':checked')].map(e=>e.value);
    const p={hr:grab('pv-hr'),cases:grab('pv-cases'),payroll:grab('pv-pay')};
    const r=await API._upsert('settings',[{key:'privileges',value:JSON.stringify(p)}]);
    if(r){this.audit('Privileges changed','Security','',JSON.stringify(p));toast('Privileges saved ✓ — takes effect at each person\'s next login');}
    else toast('Save failed — '+(API.lastError||'reason unknown'),'err');
  }

  /* ── Generic branded table print ── */
  printSection(panelId,title){
    const tbl=document.querySelector('#'+panelId+' .scr table');
    if(!tbl)return toast('Nothing to print','err');
    const w=window.open('','_blank');
    w.document.write(`<html><head><title>${title} — THP-Ghana</title><style>
      body{font-family:'Segoe UI',Arial,sans-serif;color:#1e293b;font-size:11px;padding:10mm}
      h1{font-size:15px;color:#2D3592;border-bottom:3px solid #2D3592;padding-bottom:8px;margin:0 0 4px}
      .meta{font-size:9px;color:#64748b;margin-bottom:10px}
      table{width:100%;border-collapse:collapse;font-size:10px}
      th{background:#2D3592;color:#fff;padding:6px;text-align:left;font-size:9.5px}
      td{padding:5px 6px;border:1px solid #e2e8f0;vertical-align:top}
      tr:nth-child(even) td{background:#f8fafc}
      table button{display:none}
      .c-flag{padding:1px 6px;border-radius:8px;font-size:9px;border:1px solid #cbd5e1}
      .hr-doc-chip{font-size:9px}
      @media print{.no-print{display:none}}
    </style></head><body>
    <h1>The Hunger Project — Ghana · ${title}</h1>
    <div class="meta">Generated ${new Date().toLocaleString('en-GB')} · by ${this.user.name}</div>
    ${tbl.outerHTML}
    <div class="no-print" style="margin-top:14px"><button onclick="window.print()" style="padding:8px 20px;background:#2D3592;color:#fff;border:none;border-radius:6px;cursor:pointer;font-weight:600">🖨 Print</button>
    <button onclick="window.close()" style="padding:8px 20px;background:#ef4444;color:#fff;border:none;border-radius:6px;cursor:pointer;font-weight:600;margin-left:8px">✕ Close</button></div>
    </body></html>`);
    w.document.close();
  }

  /* ═══════════════════════════════════════════
     HR PHASE 2 — standalone modules
  ═══════════════════════════════════════════ */
  _uid(p){return p+Date.now().toString(36)+Math.random().toString(36).slice(2,6);}
  _popStaffSel(elId,val){const el=$(elId);if(!el)return;el.innerHTML=Object.entries(this.staff).filter(([i,s])=>s.role!=='admin').sort((a,b)=>a[1].name.localeCompare(b[1].name)).map(([i,s])=>`<option value="${i}"${i===val?' selected':''}>${s.name} (${i})</option>`).join('');}
  _sName(id){return this.staff[id]?.name||id;}

  /* ── HR Dashboard analytics ── */
  _drivePhoto(url){if(!url)return'';const m=String(url).match(/[-\w]{25,}/);return m?'https://drive.google.com/thumbnail?id='+m[0]+'&sz=w200':'';}
  _empType(u){u=(u||'').toLowerCase();return u.includes('national service')?'National Service':u.includes('intern')?'Intern':'Full Staff';}
  _donut(items,total,centreLabel){
    if(!items.length||!total)return'<div style="color:var(--text3);font-size:.78rem">No data</div>';
    let acc=0;
    const stops=items.map(it=>{const a=acc;acc+=it.n/total*100;return `${it.color} ${a.toFixed(2)}% ${acc.toFixed(2)}%`;}).join(',');
    const legend=items.map(it=>`<div style="display:flex;align-items:center;gap:.45rem;margin-bottom:.38rem;font-size:.76rem">
      <span style="width:10px;height:10px;border-radius:3px;background:${it.color};flex:0 0 auto"></span>
      <span style="flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${it.label}</span>
      <span style="color:var(--text3);font-variant-numeric:tabular-nums">${it.n} · ${Math.round(it.n/total*100)}%</span></div>`).join('');
    return `<div style="display:flex;gap:1.2rem;align-items:center;flex-wrap:wrap">
      <div style="width:148px;height:148px;border-radius:50%;background:conic-gradient(${stops});flex:0 0 auto;position:relative;box-shadow:0 2px 10px rgba(0,0,0,.08)">
        <div style="position:absolute;inset:27%;background:var(--surf);border-radius:50%;display:flex;flex-direction:column;align-items:center;justify-content:center">
          <div style="font-size:1.35rem;font-weight:700;line-height:1">${total}</div>
          <div style="font-size:.58rem;color:var(--text3);letter-spacing:.5px;text-transform:uppercase">${centreLabel||''}</div>
        </div>
      </div>
      <div style="flex:1;min-width:160px">${legend}</div>
    </div>`;
  }
  async renderHRDash(p){
    const s1=$(p+'hd-strip1');if(!s1)return;
    s1.innerHTML='<div style="color:var(--text3);font-size:.8rem;padding:.4rem">Loading…</div>';
    const files=await API._get('hr_staff_files','select=staff_id,dob,photo_url')||[];
    const fm={};files.forEach(f=>fm[f.staff_id]=f);
    const list=Object.entries(this.staff).filter(([i,s])=>s.role!=='admin');
    const now=new Date();now.setHours(0,0,0,0);
    const box=(n,l,c,ic)=>`<div class="cs-box">${ic?`<div style="font-size:1.05rem;line-height:1">${ic}</div>`:''}<div class="cs-num"${c?` style="color:${c}"`:''}>${n}</div><div class="cs-lbl">${l}</div></div>`;
    // Live attendance stats (org-wide) — merged into one container
    const todayStr=fmtD(new Date().toISOString());
    const todayRecs=this.records.filter(r=>fmtD(r.date||r.in)===todayStr);
    const presToday=new Set(todayRecs.map(r=>r.id)).size;
    const activeNow=new Set(todayRecs.filter(r=>r.status==='Active').map(r=>r.id)).size;
    const todayISO2=new Date().toISOString().slice(0,10);
    const onLvToday=list.filter(([i])=>leaveOnDate(this.leave,i,todayISO2)).length;
    const pendLv=this.leave.filter(l=>l.status==='Pending').length;
    // Headcount
    const male=list.filter(([i,s])=>(s.gender||'male')==='male').length;
    const female=list.filter(([i,s])=>s.gender==='female').length;
    const d90=new Date(now-90*86400000);
    const newest=list.filter(([i,s])=>s.contractStart&&new Date(s.contractStart)>=d90).length;
    const attn=list.filter(([i,s])=>this._contractFlag(s.contractEnd).cls==='red').length;
    s1.innerHTML=box(list.length,'Total Staff','','👥')+box(female,'Female','#ec4899','👩')+box(male,'Male','#3b82f6','👨')
      +box(presToday,'Present Today','#16a34a','✅')+box(activeNow,'Active Now','#f59e0b','🕒')+box(onLvToday,'On Leave','#6366f1','🌴')+box(pendLv,'Pending Leave','#818cf8','⏳')
      +box(newest,'New (90 days)','#16a34a','🆕')+box(attn,'Contract Alerts','#dc2626','📄');
    // Ages + tenure
    const ages=list.map(([i])=>fm[i]?.dob).filter(Boolean).map(d=>Math.floor((now-new Date(String(d).slice(0,10)))/(365.25*86400000))).filter(a=>a>0&&a<100);
    const tenures=list.map(([i,s])=>s.contractStart).filter(Boolean).map(d=>(now-new Date(d))/(365.25*86400000)).filter(t=>t>=0);
    const s2=$(p+'hd-strip2');
    if(s2)s2.innerHTML=ages.length
      ? box(Math.min(...ages),'Youngest','','🧒')+box(Math.max(...ages),'Oldest','','🧓')+box((ages.reduce((a,b)=>a+b,0)/ages.length).toFixed(1),'Average Age','','📊')+box(ages.length+'/'+list.length,'DOBs on File','','🗂')+box(tenures.length?(tenures.reduce((a,b)=>a+b,0)/tenures.length).toFixed(1)+' yrs':'—','Avg Tenure','','⏱')
      : '<div style="color:var(--text3);font-size:.78rem;padding:.4rem">Add dates of birth in Staff Files to see age analytics.</div>';
    // Birthdays this month
    const bd=$(p+'hd-bdays');
    if(bd){
      const month=now.getMonth();
      const cel=list.map(([i,s])=>({i,s,f:fm[i]})).filter(x=>x.f?.dob&&new Date(String(x.f.dob).slice(0,10)).getMonth()===month)
        .map(x=>{const d=new Date(String(x.f.dob).slice(0,10));return{...x,d,day:d.getDate(),today:d.getDate()===now.getDate(),turns:now.getFullYear()-d.getFullYear()};})
        .sort((a,b)=>a.day-b.day);
      bd.innerHTML=cel.length?cel.map(x=>{
        const ph=this._drivePhoto(x.f.photo_url);
        const av=ph?`<img src="${ph}" style="width:56px;height:56px;border-radius:50%;object-fit:cover;border:2px solid ${x.today?'var(--gold,#F5A623)':'var(--border)'}">`
          :`<div style="width:56px;height:56px;border-radius:50%;background:${x.s.color||'#2D3592'};color:#fff;display:flex;align-items:center;justify-content:center;font-weight:700;font-size:1.2rem;border:2px solid ${x.today?'var(--gold,#F5A623)':'transparent'}">${(x.s.name||'?')[0]}</div>`;
        return`<div style="text-align:center;width:110px">${av.replace('style="','style="margin:0 auto;display:block;')}
          <div style="font-size:.76rem;font-weight:600;margin-top:.3rem">${x.s.name.split(' ')[0]} ${x.today?'🎉':''}</div>
          <div style="font-size:.68rem;color:var(--text3)">${x.day} ${x.d.toLocaleString('en',{month:'short'})} · turns ${x.turns}</div></div>`;
      }).join(''):'<div style="color:var(--text3);font-size:.78rem">No birthdays this month (or no DOBs on file).</div>';
    }
    // Ratios
    const uEl=$(p+'hd-units');
    if(uEl){
      const uc={};list.forEach(([i,s])=>{const u=(s.unit||'Unassigned').trim()||'Unassigned';uc[u]=(uc[u]||0)+1;});
      const cols=['#2D3592','#3DBFB8','#F5A623','#22c55e','#ef4444','#a855f7','#06b6d4','#ec4899','#f97316','#818cf8'];
      const items=Object.entries(uc).sort((a,b)=>b[1]-a[1]).map(([u,n],i)=>({label:u,n,color:cols[i%cols.length]}));
      uEl.innerHTML=this._donut(items,list.length,'Staff');
    }
    const eEl=$(p+'hd-emp');
    if(eEl){
      const ec={};list.forEach(([i,s])=>{const t=this._empType(s.unit);ec[t]=(ec[t]||0)+1;});
      const ecol={'Full Staff':'#22c55e','Intern':'#F5A623','National Service':'#3DBFB8'};
      const items=Object.entries(ec).sort((a,b)=>b[1]-a[1]).map(([t,n])=>({label:t,n,color:ecol[t]||'#818cf8'}));
      eEl.innerHTML=this._donut(items,list.length,'Staff')
        +`<div style="font-size:.72rem;color:var(--text3);margin-top:.7rem;text-align:center">Gender ratio — 👩 ${female} : 👨 ${male}</div>`;
    }
  }
  async _bdPrefill(){
    const id=$('m-bd-staff')?.value;if(!id)return;
    const f=await API.getHRFile(id);
    $('m-bd-dob').value=f?.dob?String(f.dob).slice(0,10):'';
  }
  async editBirthday(id){
    const sel=$('m-bd-staff');if(sel)sel.value=id;
    await this._bdPrefill();
    $('m-bd-dob')?.focus();
    $('m-bd-staff')?.scrollIntoView({behavior:'smooth',block:'center'});
  }
  async saveBirthday(){
    const id=$('m-bd-staff')?.value,dob=$('m-bd-dob')?.value;
    if(!id||!dob)return toast('Select staff and date','err');
    const r=await API._upsert('hr_staff_files',[{staff_id:id,dob}]);
    if(r){toast('Date of birth saved ✓');this.renderBirthdays('m-');this.renderHRDash('m-');}
    else toast('Save failed: '+(API.lastError||'unknown error'),'err');
  }
  async uploadHRPhoto(){
    const inp=$('hf-photo-file');const id=$('hf-id').value;
    if(!inp?.files?.length)return toast('Choose an image first','err');
    const file=inp.files[0];
    if(!file.type.startsWith('image/'))return toast('Images only','err');
    if(file.size>2*1024*1024)return toast('Image too large (max 2MB)','err');
    $('hf-msg').innerHTML='<span style="color:var(--teal)">⏳ Uploading photo…</span>';
    try{
      const b64=await this._fileToBase64(file);
      const r=await API.gasPost({action:'uploadHRDoc',staffId:id,fileName:'photo_'+file.name,fileData:b64,mimeType:file.type});
      if(r&&r.success&&r.fileUrl){
        $('hf-photo-url').value=r.fileUrl;
        const pv=$('hf-photo-prev'),ph=$('hf-photo-ph');
        pv.src=this._drivePhoto(r.fileUrl);pv.style.display='block';if(ph)ph.style.display='none';
        inp.value='';
        $('hf-msg').innerHTML='<span style="color:var(--green)">✓ Photo uploaded — click Save File to confirm.</span>';
      }else $('hf-msg').innerHTML='<span style="color:var(--red)">Upload failed.</span>';
    }catch(e){$('hf-msg').innerHTML='<span style="color:var(--red)">Upload error.</span>';}
  }

  /* ── Birthdays ── */
  async renderBirthdays(p){
    this._popStaffSel('m-bd-staff');
    const body=$(p+'bd-body');if(!body)return;
    body.innerHTML='<tr><td colspan="5" style="color:var(--text3)">Loading…</td></tr>';
    const files=await API._get('hr_staff_files','select=staff_id,dob')||[];
    const now=new Date();now.setHours(0,0,0,0);
    const rows=files.filter(f=>f.dob&&this.staff[f.staff_id]&&this.staff[f.staff_id].role!=='admin').map(f=>{
      const d=new Date(String(f.dob).slice(0,10));
      let next=new Date(now.getFullYear(),d.getMonth(),d.getDate());
      if(next<now)next=new Date(now.getFullYear()+1,d.getMonth(),d.getDate());
      const days=Math.round((next-now)/86400000);
      return{id:f.staff_id,dob:d,next,days,age:next.getFullYear()-d.getFullYear()};
    }).sort((a,b)=>a.days-b.days);
    if(!rows.length){body.innerHTML='<tr><td colspan="5"><div class="empty"><div class="empty-ico">🎂</div>No dates of birth on file yet</div></td></tr>';return;}
    body.innerHTML=rows.map(r=>{
      const s=this.staff[r.id];
      const when=r.days===0?'<span class="c-flag green">🎉 Today!</span>':r.days<=14?`<span class="c-flag amber">${r.days}d</span>`:`<span class="c-flag none">${r.days}d</span>`;
      return`<tr><td><strong>${s.name}</strong><br><span style="font-size:.72rem;color:var(--text3)">${r.id}</span></td><td style="font-size:.8rem">${s.unit||'—'}</td><td>${r.dob.getDate()} ${r.dob.toLocaleString('en',{month:'short'})}</td><td>${when}</td><td><button class="bsm bsm-navy" onclick="APP.editBirthday('${r.id}')">✏ Edit</button></td></tr>`;
    }).join('');
  }

  /* ── Announcements ── */
  async renderAnnouncements(p){
    const list=$(p+'ann-list');if(!list)return;
    list.innerHTML='<div style="color:var(--text3);font-size:.8rem">Loading…</div>';
    const rows=await API._get('announcements','order=created_at.desc&limit=50')||[];
    this._anns=rows;
    if(!rows.length){list.innerHTML='<div class="empty"><div class="empty-ico">📣</div>No announcements yet</div>';return;}
    list.innerHTML=rows.map(a=>`<div class="ann-card"><h5>${a.title}</h5><div style="font-size:.8rem;white-space:pre-wrap">${a.body||''}</div><div class="ann-meta">${a.author||''} · ${String(a.created_at).slice(0,10)} <a href="#" onclick="APP.delAnnouncement('${a.id}','${p}');return false" style="color:var(--red);margin-left:.5rem">✕ delete</a></div></div>`).join('');
  }
  async postAnnouncement(p){
    const t=$(p+'ann-title').value.trim(),b=$(p+'ann-body').value.trim();
    if(!t)return toast('Title required','err');
    const r=await API._upsert('announcements',[{id:this._uid('ANN'),title:t,body:b,author:this.user.name,created_at:new Date().toISOString()}]);
    if(r){
      this.audit('Announcement posted','HR',t,'');
      if($(p+'ann-email')?.checked){
        const recips=Object.entries(this.staff).filter(([i,s])=>s.role!=='admin'&&s.email)
          .map(([i,s])=>({name:s.name,email:s.email}));
        API.gasPost({action:'announceEmail',title:t,body:b,author:this.user.name,recipients:recips})
          .then(res=>{if(res&&res.success)toast('Emailed to '+res.sent+' staff ✓');})
          .catch(()=>{});
      }
      $(p+'ann-title').value='';$(p+'ann-body').value='';toast('Announcement posted ✓');this.renderAnnouncements(p);}
    else toast('Post failed: '+(API.lastError||'unknown error'),'err');
  }
  async delAnnouncement(id,p){
    const a=(this._anns||[]).find(x=>x.id===id);
    const t1=a?.title||'this announcement';
    if(!confirm('Delete this announcement?\n\n'+t1))return;
    const t=prompt('This cannot be undone.\n\nType DELETE to confirm removal of:\n'+t1);
    if(t===null)return;
    if(String(t).trim().toUpperCase()!=='DELETE')return toast('Not deleted — confirmation did not match','info');
    await API._delete('announcements','id=eq.'+encodeURIComponent(id));
    this.audit('Announcement deleted','HR',t1,'');
    toast('Announcement deleted');
    this.renderAnnouncements(p);
  }

  /* ── File Manager — organizational document library ── */
  _visLabel(v){return{staff:'All Staff',managers:'Managers Only',cl:'Country Leader Only',hr:'HR Only'}[v]||'All Staff';}
  async renderFileMgr(p){
    const body=$('m-fm-body');if(!body)return;
    body.innerHTML='<tr><td colspan="6" style="color:var(--text3)">Loading…</td></tr>';
    const rows=await API._get('org_documents','order=created_at.desc&limit=300')||[];
    this._orgDocs=rows;
    if(!rows.length){body.innerHTML='<tr><td colspan="6"><div class="empty"><div class="empty-ico">📁</div>No documents in the library yet — upload above</div></td></tr>';return;}
    const isHR=this.user.id===HR_MANAGER_ID||this.user.role==='admin';
    body.innerHTML=rows.map(d=>{
      const acc=d.access_level||'download';
      const visCell=isHR
        ? `<select class="fi" style="width:150px;font-size:.74rem" onchange="APP.setOrgDocField('${d.id}','visibility',this.value)">
            ${['staff','managers','cl','hr'].map(v=>`<option value="${v}" ${d.visibility===v?'selected':''}>${this._visLabel(v)}</option>`).join('')}</select>`
        : `<span class="c-flag none">${this._visLabel(d.visibility)}</span>`;
      const accCell=isHR
        ? `<select class="fi" style="width:130px;font-size:.74rem" onchange="APP.setOrgDocField('${d.id}','access_level',this.value)">
            <option value="download" ${acc==='download'?'selected':''}>View &amp; Download</option>
            <option value="view" ${acc==='view'?'selected':''}>View Only</option></select>`
        : `<span class="c-flag ${acc==='view'?'amber':'green'}">${acc==='view'?'View only':'Download'}</span>`;
      return `<tr><td><a href="${d.url}" target="_blank" style="color:var(--teal);font-weight:600">📎 ${d.name}</a></td>
        <td style="font-size:.8rem">${d.category||'General'}</td>
        <td>${visCell}</td><td>${accCell}</td>
        <td style="font-size:.76rem">${String(d.created_at).slice(0,10)}<br><span style="color:var(--text3);font-size:.68rem">${d.uploaded_by||''}</span></td>
        <td>${isHR?`<button class="bsm" style="background:rgba(239,68,68,.12);color:var(--red)" onclick="APP.delOrgDoc('${d.id}')">🗑 Delete</button>`:''}</td></tr>`;
    }).join('');
  }
  async setOrgDocField(id,field,val){
    if(this.user.id!==HR_MANAGER_ID&&this.user.role!=='admin')return toast('Only HR can change access','err');
    const r=await API._update('org_documents','id=eq.'+encodeURIComponent(id),{[field]:val});
    if(r!==null)toast('Access updated ✓');else toast('Update failed: '+(API.lastError||''),'err');
  }
  async uploadOrgDoc(){
    if(this._odBusy)return toast('An upload is already in progress…','info');
    const inp=$('m-od-file');const msg=$('m-od-msg');
    if(!inp?.files?.length)return toast('Choose a file first','err');
    const file=inp.files[0];
    if(file.size>5*1024*1024)return toast('File too large (max 5MB)','err');
    this._odBusy=true;
    const mb=(file.size/1048576).toFixed(1);
    msg.innerHTML='<span style="color:var(--teal)">⏳ Uploading '+mb+' MB to Drive — larger files can take 20-60 seconds…</span>';
    try{
      const b64=await this._fileToBase64(file);
      const r=await API.gasPost({action:'uploadHRDoc',staffId:'ORG',fileName:file.name,fileData:b64,mimeType:file.type});
      if(!r)         {msg.innerHTML='<span style="color:var(--red)">No response from Google Drive. Check the Apps Script deployment.</span>';return;}
      if(!r.success) {msg.innerHTML='<span style="color:var(--red)">Drive upload failed: '+(r.error||'unknown')+'</span>';return;}
      msg.innerHTML='<span style="color:var(--teal)">⏳ Saving to library…</span>';
      const ins=await API._upsert('org_documents',[{id:this._uid('OD'),name:file.name,url:r.fileUrl,category:$('m-od-cat').value,visibility:$('m-od-vis').value,access_level:$('m-od-acc').value,uploaded_by:this.user.name,created_at:new Date().toISOString()}]);
      if(!ins){msg.innerHTML='<span style="color:var(--red)">File reached Drive but the library record failed: '+(API.lastError||'unknown')+'</span>';return;}
      inp.value='';
      this.audit('Document uploaded','Document',file.name,$('m-od-vis').value+' / '+$('m-od-acc').value);
      // Optionally tell the people who can now see it
      if($('m-od-notify')?.checked){
        const vis=$('m-od-vis').value;
        const recips=Object.entries(this.staff).filter(([i,st])=>{
          if((st.role||'')==='admin'||!st.email)return false;
          if(vis==='staff')return true;
          if(vis==='managers')return st.role==='manager'||st.role==='country_leader';
          if(vis==='cl')return i===COUNTRY_LEADER_ID;
          if(vis==='hr')return i===HR_MANAGER_ID||i===COUNTRY_LEADER_ID;
          return false;
        }).map(([i,st])=>({name:st.name,email:st.email,phone:st.phone||''}));
        if(recips.length){
          API.gasPost({action:'docNotify',docName:file.name,category:$('m-od-cat').value,
            access:$('m-od-acc').value,by:this.user.name,recipients:recips})
            .then(r=>{if(r&&r.success)toast('Notified '+r.sent+' staff ✓');}).catch(()=>{});
        }
        $('m-od-notify').checked=false;
      }
      msg.innerHTML='<span style="color:var(--green)">✓ Uploaded and listed in the library.</span>';
      await this.renderFileMgr('m-');
    }catch(e){msg.innerHTML='<span style="color:var(--red)">Upload error: '+(e.message||e)+'</span>';}
    finally{this._odBusy=false;}
  }
  async setOrgDocVis(id,vis){
    await API._update('org_documents','id=eq.'+encodeURIComponent(id),{visibility:vis});
    toast(vis==='staff'?'Now visible to all staff':'Now HR only');
  }
  async delOrgDoc(id){
    const doc=(this._orgDocs||[]).find(d=>d.id===id);
    if(!confirm('Delete this document from the library?\n\n'+(doc?.name||'')))return;
    const t=prompt('This cannot be undone.\n\nType DELETE to confirm removal of:\n'+(doc?.name||''));
    if(t===null)return;
    if(String(t).trim().toUpperCase()!=='DELETE')return toast('Not deleted — confirmation did not match','info');
    await API._delete('org_documents','id=eq.'+encodeURIComponent(id));
    this.audit('Library document deleted','Document',doc?.name||id,doc?.category||'');
    this.renderFileMgr('m-');
  }
  _visibleDocFilter(){
    const uid=this.user.id,role=this.user.role;
    const allowed=['staff'];
    if(role==='manager'||role==='country_leader')allowed.push('managers');
    if(uid===COUNTRY_LEADER_ID)allowed.push('cl');
    if(uid===HR_MANAGER_ID||role==='admin')allowed.push('managers','cl','hr');
    return [...new Set(allowed)];
  }
  _docRow(d){
    const acc=d.access_level||'download';
    const btn=acc==='view'
      ? `<a href="${d.url}" target="_blank" class="bsm" style="text-decoration:none;background:var(--surf2);color:var(--text2);border:1px solid var(--border)">👁 View</a>`
      : `<a href="${d.url}" target="_blank" class="bsm bsm-navy" style="text-decoration:none">⬇ Download</a>`;
    return `<tr><td style="font-weight:600">📎 ${d.name}</td><td style="font-size:.8rem">${d.category||'General'}</td><td style="font-size:.78rem">${String(d.created_at).slice(0,10)}</td><td>${btn}</td></tr>`;
  }
  async renderStaffDocs(){
    const body=$('st-docs-body');if(!body)return;
    body.innerHTML='<tr><td colspan="4" style="color:var(--text3)">Loading…</td></tr>';
    const vis=this._visibleDocFilter().map(v=>'"'+v+'"').join(',');
    const rows=await API._get('org_documents','visibility=in.('+vis+')&order=created_at.desc&limit=300')||[];
    body.innerHTML=rows.length?rows.map(d=>this._docRow(d)).join('')
      :'<tr><td colspan="4"><div class="empty"><div class="empty-ico">📁</div>No documents shared yet</div></td></tr>';
  }
  /* ── New-item tracking (announcements & documents) ── */
  _seenKey(kind){return 'thp_seen_'+kind+'_'+(this.user?.id||'x');}
  _lastSeen(kind){try{return localStorage.getItem(this._seenKey(kind))||'1970-01-01';}catch(e){return '1970-01-01';}}
  markSeen(kind){
    try{localStorage.setItem(this._seenKey(kind),new Date().toISOString());}catch(e){}
    this._setBadge(kind,0);
  }
  _setBadge(kind,n){
    const el=$(kind==='ann'?'badge-ann':'badge-docs');
    if(el){el.textContent=n>0?(n>9?'9+':n):'';el.classList.toggle('on',n>0);}
    if(kind==='doc'){const m=$('mob-badge-docs');if(m)m.style.display=n>0?'block':'none';}
    if(kind==='slip'){const m2=$('badge-slip-m');if(m2){m2.textContent=n>0?n:'';m2.classList.toggle('on',n>0);}}
  }
  /* ── Group + item notification counts ── */
  _setGrpBadge(id,n){
    const el=$(id);if(!el)return;
    el.textContent=n>0?(n>9?'9+':n):'';
    el.classList.toggle('on',n>0);
  }
  /* ── First-login guidance: point new users at the Dashboard ── */
  showFirstLoginHint(){
    try{
      const k='thp_hint_'+(this.user?.id||'');
      if(localStorage.getItem(k))return;
      const nav=document.querySelector('#sb-mgr .nav-item, #sb-staff .nav-item');
      if(!nav)return;
      const r=nav.getBoundingClientRect();
      const d=document.createElement('div');
      d.className='nav-hint';
      d.style.left=(r.right+12)+'px';
      d.style.top=(r.top-4)+'px';
      d.innerHTML='<strong>Start here.</strong><br>Your Dashboard shows today at a glance. '
        +'A red dot means something needs your attention.<br>'
        +'<button onclick="this.closest(\'.nav-hint\').remove()">Got it</button>';
      document.body.appendChild(d);
      localStorage.setItem(k,'1');
      setTimeout(()=>{if(d.parentNode)d.remove();},15000);
    }catch(e){}
  }
  async refreshActionBadges(){
    try{
      const uid=this.user?.id;if(!uid)return;
      // Team: leave requests waiting on me
      const isFinal=uid===COUNTRY_LEADER_ID||this._isActiveDelegate(uid);
      const pending=this.leave.filter(l=>l.status==='Pending'&&(
        (l.supervisorId===uid&&l.supervisorStatus==='Pending')||
        (isFinal&&(l.supervisorStatus==='Approved'||l.supervisorStatus==='N/A')&&l.finalApproverStatus==='Pending')
      )).length;
      this._setGrpBadge('badge-team',pending);
      // HR: contracts needing attention + incomplete staff files

      // Finance — each item carries its own count, and the group shows the total
      let payN=0,grantN=0;
      if(this.PAY_APPROVERS.includes(uid)||this.PAY_PREPARERS.includes(uid)){
        const [runs,gr]=await Promise.all([
          API.secureGet('payroll_runs','select=status&status=eq.Submitted'),
          API._get('grants','select=end_date,status')
        ]);
        payN=(runs||[]).length;
        (gr||[]).forEach(g=>{if(String(g.status)!=='Closed'&&this._grantFlag(g.end_date).cls==='red')grantN++;});
      }
      this._setGrpBadge('badge-payroll',payN);
      this._setGrpBadge('badge-grants',grantN);
      this._setGrpBadge('badge-fin',payN+grantN);

      // HR — per-item counts
      let conN=0,fileN=0;
      const files=await API.getAllHRFiles();
      const fm={};(files||[]).forEach(f=>fm[f.staff_id]=f);
      Object.entries(this.staff).forEach(([i,st])=>{
        if((st.role||'')==='admin')return;
        if(this._contractFlag(st.contractEnd).cls==='red')conN++;
        const f=fm[i];
        const done=f?[f.dob,f.phone,f.next_of_kin,f.ssnit_number].filter(v=>v&&String(v).trim()).length:0;
        if(done<4)fileN++;
      });
      this._setGrpBadge('badge-contracts',conN);
      this._setGrpBadge('badge-files',fileN);
      const ck=await API._get('staff_checklists','select=status&status=neq.Completed')||[];
      this._setGrpBadge('badge-onboard',ck.length);
      this._setGrpBadge('badge-hr',conN+fileN+ck.length);
    }catch(e){}
  }
  async refreshNewBadges(){
    try{
      const annSeen=this._lastSeen('ann'),docSeen=this._lastSeen('doc');
      const anns=await API._get('announcements','select=created_at&order=created_at.desc&limit=50')||[];
      this._setBadge('ann',anns.filter(a=>String(a.created_at)>annSeen).length);
      const vis=this._visibleDocFilter().map(v=>'"'+v+'"').join(',');
      const docs=await API._get('org_documents','select=created_at&visibility=in.('+vis+')&order=created_at.desc&limit=50')||[];
      this._setBadge('doc',docs.filter(d=>String(d.created_at)>docSeen).length);
    }catch(e){}
  }

  /* ── Staff dashboard feed: announcements + documents ── */
  async renderStaffFeed(){
    const al=$('st-ann-list');
    if(al){
      const anns=await API._get('announcements','order=created_at.desc&limit=5')||[];
      const seen=this._lastSeen('ann');
      al.innerHTML=anns.length?anns.map(a=>`<div class="ann-card"><h5>${a.title}${String(a.created_at)>seen?'<span class="new-chip">NEW</span>':''}</h5><div style="font-size:.8rem;white-space:pre-wrap">${a.body||''}</div><div class="ann-meta">${a.author||''} · ${String(a.created_at).slice(0,10)}</div></div>`).join('')
        :'<div style="color:var(--text3);font-size:.8rem">No announcements yet.</div>';
    }
    const fd=$('st-feed-docs');
    if(fd){
      const vis=this._visibleDocFilter().map(v=>'"'+v+'"').join(',');
      const rows=await API._get('org_documents','visibility=in.('+vis+')&order=created_at.desc&limit=10')||[];
      fd.innerHTML=rows.length?rows.map(d=>this._docRow(d)).join('')
        :'<tr><td colspan="4" style="color:var(--text3);font-size:.8rem">No documents shared yet.</td></tr>';
    }    this.refreshNewBadges();
  }
  /* ── Birthday wish popup (once per year, on the staff\'s own birthday) ── */
  async checkBirthdayWish(){
    try{
      const uid=this.user?.id;if(!uid)return;
      const yr=new Date().getFullYear();
      if(localStorage.getItem('thp_bday_'+uid+'_'+yr))return;
      const f=await API.getHRFile(uid);
      if(!f?.dob)return;
      const d=new Date(String(f.dob).slice(0,10)),now=new Date();
      // Show the greeting on the day, or on the next working day if the
      // birthday fell on a weekend or public holiday and they were not in.
      const thisYear=new Date(now.getFullYear(),d.getMonth(),d.getDate());
      const daysSince=Math.floor((now-thisYear)/86400000);
      if(daysSince<0||daysSince>3)return;
      localStorage.setItem('thp_bday_'+uid+'_'+yr,'1');
      this._showBirthdayPopup(this.user.name.split(' ')[0]);
    }catch(e){}
  }
  _showBirthdayPopup(firstName){
    const ov=document.createElement('div');
    ov.className='bday-overlay';
    ov.innerHTML='<div class="bday-card">'
      +'<div class="bday-cake">🎂</div>'
      +'<h2>Happy Birthday, '+firstName+'!</h2>'
      +'<p>Wishing you a wonderful year ahead.<br>From all of us at The Hunger Project — Ghana 🎉</p>'
      +'<button class="btn-add" style="margin:.6rem auto 0" onclick="this.closest(\'.bday-overlay\').remove()">Thank you! 🎈</button>'
      +'</div>';
    const colors=['#F5A623','#3DBFB8','#2D3592','#ef4444','#22c55e','#ec4899','#a855f7'];
    for(let i=0;i<60;i++){
      const c=document.createElement('span');
      c.className='bday-confetti';
      c.style.left=Math.random()*100+'%';
      c.style.background=colors[i%colors.length];
      c.style.animationDelay=(Math.random()*2.2)+'s';
      c.style.animationDuration=(2.4+Math.random()*1.8)+'s';
      ov.appendChild(c);
    }
    document.body.appendChild(ov);
    setTimeout(()=>{if(ov.parentNode)ov.remove();},14000);
  }

  /* ═══════════════════════════════════════════
     AUDIT TRAIL — records key actions
  ═══════════════════════════════════════════ */
  async audit(action,category,target,detail){
    try{
      await API._upsert('audit_log',[{id:this._uid('AU'),actor_id:this.user?.id||'',actor_name:this.user?.name||'System',
        action,category:category||'General',target:target||'',detail:detail||'',created_at:new Date().toISOString()}]);
    }catch(e){}
  }
  async renderAudit(){
    const body=$('au-body');if(!body)return;
    body.innerHTML='<tr><td colspan="6" style="color:var(--text3)">Loading…</td></tr>';
    let rows=await API._get('audit_log','order=created_at.desc&limit=400')||[];
    const cat=$('au-cat')?.value||'',q=($('au-search')?.value||'').trim().toLowerCase();
    if(cat)rows=rows.filter(r=>r.category===cat);
    if(q)rows=rows.filter(r=>[r.actor_name,r.action,r.target,r.detail].join(' ').toLowerCase().includes(q));
    this._auditRows=rows;
    if(!rows.length){body.innerHTML='<tr><td colspan="6"><div class="empty"><div class="empty-ico">🛡</div>No activity recorded yet</div></td></tr>';return;}
    const cc={Auth:'none',Leave:'amber',Staff:'none',HR:'green',Payroll:'amber',Document:'none',Security:'red'};
    body.innerHTML=rows.map(r=>{
      const d=new Date(r.created_at);
      return `<tr class="au-row"><td style="white-space:nowrap">${d.toLocaleDateString('en-GB',{day:'2-digit',month:'short'})}<br><span style="font-size:.68rem;color:var(--text3)">${d.toLocaleTimeString([],{hour:'2-digit',minute:'2-digit'})}</span></td>
        <td><strong>${r.actor_name||'—'}</strong></td>
        <td>${r.action||''}</td>
        <td><span class="c-flag ${cc[r.category]||'none'}">${r.category||'General'}</span></td>
        <td>${r.target||'—'}</td>
        <td style="color:var(--text2)">${r.detail||''}</td></tr>`;
    }).join('');
  }
  exportAudit(){
    const rows=this._auditRows||[];
    if(!rows.length)return toast('Nothing to export','err');
    let csv='When,Who,Action,Category,Affected,Detail\n';
    rows.forEach(r=>{csv+=`"${new Date(r.created_at).toLocaleString('en-GB')}","${r.actor_name||''}","${r.action||''}","${r.category||''}","${r.target||''}","${(r.detail||'').replace(/"/g,'""')}"\n`;});
    this._dl(csv,'THP_Audit_'+Date.now()+'.csv','text/csv');
  }

  /* ═══════════════════════════════════════════
     ASK HR — answers from THP-Ghana data only
  ═══════════════════════════════════════════ */
  initAssistant(){
    if(!this._aiHist){
      this._aiHist=[];
      this._aiSay('bot','Hello '+(this.user.name||'').split(' ')[0]+" — ask me about staff, attendance, leave, contracts, birthdays or a specific person.");
    }
    // Suggestion chips hidden — the assistant still answers these when typed.
    const c=$('ai-chips');
    if(c){c.innerHTML='';c.style.display='none';}
  }
  clearAssistant(){this._aiHist=null;const c=$('ai-chat');if(c)c.innerHTML='';this.initAssistant();}
  _aiSay(who,text){
    const c=$('ai-chat');if(!c)return;
    const d=document.createElement('div');
    d.className='ai-msg '+(who==='you'?'ai-you':'ai-bot');
    d.textContent=text;
    c.appendChild(d);c.scrollTop=c.scrollHeight;
  }
  async askAssistant(preset){
    const inp=$('ai-input');
    const q=(preset||inp?.value||'').trim();
    if(!q)return;
    if(inp&&!preset)inp.value='';
    this._aiSay('you',q);
    const ans=await this._aiAnswer(q.toLowerCase());
    this._aiSay('bot',ans);
  }
  async _aiAnswer(q){
    const list=Object.entries(this.staff).filter(([i,s])=>s.role!=='admin');
    const todayISO=new Date().toISOString().slice(0,10);
    const todayStr=fmtD(new Date().toISOString());
    const has=(...w)=>w.some(x=>q.includes(x));
    // headcount / gender / unit
    if(has('how many staff','headcount','total staff','number of staff')||((has('gender','breakdown','unit','department'))&&!has('leave'))){
      const m=list.filter(([i,s])=>(s.gender||'male')==='male').length;
      const f=list.filter(([i,s])=>s.gender==='female').length;
      const uc={};list.forEach(([i,s])=>{const u=(s.unit||'Unassigned').trim();uc[u]=(uc[u]||0)+1;});
      return `We have ${list.length} staff — ${f} female, ${m} male.\n\nBy unit:\n`+
        Object.entries(uc).sort((a,b)=>b[1]-a[1]).map(([u,n])=>`• ${u}: ${n}`).join('\n');
    }
    // on leave today
    if(has('on leave','who is on leave','leave today')){
      const on=list.filter(([i])=>leaveOnDate(this.leave,i,todayISO));
      if(!on.length)return 'Nobody is on approved leave today.';
      return `${on.length} on leave today:\n`+on.map(([i,s])=>{
        const l=leaveOnDate(this.leave,i,todayISO);
        return `• ${s.name} — ${l.type} (until ${fmtISO(l.endDate)})`;}).join('\n');
    }
    // pending leave
    if(has('pending leave','awaiting approval','leave request')){
      const p=this.leave.filter(l=>l.status==='Pending');
      if(!p.length)return 'There are no pending leave requests.';
      return `${p.length} pending request(s):\n`+p.map(l=>`• ${l.name} — ${l.type}, ${fmtISO(l.startDate)} to ${fmtISO(l.endDate)} (${l.days}d)`).join('\n');
    }
    // contracts
    if(has('contract')){
      const flagged=list.map(([i,s])=>({s,f:this._contractFlag(s.contractEnd)}))
        .filter(x=>x.f.cls==='red'||x.f.cls==='amber')
        .sort((a,b)=>(a.f.days??999)-(b.f.days??999));
      if(!flagged.length)return 'No contracts are expiring within 60 days.';
      return `${flagged.length} contract(s) need attention:\n`+flagged.map(x=>`• ${x.s.name} — ${x.f.label} (ends ${fmtISO(x.s.contractEnd)})`).join('\n');
    }
    // birthdays
    if(has('birthday','birthdays','born')){
      const files=await API._get('hr_staff_files','select=staff_id,dob')||[];
      const mth=new Date().getMonth();
      const cel=files.filter(f=>f.dob&&this.staff[f.staff_id]&&new Date(String(f.dob).slice(0,10)).getMonth()===mth)
        .map(f=>({n:this.staff[f.staff_id].name,d:new Date(String(f.dob).slice(0,10))}))
        .sort((a,b)=>a.d.getDate()-b.d.getDate());
      if(!cel.length)return 'No birthdays on file for this month.';
      return `Birthdays this month:\n`+cel.map(c=>`• ${c.n} — ${c.d.getDate()} ${c.d.toLocaleString('en',{month:'long'})}`).join('\n');
    }
    // not clocked in
    if(has('not clocked','absent','who has not','clock in today','present today')){
      const rec=this.records.filter(r=>fmtD(r.date||r.in)===todayStr);
      const inSet=new Set(rec.map(r=>r.id));
      const missing=list.filter(([i])=>!inSet.has(i)&&!leaveOnDate(this.leave,i,todayISO));
      if(has('present today'))return `${inSet.size} of ${list.length} staff have clocked in today.`;
      if(!missing.length)return 'Everyone has clocked in today (or is on approved leave).';
      return `${missing.length} not yet clocked in today:\n`+missing.map(([i,s])=>`• ${s.name} (${s.unit||''})`).join('\n');
    }
    // staff files completeness
    if(has('staff file','incomplete','file status','missing')){
      const files=await API.getAllHRFiles();
      const fm={};files.forEach(f=>fm[f.staff_id]=f);
      const inc=list.filter(([i])=>{
        const f=fm[i];if(!f)return true;
        return [f.dob,f.phone,f.next_of_kin,f.ssnit_number].filter(v=>v&&String(v).trim()).length<4;});
      if(!inc.length)return 'All staff files are complete. 🎉';
      return `${inc.length} staff file(s) still incomplete:\n`+inc.map(([i,s])=>`• ${s.name}`).join('\n');
    }
    // person lookup
    const match=list.find(([i,s])=>q.includes((s.name||'').toLowerCase().split(' ')[0])&&(s.name||'').length>2);
    if(match){
      const [id,s]=match;
      const l=leaveOnDate(this.leave,id,todayISO);
      const f=this._contractFlag(s.contractEnd);
      const clockedIn=this.records.some(r=>r.id===id&&fmtD(r.date||r.in)===todayStr);
      return `${s.name} (${id})\n• Unit: ${s.unit||'—'}\n• Role: ${roleLabel(s.role)}\n• Today: ${l?('on '+l.type):(clockedIn?'clocked in':'not clocked in')}\n• Contract: ${s.contractEnd?f.label:'no date on file'}`;
    }
    return "I can answer questions about:\n• Headcount, gender and unit breakdown\n• Who is on leave today\n• Pending leave requests\n• Contracts expiring soon\n• Birthdays this month\n• Who has not clocked in today\n• Incomplete staff files\n• A specific person (type their first name)\n\nTry rephrasing, or tap one of the suggestions.";
  }

  /* ═══════════════════════════════════════════
     RECRUITMENT — vacancies & candidate pipeline
     Visible to HR and Country Leader
  ═══════════════════════════════════════════ */
  _rcStages(){return['Applied','Screening','Interview','Offer','Hired','Rejected'];}
  _rcStageFlag(st){
    const m={Applied:'none',Screening:'amber',Interview:'amber',Offer:'green',Hired:'green',Rejected:'red'};
    return `<span class="c-flag ${m[st]||'none'}">${st}</span>`;
  }
  _rcVacFlag(st){
    const m={Draft:'none',Open:'green',Interviewing:'amber',Offer:'amber',Filled:'green',Closed:'none'};
    return `<span class="c-flag ${m[st]||'none'}">${st}</span>`;
  }
  _stars(n){n=+n||0;return n?`<span class="rc-star">${'★'.repeat(n)}${'☆'.repeat(5-n)}</span>`:'<span style="color:var(--text3);font-size:.72rem">—</span>';}
  async renderRecruit(p){
    const vb=$('m-rc-vac-body'),ab=$('m-rc-app-body');if(!vb)return;
    vb.innerHTML='<tr><td colspan="8" style="color:var(--text3)">Loading…</td></tr>';
    const vacs=await API._get('recruitment_vacancies','order=created_at.desc&limit=200')||[];
    const apps=await API._get('recruitment_applicants','order=applied_date.desc.nullslast&limit=500')||[];
    this._vacs=vacs;this._apps=apps;
    const vmap={};vacs.forEach(v=>vmap[v.id]=v);
    // vacancy filter dropdown
    const fsel=$('m-rc-filter');
    const cur=fsel?.value||'';
    if(fsel)fsel.innerHTML='<option value="">All Vacancies</option>'+vacs.map(v=>`<option value="${v.id}" ${v.id===cur?'selected':''}>${v.position}</option>`).join('');
    const fApps=cur?apps.filter(a=>a.vacancy_id===cur):apps;
    // pipeline counters
    const pipe=$('m-rc-pipe');
    if(pipe){
      const counts={};this._rcStages().forEach(st=>counts[st]=fApps.filter(a=>a.stage===st).length);
      const colors={Applied:'var(--text2)',Screening:'#d97706',Interview:'#4338ca',Offer:'#0d9488',Hired:'#16a34a',Rejected:'#dc2626'};
      pipe.innerHTML=this._rcStages().map(st=>`<div class="rc-stage"><div class="rc-n" style="color:${colors[st]}">${counts[st]}</div><div class="rc-l">${st}</div></div>`).join('')
        +`<div class="rc-stage" style="background:transparent;border-style:dashed"><div class="rc-n">${vacs.filter(v=>v.status==='Open'||v.status==='Interviewing').length}</div><div class="rc-l">Open Roles</div></div>`;
    }
    // vacancies
    vb.innerHTML=vacs.length?vacs.map(v=>{
      const n=apps.filter(a=>a.vacancy_id===v.id).length;
      const closing=v.closing_date?String(v.closing_date).slice(0,10):'—';
      const late=v.closing_date&&String(v.closing_date).slice(0,10)<new Date().toISOString().slice(0,10)&&['Open','Interviewing'].includes(v.status);
      return `<tr><td><strong>${v.position}</strong>${v.description?`<br><span style="font-size:.68rem;color:var(--text3)">${String(v.description).slice(0,60)}${v.description.length>60?'…':''}</span>`:''}</td>
        <td style="font-size:.8rem">${v.unit||'—'}</td><td style="font-size:.78rem">${v.employ_type||'—'}</td>
        <td>${v.openings||1}</td><td>${this._rcVacFlag(v.status)}</td>
        <td style="font-size:.78rem">${closing}${late?' <span class="c-flag red">past</span>':''}</td>
        <td><strong>${n}</strong></td>
        <td><button class="bsm bsm-navy" onclick="APP.openVacancyModal('${v.id}')">✏</button></td></tr>`;
    }).join(''):'<tr><td colspan="8"><div class="empty"><div class="empty-ico">🧑‍💼</div>No vacancies yet — click ＋ New Vacancy</div></td></tr>';
    // candidates
    if(ab)ab.innerHTML=fApps.length?fApps.map(a=>{
      const v=vmap[a.vacancy_id];
      return `<tr><td><strong>${a.name}</strong></td>
        <td style="font-size:.8rem">${v?v.position:'—'}</td>
        <td style="font-size:.74rem">${a.email||''}${a.email&&a.phone?'<br>':''}${a.phone||''}</td>
        <td>${this._rcStageFlag(a.stage)}</td><td>${this._stars(a.rating)}</td>
        <td>${a.cv_url?`<a href="${a.cv_url}" target="_blank" style="color:var(--teal)">📎 CV</a>`:'—'}</td>
        <td style="font-size:.76rem">${a.applied_date?String(a.applied_date).slice(0,10):'—'}</td>
        <td><button class="bsm bsm-navy" onclick="APP.openApplicantModal('${a.id}')">✏</button></td></tr>`;
    }).join(''):'<tr><td colspan="8"><div class="empty"><div class="empty-ico">👤</div>No candidates yet</div></td></tr>';
  }
  openVacancyModal(id){
    const v=id?(this._vacs||[]).find(x=>x.id===id):null;
    $('vc-id').value=v?.id||'';
    $('vc-position').value=v?.position||'';
    $('vc-unit').value=v?.unit||'';
    $('vc-type').value=v?.employ_type||'Full-time';
    $('vc-openings').value=v?.openings||1;
    $('vc-status').value=v?.status||'Open';
    $('vc-posted').value=v?.posted_date?String(v.posted_date).slice(0,10):new Date().toISOString().slice(0,10);
    $('vc-closing').value=v?.closing_date?String(v.closing_date).slice(0,10):'';
    $('vc-desc').value=v?.requirements||'';
    if($('vc-jd'))$('vc-jd').value=v?.description||'';
    const mgrs=Object.entries(this.staff).filter(([i,s])=>s.role==='manager'||s.role==='country_leader')
      .sort((a,b)=>a[1].name.localeCompare(b[1].name));
    $('vc-manager').innerHTML='<option value="">— Not set —</option>'+mgrs.map(([i,s])=>`<option value="${i}" ${v?.hiring_manager===i?'selected':''}>${s.name}</option>`).join('');
    $('vc-msg').textContent='';
    $('vacancy-modal').classList.add('open');
  }
  async saveVacancy(){
    const pos=$('vc-position').value.trim();
    if(!pos)return $('vc-msg').innerHTML='<span style="color:var(--red)">Position title is required.</span>';
    const id=$('vc-id').value||this._uid('VAC');
    $('vc-msg').innerHTML='<span style="color:var(--teal)">⏳ Saving…</span>';
    const r=await API._upsert('recruitment_vacancies',[{id,position:pos,unit:$('vc-unit').value.trim(),
      employ_type:$('vc-type').value,openings:+$('vc-openings').value||1,status:$('vc-status').value,
      hiring_manager:$('vc-manager').value,posted_date:$('vc-posted').value||null,
      closing_date:$('vc-closing').value||null,description:$('vc-jd')?.value.trim()||'',requirements:$('vc-desc').value.trim(),
      created_by:this.user.name}]);
    if(r){this.audit('Vacancy saved','HR',pos,$('vc-status').value);closeModal('vacancy-modal');toast('Vacancy saved ✓');this.renderRecruit('m-');}
    else $('vc-msg').innerHTML='<span style="color:var(--red)">Save failed: '+(API.lastError||'')+'</span>';
  }
  openApplicantModal(id){
    const a=id?(this._apps||[]).find(x=>x.id===id):null;
    const vacs=this._vacs||[];
    if(!vacs.length&&!a)return toast('Create a vacancy first','err');
    $('ap-id').value=a?.id||'';
    $('ap-vacancy').innerHTML=vacs.map(v=>`<option value="${v.id}" ${a?.vacancy_id===v.id?'selected':''}>${v.position} (${v.unit||''})</option>`).join('');
    $('ap-name').value=a?.name||'';
    $('ap-email').value=a?.email||'';
    $('ap-phone').value=a?.phone||'';
    $('ap-stage').value=a?.stage||'Applied';
    $('ap-rating').value=a?.rating||0;
    $('ap-notes').value=a?.notes||'';
    $('ap-date').value=a?.applied_date?String(a.applied_date).slice(0,10):new Date().toISOString().slice(0,10);
    $('ap-cv-url').value=a?.cv_url||'';
    $('ap-cv-shown').innerHTML=a?.cv_url?`<span class="hr-doc-chip">📎 <a href="${a.cv_url}" target="_blank">View CV</a></span>`:'<span style="font-size:.74rem;color:var(--text3)">No CV uploaded</span>';
    $('ap-msg').textContent='';
    $('applicant-modal').classList.add('open');
  }
  async uploadCV(){
    const inp=$('ap-cv-file');const msg=$('ap-msg');
    if(!inp?.files?.length)return toast('Choose a file first','err');
    const file=inp.files[0];
    if(file.size>5*1024*1024)return toast('File too large (max 5MB)','err');
    msg.innerHTML='<span style="color:var(--teal)">⏳ Uploading CV…</span>';
    try{
      const b64=await this._fileToBase64(file);
      const nm=($('ap-name').value.trim()||'candidate').replace(/[^a-z0-9]/gi,'_');
      const r=await API.gasPost({action:'uploadHRDoc',staffId:'RECRUIT_'+nm,fileName:file.name,fileData:b64,mimeType:file.type});
      if(r&&r.success&&r.fileUrl){
        $('ap-cv-url').value=r.fileUrl;
        $('ap-cv-shown').innerHTML=`<span class="hr-doc-chip">📎 <a href="${r.fileUrl}" target="_blank">View CV</a></span>`;
        inp.value='';
        msg.innerHTML='<span style="color:var(--green)">✓ CV uploaded — click Save to confirm.</span>';
      }else msg.innerHTML='<span style="color:var(--red)">Upload failed: '+((r&&r.error)||'no response')+'</span>';
    }catch(e){msg.innerHTML='<span style="color:var(--red)">Upload error.</span>';}
  }
  async saveApplicant(){
    const nm=$('ap-name').value.trim();
    if(!nm)return $('ap-msg').innerHTML='<span style="color:var(--red)">Candidate name is required.</span>';
    const id=$('ap-id').value||this._uid('APP');
    $('ap-msg').innerHTML='<span style="color:var(--teal)">⏳ Saving…</span>';
    const r=await API._upsert('recruitment_applicants',[{id,vacancy_id:$('ap-vacancy').value,name:nm,
      email:$('ap-email').value.trim(),phone:$('ap-phone').value.trim(),stage:$('ap-stage').value,
      rating:+$('ap-rating').value||0,cv_url:$('ap-cv-url').value,notes:$('ap-notes').value.trim(),
      applied_date:$('ap-date').value||null,updated_at:new Date().toISOString()}]);
    if(r){this.audit('Candidate saved','HR',nm,$('ap-stage').value);closeModal('applicant-modal');toast('Candidate saved ✓');this.renderRecruit('m-');}
    else $('ap-msg').innerHTML='<span style="color:var(--red)">Save failed: '+(API.lastError||'')+'</span>';
  }

  /* ── Staff self-assessment ── */
  async renderMyAppraisals(prefix){
    const body=$((prefix||'st-')==='m-'?'m-apr-body':'st-apr-body');if(!body)return;

    body.innerHTML='<tr><td colspan="6" style="color:var(--text3)">Loading…</td></tr>';
    const rows=await API._get('performance_appraisals','staff_id=eq.'+encodeURIComponent(this.user.id)+'&order=review_date.desc.nullslast')||[];
    this._myApr=rows;
    if(!rows.length){body.innerHTML='<tr><td colspan="6"><div class="empty"><div class="empty-ico">📊</div>No appraisals yet — click ＋ Start Self-Assessment</div></td></tr>';return;}
    const sc={Draft:'none','Self-Assessed':'amber','Manager Reviewed':'amber',Acknowledged:'green',Closed:'green'};
    body.innerHTML=rows.map(r=>{
      let k=[];try{k=JSON.parse(r.kpas||'[]');}catch(e){}
      const selfScore=k.reduce((t,x)=>{
        const rs=(x.kpis||[]).map(y=>+y.self||0).filter(v=>v>0);
        const avg=rs.length?rs.reduce((a,b)=>a+b,0)/rs.length:0;
        return t+avg*(+x.weight||0)/100;},0);
      const done=r.status!=='Draft';
      return `<tr><td><strong>${r.period||'—'}</strong></td><td style="font-size:.8rem">${r.review_type||'Annual'}</td>
        <td style="font-size:.8rem">${this._sName(r.line_manager)||'—'}</td>
        <td>${selfScore.toFixed(2)} / 5</td>
        <td><span class="c-flag ${sc[r.status]||'none'}">${r.status||'Draft'}</span></td>
        <td>${done?'<span style="font-size:.72rem;color:var(--text3)">submitted</span>':`<button class="bsm bsm-navy" onclick="APP.openMyAppraisal('${r.id}')">✏</button>`}</td></tr>`;
    }).join('');
  }
  openMyAppraisal(id){
    const r=id?(this._myApr||[]).find(x=>x.id===id):null;
    if(r&&r.status!=='Draft')return toast('This review has already been submitted','info');
    const me=this.staff[this.user.id]||{};
    $('ma-id').value=r?.id||'';
    $('ma-period').value=r?.period||(new Date().getFullYear()+' Annual');
    $('ma-type').value=r?.review_type||'Annual';
    const mgrs=Object.entries(this.staff).filter(([i,s])=>(s.role==='manager'||s.role==='country_leader')&&i!==this.user.id)
      .sort((a,b)=>a[1].name.localeCompare(b[1].name));
    const pick=r?.line_manager||me.supervisor||'';
    $('ma-mgr').innerHTML='<option value="">— Select your supervisor —</option>'+
      mgrs.map(([i,s])=>`<option value="${i}" ${pick===i?'selected':''}>${s.name} — ${s.unit||''}</option>`).join('');
    $('ma-comment').value=r?.emp_comment||'';
    if($('ma-overall'))$('ma-overall').value=r?.self_overall||0;
    if($('ma-career'))$('ma-career').value=r?.career_goals||'';
    $('ma-dev').value=r?.dev_plan||'';
    let k=[];try{k=JSON.parse(r?.kpas||'[]');}catch(e){}
    this._myKpas=(k&&k.length)?k:this._defaultKPAs();
    this._renderMyKPAs();
    $('ma-msg').textContent='';
    $('myapr-modal').classList.add('open');
  }
  _renderMyKPAs(){
    const el=$('ma-kpas');if(!el)return;
    el.innerHTML=this._myKpas.map((k,i)=>`
      <div class="kpa-box">
        <div class="kpa-hd"><span style="font-size:.72rem;color:var(--text3);font-weight:700">KPA ${i+1}</span>
          <strong style="flex:1">${k.name||''}</strong>
          <span style="font-size:.74rem;color:var(--text2)">${k.weight||0}%</span></div>
        ${(k.kpis||[]).map((x,j)=>`<div class="kpi-row">
          <input class="fi ki" value="${x.kpi||''}" placeholder="What did you deliver?" oninput="APP._myKpiSet(${i},${j},'kpi',this.value)">
          <input class="fi kr" type="number" min="0" max="5" step="0.5" value="${x.self||0}" title="Your rating" oninput="APP._myKpiSet(${i},${j},'self',this.value)">
          <button class="bsm" style="background:var(--surf);border:1px solid var(--border);color:var(--text2)" onclick="APP._myDelKPI(${i},${j})">✕</button>
        </div>`).join('')}
        <button class="bsm" style="background:var(--surf);border:1px solid var(--border);color:var(--text2);margin-top:.2rem" onclick="APP._myAddKPI(${i})">＋ Add achievement</button>
      </div>`).join('');
    const w=this._myKpas.reduce((t,k)=>t+(+k.weight||0),0);
    const sc=this._myKpas.reduce((t,k)=>{
      const rs=(k.kpis||[]).map(x=>+x.self||0).filter(v=>v>0);
      const avg=rs.length?rs.reduce((a,b)=>a+b,0)/rs.length:0;
      return t+avg*(+k.weight||0)/100;},0);
    if($('ma-wtot'))$('ma-wtot').textContent=w+'%';
    if($('ma-score'))$('ma-score').textContent=sc.toFixed(2);
    return sc;
  }
  _myKpiSet(i,j,f,v){this._myKpas[i].kpis[j][f]=(f==='self')?(+v||0):v;if(f==='self')this._renderMyKPAs();}
  _myAddKPI(i){this._myKpas[i].kpis.push({kpi:'',standard:'',self:0,mgr:0});this._renderMyKPAs();}
  _myDelKPI(i,j){this._myKpas[i].kpis.splice(j,1);this._renderMyKPAs();}
  async submitMyAppraisal(){
    const mgr=$('ma-mgr').value;
    if(!mgr)return $('ma-msg').innerHTML='<span style="color:var(--red)">Please select your supervisor.</span>';
    const filled=this._myKpas.some(k=>(k.kpis||[]).some(x=>(+x.self||0)>0));
    if(!filled)return $('ma-msg').innerHTML='<span style="color:var(--red)">Please rate yourself on at least one item.</span>';
    if(!+($('ma-overall')?.value||0))return $('ma-msg').innerHTML='<span style="color:var(--red)">Select your overall self-assessment under Final Assessment.</span>';
    if(!confirm('Submit your self-assessment to '+this._sName(mgr)+'?\n\nYou will not be able to edit it afterwards.'))return;
    const score=this._renderMyKPAs();
    const id=$('ma-id').value||this._uid('APR');
    const me=this.staff[this.user.id]||{};
    $('ma-msg').innerHTML='<span style="color:var(--teal)">⏳ Submitting…</span>';
    const r=await API._upsert('performance_appraisals',[{id,staff_id:this.user.id,period:$('ma-period').value.trim(),
      review_type:$('ma-type').value,review_date:new Date().toISOString().slice(0,10),
      job_title:'',department:me.unit||'',location:'',line_manager:mgr,
      kpas:JSON.stringify(this._myKpas),final_score:+score.toFixed(2),
      dev_plan:$('ma-dev').value.trim(),emp_comment:$('ma-comment').value.trim(),
      self_overall:+($('ma-overall')?.value||0),self_score:+score.toFixed(2),
      self_submitted:new Date().toISOString(),career_goals:$('ma-career')?.value.trim()||'',
      training_needs:$('ma-dev')?.value.trim()||'',
      status:'Self-Assessed',updated_at:new Date().toISOString()}]);
    if(!r)return $('ma-msg').innerHTML='<span style="color:var(--red)">Submit failed: '+(API.lastError||'')+'</span>';
    const recips=[];
    const push=id2=>{const e=this.staff[id2]?.email;if(e)recips.push({name:this.staff[id2].name,email:e});};
    push(mgr);push(HR_MANAGER_ID);push(COUNTRY_LEADER_ID);
    API.gasPost({action:'appraisalNotify',staffName:this.user.name,period:$('ma-period').value.trim(),
      supervisor:this._sName(mgr),score:score.toFixed(2),recipients:recips}).catch(()=>{});
    this.audit('Self-assessment submitted','HR',this.user.name,$('ma-period').value.trim());
    closeModal('myapr-modal');
    toast('Self-assessment submitted to '+this._sName(mgr)+' ✓');
    this.renderMyAppraisals();
  }

  /* ── Performance Appraisal (THP Annual Review Form) ── */
  _ratingWord(v){return['—','Unsatisfactory performance','Performed Below Performance Criteria','Achieved Performance Criteria','Achieved Above Performance Criteria','Exceeded on all Performance Criteria'][Math.round(+v||0)]||'—';}
  /* KPAs exactly as in the THP Annual Performance Review Form */
  _defaultKPAs(){return[
    {name:'Financial',weight:30,kpis:[{kpi:'',standard:'',self:0,mgr:0}],commentary:''},
    {name:'Customer',weight:30,kpis:[{kpi:'',standard:'',self:0,mgr:0}],commentary:''},
    {name:'Internal Business Process',weight:20,kpis:[{kpi:'',standard:'',self:0,mgr:0}],commentary:''},
    {name:'Learning and Growth',weight:10,kpis:[{kpi:'',standard:'',self:0,mgr:0}],commentary:''},
    {name:'Ethics and Compliance',weight:10,kpis:[{kpi:'',standard:'',self:0,mgr:0}],commentary:''}];}
  async renderPerf(p){
    const body=$('m-perf-body');if(!body)return;
    body.innerHTML='<tr><td colspan="8" style="color:var(--text3)">Loading…</td></tr>';
    let rows=await API._get('performance_appraisals','order=review_date.desc.nullslast&limit=300')||[];
    // HR sees only reviews the line manager has completed
    rows=rows.filter(r=>['Manager Reviewed','Acknowledged','Closed'].includes(r.status));
    const st=$('m-ap-status')?.value||'';
    if(st)rows=rows.filter(r=>r.status===st);
    this._perfRows=rows;
    if(!rows.length){body.innerHTML='<tr><td colspan="8"><div class="empty"><div class="empty-ico">📊</div>No appraisals yet — click ＋ New Appraisal</div></td></tr>';return;}
    const sc={Draft:'none','Self-Assessed':'amber','Manager Reviewed':'amber',Acknowledged:'green',Closed:'green'};
    body.innerHTML=rows.map(r=>{
      const score=(+r.final_score||0).toFixed(2);
      const band=score>=4.5?'green':score>=3?'green':score>=2?'amber':'red';
      return `<tr><td><strong>${this._sName(r.staff_id)}</strong><br><span style="font-size:.7rem;color:var(--text3)">${r.job_title||''}</span></td>
        <td style="font-size:.8rem">${r.period||'—'}</td><td style="font-size:.78rem">${r.review_type||'Annual'}</td>
        <td><strong>${score}</strong> / 5</td>
        <td><span class="c-flag ${band}">${this._ratingWord(score)}</span></td>
        <td><span class="c-flag ${sc[r.status]||'none'}">${r.status||'Draft'}</span>${r.ack_by_staff?' ✔':''}</td>
        <td style="font-size:.76rem">${r.review_date?String(r.review_date).slice(0,10):'—'}</td>
        <td><button class="bsm bsm-navy" onclick="APP.openPerfModal('${r.id}')">✏</button>
          <button class="bsm" style="background:rgba(239,68,68,.12);color:var(--red)" onclick="APP.deleteAppraisal('${r.id}')">🗑</button></td></tr>`;
    }).join('');
  }
  openPerfModal(id){
    const r=id?(this._perfRows||[]).find(x=>x.id===id):null;
    this._popStaffSel('pf-staff',r?.staff_id);
    const mgrs=Object.entries(this.staff).filter(([i,s])=>s.role==='manager'||s.role==='country_leader').sort((a,b)=>a[1].name.localeCompare(b[1].name));
    $('pf-mgr').innerHTML='<option value="">— Not set —</option>'+mgrs.map(([i,s])=>`<option value="${i}" ${r?.line_manager===i?'selected':''}>${s.name}</option>`).join('');
    $('pf-id').value=r?.id||'';
    $('pf-type').value=r?.review_type||'Annual';
    $('pf-period').value=r?.period||(new Date().getFullYear()+' Annual');
    $('pf-date').value=r?.review_date?String(r.review_date).slice(0,10):new Date().toISOString().slice(0,10);
    $('pf-status').value=r?.status||'Draft';
    $('pf-job').value=r?.job_title||'';
    $('pf-dept').value=r?.department||'';
    $('pf-loc').value=r?.location||'Accra';
    $('pf-startco').value=r?.start_company?String(r.start_company).slice(0,10):'';
    $('pf-startpos').value=r?.start_position?String(r.start_position).slice(0,10):'';
    $('pf-dev').value=r?.dev_plan||'';
    $('pf-emp').value=r?.emp_comment||'';
    $('pf-ack').checked=!!r?.ack_by_staff;
    let k=[];try{k=JSON.parse(r?.kpas||'[]');}catch(e){}
    this._kpas=(k&&k.length)?k:this._defaultKPAs();
    if(!r)this._pfFillStaff();
    this._renderKPAs();
    $('pf-msg').textContent='';
    $('perf-modal').classList.add('open');
  }
  _pfFillStaff(){
    const id=$('pf-staff')?.value,s=this.staff[id];if(!s)return;
    if(!$('pf-dept').value)$('pf-dept').value=s.unit||'';
    if(!$('pf-startco').value&&s.contractStart)$('pf-startco').value=String(s.contractStart).slice(0,10);
    if(!$('pf-mgr').value&&s.supervisor)$('pf-mgr').value=s.supervisor;
  }
  _renderKPAs(){
    const el=$('pf-kpas');if(!el)return;
    el.innerHTML=this._kpas.map((k,i)=>`
      <div class="kpa-box">
        <div class="kpa-hd">
          <span style="font-size:.72rem;color:var(--text3);font-weight:700">KPA ${i+1}</span>
          <input class="fi kn" value="${k.name||''}" placeholder="KPA name" oninput="APP._kpaSet(${i},'name',this.value)">
          <input class="fi kw" type="number" min="0" max="100" value="${k.weight||0}" oninput="APP._kpaSet(${i},'weight',this.value)"> <span style="font-size:.74rem;color:var(--text2)">% weight</span>
          <button class="bsm" style="background:rgba(239,68,68,.12);color:var(--red)" onclick="APP.delKPA(${i})">✕</button>
        </div>
        ${(k.kpis||[]).map((x,j)=>`<div class="kpi-row">
          <input class="fi ki" value="${x.kpi||''}" placeholder="KPI" oninput="APP._kpiSet(${i},${j},'kpi',this.value)">
          <input class="fi ks" value="${x.standard||''}" placeholder="Quality standard" oninput="APP._kpiSet(${i},${j},'standard',this.value)">
          <input class="fi kr" type="number" min="0" max="5" step="0.5" value="${x.self||0}" title="Self rating" oninput="APP._kpiSet(${i},${j},'self',this.value)">
          <input class="fi kr" type="number" min="0" max="5" step="0.5" value="${x.mgr||0}" title="Manager rating" oninput="APP._kpiSet(${i},${j},'mgr',this.value)">
          <button class="bsm" style="background:var(--surf);border:1px solid var(--border);color:var(--text2)" onclick="APP.delKPI(${i},${j})">✕</button>
        </div>`).join('')}
        <button class="bsm" style="background:var(--surf);border:1px solid var(--border);color:var(--text2);margin:.2rem 0 .4rem" onclick="APP.addKPI(${i})">＋ KPI</button>
        <input class="fi" value="${k.commentary||''}" placeholder="Commentary" oninput="APP._kpaSet(${i},'commentary',this.value)">
        <div class="kpa-avg">Average rating: <strong>${this._kpaAvg(k).toFixed(2)}</strong> · Weighted: <strong>${(this._kpaAvg(k)*(+k.weight||0)/100).toFixed(3)}</strong></div>
      </div>`).join('');
    this._pfTotals();
  }
  _kpaAvg(k){
    const rs=(k.kpis||[]).map(x=>{
      const self=+x.self||0,mgr=+x.mgr||0;
      if(self&&mgr)return (self+mgr)/2;
      return mgr||self||0;
    }).filter(v=>v>0);
    return rs.length?rs.reduce((a,b)=>a+b,0)/rs.length:0;
  }
  _pfTotals(){
    const w=this._kpas.reduce((t,k)=>t+(+k.weight||0),0);
    const score=this._kpas.reduce((t,k)=>t+this._kpaAvg(k)*(+k.weight||0)/100,0);
    const wt=$('pf-wtot');if(wt){wt.textContent=w+'%';wt.style.color=Math.abs(w-100)<0.01?'var(--green)':'var(--red)';}
    const sc=$('pf-score');if(sc)sc.textContent=score.toFixed(2);
    return{w,score};
  }
  _kpaSet(i,f,v){this._kpas[i][f]=(f==='weight')?(+v||0):v;if(f==='weight')this._pfTotals();else if(f==='name'||f==='commentary')return;this._pfTotals();}
  _kpiSet(i,j,f,v){this._kpas[i].kpis[j][f]=(f==='self'||f==='mgr')?(+v||0):v;if(f==='self'||f==='mgr')this._renderKPAs();}
  addKPA(){this._kpas.push({name:'',weight:0,kpis:[{kpi:'',standard:'',self:0,mgr:0}],commentary:''});this._renderKPAs();}
  delKPA(i){this._kpas.splice(i,1);this._renderKPAs();}
  addKPI(i){this._kpas[i].kpis.push({kpi:'',standard:'',self:0,mgr:0});this._renderKPAs();}
  delKPI(i,j){this._kpas[i].kpis.splice(j,1);this._renderKPAs();}
  async savePerf(){
    const staff=$('pf-staff').value;
    if(!staff)return $('pf-msg').innerHTML='<span style="color:var(--red)">Select a staff member.</span>';
    const {w,score}=this._pfTotals();
    if(Math.abs(w-100)>0.01)return $('pf-msg').innerHTML='<span style="color:var(--red)">KPA weightings must total 100% (currently '+w+'%).</span>';
    const id=$('pf-id').value||this._uid('APR');
    $('pf-msg').innerHTML='<span style="color:var(--teal)">⏳ Saving…</span>';
    const r=await API._upsert('performance_appraisals',[{id,staff_id:staff,period:$('pf-period').value.trim(),
      review_type:$('pf-type').value,review_date:$('pf-date').value||null,job_title:$('pf-job').value.trim(),
      department:$('pf-dept').value.trim(),location:$('pf-loc').value.trim(),
      start_company:$('pf-startco').value||null,start_position:$('pf-startpos').value||null,
      line_manager:$('pf-mgr').value,kpas:JSON.stringify(this._kpas),final_score:+score.toFixed(2),
      dev_plan:$('pf-dev').value.trim(),emp_comment:$('pf-emp').value.trim(),status:$('pf-status').value,
      ack_by_staff:$('pf-ack').checked,ack_date:$('pf-ack').checked?new Date().toISOString().slice(0,10):null,
      updated_at:new Date().toISOString()}]);
    if(r){this.audit('Appraisal saved','HR',this._sName(staff),$('pf-period').value+' · score '+score.toFixed(2));
      closeModal('perf-modal');toast('Appraisal saved ✓');this.renderPerf('m-');}
    else $('pf-msg').innerHTML='<span style="color:var(--red)">Save failed: '+(API.lastError||'')+'</span>';
  }
  printAppraisal(){
    const {score}=this._pfTotals();
    const nm=this._sName($('pf-staff').value);
    const rows=this._kpas.map((k,i)=>`
      <tr><td colspan="6" style="background:#eef1ff;font-weight:bold">KPA ${i+1}: ${k.name||''} — weighting ${k.weight||0}%</td></tr>
      <tr class="hd"><th>KPI</th><th>Quality Standard</th><th>Self (1-5)</th><th>Manager (1-5)</th><th>Final Avg</th><th>Weighted</th></tr>
      ${(k.kpis||[]).map(x=>{const f=(+x.self&&+x.mgr)?((+x.self+ +x.mgr)/2):(+x.mgr|| +x.self||0);
        return `<tr><td>${x.kpi||''}</td><td>${x.standard||''}</td><td>${x.self||0}</td><td>${x.mgr||0}</td><td>${f.toFixed(2)}</td><td></td></tr>`;}).join('')}
      <tr><td colspan="4" style="font-style:italic">Commentary: ${k.commentary||'—'}</td><td><strong>${this._kpaAvg(k).toFixed(2)}</strong></td><td><strong>${(this._kpaAvg(k)*(+k.weight||0)/100).toFixed(3)}</strong></td></tr>`).join('');
    const w=window.open('','_blank');
    w.document.write(`<html><head><meta charset="UTF-8"><title>Appraisal — ${nm}</title><style>
      body{font-family:Arial,sans-serif;font-size:11px;color:#1e293b;padding:14mm}
      h1{font-size:15px;color:#2D3592;border-bottom:3px solid #2D3592;padding-bottom:7px;margin:0 0 4px}
      .meta{font-size:10px;color:#555;margin-bottom:12px}
      table{width:100%;border-collapse:collapse;margin-bottom:12px}
      td,th{border:1px solid #d5d5d5;padding:5px 7px;font-size:10px;text-align:left;vertical-align:top}
      .hd th{background:#2D3592;color:#fff;font-size:9.5px}
      .score{font-size:13px;font-weight:bold;color:#2D3592;margin:10px 0}
      .sig{display:flex;gap:40px;margin-top:36px}
      .sig div{flex:1;border-top:1px solid #777;padding-top:5px;font-size:10px}
      </style></head><body>
      <h1>The Hunger Project — Ghana · Annual Performance Review</h1>
      <div class="meta"><strong>${nm}</strong> · ${$('pf-job').value||''} · ${$('pf-dept').value||''} · ${$('pf-loc').value||''}<br>
        Period: ${$('pf-period').value||''} · Type: ${$('pf-type').value} · Review date: ${$('pf-date').value||''}<br>
        Line manager: ${this._sName($('pf-mgr').value)||'—'}</div>
      <table>${rows}</table>
      <div class="score">FINAL SCORE: ${score.toFixed(2)} / 5 — ${this._ratingWord(score)}</div>
      <table><tr><td style="width:50%"><strong>Development Plan</strong><br>${$('pf-dev').value||'—'}</td>
        <td><strong>Employee Comment</strong><br>${$('pf-emp').value||'—'}</td></tr></table>
      <div class="sig"><div>Employee Signature &amp; Date</div><div>Line Manager Signature &amp; Date</div></div>
      <div style="text-align:center;margin-top:16px" class="no-print">
        <button onclick="window.print()" style="padding:8px 22px;background:#2D3592;color:#fff;border:0;border-radius:6px;font-weight:600;cursor:pointer">🖨 Print</button></div>
      </body></html>`);
    w.document.close();
  }

  /* ── Training ── */
  async renderTraining(p){
    const body=$(p+'train-body');if(!body)return;
    body.innerHTML='<tr><td colspan="7" style="color:var(--text3)">Loading…</td></tr>';
    const rows=await API._get('training_records','order=completed_date.desc.nullslast&limit=300')||[];
    this._trainRows=rows;
    if(!rows.length){body.innerHTML='<tr><td colspan="7"><div class="empty"><div class="empty-ico">🎓</div>No training records yet</div></td></tr>';return;}
    const today=new Date().toISOString().slice(0,10);
    body.innerHTML=rows.map(r=>{
      const exp=r.expiry_date?String(r.expiry_date).slice(0,10):'';
      const expFlag=exp?(exp<today?`<span class="c-flag red">Expired ${exp}</span>`:`<span class="c-flag ${exp<new Date(Date.now()+60*86400000).toISOString().slice(0,10)?'amber':'green'}">${exp}</span>`):'—';
      return`<tr><td><strong>${this._sName(r.staff_id)}</strong></td><td style="font-size:.8rem">${r.course||'—'}</td><td style="font-size:.78rem">${r.provider||'—'}</td><td style="font-size:.78rem">${r.completed_date?String(r.completed_date).slice(0,10):'—'}</td><td>${expFlag}</td><td>${r.certificate_url?`<a href="${r.certificate_url}" target="_blank" style="color:var(--teal)">📎 View</a>`:'—'}</td><td><button class="bsm bsm-navy" onclick="APP.openTrainModal('${r.id}')">✏</button></td></tr>`;
    }).join('');
  }
  openTrainModal(id){
    const r=id?(this._trainRows||[]).find(x=>x.id===id):null;
    this._popStaffSel('tr-staff',r?.staff_id);
    $('tr-id').value=r?.id||'';
    $('tr-course').value=r?.course||'';
    $('tr-provider').value=r?.provider||'';
    $('tr-completed').value=r?.completed_date?String(r.completed_date).slice(0,10):'';
    $('tr-expiry').value=r?.expiry_date?String(r.expiry_date).slice(0,10):'';
    $('tr-cert').value=r?.certificate_url||'';
    $('tr-notes').value=r?.notes||'';
    $('tr-msg').textContent='';
    $('train-modal').classList.add('open');
  }
  async saveTrain(){
    const id=$('tr-id').value||this._uid('TR');
    if(!$('tr-course').value.trim())return $('tr-msg').innerHTML='<span style="color:var(--red)">Course is required.</span>';
    const r=await API._upsert('training_records',[{id,staff_id:$('tr-staff').value,course:$('tr-course').value.trim(),provider:$('tr-provider').value.trim(),completed_date:$('tr-completed').value||null,expiry_date:$('tr-expiry').value||null,certificate_url:$('tr-cert').value.trim(),notes:$('tr-notes').value.trim()}]);
    if(r){this.audit('Training record saved','HR',this._sName($('tr-staff').value),$('tr-course').value.trim());closeModal('train-modal');toast('Training saved ✓');this.renderTraining('m-');}
    else $('tr-msg').innerHTML='<span style="color:var(--red)">Save failed.</span>';
  }

  /* ── Lifecycle & Cases (HR + Admin only) ── */
  async renderCases(p){
    const body=$(p+'case-body');if(!body)return;
    body.innerHTML='<tr><td colspan="7" style="color:var(--text3)">Loading…</td></tr>';
    const rows=await API.secureGet('hr_cases','order=opened_date.desc.nullslast&limit=300')||[];
    this._caseRows=rows;
    if(!rows.length){body.innerHTML='<tr><td colspan="7"><div class="empty"><div class="empty-ico">⚖</div>No cases recorded</div></td></tr>';return;}
    body.innerHTML=rows.map(r=>{
      const st=r.status==='Closed'?'<span class="c-flag green">Closed</span>':r.status==='Under Review'?'<span class="c-flag amber">Under Review</span>':'<span class="c-flag red">Open</span>';
      return`<tr><td><strong>${this._sName(r.staff_id)}</strong></td><td style="font-size:.8rem">${r.case_type||'—'}</td><td style="font-size:.78rem">${r.opened_date?String(r.opened_date).slice(0,10):'—'}</td><td>${st}</td><td style="font-size:.74rem;color:var(--text2)">${r.summary||'—'}</td><td style="font-size:.74rem;color:var(--text2)">${r.outcome||'—'}</td><td><button class="bsm bsm-navy" onclick="APP.openCaseModal('${r.id}')">✏</button></td></tr>`;
    }).join('');
  }
  openCaseModal(id){
    const r=id?(this._caseRows||[]).find(x=>x.id===id):null;
    this._popStaffSel('cs-staff',r?.staff_id);
    $('cs-id').value=r?.id||'';
    $('cs-ref').value=r?.case_ref||('CASE-'+new Date().getFullYear()+'-'+String(Math.floor(Math.random()*900)+100));
    $('cs-type').value=r?.case_type||'Complaint / Grievance';
    $('cs-sev').value=r?.severity||'Minor';
    $('cs-status').value=r?.status||'Open';
    $('cs-raised').value=r?.raised_by||'';
    $('cs-opened').value=r?.opened_date?String(r.opened_date).slice(0,10):new Date().toISOString().slice(0,10);
    $('cs-closed').value=r?.closed_date?String(r.closed_date).slice(0,10):'';
    $('cs-summary').value=r?.summary||'';
    $('cs-outcome').value=r?.outcome||'';
    $('cs-appeal').value=String(r?.appeal_lodged||false);
    $('cs-appealout').value=r?.appeal_outcome||'';
    let stg=[];try{stg=JSON.parse(r?.stages||'[]');}catch(e){}
    this._csStages=stg.length?stg:this._caseStages().map(n=>({name:n,done:false,date:''}));
    this._renderCaseStages();
    let docs=[];try{docs=JSON.parse(r?.documents||'[]');}catch(e){}
    $('cs-docs-json').value=JSON.stringify(docs);
    this._renderCaseDocs(docs);
    $('cs-msg').textContent='';
    $('case-modal').classList.add('open');
  }
  async saveCase(){
    const staff=$('cs-staff').value;
    if(!staff)return $('cs-msg').innerHTML='<span style="color:var(--red)">Select the staff member.</span>';
    const id=$('cs-id').value||this._uid('CS');
    $('cs-msg').innerHTML='<span style="color:var(--teal)">⏳ Saving…</span>';
    const r=await API.secureSave('hr_cases',[{id,staff_id:staff,case_ref:$('cs-ref').value,
      case_type:$('cs-type').value,severity:$('cs-sev').value,status:$('cs-status').value,
      raised_by:$('cs-raised').value.trim(),
      opened_date:$('cs-opened').value||null,closed_date:$('cs-closed').value||null,
      summary:$('cs-summary').value.trim(),outcome:$('cs-outcome').value.trim(),
      stages:JSON.stringify(this._csStages),documents:$('cs-docs-json').value||'[]',
      appeal_lodged:$('cs-appeal').value==='true',appeal_outcome:$('cs-appealout').value.trim(),
      updated_at:new Date().toISOString()}]);
    if(r){this.audit('Case file saved','Security',this._sName(staff),
      $('cs-ref').value+' · '+$('cs-type').value+' · '+$('cs-status').value);
      closeModal('case-modal');toast('Case saved ✓');this.renderCases('m-');}
    else $('cs-msg').innerHTML='<span style="color:var(--red)">Save failed: '+(API.lastError||'')+'</span>';
  }
  async renderOrgChart(p){
    const el=$((p||'m-')+'org-tree');if(!el)return;
    el.innerHTML='<div style="color:var(--text3);font-size:.8rem">Loading…</div>';
    const rows=await API._get('org_chart','select=*')||[];
    const oc={};rows.forEach(r=>oc[r.staff_id]=r);
    this._orgOverrides=oc;
    const files=await API._get('hr_staff_files','select=staff_id,photo_url')||[];
    const ph={};files.forEach(f=>{if(f.photo_url)ph[f.staff_id]=this._drivePhoto(f.photo_url);});
    const people=Object.entries(this.staff).filter(([i,s])=>s.role!=='admin');
    const mgrOf=id=>{
      const o=oc[id];
      if(o&&o.reports_to!==undefined&&o.reports_to!==null&&o.reports_to!=='')return o.reports_to;
      if(o&&o.reports_to==='')return '';
      return (this.staff[id]?.supervisor||'').trim();
    };
    const orderOf=id=>(oc[id]?.sort_order??100);
    const titleOf=id=>(oc[id]?.title||'').trim()||(this.staff[id]?.unit||'');
    const kids=parent=>people.filter(([i])=>{
      const m=mgrOf(i);
      return parent===''?(!m||!this.staff[m]||m===i):m===parent;
    }).sort((a,b)=>orderOf(a[0])-orderOf(b[0])||a[1].name.localeCompare(b[1].name));
    const isHR=this.user.id===HR_MANAGER_ID||this.user.role==='admin';
    const seen=new Set();
    const node=([id,st],lvl)=>{
      if(seen.has(id))return'';           // guard against circular reporting
      seen.add(id);
      const av=ph[id]
        ? `<img class="oc-av" src="${ph[id]}" alt="">`
        : `<div class="oc-av" style="background:${st.color||'#2D3592'}">${(st.name||'?')[0]}</div>`;
      const ch=kids(id);
      return `<li class="oc-lvl-${lvl}"><div class="oc-node">
          ${isHR?`<button class="oc-edit" title="Edit position" onclick="APP.openOrgModal('${id}')">✏</button>`:''}
          ${av}
          <div class="oc-name">${st.name}</div>
          <div class="oc-title">${titleOf(id)}</div>
          <div class="oc-unit">${id}</div>
        </div>${ch.length?`<ul>${ch.map(c=>node(c,lvl+1)).join('')}</ul>`:''}</li>`;
    };
    const roots=kids('');
    el.innerHTML=roots.length
      ? `<div class="oc-tree"><ul>${roots.map(r=>node(r,0)).join('')}</ul></div>`
      : '<div class="empty"><div class="empty-ico">🌳</div>No staff found</div>';
    const missed=people.filter(([i])=>!seen.has(i));
    if(missed.length)el.innerHTML+=`<div style="margin-top:1rem;font-size:.76rem;color:var(--text3)">⚠ Not shown (circular reporting line): ${missed.map(m=>m[1].name).join(', ')} — use ✏ to fix.</div>`;
  }
  orgZoom(dir){
    this._orgScale=dir===0?1:Math.min(1.6,Math.max(.5,(this._orgScale||1)+dir*0.12));
    const t=document.querySelector('#m-org-tree .oc-tree');
    if(t)t.style.transform='scale('+this._orgScale+')';
  }
  openOrgModal(id){
    const st=this.staff[id];if(!st)return;
    const o=(this._orgOverrides||{})[id]||{};
    $('og-id').value=id;
    $('og-staff-name').textContent=st.name+' ('+id+')';
    const cur=o.reports_to!==undefined&&o.reports_to!==null?o.reports_to:(st.supervisor||'');
    const opts=Object.entries(this.staff).filter(([i,s])=>s.role!=='admin'&&i!==id)
      .sort((a,b)=>a[1].name.localeCompare(b[1].name))
      .map(([i,s])=>`<option value="${i}" ${i===cur?'selected':''}>${s.name} (${s.unit||''})</option>`).join('');
    $('og-reports').innerHTML=`<option value="" ${!cur?'selected':''}>— Top level (no manager) —</option>`+opts;
    $('og-title').value=o.title||st.unit||'';
    $('og-order').value=o.sort_order??100;
    $('og-msg').textContent='';
    $('org-modal').classList.add('open');
  }
  async saveOrgNode(){
    const id=$('og-id').value;
    const reports=$('og-reports').value;
    if(reports===id)return $('og-msg').innerHTML='<span style="color:var(--red)">Someone cannot report to themselves.</span>';
    // walk up the chain to prevent a loop
    let cur=reports,hops=0;
    const oc=this._orgOverrides||{};
    while(cur&&hops<50){
      if(cur===id)return $('og-msg').innerHTML='<span style="color:var(--red)">That creates a circular reporting line.</span>';
      const o=oc[cur];
      cur=(o&&o.reports_to!==undefined&&o.reports_to!==null)?o.reports_to:(this.staff[cur]?.supervisor||'');
      hops++;
    }
    $('og-msg').innerHTML='<span style="color:var(--teal)">⏳ Saving…</span>';
    const r=await API._upsert('org_chart',[{staff_id:id,reports_to:reports,title:$('og-title').value.trim(),
      sort_order:parseInt($('og-order').value)||100,updated_at:new Date().toISOString()}]);
    if(r){closeModal('org-modal');toast('Organogram updated ✓');this.renderOrgChart('m-');}
    else $('og-msg').innerHTML='<span style="color:var(--red)">Save failed: '+(API.lastError||'')+'</span>';
  }
  openOrgUnassigned(){
    const list=Object.entries(this.staff).filter(([i,s])=>s.role!=='admin')
      .sort((a,b)=>a[1].name.localeCompare(b[1].name));
    if(!list.length)return;
    this.openOrgModal(list[0][0]);
    toast('Pick the staff member from the chart (✏) or adjust this one','info');
  }
  printOrgChart(){
    const tree=$('m-org-tree');if(!tree)return toast('Nothing to print','err');
    const css=[...document.querySelectorAll('style,link[rel=stylesheet]')].map(n=>n.outerHTML).join('');
    const w=window.open('','_blank');
    w.document.write(`<html><head><title>Organogram — THP-Ghana</title>${css}
      <style>body{background:#fff;padding:12mm;font-family:'Segoe UI',Arial,sans-serif}
      h1{font-size:15px;color:#2D3592;border-bottom:3px solid #2D3592;padding-bottom:8px}
      .meta{font-size:9px;color:#64748b;margin-bottom:14px}</style></head>
      <body data-theme="light"><h1>The Hunger Project — Ghana · Organogram</h1>
      <div class="meta">Generated ${new Date().toLocaleString('en-GB')} · by ${this.user.name}</div>
      <div class="oc-wrap">${tree.innerHTML}</div>
      <div class="no-print" style="margin-top:16px"><button onclick="window.print()" style="padding:8px 20px;background:#2D3592;color:#fff;border:none;border-radius:6px;cursor:pointer;font-weight:600">🖨 Print</button></div>
      </body></html>`);
    w.document.close();
  }

  /* ── Payroll (Finance: Ernest + Emmanuel; settings-driven) ── */
  _payDefaults(){return{ssnitTotalPct:18.5,ssnitEmployeePct:5.5,ssnitCeilingMonthly:69000,pfTotalPct:10,pfEmployeePct:5,
    bands:[{w:490,r:0},{w:110,r:.05},{w:130,r:.1},{w:3166.67,r:.175},{w:16000,r:.25},{w:30520,r:.3},{w:null,r:.35}]};}
  _maskAcct(v){
    const t=String(v||'').replace(/\s+/g,'');
    if(!t)return '';
    if(t.length<=6)return '•'.repeat(Math.max(4,t.length));           // too short to reveal safely
    if(t.length<=9)return t.slice(0,2)+'•'.repeat(t.length-5)+t.slice(-3);
    return t.slice(0,3)+'•'.repeat(t.length-6)+t.slice(-3);            // e.g. 210•••••••118
  }
  editBankAcct(){
    const el=$('pm-account');if(!el)return;
    el.value='';el.readOnly=false;el.placeholder='Enter the full account number';el.focus();
    this._acctMasked=false;
    const b=$('pm-acct-btn');if(b)b.style.display='none';
  }
  _ghs(n){return 'GH₵ '+(Number(n)||0).toLocaleString('en-GH',{minimumFractionDigits:2,maximumFractionDigits:2});}
  async _loadPaySettings(){
    if(this._payS)return this._payS;
    if(this._paySPromise)return this._paySPromise;   // avoid duplicate in-flight calls
    const r=await API.secureGet('payroll_settings','key=eq.main');
    let s=this._payDefaults();
    if(r&&r.length){try{const j=JSON.parse(r[0].value);
      if(j.extraReliefPct!==undefined&&j.pfMatchPct===undefined)j.pfMatchPct=j.extraReliefPct; // legacy key
      s={...s,...j};}catch(e){}}
    this._payS=s;return s;
  }
  async savePaySettings(p){
    const s=await this._loadPaySettings();
    s.ssnitTotalPct=+($(p+'ps-emp').value)||s.ssnitTotalPct;
    s.ssnitEmployeePct=+($(p+'ps-empr').value)||s.ssnitEmployeePct;
    s.ssnitCeilingMonthly=+($(p+'ps-ceil').value)||s.ssnitCeilingMonthly;
    s.pfTotalPct=+($(p+'ps-t3').value)||s.pfTotalPct;
    s.pfEmployeePct=+($(p+'ps-relief').value)||0;
    this._payS=s;
    const r=await API.secureSave('payroll_settings',[{key:'main',value:JSON.stringify(s)}]);
    if(r){toast('Settings saved ✓');this.renderPayroll(p);}else toast('Save failed: '+(API.lastError||'unknown'),'err');
  }
  _calcPay(ps,S){
    const n=v=>+v||0;
    let allow=[];try{allow=JSON.parse(ps.allowances||'[]');}catch(e){}
    const basic=n(ps.basic);
    const arrears=n(ps.arrears),incent=n(ps.incentives),bonus=n(ps.bonus),ot=n(ps.overtime),fuel=n(ps.fuel_allowance);
    const taxA=allow.filter(a=>a.tax).reduce((t,a)=>t+n(a.a),0);
    const nonTax=allow.filter(a=>!a.tax).reduce((t,a)=>t+n(a.a),0);
    // Cash earnings that attract tax alongside basic
    const taxableExtras=arrears+incent+bonus+ot+taxA;
    const gross=basic+arrears+incent+bonus+ot+fuel+taxA+nonTax;
    const capped=Math.min(basic,S.ssnitCeilingMonthly);
    // SSNIT: 18.5% total — 5.5% deducted from the employee, the rest paid by THP.
    const ssnitTotal=(+S.ssnitTotalPct||18.5)/100*capped;
    const ssnitEmp=(+S.ssnitEmployeePct||5.5)/100*capped;
    // Provident Fund: 10% total — 5% deducted from the employee, the rest paid by THP.
    const pfTotal=(+S.pfTotalPct||10)/100*basic;
    const tier3=(+S.pfEmployeePct||5)/100*basic;
    // Total tax relief allowed on basic, as a single percentage.
    // Set it so PAYE matches the payslip: 10% of basic for THP
    // (5% SSNIT-side + 5% provident fund) is already covered by
    // ssnitEmp + tier3; reliefPct adds any further relief on basic.
    // Tax relief: the full SSNIT employee share plus the FULL provident fund
    // (both portions), which is what reconciles to the THP payslip.
    let taxable=Math.max(0,basic-ssnitEmp-pfTotal+taxableExtras);
    let paye=0,rem=taxable;
    for(const b of S.bands){const chunk=b.w===null?rem:Math.min(rem,b.w);paye+=chunk*b.r;rem-=chunk;if(rem<=0)break;}
    if(ps.paye_override!==null&&ps.paye_override!==undefined&&ps.paye_override!=='')paye=n(ps.paye_override);
    const advance=n(ps.salary_advance),ug=n(ps.ug_credit),other=n(ps.other_deductions);
    const totalDed=ssnitEmp+tier3+paye+advance+ug+other;
    const net=gross-totalDed;
    const emprSSNIT=ssnitTotal-ssnitEmp;
    return{gross,taxA,nonTax,arrears,incent,bonus,ot,fuel,ssnitEmp,ssnitTotal,tier3,pfTotal,paye,advance,ug,other,totalDed,net,
      cost:gross+emprSSNIT+(pfTotal-tier3),
      allowStr:allow.map(a=>`${a.n} ${this._ghs(a.a)}${a.tax?'':' (nt)'}`).join(', ')||'—'};
  }
  async renderPayroll(p){
    const body=$(p+'pay-body');if(!body)return;
    body.innerHTML='<tr><td colspan="10" style="color:var(--text3)">Loading…</td></tr>';
    // One round trip each to Apps Script is slow; run them together.
    const [S,rowsRaw]=await Promise.all([
      this._loadPaySettings(),
      API.secureGet('payroll_staff','select=*')
    ]);
    $(p+'ps-emp').value=S.ssnitTotalPct;$(p+'ps-empr').value=S.ssnitEmployeePct;
    $(p+'ps-ceil').value=S.ssnitCeilingMonthly;$(p+'ps-t3').value=S.pfTotalPct;
    if($(p+'ps-relief'))$(p+'ps-relief').value=S.pfEmployeePct??5;
    if(!$(p+'pay-month').value)$(p+'pay-month').value=new Date().toISOString().slice(0,7);
    const ap=$(p+'ps-applied');
    if(ap)ap.innerHTML='<strong>Rates currently applied:</strong> '
      +'SSNIT <strong>'+S.ssnitTotalPct+'%</strong> total, of which <strong>'+S.ssnitEmployeePct+'%</strong> is deducted from the employee · '
      +'Provident Fund <strong>'+S.pfTotalPct+'%</strong> total, of which <strong>'+S.pfEmployeePct+'%</strong> is deducted · '
      +'SSNIT ceiling <strong>'+this._ghs(S.ssnitCeilingMonthly)+'</strong>/month.';
    this.renderPayStatus(p);
    const rows=rowsRaw||[];
    this._payRows=rows;const payMap={};rows.forEach(r=>payMap[r.staff_id]=r);
    const list=Object.entries(this.staff).filter(([i,s])=>s.role!=='admin').sort((a,b)=>a[1].name.localeCompare(b[1].name));
    let tot={gross:0,ssnitEmp:0,tier3:0,paye:0,net:0,cost:0};
    this._payCalc=[];
    body.innerHTML=list.map(([id,s])=>{
      const ps=payMap[id];
      if(!ps||!+ps.basic)return`<tr><td><strong>${s.name}</strong><br><span style="font-size:.72rem;color:var(--text3)">${id}</span></td><td colspan="8" style="color:var(--text3);font-size:.78rem">No salary set</td><td><button class="bsm bsm-navy" onclick="APP.openPayModal('${id}')">✏ Setup</button></td></tr>`;
      const c=this._calcPay(ps,S);
      Object.keys(tot).forEach(k=>tot[k]+=c[k]);
      let alloc=[];try{alloc=JSON.parse(ps.cost_allocation||'[]');}catch(e){}
      let allowList=[];try{allowList=JSON.parse(ps.allowances||'[]');}catch(e){}
      this._payCalc.push({id,name:s.name,unit:s.unit,email:s.email||'',designation:ps.designation||'',...c,basic:+ps.basic,alloc,allowances:allowList});
      return`<tr><td><strong>${s.name}</strong><br><span style="font-size:.72rem;color:var(--text3)">${id}</span></td><td>${this._ghs(ps.basic)}</td><td style="font-size:.72rem">${c.allowStr}</td><td>${this._ghs(c.gross)}</td><td>${this._ghs(c.ssnitEmp)}</td><td>${this._ghs(c.tier3)}</td><td>${this._ghs(c.paye)}</td><td><strong>${this._ghs(c.net)}</strong></td><td>${this._ghs(c.cost)}</td><td><button class="bsm bsm-navy" onclick="APP.openPayModal('${id}')">✏</button></td></tr>`;
    }).join('');
    const sm=$(p+'pay-summary');
    if(sm)sm.innerHTML=`<div class="cs-box"><div class="cs-num" style="font-size:1rem">${this._ghs(tot.gross)}</div><div class="cs-lbl">Gross</div></div>
      <div class="cs-box"><div class="cs-num" style="font-size:1rem">${this._ghs(tot.paye)}</div><div class="cs-lbl">PAYE</div></div>
      <div class="cs-box"><div class="cs-num" style="font-size:1rem">${this._ghs(tot.ssnitEmp)}</div><div class="cs-lbl">SSNIT (Emp)</div></div>
      <div class="cs-box"><div class="cs-num" style="font-size:1rem;color:var(--green)">${this._ghs(tot.net)}</div><div class="cs-lbl">Net Payout</div></div>
      <div class="cs-box"><div class="cs-num" style="font-size:1rem">${this._ghs(tot.cost)}</div><div class="cs-lbl">Employer Cost</div></div>`;
  }
  /* ── Allowance builder: name · amount · taxable ── */
  _renderAllowRows(){
    const el=$('pm-allow-rows');if(!el)return;
    el.innerHTML=(this._allowList||[]).map((x,i)=>`
      <div class="allow-row">
        <input class="fi an" value="${(x.n||'').replace(/"/g,'&quot;')}" placeholder="Allowance name" oninput="APP._allowSet(${i},'n',this.value)">
        <input class="fi aa" type="number" step="0.01" value="${+x.a||0}" oninput="APP._allowSet(${i},'a',this.value)">
        <select class="fi at" onchange="APP._allowSet(${i},'tax',this.value==='t')">
          <option value="t" ${x.tax?'selected':''}>Taxable</option>
          <option value="n" ${x.tax?'':'selected'}>Non-taxable</option>
        </select>
        <button class="bsm" style="background:var(--surf);border:1px solid var(--border);color:var(--text2)" onclick="APP._allowDel(${i})">✕</button>
      </div>`).join('')
      ||'<div style="font-size:.76rem;color:var(--text3);margin-bottom:.4rem">No allowances added.</div>';
    $('pm-allow').value=JSON.stringify(this._allowList||[]);
  }
  _allowSet(i,f,v){this._allowList[i][f]=(f==='a')?(+v||0):v;$('pm-allow').value=JSON.stringify(this._allowList);}
  _allowDel(i){this._allowList.splice(i,1);this._renderAllowRows();}
  addAllowRow(){(this._allowList=this._allowList||[]).push({n:'',a:0,tax:false});this._renderAllowRows();}

  async openPayModal(id){
    const s=this.staff[id];if(!s)return;
    const ps=(this._payRows||[]).find(r=>r.staff_id===id);
    let allow=[];try{allow=JSON.parse(ps?.allowances||'[]');}catch(e){}
    $('pm-id').value=id;
    $('pm-staff-name').textContent=s.name+' ('+id+')';
    $('pm-basic').value=ps?.basic||'';
    $('pm-tier3').value=ps?.tier3_pct||0;
    $('pm-grade').value=ps?.grade||'junior';
    $('pm-bank').value='';$('pm-account').value='';
    this._acctMasked=false;this._acctReal='';
    Promise.resolve(ps||{}).then(hf=>{
      if($('pm-id').value!==id)return;
      $('pm-bank').value=hf?.bank_name||'';
      const acct=hf?.bank_account||'';
      this._acctReal=acct;
      const el=$('pm-account'),btn=$('pm-acct-btn');
      if(acct){el.value=this._maskAcct(acct);el.readOnly=true;this._acctMasked=true;if(btn)btn.style.display='inline-block';}
      else{el.value='';el.readOnly=false;this._acctMasked=false;if(btn)btn.style.display='none';}
    });
    this._allowList=allow.slice();
    const legacyFuel=+(ps?.fuel_allowance)||0;
    if(legacyFuel>0&&!this._allowList.some(x=>/fuel/i.test(x.n||'')))
      this._allowList.push({n:'Fuel Allowance',a:legacyFuel,tax:false});
    this._renderAllowRows();
    $('pm-designation').value=ps?.designation||'';
    ['arrears','incentives','bonus','overtime','fuel','advance','ug','other'].forEach(k=>{
      const map={arrears:'arrears',incentives:'incentives',bonus:'bonus',overtime:'overtime',
        fuel:'fuel_allowance',advance:'salary_advance',ug:'ug_credit',other:'other_deductions'};
      const el=$('pm-'+k);if(el)el.value=ps?.[map[k]]||0;});
    $('pm-paye-ov').value=(ps?.paye_override===null||ps?.paye_override===undefined)?'':ps.paye_override;
    let alloc=[];try{alloc=JSON.parse(ps?.cost_allocation||'[]');}catch(e){}
    $('pm-alloc').value=alloc.map(x=>`${x.project} : ${x.pct}`).join('\n');
    $('pm-msg').textContent='';
    $('pay-modal').classList.add('open');
  }
  async savePayStaff(){
    const id=$('pm-id').value;
    const allow=(this._allowList||[]).filter(x=>(x.n||'').trim()&&(+x.a||0)!==0)
      .map(x=>({n:x.n.trim(),a:+x.a||0,tax:!!x.tax}));
    const alloc=$('pm-alloc').value.split('\n').map(l=>l.trim()).filter(Boolean).map(l=>{
      const q=l.split(':').map(x=>x.trim());return{project:q[0]||'Unallocated',pct:+q[1]||0};});
    const totPct=alloc.reduce((t,x)=>t+x.pct,0);
    if(alloc.length&&Math.abs(totPct-100)>0.01)return $('pm-msg').innerHTML='<span style="color:var(--red)">Allocation must total 100% (currently '+totPct+'%).</span>';
    const ov=$('pm-paye-ov').value.trim();
    const r=await API.secureSave('payroll_staff',[{staff_id:id,basic:+$('pm-basic').value||0,allowances:JSON.stringify(allow),
      tier3_pct:+$('pm-tier3').value||0,grade:$('pm-grade').value,cost_allocation:JSON.stringify(alloc),
      designation:$('pm-designation').value.trim(),
      arrears:+$('pm-arrears').value||0,incentives:+$('pm-incentives').value||0,bonus:+$('pm-bonus').value||0,
      overtime:+$('pm-overtime').value||0,fuel_allowance:0,
      salary_advance:+$('pm-advance').value||0,ug_credit:+$('pm-ug').value||0,other_deductions:+$('pm-other').value||0,
      paye_override:ov===''?null:+ov,
      bank_name:$('pm-bank').value.trim(),
      bank_account:this._acctMasked?this._acctReal:$('pm-account').value.trim(),
      updated_at:new Date().toISOString()}]);
    if(r){

      closeModal('pay-modal');toast('Pay setup saved ✓');this.renderPayroll('m-');this.renderPayroll('st-');
    }
    else $('pm-msg').innerHTML='<span style="color:var(--red)">Save failed — '+(API.lastError||'reason unknown')+'</span>';
  }
  /* ── Phase B: bank advice, statutory returns, allocation, payslips ── */
  _payGuard(){
    if(!this._payCalc||!this._payCalc.length){toast('Recalculate the payroll first','err');return false;}
    return true;
  }
  async exportBankAdvice(p){
    if(!this._requireApproval())return;
    if(!this._payGuard())return;
    const month=$(p+'pay-month').value||'';
    const files=await API.secureGet('payroll_staff','select=staff_id,bank_name,bank_account')||[];
    const bk={};files.forEach(f=>bk[f.staff_id]=f);
    let csv='THP-GHANA BANK ADVICE,'+month+'\nStaff ID,Name,Bank,Account Number,Net Pay (GHS)\n';
    let tot=0,missing=[];
    this._payCalc.forEach(r=>{
      const b=bk[r.id]||{};
      if(!b.bank_account)missing.push(r.name);
      tot+=r.net;
      csv+=`"${r.id}","${r.name}","${b.bank_name||''}","${b.bank_account||''}",${r.net.toFixed(2)}\n`;
    });
    csv+=`,,,TOTAL,${tot.toFixed(2)}\n`;
    this._dl(csv,'THP_Bank_Advice_'+month+'.csv','text/csv');
    if(missing.length)toast('No bank account on file for: '+missing.join(', '),'info');
  }
  async exportStatutory(p){
    if(!this._payGuard())return;
    const S=await this._loadPaySettings();
    const month=$(p+'pay-month').value||'';
    let csv='THP-GHANA STATUTORY RETURNS,'+month+'\n';
    csv+='Staff ID,Name,Basic (GHS),SSNIT Employee 5.5%,SSNIT Employer 13%,Tier 1 (13.5%),Tier 2 (5%),Provident Fund,PAYE (GHS)\n';
    let t={emp:0,empr:0,t1:0,t2:0,t3:0,paye:0};
    this._payCalc.forEach(r=>{
      const capped=Math.min(r.basic,S.ssnitCeilingMonthly);
      const empr=S.ssnitEmployerPct/100*capped;
      const total=r.ssnitEmp+empr;          // 18.5% combined
      const t1=total*(13.5/18.5), t2=total*(5/18.5);
      t.emp+=r.ssnitEmp;t.empr+=empr;t.t1+=t1;t.t2+=t2;t.t3+=r.tier3;t.paye+=r.paye;
      csv+=`"${r.id}","${r.name}",${r.basic.toFixed(2)},${r.ssnitEmp.toFixed(2)},${empr.toFixed(2)},${t1.toFixed(2)},${t2.toFixed(2)},${r.tier3.toFixed(2)},${r.paye.toFixed(2)}\n`;
    });
    csv+=`,TOTALS,,${t.emp.toFixed(2)},${t.empr.toFixed(2)},${t.t1.toFixed(2)},${t.t2.toFixed(2)},${t.t3.toFixed(2)},${t.paye.toFixed(2)}\n`;
    csv+='\nNOTE,"Tier 1 to SSNIT; Tier 2 to a licensed private trustee; PAYE to GRA. Confirm current rates before filing."\n';
    this._dl(csv,'THP_Statutory_Returns_'+month+'.csv','text/csv');
  }
  exportAllocation(p){
    if(!this._payGuard())return;
    const month=$(p+'pay-month').value||'';
    const proj={};
    let csv='THP-GHANA PAYROLL COST ALLOCATION,'+month+'\nStaff ID,Name,Project/Grant,Share %,Allocated Cost (GHS)\n';
    let unalloc=0;
    this._payCalc.forEach(r=>{
      const list=(r.alloc&&r.alloc.length)?r.alloc:[{project:'Unallocated',pct:100}];
      if(!r.alloc||!r.alloc.length)unalloc++;
      list.forEach(a=>{
        const amt=r.cost*a.pct/100;
        proj[a.project]=(proj[a.project]||0)+amt;
        csv+=`"${r.id}","${r.name}","${a.project}",${a.pct},${amt.toFixed(2)}\n`;
      });
    });
    csv+='\nPROJECT/GRANT,TOTAL EMPLOYER COST (GHS)\n';
    Object.entries(proj).sort((a,b)=>b[1]-a[1]).forEach(([k,v])=>{csv+=`"${k}",${v.toFixed(2)}\n`;});
    this._dl(csv,'THP_Cost_Allocation_'+month+'.csv','text/csv');
    if(unalloc)toast(unalloc+' staff have no allocation set — shown as "Unallocated"','info');
  }
  async emailPayslips(p){
    if(!this._requireApproval())return;
    if(!this._payGuard())return;
    const month=this._payMonthLabel(p);
    const withEmail=this._payCalc.filter(r=>r.email);
    const without=this._payCalc.filter(r=>!r.email);

    // ── Safety check: two staff must never share an email address,
    //    or one person would receive the other's payslip. ──
    const byEmail={};
    withEmail.forEach(r=>{const e=r.email.trim().toLowerCase();(byEmail[e]=byEmail[e]||[]).push(r.name);});
    const dupes=Object.entries(byEmail).filter(([e,names])=>names.length>1);
    if(dupes.length){
      const detail=dupes.map(([e,names])=>'• '+e+'  →  '+names.join(' AND ')).join('\n');
      alert('PAYSLIPS NOT SENT\n\nThese staff share the same email address, so one would receive '
        +'another person\'s payslip:\n\n'+detail
        +'\n\nPlease correct the email addresses in Staff records, then try again.');
      this.audit('Payslip send blocked — duplicate emails','Payroll',month,detail.replace(/\n/g,' | '));
      return;
    }
    // Warn if an address does not look valid
    const bad=withEmail.filter(r=>!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(r.email.trim()));
    if(bad.length&&!confirm('These addresses do not look valid:\n\n'
      +bad.map(r=>'• '+r.name+' — '+r.email).join('\n')
      +'\n\nSend anyway?'))return;
    if(!withEmail.length)return toast('No staff have email addresses on file','err');
    if(!confirm('Release payslips for '+month+'?\n\n'+withEmail.length+' staff will be notified by email that their payslip is ready to view in their portal.\n\nEach person can only open their own.'+(without.length?'\n\nNo email on file for: '+without.map(r=>r.name).join(', '):'')))return;
    showLoader('Sending payslips…');
    // Payslip password: first name (lowercase) + date of birth ddmmyyyy
    // e.g. eric14091990 — private to the individual, unlike a Staff ID.
    const files=await API._get('hr_staff_files','select=staff_id,dob')||[];
    const dobMap={};files.forEach(x=>{if(x.dob)dobMap[x.staff_id]=String(x.dob).slice(0,10);});
    const noDob=withEmail.filter(r=>!dobMap[r.id]);
    if(noDob.length&&!confirm('These staff have no date of birth on file, so their payslip '
      +'cannot be password-protected:\n\n'+noDob.map(r=>'• '+r.name).join('\n')
      +'\n\nAdd their date of birth in Staff Files first, or continue and send theirs unprotected?'))return;
    const pwFor=r=>{
      const d=dobMap[r.id];if(!d)return '';
      const [y,m,dd]=d.split('-');
      const first=String(r.name||'').trim().split(/\s+/)[0].toLowerCase().replace(/[^a-z]/g,'');
      return first+dd+m+y;
    };
    const payload=[];
    for(const r of withEmail){
      payload.push({name:r.name,id:r.id,email:r.email,unit:r.unit||'',month,
        net:r.net.toFixed(2),pw:pwFor(r),html:await this._payslipHTML(r,month)});
    }
    const res=await API.gasPost({action:'sendPayslips',month,slips:payload});
    hideLoader();
    if(res&&res.success)toast('Payslips sent: '+res.sent+(res.failed?(' · failed: '+res.failed):'')+' ✓');
    else toast('Payslip sending failed'+(res&&res.error?': '+res.error:' — check the Apps Script deployment'),'err');
  }
  get _logo(){return 'data:image/png;base64,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b7F64GBx2vSp+/YeWnbZoqamXKEx29b29sQJkydNGre5bWs+yytvvvKyj537qU+fe86imatWvfT6q9uXffSsK68568PnzNuwYcv8s6bP+9DMKJJ8IffGph3G1LGPlOrRyc4op+AWj+YXIHWnUqh60M9OO73gZ4ZSCacJkEJ8Q571auVyGGQy1hgCR7WKcpoNslGoYRQ3t/ilYihicznf80hEqpU0G2TJaBJzpRIvXjT5775744ov3PNWW/tFF0+8554vXXXV377zdrcfhGPG+V3tlMReU3OmEhaNoWyQq8VQJZDWnRJQI0gtONT01Ky3YnLihEAgOGhO9OeXXDsp32CHUlnJgSKQqKbiKNeQd84myqoJvBwrwlCtCQoNGoVRNtsKJVGp1UDQTDYXO2WF05rxTWdPJBbnXTCjqaVw2eWLn3jytT17u1paRolWuzsi6xVsJomS1HoFQCtRYuoEFoAOVQZpELWggdAwEIkJAKgGiglWHBd7K84Z5ZJoam2+oTnDSNlYh5gUbI3TtD5jfeivKRmGc6NHB6Vqdev29k8v/78Lzp02eeq41atfeXlDm5fJplpRwHpZEacaExsVByJjbP9j0Qj6w/EVNAATgjRxra3m5luvCDyfKCZjX31155o1rxkvX58mVQXqCldijOdSIoI1XqlYPP/cCf/6wO1/9e2fPHD/G51d+m+r1oujIJPP5loECRsnYkRhLJFrSl2FWTybSROriJWE/ns1MYBE4aw1XZ3FbdvevnHFRyQxG9a3/b97brjt1uW9XWmxu9rTWe7tDks9pagEpA1dBytpWo4q1d7O3qxf2LTx4E0r7nnql20Nza4S7/OMbSoU0iQVxEReb1ccVStJXO3t6oqisrryfffffNLYfLHYyYb664FBAoUek08MpRIMEiViSRP7ztZuAPsPVR5++KUv/9k1S5bM/tV/rP+j267dsWtXtYbLl511/71Pr1699s47rhw/ZXTGZJ59YuOjj7918x9fesq0UT/7t02P/eqVS5bO+dznl7FJN7ftuf++X8VS/sy1H75w6Xw1urVt908fefHP77j5nLMn3/G1a37w4Jq1a7fmsllVGVDlvBcdGmYFlPpLewSBDyCNklGjMpPGNx882P3aK3uqYfWG6y/Zu3dfGNdyjXLtZ8+5aeWl3/3WY4/8Yt3d375uzmlNL67b8DsXz58y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';}

  /* Official THP-Ghana stamp — applied to APPROVED payslips only */
  get _stamp(){return 'data:image/png;base64,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';}

  /* ── Payslip in THP format ── */
  /* Organisation details, from Settings rather than hardcoded */
  _org(k,d){return (this.settings&&this.settings[k])||d;}
  async _payslipHTML(r,month){
    const approved=!!this._payApproved;
    const f=v=>(+v||0).toLocaleString('en-GH',{minimumFractionDigits:2,maximumFractionDigits:2});
    const dash=v=>(+v)?f(v):'-';
    let bank={};
    try{const rows=await API.secureGet('payroll_staff','select=bank_name,bank_account&staff_id=eq.'+encodeURIComponent(r.id));
        bank=(rows&&rows[0])||{};}catch(e){}
    const s=this.staff[r.id]||{};
    const row=(l,v)=>`<tr><td class="lbl">${l}</td><td class="cur">GH₵</td><td class="amt">${v}</td></tr>`;
    return `<html><head><meta charset="UTF-8"><style>
      body{font-family:Arial,Helvetica,sans-serif;color:#222;font-size:11px;margin:0;padding:16px}
      .sheet{border:1px solid #bbb;padding:0 0 18px}
      .top{background:#efefef;padding:14px 18px;display:flex;justify-content:space-between;align-items:flex-start}
      .org{font-size:13px;font-weight:bold;letter-spacing:.3px}
      .addr{font-size:9px;font-style:italic;color:#444;margin-top:2px}
      .mail{font-size:9px;color:#1155cc;text-decoration:underline}
      .ptitle{font-size:22px;font-weight:bold;letter-spacing:6px;color:#555;text-align:right}
      .bar{height:16px;background:#4a4a4a;margin-bottom:14px}
      .pad{padding:0 18px}
      .net{text-align:right;font-size:12px;font-weight:bold;margin-bottom:10px}
      .net b{font-size:17px}
      .cols{display:flex;gap:22px}
      .col{flex:1}
      table{width:100%;border-collapse:collapse}
      .det td{border:1px solid #d5d5d5;padding:4px 7px;font-size:10px;background:#fafafa}
      .det td.k{width:44%;color:#333}
      .hd th{background:#efefef;border:1px solid #d5d5d5;padding:5px 7px;font-size:10.5px;font-weight:bold;text-align:left}
      .hd th.r{text-align:right}
      .ln td{border-bottom:1px solid #e8e8e8;padding:4px 7px;font-size:10px}
      .ln td.lbl{font-style:italic}
      .ln td.cur{width:34px;color:#555}
      .ln td.amt{text-align:right;width:78px}
      .tot td{padding:7px;font-size:11px;font-weight:bold;border-top:1px solid #999}
      .tot td.amt{text-align:right}
      .sig{display:flex;gap:22px;margin-top:34px;padding:0 18px}
      .sig div{flex:1;font-size:10px;font-style:italic;font-weight:bold}
      .sig span{display:inline-block;border-bottom:1px solid #777;width:58%;margin-left:6px}
    </style></head><body><div class="sheet">
      <div class="top">
        <div><div class="org">${this._org('org_name','The Hunger Project — Ghana').toUpperCase()}</div>
          <div class="addr">${this._org('org_address','PMB CT 7, Cantonments Accra, Ghana')}</div>
          <div class="mail">email: ${this._org('org_email','thpghana@thp.org')}</div></div>
        <div style="text-align:right">
          <img src="${this._logo}" alt="THP" style="width:74px;height:auto;display:inline-block;margin-bottom:4px">
          <div class="ptitle">PAYSLIP</div>
        </div>
      </div>
      <div class="bar"></div>
      <div class="pad">
        <div class="net">Net Pay: &nbsp; GH₵ &nbsp; <b>${f(r.net)}</b></div>
        <div class="cols">
          <div class="col"><table class="det">
            <tr><td class="k">Employee Name :</td><td>${r.name}</td></tr>
            <tr><td class="k">Employee ID :</td><td>${r.id}</td></tr>
            <tr><td class="k">E-mail ID</td><td>${r.email||''}</td></tr>
            <tr><td class="k">Contact No :</td><td>${bank.phone||s.phone||''}</td></tr>
          </table></div>
          <div class="col"><table class="det">
            <tr><td class="k">Department :</td><td>${r.unit||''}</td></tr>
            <tr><td class="k">Designation :</td><td>${r.designation||''}</td></tr>
            <tr><td class="k">Bank Account No.</td><td>${this._maskAcct(bank.bank_account)}</td></tr>
            <tr><td class="k">Pay Period:</td><td>${month}</td></tr>
            <tr><td class="k">Generated:</td><td>${new Date().toLocaleString('en-GB',{day:'2-digit',month:'short',year:'numeric',hour:'2-digit',minute:'2-digit'})}</td></tr>
          </table></div>
        </div>
        <div class="cols" style="margin-top:16px">
          <div class="col"><table>
            <tr class="hd"><th>EARNINGS</th><th></th><th class="r">AMOUNT</th></tr>
            ${row('Basic Salary',f(r.basic))}${row('Arrears',dash(r.arrears))}${row('Incentives',dash(r.incent))}
            ${row('Bonus',dash(r.bonus))}${row('Over Time Pay',dash(r.ot))}
            ${(r.allowances||[]).map(a=>row(a.n,f(a.a))).join('')}
          </table></div>
          <div class="col"><table>
            <tr class="hd"><th>DEDUCTIONS</th><th></th><th class="r">AMOUNT</th></tr>
            ${row('Provident Fund',dash(r.tier3))}${row('SSNIT',f(r.ssnitEmp))}${row('PAYE',f(r.paye))}
            ${row('Salary Advance',dash(r.advance))}${row('UG Credit',dash(r.ug))}${row('Others Deductions',dash(r.other))}
          </table></div>
        </div>
        <div class="cols" style="margin-top:20px">
          <div class="col"><table><tr class="tot"><td>Gross Salary</td><td>GH₵</td><td class="amt">${f(r.gross)}</td></tr></table></div>
          <div class="col"><table>
            <tr class="tot"><td>Total Deductions</td><td>GH₵</td><td class="amt">${f(r.totalDed)}</td></tr>
            <tr class="tot"><td style="text-align:right">NET Salary</td><td>GH₵</td><td class="amt">${f(r.net)}</td></tr>
          </table></div>
        </div>
      </div>
      <div class="sig">
        <div>Employee Signature :<span></span></div>
        <div style="position:relative">Employer Signature :<span></span>
          ${approved?`<img src="${this._stamp}" alt="" style="position:absolute;right:6%;bottom:-6px;width:150px;opacity:.92">`:''}
        </div>
      </div>
    </div></body></html>`;
  }
  async previewPayslip(p){
    if(!this._payGuard())return;
    // Open synchronously with the click — once we await, the browser
    // treats a new window as an unsolicited popup and blocks it.
    const w=window.open('','_blank');
    if(!w){toast('Your browser blocked the preview window. Allow pop-ups for this site, then try again.','err');return;}
    w.document.write('<html><head><meta charset="UTF-8"><title>Payslips — THP-Ghana</title></head>'
      +'<body style="font-family:Arial,sans-serif;padding:40px;text-align:center;color:#2D3592">'
      +'<h3>Preparing payslips…</h3><p style="color:#64748b">Please wait.</p></body></html>');
    const month=this._payMonthLabel(p);
    const list=this._payCalc||[];
    try{
      const pages=[];
      for(const r of list){pages.push(await this._payslipHTML(r,month));}
      const bodies=pages.map(h=>{
        const m=h.match(/<body>([\s\S]*)<\/body>/i);
        return '<div class="slip-page">'+(m?m[1]:h)+'</div>';
      }).join('');
      const head=pages[0]?pages[0].match(/<head>([\s\S]*)<\/head>/i):null;
      w.document.open();
      w.document.write('<html><head>'+(head?head[1]:'')
        +'<style>.slip-page{page-break-after:always;margin-bottom:26px}'
        +'.slip-page:last-child{page-break-after:auto}'
        +'@media screen{.slip-page{box-shadow:0 2px 14px rgba(0,0,0,.12);margin:0 auto 26px;max-width:960px}}'
        +'</style></head><body>'+bodies
        +'<div style="text-align:center;padding:16px" class="no-print">'
        +'<button onclick="window.print()" style="padding:10px 26px;background:#2D3592;color:#fff;border:0;border-radius:8px;font-weight:600;cursor:pointer">🖨 Print / Save all as PDF</button></div>'
        +'</body></html>');
      w.document.close();
      toast(list.length+' payslip(s) ready — one per page','info');
    }catch(e){
      try{w.document.open();w.document.write('<p style="font-family:Arial;padding:40px;color:#dc2626">Could not build the payslips: '+(e.message||e)+'</p>');w.document.close();}catch(_){}
      toast('Preview failed: '+(e.message||e),'err');
    }
  }
  _payMonthLabel(p){
    const v=$(p+'pay-month').value||new Date().toISOString().slice(0,7);
    const d=new Date(v+'-01');
    return d.toLocaleString('en',{month:'short'})+'-'+d.getFullYear();
  }
  /* ── Payroll approval workflow ──
     Draft → Submitted → Approved. Exports and payslips are locked
     until the Country Leader approves. Ernest and Emmanuel prepare
     and submit; Agatha approves or returns. ── */
  get PAY_PREPARERS(){return['THPG/05/2025','THPG/01/2026-3','ADMIN01'];}
  get PAY_APPROVERS(){return['THPG/12/2024','ADMIN01'];}
  _canPrepare(){return this.PAY_PREPARERS.includes(this.user?.id);}
  _canApprove(){return this.PAY_APPROVERS.includes(this.user?.id);}
  _runId(p){return 'RUN-'+($(p+'pay-month').value||new Date().toISOString().slice(0,7));}
  async _loadRun(p){
    const rows=await API.secureGet('payroll_runs','id=eq.'+encodeURIComponent(this._runId(p)));
    return (rows&&rows[0])||null;
  }
  async renderPayStatus(p){
    const el=$(p+'pay-status');if(!el)return;
    const sub0=$(p+'pay-submit'),app0=$(p+'pay-approve'),ret0=$(p+'pay-return');
    // Hide all three first — only the role-appropriate ones are shown below.
    [sub0,app0,ret0].forEach(b=>{if(b)b.style.display='none';});
    const run=await this._loadRun(p);
    this._curRun=run;
    const sub=$(p+'pay-submit'),app=$(p+'pay-approve'),ret=$(p+'pay-return');
    const st=run?.status||'None';
    const show=(e,v)=>{if(e)e.style.display=v?'inline-block':'none';};
    if(!run){
      el.style.display='none';
      show(sub,this._canPrepare());show(app,false);show(ret,false);
      return;
    }
    el.style.display='block';
    const badge={Draft:'none',Submitted:'amber',Approved:'green',Returned:'red'}[st]||'none';
    let line=`<span class="c-flag ${badge}">${st}</span>`;
    if(run.prepared_by)line+=` &nbsp;Prepared by <strong>${run.prepared_by}</strong>`;
    if(run.submitted_by)line+=` &nbsp;· Submitted by <strong>${run.submitted_by}</strong> ${run.submitted_at?String(run.submitted_at).slice(0,10):''}`;
    if(run.approved_by)line+=` &nbsp;· Approved by <strong>${run.approved_by}</strong> ${run.approved_at?String(run.approved_at).slice(0,10):''}`;
    if(run.approver_note)line+=`<br><span style="color:var(--red)">Note: ${run.approver_note}</span>`;
    el.innerHTML=line;
    show(sub,this._canPrepare()&&(st==='Draft'||st==='Returned'));
    show(app,this._canApprove()&&st==='Submitted');
    show(ret,this._canApprove()&&st==='Submitted');
    // Lock distribution until approved
    this._payApproved=(st==='Approved');
  }
  _requireApproval(){
    if(this._payApproved)return true;
    toast('This payroll run must be approved by the Country Leader first','err');
    return false;
  }
  async savePayrollRun(p){
    if(!this._canPrepare())return toast('Only Finance can prepare payroll','err');
    if(!this._payCalc||!this._payCalc.length)return toast('Recalculate first','err');
    const month=$(p+'pay-month').value||new Date().toISOString().slice(0,7);
    const cur=await this._loadRun(p);
    if(cur&&cur.status==='Approved')return toast('This run is already approved and cannot be changed','err');
    if(cur&&cur.status==='Submitted')return toast('This run is awaiting approval — it cannot be edited','err');
    const tot=this._payTotals();
    const r=await API.secureSave('payroll_runs',[{id:'RUN-'+month,month,
      data:JSON.stringify(this._payCalc),totals:JSON.stringify(tot),
      status:cur?.status==='Returned'?'Returned':'Draft',
      prepared_by:this.user.name,prepared_at:new Date().toISOString(),
      processed_at:new Date().toISOString()}]);
    if(r){this.audit('Payroll draft saved','Payroll',month,this._payCalc.length+' staff');
      toast('Draft saved for '+month+' ✓');this.renderPayStatus(p);}
    else toast('Save failed: '+(API.lastError||'unknown'),'err');
  }
  _payTotals(){
    const t={gross:0,ssnitEmp:0,tier3:0,paye:0,net:0,cost:0};
    (this._payCalc||[]).forEach(r=>Object.keys(t).forEach(k=>t[k]+=(+r[k]||0)));
    t.staff=(this._payCalc||[]).length;
    return t;
  }
  async submitPayrollRun(p){
    if(!this._canPrepare())return toast('Only Finance can submit payroll','err');
    if(!this._payCalc||!this._payCalc.length)return toast('Recalculate first','err');
    const month=$(p+'pay-month').value;
    const t=this._payTotals();
    if(!confirm('Submit payroll for '+month+' to the Country Leader?\n\n'
      +t.staff+' staff · Net payout '+this._ghs(t.net)+'\n\nOnce submitted it cannot be edited until it is approved or returned.'))return;
    await this.savePayrollRun(p);
    const r=await API.secureSave('payroll_runs',[{id:'RUN-'+month,month,
      status:'Submitted',submitted_by:this.user.name,submitted_at:new Date().toISOString(),approver_note:''}]);
    if(!r)return toast('Submit failed: '+(API.lastError||'unknown'),'err');
    this.audit('Payroll submitted for approval','Payroll',month,t.staff+' staff · net '+t.net.toFixed(2));
    const cl=this.staff[COUNTRY_LEADER_ID];
    if(cl?.email)API.gasPost({action:'payrollNotify',stage:'submitted',month,
      to:cl.email,toName:cl.name,by:this.user.name,
      staffCount:t.staff,net:t.net.toFixed(2)}).catch(()=>{});
    toast('Submitted to the Country Leader for approval ✓');
    this.renderPayStatus(p);
  }
  async approvePayrollRun(p){
    if(!this._canApprove())return toast('Only the Country Leader can approve payroll','err');
    const month=$(p+'pay-month').value;
    const t=this._curRun?.totals?JSON.parse(this._curRun.totals):this._payTotals();
    if(!confirm('Approve payroll for '+month+'?\n\n'+(t.staff||'?')+' staff · Net payout '
      +this._ghs(t.net||0)+'\n\nOnce approved, Finance can issue payslips and the bank advice.'))return;
    const r=await API.secureSave('payroll_runs',[{id:'RUN-'+month,month,
      status:'Approved',approved_by:this.user.name,approved_at:new Date().toISOString(),approver_note:''}]);
    if(!r)return toast('Approval failed: '+(API.lastError||'unknown'),'err');
    this.audit('Payroll APPROVED','Payroll',month,'net '+(t.net||0));
    ['THPG/05/2025','THPG/01/2026-3'].forEach(id=>{
      const e=this.staff[id]?.email;
      if(e)API.gasPost({action:'payrollNotify',stage:'approved',month,to:e,toName:this.staff[id].name,
        by:this.user.name,staffCount:t.staff||0,net:(t.net||0).toFixed(2)}).catch(()=>{});
    });
    toast('Payroll approved ✓ — Finance can now issue payslips');
    this.renderPayStatus(p);
  }
  async returnPayrollRun(p){
    if(!this._canApprove())return toast('Only the Country Leader can return payroll','err');
    const note=prompt('Why are you returning this payroll run?\n\nFinance will see this note.');
    if(note===null)return;
    if(!note.trim())return toast('Please give a reason','err');
    const month=$(p+'pay-month').value;
    const r=await API.secureSave('payroll_runs',[{id:'RUN-'+month,month,
      status:'Returned',approver_note:note.trim(),approved_by:'',approved_at:null}]);
    if(!r)return toast('Return failed: '+(API.lastError||'unknown'),'err');
    this.audit('Payroll returned to Finance','Payroll',month,note.trim());
    ['THPG/05/2025','THPG/01/2026-3'].forEach(id=>{
      const e=this.staff[id]?.email;
      if(e)API.gasPost({action:'payrollNotify',stage:'returned',month,to:e,toName:this.staff[id].name,
        by:this.user.name,note:note.trim()}).catch(()=>{});
    });
    toast('Returned to Finance with your note');
    this.renderPayStatus(p);
  }
  exportPayrollCSV(p){
    if(!this._payCalc||!this._payCalc.length)return toast('Nothing to export','err');
    const month=$(p+'pay-month').value||'';
    let csv='Staff ID,Name,Unit,Basic,Gross,SSNIT Employee,Provident Fund,PAYE,Net Pay,Employer Cost\n';
    this._payCalc.forEach(r=>{csv+=`"${r.id}","${r.name}","${r.unit||''}",${r.basic.toFixed(2)},${r.gross.toFixed(2)},${r.ssnitEmp.toFixed(2)},${r.tier3.toFixed(2)},${r.paye.toFixed(2)},${r.net.toFixed(2)},${r.cost.toFixed(2)}\n`;});
    this._dl(csv,'THP_Payroll_'+month+'.csv','text/csv');
  }

  /* ═══════════════════════════════════════════
     MY PAYSLIPS — staff view their own, in-system
  ═══════════════════════════════════════════ */
  async renderMyPayslips(prefix){
    const body=$((prefix==='m-')?'m-slip-body':'st-slip-body');if(!body)return;
    body.innerHTML='<tr><td colspan="6" style="color:var(--text3)">Loading…</td></tr>';
    const r=await API.gasPost({action:'myPayslips',staffId:this.user.id,token:API._sessTok()});
    if(!r||!r.success){
      body.innerHTML='<tr><td colspan="6" style="color:var(--red)">Could not load payslips'+(r?.error?': '+r.error:'')+'</td></tr>';
      return;
    }
    const slips=r.slips||[];
    this._mySlips=slips;
    if(!slips.length){body.innerHTML='<tr><td colspan="6"><div class="empty"><div class="empty-ico">🧾</div>No approved payslips yet</div></td></tr>';return;}
    body.innerHTML=slips.map((sl,i)=>`<tr><td><strong>${sl.month}</strong></td>
      <td>${this._ghs(sl.gross)}</td><td>${this._ghs(sl.totalDed)}</td>
      <td><strong>${this._ghs(sl.net)}</strong></td>
      <td style="font-size:.76rem">${sl.approvedAt?String(sl.approvedAt).slice(0,10):'—'}</td>
      <td><button class="bsm bsm-navy" onclick="APP.viewMyPayslip(${i})">🧾 View</button></td></tr>`).join('');
    this._setBadge('slip',0);
  }
  async viewMyPayslip(i){
    const sl=(this._mySlips||[])[i];if(!sl)return;
    // Open synchronously with the click, or the browser blocks the popup.
    const w=window.open('','_blank');
    if(!w){toast('Your browser blocked the window. Allow pop-ups for this site, then try again.','err');return;}
    w.document.write('<html><body style="font-family:Arial;padding:40px;text-align:center;color:#2D3592"><h3>Opening your payslip…</h3></body></html>');
    try{
      const me=this.staff[this.user.id]||{};
      const r={...sl,id:this.user.id,name:this.user.name,unit:me.unit||'',email:me.email||''};
      this._payApproved=true;              // only approved runs are returned
      const html=await this._payslipHTML(r,sl.month);
      w.document.open();
      w.document.write(html+'<div style="text-align:center;padding:14px" class="no-print">'
        +'<button onclick="window.print()" style="padding:9px 24px;background:#2D3592;color:#fff;border:0;border-radius:8px;font-weight:600;cursor:pointer">🖨 Print / Save as PDF</button></div>');
      w.document.close();
    }catch(e){
      try{w.document.open();w.document.write('<p style="font-family:Arial;padding:40px;color:#dc2626">Could not open the payslip.</p>');w.document.close();}catch(_){}
    }
  }

  /* ═══════════════════════════════════════════
     ONBOARDING / OFFBOARDING CHECKLISTS
  ═══════════════════════════════════════════ */
  _obTemplate(kind){
    if(kind==='offboarding')return [
      {task:'Resignation / end-of-contract letter received',owner:'HR',done:false},
      {task:'Exit interview conducted',owner:'HR',done:false},
      {task:'Handover notes completed and accepted',owner:'Supervisor',done:false},
      {task:'All THP assets returned (laptop, phone, ID, keys)',owner:'HR / Admin',done:false},
      {task:'Email and system accounts deactivated',owner:'IT',done:false},
      {task:'Attendance system account disabled',owner:'Admin',done:false},
      {task:'Outstanding leave balance calculated',owner:'HR',done:false},
      {task:'Final salary and entitlements processed',owner:'Finance',done:false},
      {task:'SSNIT and statutory obligations settled',owner:'Finance',done:false},
      {task:'Clearance certificate issued',owner:'HR',done:false}];
    return [
      {task:'Signed contract received and filed',owner:'HR',done:false},
      {task:'Staff ID created in attendance system',owner:'Admin',done:false},
      {task:'THP email address created',owner:'IT',done:false},
      {task:'Staff file completed (DOB, next of kin, SSNIT)',owner:'HR',done:false},
      {task:'Bank details collected for payroll',owner:'Finance',done:false},
      {task:'Laptop / equipment issued',owner:'HR / Admin',done:false},
      {task:'Office ID card issued',owner:'HR',done:false},
      {task:'Orientation and induction completed',owner:'HR',done:false},
      {task:'Introduced to team and supervisor assigned',owner:'Supervisor',done:false},
      {task:'Safeguarding and code of conduct signed',owner:'HR',done:false},
      {task:'Added to payroll',owner:'Finance',done:false}];
  }
  async renderChecklists(){
    const body=$('ob-body');if(!body)return;
    body.innerHTML='<tr><td colspan="7" style="color:var(--text3)">Loading…</td></tr>';
    let rows=await API._get('staff_checklists','order=created_at.desc&limit=300')||[];
    this._checklists=rows;
    const f=$('ob-filter')?.value||'';
    let list=rows;
    if(f==='open')list=list.filter(r=>r.status!=='Completed');
    else if(f)list=list.filter(r=>r.kind===f);
    if(!list.length){body.innerHTML='<tr><td colspan="7"><div class="empty"><div class="empty-ico">🚀</div>No checklists yet</div></td></tr>';return;}
    body.innerHTML=list.map(c=>{
      let items=[];try{items=JSON.parse(c.items||'[]');}catch(e){}
      const done=items.filter(i=>i.done).length;
      const pct=items.length?Math.round(done/items.length*100):0;
      const bar=`<div style="display:flex;align-items:center;gap:.4rem"><div style="flex:1;height:7px;background:var(--surf2);border-radius:4px;overflow:hidden;min-width:70px"><div style="width:${pct}%;height:100%;background:${pct===100?'#16a34a':'#2D3592'}"></div></div><span style="font-size:.72rem;color:var(--text2)">${done}/${items.length}</span></div>`;
      const st=c.status==='Completed'?'<span class="c-flag green">Completed</span>':'<span class="c-flag amber">In Progress</span>';
      return `<tr><td><strong>${this._sName(c.staff_id)}</strong></td>
        <td>${c.kind==='offboarding'?'📤 Offboarding':'🚀 Onboarding'}</td>
        <td style="font-size:.78rem">${c.start_date?fmtISO(c.start_date):'—'}</td>
        <td style="font-size:.78rem">${c.target_date?fmtISO(c.target_date):'—'}</td>
        <td>${bar}</td><td>${st}</td>
        <td><button class="bsm bsm-navy" onclick="APP.openChecklistModal('${c.id}')">✏</button></td></tr>`;
    }).join('');
  }
  openChecklistModal(id){
    const c=id?(this._checklists||[]).find(x=>x.id===id):null;
    this._popStaffSel('ob-staff',c?.staff_id);
    $('ob-id').value=c?.id||'';
    $('ob-kind').value=c?.kind||'onboarding';
    $('ob-start').value=c?.start_date?String(c.start_date).slice(0,10):new Date().toISOString().slice(0,10);
    $('ob-target').value=c?.target_date?String(c.target_date).slice(0,10):'';
    $('ob-notes').value=c?.notes||'';
    $('ob-title').textContent=(c?.kind==='offboarding')?'📤 Offboarding Checklist':'🚀 Onboarding Checklist';
    let items=[];try{items=JSON.parse(c?.items||'[]');}catch(e){}
    this._obItems=items.length?items:this._obTemplate($('ob-kind').value);
    this._renderChecklistItems();
    $('ob-msg').textContent='';
    $('checklist-modal').classList.add('open');
  }
  _obLoadTemplate(){
    if($('ob-id').value)return;           // don't overwrite an existing checklist
    this._obItems=this._obTemplate($('ob-kind').value);
    $('ob-title').textContent=$('ob-kind').value==='offboarding'?'📤 Offboarding Checklist':'🚀 Onboarding Checklist';
    this._renderChecklistItems();
  }
  _renderChecklistItems(){
    const el=$('ob-items');if(!el)return;
    el.innerHTML=this._obItems.map((it,i)=>`
      <div style="display:flex;gap:.5rem;align-items:center;margin-bottom:.4rem;flex-wrap:wrap">
        <input type="checkbox" ${it.done?'checked':''} onchange="APP._obSet(${i},'done',this.checked)">
        <input class="fi" style="flex:2;min-width:180px" value="${(it.task||'').replace(/"/g,'&quot;')}" oninput="APP._obSet(${i},'task',this.value)">
        <input class="fi" style="width:110px" value="${it.owner||''}" placeholder="Owner" oninput="APP._obSet(${i},'owner',this.value)">
        <button class="bsm" style="background:var(--surf);border:1px solid var(--border);color:var(--text2)" onclick="APP._obDel(${i})">✕</button>
      </div>`).join('');
  }
  _obSet(i,f,v){this._obItems[i][f]=v;if(f==='done'){this._obItems[i].done_at=v?new Date().toISOString():null;}}
  _obDel(i){this._obItems.splice(i,1);this._renderChecklistItems();}
  addChecklistItem(){this._obItems.push({task:'',owner:'',done:false});this._renderChecklistItems();}
  async saveChecklist(){
    const staff=$('ob-staff').value;
    if(!staff)return $('ob-msg').innerHTML='<span style="color:var(--red)">Select a staff member.</span>';
    const id=$('ob-id').value||this._uid('CHK');
    const allDone=this._obItems.length&&this._obItems.every(i=>i.done);
    $('ob-msg').innerHTML='<span style="color:var(--teal)">⏳ Saving…</span>';
    const r=await API._upsert('staff_checklists',[{id,staff_id:staff,kind:$('ob-kind').value,
      status:allDone?'Completed':'In Progress',items:JSON.stringify(this._obItems),
      start_date:$('ob-start').value||null,target_date:$('ob-target').value||null,
      completed_at:allDone?new Date().toISOString():null,
      notes:$('ob-notes').value.trim(),created_by:this.user.name,updated_at:new Date().toISOString()}]);
    if(r){this.audit('Checklist saved','HR',this._sName(staff),$('ob-kind').value+(allDone?' · COMPLETED':''));
      closeModal('checklist-modal');toast('Checklist saved ✓');this.renderChecklists();}
    else $('ob-msg').innerHTML='<span style="color:var(--red)">Save failed: '+(API.lastError||'')+'</span>';
  }

  /* ═══════════════════════════════════════════
     ASSET / EQUIPMENT REGISTER
  ═══════════════════════════════════════════ */
  async renderAssets(){
    const body=$('as-body');if(!body)return;
    body.innerHTML='<tr><td colspan="8" style="color:var(--text3)">Loading…</td></tr>';
    const rows=await API._get('assets','order=id.asc&limit=500')||[];
    this._assets=rows;
    const q=($('as-search')?.value||'').trim().toLowerCase();
    const st=$('as-status')?.value||'';
    let list=rows;
    if(q)list=list.filter(a=>[a.id,a.name,a.serial_no,a.category].join(' ').toLowerCase().includes(q));
    if(st)list=list.filter(a=>(a.status||'Available')===st);
    const sm=$('as-summary');
    if(sm){
      const c=k=>rows.filter(a=>(a.status||'Available')===k).length;
      const val=rows.reduce((t,a)=>t+(+a.value||0),0);
      const box=(n,l,col,ic)=>`<div class="cs-box"><div style="font-size:1.05rem;line-height:1">${ic}</div><div class="cs-num" style="color:${col}">${n}</div><div class="cs-lbl">${l}</div></div>`;
      sm.innerHTML=box(rows.length,'Total Assets','','💻')+box(c('Assigned'),'Assigned','#2D3592','👤')
        +box(c('Available'),'Available','#16a34a','✅')+box(c('Repair')+c('Lost'),'Repair / Lost','#dc2626','⚠')
        +box(this._ghs(val),'Register Value','','₵');
    }
    if(!list.length){body.innerHTML='<tr><td colspan="8"><div class="empty"><div class="empty-ico">💻</div>No assets found</div></td></tr>';return;}
    const cc={Available:'green',Assigned:'none',Repair:'amber',Retired:'none',Lost:'red'};
    body.innerHTML=list.map(a=>{
      const assigned=(a.status==='Assigned');
      return `<tr><td style="font-weight:600">${a.id}</td>
        <td>${a.name||''}${a.value>0?`<br><span style="font-size:.68rem;color:var(--text3)">${this._ghs(a.value)}</span>`:''}</td>
        <td style="font-size:.78rem">${a.category||'—'}</td>
        <td style="font-size:.74rem">${a.serial_no||'—'}</td>
        <td style="font-size:.78rem">${a.condition||'—'}</td>
        <td><span class="c-flag ${cc[a.status]||'none'}">${a.status||'Available'}</span></td>
        <td style="font-size:.78rem">${assigned?this._sName(a.assigned_to):'—'}</td>
        <td>${assigned
          ? `<button class="bsm" style="background:rgba(22,163,74,.14);color:#16a34a" onclick="APP.openIssueModal('${a.id}','Returned')">📥 Return</button>`
          : `<button class="bsm bsm-navy" onclick="APP.openIssueModal('${a.id}','Issued')">📤 Issue</button>`}
          <button class="bsm" style="background:var(--surf2);color:var(--text2);border:1px solid var(--border)" onclick="APP.openAssetModal('${a.id}')">✏</button>
          <button class="bsm" style="background:var(--surf2);color:var(--text2);border:1px solid var(--border)" onclick="APP.assetHistory('${a.id}')">🕘</button>
          <button class="bsm" style="background:rgba(239,68,68,.12);color:var(--red)" onclick="APP.deleteAsset('${a.id}')">🗑</button></td></tr>`;
    }).join('');
  }
  _asCatChange(){
    const sel=$('as-cat'),box=$('as-cat-other');
    if(!sel||!box)return;
    const other=sel.value==='Other';
    box.style.display=other?'block':'none';
    if(other)box.focus();
  }
  openAssetModal(id){
    const a=id?(this._assets||[]).find(x=>String(x.id)===String(id)):null;
    $('as-edit-id').value=a?.id||'';
    $('as-id').value=a?.id||'';$('as-id').readOnly=!!a;
    $('as-name').value=a?.name||'';
    const known=['IT Equipment','Furniture','Vehicle','Phone','Office Equipment','Generator / Power','Field Equipment'];
    const cat=a?.category||'IT Equipment';
    if(known.includes(cat)){$('as-cat').value=cat;$('as-cat-other').value='';}
    else{$('as-cat').value='Other';$('as-cat-other').value=cat;}
    this._asCatChange();
    $('as-make').value=a?.make||'';
    $('as-model').value=a?.model||'';
    $('as-serial').value=a?.serial_no||'';
    $('as-pdate').value=a?.purchase_date?String(a.purchase_date).slice(0,10):'';
    $('as-value').value=a?.value||'';
    $('as-supplier').value=a?.supplier||'';
    $('as-invoice').value=a?.invoice_no||'';
    $('as-funding').value=a?.funding_source||'';
    $('as-grant').value=a?.grant_id||'';
    $('as-warranty').value=a?.warranty_end?String(a.warranty_end).slice(0,10):'';
    $('as-life').value=a?.useful_life_yrs||'';
    $('as-insured').value=String(a?.insured||false);
    $('as-cond').value=a?.condition||'Good';
    $('as-st').value=a?.status||'Available';
    $('as-loc').value=a?.location||'';
    $('as-verified').value=a?.last_verified?String(a.last_verified).slice(0,10):'';
    $('as-ddate').value=a?.disposal_date?String(a.disposal_date).slice(0,10):'';
    $('as-dmethod').value=a?.disposal_method||'';
    $('as-notes').value=a?.notes||'';
    $('as-msg').textContent='';
    $('asset-modal').classList.add('open');
  }
  async saveAsset(){
    const id=$('as-id').value.trim(),name=$('as-name').value.trim();
    if(!id)return $('as-msg').innerHTML='<span style="color:var(--red)">Asset tag is required.</span>';
    if(!name)return $('as-msg').innerHTML='<span style="color:var(--red)">Asset name is required.</span>';
    $('as-msg').innerHTML='<span style="color:var(--teal)">⏳ Saving…</span>';
    let cat=$('as-cat').value;
    if(cat==='Other'){
      const typed=$('as-cat-other').value.trim();
      if(!typed)return $('as-msg').innerHTML='<span style="color:var(--red)">Type the category name.</span>';
      cat=typed;
    }
    const r=await API._upsert('assets',[{id,name,category:cat,
      make:$('as-make').value.trim(),model:$('as-model').value.trim(),
      serial_no:$('as-serial').value.trim(),purchase_date:$('as-pdate').value||null,
      value:+$('as-value').value||0,supplier:$('as-supplier').value.trim(),
      invoice_no:$('as-invoice').value.trim(),funding_source:$('as-funding').value.trim(),
      grant_id:$('as-grant').value.trim(),warranty_end:$('as-warranty').value||null,
      useful_life_yrs:+$('as-life').value||0,insured:$('as-insured').value==='true',
      condition:$('as-cond').value,status:$('as-st').value,location:$('as-loc').value.trim(),
      last_verified:$('as-verified').value||null,disposal_date:$('as-ddate').value||null,
      disposal_method:$('as-dmethod').value.trim(),
      notes:$('as-notes').value.trim(),updated_at:new Date().toISOString()}]);
    if(r){this.audit('Asset saved','HR',name,id+' · '+$('as-st').value);
      closeModal('asset-modal');toast('Asset saved ✓');this.renderAssets();}
    else $('as-msg').innerHTML='<span style="color:var(--red)">Save failed: '+(API.lastError||'')+'</span>';
  }
  async deleteAsset(id){
    const a=(this._assets||[]).find(x=>String(x.id)===String(id));if(!a)return;
    const who=id+' — '+(a.name||'');
    if(a.status==='Assigned')return toast('This asset is still assigned to '+this._sName(a.assigned_to)+'. Record a return first.','err');
    if(!confirm('Delete this asset from the register?\n\n'+who+'\n\nConsider setting the status to Retired or Disposed instead, so the history is preserved.'))return;
    const t=prompt('This cannot be undone.\n\nType DELETE to confirm removal of:\n'+who);
    if(t===null)return;
    if(String(t).trim().toUpperCase()!=='DELETE')return toast('Not deleted — confirmation did not match','info');
    await API._delete('assets','id=eq.'+encodeURIComponent(id));
    this.audit('Asset deleted','HR',a.name||'',id);
    toast('Asset deleted');this.renderAssets();
  }
  openIssueModal(assetId,action){
    const a=(this._assets||[]).find(x=>String(x.id)===String(assetId));if(!a)return;
    $('iss-id').value=assetId;$('iss-action').value=action;
    $('iss-title').textContent=action==='Issued'?'📤 Issue Asset':'📥 Return Asset';
    $('iss-asset').textContent=a.id+' — '+(a.name||'');
    $('iss-date').value=new Date().toISOString().slice(0,10);
    $('iss-cond').value=a.condition||'Good';
    $('iss-note').value='';
    const wrap=$('iss-staff-wrap');
    if(action==='Issued'){wrap.style.display='block';this._popStaffSel('iss-staff');}
    else{wrap.style.display='none';}
    $('iss-msg').textContent='';
    $('issue-modal').classList.add('open');
  }
  async confirmIssue(){
    const assetId=$('iss-id').value,action=$('iss-action').value;
    const a=(this._assets||[]).find(x=>String(x.id)===String(assetId));if(!a)return;
    const staffId=action==='Issued'?$('iss-staff').value:(a.assigned_to||'');
    if(action==='Issued'&&!staffId)return $('iss-msg').innerHTML='<span style="color:var(--red)">Select who it is issued to.</span>';
    const cond=$('iss-cond').value,note=$('iss-note').value.trim();
    $('iss-msg').innerHTML='<span style="color:var(--teal)">⏳ Recording…</span>';
    const upd=await API._upsert('assets',[{id:assetId,name:a.name,category:a.category,serial_no:a.serial_no,
      purchase_date:a.purchase_date,value:a.value,notes:a.notes,condition:cond,
      status:action==='Issued'?'Assigned':'Available',
      assigned_to:action==='Issued'?staffId:'',
      assigned_at:action==='Issued'?new Date().toISOString():null,
      updated_at:new Date().toISOString()}]);
    if(!upd)return $('iss-msg').innerHTML='<span style="color:var(--red)">Failed: '+(API.lastError||'')+'</span>';
    await API._upsert('asset_movements',[{id:this._uid('MOV'),asset_id:assetId,asset_name:a.name||'',
      staff_id:staffId,staff_name:this._sName(staffId)||'',action,condition:cond,note,
      handled_by:this.user.name,created_at:new Date().toISOString()}]);
    this.audit('Asset '+action.toLowerCase(),'HR',this._sName(staffId)||'—',assetId+' · '+(a.name||''));
    // Email record to HR and the staff member — the paper trail if the system is ever doubted
    const st=this.staff[staffId]||{};
    API.gasPost({action:'assetNotify',assetId,assetName:a.name||'',serial:a.serial_no||'',
      staffId,staffName:this._sName(staffId)||'',staffEmail:st.email||'',staffPhone:st.phone||'',
      movement:action,condition:cond,note,handledBy:this.user.name,
      date:$('iss-date').value||new Date().toISOString().slice(0,10)}).catch(()=>{});
    closeModal('issue-modal');
    toast('Asset '+action.toLowerCase()+' ✓ — email record sent');
    this.renderAssets();
  }
  async assetHistory(assetId){
    const rows=await API._get('asset_movements','asset_id=eq.'+encodeURIComponent(assetId)+'&order=created_at.desc')||[];
    const a=(this._assets||[]).find(x=>String(x.id)===String(assetId));
    const w=window.open('','_blank');
    w.document.write(`<html><head><meta charset="UTF-8"><title>Asset history — ${assetId}</title><style>
      body{font-family:Arial,sans-serif;font-size:11px;color:#1e293b;padding:14mm}
      h1{font-size:15px;color:#2D3592;border-bottom:3px solid #2D3592;padding-bottom:7px;margin:0 0 4px}
      .meta{font-size:10px;color:#555;margin-bottom:12px}
      table{width:100%;border-collapse:collapse}
      th{background:#2D3592;color:#fff;padding:6px;text-align:left;font-size:9.5px}
      td{padding:5px 6px;border:1px solid #e2e8f0;font-size:10px}
      tr:nth-child(even) td{background:#f8fafc}
      </style></head><body>
      <h1>The Hunger Project — Ghana · Asset Movement History</h1>
      <div class="meta"><strong>${assetId}</strong> — ${a?.name||''} ${a?.serial_no?('· Serial '+a.serial_no):''}<br>
        Generated ${new Date().toLocaleString('en-GB')}</div>
      <table><tr><th>Date</th><th>Action</th><th>Staff</th><th>Condition</th><th>Note</th><th>Handled by</th></tr>
      ${rows.length?rows.map(m=>`<tr><td>${new Date(m.created_at).toLocaleString('en-GB')}</td>
        <td>${m.action}</td><td>${m.staff_name||'—'}</td><td>${m.condition||''}</td>
        <td>${m.note||''}</td><td>${m.handled_by||''}</td></tr>`).join('')
        :'<tr><td colspan="6">No movements recorded</td></tr>'}
      </table>
      <div style="text-align:center;margin-top:16px" class="no-print"><button onclick="window.print()" style="padding:8px 22px;background:#2D3592;color:#fff;border:0;border-radius:6px;font-weight:600;cursor:pointer">🖨 Print</button></div>
      </body></html>`);
    w.document.close();
  }
  exportAssets(){
    const rows=this._assets||[];
    if(!rows.length)return toast('Nothing to export','err');
    let csv='Asset Tag,Name,Category,Serial,Purchase Date,Value,Condition,Status,Assigned To\n';
    rows.forEach(a=>{csv+=`"${a.id}","${(a.name||'').replace(/"/g,'""')}","${a.category||''}","${a.serial_no||''}","${a.purchase_date||''}",${a.value||0},"${a.condition||''}","${a.status||''}","${this._sName(a.assigned_to)||''}"\n`;});
    this._dl(csv,'THP_Assets_'+Date.now()+'.csv','text/csv');
  }

  /* ═══════════════════════════════════════════
     APPRAISAL WORKFLOW v2
     Staff rates self → line manager rates → HR record.
     Neither side can alter the other's rating.
  ═══════════════════════════════════════════ */
  _overallWord(v){return ['— not set','Unsatisfactory Performance','Development Needed','Meets Role Requirements','Exceeds Role Requirements','Exceptional Performance'][+v||0];}
  _kpaSelfAvg(k){const r=(k.kpis||[]).map(x=>+x.self||0).filter(v=>v>0);return r.length?r.reduce((a,b)=>a+b,0)/r.length:0;}
  _kpaMgrAvg(k){const r=(k.kpis||[]).map(x=>+x.mgr||0).filter(v=>v>0);return r.length?r.reduce((a,b)=>a+b,0)/r.length:0;}
  _weighted(kpas,fn){return (kpas||[]).reduce((t,k)=>t+fn.call(this,k)*(+k.weight||0)/100,0);}

  /* ── Case management v2: stages, documents, two-step delete ── */
  _caseStages(){return ['Complaint / allegation received','Investigation carried out','Query letter issued to employee',
    'Employee written response received','Notice of hearing issued','Disciplinary hearing held','Hearing minutes recorded',
    'Decision / outcome issued','Right of appeal communicated','Appeal heard and concluded'];}
  _renderCaseStages(){
    const el=$('cs-stages');if(!el)return;
    el.innerHTML=this._csStages.map((st,i)=>`
      <div style="display:flex;gap:.5rem;align-items:center;margin-bottom:.4rem;flex-wrap:wrap">
        <input type="checkbox" ${st.done?'checked':''} onchange="APP._csSet(${i},'done',this.checked)">
        <span style="flex:2;min-width:200px;font-size:.8rem">${st.name}</span>
        <input class="fi" type="date" style="width:150px" value="${st.date||''}" onchange="APP._csSet(${i},'date',this.value)">
      </div>`).join('');
  }
  _csSet(i,f,v){this._csStages[i][f]=v;if(f==='done'&&v&&!this._csStages[i].date)
    {this._csStages[i].date=new Date().toISOString().slice(0,10);this._renderCaseStages();}}
  _renderCaseDocs(docs){
    $('cs-docs').innerHTML=docs.length
      ? docs.map((d,i)=>`<span class="hr-doc-chip">📎 <a href="${d.url}" target="_blank">${d.name}</a> <a href="#" onclick="APP.removeCaseDoc(${i});return false" style="color:var(--red)">✕</a></span>`).join('')
      : '<span style="font-size:.76rem;color:var(--text3)">No documents attached yet.</span>';
  }
  removeCaseDoc(i){
    let docs=[];try{docs=JSON.parse($('cs-docs-json').value||'[]');}catch(e){}
    const d=docs[i];if(!d)return;
    if(!confirm('Remove this document from the case file?\n\n'+d.name))return;
    const t=prompt('This cannot be undone.\n\nType DELETE to confirm removal of:\n'+d.name);
    if(t===null)return;
    if(String(t).trim().toUpperCase()!=='DELETE')return toast('Not deleted — confirmation did not match','info');
    this.audit('Case document removed','Security',this._sName($('cs-staff').value),d.name);
    docs.splice(i,1);
    $('cs-docs-json').value=JSON.stringify(docs);
    this._renderCaseDocs(docs);
  }
  async uploadCaseDoc(){
    if(this._csBusy)return toast('An upload is already in progress','info');
    const inp=$('cs-doc-file');if(!inp?.files?.length)return toast('Choose a file first','err');
    const file=inp.files[0];
    if(file.size>5*1024*1024)return toast('File too large — maximum 5 MB','err');
    this._csBusy=true;
    $('cs-msg').innerHTML='<span style="color:var(--teal)">⏳ Uploading '+(file.size/1048576).toFixed(1)+' MB — this can take 30–60 seconds…</span>';
    try{
      const b64=await this._fileToBase64(file);
      const r=await API.gasPost({action:'uploadHRDoc',staffId:'CASE_'+($('cs-id').value||'new'),
        fileName:file.name,fileData:b64,mimeType:file.type});
      if(r&&r.success&&r.fileUrl){
        let docs=[];try{docs=JSON.parse($('cs-docs-json').value||'[]');}catch(e){}
        docs.push({name:file.name,url:r.fileUrl,at:new Date().toISOString().slice(0,10)});
        $('cs-docs-json').value=JSON.stringify(docs);
        this._renderCaseDocs(docs);inp.value='';
        this.audit('Case document uploaded','Security',this._sName($('cs-staff').value),file.name);
        $('cs-msg').innerHTML='<span style="color:var(--green)">✓ Uploaded — click Save to confirm.</span>';
      }else $('cs-msg').innerHTML='<span style="color:var(--red)">Upload failed'+((r&&r.error)?': '+r.error:'')+'</span>';
    }catch(e){$('cs-msg').innerHTML='<span style="color:var(--red)">Upload error: '+(e.message||e)+'</span>';}
    finally{this._csBusy=false;}
  }
  async deleteCase(id){
    if(this.user.id!==HR_MANAGER_ID&&this.user.role!=='admin')
      return toast('Only HR or the Administrator can delete a case file','err');
    const c=(this._caseRows||[]).find(x=>x.id===id);if(!c)return;
    const who=(c.case_ref||id)+' — '+this._sName(c.staff_id);
    if(!confirm('Delete this case file?\n\n'+who+'\n\nDisciplinary records are normally retained. Consider closing the case instead.'))return;
    const t=prompt('This cannot be undone.\n\nType DELETE to confirm removal of:\n'+who);
    if(t===null)return;
    if(String(t).trim().toUpperCase()!=='DELETE')return toast('Not deleted — confirmation did not match','info');
    await API.secureDelete('hr_cases','id=eq.'+encodeURIComponent(id));
    this.audit('CASE FILE DELETED','Security',this._sName(c.staff_id),(c.case_ref||id)+' · '+(c.case_type||''));
    toast('Case deleted');
    this.renderCases('m-');
  }

  /* ── Reviews waiting on me as line manager ── */
  async renderToRate(){
    const body=$('m-torate-body');if(!body)return;
    body.innerHTML='<tr><td colspan="5" style="color:var(--text3)">Loading…</td></tr>';
    const rows=await API._get('performance_appraisals','line_manager=eq.'+encodeURIComponent(this.user.id)
      +'&status=eq.Self-Assessed&order=self_submitted.desc')||[];
    this._toRate=rows;
    this._setGrpBadge('badge-myapr',rows.length);
    if(!rows.length){body.innerHTML='<tr><td colspan="5" style="color:var(--text3);font-size:.8rem">Nothing awaiting your rating.</td></tr>';return;}
    body.innerHTML=rows.map(r=>`<tr><td><strong>${this._sName(r.staff_id)}</strong></td>
      <td>${r.period||'—'}</td><td>${(+r.self_score||0).toFixed(2)}/5</td>
      <td style="font-size:.76rem">${r.self_submitted?String(r.self_submitted).slice(0,10):'—'}</td>
      <td><button class="bsm bsm-navy" onclick="APP.openMgrRating('${r.id}')">📝 Rate</button></td></tr>`).join('');
  }
  openMgrRating(id){
    const r=(this._toRate||[]).find(x=>x.id===id);if(!r)return;
    this._mrRec=r;
    let k=[];try{k=JSON.parse(r.kpas||'[]');}catch(e){}
    this._mrKpas=k;
    $('mr-id').value=id;
    $('mr-who').textContent=this._sName(r.staff_id)+' — '+(r.period||'')+' ('+(r.review_type||'Annual')+')';
    $('mr-selfoverall').value=this._overallWord(r.self_overall);
    $('mr-mgroverall').value=r.mgr_overall||0;
    $('mr-comment').value=r.mgr_comment||'';
    $('mr-strengths').value=r.strengths||'';
    $('mr-devareas').value=r.dev_areas||'';
    $('mr-training').value=r.training_needs||r.dev_plan||'';
    $('mr-support').value=r.support_needed||'';
    $('mr-next').value=r.next_period_obj||'';
    this._renderMrKpas();
    $('mr-msg').textContent='';
    $('mgrate-modal').classList.add('open');
  }
  _renderMrKpas(){
    const el=$('mr-kpas');if(!el)return;
    el.innerHTML=this._mrKpas.map((k,i)=>`
      <div class="kpa-box">
        <div class="kpa-hd"><span style="font-size:.72rem;color:var(--text3);font-weight:700">KPA ${i+1}</span>
          <strong style="flex:1">${k.name||''}</strong><span style="font-size:.74rem;color:var(--text2)">${k.weight||0}%</span></div>
        ${(k.kpis||[]).map((x,j)=>{
          const self=+x.self||0,mgr=+x.mgr||0;
          const fin=(self&&mgr)?((self+mgr)/2).toFixed(2):(mgr||self||0).toFixed(2);
          return `<div class="kpi-row">
            <div style="flex:2;min-width:160px;font-size:.78rem">${x.kpi||'<em style="color:var(--text3)">(no detail given)</em>'}</div>
            <span style="font-size:.72rem;color:var(--text3)">self</span>
            <input class="fi kr" type="number" value="${self}" readonly style="opacity:.6">
            <span style="font-size:.72rem;color:var(--teal);font-weight:600">mine</span>
            <input class="fi kr" type="number" min="0" max="5" step="0.5" value="${mgr}" oninput="APP._mrSet(${i},${j},this.value)">
            <span style="font-size:.74rem">= <strong>${fin}</strong></span>
          </div>`;}).join('')}
        <div class="kpa-avg">Self ${this._kpaSelfAvg(k).toFixed(2)} · Manager ${this._kpaMgrAvg(k).toFixed(2)}</div>
      </div>`).join('');
    const self=this._weighted(this._mrKpas,this._kpaSelfAvg);
    const mgr=this._weighted(this._mrKpas,this._kpaMgrAvg);
    const fin=(self&&mgr)?(self+mgr)/2:(mgr||self);
    if($('mr-self'))$('mr-self').textContent=self.toFixed(2);
    if($('mr-mgr'))$('mr-mgr').textContent=mgr.toFixed(2);
    if($('mr-final'))$('mr-final').textContent=fin.toFixed(2);
  }
  _mrSet(i,j,v){this._mrKpas[i].kpis[j].mgr=+v||0;this._renderMrKpas();}
  async saveMgrRating(complete){
    const id=$('mr-id').value,r=this._mrRec;if(!r)return;
    const mgrScore=this._weighted(this._mrKpas,this._kpaMgrAvg);
    const selfScore=this._weighted(this._mrKpas,this._kpaSelfAvg);
    const finalScore=(selfScore&&mgrScore)?(selfScore+mgrScore)/2:(mgrScore||selfScore);
    if(complete){
      if(!+$('mr-mgroverall').value)return $('mr-msg').innerHTML='<span style="color:var(--red)">Select your overall assessment.</span>';
      const unrated=this._mrKpas.some(k=>(k.kpis||[]).some(x=>(+x.self>0)&&!(+x.mgr>0)));
      if(unrated&&!confirm('Some items the staff member rated have no rating from you.\n\nComplete anyway?'))return;
      if(!confirm('Complete this review and send it to HR?\n\nFinal score: '+finalScore.toFixed(2)+'/5\n\nIt cannot be edited afterwards.'))return;
    }
    $('mr-msg').innerHTML='<span style="color:var(--teal)">⏳ Saving…</span>';
    const upd={id,staff_id:r.staff_id,period:r.period,review_type:r.review_type,line_manager:this.user.id,
      kpas:JSON.stringify(this._mrKpas),self_score:+selfScore.toFixed(2),mgr_score:+mgrScore.toFixed(2),
      final_score:+finalScore.toFixed(2),mgr_overall:+$('mr-mgroverall').value||0,
      mgr_comment:$('mr-comment').value.trim(),strengths:$('mr-strengths').value.trim(),
      dev_areas:$('mr-devareas').value.trim(),training_needs:$('mr-training').value.trim(),
      support_needed:$('mr-support').value.trim(),next_period_obj:$('mr-next').value.trim(),
      updated_at:new Date().toISOString()};
    if(complete){upd.status='Manager Reviewed';upd.mgr_submitted=new Date().toISOString();}
    const res=await API._upsert('performance_appraisals',[upd]);
    if(!res)return $('mr-msg').innerHTML='<span style="color:var(--red)">Save failed: '+(API.lastError||'')+'</span>';
    this.audit(complete?'Appraisal completed by line manager':'Appraisal rating saved','HR',
      this._sName(r.staff_id),(r.period||'')+' · final '+finalScore.toFixed(2));
    if(complete){
      const st=this.staff[r.staff_id]||{},edna=this.staff[HR_MANAGER_ID];
      [{to:st.email,name:st.name,role:'staff'},{to:edna?.email,name:edna?.name,role:'hr'}]
        .filter(x=>x.to).forEach(x=>API.gasPost({action:'appraisalDone',to:x.to,toName:x.name,role:x.role,
          staffName:this._sName(r.staff_id),period:r.period||'',score:finalScore.toFixed(2),
          rating:this._overallWord($('mr-mgroverall').value),by:this.user.name}).catch(()=>{}));
    }
    closeModal('mgrate-modal');
    toast(complete?'Review completed and sent to HR ✓':'Draft saved ✓');
    this.renderToRate();this.renderPerf('m-');
  }

  /* ── Delete an appraisal (HR) — two-step ── */
  async deleteAppraisal(id){
    const r=(this._perfRows||[]).find(x=>x.id===id);if(!r)return;
    const who=this._sName(r.staff_id)+' — '+(r.period||'');
    if(!confirm('Delete this appraisal record?\n\n'+who))return;
    const t=prompt('This cannot be undone.\n\nType DELETE to confirm removal of:\n'+who);
    if(t===null)return;
    if(String(t).trim().toUpperCase()!=='DELETE')return toast('Not deleted — confirmation did not match','info');
    const ok=await API._delete('performance_appraisals','id=eq.'+encodeURIComponent(id));
    this.audit('Appraisal deleted','HR',this._sName(r.staff_id),(r.period||'')+' · score '+(r.final_score||0));
    toast('Appraisal deleted');
    this.renderPerf('m-');
  }

  /* ═══════════════════════════════════════════
     ADMIN — settings, security, health, lifecycle
  ═══════════════════════════════════════════ */
  async renderOrgSettings(){
    const rows=await API._get('settings','select=key,value')||[];
    const m={};rows.forEach(r=>m[r.key]=r.value);
    const set=(el,k,d)=>{const e=$(el);if(e)e.value=m[k]??d;};
    set('set-org-name','org_name','The Hunger Project — Ghana');
    set('set-org-address','org_address','PMB CT 7, Cantonments Accra, Ghana');
    set('set-org-email','org_email','thpghana@thp.org');
    set('set-work-start','work_start','08:00');
    set('set-work-end','work_end','17:00');
    set('set-late','late_after','08:30');
    set('set-early','early_exit_before','16:30');
    set('set-lv-annual','leave_annual','24');
    set('set-lv-mat','leave_maternity','65');
    set('set-lv-pat','leave_paternity','5');
    set('set-lv-comp','leave_compassionate','5');
    set('set-session','session_hours','12');
    $('set-msg').textContent='';
  }
  async saveOrgSettings(){
    const pairs=[['org_name','set-org-name'],['org_address','set-org-address'],['org_email','set-org-email'],
      ['work_start','set-work-start'],['work_end','set-work-end'],['late_after','set-late'],
      ['early_exit_before','set-early'],['leave_annual','set-lv-annual'],['leave_maternity','set-lv-mat'],
      ['leave_paternity','set-lv-pat'],['leave_compassionate','set-lv-comp'],['session_hours','set-session']];
    const rows=pairs.map(([k,el])=>({key:k,value:String($(el)?.value??'')}));
    $('set-msg').innerHTML='<span style="color:var(--teal)">⏳ Saving…</span>';
    const r=await API._upsert('settings',rows);
    if(r){this.audit('System settings updated','Security','',rows.length+' values');
      $('set-msg').innerHTML='<span style="color:var(--green)">✓ Saved. Some changes apply at next sign-in.</span>';
      toast('Settings saved ✓');}
    else $('set-msg').innerHTML='<span style="color:var(--red)">Save failed: '+(API.lastError||'')+'</span>';
  }

  async renderSecurity(){
    const sb=$('sec-sess-body'),lb=$('sec-log-body');
    if(sb)sb.innerHTML='<tr><td colspan="4" style="color:var(--text3)">Loading…</td></tr>';
    const [sess,log]=await Promise.all([
      API._get('sessions','select=id,staff_id,expires_at,created_at&order=created_at.desc&limit=100'),
      API._get('login_log','select=*&order=created_at.desc&limit=100')
    ]);
    const now=new Date();
    const live=(sess||[]).filter(x=>!x.expires_at||new Date(x.expires_at)>now);
    if(sb)sb.innerHTML=live.length?live.map(x=>`<tr>
      <td><strong>${this._sName(x.staff_id)||x.staff_id}</strong><br><span style="font-size:.7rem;color:var(--text3)">${x.staff_id}</span></td>
      <td style="font-size:.78rem">${x.created_at?new Date(x.created_at).toLocaleString('en-GB'):'—'}</td>
      <td style="font-size:.78rem">${x.expires_at?new Date(x.expires_at).toLocaleString('en-GB'):'—'}</td>
      <td><button class="bsm" style="background:rgba(239,68,68,.12);color:var(--red)" onclick="APP.endSession('${x.id}','${x.staff_id}')">⏻ End</button></td></tr>`).join('')
      :'<tr><td colspan="4" style="color:var(--text3);font-size:.8rem">Nobody is signed in.</td></tr>';
    const f=$('sec-filter')?.value||'';
    let rows=log||[];
    if(f)rows=rows.filter(r=>r.outcome===f);
    if(lb)lb.innerHTML=rows.length?rows.map(r=>`<tr>
      <td style="font-size:.76rem;white-space:nowrap">${new Date(r.created_at).toLocaleString('en-GB')}</td>
      <td style="font-size:.76rem">${r.staff_id||'—'}</td><td>${r.name||'—'}</td>
      <td><span class="c-flag ${r.outcome==='success'?'green':'red'}">${r.outcome}</span></td>
      <td style="font-size:.74rem;color:var(--text2)">${r.reason||''}</td></tr>`).join('')
      :'<tr><td colspan="5" style="color:var(--text3);font-size:.8rem">No sign-in attempts recorded yet.</td></tr>';
  }
  async endSession(sid,staffId){
    if(!confirm('Sign out '+(this._sName(staffId)||staffId)+'?\nThey will need to sign in again.'))return;
    await API._delete('sessions','id=eq.'+encodeURIComponent(sid));
    this.audit('Session ended by admin','Security',this._sName(staffId)||staffId,'');
    toast('Signed out');this.renderSecurity();
  }
  async endAllSessions(){
    if(!confirm('Sign EVERYONE out, including yourself?\n\nAll staff will need to sign in again.'))return;
    const t=prompt('Type SIGNOUT to confirm.');
    if(String(t||'').trim().toUpperCase()!=='SIGNOUT')return toast('Cancelled','info');
    await API._delete('sessions','id=neq.__none__');
    this.audit('All sessions ended','Security','','admin action');
    toast('Everyone signed out');this.renderSecurity();
  }

  async renderHealth(){
    const el=$('health-body');if(!el)return;
    el.innerHTML='<div style="color:var(--text3);font-size:.8rem">Running checks…</div>';
    const line=(ok,label,detail)=>`<div style="display:flex;gap:.6rem;align-items:flex-start;padding:.55rem 0;border-bottom:1px solid var(--border)">
      <span style="font-size:1rem">${ok===null?'⚪':ok?'✅':'⚠️'}</span>
      <div style="flex:1"><div style="font-size:.85rem;font-weight:600">${label}</div>
      <div style="font-size:.76rem;color:var(--text2)">${detail}</div></div></div>`;
    const out=[];
    const staffList=Object.entries(this.staff).filter(([i,s])=>s.role!=='admin');
    // data completeness
    const files=await API.getAllHRFiles();
    const fm={};(files||[]).forEach(f=>fm[f.staff_id]=f);
    const noDob=staffList.filter(([i])=>!fm[i]?.dob).length;
    const noPhone=staffList.filter(([i,s])=>!(fm[i]?.phone||s.phone)).length;
    const noEmail=staffList.filter(([i,s])=>!s.email).length;
    const noContract=staffList.filter(([i,s])=>!s.contractEnd).length;
    out.push(line(noEmail===0,'Email addresses',noEmail===0?'All staff have an email on file.':noEmail+' staff have no email — they cannot receive notifications.'));
    out.push(line(noPhone===0,'Phone numbers',noPhone===0?'All staff have a phone number.':noPhone+' staff have no phone — SMS cannot reach them.'));
    out.push(line(noDob===0,'Dates of birth',noDob===0?'All staff files have a date of birth.':noDob+' missing — birthday greetings will skip them.'));
    out.push(line(noContract===0,'Contract dates',noContract===0?'All staff have a contract end date.':noContract+' have no contract end date.'));
    // duplicate emails would misdirect payslips
    const byEmail={};staffList.forEach(([i,s])=>{if(s.email){const e=s.email.toLowerCase();(byEmail[e]=byEmail[e]||[]).push(s.name);}});
    const dup=Object.entries(byEmail).filter(([e,n])=>n.length>1);
    out.push(line(dup.length===0,'Duplicate email addresses',dup.length===0?'No two staff share an address.'
      :dup.map(([e,n])=>e+' → '+n.join(' & ')).join('; ')+' — payslips would be misdirected.'));
    // backup mirror
    const jr=await API._get('job_runs','select=run_at&order=run_at.desc&limit=1')||[];
    const last=jr[0]?.run_at;
    const ageH=last?Math.round((Date.now()-new Date(last))/3600000):null;
    out.push(line(ageH!==null&&ageH<36,'Scheduled jobs',last?('Last automated job ran '+new Date(last).toLocaleString('en-GB')+(ageH>36?' — that is over a day ago.':'')):'No automated job has run yet.'));
    // payroll state
    try{
      const runs=await API.secureGet('payroll_runs','select=month,status&order=month.desc&limit=3')||[];
      out.push(line(true,'Payroll',runs.length?runs.map(r=>r.month+': '+r.status).join(' · '):'No payroll run recorded yet.'));
    }catch(e){out.push(line(null,'Payroll','Not available to this account.'));}
    // sessions
    const sess=await API._get('sessions','select=id&limit=200')||[];
    out.push(line(true,'Active sessions',sess.length+' session(s) on record.'));
    out.push(line(true,'Staff on record',staffList.length+' active staff, '+Object.keys(this.staff).length+' records in total.'));
    el.innerHTML=out.join('');
  }

  /* ── Deactivate instead of delete ── */
  async toggleStaffActive(id){
    const s=this.staff[id];if(!s)return;
    const isActive=s.active!==false;
    if(isActive){
      if(!confirm('Deactivate '+s.name+'?\n\nThey will no longer be able to sign in or clock in, but all their\nrecords, leave history and appraisals are kept.'))return;
      const reason=prompt('Reason for leaving (resignation, end of contract, etc.)');
      if(reason===null)return;
      const r=await API._update('staff','id=eq.'+encodeURIComponent(id),
        {active:false,exit_date:new Date().toISOString().slice(0,10),exit_reason:String(reason||'').trim()});
      if(r===null)return toast('Failed: '+(API.lastError||''),'err');
      s.active=false;this._cacheS();
      this.audit('Staff deactivated','Staff',s.name,String(reason||''));
      toast(s.name+' deactivated — records kept');
    } else {
      if(!confirm('Reactivate '+s.name+'?\n\nThey will be able to sign in and clock in again.'))return;
      const r=await API._update('staff','id=eq.'+encodeURIComponent(id),{active:true,exit_date:null,exit_reason:''});
      if(r===null)return toast('Failed: '+(API.lastError||''),'err');
      s.active=true;this._cacheS();
      this.audit('Staff reactivated','Staff',s.name,'');
      toast(s.name+' reactivated');
    }
    this._renderStaffGrid();
  }

  /* ═══════════════════════════════════════════
     GRANTS REGISTER — Finance (Ernest) + Country Leader
  ═══════════════════════════════════════════ */
  /* Grants use dd-mm-yyyy, per the Grant Register */
  _grDate(v){
    if(!v)return '—';
    const d=new Date(String(v).slice(0,10));
    if(isNaN(d))return String(v);
    const p=n=>String(n).padStart(2,'0');
    return p(d.getDate())+'-'+p(d.getMonth()+1)+'-'+d.getFullYear();
  }
  _grantFlag(endDate){
    if(!endDate)return{cls:'none',label:'No end date',days:null,months:null};
    const end=new Date(String(endDate).slice(0,10));
    if(isNaN(end))return{cls:'none',label:'Invalid date',days:null,months:null};
    const now=new Date();now.setHours(0,0,0,0);
    const days=Math.round((end-now)/86400000);
    const months=(days/30.44).toFixed(1);
    if(days<0)return{cls:'red',label:'⚠ Ended '+Math.abs(days)+'d ago',days,months};
    if(days<=90)return{cls:'red',label:'🔴 '+days+' days left',days,months};
    if(days<=180)return{cls:'amber',label:'🟠 '+days+' days left',days,months};
    return{cls:'green',label:'🟢 '+months+' months left',days,months};
  }
  async renderGrants(){
    const body=$('gr-body');if(!body)return;
    body.innerHTML='<tr><td colspan="8" style="color:var(--text3)">Loading…</td></tr>';
    let rows=await API._get('grants','order=end_date.asc.nullslast&limit=300')||[];
    this._grants=rows;
    const q=($('gr-search')?.value||'').trim().toLowerCase();
    const st=$('gr-status')?.value||'';
    let list=rows;
    if(q)list=list.filter(g=>(g.name||'').toLowerCase().includes(q)||String(g.id).includes(q));
    if(st)list=list.filter(g=>(g.status||'Active')===st);
    // summary across ALL grants, not just the filtered view
    const sm=$('gr-summary');
    if(sm){
      let red=0,amber=0,green=0;
      rows.forEach(g=>{const f=this._grantFlag(g.end_date);
        if(f.cls==='red')red++;else if(f.cls==='amber')amber++;else if(f.cls==='green')green++;});
      const box=(n,l,c,ic)=>`<div class="cs-box"><div style="font-size:1.05rem;line-height:1">${ic}</div><div class="cs-num" style="color:${c}">${n}</div><div class="cs-lbl">${l}</div></div>`;
      sm.innerHTML=box(rows.length,'Total Grants','','🎗')
        +box(red,'Ending ≤90 days','#dc2626','🔴')
        +box(amber,'Ending ≤180 days','#d97706','🟠')
        +box(green,'Running','#16a34a','🟢');
    }
    if(!list.length){body.innerHTML='<tr><td colspan="8"><div class="empty"><div class="empty-ico">🎗</div>No grants found</div></td></tr>';return;}
    body.innerHTML=list.map(g=>{
      const f=this._grantFlag(g.end_date);
      return `<tr><td style="font-weight:600">${g.id}</td>
        <td>${g.name||''}${g.amount>0?`<br><span style="font-size:.7rem;color:var(--text3)">${g.currency||''} ${(+g.amount).toLocaleString()}</span>`:''}</td>
        <td style="font-size:.78rem">${g.pc_local||'—'}</td>
        <td style="font-size:.78rem">${this._grDate(g.start_date)}</td>
        <td style="font-size:.78rem;font-weight:600">${this._grDate(g.end_date)}</td>
        <td style="font-size:.78rem">${f.months!==null&&f.days>=0?f.months:'—'}</td>
        <td><span class="c-flag ${f.cls}">${f.label}</span></td>
        <td><button class="bsm bsm-navy" onclick="APP.openGrantModal('${g.id}')">✏</button></td></tr>`;
    }).join('');
  }
  openGrantModal(id){
    const g=id?(this._grants||[]).find(x=>String(x.id)===String(id)):null;
    $('gr-edit-id').value=g?.id||'';
    $('gr-id').value=g?.id||'';
    $('gr-id').readOnly=!!g;
    $('gr-name').value=g?.name||'';
    $('gr-pc').value=g?.pc_local||'US';
    $('gr-start').value=g?.start_date?String(g.start_date).slice(0,10):'';
    $('gr-end').value=g?.end_date?String(g.end_date).slice(0,10):'';
    $('gr-amount').value=g?.amount||'';
    $('gr-curr').value=g?.currency||'USD';
    $('gr-st').value=g?.status||'Active';
    $('gr-focal').value=g?.focal_person||'';
    $('gr-notes').value=g?.notes||'';
    $('gr-msg').textContent='';
    $('grant-modal').classList.add('open');
  }
  async saveGrant(){
    const id=$('gr-id').value.trim(),name=$('gr-name').value.trim();
    if(!id)return $('gr-msg').innerHTML='<span style="color:var(--red)">Grant ID is required.</span>';
    if(!name)return $('gr-msg').innerHTML='<span style="color:var(--red)">Grant name is required.</span>';
    const start=$('gr-start').value,end=$('gr-end').value;
    if(start&&end&&end<start)return $('gr-msg').innerHTML='<span style="color:var(--red)">End date is before start date.</span>';
    let months=null;
    if(start&&end)months=Math.round((new Date(end)-new Date(start))/86400000/30.44);
    $('gr-msg').innerHTML='<span style="color:var(--teal)">⏳ Saving…</span>';
    const r=await API._upsert('grants',[{id,name,pc_local:$('gr-pc').value,
      start_date:start||null,end_date:end||null,months_active:months,
      amount:+$('gr-amount').value||0,currency:$('gr-curr').value,status:$('gr-st').value,
      focal_person:$('gr-focal').value.trim(),notes:$('gr-notes').value.trim(),
      updated_at:new Date().toISOString()}]);
    if(r){this.audit('Grant saved','Payroll',name,id+' · ends '+(end||'—'));
      closeModal('grant-modal');toast('Grant saved ✓');this.renderGrants();}
    else $('gr-msg').innerHTML='<span style="color:var(--red)">Save failed: '+(API.lastError||'')+'</span>';
  }
  exportGrants(){
    const rows=this._grants||[];
    if(!rows.length)return toast('Nothing to export','err');
    let csv='Grant ID,Grant Name / Funding Source,PC/Local,Start (dd-mm-yyyy),End (dd-mm-yyyy),Months Left,Amount,Currency,Status,Focal Person\n';
    rows.forEach(g=>{const f=this._grantFlag(g.end_date);
      csv+=`"${g.id}","${(g.name||'').replace(/"/g,'""')}","${g.pc_local||''}","${this._grDate(g.start_date)}","${this._grDate(g.end_date)}","${f.months??''}",${g.amount||0},"${g.currency||''}","${g.status||''}","${g.focal_person||''}"\n`;});
    this._dl(csv,'THP_Grants_'+Date.now()+'.csv','text/csv');
  }

  /* ═══════════════════════════════════════════
     STAFF CONTRACT REMINDERS
     Visible to Admin, Edna (HR), and Agatha (CL)
     Excludes Interns and National Service personnel
  ═══════════════════════════════════════════ */
  _contractFlag(endDate){
    // Returns {cls, label, days} based on days until contract end
    if(!endDate)return{cls:'none',label:'No contract date',days:null};
    const end=new Date(endDate);if(isNaN(end))return{cls:'none',label:'Invalid date',days:null};
    const now=new Date();now.setHours(0,0,0,0);
    const days=Math.round((end-now)/86400000);
    if(days<0)return{cls:'red',label:'⚠ Expired '+Math.abs(days)+'d ago',days};
    if(days<=30)return{cls:'red',label:'🔴 Expires in '+days+'d',days};
    if(days<=60)return{cls:'amber',label:'🟠 Expires in '+days+'d',days};
    return{cls:'green',label:'🟢 '+days+'d remaining',days};
  }
  renderContracts(prefix){
    const p=prefix||'a-';
    const body=$(p+'contracts-body');if(!body)return;
    const summary=$(p+'contract-summary');
    const q=($(p+'contract-search')?.value||'').trim().toLowerCase();
    // Build contract list — include all staff (interns, NSS, and Country Leader included)
    let list=Object.entries(this.staff).filter(([id,s])=>{
      return (s.role||'')!=='admin';
    });
    if(q)list=list.filter(([id,s])=>id.toLowerCase().includes(q)||(s.name||'').toLowerCase().includes(q));
    // Sort: soonest expiry first, no-date staff last
    list.sort((a,b)=>{
      const ea=a[1].contractEnd,eb=b[1].contractEnd;
      if(!ea&&!eb)return a[1].name.localeCompare(b[1].name);
      if(!ea)return 1;if(!eb)return -1;
      return new Date(ea)-new Date(eb);
    });
    // Summary counts
    let red=0,amber=0,green=0,nodate=0;
    list.forEach(([id,s])=>{const f=this._contractFlag(s.contractEnd);if(f.cls==='red')red++;else if(f.cls==='amber')amber++;else if(f.cls==='green')green++;else nodate++;});
    if(summary)summary.innerHTML=
      `<div class="cs-box"><div class="cs-num" style="color:#dc2626">${red}</div><div class="cs-lbl">Expiring / Expired</div></div>`+
      `<div class="cs-box"><div class="cs-num" style="color:#d97706">${amber}</div><div class="cs-lbl">Within 60 Days</div></div>`+
      `<div class="cs-box"><div class="cs-num" style="color:#16a34a">${green}</div><div class="cs-lbl">Active</div></div>`+
      `<div class="cs-box"><div class="cs-num" style="color:var(--text3)">${nodate}</div><div class="cs-lbl">No Date Set</div></div>`;
    if(!list.length){body.innerHTML='<tr><td colspan="6"><div class="empty"><div class="empty-ico">📭</div>No staff found</div></td></tr>';return;}
    body.innerHTML=list.map(([id,s])=>{
      const f=this._contractFlag(s.contractEnd);
      const start=s.contractStart?fmtISO(s.contractStart):'—';
      const end=s.contractEnd?fmtISO(s.contractEnd):'—';
      return `<tr><td><strong>${s.name}</strong><br><span style="font-size:.72rem;color:var(--text3)">${id}</span></td>`+
        `<td style="font-size:.8rem">${s.unit||'—'}</td>`+
        `<td style="font-size:.8rem">${start}</td>`+
        `<td style="font-size:.8rem">${end}</td>`+
        `<td><span class="c-flag ${f.cls}">${f.label}</span></td>`+
        `<td><button class="bsm bsm-navy" onclick="APP.openContractModal('${id}')">✏ Edit</button></td></tr>`;
    }).join('');
  }
  openContractModal(id){
    const s=this.staff[id];if(!s)return;
    $('cm-id').value=id;
    $('cm-staff-name').textContent=s.name+' ('+id+')';
    $('cm-start').value=s.contractStart?String(s.contractStart).slice(0,10):'';
    $('cm-end').value=s.contractEnd?String(s.contractEnd).slice(0,10):'';
    $('cm-msg').textContent='';
    $('contract-modal').classList.add('open');
  }
  async saveContract(){
    const id=$('cm-id').value;
    const start=$('cm-start').value||'';
    const end=$('cm-end').value||'';
    const msg=$('cm-msg');
    if(end&&start&&new Date(end)<new Date(start)){if(msg)msg.innerHTML='<span style="color:var(--red)">End date is before start date.</span>';return;}
    if(msg)msg.innerHTML='<span style="color:var(--teal)">⏳ Saving…</span>';
    const r=await API.updateContract(id,start,end);
    if(r&&r.success){
      this.staff[id].contractStart=start;
      this.staff[id].contractEnd=end;
      this._cacheS();
      this.audit('Contract updated','HR',this.staff[id]?.name||id,(start||'—')+' → '+(end||'—'));
      closeModal('contract-modal');
      // Re-render whichever panel is active
      this.renderContracts('a-');this.renderContracts('m-');
      toast('Contract dates saved ✓');
    } else {
      if(msg)msg.innerHTML='<span style="color:var(--red)">Save failed. Try again.</span>';
    }
  }

  /* ── Contract-expiry email is now handled by the GAS scheduled
     job `checkContractExpiry` (60-day + 30-day + expired notices).
     The frontend no longer sends emails to avoid duplicates — it
     only renders the visual Contracts panel. ── */
  _checkContractReminders(){
    // Intentionally a no-op: GAS time-driven trigger sends the emails.
    return;
  }

  /* ═══════════════════════════════════════════
     COUNTRY LEADER DELEGATION
     Allows CL to delegate leave approval to another manager
  ═══════════════════════════════════════════ */
  _isActiveDelegate(uid){
    try{
      const d=JSON.parse(localStorage.getItem('thp_delegation')||'null');
      if(!d||!d.active||d.delegateId!==uid)return false;
      const now=new Date().toISOString().slice(0,10);
      return now>=d.startDate&&now<=d.endDate;
    }catch(e){return false;}
  }

  _getActiveDelegate(){
    try{
      const d=JSON.parse(localStorage.getItem('thp_delegation')||'null');
      if(!d||!d.active)return null;
      const now=new Date().toISOString().slice(0,10);
      if(now>=d.startDate&&now<=d.endDate)return d;
      return null;
    }catch(e){return null;}
  }

  async renderDelegation(prefix){
    const p=prefix||'';
    const statusEl=$(p+'deleg-status')||$('deleg-status');
    const sel=$(p+'deleg-person')||$('deleg-person');
    if(!sel)return;

    // Populate delegate dropdown — ONLY managers/supervisors (exclude CL herself)
    const managers=Object.entries(this.staff).filter(([id,s])=>{
      if(id===COUNTRY_LEADER_ID)return false;
      const r=(s.role||'staff').toLowerCase().trim();
      return r==='manager'||r==='country_leader';
    }).sort((a,b)=>a[1].name.localeCompare(b[1].name));
    sel.innerHTML='<option value="">— Select a manager —</option>'+
      (managers.length
        ? managers.map(([id,s])=>`<option value="${id}">${s.name} (${s.unit||'—'})</option>`).join('')
        : '<option value="" disabled>⚠ No managers found — check staff roles in admin panel</option>');

    // Load current delegation from Supabase settings
    const settings=await API._get('settings','key=eq.cl_delegation');
    let deleg=null;
    if(settings&&settings.length){
      try{deleg=JSON.parse(settings[0].value);}catch(e){}
    }
    if(deleg&&deleg.active){
      localStorage.setItem('thp_delegation',JSON.stringify(deleg));
      const delegName=this.staff[deleg.delegateId]?.name||deleg.delegateId;
      const now=new Date().toISOString().slice(0,10);
      const isActive=now>=deleg.startDate&&now<=deleg.endDate;
      if(statusEl)statusEl.innerHTML=`<span style="color:${isActive?'var(--green)':'var(--gold)'}">● ${isActive?'Active':'Scheduled'} Delegation</span><br>
        <strong>${delegName}</strong> can approve leave on behalf of the Country Leader<br>
        <span style="font-size:.76rem;color:var(--text3)">${fmtISO(deleg.startDate)} → ${fmtISO(deleg.endDate)}</span>`;
      sel.value=deleg.delegateId;
      const startEl=$(p+'deleg-start')||$('deleg-start');if(startEl)startEl.value=deleg.startDate;
      const endEl=$(p+'deleg-end')||$('deleg-end');if(endEl)endEl.value=deleg.endDate;
    } else {
      localStorage.removeItem('thp_delegation');
      if(statusEl)statusEl.innerHTML='<span style="color:var(--text3)">● No active delegation</span><br><span style="font-size:.78rem">The Country Leader is currently the sole final approver for all leave requests.</span>';
    }
  }

  async saveDelegation(prefix){
    const p=prefix||'';
    const delegateId=($(p+'deleg-person')||$('deleg-person'))?.value;
    const startDate=($(p+'deleg-start')||$('deleg-start'))?.value;
    const endDate=($(p+'deleg-end')||$('deleg-end'))?.value;
    const msg=$(p+'deleg-msg')||$('deleg-msg');
    if(!delegateId){if(msg)msg.innerHTML='<span style="color:var(--red)">Select a manager.</span>';return;}
    if(!startDate||!endDate){if(msg)msg.innerHTML='<span style="color:var(--red)">Set start and end dates.</span>';return;}
    if(new Date(endDate)<new Date(startDate)){if(msg)msg.innerHTML='<span style="color:var(--red)">End date before start date.</span>';return;}

    const deleg={active:true,delegateId,startDate,endDate,updatedAt:new Date().toISOString()};
    if(msg)msg.innerHTML='<span style="color:var(--teal)">⏳ Saving…</span>';
    await API._upsert('settings',[{key:'cl_delegation',value:JSON.stringify(deleg),updated_at:new Date().toISOString()}]);
    localStorage.setItem('thp_delegation',JSON.stringify(deleg));

    // Notify the delegate via email
    const delegName=this.staff[delegateId]?.name||'';
    const delegEmail=this.staff[delegateId]?.email||'';
    if(delegEmail){
      API.gasPost({action:'delegationNotify',delegateName:delegName,delegateEmail:delegEmail,startDate,endDate,active:true}).catch(()=>{});
    }

    this.renderDelegation(prefix);
    if(msg)msg.innerHTML='<span style="color:var(--green)">✓ Delegation activated!</span>';
    toast(`${delegName} can now approve leave as delegate.`);
  }

  async deactivateDelegation(prefix){
    if(!confirm('Deactivate the current delegation?'))return;
    const deleg={active:false,delegateId:'',startDate:'',endDate:'',updatedAt:new Date().toISOString()};
    await API._upsert('settings',[{key:'cl_delegation',value:JSON.stringify(deleg),updated_at:new Date().toISOString()}]);
    localStorage.removeItem('thp_delegation');
    this.renderDelegation(prefix);
    toast('Delegation deactivated.');
  }

  /* ═══════════════════════════════════════════
     AUTO CLOCK-OUT AT MIDNIGHT
  ═══════════════════════════════════════════ */
  _startAutoClockOut(){
    setInterval(()=>{
      if(!this.user)return;
      const now=new Date();
      const rec=this.records.find(r=>r.id===this.user.id&&!r.out);
      if(!rec)return;
      const clockInDate=new Date(rec.in).toISOString().slice(0,10);
      const todayDate=now.toISOString().slice(0,10);
      if(clockInDate!==todayDate){
        const midnight=new Date(clockInDate+'T23:59:59');
        const hrs=(midnight-new Date(rec.in))/3600000;
        rec.out=midnight.toISOString();rec.hours=fx(hrs);rec.status='Auto Clock-Out (Midnight)';
        API.updateRecord(rec).then(()=>{
          this._cacheR();
          const p=this._pfx();
          if($(p+'btn-co'))$(p+'btn-co').disabled=true;
          this._sess(false);this._stats();
          toast('⏰ Auto-clocked out at midnight.','info');
        });
      }
    },60000);
  }

  /* ═══════════════════════════════════════════
     MORNING CLOCK-IN REMINDER
  ═══════════════════════════════════════════ */
  _checkClockInReminder(){
    if(!this.user||this.user.role==='admin')return;
    const now=new Date();
    if(isWeekend(now)||isHoliday(now))return;
    const hour=now.getHours(),min=now.getMinutes();
    if(hour<8||(hour===8&&min<30))return;
    if(hour>12)return;
    const todayStr=todayISO();
    const onLeave=leaveOnDate(this.leave,this.user.id,todayStr);
    if(onLeave)return;
    const alreadyIn=this.records.find(r=>r.id===this.user.id&&((r.date||r.in||'').slice(0,10)===todayStr||(r.in&&new Date(r.in).toISOString().slice(0,10)===todayStr)));
    if(!alreadyIn){
      setTimeout(()=>toast('⏰ Reminder: You haven\'t clocked in today.','info'),2000);
    }
  }
}

const APP=new App();

/* ═══════════════════════════════════════════════
   8. SESSION RESTORE — Server-validated
   On page load, send stored token to server for
   validation instead of trusting localStorage.
═══════════════════════════════════════════════ */
(async function restoreSession(){
  try{
    const session=getSession();
    if(!session)return;
    const {id,token}=session;
    if(!id||!token)return;

    // Show loading overlay & hide login to prevent flash
    showLoader('Verifying your session…');
    const loginEl=$('login-view');
    if(loginEl)loginEl.style.display='none';

    /* ── SERVER VALIDATION ── */
    const result=await API.validateSession(id,token);
    if(!result||!result.success){
      // Invalid session — back to login
      clearSession();
      hideLoader();
      if(loginEl)loginEl.style.display='';
      return;
    }

    APP.user=result.user;

    /* Hydrate all data from server */
    const loT=$('lo-text');if(loT)loT.textContent='Loading your data…';
    const data=await API.hydrate();
    if(data&&data.success){
      APP.staff=data.staff||{};
      APP.records=data.records||[];
      APP.leave=data.leave||[];
      APP.holidays=data.holidays||[];
      APP._cacheH();
    }

    if(loT)loT.textContent='Setting up your dashboard…';
    const role=APP.user.role;

    if(role==='admin'){
      showView('admin-view');
      setTimeout(()=>{
        try{
        APP.renderAdmin();APP._renderDash();APP._renderStaffGrid();APP._renderReports();APP.renderAdminLeave();APP._updateNotifBadges();
        APP._populateSupervisorDropdown();APP._initEntQR();APP.renderAdminHolidays();
        APP._checkContractReminders();
        if($('script-url-input')&&API.getGasUrl())$('script-url-input').value=API.getGasUrl();
        
      }catch(e){console.error('Dashboard setup error:',e);toast('Some parts of the dashboard did not load. Refresh to try again.','err');}
        finally{hideLoader();}
      },100);
      API.updateChips();
      return;
    }

    if(isManagerRole(role)){
      showView('manager-view');
      setTimeout(()=>{
        try{
        if($('m-unit-display'))$('m-unit-display').textContent=APP.user.unit;
        APP._toggleMgrReports(id);APP._setLeaveTabLabel(id);
        if($('mgr-name'))$('mgr-name').textContent=APP.user.name;
        const av=$('mgr-av');if(av){av.textContent=ini(APP.user.name);av.style.background=APP.user.color||avColor(APP.user.name);}
        const mav=$('mob-mgr-av');if(mav){mav.textContent=ini(APP.user.name);mav.style.background=APP.user.color||avColor(APP.user.name);}
        const mn=$('mob-mgr-name');if(mn)mn.textContent=APP.user.name;
        APP._sessCheck();APP._initWorkModeListeners();APP._stats();APP._renderMgrDash();APP.renderMgrRecs();APP.loadLeave();APP._updateNotifBadges();
        if($('m-chpw-name'))$('m-chpw-name').textContent=APP.user.name;
        APP._checkDefaultPass('mgr');APP._renderProfileForm('m-');
        if(id===COUNTRY_LEADER_ID){const dn=$('nav-mgr-deleg');if(dn)dn.classList.remove('cl-only-tab');const dm=$('mob-mgr-deleg');if(dm)dm.classList.remove('cl-only-tab');}
        if(typeof APP!=='undefined'&&APP._applyPrivileges)APP._applyPrivileges(id);
        if(typeof APP!=='undefined'&&APP._checkContractReminders)APP._checkContractReminders();
        APP._startAutoClockOut();APP._checkClockInReminder();
        
      }catch(e){console.error('Dashboard setup error:',e);toast('Some parts of the dashboard did not load. Refresh to try again.','err');}
        finally{hideLoader();}
      },100);
    } else {
      showView('staff-view');
      setTimeout(()=>{
        try{
        $('st-name').textContent=APP.user.name;
        const av=$('st-av');if(av){av.textContent=ini(APP.user.name);av.style.background=APP.user.color||avColor(APP.user.name);}
        const mav=$('mob-st-av');if(mav){mav.textContent=ini(APP.user.name);mav.style.background=APP.user.color||avColor(APP.user.name);}
        const mn=$('mob-st-name');if(mn)mn.textContent=APP.user.name;
        APP._stats();APP.renderStaffLogs();APP._staffQR();APP._sessCheck();APP._initWorkModeListeners();APP._renderLeaveBal();APP.renderStaffLeave();APP._initLeaveForm();APP._updateNotifBadges();
        APP.renderStaffFeed();APP.checkBirthdayWish();
        (this._applyPrivileges?this:APP)._applyPrivileges(id);
        if($('unit-display'))$('unit-display').textContent=APP.user.unit;
        APP._filterLeaveByGender();APP._checkDefaultPass('');APP._renderProfileForm('');
        APP._startAutoClockOut();APP._checkClockInReminder();
        
      }catch(e){console.error('Dashboard setup error:',e);toast('Some parts of the dashboard did not load. Refresh to try again.','err');}
        finally{hideLoader();}
      },100);
    }
    API.updateChips();

    // Auto-refresh leave data every 60s from Supabase
    setInterval(async()=>{
      if(!APP.user)return;
      try{
        const rows=await API._get('leave_requests','order=applied_at.desc&limit=2000');
        if(rows){
          APP.leave=rows.map(r=>({id:r.id,staffId:r.staff_id,name:r.name,unit:(r.unit||'').trim(),type:r.type,
            startDate:r.start_date,endDate:r.end_date,days:r.days,reason:r.reason,sickNote:r.sick_note,
            staffEmail:r.staff_email||'',supervisorId:r.supervisor_id||'',supervisorStatus:r.supervisor_status||'Pending',
            supervisorNote:r.supervisor_note||'',finalApproverId:r.final_approver_id||'',
            finalApproverStatus:r.final_approver_status||'Pending',finalApproverNote:r.final_approver_note||'',
            status:r.overall_status||'Pending',hrStatus:r.final_approver_status||r.overall_status||'Pending',
            hrNote:r.final_approver_note||'',appliedAt:r.applied_at||'',updatedAt:r.updated_at||'',
            handoverNote:r.handover_note||'',compRef:r.comp_ref||''}));
          APP._cacheL();APP._updateNotifBadges();
        }
      }catch(e){}
    },60000);
  }catch(e){clearSession();}
})();

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

/* ── THEME ── */
function toggleTheme(){
  const d=document.documentElement;
  const isDark=d.getAttribute('data-theme')==='dark';
  d.setAttribute('data-theme',isDark?'light':'dark');
  $('theme-toggle').textContent=isDark?'🌙':'☀️';
  localStorage.setItem('thp_theme',isDark?'light':'dark');
}
(function initTheme(){
  const t=localStorage.getItem('thp_theme')||'dark';
  document.documentElement.setAttribute('data-theme',t);
  const btn=$('theme-toggle'); if(btn) btn.textContent=t==='dark'?'☀️':'🌙';
})();

/* ── LOADING OVERLAY ── */
function showLoader(msg){
  const el=$('loading-overlay');if(!el)return;
  if(msg){const t=$('lo-text');if(t)t.textContent=msg;}
  el.classList.remove('fade-out');
  el.classList.add('active');
}
function hideLoader(){
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
const SUPABASE={
  URL:'https://jhpqzkwzxprsnaczkyjq.supabase.co',  // ← paste your Project URL
  KEY:'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImpocHF6a3d6eHByc25hY3preWpxIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzQxOTE4NTMsImV4cCI6MjA4OTc2Nzg1M30.GKJz9EhxGP1wTQBiufLoVLxWOstx-9Z0MPWHxj2c8VM',                                 // ← paste your anon key
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
      if(String(pass)!==String(adminPass))return{success:false,error:'Incorrect password'};
      const token=this._genToken();
      await this._cleanSessions(id);
      await this._insert('sessions',{staff_id:id,token,expires_at:new Date(Date.now()+12*3600000).toISOString()});
      return{success:true,user:{id:'ADMIN01',name:'Administrator',role:'admin'},token};
    }

    // Staff login
    const rows=await this._get('staff','id=eq.'+encodeURIComponent(id));
    if(!rows||!rows.length)return{success:false,error:'Staff ID not found'};
    const s=rows[0];
    if(String(s.password)!==String(pass))return{success:false,error:'Incorrect password'};
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
        contractStart:s.contract_start||'',contractEnd:s.contract_end||''};
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
    if(r===null)return msg.innerHTML='<span style="color:var(--red)">Save failed. Try again.</span>';
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
      return`<div class="scard"><div class="scard-top"><div class="av" style="background:${col}">${ini(s.name)}</div><div class="s-info"><div class="s-name">${s.name}</div><div class="s-id">${id}</div></div></div><div class="s-meta"><span class="s-unit">${s.unit||'—'}</span><span class="s-role-badge role-${s.role||'staff'}">${roleLabel(s.role)}</span></div><div class="scard-btns"><button class="btn-edit" onclick="APP.openEdit('${id}')">✏ Edit</button><button class="btn-edit" style="background:rgba(245,166,35,.1);border-color:rgba(245,166,35,.2);color:var(--gold)" onclick="APP.adminResetPass('${id}')">🔑 Reset</button><button class="btn-del" onclick="APP.delStaff('${id}')">🗑</button></div></div>`;
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
    const staffData={name,unit,role,pass,color,email,gender,supervisor};
    /* SERVER FIRST */
    const r=await API.saveStaff(id,staffData);
    if(!r||!r.success){toast('Server error','err');return;}
    this.staff[id]=staffData;this._cacheS();this._renderStaffGrid();this._populateSupervisorDropdown();
    ['ns-id','ns-nm','ns-pw','ns-email'].forEach(i=>$(i).value='');
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
      let staffList=Object.entries(this.staff).filter(([,s])=>!EXCLUDED_UNITS.includes((s.unit||'').trim()));
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
    else{const EXCLUDED_UNITS=['National Service','Intern'];const staffList=Object.entries(this.staff).filter(([,s])=>!EXCLUDED_UNITS.includes((s.unit||'').trim()));const allDays=this._mgrRepDays();let csv='Staff ID,Date,Name,Unit,Status\n';allDays.forEach(dt=>{const dateStr=fmtD(dt.toISOString());const hol=isHoliday(dt);if(hol){const holName=getHolidayName(dt)||'Public Holiday';csv+=`"—","${dateStr}","ALL STAFF","—","Holiday — ${holName}"\n`;}else{staffList.forEach(([id,s])=>{const present=this.records.some(r=>r.id===id&&fmtD(r.date||r.in)===dateStr);const onLeave=present?null:leaveOnDate(this.leave,id,dt.toISOString().slice(0,10));csv+=`"${id}","${dateStr}","${s.name}","${s.unit}","${present?'Present':onLeave?'On Leave':'Absent'}"\n`;});}});this._dl(csv,'THP_Report_'+Date.now()+'.csv','text/csv');}
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
    const s=this.staff[id];if(!s)return;
    $('hf-id').value=id;
    $('hf-staff-name').textContent=s.name+' ('+id+') — '+(s.unit||'');
    $('hf-msg').textContent='Loading…';
    $('hrfile-modal').classList.add('open');
    const f=await API.getHRFile(id)||{};
    $('hf-dob').value=f.dob?String(f.dob).slice(0,10):'';
    $('hf-phone').value=f.phone||s.phone||'';
    $('hf-emergency').value=f.emergency_contact||s.emergencyContact||'';
    $('hf-nok').value=f.next_of_kin||'';
    $('hf-nok-phone').value=f.next_of_kin_phone||'';
    $('hf-ssnit').value=f.ssnit_number||'';
    $('hf-quals').value=f.qualifications||'';
    $('hf-notes').value=f.notes||'';
    $('hf-photo-url').value=f.photo_url||'';
    const _pv=$('hf-photo-prev');if(_pv){const _t=this._drivePhoto(f.photo_url);if(_t){_pv.src=_t;_pv.style.display='block';}else _pv.style.display='none';}
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
    docs.splice(i,1);
    $('hf-docs-json').value=JSON.stringify(docs);
    this._renderHRDocs(docs);
    toast('Removed — click Save File to confirm.','info');
  }
  async uploadHRDoc(){
    const fileInput=$('hf-doc-file');const id=$('hf-id').value;
    if(!fileInput?.files?.length)return toast('Choose a file first','err');
    const file=fileInput.files[0];
    if(file.size>5*1024*1024)return toast('File too large (max 5MB)','err');
    $('hf-msg').innerHTML='<span style="color:var(--teal)">⏳ Uploading…</span>';
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
      }else $('hf-msg').innerHTML='<span style="color:var(--red)">Upload failed. Try again.</span>';
    }catch(e){$('hf-msg').innerHTML='<span style="color:var(--red)">Upload error.</span>';}
  }
  async saveHRFile(){
    const id=$('hf-id').value;if(!id)return;
    const data={
      dob:$('hf-dob').value||null,
      phone:$('hf-phone').value.trim(),
      emergency_contact:$('hf-emergency').value.trim(),
      next_of_kin:$('hf-nok').value.trim(),
      next_of_kin_phone:$('hf-nok-phone').value.trim(),
      ssnit_number:$('hf-ssnit').value.trim(),
      qualifications:$('hf-quals').value.trim(),
      notes:$('hf-notes').value.trim(),
      photo_url:$('hf-photo-url').value||'',
      documents:$('hf-docs-json').value||'[]'
    };
    $('hf-msg').innerHTML='<span style="color:var(--teal)">⏳ Saving…</span>';
    const r=await API.saveHRFile(id,data);
    if(r&&r.success){
      this.audit('Staff file updated','HR',this.staff[id]?.name||id,'');
      closeModal('hrfile-modal');
      this.renderHRFiles('a-');this.renderHRFiles('m-');
      toast('Staff file saved ✓');
    }else $('hf-msg').innerHTML='<span style="color:var(--red)">Save failed. Try again.</span>';
  }

  /* ── Privileges (admin-assigned access) ── */
  _privDefaults(){return{hr:[COUNTRY_LEADER_ID,HR_MANAGER_ID],cases:[HR_MANAGER_ID],payroll:['THPG/05/2025','THPG/01/2026-3']};}
  async _fetchPriv(){
    let p={};
    try{const r=await API._get('settings','key=eq.privileges');if(r&&r.length&&r[0].value)p=JSON.parse(r[0].value);}catch(e){}
    const d=this._privDefaults();
    return{hr:p.hr||d.hr,cases:p.cases||d.cases,payroll:p.payroll||d.payroll};
  }
  async _applyPrivileges(id){
    const p=await this._fetchPriv();this._priv=p;
    if(p.hr.includes(id)){
      document.querySelectorAll('.contract-tab').forEach(e=>e.classList.remove('contract-tab'));
      document.querySelectorAll('.hr-tab').forEach(e=>e.classList.remove('hr-tab'));
      document.body.classList.add('hr-mode');
      this.renderHRDash('m-');
    }
    if(p.cases.includes(id))document.querySelectorAll('.cases-tab').forEach(e=>e.classList.remove('cases-tab'));
    if(p.payroll.includes(id))document.querySelectorAll('.payroll-tab').forEach(e=>e.classList.remove('payroll-tab'));
    // Sync sidebar groups AFTER privileges resolve — never on a timer,
    // otherwise a slow fetch leaves a granted section marked "empty".
    _restoreNavGroups();_hideEmptyNavGroups();
    setTimeout(()=>{_restoreNavGroups();_hideEmptyNavGroups();},150);
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
    else toast('Save failed','err');
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
        const pv=$('hf-photo-prev');pv.src=this._drivePhoto(r.fileUrl);pv.style.display='block';
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
    await API._delete('announcements','id=eq.'+encodeURIComponent(id));
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
    if(!confirm('Delete this document from the library?'))return;
    await API._delete('org_documents','id=eq.'+encodeURIComponent(id));
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
      if(d.getMonth()!==now.getMonth()||d.getDate()!==now.getDate())return;
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
    $('vc-desc').value=v?.description||'';
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
      closing_date:$('vc-closing').value||null,description:$('vc-desc').value.trim(),created_by:this.user.name}]);
    if(r){closeModal('vacancy-modal');toast('Vacancy saved ✓');this.renderRecruit('m-');}
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
    if(r){closeModal('applicant-modal');toast('Candidate saved ✓');this.renderRecruit('m-');}
    else $('ap-msg').innerHTML='<span style="color:var(--red)">Save failed: '+(API.lastError||'')+'</span>';
  }

  /* ── Staff self-assessment ── */
  async renderMyAppraisals(){
    const body=$('st-apr-body');if(!body)return;
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
  _ratingWord(v){return['—','Unsatisfactory','Below Criteria','Achieved Criteria','Above Criteria','Exceeded All'][Math.round(+v||0)]||'—';}
  _defaultKPAs(){return[
    {name:'Financial',weight:30,kpis:[{kpi:'',standard:'',self:0,mgr:0}],commentary:''},
    {name:'Customer',weight:30,kpis:[{kpi:'',standard:'',self:0,mgr:0}],commentary:''},
    {name:'Internal Business Process',weight:20,kpis:[{kpi:'',standard:'',self:0,mgr:0}],commentary:''},
    {name:'Learning & Growth',weight:10,kpis:[{kpi:'',standard:'',self:0,mgr:0}],commentary:''},
    {name:'Values & Behaviours',weight:10,kpis:[{kpi:'',standard:'',self:0,mgr:0}],commentary:''}];}
  async renderPerf(p){
    const body=$('m-perf-body');if(!body)return;
    body.innerHTML='<tr><td colspan="8" style="color:var(--text3)">Loading…</td></tr>';
    let rows=await API._get('performance_appraisals','order=review_date.desc.nullslast&limit=300')||[];
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
        <td><button class="bsm bsm-navy" onclick="APP.openPerfModal('${r.id}')">✏</button></td></tr>`;
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
    if(r){closeModal('train-modal');toast('Training saved ✓');this.renderTraining('m-');}
    else $('tr-msg').innerHTML='<span style="color:var(--red)">Save failed.</span>';
  }

  /* ── Lifecycle & Cases (HR + Admin only) ── */
  async renderCases(p){
    const body=$(p+'case-body');if(!body)return;
    body.innerHTML='<tr><td colspan="7" style="color:var(--text3)">Loading…</td></tr>';
    const rows=await API._get('hr_cases','order=opened_date.desc.nullslast&limit=300')||[];
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
    $('cs-type').value=r?.case_type||'Disciplinary';
    $('cs-status').value=r?.status||'Open';
    $('cs-opened').value=r?.opened_date?String(r.opened_date).slice(0,10):new Date().toISOString().slice(0,10);
    $('cs-closed').value=r?.closed_date?String(r.closed_date).slice(0,10):'';
    $('cs-summary').value=r?.summary||'';
    $('cs-outcome').value=r?.outcome||'';
    $('cs-msg').textContent='';
    $('case-modal').classList.add('open');
  }
  async saveCase(){
    const id=$('cs-id').value||this._uid('CS');
    const r=await API._upsert('hr_cases',[{id,staff_id:$('cs-staff').value,case_type:$('cs-type').value,status:$('cs-status').value,opened_date:$('cs-opened').value||null,closed_date:$('cs-closed').value||null,summary:$('cs-summary').value.trim(),outcome:$('cs-outcome').value.trim()}]);
    if(r){this.audit('HR case saved','HR',this._sName($('cs-staff').value),$('cs-type').value+' · '+$('cs-status').value);closeModal('case-modal');toast('Case saved ✓');this.renderCases('m-');}
    else $('cs-msg').innerHTML='<span style="color:var(--red)">Save failed.</span>';
  }

  /* ── Org Chart (editable organogram) ── */
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
  _payDefaults(){return{ssnitEmployeePct:5.5,ssnitEmployerPct:13,ssnitCeilingMonthly:69000,tier3MaxPct:16.5,extraReliefPct:5,
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
    const r=await API._get('payroll_settings','key=eq.main');
    let s=this._payDefaults();
    if(r&&r.length){try{s={...s,...JSON.parse(r[0].value)};}catch(e){}}
    this._payS=s;return s;
  }
  async savePaySettings(p){
    const s=await this._loadPaySettings();
    s.ssnitEmployeePct=+($(p+'ps-emp').value)||s.ssnitEmployeePct;
    s.ssnitEmployerPct=+($(p+'ps-empr').value)||s.ssnitEmployerPct;
    s.ssnitCeilingMonthly=+($(p+'ps-ceil').value)||s.ssnitCeilingMonthly;
    s.tier3MaxPct=+($(p+'ps-t3').value)||s.tier3MaxPct;
    s.extraReliefPct=+($(p+'ps-relief').value)||0;
    this._payS=s;
    const r=await API._upsert('payroll_settings',[{key:'main',value:JSON.stringify(s)}]);
    if(r){toast('Settings saved ✓');this.renderPayroll(p);}else toast('Save failed','err');
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
    const ssnitEmp=S.ssnitEmployeePct/100*capped;
    const tier3=Math.min(n(ps.tier3_pct),S.tier3MaxPct)/100*basic;   // Provident Fund
    const extraRelief=(+S.extraReliefPct||0)/100*basic;   // Tier 2 relief (matches THP payslips)
    let taxable=Math.max(0,basic-ssnitEmp-tier3-extraRelief+taxableExtras);
    let paye=0,rem=taxable;
    for(const b of S.bands){const chunk=b.w===null?rem:Math.min(rem,b.w);paye+=chunk*b.r;rem-=chunk;if(rem<=0)break;}
    if(ps.paye_override!==null&&ps.paye_override!==undefined&&ps.paye_override!=='')paye=n(ps.paye_override);
    const advance=n(ps.salary_advance),ug=n(ps.ug_credit),other=n(ps.other_deductions);
    const totalDed=ssnitEmp+tier3+paye+advance+ug+other;
    const net=gross-totalDed;
    const emprSSNIT=S.ssnitEmployerPct/100*capped;
    return{gross,taxA,nonTax,arrears,incent,bonus,ot,fuel,ssnitEmp,tier3,paye,advance,ug,other,totalDed,net,
      cost:gross+emprSSNIT,
      allowStr:allow.map(a=>`${a.n} ${this._ghs(a.a)}${a.tax?'':' (nt)'}`).join(', ')||'—'};
  }
  async renderPayroll(p){
    const body=$(p+'pay-body');if(!body)return;
    body.innerHTML='<tr><td colspan="10" style="color:var(--text3)">Loading…</td></tr>';
    const S=await this._loadPaySettings();
    $(p+'ps-emp').value=S.ssnitEmployeePct;$(p+'ps-empr').value=S.ssnitEmployerPct;
    $(p+'ps-ceil').value=S.ssnitCeilingMonthly;$(p+'ps-t3').value=S.tier3MaxPct;
    if($(p+'ps-relief'))$(p+'ps-relief').value=S.extraReliefPct??0;
    if(!$(p+'pay-month').value)$(p+'pay-month').value=new Date().toISOString().slice(0,7);
    const rows=await API._get('payroll_staff','select=*')||[];
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
      this._payCalc.push({id,name:s.name,unit:s.unit,email:s.email||'',designation:ps.designation||'',...c,basic:+ps.basic,alloc});
      return`<tr><td><strong>${s.name}</strong><br><span style="font-size:.72rem;color:var(--text3)">${id}</span></td><td>${this._ghs(ps.basic)}</td><td style="font-size:.72rem">${c.allowStr}</td><td>${this._ghs(c.gross)}</td><td>${this._ghs(c.ssnitEmp)}</td><td>${this._ghs(c.tier3)}</td><td>${this._ghs(c.paye)}</td><td><strong>${this._ghs(c.net)}</strong></td><td>${this._ghs(c.cost)}</td><td><button class="bsm bsm-navy" onclick="APP.openPayModal('${id}')">✏</button></td></tr>`;
    }).join('');
    const sm=$(p+'pay-summary');
    if(sm)sm.innerHTML=`<div class="cs-box"><div class="cs-num" style="font-size:1rem">${this._ghs(tot.gross)}</div><div class="cs-lbl">Gross</div></div>
      <div class="cs-box"><div class="cs-num" style="font-size:1rem">${this._ghs(tot.paye)}</div><div class="cs-lbl">PAYE</div></div>
      <div class="cs-box"><div class="cs-num" style="font-size:1rem">${this._ghs(tot.ssnitEmp)}</div><div class="cs-lbl">SSNIT (Emp)</div></div>
      <div class="cs-box"><div class="cs-num" style="font-size:1rem;color:var(--green)">${this._ghs(tot.net)}</div><div class="cs-lbl">Net Payout</div></div>
      <div class="cs-box"><div class="cs-num" style="font-size:1rem">${this._ghs(tot.cost)}</div><div class="cs-lbl">Employer Cost</div></div>`;
  }
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
    API.getHRFile(id).then(hf=>{
      if($('pm-id').value!==id)return;
      $('pm-bank').value=hf?.bank_name||'';
      const acct=hf?.bank_account||'';
      this._acctReal=acct;
      const el=$('pm-account'),btn=$('pm-acct-btn');
      if(acct){el.value=this._maskAcct(acct);el.readOnly=true;this._acctMasked=true;if(btn)btn.style.display='inline-block';}
      else{el.value='';el.readOnly=false;this._acctMasked=false;if(btn)btn.style.display='none';}
    });
    $('pm-allow').value=allow.map(a=>`${a.n} : ${a.a} : ${a.tax?'t':'n'}`).join('\n');
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
    const allow=$('pm-allow').value.split('\n').map(l=>l.trim()).filter(Boolean).map(l=>{
      const parts=l.split(':').map(x=>x.trim());
      return{n:parts[0]||'Allowance',a:+parts[1]||0,tax:(parts[2]||'t').toLowerCase()!=='n'};
    });
    const alloc=$('pm-alloc').value.split('\n').map(l=>l.trim()).filter(Boolean).map(l=>{
      const q=l.split(':').map(x=>x.trim());return{project:q[0]||'Unallocated',pct:+q[1]||0};});
    const totPct=alloc.reduce((t,x)=>t+x.pct,0);
    if(alloc.length&&Math.abs(totPct-100)>0.01)return $('pm-msg').innerHTML='<span style="color:var(--red)">Allocation must total 100% (currently '+totPct+'%).</span>';
    const ov=$('pm-paye-ov').value.trim();
    const r=await API._upsert('payroll_staff',[{staff_id:id,basic:+$('pm-basic').value||0,allowances:JSON.stringify(allow),
      tier3_pct:+$('pm-tier3').value||0,grade:$('pm-grade').value,cost_allocation:JSON.stringify(alloc),
      designation:$('pm-designation').value.trim(),
      arrears:+$('pm-arrears').value||0,incentives:+$('pm-incentives').value||0,bonus:+$('pm-bonus').value||0,
      overtime:+$('pm-overtime').value||0,fuel_allowance:+$('pm-fuel').value||0,
      salary_advance:+$('pm-advance').value||0,ug_credit:+$('pm-ug').value||0,other_deductions:+$('pm-other').value||0,
      paye_override:ov===''?null:+ov,
      updated_at:new Date().toISOString()}]);
    if(r){
      const acctVal=this._acctMasked?this._acctReal:$('pm-account').value.trim();
      await API._upsert('hr_staff_files',[{staff_id:id,bank_name:$('pm-bank').value.trim(),bank_account:acctVal}]);
      closeModal('pay-modal');toast('Pay setup saved ✓');this.renderPayroll('m-');this.renderPayroll('st-');
    }
    else $('pm-msg').innerHTML='<span style="color:var(--red)">Save failed.</span>';
  }
  /* ── Phase B: bank advice, statutory returns, allocation, payslips ── */
  _payGuard(){
    if(!this._payCalc||!this._payCalc.length){toast('Recalculate the payroll first','err');return false;}
    return true;
  }
  async exportBankAdvice(p){
    if(!this._payGuard())return;
    const month=$(p+'pay-month').value||'';
    const files=await API._get('hr_staff_files','select=staff_id,bank_name,bank_account')||[];
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
    if(!this._payGuard())return;
    const month=this._payMonthLabel(p);
    const withEmail=this._payCalc.filter(r=>r.email);
    const without=this._payCalc.filter(r=>!r.email);
    if(!withEmail.length)return toast('No staff have email addresses on file','err');
    if(!confirm('Send '+withEmail.length+' payslip(s) for '+month+'?\n\nEach person receives only their own payslip.'+(without.length?'\n\nNo email on file for: '+without.map(r=>r.name).join(', '):'')))return;
    showLoader('Sending payslips…');
    const payload=[];
    for(const r of withEmail){
      payload.push({name:r.name,id:r.id,email:r.email,unit:r.unit||'',month,
        net:r.net.toFixed(2),html:await this._payslipHTML(r,month)});
    }
    const res=await API.gasPost({action:'sendPayslips',month,slips:payload});
    hideLoader();
    if(res&&res.success)toast('Payslips sent: '+res.sent+(res.failed?(' · failed: '+res.failed):'')+' ✓');
    else toast('Payslip sending failed'+(res&&res.error?': '+res.error:' — check the Apps Script deployment'),'err');
  }
  /* ── Payslip in THP format ── */
  async _payslipHTML(r,month){
    const f=v=>(+v||0).toLocaleString('en-GH',{minimumFractionDigits:2,maximumFractionDigits:2});
    const dash=v=>(+v)?f(v):'-';
    let bank={};
    try{const hf=await API.getHRFile(r.id);bank=hf||{};}catch(e){}
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
        <div><div class="org">THE HUNGER PROJECT – GHANA</div>
          <div class="addr">PMB CT 7, Cantonments Accra, Ghana</div>
          <div class="mail">email: thpghana@thp.org</div></div>
        <div class="ptitle">PAYSLIP</div>
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
          </table></div>
        </div>
        <div class="cols" style="margin-top:16px">
          <div class="col"><table>
            <tr class="hd"><th>EARNINGS</th><th></th><th class="r">AMOUNT</th></tr>
            ${row('Basic Salary',f(r.basic))}${row('Arrears',dash(r.arrears))}${row('Incentives',dash(r.incent))}
            ${row('Bonus',dash(r.bonus))}${row('Over Time Pay',dash(r.ot))}${row('Fuel Allowance',dash(r.fuel))}
            ${r.taxA+r.nonTax>0?row('Other Allowances',f(r.taxA+r.nonTax)):''}
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
      <div class="sig"><div>Employee Signature :<span></span></div><div>Employer Signature :<span></span></div></div>
    </div></body></html>`;
  }
  async previewPayslip(p){
    if(!this._payGuard())return;
    const month=this._payMonthLabel(p);
    const r=this._payCalc[0];
    const html=await this._payslipHTML(r,month);
    const w=window.open('','_blank');
    w.document.write(html+'<div style="text-align:center;padding:14px" class="no-print"><button onclick="window.print()" style="padding:9px 24px;background:#2D3592;color:#fff;border:0;border-radius:8px;font-weight:600;cursor:pointer">🖨 Print / Save as PDF</button></div>');
    w.document.close();
    toast('Showing payslip for '+r.name+' — use Print to save as PDF','info');
  }
  _payMonthLabel(p){
    const v=$(p+'pay-month').value||new Date().toISOString().slice(0,7);
    const d=new Date(v+'-01');
    return d.toLocaleString('en',{month:'short'})+'-'+d.getFullYear();
  }
  async savePayrollRun(p){
    if(!this._payCalc||!this._payCalc.length)return toast('Nothing calculated yet','err');
    const month=$(p+'pay-month').value||new Date().toISOString().slice(0,7);
    const r=await API._upsert('payroll_runs',[{id:'RUN-'+month,month,data:JSON.stringify(this._payCalc),processed_at:new Date().toISOString()}]);
    if(r){this.audit('Payroll run saved','Payroll',month,this._payCalc.length+' staff');toast('Payroll run for '+month+' saved ✓');}else toast('Save failed','err');
  }
  /* ── Payroll Phase B: bank advice, statutory returns, payslips ── */
  exportPayrollCSV(p){
    if(!this._payCalc||!this._payCalc.length)return toast('Nothing to export','err');
    const month=$(p+'pay-month').value||'';
    let csv='Staff ID,Name,Unit,Basic,Gross,SSNIT Employee,Provident Fund,PAYE,Net Pay,Employer Cost\n';
    this._payCalc.forEach(r=>{csv+=`"${r.id}","${r.name}","${r.unit||''}",${r.basic.toFixed(2)},${r.gross.toFixed(2)},${r.ssnitEmp.toFixed(2)},${r.tier3.toFixed(2)},${r.paye.toFixed(2)},${r.net.toFixed(2)},${r.cost.toFixed(2)}\n`;});
    this._dl(csv,'THP_Payroll_'+month+'.csv','text/csv');
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
        APP.renderAdmin();APP._renderDash();APP._renderStaffGrid();APP._renderReports();APP.renderAdminLeave();APP._updateNotifBadges();
        APP._populateSupervisorDropdown();APP._initEntQR();APP.renderAdminHolidays();
        APP._checkContractReminders();
        if($('script-url-input')&&API.getGasUrl())$('script-url-input').value=API.getGasUrl();
        hideLoader();
      },100);
      API.updateChips();
      return;
    }

    if(isManagerRole(role)){
      showView('manager-view');
      setTimeout(()=>{
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
        APP._applyPrivileges(id);APP._checkContractReminders();
        APP._startAutoClockOut();APP._checkClockInReminder();
        hideLoader();
      },100);
    } else {
      showView('staff-view');
      setTimeout(()=>{
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
        hideLoader();
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

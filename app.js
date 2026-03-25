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
function showPanel(id,sbId,e){
  $(sbId).nextElementSibling.querySelectorAll('.panel').forEach(p=>p.classList.remove('active'));
  $(id).classList.add('active');
  document.querySelectorAll('#'+sbId+' .nav-item').forEach(n=>n.classList.remove('active'));
  if(e)e.currentTarget.classList.add('active');
  if(window.innerWidth<=768)$(sbId).classList.remove('open');
}
function toggleSB(id){$(id).classList.toggle('open');}
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
  async _supa(path,opts={}){
    try{
      const r=await fetch(SUPABASE.URL+'/rest/v1/'+path,{headers:this._headers(),...opts});
      if(!r.ok){console.warn('Supabase error:',r.status,await r.text());return null;}
      const text=await r.text();
      return text?JSON.parse(text):[];
    }catch(e){console.warn('Supabase fetch:',e);return null;}
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
    const r=await this._insert('attendance',{
      date:rec.date,staff_id:rec.id,name:rec.name,unit:(rec.unit||'').trim(),
      clock_in:rec.in,clock_out:rec.out||null,hours:rec.hours||null,status:rec.status||'Active'
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
      phone:data.phone||'',emergency_contact:data.emergencyContact||''
    }]);
    return r?{success:true}:{success:false};
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
      overall_status:leave.status||'Pending'
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
      {name:'African Union Day',date:year+'-05-25',type:'fixed',recurring:'yes'},
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
        supervisor:s.supervisor||'',phone:s.phone||'',emergencyContact:s.emergency_contact||''};
    });

    // Transform attendance rows
    const records=(attRows||[]).map(r=>({
      date:r.date,id:r.staff_id,name:r.name,unit:r.unit,
      in:r.clock_in,out:r.clock_out||null,hours:r.hours||null,status:r.status||'Active'
    }));

    // Transform leave rows
    const leave=(leaveRows||[]).map(r=>({
      id:r.id,staffId:r.staff_id,name:r.name,unit:(r.unit||'').trim(),type:r.type,
      startDate:r.start_date,endDate:r.end_date,days:r.days,reason:r.reason,sickNote:r.sick_note,
      staffEmail:r.staff_email||'',
      supervisorId:r.supervisor_id||'',supervisorStatus:r.supervisor_status||'Pending',supervisorNote:r.supervisor_note||'',
      finalApproverId:r.final_approver_id||'',finalApproverStatus:r.final_approver_status||'Pending',finalApproverNote:r.final_approver_note||'',
      status:r.overall_status||'Pending',hrStatus:r.final_approver_status||r.overall_status||'Pending',hrNote:r.final_approver_note||'',
      appliedAt:r.applied_at||'',updatedAt:r.updated_at||''
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
    `${year}-05-25`,               // African Union Day
    `${year}-07-01`,               // Republic Day
    `${year}-08-04`,               // Founders Day
    `${year}-09-21`,               // Kwame Nkrumah Memorial Day
    farmersDayISO(year),           // Farmer's Day (1st Friday Dec)
    `${year}-12-25`,               // Christmas
    `${year}-12-26`,               // Boxing Day
    eids.eidFitr,                  // Eid al-Fitr (estimated)
    eids.eidAdha,                  // Eid al-Adha (estimated)
  ]);
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
    [`${year}-05-25`]:'African Union Day',
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
const LEAVE_LIMITS={'Annual Leave':24,'Sick Leave':null,'Maternity Leave':65,'Paternity Leave':5,'Compassionate Leave':5};
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
        if($('script-url-input')&&API.getUrl())$('script-url-input').value=API.getUrl();
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
        this._sessCheck();this._stats();this._renderMgrDash();this.renderMgrRecs();this.loadLeave();this._updateNotifBadges();
        if($('m-chpw-name'))$('m-chpw-name').textContent=this.user.name;
        this._checkDefaultPass('mgr');this._renderProfileForm('m-');
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
        this._stats();this.renderStaffLogs();this._staffQR();this._sessCheck();this._renderLeaveBal();this.renderStaffLeave();this._initLeaveForm();this._updateNotifBadges();
        if($('unit-display'))$('unit-display').textContent=this.user.unit;
        this._filterLeaveByGender();this._checkDefaultPass('');this._renderProfileForm('');
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
  async clockIn(){
    const now=new Date();
    if(isWeekend(now))return toast('Not allowed on weekends.','err');
    if(isHoliday(now)){const hName=getHolidayName(now);return toast(`Today is a public holiday${hName?' — '+hName:''}.`,'info');}
    if(this.records.find(r=>r.id===this.user.id&&!r.out))return toast('Already clocked in','err');
    const todayStr=todayISO();
    const alreadyToday=this.records.find(r=>r.id===this.user.id&&((r.date||r.in||'').slice(0,10)===todayStr||(r.in&&new Date(r.in).toISOString().slice(0,10)===todayStr)));
    if(alreadyToday)return toast('Already clocked in today.','err');
    const onLeave=leaveOnDate(this.leave,this.user.id,todayStr);
    if(onLeave)return toast(`On approved ${onLeave.type} today.`,'info');

    const unit=(this.user.unit||'').trim();
    const rec={date:fmtD(now.toISOString()),id:this.user.id,name:this.user.name,unit,in:now.toISOString(),out:null,hours:null,status:'Active'};

    /* SERVER FIRST */
    const result=await API.saveRecord(rec);
    if(!result||!result.success){toast('Server error — try again','err');return;}

    this.records.push(rec);this._cacheR();
    const p=this._pfx();$(p+'btn-ci').disabled=true;$(p+'btn-co').disabled=false;this._sess(true);this._stats();
    toast('Clocked in at '+fmtT(now.toISOString()));
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
    if(txt)txt.textContent=on?'On Shift':'Signed Out';
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
    body.innerHTML=recs.map(r=>`<tr><td>${fmtD(r.date||r.in)}</td><td>${r.unit}</td><td>${fmtT(r.in)}</td><td>${r.out?fmtT(r.out):'<span style="color:var(--teal)">Active</span>'}</td><td>${r.hours||'--'}</td><td>${this._bdg(r.status)}</td></tr>`).join('');
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
    body.innerHTML=recs.map(r=>`<tr><td>${fmtD(r.date||r.in)}</td><td>${r.unit}</td><td>${fmtT(r.in)}</td><td>${r.out?fmtT(r.out):'<span style="color:var(--teal)">Active</span>'}</td><td>${r.hours||'--'}</td><td>${this._bdg(r.status)}</td></tr>`).join('');
  }

  /* ── Leave balances ── */
  _leaveDaysUsed(staffId,type){
    const yr=new Date().getFullYear();
    return this.leave.filter(l=>l.staffId===staffId&&l.type===type&&l.status==='Approved'&&new Date(l.startDate).getFullYear()===yr).reduce((a,l)=>a+parseInt(l.days||0),0);
  }
  _renderLeaveBal(){
    const gender=this.staff[this.user.id]?.gender||'';
    let types=['Annual Leave','Sick Leave','Paternity Leave','Compassionate Leave'];
    if(gender==='female')types=['Annual Leave','Sick Leave','Maternity Leave','Compassionate Leave'];
    const icons={'Annual Leave':'🌴','Sick Leave':'🏥','Maternity Leave':'👶','Paternity Leave':'👨‍👶','Compassionate Leave':'🕊'};
    $('st-leave-bal').innerHTML=`<h4>Leave Balances (${new Date().getFullYear()})</h4>`+
      types.map(t=>{
        const limit=LEAVE_LIMITS[t];const used=this._leaveDaysUsed(this.user.id,t);
        if(limit===null)return`<div class="bal-row"><div class="bal-icon">${icons[t]||'📋'}</div><div class="bal-info"><div class="bal-lbl">${t}</div><div style="font-size:.72rem;color:var(--text2)">As certified by medical professional</div></div><div class="bal-num">${used} days used</div></div>`;
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
      const isFinalApprover=uid===COUNTRY_LEADER_ID;
      const pending=isFinalApprover
        ? this.leave.filter(l=>l.finalApproverId===uid&&(l.finalApproverStatus==='Pending'||l.hrStatus==='Pending')&&(l.supervisorStatus==='Approved'||l.supervisorStatus==='N/A')).length
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
    let types=['Annual Leave','Sick Leave','Paternity Leave','Compassionate Leave'];
    if(gender==='female')types=['Annual Leave','Sick Leave','Maternity Leave','Compassionate Leave'];
    const icons={'Annual Leave':'🌴','Sick Leave':'🏥','Maternity Leave':'👶','Paternity Leave':'👨‍👶','Compassionate Leave':'🕊'};
    const el=$('mgr-leave-bal');if(!el)return;
    el.innerHTML=`<h4>Leave Balances (${new Date().getFullYear()})</h4>`+
      types.map(t=>{
        const limit=LEAVE_LIMITS[t];const used=this._leaveDaysUsed(this.user.id,t);
        if(limit===null)return`<div class="bal-row"><div class="bal-icon">${icons[t]||'📋'}</div><div class="bal-info"><div class="bal-lbl">${t}</div></div><div class="bal-num">${used} used</div></div>`;
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
      <div class="stat"><div class="stat-lbl">Team Size</div><div class="stat-val">${teamStaff.length}</div></div>
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
    const days=workingDaysBetween(start,end);
    if(days===0)return setErr('Dates fall on weekends/holidays.');
    const limit=LEAVE_LIMITS[type];
    if(limit!==null){
      const used=this.leave.filter(l=>l.staffId===this.user.id&&l.type===type&&l.status!=='Rejected').reduce((a,l)=>a+(parseInt(l.days)||0),0);
      if(used+days>limit)return setErr(`${limit-used} days left for ${type}.`);
    }
    const overlap=this.leave.find(l=>l.staffId===this.user.id&&l.type===type&&l.status!=='Rejected'&&new Date(l.startDate)<=new Date(end)&&new Date(l.endDate)>=new Date(start));
    if(overlap)return setErr('Overlapping request exists.');

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
      _supervisorEmail:this.staff[supervisorId]?.email||'',
      _finalApproverEmail:this.staff[finalApproverId]?.email||''
    };

    /* SERVER FIRST */
    const result=await API.applyLeave(lv);
    if(!result||!result.success){toast('Server error — try again','err');return;}

    /* Upload sick note to Google Drive if present */
    if(type==='Sick Leave'){
      const fileInput=$(p+'sick-file');
      if(fileInput&&fileInput.files.length){
        const file=fileInput.files[0];
        try{
          toast('Uploading medical document to Drive…','info');
          const base64=await this._fileToBase64(file);
          const uploadResult=await API.uploadSickNote(result.leaveId||id,file.name,base64,file.type);
          if(uploadResult&&uploadResult.success){
            lv.sickNote=file.name+' | '+uploadResult.fileUrl;
            lv.sickNoteUrl=uploadResult.fileUrl;
            lv.sickNoteDownload=uploadResult.downloadUrl;
            toast('Medical document uploaded to Drive ✓');
          } else {
            toast('Document saved locally but Drive upload failed','err');
          }
        }catch(e){console.warn('Sick note upload:',e);toast('Document upload error — leave still submitted','err');}
      }
    }

    this.leave.push(lv);this._cacheL();this._updateNotifBadges();
    if(isManagerRole(this.user.role))this.renderMgrMyLeave();else this.renderStaffLeave();
    // Clear form
    $(p+'start').value='';$(p+'end').value='';$(p+'reason').value='';
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
      return`<tr><td>${l.type}</td><td>${fmtISO(l.startDate)}</td><td>${fmtISO(l.endDate)}</td><td>${l.days}</td><td><div style="font-size:.7rem;color:var(--text3)">${supName}</div>${_bdg(l.supervisorStatus,true)}</td><td><div style="font-size:.7rem;color:var(--text3)">${finName}</div>${_bdg(l.finalApproverStatus||l.hrStatus)}</td><td>${_bdg(l.status)}</td><td style="font-size:.74rem;color:var(--text2)">${note}</td></tr>`;
    }).join('');
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
          hrNote:r.final_approver_note||'',appliedAt:r.applied_at||'',updatedAt:r.updated_at||''}));
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
    const isFinalApprover=uid===COUNTRY_LEADER_ID;
    const items=isFinalApprover
      ? this.leave.filter(l=>l.finalApproverId===uid&&(l.finalApproverStatus==='Pending'||l.hrStatus==='Pending')&&(l.supervisorStatus==='Approved'||l.supervisorStatus==='N/A'))
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
    const items=f?this.leave.filter(l=>l.status===f):this.leave;
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
    const isFinalApprover=uid===COUNTRY_LEADER_ID;
    const supName=this.staff[lv.supervisorId]?.name||'—';
    const finalName=this.staff[lv.finalApproverId]?.name||'—';
    const _bs=(s)=>s==='Approved'?'stage-ok':s==='Rejected'?'stage-rej':s==='N/A'?'stage-ok':'stage-pend';
    $('lm-title').textContent=(isFinalApprover?'✅ Final Approval — ':'👤 Supervisor Review — ')+lv.name;
    $('lm-info').innerHTML=`<strong>Type:</strong> ${lv.type} &nbsp; <strong>Days:</strong> ${lv.days}<br>
      <strong>Dates:</strong> ${fmtISO(lv.startDate)} → ${fmtISO(lv.endDate)}<br>
      <strong>Reason:</strong> ${lv.reason||'—'}<br>
      ${lv.sickNote?`<strong>Medical Doc:</strong> ${this._renderSickNoteLink(lv.sickNote)}<br>`:''}
      <strong>Supervisor (${supName}):</strong> <span class="stage-badge ${_bs(lv.supervisorStatus)}">${lv.supervisorStatus}</span><br>
      <strong>Final (${finalName}):</strong> <span class="stage-badge ${_bs(lv.finalApproverStatus||lv.hrStatus)}">${lv.finalApproverStatus||lv.hrStatus||'Pending'}</span>`;
    $('lm-note').value='';$('lm-id').value=id;$('leave-modal').classList.add('open');
  }

  /* ── Decide leave — SERVER FIRST ── */
  async decideLeave(status){
    const id=$('lm-id').value,note=$('lm-note').value.trim();
    const lv=this.leave.find(l=>l.id===id);if(!lv)return;
    const uid=this.user?.id;
    const isFinalApprover=uid===COUNTRY_LEADER_ID;
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
  }

  async saveProfile(prefix){
    const uid=this.user?.id;if(!uid)return;
    const p=prefix||'';
    const email=$(p+'prof-email')?.value.trim()||'';
    const phone=$(p+'prof-phone')?.value.trim()||'';
    const emergencyContact=$(p+'prof-emergency')?.value.trim()||'';
    const msgEl=$(p+'prof-msg');if(msgEl)msgEl.textContent='';

    if(msgEl)msgEl.innerHTML='<span style="color:var(--teal)">⏳ Saving…</span>';
    const r=await API.updateProfile(uid,{email,phone,emergencyContact});
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
  _mgrRepFilter(recs){
    const from=$('mgr-rep-from')?.value,to=$('mgr-rep-to')?.value;
    if(from)recs=recs.filter(r=>new Date(r.date||r.in)>=new Date(from));
    if(to)recs=recs.filter(r=>new Date(r.date||r.in)<=new Date(to+'T23:59:59'));
    return recs;
  }
  _mgrRepDays(){
    const from=$('mgr-rep-from')?.value,to=$('mgr-rep-to')?.value;
    const now=new Date(),y=now.getFullYear(),m=now.getMonth();
    const start=from?new Date(from):new Date(y,m,1);
    const end=to?new Date(to+'T23:59:59'):now;
    const days=[];
    for(let d=new Date(start);d<=end;d.setDate(d.getDate()+1)){if(isWorkingDay(d))days.push(new Date(d));}
    return days;
  }
  clearMgrRepDates(){if($('mgr-rep-from'))$('mgr-rep-from').value='';if($('mgr-rep-to'))$('mgr-rep-to').value='';this.renderMgrReport();}
  renderMgrReport(){
    const isHR=this.user.id===HR_MANAGER_ID;
    const hdr=$('m-report-hdr'),body=$('m-report-body');if(!hdr||!body)return;
    if(isHR){
      let recs=this._mgrRepFilter(this.records.slice());
      hdr.innerHTML='<th>Date</th><th>Staff ID</th><th>Name</th><th>Unit</th><th>Clock In</th><th>Clock Out</th><th>Hours</th><th>Status</th>';
      body.innerHTML=recs.length?recs.slice().reverse().map(r=>`<tr><td>${fmtD(r.date||r.in)}</td><td style="color:var(--text2);font-size:.76rem">${r.id}</td><td><strong>${r.name}</strong></td><td>${r.unit}</td><td>${fmtT(r.in)}</td><td>${r.out?fmtT(r.out):'Active'}</td><td>${r.hours||'--'}</td><td>${this._bdg(r.status)}</td></tr>`).join(''):'<tr><td colspan="8"><div class="empty"><div class="empty-ico">📭</div>No records</div></td></tr>';
    } else {
      const EXCLUDED_UNITS=['National Service','Intern'];
      const staffList=Object.entries(this.staff).filter(([,s])=>!EXCLUDED_UNITS.includes((s.unit||'').trim()));
      const allDays=this._mgrRepDays();
      const rows=[];
      allDays.forEach(dt=>{const dateStr=fmtD(dt.toISOString());staffList.forEach(([id,s])=>{const present=this.records.some(r=>r.id===id&&fmtD(r.date||r.in)===dateStr);const onLeave=present?null:leaveOnDate(this.leave,id,dt.toISOString().slice(0,10));rows.push({id,name:s.name,unit:s.unit,date:dateStr,dt,present,onLeave});});});
      rows.sort((a,b)=>b.dt-a.dt);
      hdr.innerHTML='<th>Staff ID</th><th>Date</th><th>Name</th><th>Unit</th><th>Status</th>';
      body.innerHTML=rows.length?rows.map(r=>{let badge;if(r.present)badge='<span class="badge b-ok">✓ Present</span>';else if(r.onLeave)badge=`<span class="badge" style="background:rgba(99,102,241,.15);color:#4338ca">🌴 ${r.onLeave.type}</span>`;else badge='<span class="badge b-err">✗ Absent</span>';return`<tr><td style="color:var(--text2);font-size:.76rem">${r.id}</td><td>${r.date}</td><td><strong>${r.name}</strong></td><td>${r.unit}</td><td>${badge}</td></tr>`;}).join(''):'<tr><td colspan="5"><div class="empty"><div class="empty-ico">📭</div>No data</div></td></tr>';
    }
  }
  exportMgrReport(){
    const isHR=this.user.id===HR_MANAGER_ID;
    if(isHR){let recs=this._mgrRepFilter(this.records.slice()).reverse();let csv='Date,Staff ID,Name,Unit,Clock In,Clock Out,Hours,Status\n';recs.forEach(r=>{csv+=`"${fmtD(r.date||r.in)}","${r.id}","${r.name}","${r.unit}","${fmtT(r.in)}","${r.out?fmtT(r.out):'Active'}","${r.hours||'--'}","${r.status}"\n`;});this._dl(csv,'THP_FullReport_'+Date.now()+'.csv','text/csv');}
    else{const EXCLUDED_UNITS=['National Service','Intern'];const staffList=Object.entries(this.staff).filter(([,s])=>!EXCLUDED_UNITS.includes((s.unit||'').trim()));const allDays=this._mgrRepDays();let csv='Staff ID,Date,Name,Unit,Status\n';allDays.forEach(dt=>{const dateStr=fmtD(dt.toISOString());staffList.forEach(([id,s])=>{const present=this.records.some(r=>r.id===id&&fmtD(r.date||r.in)===dateStr);const onLeave=present?null:leaveOnDate(this.leave,id,dt.toISOString().slice(0,10));csv+=`"${id}","${dateStr}","${s.name}","${s.unit}","${present?'Present':onLeave?'On Leave':'Absent'}"\n`;});});this._dl(csv,'THP_Report_'+Date.now()+'.csv','text/csv');}
  }
  printMgrReport(){
    const tbl=$('m-report-table');if(!tbl)return;
    const logoImg=document.querySelector('img[alt="THP-Ghana"],img[alt="Logo"],img[alt="THP"]');
    const w=window.open('','_blank');
    w.document.write(`<!DOCTYPE html><html><head><title>THP-Ghana Report</title><style>body{font-family:Arial;padding:20px}table{width:100%;border-collapse:collapse;font-size:12px}th{background:#2D3592;color:#fff;padding:8px}td{padding:7px;border-bottom:1px solid #eee}tr:nth-child(even) td{background:#f9f9f9}@media print{button{display:none}}</style></head><body>${tbl.outerHTML}<script>window.onload=()=>window.print()<\/script></body></html>`);
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
        APP._sessCheck();APP._stats();APP._renderMgrDash();APP.renderMgrRecs();APP.loadLeave();APP._updateNotifBadges();
        if($('m-chpw-name'))$('m-chpw-name').textContent=APP.user.name;
        APP._checkDefaultPass('mgr');APP._renderProfileForm('m-');
        hideLoader();
      },100);
    } else {
      showView('staff-view');
      setTimeout(()=>{
        $('st-name').textContent=APP.user.name;
        const av=$('st-av');if(av){av.textContent=ini(APP.user.name);av.style.background=APP.user.color||avColor(APP.user.name);}
        const mav=$('mob-st-av');if(mav){mav.textContent=ini(APP.user.name);mav.style.background=APP.user.color||avColor(APP.user.name);}
        const mn=$('mob-st-name');if(mn)mn.textContent=APP.user.name;
        APP._stats();APP.renderStaffLogs();APP._staffQR();APP._sessCheck();APP._renderLeaveBal();APP.renderStaffLeave();APP._initLeaveForm();APP._updateNotifBadges();
        if($('unit-display'))$('unit-display').textContent=APP.user.unit;
        APP._filterLeaveByGender();APP._checkDefaultPass('');APP._renderProfileForm('');
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
            hrNote:r.final_approver_note||'',appliedAt:r.applied_at||'',updatedAt:r.updated_at||''}));
          APP._cacheL();APP._updateNotifBadges();
        }
      }catch(e){}
    },60000);
  }catch(e){clearSession();}
})();

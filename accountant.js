(() => {
  const session=JSON.parse(localStorage.getItem("sa_session")||"null");
  if(!session || session.role!=="accountant"){ location.href="index.html"; return; }

  const $=s=>document.querySelector(s);
  const views=["dashboard","employees","import","attendance","reports","rules"];
  const state={month:new Date()};
  const titleMap={dashboard:"Dashboard",employees:"Employees",import:"Import Fingerprints",attendance:"Attendance",reports:"Reports",rules:"Rules"};

  function setView(name){
    views.forEach(v=>document.getElementById(`view-${v}`).classList.toggle("active",v===name));
    document.querySelectorAll(".nav-item[data-view]").forEach(b=>b.classList.toggle("active",b.dataset.view===name));
    $("#pageTitle").textContent=titleMap[name];
    if(name==="dashboard")renderDashboard();
    if(name==="employees")renderEmployees();
    if(name==="attendance")renderAttendance();
    if(name==="reports")renderReportSelector();
  }
  document.querySelectorAll(".nav-item[data-view]").forEach(b=>b.addEventListener("click",()=>setView(b.dataset.view)));
  document.querySelectorAll("[data-go]").forEach(b=>b.addEventListener("click",()=>setView(b.dataset.go)));
  $("#logoutBtn").onclick=()=>{localStorage.removeItem("sa_session");location.href="index.html"};

  function initMonth(){
    const sel=$("#monthSelect");
    const now=new Date();
    for(let i=11;i>=-1;i--){ const d=new Date(now.getFullYear(),now.getMonth()-i,1); const o=document.createElement("option");o.value=`${d.getFullYear()}-${d.getMonth()}`;o.textContent=d.toLocaleDateString("en-US",{month:"long",year:"numeric"});sel.appendChild(o); }
    sel.value=`${now.getFullYear()}-${now.getMonth()}`;
    sel.onchange=()=>{const [y,m]=sel.value.split("-").map(Number);state.month=new Date(y,m,1);renderDashboard();renderAttendance();renderReportSelector()};
  }
  function esc(x){return String(x??"").replace(/[&<>"']/g,m=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#039;"}[m]))}
  function monthRows(){
    const y=state.month.getFullYear(),m=state.month.getMonth(), emps=SAAttendance.getEmployees(), records=SAAttendance.getRecords();
    const byEmp=new Map();records.forEach(r=>{if(!byEmp.has(r.id))byEmp.set(r.id,[]);byEmp.get(r.id).push(r)});
    return emps.map(e=>{const days=SAAttendance.employeeMonth(e.code,y,m);return {e,days}});
  }
  function renderDashboard(){
    const emps=SAAttendance.getEmployees(), rec=SAAttendance.getRecords(), rows=monthRows();
    $("#kpiEmployees").textContent=emps.length;
    $("#kpiFingerprints").textContent=rec.length;
    let present=0,review=0; rows.forEach(x=>x.days.forEach(d=>{if(d.status==="Present"||d.status==="Late")present++;if(d.needsReview)review++}));
    $("#kpiPresent").textContent=present;$("#kpiReview").textContent=review;$("#reviewBadge").textContent=`${review} items`;
    const issues=[];rows.forEach(x=>x.days.filter(d=>d.needsReview).slice(0,3).forEach(d=>issues.push({name:x.e.name,code:x.e.code,date:d.date,status:d.status})));
    if(!issues.length){$("#reviewTable").innerHTML='<div class="empty-state">✓ Nothing needs review right now.</div>';return;}
    $("#reviewTable").innerHTML=`<table class="data-table"><thead><tr><th>Employee</th><th>Code</th><th>Date</th><th>Issue</th></tr></thead><tbody>${issues.map(i=>`<tr><td>${esc(i.name)}</td><td>${esc(i.code)}</td><td>${i.date}</td><td><span class="status ${i.status==="No fingerprints"?"warn":"danger"}">${esc(i.status)}</span></td></tr>`).join("")}</tbody></table>`;
  }
  function renderEmployees(){
    const q=($("#employeeSearch").value||"").toLowerCase(), emps=SAAttendance.getEmployees().filter(e=>[e.code,e.name,e.job,e.department].join(" ").toLowerCase().includes(q));
    $("#employeesTable").innerHTML=emps.length?`<table class="data-table"><thead><tr><th>Code</th><th>Name</th><th>Job</th><th>Department</th><th>Weekly off</th></tr></thead><tbody>${emps.map(e=>`<tr><td><strong>${esc(e.code)}</strong></td><td>${esc(e.name)}</td><td>${esc(e.job)}</td><td>${esc(e.department)}</td><td>${esc(e.weeklyOff||"Friday")}</td></tr>`).join("")}</tbody></table>`:'<div class="empty-state">No employee data found.</div>';
  }
  $("#employeeSearch").addEventListener("input",renderEmployees);

  $("#excelInput").addEventListener("change",async e=>{
    const file=e.target.files[0]; if(!file)return;
    try{
      const data=await file.arrayBuffer(), wb=XLSX.read(data,{type:"array",cellDates:true});
      const result=SAAttendance.importWorkbook(wb);
      $("#importResult").innerHTML=`<div class="panel"><strong>✓ Import complete</strong><br><span class="muted">Source sheet: ${esc(result.sheetName)} · Fingerprints: ${result.records.length} · Employees imported: ${result.employees.length}</span></div>`;
      renderDashboard();
    }catch(err){$("#importResult").innerHTML=`<div class="status danger">Import failed: ${esc(err.message)}</div>`}
  });

  function renderAttendance(){
    const q=($("#attendanceSearch").value||"").toLowerCase(), rows=monthRows().filter(x=>[x.e.code,x.e.name].join(" ").toLowerCase().includes(q));
    if(!rows.length){$("#attendanceTable").innerHTML='<div class="empty-state">Import employee/fingerprint data first.</div>';return;}
    $("#attendanceTable").innerHTML=`<table class="data-table"><thead><tr><th>Code</th><th>Employee</th><th>Present</th><th>Late</th><th>OT</th><th>Review</th></tr></thead><tbody>${rows.map(x=>{let p=0,l=0,ot=0,r=0;x.days.forEach(d=>{if(d.status==="Present")p++;if(d.status==="Late")l++;ot+=d.totalOT||0;if(d.needsReview)r++});return `<tr><td>${esc(x.e.code)}</td><td>${esc(x.e.name)}</td><td>${p}</td><td>${l}</td><td>${SAAttendance.fmtMin(ot)}</td><td><span class="status ${r?"warn":"ok"}">${r}</span></td></tr>`}).join("")}</tbody></table>`;
  }
  $("#attendanceSearch").addEventListener("input",renderAttendance);

  function renderReportSelector(){
    const sel=$("#reportEmployee"), old=sel.value, emps=SAAttendance.getEmployees();sel.innerHTML='<option value="">Select employee</option>'+emps.map(e=>`<option value="${esc(e.code)}">${esc(e.code)} — ${esc(e.name)}</option>`).join("");if(old)sel.value=old;
    if(sel.value)renderReport(sel.value);else $("#reportPanel").innerHTML='<div class="empty-state">Select an employee to view a report.</div>';
  }
  function renderReport(code){
    const e=SAAttendance.getEmployees().find(x=>x.code===code), days=SAAttendance.employeeMonth(code,state.month.getFullYear(),state.month.getMonth());
    let p=0,l=0,r=0,ot=0;days.forEach(d=>{if(d.status==="Present")p++;if(d.status==="Late")l++;if(d.needsReview)r++;ot+=d.totalOT||0});
    $("#reportPanel").innerHTML=`<div class="report-header"><div class="report-title"><p class="eyebrow">MONTHLY REPORT</p><h2>${esc(e?.name||code)}</h2><p>Code ${esc(code)} · ${esc(e?.job||"")} · ${esc(e?.department||"")}</p></div><div class="report-metrics"><div class="metric"><strong>${p}</strong><small>Present</small></div><div class="metric"><strong>${l}</strong><small>Late</small></div><div class="metric"><strong>${SAAttendance.fmtMin(ot)}</strong><small>Overtime</small></div><div class="metric"><strong>${r}</strong><small>Review</small></div></div></div><table class="data-table"><thead><tr><th>Date</th><th>Day</th><th>In</th><th>Out</th><th>Late</th><th>Early leave</th><th>OT</th><th>Status</th></tr></thead><tbody>${days.map(d=>`<tr><td>${d.date}</td><td>${d.dayName}</td><td>${d.first||"—"}</td><td>${d.last||"—"}</td><td>${d.late?SAAttendance.fmtMin(d.late):"—"}</td><td>${d.earlyLeave?SAAttendance.fmtMin(d.earlyLeave):"—"}</td><td>${d.totalOT?SAAttendance.fmtMin(d.totalOT):"—"}</td><td><span class="status ${d.statusClass}">${esc(d.status)}</span></td></tr>`).join("")}</tbody></table>`;
  }
  $("#reportEmployee").addEventListener("change",e=>renderReport(e.target.value));
  initMonth();renderDashboard();
})();
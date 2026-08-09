(() => {
  const session=JSON.parse(localStorage.getItem("sa_session")||"null");
  if(!session || session.role!=="employee"){location.href="index.html";return;}
  const $=s=>document.querySelector(s), code=session.code;
  $("#empName").textContent=session.name;$("#empMeta").textContent=`${session.code} · ${session.job||"Employee"} · ${session.site||"Site"}`;
  const sel=$("#monthSelect"),now=new Date();
  for(let i=11;i>=-1;i--){const d=new Date(now.getFullYear(),now.getMonth()-i,1),o=document.createElement("option");o.value=`${d.getFullYear()}-${d.getMonth()}`;o.textContent=d.toLocaleDateString("en-US",{month:"long",year:"numeric"});sel.appendChild(o);}
  sel.value=`${now.getFullYear()}-${now.getMonth()}`;
  function render(){const [y,m]=sel.value.split("-").map(Number),days=SAAttendance.employeeMonth(code,y,m);let p=0,l=0,r=0,ot=0;days.forEach(d=>{if(d.status==="Present")p++;if(d.status==="Late")l++;if(d.needsReview)r++;ot+=d.totalOT||0});$("#present").textContent=p;$("#late").textContent=l;$("#review").textContent=r;$("#overtime").textContent=SAAttendance.fmtMin(ot);$("#employeeTable").innerHTML=`<table class="data-table"><thead><tr><th>Date</th><th>Day</th><th>Check-in</th><th>Check-out</th><th>Late</th><th>Early</th><th>OT</th><th>Status</th></tr></thead><tbody>${days.map(d=>`<tr><td>${d.date}</td><td>${d.dayName}</td><td>${d.first||"—"}</td><td>${d.last||"—"}</td><td>${d.late?SAAttendance.fmtMin(d.late):"—"}</td><td>${d.earlyLeave?SAAttendance.fmtMin(d.earlyLeave):"—"}</td><td>${d.totalOT?SAAttendance.fmtMin(d.totalOT):"—"}</td><td><span class="status ${d.statusClass}">${d.status}</span></td></tr>`).join("")}</tbody></table>`;}
  sel.onchange=render;$("#logoutBtn").onclick=()=>{localStorage.removeItem("sa_session");location.href="index.html"};render();
})();
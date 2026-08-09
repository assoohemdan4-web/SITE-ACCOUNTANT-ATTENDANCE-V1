window.SAAttendance = (() => {
  const pad = n => String(n).padStart(2, "0");
  function minutes(t){ if(t==null || t==="") return null; const p=String(t).split(":").map(Number); return p[0]*60+(p[1]||0); }
  function fmtMin(m){ if(m==null || Number.isNaN(m)) return "—"; const sign=m<0?"-":""; m=Math.abs(Math.round(m)); return `${sign}${pad(Math.floor(m/60))}:${pad(m%60)}`; }
  function fmtDate(d){ return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`; }
  function parseDateTime(v){
    if(v instanceof Date && !Number.isNaN(v.getTime())) return v;
    if(typeof v === "number" && window.XLSX?.SSF){ const p=XLSX.SSF.parse_date_code(v); if(p) return new Date(p.y,p.m-1,p.d,p.H,p.M,p.S||0); }
    if(typeof v === "string"){
      const s=v.trim().replace("T"," ");
      const m=s.match(/^(\d{1,4})[\/\-](\d{1,2})[\/\-](\d{1,4})\s*(\d{1,2})?:?(\d{2})?:?(\d{2})?/);
      if(m){
        let a=+m[1],b=+m[2],c=+m[3]; let y,mn,d;
        if(a>31){ y=a; mn=b; d=c; } else { d=a; mn=b; y=c; }
        return new Date(y,mn-1,d,+(m[4]||0),+(m[5]||0),+(m[6]||0));
      }
      const d=new Date(v); if(!Number.isNaN(d.getTime())) return d;
    }
    return null;
  }
  function normalizeRow(row){
    const keys=Object.keys(row), lower=keys.map(k=>k.toLowerCase());
    const find = names => { for(const n of names){ const i=lower.findIndex(k=>k===n.toLowerCase() || k.includes(n.toLowerCase())); if(i>=0) return row[keys[i]]; } return null; };
    const id=find(["id","code","الكود"]);
    const raw=find(["date/time","datetime","date time","التاريخ والوقت"]);
    const dt=parseDateTime(raw);
    if(id==null || !dt) return null;
    return { id:String(id).trim(), datetime:dt, date:fmtDate(dt), time:`${pad(dt.getHours())}:${pad(dt.getMinutes())}`, raw:row };
  }
  function group(records){
    const map=new Map();
    records.forEach(r=>{ const k=`${r.id}|${r.date}`; if(!map.has(k)) map.set(k,[]); map.get(k).push(r); });
    for(const arr of map.values()) arr.sort((a,b)=>a.datetime-b.datetime);
    return map;
  }
  function scheduleFor(date){
    const s=SA_CONFIG.schedule[date.getDay()];
    return {...s};
  }
  function calculateDay(date, punches, decision=null){
    const s=scheduleFor(date), day={date:fmtDate(date), dayName:s.name, punches:punches||[], decision};
    if(!punches || punches.length===0){
      if(decision) return {...day,status:decision,statusClass:"neutral",needsReview:false,first:null,last:null,late:0,earlyLeave:0,earlyOT:0,lateOT:0,totalOT:0};
      return {...day,status:"No fingerprints",statusClass:"warn",needsReview:true,first:null,last:null,late:0,earlyLeave:0,earlyOT:0,lateOT:0,totalOT:0};
    }
    const first=punches[0], last=punches[punches.length-1], inM=minutes(first.time), outM=minutes(last.time);
    if(punches.length===1) return {...day,first:first.time,last:null,status:"Fingerprint lost",statusClass:"warn",needsReview:true,late:Math.max(0,inM-minutes(s.start)),earlyLeave:0,earlyOT:Math.min(Math.max(0,minutes(s.start)-inM),60),lateOT:0,totalOT:Math.min(Math.max(0,minutes(s.start)-inM),60)};
    if(outM<inM) return {...day,first:first.time,last:last.time,status:"Fingerprint error",statusClass:"danger",needsReview:true,late:0,earlyLeave:0,earlyOT:0,lateOT:0,totalOT:0};
    const start=minutes(s.start), end=minutes(s.end);
    let late=Math.max(0,inM-start), earlyLeave=Math.max(0,end-outM);
    // Friday is off by default, but any actual punches make it an attendance day 08:00-16:00.
    const effectiveStart = date.getDay()===5 ? 480 : start;
    const effectiveEnd = date.getDay()===5 ? 960 : end;
    late=Math.max(0,inM-effectiveStart); earlyLeave=Math.max(0,effectiveEnd-outM);
    const earlyOT=Math.min(Math.max(0,effectiveStart-inM),60);
    const lateOT=Math.min(Math.max(0,outM-effectiveEnd),60);
    const status = (late||earlyLeave) ? "Late" : "Present";
    return {...day,first:first.time,last:last.time,status,statusClass:status==="Present"?"ok":"warn",needsReview:false,late,earlyLeave,earlyOT,lateOT,totalOT:earlyOT+lateOT,actualMinutes:outM-inM};
  }
  function monthDates(year, monthIndex){ const arr=[],d=new Date(year,monthIndex,1); while(d.getMonth()===monthIndex){arr.push(new Date(d));d.setDate(d.getDate()+1)} return arr; }
  function getRecords(){ return JSON.parse(localStorage.getItem("sa_records")||"[]"); }
  function saveRecords(r){ localStorage.setItem("sa_records",JSON.stringify(r)); }
  function getEmployees(){ return JSON.parse(localStorage.getItem("sa_employees")||"[]"); }
  function saveEmployees(e){ localStorage.setItem("sa_employees",JSON.stringify(e)); }
  function importWorkbook(workbook){
    let sheetName=workbook.SheetNames.find(n=>["السحب","attendance","fingerprints","data"].some(x=>n.toLowerCase().includes(x))) || workbook.SheetNames[0];
    const sheet=workbook.Sheets[sheetName], rows=XLSX.utils.sheet_to_json(sheet,{defval:null,raw:true});
    const records=rows.map(normalizeRow).filter(Boolean);
    saveRecords(records);
    // Try to import employee data from a sheet with a recognizable header.
    let employees=[];
    for(const name of workbook.SheetNames){
      const sh=workbook.Sheets[name], rs=XLSX.utils.sheet_to_json(sh,{defval:null,raw:true});
      if(!rs.length) continue;
      const keys=Object.keys(rs[0]).map(x=>x.toLowerCase());
      if(keys.some(k=>k.includes("الكود")) && keys.some(k=>k.includes("الاسم"))){
        employees=rs.map(r=>({code:String(r["الكود"]??r["Code"]??"").trim(),name:String(r["الاسم"]??r["Name"]??"").trim(),job:String(r["الوظيفة"]??r["Job"]??"").trim(),department:String(r["القسم"]??r["Department"]??"").trim(),weeklyOff:String(r["الراحة الاسبوعية"]??"الجمعة") })).filter(e=>e.code);
        break;
      }
    }
    if(employees.length) saveEmployees(employees);
    return {sheetName,records,employees};
  }
  function employeeMonth(code, year, monthIndex){
    const records=getRecords().filter(r=>r.id===String(code));
    const byDate=new Map();
    records.forEach(r=>{if(!byDate.has(r.date))byDate.set(r.date,[]);byDate.get(r.date).push(r)});
    const rows=monthDates(year,monthIndex).map(d=>calculateDay(d,byDate.get(fmtDate(d))||[]));
    return rows;
  }
  return {minutes,fmtMin,parseDateTime,group,calculateDay,monthDates,getRecords,getEmployees,saveEmployees,importWorkbook,employeeMonth};
})();
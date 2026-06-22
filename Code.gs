// ============================================================
// TC LAB MANAGER — Code.gs (PIR + PGR + WASHING + DISPATCH + STORE)
// LockService added to all save functions (concurrent safe)
// Store: Category + Unit-in-master + unit conversion
// ============================================================

var CONFIG = {
  PIR_SHEET_PREFIX: "PIR-ROOM-",
  PIR_TOTAL_ROOMS:  10,
  USERS_SHEET:      "USERS",
  CROPS_SHEET:      "CROPS",
  CYCLES_SHEET:     "CYCLES",
  OPERATORS_SHEET:  "OPERATORS"
};

var PGR_SS_ID = "1F46wm8HKyPjxtsjSEthznpdIXvp-31KtUV4JXpH2ubU";

// ── ENTRY POINT ──────────────────────────────────────────
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('TC Lab Manager')
    .addMetaTag('viewport','width=device-width,initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── AUTH ─────────────────────────────────────────────────
function loginUser(username, password) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.USERS_SHEET);
    if (!sheet) return { success:false, message:"USERS sheet nahi mili." };
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      if (data[i][0].toString().trim() === username.trim() &&
          data[i][1].toString() === password &&
          data[i][5].toString().toLowerCase() !== 'no') {
        return {
          success  : true,
          username : data[i][0].toString().trim(),
          name     : data[i][2].toString(),
          role     : data[i][3].toString().toLowerCase(),
          dept     : data[i][4].toString() || 'ALL'
        };
      }
    }
    return { success:false, message:"Galat username ya password!" };
  } catch(e) { return { success:false, message:"Login error: "+e.message }; }
}

// ── INIT PIR ─────────────────────────────────────────────
function initializeApp() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  for (var r = 1; r <= CONFIG.PIR_TOTAL_ROOMS; r++) {
    var nm = CONFIG.PIR_SHEET_PREFIX + (r<10?"0"+r:r);
    if (!ss.getSheetByName(nm)) setupPirSheet(ss.insertSheet(nm));
  }

  if (!ss.getSheetByName(CONFIG.USERS_SHEET)) {
    var us = ss.insertSheet(CONFIG.USERS_SHEET);
    us.appendRow(["Username","Password","Name","Role","Dept","Active"]);
    us.getRange(1,1,1,6).setBackground("#1D9E75").setFontColor("#fff").setFontWeight("bold");
    us.appendRow(["admin",         "lab@123",  "Administrator",    "admin",            "ALL","yes"]);
    us.appendRow(["lab_incharge",  "lab@456",  "Lab Incharge",     "lab_incharge",     "ALL","yes"]);
    us.appendRow(["manager",       "mgr@789",  "Manager",          "manager",          "ALL","yes"]);
    us.appendRow(["sup_room01",    "sup@r01",  "Supervisor R1",    "supervisor",       "PIR","yes"]);
    us.appendRow(["conta_sup",     "cnt@r01",  "Conta Supervisor", "conta_supervisor", "PIR","yes"]);
    us.appendRow(["pgr_sup",       "pgr@123",  "PGR Supervisor",   "supervisor",       "PGR","yes"]);
    us.setFrozenRows(1);
  }

  if (!ss.getSheetByName(CONFIG.CROPS_SHEET)) {
    var cs = ss.insertSheet(CONFIG.CROPS_SHEET);
    cs.appendRow(["Crop","Varieties","CycleType"]);
    cs.getRange(1,1,1,3).setBackground("#1D9E75").setFontColor("#fff").setFontWeight("bold");
    cs.appendRow(["BANANA","G-09","BANANA"]);
    cs.appendRow(["BAMBOO","Balkua,Tulda,Golden","DEFAULT"]);
    cs.appendRow(["TEAK","India","DEFAULT"]);
    cs.setFrozenRows(1);
  }

  if (!ss.getSheetByName(CONFIG.CYCLES_SHEET)) {
    var cyc = ss.insertSheet(CONFIG.CYCLES_SHEET);
    cyc.appendRow(["Crop","Cycle","Active"]);
    cyc.getRange(1,1,1,3).setBackground("#1D9E75").setFontColor("#fff").setFontWeight("bold");
    ["S-00","S-01","S-02","S-03","M-01","M-02","M-03","M-05","M-06","SHOOTING","ROOTING"]
      .forEach(function(c){ cyc.appendRow(["BANANA",c,"yes"]); });
    for (var i=1;i<=20;i++) cyc.appendRow(["DEFAULT","C-"+(i<10?"0"+i:i),"yes"]);
    cyc.appendRow(["DEFAULT","SHOOTING","yes"]);
    cyc.appendRow(["DEFAULT","ROOTING","yes"]);
    cyc.setFrozenRows(1);
  }

  if (!ss.getSheetByName(CONFIG.OPERATORS_SHEET)) {
    var ops = ss.insertSheet(CONFIG.OPERATORS_SHEET);
    ops.appendRow(["Operator Name","Operator Code","Active"]);
    ops.getRange(1,1,1,3).setBackground("#1D9E75").setFontColor("#fff").setFontWeight("bold");
    ops.setFrozenRows(1);
  }

  return { success:true };
}

// ── PIR SHEET SETUP ───────────────────────────────────────
function setupPirSheet(sheet) {
  var h = [
    "Sr.No.","Date","Shift","Crop","Variety","LOT","Cycle",
    "Supervisor","Operator","Operator Code",
    "Used Bottle","Used Clams","Used Culture",
    "SH Bottle","SH Clums","SH Culture",
    "MUL Bottle","MUL Clums","MUL Culture",
    "RT Bottle","RT Clums","RT Culture",
    "TOT Bottle","TOT Culture","MR","Media Date",
    "Conta.Bottle","Conta.%","Conta.Date","Conta.By","Conta.Reason"
  ];
  sheet.appendRow(h);
  sheet.getRange(1,1,1,h.length)
    .setBackground("#1D9E75").setFontColor("#fff")
    .setFontWeight("bold").setFontSize(10);
  sheet.setFrozenRows(1);
  sheet.setColumnWidths(1,h.length,100);
}

// ── PIR SAVE ENTRIES (LOCK) ──────────────────────────────
function saveBulkEntries(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var rn    = parseInt(payload.room);
    var sName = CONFIG.PIR_SHEET_PREFIX+(rn<10?"0"+rn:rn);
    var sheet = ss.getSheetByName(sName);
    if (!sheet) { initializeApp(); sheet = ss.getSheetByName(sName); }
    var date  = new Date(payload.date);
    var saved = 0;
    payload.rows.forEach(function(row) {
      if (!row.operator || !row.operator.trim()) return;
      var uB  = n(row.usedBottle);
      var uCl = n(row.usedClams) || 4;
      var uC  = uB * uCl;
      var shB = n(row.shBottle),  shCl = n(row.shClums)  || 0, shC = shB * shCl;
      var mB  = n(row.mulBottle), mCl  = n(row.mulClums) || 0, mC  = mB  * mCl;
      var rB  = n(row.rtBottle),  rCl  = n(row.rtClums)  || 0, rC  = rB  * rCl;
      var tB  = shB+mB+rB;
      var tC  = shC+mC+rC;
      var mr  = uC > 0 ? Math.round((tC/uC)*100)/100 : 0;
      var sr  = sheet.getLastRow();
      var oCode = (row.operatorCode || "").toString().trim();
      sheet.appendRow([
        sr, date, payload.shift, payload.crop, payload.variety,
        payload.lot, payload.cycle, payload.supervisorName,
        row.operator.trim(), oCode,
        uB, uCl, uC,
        shB, shCl, shC,
        mB,  mCl,  mC,
        rB,  rCl,  rC,
        tB,  tC,   mr,
        payload.mediaDate||"",
        "","","","",""
      ]);
      var nr = sheet.getLastRow();
      sheet.getRange(nr,2).setNumberFormat("dd-mmm-yy");
      sheet.getRange(nr,25).setNumberFormat("0.00");
      if (nr%2===0) sheet.getRange(nr,1,1,31).setBackground("#E1F5EE");
      saved++;
    });
    SpreadsheetApp.flush();
    return { success:true, saved:saved, message:saved+" operators ki entries save ho gayi!" };
  } catch(e) { return { success:false, message:"Error: "+e.message }; }
  finally { try{lock.releaseLock();}catch(e2){} }
}

// ── PIR GET ROOM ENTRIES ──────────────────────────────────
function getPirRoomEntries(room, dateFrom, dateTo, supervisorFilter) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var rn    = parseInt(room);
    var sheet = ss.getSheetByName("PIR-ROOM-"+(rn<10?"0"+rn:rn));
    if (!sheet||sheet.getLastRow()<=1) return [];
    var tz    = Session.getScriptTimeZone();
    var fromStr = Utilities.formatDate(new Date(dateFrom), tz, "yyyy-MM-dd");
    var toStr   = Utilities.formatDate(new Date(dateTo),   tz, "yyyy-MM-dd");
    var data  = sheet.getDataRange().getValues();
    var out   = [];
    for (var i=1;i<data.length;i++) {
      var cell=data[i][1]; if (!cell) continue;
      var rd;
      if(cell instanceof Date){rd=new Date(cell.getTime());}else{var ds3=String(cell);if(ds3.indexOf('/')>0){var p3=ds3.split('/');if(p3.length===3)rd=new Date(p3[2],p3[1]-1,p3[0]);else rd=new Date(ds3);}else rd=new Date(ds3);}
      if(!rd||isNaN(rd.getTime()))continue;
      var rdStr = Utilities.formatDate(rd, tz, "yyyy-MM-dd");
      if (rdStr<fromStr||rdStr>toStr) continue;
      if (supervisorFilter&&supervisorFilter!=='ALL'&&
          data[i][7].toString().trim()!==supervisorFilter.trim()) continue;
      out.push(makeRow(data[i],rd,tz,rn));
    }
    return out;
  } catch(e) { return []; }
}

function makeRow(d,rd,tz,rn) {
  return {
    srNo:d[0], date:Utilities.formatDate(rd,tz,"dd-MMM-yy"),
    shift:d[2], crop:d[3], variety:d[4], lot:d[5], cycle:d[6],
    supervisor:d[7], operator:d[8], operatorCode:d[9],
    usedBottle:d[10], usedCulture:d[12],
    shBottle:d[13], shClums:d[14], shCulture:d[15],
    mulBottle:d[16], mulClums:d[17], mulCulture:d[18],
    rtBottle:d[19], rtClums:d[20], rtCulture:d[21],
    totBottle:d[22], totCulture:d[23],
    mr:n(d[24]).toFixed(2),
    contaBottle:d[26]!==""?d[26]:"—",
    contaPct:d[27]!==""?n(d[27]).toFixed(2)+"%":"⏳ Pending"
  };
}

function getPirPendingForExchange(room, date) {
  try {
    var tz = Session.getScriptTimeZone();
    var produced = {};
    var rn = parseInt(room);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PIR-ROOM-"+(rn<10?"0"+rn:rn));
    if (sheet && sheet.getLastRow() > 1) {
      var targStr = Utilities.formatDate(new Date(date), tz, "yyyy-MM-dd");
      var data = sheet.getDataRange().getValues();
      for (var i=1; i<data.length; i++) {
        var cell = data[i][1]; if (!cell) continue;
        var rd;
        if(cell instanceof Date){rd=new Date(cell.getTime());}else{var ds=String(cell);if(ds.indexOf('/')>0){var p=ds.split('/');if(p.length===3)rd=new Date(p[2],p[1]-1,p[0]);else rd=new Date(ds);}else rd=new Date(ds);}
        if(!rd||isNaN(rd.getTime()))continue;
        if (Utilities.formatDate(rd, tz, "yyyy-MM-dd") !== targStr) continue;
        var crop=String(data[i][3]||''), variety=String(data[i][4]||''), lot=String(data[i][5]||''), cycle=String(data[i][6]||'');
        var key = crop+'|'+variety+'|'+lot+'|'+cycle;
        if (!produced[key]) produced[key] = {crop:crop,variety:variety,lot:lot,cycle:cycle,bottles:0};
        produced[key].bottles += (parseFloat(data[i][22])||0);
      }
    }
    var sent = {};
    var exSheet = SpreadsheetApp.openById(PGR_SS_ID).getSheetByName("PGR_EXCHANGE");
    if (exSheet && exSheet.getLastRow() > 1) {
      var ed = exSheet.getDataRange().getValues();
      for (var j=1; j<ed.length; j++) {
        if (!ed[j][1]) continue;
        var len = ed[j].length;
        var crop2, variety2, lot2, cycle2, dir2, sentQty2;
        if (len >= 17) {
          lot2=String(ed[j][3]||''); crop2=String(ed[j][4]||''); variety2=String(ed[j][5]||'');
          cycle2=String(ed[j][6]||''); dir2=String(ed[j][7]||''); sentQty2=parseFloat(ed[j][10])||0;
        } else {
          lot2=String(ed[j][3]||''); crop2=''; variety2='';
          cycle2=String(ed[j][4]||''); dir2=String(ed[j][5]||''); sentQty2=parseFloat(ed[j][7])||0;
        }
        if (dir2.indexOf('PIR to PGR') < 0 && dir2.indexOf('PIR') !== 0) continue;
        var key2 = crop2+'|'+variety2+'|'+lot2+'|'+cycle2;
        sent[key2] = (sent[key2]||0) + sentQty2;
      }
    }
    var out = [];
    Object.keys(produced).forEach(function(k){
      var g = produced[k];
      var alreadySent = sent[k] || 0;
      var remaining = g.bottles - alreadySent;
      if (remaining > 0) {
        out.push({crop:g.crop,variety:g.variety,lot:g.lot,cycle:g.cycle,
          produced:g.bottles,sent:alreadySent,remaining:remaining});
      }
    });
    return {success:true, groups:out};
  } catch(e) { return {success:false, groups:[], message:e.message}; }
}

// ── PIR CONTAMINATION FETCH ───────────────────────────────
function getPendingContaEntries(room, date, supervisorFilter) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var rn    = parseInt(room);
    var sheet = ss.getSheetByName("PIR-ROOM-"+(rn<10?"0"+rn:rn));
    if (!sheet||sheet.getLastRow()<=1) return {success:true,rows:[]};
    var tz    = Session.getScriptTimeZone();
    var targStr = Utilities.formatDate(new Date(date), tz, "yyyy-MM-dd");
    var data  = sheet.getDataRange().getValues();
    var out   = [];
    for (var i=1;i<data.length;i++) {
      var cell=data[i][1]; if (!cell) continue;
      var rd;
      if(cell instanceof Date){rd=new Date(cell.getTime());}else{var ds3=String(cell);if(ds3.indexOf('/')>0){var p3=ds3.split('/');if(p3.length===3)rd=new Date(p3[2],p3[1]-1,p3[0]);else rd=new Date(ds3);}else rd=new Date(ds3);}
      if(!rd||isNaN(rd.getTime()))continue;
      var rdStr = Utilities.formatDate(rd, tz, "yyyy-MM-dd");
      if (rdStr!==targStr) continue;
      if (supervisorFilter&&supervisorFilter!=='ALL'&&
          data[i][7].toString().trim()!==supervisorFilter.trim()) continue;
      out.push({
        rowIndex:i+1, srNo:data[i][0], shift:data[i][2],
        crop:data[i][3], variety:data[i][4], lot:data[i][5], cycle:data[i][6],
        supervisor:data[i][7], operator:data[i][8], operatorCode:data[i][9],
        usedBottle:n(data[i][10]), usedCulture:n(data[i][12]),
        totBottle:n(data[i][22]), totCulture:n(data[i][23]),
        mr:n(data[i][24]).toFixed(2),
        contaBottle:data[i][26]!==""?data[i][26]:"",
        contaPct:data[i][27]!==""?n(data[i][27]).toFixed(2):"",
        contaReason:data[i][30]||""
      });
    }
    return {success:true,rows:out};
  } catch(e) { return {success:false,message:e.message}; }
}

// ── PIR CONTAMINATION SAVE (LOCK) ────────────────────────
function saveContaEntries(room, updates, contaBy) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var rn    = parseInt(room);
    var sheet = ss.getSheetByName("PIR-ROOM-"+(rn<10?"0"+rn:rn));
    if (!sheet) return {success:false,message:"Sheet nahi mili"};
    var tz    = Session.getScriptTimeZone();
    var today = Utilities.formatDate(new Date(),tz,"dd-MMM-yy");
    var saved=0, alerts=[];
    updates.forEach(function(u) {
      if (u.contaBottle==null||u.contaBottle==="") return;
      var ri   = parseInt(u.rowIndex);
      var totB = n(sheet.getRange(ri,23).getValue());
      var cB   = n(u.contaBottle);
      var cPct = totB>0?(cB/totB)*100:0;
      sheet.getRange(ri,27).setValue(cB);
      sheet.getRange(ri,28).setValue(cPct).setNumberFormat("0.00");
      sheet.getRange(ri,29).setValue(today);
      sheet.getRange(ri,30).setValue(contaBy||"Supervisor");
      sheet.getRange(ri,31).setValue(u.contaReason||"");
      if (cPct>5) {
        sheet.getRange(ri,27,1,5).setBackground("#FAECE7").setFontColor("#993C1D");
        alerts.push(sheet.getRange(ri,9).getValue()+" ("+cPct.toFixed(2)+"%)");
      }
      saved++;
    });
    SpreadsheetApp.flush();
    return {success:true,saved:saved,
      alertMsg:alerts.length?"⚠️ HIGH CONTA: "+alerts.join(", "):"",
      message:saved+" entries mein contamination save ho gaya!"};
  } catch(e) { return {success:false,message:"Error: "+e.message}; }
  finally { try{lock.releaseLock();}catch(e2){} }
}

// ── PIR DASHBOARD ────────────────────────────────────────
function getMainDashboard(dateFrom, dateTo) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var fromD = new Date(dateFrom); fromD.setHours(0,0,0,0);
  var toD   = new Date(dateTo);   toD.setHours(23,59,59,999);
  var entries=0,bottles=0,culture=0,usedBottles=0,pendC=0;
  var mrcSum=0,mrcCnt=0;
  var opSet={};
  var roomData=[];
  for (var r=1;r<=CONFIG.PIR_TOTAL_ROOMS;r++) {
    var rn    = r<10?"0"+r:""+r;
    var sheet = ss.getSheetByName("PIR-ROOM-"+rn);
    if (!sheet||sheet.getLastRow()<=1) {
      roomData.push({room:rn,entries:0,bottles:0,avgMr:"—",pendingConta:0}); continue;
    }
    var lastRow = sheet.getLastRow();
    var data = sheet.getRange(2, 2, lastRow-1, 27).getValues();
    var re=0,rb=0,rmSum=0,rmCnt=0,rp=0;
    for (var i=0;i<data.length;i++) {
      var cell=data[i][0]; if (!cell) continue;
      var rd;
      if(cell instanceof Date){rd=new Date(cell.getTime());}else{var ds3=String(cell);if(ds3.indexOf('/')>0){var p3=ds3.split('/');if(p3.length===3)rd=new Date(p3[2],p3[1]-1,p3[0]);else rd=new Date(ds3);}else rd=new Date(ds3);}
      if(!rd||isNaN(rd.getTime()))continue;
      rd.setHours(0,0,0,0);
      if (rd<fromD||rd>toD) continue;
      var usedBtl = parseFloat(data[i][9])||0;
      var opName = data[i][7]?String(data[i][7]).trim():'';
      var usedCult = parseFloat(data[i][11])||0;
      var totBtl = parseFloat(data[i][21])||0;
      var totCult = parseFloat(data[i][22])||0;
      var mr = parseFloat(data[i][23])||0;
      var contaCell = data[i][26];
      re++; rb+=totBtl; rmSum+=mr; rmCnt++;
      if (contaCell==="") rp++;
      entries++; bottles+=totBtl; culture+=totCult; usedBottles+=usedBtl;
      if (opName) opSet[opName.toLowerCase()]=true;
      if (usedCult>0){ mrcSum+=(totCult/usedCult); mrcCnt++; }
      if (contaCell==="") pendC++;
    }
    roomData.push({room:rn,entries:re,bottles:rb,pendingConta:rp,
      avgMr:rmCnt>0?(rmSum/rmCnt).toFixed(2):"—"});
  }
  var presentOps = Object.keys(opSet).length;
  var mrBottle = usedBottles>0 ? (bottles/usedBottles).toFixed(2) : "—";
  return {
    totalEntries:entries,
    totalBottles:bottles,
    totalCulture:culture,
    totalUsedBottle:usedBottles,
    contaPending:pendC,
    presentOperators:presentOps,
    avgMr:mrBottle,
    avgMrCulture:mrcCnt>0?(mrcSum/mrcCnt).toFixed(2):"—",
    outputBottle:presentOps>0?(bottles/presentOps).toFixed(0):"—",
    outputCulture:presentOps>0?(culture/presentOps).toFixed(0):"—",
    rooms:roomData
  };
}

// ── PIR FILTERED REPORT ───────────────────────────────────
function getPirFilteredReport(params) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var tz    = Session.getScriptTimeZone();
    var fromD = params.dateFrom?new Date(params.dateFrom):null;
    var toD   = params.dateTo?new Date(params.dateTo):null;
    if (fromD) fromD.setHours(0,0,0,0);
    if (toD)   toD.setHours(23,59,59,999);
    var rooms=[];
    if (params.room==="ALL") {
      for (var r=1;r<=CONFIG.PIR_TOTAL_ROOMS;r++) rooms.push(r<10?"0"+r:""+r);
    } else { var rv=parseInt(params.room); rooms.push(rv<10?"0"+rv:""+rv); }
    var allRows=[];
    rooms.forEach(function(rn) {
      var sheet=ss.getSheetByName("PIR-ROOM-"+rn);
      if (!sheet||sheet.getLastRow()<=1) return;
      var data=sheet.getDataRange().getValues();
      for (var i=1;i<data.length;i++) {
        var cell=data[i][1]; if (!cell) continue;
        var rd=cell instanceof Date?new Date(cell.getTime()):new Date(cell);
        rd.setHours(0,0,0,0);
        if (fromD&&rd<fromD) continue;
        if (toD&&rd>toD) continue;
        if (params.supervisorFilter&&params.supervisorFilter!=='ALL'&&
            data[i][7].toString().trim()!==params.supervisorFilter) continue;
        if (params.shift&&params.shift!=="ALL"&&data[i][2]!==params.shift) continue;
        if (params.operator&&params.operator.trim()&&
            data[i][8].toLowerCase().indexOf(params.operator.toLowerCase())<0) continue;
        if (params.crop&&params.crop!=="ALL"&&data[i][3]!==params.crop) continue;
        if (params.variety&&params.variety!=="ALL"&&
            String(data[i][4]).trim()!==String(params.variety).trim()) continue;
        if (params.lot&&params.lot.trim()&&
            String(data[i][5]).toLowerCase().indexOf(params.lot.toLowerCase())<0) continue;
        if (params.cycle&&params.cycle.trim()&&
            String(data[i][6]).toLowerCase().indexOf(params.cycle.toLowerCase())<0) continue;
        if (params.contaByFilter&&params.contaByFilter!=='ALL'&&
            data[i][29].toString().trim()!==params.contaByFilter.trim()) continue;
        allRows.push({
          srNo:data[i][0], room:rn, date:Utilities.formatDate(rd,tz,"dd-MMM-yy"), dateRaw:rd.getTime(),
          shift:data[i][2], crop:data[i][3], variety:data[i][4],
          lot:String(data[i][5]), cycle:data[i][6],
          supervisor:data[i][7], operator:data[i][8], operatorCode:data[i][9],
          usedBottle:n(data[i][10]), usedCulture:n(data[i][12]),
          shBottle:n(data[i][13]), shClums:n(data[i][14]), shCulture:n(data[i][15]),
          mulBottle:n(data[i][16]), mulClums:n(data[i][17]), mulCulture:n(data[i][18]),
          rtBottle:n(data[i][19]), rtClums:n(data[i][20]), rtCulture:n(data[i][21]),
          totBottle:n(data[i][22]), totCulture:n(data[i][23]),
          mr:n(data[i][24]), contaFilled:data[i][27]!=="",
          contaPct:n(data[i][27]), contaBottle:n(data[i][26]), contaBy:data[i][29]||""
        });
      }
    });
    allRows.sort(function(a,b){return a.dateRaw-b.dateRaw;});
    if (params.groupBy==="detail")
      return {success:true,type:"detail",rows:allRows,count:allRows.length};
    var groups={};
    allRows.forEach(function(r) {
      var key;
      switch(params.groupBy) {
        case "daily":      key=r.date; break;
        case "monthly":    key=r.date.substr(3); break;
        case "operator":   key=r.operator+(r.operatorCode?" ("+r.operatorCode+")":""); break;
        case "supervisor": key=r.supervisor; break;
        case "room":       key="Room "+r.room; break;
        case "crop":       key=r.crop; break;
        case "variety":    key=(r.crop||"")+" - "+(r.variety||"—"); break;
        case "lot":        key="LOT-"+r.lot; break;
        case "cycle":      key=r.cycle; break;
        default:           key=r.date;
      }
      if (!groups[key]) groups[key]={key:key,entries:0,usedBottle:0,usedCulture:0,
        totBottle:0,totCulture:0,mrSum:0,mrCnt:0,contaPctSum:0,contaFilled:0,contaBottleSum:0,
        rooms:{},details:[]};
      var g=groups[key];
      g.entries++; g.usedBottle+=r.usedBottle; g.usedCulture+=r.usedCulture;
      g.totBottle+=r.totBottle; g.totCulture+=r.totCulture;
      g.contaBottleSum+=r.contaBottle||0;
      g.rooms[r.room]=true;
      g.details.push(r);
      if (r.usedCulture>0){g.mrSum+=r.mr;g.mrCnt++;}
      if (r.contaFilled){g.contaPctSum+=r.contaPct;g.contaFilled++;}
    });
    var summary=Object.keys(groups).map(function(k){
      var g=groups[k];
      return {key:g.key,entries:g.entries,
        usedBottle:g.usedBottle,usedCulture:g.usedCulture,
        totBottle:g.totBottle,totCulture:g.totCulture,
        contaBottle:g.contaBottleSum,
        rooms:Object.keys(g.rooms).sort().join(", "),
        details:g.details,
        avgMr:g.mrCnt>0?(g.mrSum/g.mrCnt).toFixed(2):"0.00",
        avgConta:g.contaFilled>0?(g.contaPctSum/g.contaFilled).toFixed(2):"—"};
    });
    return {success:true,type:"summary",groupBy:params.groupBy,rows:summary,count:allRows.length};
  } catch(e) { return {success:false,message:"Error: "+e.message}; }
}

function getDashboardAll(from, to) {
  var result = { pir:null, pgr:null, wash:null };
  try { result.pir = getMainDashboard(from, to); } catch(e) { result.pir = null; }
  try {
    result.pgr = getPGRReport({dateFrom:from,dateTo:to,shift:'ALL',groupBy:'detail',
      operator:'',crop:'',lots:[],varieties:[],cycles:[]});
  } catch(e) { result.pgr = null; }
  try {
    result.wash = getWashingReport({dateFrom:from,dateTo:to,shift:'ALL',
      workType:'ALL',operator:'',groupBy:'detail'});
  } catch(e) { result.wash = null; }
  return result;
}

// ── PIR OPERATORS ─────────────────────────────────────────
function getOperatorList() {
  try {
    var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.OPERATORS_SHEET);
    if (!sheet) return {success:false,operators:[]};
    var lastRow=sheet.getLastRow();
    if (lastRow<=1) return {success:true,operators:[]};
    var data=sheet.getRange(1,1,lastRow,3).getValues(), ops=[];
    for (var i=1;i<data.length;i++) {
      var name=data[i][0]?data[i][0].toString().trim():'';
      var code=data[i][1]?data[i][1].toString().trim():'';
      var active=data[i][2]?data[i][2].toString().trim().toLowerCase():'yes';
      if (!name||active==='no') continue;
      ops.push({rowIndex:i+1,name:name,code:code});
    }
    return {success:true,operators:ops};
  } catch(e) { return {success:false,operators:[],message:e.message}; }
}

function getAllOperators() {
  try {
    var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.OPERATORS_SHEET);
    if (!sheet) return {success:false,operators:[]};
    var data=sheet.getDataRange().getValues(), ops=[];
    for (var i=1;i<data.length;i++) {
      if (!data[i][0]) continue;
      ops.push({rowIndex:i+1,name:data[i][0].toString().trim(),
        code:data[i][1].toString().trim(),active:data[i][2].toString().toLowerCase()});
    }
    return {success:true,operators:ops};
  } catch(e) { return {success:false,operators:[],message:e.message}; }
}

function saveOperator(opData) {
  try {
    var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.OPERATORS_SHEET);
    if (!sheet) return {success:false,message:"OPERATORS sheet nahi mili"};
    var row=[opData.name.trim(),opData.code.trim().toUpperCase(),opData.active||'yes'];
    if (opData.rowIndex) {
      sheet.getRange(opData.rowIndex,1,1,3).setValues([row]);
      return {success:true,message:"Operator update ho gaya!"};
    }
    var data=sheet.getDataRange().getValues();
    for (var i=1;i<data.length;i++) {
      if (data[i][1].toString().trim().toUpperCase()===opData.code.trim().toUpperCase())
        return {success:false,message:'Code "'+opData.code+'" pehle se exist karta hai!'};
    }
    sheet.appendRow(row);
    return {success:true,message:'"'+opData.name+'" add ho gaya!'};
  } catch(e) { return {success:false,message:e.message}; }
}

function deleteOperator(rowIndex) {
  try {
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.OPERATORS_SHEET)
      .getRange(rowIndex,3).setValue("no");
    return {success:true,message:"Operator deactivate ho gaya!"};
  } catch(e) { return {success:false,message:e.message}; }
}

// ── PIR CYCLES ────────────────────────────────────────────
function getCyclesForCrop(cropName) {
  try {
    var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.CYCLES_SHEET);
    if (!sheet) return {success:false,cycles:[]};
    var data=sheet.getDataRange().getValues(), cycles=[], found=false;
    var cropKey=cropName?cropName.toUpperCase():"DEFAULT";
    for (var i=1;i<data.length;i++) {
      if (!data[i][0]) continue;
      if (data[i][0].toString().toUpperCase()===cropKey&&data[i][2].toString().toLowerCase()==="yes") {
        cycles.push({rowIndex:i+1,cycle:data[i][1].toString()}); found=true;
      }
    }
    if (!found) {
      for (var j=1;j<data.length;j++) {
        if (!data[j][0]) continue;
        if (data[j][0].toString().toUpperCase()==="DEFAULT"&&data[j][2].toString().toLowerCase()==="yes")
          cycles.push({rowIndex:j+1,cycle:data[j][1].toString()});
      }
    }
    return {success:true,cycles:cycles};
  } catch(e) { return {success:false,cycles:[]}; }
}

function addCycleToSheet(cropName, cycleName) {
  try {
    var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.CYCLES_SHEET);
    if (!sheet) return {success:false,message:"CYCLES sheet nahi mili"};
    var cropKey=cropName?cropName.toUpperCase():"DEFAULT";
    var cu=cycleName.trim().toUpperCase();
    var data=sheet.getDataRange().getValues();
    for (var i=1;i<data.length;i++) {
      if (data[i][0].toString().toUpperCase()===cropKey&&data[i][1].toString().toUpperCase()===cu)
        return {success:false,message:'"'+cu+'" pehle se exist karta hai!'};
    }
    sheet.appendRow([cropKey,cu,"yes"]);
    return {success:true,message:'"'+cu+'" cycle add ho gaya!'};
  } catch(e) { return {success:false,message:e.message}; }
}

function getAllCycles() {
  try {
    var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.CYCLES_SHEET);
    if (!sheet) return {success:false,cycles:[]};
    var data=sheet.getDataRange().getValues(), cycles=[];
    for (var i=1;i<data.length;i++) {
      if (!data[i][0]&&!data[i][1]) continue;
      cycles.push({rowIndex:i+1,crop:data[i][0].toString(),
        cycle:data[i][1].toString(),active:data[i][2].toString()});
    }
    return {success:true,cycles:cycles};
  } catch(e) { return {success:false,cycles:[]}; }
}

function getAllCyclesForFilter() {
  try {
    var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.CYCLES_SHEET);
    if (!sheet) return {success:false,cycles:[]};
    var data=sheet.getDataRange().getValues();
    var cycles=[], seen={};
    for (var i=1;i<data.length;i++) {
      if (!data[i][1]) continue;
      var c=data[i][1].toString().trim();
      if (!seen[c]&&data[i][2].toString().toLowerCase()==='yes') {
        cycles.push(c); seen[c]=true;
      }
    }
    return {success:true,cycles:cycles};
  } catch(e) { return {success:false,cycles:[]}; }
}

function deleteCycleFromSheet(rowIndex) {
  try {
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.CYCLES_SHEET)
      .getRange(rowIndex,3).setValue("no");
    return {success:true,message:"Cycle deactivate ho gaya!"};
  } catch(e) { return {success:false,message:e.message}; }
}

// ── PIR USERS ─────────────────────────────────────────────
function getAllUsers() {
  try {
    var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.USERS_SHEET);
    if (!sheet) return {success:false,message:"USERS sheet nahi mili"};
    var data=sheet.getDataRange().getValues(), users=[];
    for (var i=1;i<data.length;i++) {
      if (!data[i][0]) continue;
      users.push({rowIndex:i+1,username:data[i][0],password:data[i][1],
        name:data[i][2],role:data[i][3],dept:data[i][4],active:data[i][5]});
    }
    return {success:true,users:users};
  } catch(e) { return {success:false,message:e.message}; }
}

function saveUser(userData) {
  try {
    var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.USERS_SHEET);
    var row=[userData.username.trim(),userData.password,userData.name.trim(),
             userData.role,userData.dept||'ALL',userData.active||'yes'];
    if (userData.rowIndex) {
      sheet.getRange(userData.rowIndex,1,1,6).setValues([row]);
    } else {
      var data=sheet.getDataRange().getValues();
      for (var i=1;i<data.length;i++) {
        if (data[i][0].toString().trim()===userData.username.trim())
          return {success:false,message:"Username pehle se exist karta hai!"};
      }
      sheet.appendRow(row);
    }
    return {success:true,message:"User save ho gaya!"};
  } catch(e) { return {success:false,message:e.message}; }
}

function deleteUser(rowIndex) {
  try {
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.USERS_SHEET).deleteRow(rowIndex);
    return {success:true,message:"User delete ho gaya!"};
  } catch(e) { return {success:false,message:e.message}; }
}

// ── PIR CROPS ─────────────────────────────────────────────
function getAllCrops() {
  try {
    var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.CROPS_SHEET);
    if (!sheet) return {success:false,crops:[]};
    var data=sheet.getDataRange().getValues(), crops=[];
    for (var i=1;i<data.length;i++) {
      if (!data[i][0]) continue;
      crops.push({rowIndex:i+1,crop:data[i][0].toString(),
        varieties:data[i][1].toString(),cycleType:data[i][2].toString()||'DEFAULT'});
    }
    return {success:true,crops:crops};
  } catch(e) { return {success:false,crops:[]}; }
}

function getCropList() {
  try {
    var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.CROPS_SHEET);
    if (!sheet) return {success:false,crops:[]};
    var data=sheet.getDataRange().getValues(), crops=[];
    for (var i=1;i<data.length;i++) {
      if (!data[i][0]) continue;
      crops.push({crop:data[i][0].toString(),
        varieties:data[i][1].toString().split(',').map(function(v){return v.trim();}).filter(Boolean),
        cycleType:data[i][2].toString()||'DEFAULT'});
    }
    return {success:true,crops:crops};
  } catch(e) { return {success:false,crops:[]}; }
}

function saveCrop(cropData) {
  try {
    var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.CROPS_SHEET);
    var row=[cropData.crop.trim().toUpperCase(),cropData.varieties.trim(),cropData.cycleType||'DEFAULT'];
    if (cropData.rowIndex) sheet.getRange(cropData.rowIndex,1,1,3).setValues([row]);
    else sheet.appendRow(row);
    return {success:true,message:"Crop save ho gaya!"};
  } catch(e) { return {success:false,message:e.message}; }
}

function deleteCrop(rowIndex) {
  try {
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.CROPS_SHEET).deleteRow(rowIndex);
    return {success:true};
  } catch(e) { return {success:false,message:e.message}; }
}

// ── PIR FIX OLD ENTRIES ───────────────────────────────────
function fixOldEntries() {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var opSheet=ss.getSheetByName(CONFIG.OPERATORS_SHEET);
  var codeMap={};
  if (opSheet) {
    var opData=opSheet.getDataRange().getValues();
    for (var i=1;i<opData.length;i++) {
      var nm=opData[i][0]?opData[i][0].toString().trim().toLowerCase():'';
      var cd=opData[i][1]?opData[i][1].toString().trim():'';
      if (nm&&cd) codeMap[nm]=cd;
    }
  }
  var totalFixed=0, results=[];
  for (var r=1;r<=CONFIG.PIR_TOTAL_ROOMS;r++) {
    var sName='PIR-ROOM-'+(r<10?'0'+r:r);
    var sheet=ss.getSheetByName(sName);
    if (!sheet||sheet.getLastRow()<=1){results.push(sName+': empty');continue;}
    var lastRow=sheet.getLastRow();
    var data=sheet.getRange(2,9,lastRow-1,2).getValues();
    var fixed=0,skip=0;
    for (var i=0;i<data.length;i++) {
      var col10=data[i][1];
      var isNum=(typeof col10==='number')||(!isNaN(parseFloat(col10))&&col10!==''&&col10!==null);
      if (!isNum){skip++;continue;}
      var rowNum=i+2;
      var fr=sheet.getRange(rowNum,1,1,31).getValues()[0];
      var opName=fr[8]?fr[8].toString().trim():'';
      var opCode=codeMap[opName.toLowerCase()]||'';
      if (!opCode) {
        var keys=Object.keys(codeMap);
        for (var k=0;k<keys.length;k++) {
          if (keys[k].indexOf(opName.toLowerCase())>=0||opName.toLowerCase().indexOf(keys[k])>=0)
            {opCode=codeMap[keys[k]];break;}
        }
      }
      var nr=[fr[0],fr[1],fr[2],fr[3],fr[4],fr[5],fr[6],fr[7],fr[8],opCode,
              fr[9],fr[10],fr[11],fr[12],fr[13],fr[14],fr[15],fr[16],fr[17],
              fr[18],fr[19],fr[20],fr[21],fr[22],fr[23],fr[24],
              fr[25],fr[26],fr[27],fr[28],fr[29]];
      sheet.getRange(rowNum,1,1,31).setValues([nr]);
      sheet.getRange(rowNum,2).setNumberFormat("dd-mmm-yy");
      sheet.getRange(rowNum,25).setNumberFormat("0.00");
      if (rowNum%2===0) sheet.getRange(rowNum,1,1,31).setBackground("#E1F5EE");
      fixed++;totalFixed++;
      if (fixed%50===0) SpreadsheetApp.flush();
    }
    results.push(sName+': Fixed='+fixed+' | OK='+skip);
  }
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('Fix Complete!\nTotal: '+totalFixed+'\n\n'+results.join('\n'));
}

// ════════════════════════════════════════════════════════════
// PGR MODULE
// ════════════════════════════════════════════════════════════

function initializePGR() {
  var ss = SpreadsheetApp.openById(PGR_SS_ID);
  var prod = ss.getSheetByName("PGR_PRODUCTION");
  if (!prod) {
    prod = ss.insertSheet("PGR_PRODUCTION");
    var ph = ["Sr.No","Date","Shift","LOT","Crop","Variety","Cycle","PGR Room",
      "Operator Name","Operator Code","Tray","Bottle","Own Contam Bottle",
      "Cross Name","Cross Code","Cross Bottle","Cross Contam Bottle",
      "Total Contam Bottle","Total Contam %","Supervisor"];
    prod.appendRow(ph);
    prod.getRange(1,1,1,ph.length).setBackground("#1D9E75").setFontColor("#fff").setFontWeight("bold").setFontSize(10);
    prod.setFrozenRows(1);
    prod.setColumnWidths(1,ph.length,120);
    SpreadsheetApp.flush();
  }
  var exch = ss.getSheetByName("PGR_EXCHANGE");
  if (!exch) {
    exch = ss.insertSheet("PGR_EXCHANGE");
    var eh = ["Sr.No","Date","Shift","LOT","Crop","Variety","Cycle","Direction","PGR\/PIR Room","PIR\/PGR Room (Source)",
      "Sent Qty","Sent By","Received Qty","Status","Received By",
      "Received Date","Difference","Remarks"];
    exch.appendRow(eh);
    exch.getRange(1,1,1,eh.length).setBackground("#185FA5").setFontColor("#fff").setFontWeight("bold").setFontSize(10);
    exch.setFrozenRows(1);
    exch.setColumnWidths(1,eh.length,120);
    SpreadsheetApp.flush();
  } else {
    var hdr = exch.getRange(1,1,1,exch.getLastColumn()).getValues()[0];
    if (hdr.length < 18 || String(hdr[9]||'').indexOf('Source') < 0) {
      exch.insertColumnAfter(9);
      exch.getRange(1,10).setValue("PIR\/PGR Room (Source)").setBackground("#185FA5").setFontColor("#fff").setFontWeight("bold");
      exch.setColumnWidth(10, 120);
      SpreadsheetApp.flush();
      SpreadsheetApp.getUi().alert("PGR_EXCHANGE sheet updated: PIR/PGR Room (Source) column added!");
    }
  }
  var ops = ss.getSheetByName("PGR_OPERATORS");
  if (!ops) {
    ops = ss.insertSheet("PGR_OPERATORS");
    ops.appendRow(["Operator Name","Operator Code","Active"]);
    ops.getRange(1,1,1,3).setBackground("#1D9E75").setFontColor("#fff").setFontWeight("bold");
    ops.setFrozenRows(1);
    SpreadsheetApp.flush();
  }
  var crops = ss.getSheetByName("PGR_CROPS");
  if (!crops) {
    crops = ss.insertSheet("PGR_CROPS");
    crops.appendRow(["Crop","Varieties","CycleType"]);
    crops.getRange(1,1,1,3).setBackground("#1D9E75").setFontColor("#fff").setFontWeight("bold");
    crops.appendRow(["BANANA","G-09","BANANA"]);
    crops.appendRow(["BAMBOO","Balkua,Tulda,Golden","DEFAULT"]);
    crops.appendRow(["TEAK","India","DEFAULT"]);
    crops.setFrozenRows(1);
    SpreadsheetApp.flush();
  }
  try { var def=ss.getSheetByName("Sheet1"); if(def) ss.deleteSheet(def); } catch(e){}
  SpreadsheetApp.getUi().alert('✅ PGR Setup Complete!\n✓ PGR_PRODUCTION\n✓ PGR_EXCHANGE\n✓ PGR_OPERATORS\n✓ PGR_CROPS');
}

function getPGROperators() {
  try {
    var sheet=SpreadsheetApp.openById(PGR_SS_ID).getSheetByName("PGR_OPERATORS");
    if (!sheet) return {success:false,operators:[]};
    var data=sheet.getDataRange().getValues(), ops=[];
    for (var i=1;i<data.length;i++) {
      if (!data[i][0]) continue;
      if (data[i][2]&&data[i][2].toString().toLowerCase()==='no') continue;
      ops.push({rowIndex:i+1,name:data[i][0].toString().trim(),code:data[i][1].toString().trim()});
    }
    return {success:true,operators:ops};
  } catch(e) { return {success:false,operators:[],message:e.message}; }
}

function getAllPGROperators() {
  try {
    var sheet=SpreadsheetApp.openById(PGR_SS_ID).getSheetByName("PGR_OPERATORS");
    if (!sheet) return {success:false,operators:[]};
    var data=sheet.getDataRange().getValues(), ops=[];
    for (var i=1;i<data.length;i++) {
      if (!data[i][0]) continue;
      ops.push({rowIndex:i+1,name:data[i][0].toString().trim(),
        code:data[i][1].toString().trim(),active:data[i][2].toString().toLowerCase()});
    }
    return {success:true,operators:ops};
  } catch(e) { return {success:false,operators:[],message:e.message}; }
}

function savePGROperator(opData) {
  try {
    var sheet=SpreadsheetApp.openById(PGR_SS_ID).getSheetByName("PGR_OPERATORS");
    if (!sheet) return {success:false,message:"PGR_OPERATORS sheet nahi mili"};
    var row=[opData.name.trim(),opData.code.trim().toUpperCase(),opData.active||'yes'];
    if (opData.rowIndex) {
      sheet.getRange(opData.rowIndex,1,1,3).setValues([row]);
      return {success:true,message:"Operator updated!"};
    }
    var data=sheet.getDataRange().getValues();
    for (var i=1;i<data.length;i++) {
      if (data[i][1].toString().trim().toUpperCase()===opData.code.trim().toUpperCase())
        return {success:false,message:'Code "'+opData.code+'" already exists!'};
    }
    sheet.appendRow(row);
    return {success:true,message:'"'+opData.name+'" added!'};
  } catch(e) { return {success:false,message:e.message}; }
}

function deletePGROperator(rowIndex) {
  try {
    SpreadsheetApp.openById(PGR_SS_ID).getSheetByName("PGR_OPERATORS")
      .getRange(rowIndex,3).setValue("no");
    return {success:true,message:"Operator deactivated!"};
  } catch(e) { return {success:false,message:e.message}; }
}

function getPGRCropList() {
  try {
    var sheet=SpreadsheetApp.openById(PGR_SS_ID).getSheetByName("PGR_CROPS");
    if (!sheet) return {success:false,crops:[]};
    var data=sheet.getDataRange().getValues(), crops=[];
    for (var i=1;i<data.length;i++) {
      if (!data[i][0]) continue;
      crops.push({crop:data[i][0].toString(),
        varieties:data[i][1].toString().split(',').map(function(v){return v.trim();}).filter(Boolean),
        cycleType:data[i][2].toString()||'DEFAULT'});
    }
    return {success:true,crops:crops};
  } catch(e) { return {success:false,crops:[]}; }
}

// ── PGR SAVE ENTRIES (LOCK) ──────────────────────────────
function savePGREntries(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var sheet=SpreadsheetApp.openById(PGR_SS_ID).getSheetByName("PGR_PRODUCTION");
    if (!sheet) return {success:false,message:"PGR_PRODUCTION sheet not found"};
    var date=payload.date ? new Date(payload.date) : new Date();
    var saved=0;
    payload.rows.forEach(function(row) {
      if (!row.opName||!row.opName.trim()) return;
      var rowLot = (row.lot||payload.lot||'').trim();
      var rowCrop = (row.crop||payload.crop||'').trim();
      var rowVariety = (row.variety||payload.variety||'').trim();
      var rowCycle = (row.cycle||payload.cycle||'').trim();
      var sr=sheet.getLastRow();
      var ownContamBtl = n(row.ownContamBtl||row.contamBtl||0);
      var crossContamBtl = n(row.crossContamBtl||0);
      var totalContamBtl = ownContamBtl + crossContamBtl;
      var baseBottle = n(row.bottle)>0 ? n(row.bottle) : n(row.crossBottle||0);
      var totalContamPct = baseBottle>0 ? Math.round((totalContamBtl/baseBottle)*10000)/100 : 0;
      var rowRoom = (row.pgrRoom||row.room||'').toString().trim();
      sheet.appendRow([
        sr, date, payload.shift,
        rowLot, rowCrop, rowVariety, rowCycle, rowRoom,
        row.opName.trim(), (row.opCode||'').toString().trim(),
        n(row.tray), n(row.bottle), ownContamBtl,
        (row.crossName||'').trim(), (row.crossCode||'').trim(),
        n(row.crossBottle||0), crossContamBtl,
        totalContamBtl, totalContamPct, payload.supervisorName
      ]);
      var nr=sheet.getLastRow();
      sheet.getRange(nr,2).setNumberFormat("dd-mmm-yy");
      sheet.getRange(nr,19).setNumberFormat("0.00");
      if (nr%2===0) sheet.getRange(nr,1,1,20).setBackground("#E1F5EE");
      saved++;
    });
    SpreadsheetApp.flush();
    return {success:true,saved:saved,message:saved+" entries saved!"};
  } catch(e) { return {success:false,message:"Error: "+e.message}; }
  finally { try{lock.releaseLock();}catch(e2){} }
}

function getPGREntries(dateFrom, dateTo, shiftFilter) {
  try {
    var sheet=SpreadsheetApp.openById(PGR_SS_ID).getSheetByName("PGR_PRODUCTION");
    if (!sheet||sheet.getLastRow()<=1) return {success:true,rows:[]};
    var tz=Session.getScriptTimeZone();
    var fromD=new Date(dateFrom); fromD.setHours(0,0,0,0);
    var toD=new Date(dateTo); toD.setHours(23,59,59,999);
    var data=sheet.getDataRange().getValues(), out=[];
    for (var i=1;i<data.length;i++) {
      var cell=data[i][1]; if (!cell) continue;
      var cellDate=cell;
      var rd;
      if (cellDate instanceof Date) {
        rd=new Date(cellDate.getTime());
      } else {
        var ds2=String(cellDate);
        if (ds2.indexOf('/')>0){var p=ds2.split('/');if(p.length===3)rd=new Date(p[2],p[1]-1,p[0]);else rd=new Date(ds2);}
        else rd=new Date(ds2);
      }
      if (!rd||isNaN(rd.getTime())) continue;
      var rdStrP = Utilities.formatDate(rd, tz, "yyyy-MM-dd");
      var fromStrP = Utilities.formatDate(fromD, tz, "yyyy-MM-dd");
      var toStrP = Utilities.formatDate(toD, tz, "yyyy-MM-dd");
      if (rdStrP<fromStrP||rdStrP>toStrP) continue;
      if (shiftFilter&&shiftFilter!=='ALL'&&data[i][2]!==shiftFilter) continue;
      var isNew=data[i].length>=18;
      var row;
      if (isNew) {
        row = {
          srNo:data[i][0], date:Utilities.formatDate(rd,tz,"dd-MMM-yy"),
          shift:data[i][2], lot:String(data[i][3]||''),
          crop:String(data[i][4]||''), variety:String(data[i][5]||''),
          cycle:String(data[i][6]||''), room:String(data[i][7]||''),
          opName:String(data[i][8]||''), opCode:String(data[i][9]||''),
          tray:n(data[i][10]), bottle:n(data[i][11]), contamBtl:n(data[i][12]),
          crossName:String(data[i][13]||''), crossCode:String(data[i][14]||''),
          crossBottle:n(data[i][15]), crossContamBtl:n(data[i][16]),
          totalContamBtl:n(data[i][17]), totalContamPct:n(data[i][18]).toFixed(2)
        };
      } else {
        row = {
          srNo:data[i][0], date:Utilities.formatDate(rd,tz,"dd-MMM-yy"),
          shift:data[i][2], lot:String(data[i][3]||''),
          crop:'—', variety:'—', cycle:String(data[i][4]||''),
          opName:String(data[i][5]||''), opCode:String(data[i][6]||''),
          tray:n(data[i][7]), bottle:n(data[i][8]), contamBtl:n(data[i][9]),
          crossName:String(data[i][10]||''), crossCode:String(data[i][11]||''),
          crossBottle:n(data[i][12]), crossContamBtl:n(data[i][13]),
          totalContamBtl:n(data[i][14]), totalContamPct:n(data[i][15]).toFixed(2)
        };
      }
      out.push(row);
    }
    return {success:true,rows:out};
  } catch(e) { return {success:false,rows:[],message:e.message}; }
}

function getPGRReport(params) {
  try {
    var sheet=SpreadsheetApp.openById(PGR_SS_ID).getSheetByName("PGR_PRODUCTION");
    if (!sheet||sheet.getLastRow()<=1) return {success:true,type:'detail',rows:[],count:0};
    var tz=Session.getScriptTimeZone();
    var fromD=params.dateFrom?new Date(params.dateFrom):null;
    var toD=params.dateTo?new Date(params.dateTo):null;
    if (fromD) fromD.setHours(0,0,0,0);
    if (toD)   toD.setHours(23,59,59,999);
    var data=sheet.getDataRange().getValues(), allRows=[];
    for (var i=1;i<data.length;i++) {
      var cell=data[i][1]; if (!cell) continue;
      var rd;
      if(cell instanceof Date){rd=new Date(cell.getTime());}else{var ds3=String(cell);if(ds3.indexOf('/')>0){var p3=ds3.split('/');if(p3.length===3)rd=new Date(p3[2],p3[1]-1,p3[0]);else rd=new Date(ds3);}else rd=new Date(ds3);}
      if(!rd||isNaN(rd.getTime()))continue;
      rd.setHours(0,0,0,0);
      if (fromD&&rd<fromD) continue;
      if (toD&&rd>toD) continue;
      if (params.shift&&params.shift!=='ALL'&&data[i][2]!==params.shift) continue;
      if (params.lots&&params.lots.length>0) {
        if (params.lots.indexOf(String(data[i][3]))<0) continue;
      }
      if (params.crop&&params.crop.trim()&&String(data[i][4]).toLowerCase()!==params.crop.toLowerCase()) continue;
      if (params.varieties&&params.varieties.length>0) {
        if (params.varieties.indexOf(String(data[i][5]))<0) continue;
      }
      if (params.cycles&&params.cycles.length>0) {
        if (params.cycles.indexOf(String(data[i][6]))<0) continue;
      }
      if (params.operator&&params.operator.trim()) {
        var opMatch=String(data[i][8]).toLowerCase().indexOf(params.operator.toLowerCase())>=0;
        if (!opMatch) continue;
      }
      var rowLen2=data[i].length;
      var isNew2=rowLen2>=18;
      var rw;
      if(isNew2){
        rw={srNo:data[i][0], date:Utilities.formatDate(rd,tz,"dd-MMM-yy"), dateRaw:rd.getTime(),
          shift:data[i][2], lot:String(data[i][3]||''),
          crop:String(data[i][4]||''), variety:String(data[i][5]||''),
          cycle:String(data[i][6]||''), room:String(data[i][7]||''),
          opName:String(data[i][8]||''), opCode:String(data[i][9]||''),
          tray:n(data[i][10]), bottle:n(data[i][11]), contamBtl:n(data[i][12]),
          crossName:String(data[i][13]||''), crossCode:String(data[i][14]||''),
          crossBottle:n(data[i][15]), crossContamBtl:n(data[i][16]),
          totalContamBtl:n(data[i][17]), totalContamPct:n(data[i][18])};
      } else {
        rw={srNo:data[i][0], date:Utilities.formatDate(rd,tz,"dd-MMM-yy"), dateRaw:rd.getTime(),
          shift:data[i][2], lot:String(data[i][3]||''),
          crop:'—', variety:'—', cycle:String(data[i][4]||''), room:'—',
          opName:String(data[i][5]||''), opCode:String(data[i][6]||''),
          tray:n(data[i][7]), bottle:n(data[i][8]), contamBtl:n(data[i][9]),
          crossName:String(data[i][10]||''), crossCode:String(data[i][11]||''),
          crossBottle:n(data[i][12]), crossContamBtl:n(data[i][13]),
          totalContamBtl:n(data[i][14]), totalContamPct:n(data[i][15])};
      }
      allRows.push(rw);
    }
    allRows.sort(function(a,b){return a.dateRaw-b.dateRaw;});
    if (params.groupBy==='detail')
      return {success:true,type:'detail',rows:allRows,count:allRows.length};
    var groups={};
    allRows.forEach(function(r) {
      var key;
      switch(params.groupBy) {
        case 'daily':    key=r.date; break;
        case 'monthly':  key=r.date.substr(3); break;
        case 'operator': key=r.opName+(r.opCode?' ('+r.opCode+')':''); break;
        case 'lot':      key='LOT: '+r.lot; break;
        case 'cycle':    key='Cycle: '+r.cycle; break;
        case 'crop':     key='Crop: '+(r.crop||'Unknown'); break;
        case 'variety':  key='Var: '+(r.variety||'Unknown'); break;
        default:         key=r.date;
      }
      if (!groups[key]) groups[key]={key:key,entries:0,tray:0,bottle:0,contamBtl:0,
        crossBottle:0,crossContamBtl:0,totalContamBtl:0,pctSum:0,pctCnt:0,rooms:{},details:[]};
      var g=groups[key];
      g.entries++; g.tray+=r.tray; g.bottle+=r.bottle; g.contamBtl+=r.contamBtl;
      g.crossBottle+=r.crossBottle; g.crossContamBtl+=r.crossContamBtl;
      g.totalContamBtl+=r.totalContamBtl;
      if(r.room&&r.room!=='—')g.rooms[r.room]=true;
      g.details.push(r);
      if (r.totalContamPct>0){g.pctSum+=r.totalContamPct;g.pctCnt++;}
    });
    var summary=Object.keys(groups).map(function(k){
      var g=groups[k];
      return {key:g.key,entries:g.entries,tray:g.tray,bottle:g.bottle,
        contamBtl:g.contamBtl,crossBottle:g.crossBottle,
        crossContamBtl:g.crossContamBtl,totalContamBtl:g.totalContamBtl,
        rooms:Object.keys(g.rooms).sort().join(", "),details:g.details,
        avgContamPct:g.pctCnt>0?(g.pctSum/g.pctCnt).toFixed(2):'0.00'};
    });
    return {success:true,type:'summary',groupBy:params.groupBy,rows:summary,count:allRows.length};
  } catch(e) { return {success:false,message:e.message}; }
}

// ── PGR EXCHANGE SAVE (LOCK) ─────────────────────────────
function savePGRExchange(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var sheet=SpreadsheetApp.openById(PGR_SS_ID).getSheetByName("PGR_EXCHANGE");
    if (!sheet) return {success:false,message:"PGR_EXCHANGE sheet nahi mili"};
    var tz=Session.getScriptTimeZone();
    var sr=sheet.getLastRow();
    var exDate = payload.date ? new Date(payload.date) : new Date();
    sheet.appendRow([
      sr, exDate, payload.shift,
      (payload.lot||''), (payload.crop||''), (payload.variety||''), (payload.cycle||''),
      (payload.direction||'PGR to PIR'), (payload.room||''), (payload.pirRoom||''),
      n(payload.sentQty), (payload.sentBy||''),
      '', 'Pending', '', '', '', (payload.remarks||'')
    ]);
    var nr=sheet.getLastRow();
    sheet.getRange(nr,2).setNumberFormat("dd-mmm-yy");
    if (nr%2===0) sheet.getRange(nr,1,1,18).setBackground("#E6F1FB");
    SpreadsheetApp.flush();
    return {success:true,message:"Exchange entry saved!",rowIndex:nr};
  } catch(e) { return {success:false,message:e.message}; }
  finally { try{lock.releaseLock();}catch(e2){} }
}

function getPGRExchangeForRoom(room) {
  try {
    var sheet=SpreadsheetApp.openById(PGR_SS_ID).getSheetByName("PGR_EXCHANGE");
    if (!sheet||sheet.getLastRow()<=1) return {success:true,rows:[]};
    var tz=Session.getScriptTimeZone();
    var data=sheet.getDataRange().getValues(), out=[];
    for (var i=1;i<data.length;i++) {
      if (!data[i][1]) continue;
      if (data[i][8].toString().trim()!==room.toString().trim()) continue;
      var rd=data[i][1] instanceof Date?new Date(data[i][1].getTime()):new Date(data[i][1]);
      out.push({
        rowIndex:i+1, date:Utilities.formatDate(rd,tz,"dd-MMM-yy"),
        shift:data[i][2], lot:data[i][3], cycle:data[i][4],
        direction:data[i][5], room:data[i][6],
        sentQty:n(data[i][7]), sentBy:data[i][8],
        receivedQty:n(data[i][9]), status:data[i][10],
        receivedBy:data[i][11], receivedDate:data[i][12],
        difference:n(data[i][13]), remarks:data[i][14]
      });
    }
    return {success:true,rows:out};
  } catch(e) { return {success:false,rows:[],message:e.message}; }
}

function getAllPGRExchange(dateFrom, dateTo) {
  try {
    var sheet=SpreadsheetApp.openById(PGR_SS_ID).getSheetByName("PGR_EXCHANGE");
    if (!sheet||sheet.getLastRow()<=1) return {success:true,rows:[]};
    var tz=Session.getScriptTimeZone();
    var fromD=new Date(dateFrom); fromD.setHours(0,0,0,0);
    var toD=new Date(dateTo); toD.setHours(23,59,59,999);
    var data=sheet.getDataRange().getValues(), out=[];
    for (var i=1;i<data.length;i++) {
      if (!data[i][1]) continue;
      var rawDate=data[i][1];
      var rd;
      if (rawDate instanceof Date) {
        rd=new Date(rawDate.getTime());
      } else {
        var ds=String(rawDate);
        if (ds.indexOf('/')>0) {
          var parts=ds.split('/');
          if (parts.length===3) rd=new Date(parts[2],parts[1]-1,parts[0]);
          else rd=new Date(ds);
        } else {
          rd=new Date(ds);
        }
      }
      if (!rd||isNaN(rd.getTime())) continue;
      var rdStr2 = Utilities.formatDate(rd, tz, "yyyy-MM-dd");
      var fromStr2 = Utilities.formatDate(fromD, tz, "yyyy-MM-dd");
      var toStr2 = Utilities.formatDate(toD, tz, "yyyy-MM-dd");
      if (rdStr2<fromStr2||rdStr2>toStr2) continue;
      var isNew=data[i].length>=17 || (data[i].length>=9 && !String(data[i][4]||'').match(/^[A-Z]-\d/));
      var row;
      if (isNew) {
        row = {
          rowIndex:i+1, date:Utilities.formatDate(rd,tz,"dd-MMM-yy"),
          shift:String(data[i][2]||''), lot:String(data[i][3]||''),
          crop:String(data[i][4]||''), variety:String(data[i][5]||''),
          cycle:String(data[i][6]||''),
          direction:String(data[i][7]||''), room:String(data[i][8]||''),
          pirRoom:String(data[i][9]||''),
          sentQty:n(data[i][10]), sentBy:String(data[i][11]||''),
          receivedQty:n(data[i][12])||'', status:String(data[i][13]||'Pending'),
          receivedBy:String(data[i][14]||''), receivedDate:data[i][15],
          difference:n(data[i][16]), remarks:String(data[i][17]||'')
        };
      } else {
        row = {
          rowIndex:i+1, date:Utilities.formatDate(rd,tz,"dd-MMM-yy"),
          shift:String(data[i][2]||''), lot:String(data[i][3]||''),
          crop:'—', variety:'—', cycle:String(data[i][4]||''),
          direction:String(data[i][5]||''), room:String(data[i][6]||''),
          sentQty:n(data[i][7]), sentBy:String(data[i][8]||''),
          receivedQty:n(data[i][9])||'', status:String(data[i][10]||'Pending'),
          receivedBy:String(data[i][11]||''), receivedDate:data[i][12],
          difference:n(data[i][13]), remarks:String(data[i][14]||'')
        };
      }
      out.push(row);
    }
    return {success:true,rows:out};
  } catch(e) { return {success:false,rows:[],message:e.message}; }
}

// ── PGR EXCHANGE STATUS UPDATE (LOCK) ────────────────────
function updatePGRExchangeStatus(rowIndex, status, receivedQty, receivedBy, remarks) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var sheet=SpreadsheetApp.openById(PGR_SS_ID).getSheetByName("PGR_EXCHANGE");
    if (!sheet) return {success:false,message:"Exchange sheet not found"};
    var tz=Session.getScriptTimeZone();
    var today=Utilities.formatDate(new Date(),tz,"dd/MM/yyyy");
    var rowData=sheet.getRange(rowIndex,1,1,20).getValues()[0];
    var dirField=String(rowData[7]||'');
    var isNewFmt=dirField.indexOf('PIR')>=0||dirField.indexOf('PGR')>=0||dirField.indexOf('to')>=0;
    var sentQtyCol, rcvQtyCol, statusCol, rcvByCol, rcvDateCol, diffCol, remarksCol, numCols;
    if (isNewFmt) {
      sentQtyCol=11; rcvQtyCol=13; statusCol=14; rcvByCol=15; rcvDateCol=16; diffCol=17; remarksCol=18; numCols=18;
    } else {
      sentQtyCol=8; rcvQtyCol=10; statusCol=11; rcvByCol=12; rcvDateCol=13; diffCol=14; remarksCol=15; numCols=15;
    }
    var sentQty=n(sheet.getRange(rowIndex,sentQtyCol).getValue());
    var rcvQty=n(receivedQty);
    var diff=rcvQty-sentQty;
    sheet.getRange(rowIndex,rcvQtyCol).setValue(rcvQty);
    sheet.getRange(rowIndex,statusCol).setValue(status);
    sheet.getRange(rowIndex,rcvByCol).setValue(receivedBy);
    sheet.getRange(rowIndex,rcvDateCol).setValue(today);
    sheet.getRange(rowIndex,diffCol).setValue(diff);
    sheet.getRange(rowIndex,remarksCol).setValue(remarks||'');
    if (status==='Rejected') {
      sheet.getRange(rowIndex,1,1,numCols).setBackground("#FAECE7").setFontColor("#993C1D");
    } else if (diff!==0) {
      sheet.getRange(rowIndex,1,1,numCols).setBackground("#FAEEDA");
    } else {
      sheet.getRange(rowIndex,1,1,numCols).setBackground("#E1F5EE");
    }
    SpreadsheetApp.flush();
    return {success:true,message:"Status updated to "+status};
  } catch(e) { return {success:false,message:e.message}; }
  finally { try{lock.releaseLock();}catch(e2){} }
}

// ── HELPER ────────────────────────────────────────────────
function n(v){return parseFloat(v)||0;}

// ════════════════════════════════════════════════════════════
// WASHING MODULE
// ════════════════════════════════════════════════════════════

var WASHING_SS_ID = "1JifnH-ov7AphKHv3wfPjCd_GTQWe1ARbVg6humaebmA";

function initializeWashing() {
  var ss = SpreadsheetApp.openById(WASHING_SS_ID);
  var ws = ss.getSheetByName("WASHING_DATA");
  if (!ws) {
    ws = ss.insertSheet("WASHING_DATA");
    var wh = ["Sr.No","Date","Shift","Supervisor","Operator Name","Category","Sub-Category","Batch No","Quantity"];
    ws.appendRow(wh);
    ws.getRange(1,1,1,wh.length).setBackground("#185FA5").setFontColor("#fff").setFontWeight("bold").setFontSize(10);
    ws.setFrozenRows(1);
    ws.setColumnWidths(1,wh.length,130);
    SpreadsheetApp.flush();
  }
  var wo = ss.getSheetByName("WASHING_OPERATORS");
  if (!wo) {
    wo = ss.insertSheet("WASHING_OPERATORS");
    wo.appendRow(["Operator Name","Operator Code","Active"]);
    wo.getRange(1,1,1,3).setBackground("#1D9E75").setFontColor("#fff").setFontWeight("bold");
    wo.setFrozenRows(1);
    SpreadsheetApp.flush();
  }
  try { var def=ss.getSheetByName("Sheet1"); if(def) ss.deleteSheet(def); } catch(e){}
  SpreadsheetApp.getUi().alert("✅ Washing Setup Complete!\n✓ WASHING_DATA\n✓ WASHING_OPERATORS");
}

function getWashingOperators() {
  try {
    var sheet = SpreadsheetApp.openById(WASHING_SS_ID).getSheetByName("WASHING_OPERATORS");
    if (!sheet) return {success:false, operators:[]};
    var data = sheet.getDataRange().getValues(), ops = [];
    for (var i=1; i<data.length; i++) {
      if (!data[i][0]) continue;
      if (data[i][2] && data[i][2].toString().toLowerCase() === 'no') continue;
      ops.push({rowIndex:i+1, name:data[i][0].toString().trim(), code:data[i][1].toString().trim()});
    }
    return {success:true, operators:ops};
  } catch(e) { return {success:false, operators:[], message:e.message}; }
}

function getAllWashingOperators() {
  try {
    var sheet = SpreadsheetApp.openById(WASHING_SS_ID).getSheetByName("WASHING_OPERATORS");
    if (!sheet) return {success:false, operators:[]};
    var data = sheet.getDataRange().getValues(), ops = [];
    for (var i=1; i<data.length; i++) {
      if (!data[i][0]) continue;
      ops.push({rowIndex:i+1, name:data[i][0].toString().trim(), code:data[i][1].toString().trim(), active:data[i][2].toString().toLowerCase()});
    }
    return {success:true, operators:ops};
  } catch(e) { return {success:false, operators:[], message:e.message}; }
}

function saveWashingOperator(opData) {
  try {
    var sheet = SpreadsheetApp.openById(WASHING_SS_ID).getSheetByName("WASHING_OPERATORS");
    if (!sheet) return {success:false, message:"WASHING_OPERATORS sheet nahi mili"};
    var row = [opData.name.trim(), opData.code.trim().toUpperCase(), opData.active||'yes'];
    if (opData.rowIndex) {
      sheet.getRange(opData.rowIndex,1,1,3).setValues([row]);
      return {success:true, message:"Operator updated!"};
    }
    var data = sheet.getDataRange().getValues();
    for (var i=1; i<data.length; i++) {
      if (data[i][1].toString().trim().toUpperCase() === opData.code.trim().toUpperCase())
        return {success:false, message:'Code "'+opData.code+'" already exists!'};
    }
    sheet.appendRow(row);
    return {success:true, message:'"'+opData.name+'" added!'};
  } catch(e) { return {success:false, message:e.message}; }
}

function deleteWashingOperator(rowIndex) {
  try {
    SpreadsheetApp.openById(WASHING_SS_ID).getSheetByName("WASHING_OPERATORS").getRange(rowIndex,3).setValue("no");
    return {success:true, message:"Operator deactivated!"};
  } catch(e) { return {success:false, message:e.message}; }
}

// ── WASHING SAVE (LOCK) — single function, duplicate removed ──
function saveWashingEntries(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var sheet = SpreadsheetApp.openById(WASHING_SS_ID).getSheetByName("WASHING_DATA");
    if (!sheet) return {success:false, message:"WASHING_DATA sheet nahi mili"};
    var date = payload.date ? new Date(payload.date) : new Date();
    var saved = 0;
    payload.rows.forEach(function(row) {
      if (!row.opName || !row.opName.trim()) return;
      if (!row.qty || n(row.qty) <= 0) return;
      var sr = sheet.getLastRow();
      sheet.appendRow([sr, date, payload.shift, payload.supervisorName,
        row.opName.trim(), row.workType, (row.subCategory||''),
        (row.batchNo||''), n(row.qty)]);
      var nr = sheet.getLastRow();
      sheet.getRange(nr,2).setNumberFormat("dd-mmm-yy");
      if (nr%2===0) sheet.getRange(nr,1,1,9).setBackground("#E6F1FB");
      saved++;
    });
    SpreadsheetApp.flush();
    return {success:true, saved:saved, message:saved+" entries save ho gayi!"};
  } catch(e) { return {success:false, message:"Error: "+e.message}; }
  finally { try{lock.releaseLock();}catch(e2){} }
}

function getWashingReport(params) {
  try {
    var sheet = SpreadsheetApp.openById(WASHING_SS_ID).getSheetByName("WASHING_DATA");
    if (!sheet || sheet.getLastRow()<=1) return {success:true, rows:[], count:0};
    var tz = Session.getScriptTimeZone();
    var fromD = new Date(params.dateFrom); fromD.setHours(0,0,0,0);
    var toD = new Date(params.dateTo); toD.setHours(23,59,59,999);
    var fromStr = Utilities.formatDate(fromD, tz, "yyyy-MM-dd");
    var toStr = Utilities.formatDate(toD, tz, "yyyy-MM-dd");
    var data = sheet.getDataRange().getValues(), out = [];
    for (var i=1; i<data.length; i++) {
      var cell = data[i][1]; if (!cell) continue;
      var rd = cell instanceof Date ? new Date(cell.getTime()) : new Date(cell);
      var rdStr = Utilities.formatDate(rd, tz, "yyyy-MM-dd");
      if (rdStr < fromStr || rdStr > toStr) continue;
      if (params.shift && params.shift !== 'ALL' && data[i][2] !== params.shift) continue;
      var len = data[i].length;
      var cat, sub, batchNo, qty;
      if (len >= 9) {
        cat = String(data[i][5]||''); sub = String(data[i][6]||'');
        batchNo = String(data[i][7]||''); qty = n(data[i][8]);
      } else if (len === 8) {
        cat = String(data[i][5]||''); sub = String(data[i][6]||'');
        batchNo = ''; qty = n(data[i][7]);
      } else {
        cat = String(data[i][5]||''); sub = '';
        batchNo = ''; qty = n(data[i][6]);
      }
      if (params.workType && params.workType !== 'ALL' && cat !== params.workType) continue;
      if (params.operator && params.operator.trim() && data[i][4].toString().toLowerCase().indexOf(params.operator.toLowerCase()) < 0) continue;
      out.push({
        srNo:data[i][0], date:Utilities.formatDate(rd,tz,"dd-MMM-yy"),
        shift:data[i][2], supervisor:data[i][3],
        opName:data[i][4], workType:cat, subCategory:sub, batchNo:batchNo, qty:qty
      });
    }
    if (params.groupBy === 'detail') return {success:true, type:'detail', rows:out, count:out.length};
    var groups = {};
    out.forEach(function(r) {
      var key;
      switch(params.groupBy) {
        case 'daily':    key = r.date; break;
        case 'operator': key = r.opName; break;
        case 'worktype': key = r.workType + (r.subCategory?' - '+r.subCategory:''); break;
        case 'monthly':  key = r.date.substr(3); break;
        default:         key = r.date;
      }
      if (!groups[key]) groups[key] = {key:key, entries:0, bottleWash:0, capWash:0, mediaDiscard:0, total:0, details:[], _batchSeen:{}};
      var g = groups[key];
      g.entries++;
      if (r.workType === 'Bottle Washing') {
        if (r.batchNo) {
          if (!g._batchSeen[r.batchNo]) { g._batchSeen[r.batchNo] = true; g.bottleWash += r.qty; g.total += r.qty; }
        } else { g.bottleWash += r.qty; g.total += r.qty; }
      } else if (r.workType === 'Cap Wash') { g.capWash += r.qty; g.total += r.qty; }
      else if (r.workType === 'Media Discard') { g.mediaDiscard += r.qty; g.total += r.qty; }
      g.details.push(r);
    });
    var summary = Object.keys(groups).map(function(k){
      var g = groups[k];
      return {key:g.key, entries:g.entries, bottleWash:g.bottleWash,
        capWash:g.capWash, mediaDiscard:g.mediaDiscard, total:g.total, details:g.details};
    });
    return {success:true, type:'summary', groupBy:params.groupBy, rows:summary, count:out.length};
  } catch(e) { return {success:false, message:e.message}; }
}

// ════════════════════════════════════════════════════════════
// DISPATCH MODULE
// ════════════════════════════════════════════════════════════
var DISPATCH_SS_ID = "1koOk1hEwM4lBp-5uBzb_NUXZpy3CnJyJ0hpTW9soyQ8";

function initializeDispatch() {
  var ss = SpreadsheetApp.openById(DISPATCH_SS_ID);
  var inw = ss.getSheetByName("DISPATCH_INWARD");
  if (!inw) {
    inw = ss.insertSheet("DISPATCH_INWARD");
    var h = ["Sr.No","Received Date","Crop","Variety","LOT","Cycle","Room","Bottles","Source","Received By","Received On"];
    inw.appendRow(h);
    inw.getRange(1,1,1,h.length).setBackground("#1D9E75").setFontColor("#fff").setFontWeight("bold").setFontSize(10);
    inw.setFrozenRows(1);
    inw.setColumnWidths(1,h.length,110);
    SpreadsheetApp.flush();
  }
  try { var def=ss.getSheetByName("Sheet1"); if(def) ss.deleteSheet(def); } catch(e){}
  SpreadsheetApp.getUi().alert("✅ Dispatch Setup Complete!\n✓ DISPATCH_INWARD");
}

function getDispatchPending(from, to) {
  try {
    var tz = Session.getScriptTimeZone();
    var fromD = new Date(from); fromD.setHours(0,0,0,0);
    var toD = new Date(to); toD.setHours(23,59,59,999);
    var fromStr = Utilities.formatDate(fromD, tz, "yyyy-MM-dd");
    var toStr = Utilities.formatDate(toD, tz, "yyyy-MM-dd");
    var produced = {};
    var exSheet = SpreadsheetApp.openById(PGR_SS_ID).getSheetByName("PGR_EXCHANGE");
    if (exSheet && exSheet.getLastRow() > 1) {
      var ed = exSheet.getDataRange().getValues();
      for (var i=1; i<ed.length; i++) {
        var cell = ed[i][1]; if (!cell) continue;
        var rd;
        if(cell instanceof Date){rd=new Date(cell.getTime());}else{var ds=String(cell);if(ds.indexOf('/')>0){var p=ds.split('/');if(p.length===3)rd=new Date(p[2],p[1]-1,p[0]);else rd=new Date(ds);}else rd=new Date(ds);}
        if(!rd||isNaN(rd.getTime()))continue;
        var rdStr = Utilities.formatDate(rd, tz, "yyyy-MM-dd");
        if (rdStr < fromStr || rdStr > toStr) continue;
        var len = ed[i].length;
        var lot, crop, variety, cycle, dir, pgrRoom, recvQty, status;
        if (len >= 17) {
          lot=String(ed[i][3]||''); crop=String(ed[i][4]||''); variety=String(ed[i][5]||'');
          cycle=String(ed[i][6]||''); dir=String(ed[i][7]||''); pgrRoom=String(ed[i][8]||'');
          recvQty=parseFloat(ed[i][12])||0; status=String(ed[i][13]||'');
        } else {
          lot=String(ed[i][3]||''); crop='—'; variety='—';
          cycle=String(ed[i][4]||''); dir=String(ed[i][5]||''); pgrRoom=String(ed[i][6]||'');
          recvQty=parseFloat(ed[i][9])||0; status=String(ed[i][10]||'');
        }
        if (dir.indexOf('PIR to PGR') < 0) continue;
        if (cycle.toUpperCase().indexOf('ROOT') < 0) continue;
        if (status.toLowerCase() !== 'received') continue;
        if (recvQty <= 0) continue;
        var key = rdStr+'|'+crop+'|'+variety+'|'+lot+'|'+cycle+'|'+pgrRoom;
        if (!produced[key]) produced[key] = {date:Utilities.formatDate(rd,tz,"dd-MMM-yy"),crop:crop,variety:variety,lot:lot,cycle:cycle,room:pgrRoom,bottles:0};
        produced[key].bottles += recvQty;
      }
    }
    var received = {};
    var inSheet = SpreadsheetApp.openById(DISPATCH_SS_ID).getSheetByName("DISPATCH_INWARD");
    if (inSheet && inSheet.getLastRow() > 1) {
      var idata = inSheet.getDataRange().getValues();
      for (var j=1; j<idata.length; j++) {
        if (!idata[j][1]) continue;
        var ird = idata[j][1] instanceof Date ? new Date(idata[j][1].getTime()) : new Date(idata[j][1]);
        var irdStr = Utilities.formatDate(ird, tz, "yyyy-MM-dd");
        var rkey = irdStr+'|'+String(idata[j][2]||'')+'|'+String(idata[j][3]||'')+'|'+String(idata[j][4]||'')+'|'+String(idata[j][5]||'')+'|'+String(idata[j][6]||'');
        received[rkey] = (received[rkey]||0) + (parseFloat(idata[j][7])||0);
      }
    }
    var out = [];
    Object.keys(produced).forEach(function(k){
      var g = produced[k];
      var alreadyRecv = received[k] || 0;
      var remaining = g.bottles - alreadyRecv;
      if (remaining > 0) {
        out.push({date:g.date,crop:g.crop,variety:g.variety,lot:g.lot,cycle:g.cycle,room:g.room,bottles:remaining});
      }
    });
    return {success:true, rows:out};
  } catch(e) { return {success:false, rows:[], message:e.message}; }
}

// ── DISPATCH RECEIVE single (LOCK) ───────────────────────
function receiveDispatchEntry(entry) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var sheet = SpreadsheetApp.openById(DISPATCH_SS_ID).getSheetByName("DISPATCH_INWARD");
    if (!sheet) { initializeDispatch(); sheet = SpreadsheetApp.openById(DISPATCH_SS_ID).getSheetByName("DISPATCH_INWARD"); }
    var sr = sheet.getLastRow();
    var recvDate = entry.date ? new Date(entry.date) : new Date();
    sheet.appendRow([sr, recvDate, entry.crop, entry.variety, entry.lot, entry.cycle,
      entry.room, n(entry.bottles), "PGR", entry.receivedBy||"", new Date(), (entry.remarks||"")]);
    var nr = sheet.getLastRow();
    sheet.getRange(nr,2).setNumberFormat("dd-mmm-yy");
    sheet.getRange(nr,11).setNumberFormat("dd-mmm-yy HH:mm");
    if (nr%2===0) sheet.getRange(nr,1,1,12).setBackground("#E1F5EE");
    SpreadsheetApp.flush();
    return {success:true, message:"Received!"};
  } catch(e) { return {success:false, message:"Error: "+e.message}; }
  finally { try{lock.releaseLock();}catch(e2){} }
}

// ── DISPATCH RECEIVE bulk (LOCK) ─────────────────────────
function receiveDispatchBulk(entries) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var sheet = SpreadsheetApp.openById(DISPATCH_SS_ID).getSheetByName("DISPATCH_INWARD");
    if (!sheet) { initializeDispatch(); sheet = SpreadsheetApp.openById(DISPATCH_SS_ID).getSheetByName("DISPATCH_INWARD"); }
    var saved = 0;
    entries.forEach(function(entry){
      if (!entry.bottles || n(entry.bottles) <= 0) return;
      var sr = sheet.getLastRow();
      var recvDate = entry.date ? new Date(entry.date) : new Date();
      sheet.appendRow([sr, recvDate, entry.crop, entry.variety, entry.lot, entry.cycle,
        entry.room, n(entry.bottles), "PGR", entry.receivedBy||"", new Date(), (entry.remarks||"")]);
      var nr = sheet.getLastRow();
      sheet.getRange(nr,2).setNumberFormat("dd-mmm-yy");
      sheet.getRange(nr,11).setNumberFormat("dd-mmm-yy HH:mm");
      if (nr%2===0) sheet.getRange(nr,1,1,12).setBackground("#E1F5EE");
      saved++;
    });
    SpreadsheetApp.flush();
    return {success:true, saved:saved, message:saved+" entries received!"};
  } catch(e) { return {success:false, message:"Error: "+e.message}; }
  finally { try{lock.releaseLock();}catch(e2){} }
}

function getDispatchStock(from, to) {
  try {
    var sheet = SpreadsheetApp.openById(DISPATCH_SS_ID).getSheetByName("DISPATCH_INWARD");
    if (!sheet || sheet.getLastRow() <= 1) return {success:true, rows:[]};
    var tz = Session.getScriptTimeZone();
    var fromD = new Date(from); fromD.setHours(0,0,0,0);
    var toD = new Date(to); toD.setHours(23,59,59,999);
    var fromStr = Utilities.formatDate(fromD, tz, "yyyy-MM-dd");
    var toStr = Utilities.formatDate(toD, tz, "yyyy-MM-dd");
    var data = sheet.getDataRange().getValues(), out = [];
    for (var i=1; i<data.length; i++) {
      if (!data[i][1]) continue;
      var rd = data[i][1] instanceof Date ? new Date(data[i][1].getTime()) : new Date(data[i][1]);
      var rdStr = Utilities.formatDate(rd, tz, "yyyy-MM-dd");
      if (rdStr < fromStr || rdStr > toStr) continue;
      out.push({date:Utilities.formatDate(rd,tz,"dd-MMM-yy"),crop:data[i][2],variety:data[i][3],
        lot:data[i][4],cycle:data[i][5],room:data[i][6],bottles:n(data[i][7]),
        receivedBy:data[i][9],remarks:String(data[i][11]||'')});
    }
    return {success:true, rows:out};
  } catch(e) { return {success:false, rows:[], message:e.message}; }
}

function getDispatchOperators() {
  try {
    var ss = SpreadsheetApp.openById(DISPATCH_SS_ID);
    var sheet = ss.getSheetByName("DISPATCH_OPERATORS");
    if (!sheet) {
      sheet = ss.insertSheet("DISPATCH_OPERATORS");
      sheet.appendRow(["Operator Name","Operator Code","Active"]);
      sheet.getRange(1,1,1,3).setBackground("#1D9E75").setFontColor("#fff").setFontWeight("bold");
      sheet.setFrozenRows(1);
      SpreadsheetApp.flush();
      return {success:true, operators:[]};
    }
    var data = sheet.getDataRange().getValues(), ops = [];
    for (var i=1; i<data.length; i++) {
      if (!data[i][0]) continue;
      if (data[i][2] && data[i][2].toString().toLowerCase() === 'no') continue;
      ops.push({rowIndex:i+1, name:data[i][0].toString().trim(), code:data[i][1].toString().trim()});
    }
    return {success:true, operators:ops};
  } catch(e) { return {success:false, operators:[], message:e.message}; }
}

function getAllDispatchOperators() {
  try {
    var sheet = SpreadsheetApp.openById(DISPATCH_SS_ID).getSheetByName("DISPATCH_OPERATORS");
    if (!sheet) return {success:false, operators:[]};
    var data = sheet.getDataRange().getValues(), ops = [];
    for (var i=1; i<data.length; i++) {
      if (!data[i][0]) continue;
      ops.push({rowIndex:i+1, name:data[i][0].toString().trim(), code:data[i][1].toString().trim(), active:data[i][2].toString().toLowerCase()});
    }
    return {success:true, operators:ops};
  } catch(e) { return {success:false, operators:[], message:e.message}; }
}

function saveDispatchOperator(opData) {
  try {
    var ss = SpreadsheetApp.openById(DISPATCH_SS_ID);
    var sheet = ss.getSheetByName("DISPATCH_OPERATORS");
    if (!sheet) {
      sheet = ss.insertSheet("DISPATCH_OPERATORS");
      sheet.appendRow(["Operator Name","Operator Code","Active"]);
      sheet.getRange(1,1,1,3).setBackground("#1D9E75").setFontColor("#fff").setFontWeight("bold");
      sheet.setFrozenRows(1);
    }
    var row = [opData.name.trim(), opData.code.trim().toUpperCase(), opData.active||'yes'];
    if (opData.rowIndex) {
      sheet.getRange(opData.rowIndex,1,1,3).setValues([row]);
      return {success:true, message:"Operator updated!"};
    }
    var data = sheet.getDataRange().getValues();
    for (var i=1; i<data.length; i++) {
      if (data[i][1].toString().trim().toUpperCase() === opData.code.trim().toUpperCase())
        return {success:false, message:'Code "'+opData.code+'" already exists!'};
    }
    sheet.appendRow(row);
    return {success:true, message:'"'+opData.name+'" added!'};
  } catch(e) { return {success:false, message:e.message}; }
}

function deleteDispatchOperator(rowIndex) {
  try {
    SpreadsheetApp.openById(DISPATCH_SS_ID).getSheetByName("DISPATCH_OPERATORS").getRange(rowIndex,3).setValue("no");
    return {success:true, message:"Operator deactivated!"};
  } catch(e) { return {success:false, message:e.message}; }
}

// ── DISPATCH CONTAMINATION SAVE (LOCK) ───────────────────
function saveDispatchContamination(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var ss = SpreadsheetApp.openById(DISPATCH_SS_ID);
    var sheet = ss.getSheetByName("DISPATCH_CONTAMINATION");
    if (!sheet) {
      sheet = ss.insertSheet("DISPATCH_CONTAMINATION");
      var h = ["Sr.No","Date","Crop","Variety","LOT","Qty Checked","Total Contaminated","Total %","Operator","Op Checked","Op Contaminated","Op %","Checked By","Entry Time"];
      sheet.appendRow(h);
      sheet.getRange(1,1,1,h.length).setBackground("#854F0B").setFontColor("#fff").setFontWeight("bold").setFontSize(10);
      sheet.setFrozenRows(1);
      sheet.setColumnWidths(1,h.length,105);
      SpreadsheetApp.flush();
    }
    var date = payload.date ? new Date(payload.date) : new Date();
    var ops = payload.operators || [];
    var qChecked = 0, totalConta = 0;
    ops.forEach(function(o){ qChecked += n(o.checked); totalConta += n(o.contaminated); });
    var totalPct = qChecked>0 ? Math.round((totalConta/qChecked)*10000)/100 : 0;
    var saved = 0;
    var entryTime = new Date();
    var isFirst = true;
    ops.forEach(function(o){
      if (!o.opName || !o.opName.trim()) return;
      var opChk = n(o.checked);
      var opCon = n(o.contaminated);
      var opPct = opChk>0 ? Math.round((opCon/opChk)*10000)/100 : 0;
      var sr = sheet.getLastRow();
      var qC = isFirst ? qChecked : "";
      var tC = isFirst ? totalConta : "";
      var tP = isFirst ? totalPct : "";
      sheet.appendRow([sr, date, payload.crop, payload.variety, payload.lot,
        qC, tC, tP,
        o.opName.trim(), opChk, opCon, opPct,
        payload.checkedBy||"", entryTime]);
      var nr = sheet.getLastRow();
      sheet.getRange(nr,2).setNumberFormat("dd-mmm-yy");
      sheet.getRange(nr,14).setNumberFormat("dd-mmm-yy HH:mm");
      if (nr%2===0) sheet.getRange(nr,1,1,14).setBackground("#FAEEDA");
      isFirst = false;
      saved++;
    });
    SpreadsheetApp.flush();
    return {success:true, saved:saved, message:saved+" operators saved! Checked: "+qChecked+", Conta: "+totalConta+" ("+totalPct+"%)"};
  } catch(e) { return {success:false, message:"Error: "+e.message}; }
  finally { try{lock.releaseLock();}catch(e2){} }
}

function getDispatchReceivedLots(crop, variety) {
  try {
    var sheet = SpreadsheetApp.openById(DISPATCH_SS_ID).getSheetByName("DISPATCH_INWARD");
    if (!sheet || sheet.getLastRow() <= 1) return {success:true, lots:[]};
    var data = sheet.getDataRange().getValues();
    var seen = {};
    var lots = [];
    for (var i=1; i<data.length; i++) {
      if (!data[i][1]) continue;
      var c = String(data[i][2]||'').trim();
      var v = String(data[i][3]||'').trim();
      var lot = String(data[i][4]||'').trim();
      if (!lot) continue;
      if (crop && c.toLowerCase() !== crop.toLowerCase()) continue;
      if (variety && v && variety && v.toLowerCase() !== variety.toLowerCase()) continue;
      if (!seen[lot]) { seen[lot] = true; lots.push(lot); }
    }
    lots.sort();
    return {success:true, lots:lots};
  } catch(e) { return {success:false, lots:[], message:e.message}; }
}

// ── DISPATCH OUTWARD SAVE (LOCK) ─────────────────────────
function saveDispatchOutward(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var ss = SpreadsheetApp.openById(DISPATCH_SS_ID);
    var sheet = ss.getSheetByName("DISPATCH_OUTWARD");
    if (!sheet) {
      sheet = ss.insertSheet("DISPATCH_OUTWARD");
      var h = ["Sr.No","Dispatch Date","Crop","Variety","LOT","Cutting From","Cutting To","Tray","Bottle/Tray","Extra Bottle","Total Bottle","Dispatched By","Entry Time"];
      sheet.appendRow(h);
      sheet.getRange(1,1,1,h.length).setBackground("#1D9E75").setFontColor("#fff").setFontWeight("bold").setFontSize(10);
      sheet.setFrozenRows(1);
      sheet.setColumnWidths(1,h.length,105);
      SpreadsheetApp.flush();
    }
    var date = payload.date ? new Date(payload.date) : new Date();
    var rows = payload.rows || [];
    var saved = 0;
    var entryTime = new Date();
    rows.forEach(function(r){
      if (!r.crop || !r.lot) return;
      if (!r.totalBottle || n(r.totalBottle) <= 0) return;
      var sr = sheet.getLastRow();
      var cFrom = r.cuttingFrom ? new Date(r.cuttingFrom) : "";
      var cTo = r.cuttingTo ? new Date(r.cuttingTo) : "";
      sheet.appendRow([sr, date, r.crop, r.variety, r.lot,
        cFrom, cTo, n(r.tray), n(r.bottlePerTray), n(r.extraBottle), n(r.totalBottle),
        payload.dispatchedBy||"", entryTime]);
      var nr = sheet.getLastRow();
      sheet.getRange(nr,2).setNumberFormat("dd-mmm-yy");
      if (cFrom) sheet.getRange(nr,6).setNumberFormat("dd-mmm-yy");
      if (cTo) sheet.getRange(nr,7).setNumberFormat("dd-mmm-yy");
      sheet.getRange(nr,13).setNumberFormat("dd-mmm-yy HH:mm");
      if (nr%2===0) sheet.getRange(nr,1,1,13).setBackground("#E1F5EE");
      saved++;
    });
    SpreadsheetApp.flush();
    return {success:true, saved:saved, message:saved+" rows dispatch ho gaye!"};
  } catch(e) { return {success:false, message:"Error: "+e.message}; }
  finally { try{lock.releaseLock();}catch(e2){} }
}

function getDispatchOutward(from, to) {
  try {
    var sheet = SpreadsheetApp.openById(DISPATCH_SS_ID).getSheetByName("DISPATCH_OUTWARD");
    if (!sheet || sheet.getLastRow() <= 1) return {success:true, rows:[]};
    var tz = Session.getScriptTimeZone();
    var fromD = new Date(from); fromD.setHours(0,0,0,0);
    var toD = new Date(to); toD.setHours(23,59,59,999);
    var fromStr = Utilities.formatDate(fromD, tz, "yyyy-MM-dd");
    var toStr = Utilities.formatDate(toD, tz, "yyyy-MM-dd");
    var data = sheet.getDataRange().getValues(), out = [];
    for (var i=1; i<data.length; i++) {
      if (!data[i][1]) continue;
      var rd = data[i][1] instanceof Date ? new Date(data[i][1].getTime()) : new Date(data[i][1]);
      var rdStr = Utilities.formatDate(rd, tz, "yyyy-MM-dd");
      if (rdStr < fromStr || rdStr > toStr) continue;
      var cf = data[i][5] instanceof Date ? Utilities.formatDate(data[i][5],tz,"dd-MMM-yy") : "";
      var ct = data[i][6] instanceof Date ? Utilities.formatDate(data[i][6],tz,"dd-MMM-yy") : "";
      out.push({date:Utilities.formatDate(rd,tz,"dd-MMM-yy"),crop:data[i][2],variety:data[i][3],
        lot:data[i][4],cuttingFrom:cf,cuttingTo:ct,tray:n(data[i][7]),bottlePerTray:n(data[i][8]),
        extraBottle:n(data[i][9]),totalBottle:n(data[i][10]),dispatchedBy:data[i][11]});
    }
    return {success:true, rows:out};
  } catch(e) { return {success:false, rows:[], message:e.message}; }
}

// ── DISPATCH OPENING (LOCK) ──────────────────────────────
function saveDispatchOpening(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var ss = SpreadsheetApp.openById(DISPATCH_SS_ID);
    var sheet = ss.getSheetByName("DISPATCH_OPENING");
    if (!sheet) {
      sheet = ss.insertSheet("DISPATCH_OPENING");
      var h = ["Sr.No","As On Date","Crop","Variety","LOT","Opening Bottles","Added By","Entry Time"];
      sheet.appendRow(h);
      sheet.getRange(1,1,1,h.length).setBackground("#6D28D9").setFontColor("#fff").setFontWeight("bold").setFontSize(10);
      sheet.setFrozenRows(1);
      sheet.setColumnWidths(1,h.length,115);
      SpreadsheetApp.flush();
    }
    var date = payload.date ? new Date(payload.date) : new Date();
    var rows = payload.rows || [];
    var saved = 0;
    var entryTime = new Date();
    rows.forEach(function(r){
      if (!r.crop || !r.lot) return;
      if (n(r.opening) <= 0) return;
      var sr = sheet.getLastRow();
      sheet.appendRow([sr, date, r.crop, r.variety, r.lot, n(r.opening), payload.addedBy||"", entryTime]);
      var nr = sheet.getLastRow();
      sheet.getRange(nr,2).setNumberFormat("dd-mmm-yy");
      sheet.getRange(nr,8).setNumberFormat("dd-mmm-yy HH:mm");
      if (nr%2===0) sheet.getRange(nr,1,1,8).setBackground("#EDE9FE");
      saved++;
    });
    SpreadsheetApp.flush();
    return {success:true, saved:saved, message:saved+" opening stock entries saved!"};
  } catch(e) { return {success:false, message:"Error: "+e.message}; }
  finally { try{lock.releaseLock();}catch(e2){} }
}

function getDispatchStockReport(params) {
  try {
    var tz = Session.getScriptTimeZone();
    var toD = new Date(params.dateTo); toD.setHours(23,59,59,999);
    var toStr = Utilities.formatDate(toD, tz, "yyyy-MM-dd");
    var fCrop = (params.crop||'').toLowerCase();
    var fVariety = (params.variety||'').toLowerCase();
    var fLot = (params.lot||'').toLowerCase();
    var stock = {};
    function getKey(crop,variety,lot){ return crop+'|'+variety+'|'+lot; }
    function ensure(crop,variety,lot){
      var k = getKey(crop,variety,lot);
      if (!stock[k]) stock[k] = {crop:crop,variety:variety,lot:lot,opening:0,inward:0,conta:0,dispatch:0};
      return stock[k];
    }
    function dateOk(cell){
      if (!cell) return false;
      var rd = cell instanceof Date ? new Date(cell.getTime()) : new Date(cell);
      if (isNaN(rd.getTime())) return false;
      return Utilities.formatDate(rd, tz, "yyyy-MM-dd") <= toStr;
    }
    var ss = SpreadsheetApp.openById(DISPATCH_SS_ID);
    var op = ss.getSheetByName("DISPATCH_OPENING");
    if (op && op.getLastRow()>1) {
      var od = op.getDataRange().getValues();
      for (var i=1;i<od.length;i++){
        if (!dateOk(od[i][1])) continue;
        var c=String(od[i][2]||''),v=String(od[i][3]||''),l=String(od[i][4]||'');
        ensure(c,v,l).opening += (parseFloat(od[i][5])||0);
      }
    }
    var inw = ss.getSheetByName("DISPATCH_INWARD");
    if (inw && inw.getLastRow()>1) {
      var id = inw.getDataRange().getValues();
      for (var j=1;j<id.length;j++){
        if (!dateOk(id[j][1])) continue;
        var c2=String(id[j][2]||''),v2=String(id[j][3]||''),l2=String(id[j][4]||'');
        ensure(c2,v2,l2).inward += (parseFloat(id[j][7])||0);
      }
    }
    var con = ss.getSheetByName("DISPATCH_CONTAMINATION");
    if (con && con.getLastRow()>1) {
      var cd = con.getDataRange().getValues();
      for (var k2=1;k2<cd.length;k2++){
        if (!dateOk(cd[k2][1])) continue;
        var tc = cd[k2][6];
        if (tc==='' || tc===null) continue;
        var c3=String(cd[k2][2]||''),v3=String(cd[k2][3]||''),l3=String(cd[k2][4]||'');
        ensure(c3,v3,l3).conta += (parseFloat(tc)||0);
      }
    }
    var out = ss.getSheetByName("DISPATCH_OUTWARD");
    if (out && out.getLastRow()>1) {
      var otd = out.getDataRange().getValues();
      for (var m=1;m<otd.length;m++){
        if (!dateOk(otd[m][1])) continue;
        var c4=String(otd[m][2]||''),v4=String(otd[m][3]||''),l4=String(otd[m][4]||'');
        ensure(c4,v4,l4).dispatch += (parseFloat(otd[m][10])||0);
      }
    }
    var rows = [];
    Object.keys(stock).forEach(function(key){
      var s = stock[key];
      if (fCrop && s.crop.toLowerCase()!==fCrop) return;
      if (fVariety && s.variety.toLowerCase()!==fVariety) return;
      if (fLot && s.lot.toLowerCase().indexOf(fLot)<0) return;
      s.available = s.opening + s.inward - s.conta - s.dispatch;
      if (s.opening===0 && s.inward===0 && s.conta===0 && s.dispatch===0) return;
      rows.push(s);
    });
    rows.sort(function(a,b){ return (a.crop+a.variety+a.lot).localeCompare(b.crop+b.variety+b.lot); });
    return {success:true, rows:rows};
  } catch(e) { return {success:false, rows:[], message:e.message}; }
}

function getDispatchStockFilters() {
  try {
    var ss = SpreadsheetApp.openById(DISPATCH_SS_ID);
    var crops={}, lots={};
    ["DISPATCH_INWARD","DISPATCH_OPENING","DISPATCH_OUTWARD"].forEach(function(name){
      var sh = ss.getSheetByName(name);
      if (!sh || sh.getLastRow()<=1) return;
      var d = sh.getDataRange().getValues();
      for (var i=1;i<d.length;i++){
        var c=String(d[i][2]||'').trim(), l=String(d[i][4]||'').trim();
        if (c) crops[c]=true;
        if (l) lots[l]=true;
      }
    });
    return {success:true, crops:Object.keys(crops).sort(), lots:Object.keys(lots).sort()};
  } catch(e) { return {success:false, crops:[], lots:[]}; }
}

// ════════════════════════════════════════════════════════════
// STORE ROOM MODULE (Category + Unit-in-master + conversion + LOCK)
// ════════════════════════════════════════════════════════════
var STORE_SS_ID = "1_hbBsbdOoz3zalqsJC6lGdcxzKYuRD_BKY5-Jpj4Qq0";

// ── STORE ITEMS (Item + Category + Unit) ──
function getStoreItems() {
  try {
    var ss = SpreadsheetApp.openById(STORE_SS_ID);
    var sheet = ss.getSheetByName("STORE_ITEMS");
    if (!sheet) {
      sheet = ss.insertSheet("STORE_ITEMS");
      sheet.appendRow(["Item Name","Category","Unit"]);
      sheet.getRange(1,1,1,3).setBackground("#1D9E75").setFontColor("#fff").setFontWeight("bold");
      sheet.setFrozenRows(1);
      SpreadsheetApp.flush();
      return {success:true, items:[], itemsWithCat:[], categories:[]};
    }
    var data = sheet.getDataRange().getValues(), items = [], itemsWithCat = [], catSet = {};
    for (var i=1; i<data.length; i++) {
      var it = String(data[i][0]||'').trim();
      var cat = String(data[i][1]||'').trim();
      var unit = String(data[i][2]||'').trim();
      if (!it) continue;
      items.push(it);
      itemsWithCat.push({item:it, category:cat, unit:unit});
      if (cat) catSet[cat] = true;
    }
    return {success:true, items:items, itemsWithCat:itemsWithCat, categories:Object.keys(catSet).sort()};
  } catch(e) { return {success:false, items:[], itemsWithCat:[], categories:[], message:e.message}; }
}

function addStoreItemIfNew(item, category, unit) {
  if (!item) return;
  var ss = SpreadsheetApp.openById(STORE_SS_ID);
  var sheet = ss.getSheetByName("STORE_ITEMS");
  if (!sheet) { sheet = ss.insertSheet("STORE_ITEMS"); sheet.appendRow(["Item Name","Category","Unit"]); sheet.setFrozenRows(1); }
  var data = sheet.getDataRange().getValues();
  for (var i=1; i<data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === item.trim().toLowerCase()) {
      if (category && !String(data[i][1]||'').trim()) sheet.getRange(i+1,2).setValue(category);
      if (unit && !String(data[i][2]||'').trim()) sheet.getRange(i+1,3).setValue(unit);
      return;
    }
  }
  sheet.appendRow([item.trim(), category||"", unit||""]);
}

// ── PURCHASE (Category + Unit-to-master + LOCK) ──
function saveStorePurchase(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var ss = SpreadsheetApp.openById(STORE_SS_ID);
    var sheet = ss.getSheetByName("STORE_PURCHASE");
    if (!sheet) {
      sheet = ss.insertSheet("STORE_PURCHASE");
      var h = ["Sr.No","Date","Item","Category","Quantity","Unit","Rate/Amount","Supplier","Bill No","Added By","Entry Time"];
      sheet.appendRow(h);
      sheet.getRange(1,1,1,h.length).setBackground("#1D9E75").setFontColor("#fff").setFontWeight("bold").setFontSize(10);
      sheet.setFrozenRows(1); sheet.setColumnWidths(1,h.length,110);
      SpreadsheetApp.flush();
    }
    var date = payload.date ? new Date(payload.date) : new Date();
    var supplier = payload.supplier || "", billNo = payload.billNo || "";
    var saved = 0, entryTime = new Date();
    payload.rows.forEach(function(r){
      if (!r.item || n(r.qty)<=0) return;
      var sr = sheet.getLastRow();
      sheet.appendRow([sr, date, r.item, r.category||"", n(r.qty), r.unit||"", n(r.rate)||"", supplier, billNo, payload.addedBy||"", entryTime]);
      var nr = sheet.getLastRow();
      sheet.getRange(nr,2).setNumberFormat("dd-mmm-yy");
      sheet.getRange(nr,11).setNumberFormat("dd-mmm-yy HH:mm");
      if (nr%2===0) sheet.getRange(nr,1,1,11).setBackground("#E1F5EE");
      addStoreItemIfNew(r.item, r.category, r.unit);
      saved++;
    });
    SpreadsheetApp.flush();
    return {success:true, saved:saved, message:saved+" items purchase save!"};
  } catch(e) { return {success:false, message:"Error: "+e.message}; }
  finally { try{lock.releaseLock();}catch(e2){} }
}

// ── ISSUE (Category + LOCK) ──
function saveStoreIssue(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var ss = SpreadsheetApp.openById(STORE_SS_ID);
    var sheet = ss.getSheetByName("STORE_ISSUE");
    if (!sheet) {
      sheet = ss.insertSheet("STORE_ISSUE");
      var h = ["Sr.No","Date","Department","Item","Category","Quantity","Unit","Issued To","Added By","Entry Time"];
      sheet.appendRow(h);
      sheet.getRange(1,1,1,h.length).setBackground("#2563EB").setFontColor("#fff").setFontWeight("bold").setFontSize(10);
      sheet.setFrozenRows(1); sheet.setColumnWidths(1,h.length,110);
      SpreadsheetApp.flush();
    }
    var date = payload.date ? new Date(payload.date) : new Date();
    var saved = 0, entryTime = new Date();
    payload.rows.forEach(function(r){
      if (!r.department || !r.item || n(r.qty)<=0) return;
      var sr = sheet.getLastRow();
      sheet.appendRow([sr, date, r.department, r.item, r.category||"", n(r.qty), r.unit||"", payload.issuedTo||"", payload.addedBy||"", entryTime]);
      var nr = sheet.getLastRow();
      sheet.getRange(nr,2).setNumberFormat("dd-mmm-yy");
      sheet.getRange(nr,10).setNumberFormat("dd-mmm-yy HH:mm");
      if (nr%2===0) sheet.getRange(nr,1,1,10).setBackground("#DBEAFE");
      addStoreItemIfNew(r.item, r.category, r.unit);
      saved++;
    });
    SpreadsheetApp.flush();
    return {success:true, saved:saved, message:saved+" items issue save!"};
  } catch(e) { return {success:false, message:"Error: "+e.message}; }
  finally { try{lock.releaseLock();}catch(e2){} }
}

// ── OPENING (Category + LOCK) ──
function saveStoreOpening(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var ss = SpreadsheetApp.openById(STORE_SS_ID);
    var sheet = ss.getSheetByName("STORE_OPENING");
    if (!sheet) {
      sheet = ss.insertSheet("STORE_OPENING");
      var h = ["Sr.No","As On Date","Item","Category","Opening Qty","Unit","Added By","Entry Time"];
      sheet.appendRow(h);
      sheet.getRange(1,1,1,h.length).setBackground("#6D28D9").setFontColor("#fff").setFontWeight("bold").setFontSize(10);
      sheet.setFrozenRows(1); sheet.setColumnWidths(1,h.length,110);
      SpreadsheetApp.flush();
    }
    var date = payload.date ? new Date(payload.date) : new Date();
    var saved = 0, entryTime = new Date();
    payload.rows.forEach(function(r){
      if (!r.item || n(r.qty)<=0) return;
      var sr = sheet.getLastRow();
      sheet.appendRow([sr, date, r.item, r.category||"", n(r.qty), r.unit||"", payload.addedBy||"", entryTime]);
      var nr = sheet.getLastRow();
      sheet.getRange(nr,2).setNumberFormat("dd-mmm-yy");
      sheet.getRange(nr,8).setNumberFormat("dd-mmm-yy HH:mm");
      if (nr%2===0) sheet.getRange(nr,1,1,8).setBackground("#EDE9FE");
      addStoreItemIfNew(r.item, r.category, r.unit);
      saved++;
    });
    SpreadsheetApp.flush();
    return {success:true, saved:saved, message:saved+" opening items save!"};
  } catch(e) { return {success:false, message:"Error: "+e.message}; }
  finally { try{lock.releaseLock();}catch(e2){} }
}

// ── STOCK REPORT (Category + unit conversion + filter) ──
function getStoreStockReport(params) {
  try {
    var tz = Session.getScriptTimeZone();
    var toD = new Date(params.dateTo); toD.setHours(23,59,59,999);
    var toStr = Utilities.formatDate(toD, tz, "yyyy-MM-dd");
    var fItem = (params.item||'').toLowerCase();
    var fCat = (params.category||'').toLowerCase();

    function normalize(qty, unit) {
      var u = String(unit||'').toLowerCase().trim();
      if (u==='gram'||u==='gm'||u==='g') return {qty:qty/1000, unit:'kg'};
      if (u==='kg'||u==='kilogram') return {qty:qty, unit:'kg'};
      if (u==='ml'||u==='millilitre'||u==='milliliter') return {qty:qty/1000, unit:'litre'};
      if (u==='litre'||u==='liter'||u==='l') return {qty:qty, unit:'litre'};
      return {qty:qty, unit:u||'pcs'};
    }

    var itemCat = {};
    var itSheet = SpreadsheetApp.openById(STORE_SS_ID).getSheetByName("STORE_ITEMS");
    if (itSheet && itSheet.getLastRow()>1) {
      var itd = itSheet.getDataRange().getValues();
      for (var x=1;x<itd.length;x++){ var inm=String(itd[x][0]||'').trim().toLowerCase(); if(inm) itemCat[inm]=String(itd[x][1]||'').trim(); }
    }

    var stock = {};
    function ensure(item, baseUnit) {
      var k = item.toLowerCase()+'|'+baseUnit;
      if (!stock[k]) stock[k] = {item:item, category:(itemCat[item.toLowerCase()]||''), unit:baseUnit, opening:0, purchase:0, issue:0};
      return stock[k];
    }
    function dateOk(cell){
      if (!cell) return false;
      var rd = cell instanceof Date ? new Date(cell.getTime()) : new Date(cell);
      if (isNaN(rd.getTime())) return false;
      return Utilities.formatDate(rd, tz, "yyyy-MM-dd") <= toStr;
    }
    var ss = SpreadsheetApp.openById(STORE_SS_ID);

    // Opening: Date(1) Item(2) Cat(3) Qty(4) Unit(5) [new] | old: Date(1) Item(2) Qty(3) Unit(4)
    var op = ss.getSheetByName("STORE_OPENING");
    if (op && op.getLastRow()>1) {
      var od = op.getDataRange().getValues();
      var oNew = od[0].length >= 8;
      for (var i=1;i<od.length;i++){ if(!dateOk(od[i][1]))continue; var it=String(od[i][2]||'').trim(); if(!it)continue;
        var q=oNew?n(od[i][4]):n(od[i][3]); var u=oNew?od[i][5]:od[i][4];
        var nm=normalize(q,u); ensure(it,nm.unit).opening += nm.qty; }
    }
    // Purchase: Date(1) Item(2) Cat(3) Qty(4) Unit(5) [new] | old: Date(1) Item(2) Qty(3) Unit(4)
    var pu = ss.getSheetByName("STORE_PURCHASE");
    if (pu && pu.getLastRow()>1) {
      var pd = pu.getDataRange().getValues();
      var pNew = pd[0].length >= 11;
      for (var j=1;j<pd.length;j++){ if(!dateOk(pd[j][1]))continue; var it2=String(pd[j][2]||'').trim(); if(!it2)continue;
        var q2=pNew?n(pd[j][4]):n(pd[j][3]); var u2=pNew?pd[j][5]:pd[j][4];
        var nm2=normalize(q2,u2); ensure(it2,nm2.unit).purchase += nm2.qty; }
    }
    // Issue: Date(1) Dept(2) Item(3) Cat(4) Qty(5) Unit(6) [new] | old: Date(1) Dept(2) Item(3) Qty(4) Unit(5)
    var iss = ss.getSheetByName("STORE_ISSUE");
    if (iss && iss.getLastRow()>1) {
      var idd = iss.getDataRange().getValues();
      var iNew = idd[0].length >= 10;
      for (var k=1;k<idd.length;k++){ if(!dateOk(idd[k][1]))continue; var it3=String(idd[k][3]||'').trim(); if(!it3)continue;
        var q3=iNew?n(idd[k][5]):n(idd[k][4]); var u3=iNew?idd[k][6]:idd[k][5];
        var nm3=normalize(q3,u3); ensure(it3,nm3.unit).issue += nm3.qty; }
    }

    var rows = [];
    Object.keys(stock).forEach(function(key){
      var s = stock[key];
      if (fItem && s.item.toLowerCase().indexOf(fItem)<0) return;
      if (fCat && s.category.toLowerCase()!==fCat) return;
      s.available = Math.round((s.opening + s.purchase - s.issue)*1000)/1000;
      s.opening = Math.round(s.opening*1000)/1000;
      s.purchase = Math.round(s.purchase*1000)/1000;
      s.issue = Math.round(s.issue*1000)/1000;
      if (s.opening===0 && s.purchase===0 && s.issue===0) return;
      rows.push(s);
    });
    rows.sort(function(a,b){ return (a.category+a.item+a.unit).localeCompare(b.category+b.item+b.unit); });
    return {success:true, rows:rows};
  } catch(e) { return {success:false, rows:[], message:e.message}; }
}


// ════════════════════════════════════════════════════════════
// MPR MODULE (Media Preparation) — Production + Issue + Stock + LOCK
// ════════════════════════════════════════════════════════════
var MPR_SS_ID = "19DT2ywzMr6Ocss7ZxbMVeKdafY6SGyMkxAoKZVugII4";   // MPR Google Sheet ID (sirf ID, poora link nahi)

// ── MEDIA TYPES (master) ──
function getMprTypes() {
  try {
    var ss = SpreadsheetApp.openById(MPR_SS_ID);
    var sheet = ss.getSheetByName("MPR_TYPES");
    if (!sheet) {
      sheet = ss.insertSheet("MPR_TYPES");
      sheet.appendRow(["Media Type","Recipe/Notes","Active"]);
      sheet.getRange(1,1,1,3).setBackground("#1D9E75").setFontColor("#fff").setFontWeight("bold");
      sheet.setFrozenRows(1); SpreadsheetApp.flush();
      return {success:true, types:[], typesWithNote:[]};
    }
    var data = sheet.getDataRange().getValues(), types=[], withNote=[];
    for (var i=1;i<data.length;i++){
      var t=String(data[i][0]||'').trim(), note=String(data[i][1]||'').trim(), active=String(data[i][2]||'yes').trim().toLowerCase();
      if(!t||active==='no')continue;
      types.push(t); withNote.push({type:t,note:note});
    }
    return {success:true, types:types, typesWithNote:withNote};
  } catch(e){ return {success:false, types:[], typesWithNote:[], message:e.message}; }
}
function getAllMprTypes() {
  try {
    var sheet = SpreadsheetApp.openById(MPR_SS_ID).getSheetByName("MPR_TYPES");
    if(!sheet) return {success:true, types:[]};
    var data=sheet.getDataRange().getValues(), out=[];
    for(var i=1;i<data.length;i++){
      if(!data[i][0])continue;
      out.push({rowIndex:i+1,type:String(data[i][0]).trim(),note:String(data[i][1]||'').trim(),active:String(data[i][2]||'yes').toLowerCase()});
    }
    return {success:true, types:out};
  } catch(e){ return {success:false, types:[], message:e.message}; }
}
function addMprTypeIfNew(type, note) {
  if(!type) return;
  var ss=SpreadsheetApp.openById(MPR_SS_ID);
  var sheet=ss.getSheetByName("MPR_TYPES");
  if(!sheet){sheet=ss.insertSheet("MPR_TYPES");sheet.appendRow(["Media Type","Recipe/Notes","Active"]);sheet.setFrozenRows(1);}
  var data=sheet.getDataRange().getValues();
  for(var i=1;i<data.length;i++){
    if(String(data[i][0]).trim().toLowerCase()===type.trim().toLowerCase()){
      if(note&&!String(data[i][1]||'').trim())sheet.getRange(i+1,2).setValue(note);
      return;
    }
  }
  sheet.appendRow([type.trim(), note||"", "yes"]);
}
function saveMprType(data) {
  try {
    var ss=SpreadsheetApp.openById(MPR_SS_ID);
    var sheet=ss.getSheetByName("MPR_TYPES");
    if(!sheet){sheet=ss.insertSheet("MPR_TYPES");sheet.appendRow(["Media Type","Recipe/Notes","Active"]);sheet.setFrozenRows(1);}
    var row=[data.type.trim(), data.note||"", data.active||"yes"];
    if(data.rowIndex){ sheet.getRange(data.rowIndex,1,1,3).setValues([row]); return {success:true,message:"Media type updated!"}; }
    var d=sheet.getDataRange().getValues();
    for(var i=1;i<d.length;i++){ if(String(d[i][0]).trim().toLowerCase()===data.type.trim().toLowerCase()) return {success:false,message:'"'+data.type+'" pehle se hai!'}; }
    sheet.appendRow(row);
    return {success:true,message:'"'+data.type+'" add ho gaya!'};
  } catch(e){ return {success:false,message:e.message}; }
}
function deleteMprType(rowIndex) {
  try { SpreadsheetApp.openById(MPR_SS_ID).getSheetByName("MPR_TYPES").getRange(rowIndex,3).setValue("no"); return {success:true,message:"Deactivate ho gaya!"}; }
  catch(e){ return {success:false,message:e.message}; }
}

// ── MPR OPERATORS ──
function getMprOperators() {
  try {
    var ss=SpreadsheetApp.openById(MPR_SS_ID);
    var sheet=ss.getSheetByName("MPR_OPERATORS");
    if(!sheet){sheet=ss.insertSheet("MPR_OPERATORS");sheet.appendRow(["Operator Name","Operator Code","Active"]);sheet.getRange(1,1,1,3).setBackground("#1D9E75").setFontColor("#fff").setFontWeight("bold");sheet.setFrozenRows(1);SpreadsheetApp.flush();return {success:true,operators:[]};}
    var data=sheet.getDataRange().getValues(), ops=[];
    for(var i=1;i<data.length;i++){
      if(!data[i][0])continue;
      if(data[i][2]&&data[i][2].toString().toLowerCase()==='no')continue;
      ops.push({rowIndex:i+1,name:data[i][0].toString().trim(),code:data[i][1].toString().trim()});
    }
    return {success:true,operators:ops};
  } catch(e){ return {success:false,operators:[],message:e.message}; }
}
function getAllMprOperators() {
  try {
    var sheet=SpreadsheetApp.openById(MPR_SS_ID).getSheetByName("MPR_OPERATORS");
    if(!sheet) return {success:true,operators:[]};
    var data=sheet.getDataRange().getValues(), ops=[];
    for(var i=1;i<data.length;i++){
      if(!data[i][0])continue;
      ops.push({rowIndex:i+1,name:data[i][0].toString().trim(),code:data[i][1].toString().trim(),active:data[i][2].toString().toLowerCase()});
    }
    return {success:true,operators:ops};
  } catch(e){ return {success:false,operators:[],message:e.message}; }
}
function saveMprOperator(opData) {
  try {
    var ss=SpreadsheetApp.openById(MPR_SS_ID);
    var sheet=ss.getSheetByName("MPR_OPERATORS");
    if(!sheet){sheet=ss.insertSheet("MPR_OPERATORS");sheet.appendRow(["Operator Name","Operator Code","Active"]);sheet.setFrozenRows(1);}
    var row=[opData.name.trim(),opData.code.trim().toUpperCase(),opData.active||'yes'];
    if(opData.rowIndex){sheet.getRange(opData.rowIndex,1,1,3).setValues([row]);return {success:true,message:"Operator updated!"};}
    var data=sheet.getDataRange().getValues();
    for(var i=1;i<data.length;i++){ if(data[i][1].toString().trim().toUpperCase()===opData.code.trim().toUpperCase()) return {success:false,message:'Code "'+opData.code+'" already exists!'}; }
    sheet.appendRow(row);
    return {success:true,message:'"'+opData.name+'" added!'};
  } catch(e){ return {success:false,message:e.message}; }
}
function deleteMprOperator(rowIndex) {
  try { SpreadsheetApp.openById(MPR_SS_ID).getSheetByName("MPR_OPERATORS").getRange(rowIndex,3).setValue("no"); return {success:true,message:"Operator deactivated!"}; }
  catch(e){ return {success:false,message:e.message}; }
}

// ── PRODUCTION (banaya) — LOCK ──
function saveMprProduction(payload) {
  var lock=LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var ss=SpreadsheetApp.openById(MPR_SS_ID);
    var sheet=ss.getSheetByName("MPR_PRODUCTION");
    if(!sheet){
      sheet=ss.insertSheet("MPR_PRODUCTION");
      var h=["Sr.No","Date","Shift","Media Type","Batch No","Container","Total Filled","Autoclave Temp","Autoclave Time","Autoclave Done","Operator","Op Code","Supervisor","Entry Time"];
      sheet.appendRow(h);
      sheet.getRange(1,1,1,h.length).setBackground("#0F6E56").setFontColor("#fff").setFontWeight("bold").setFontSize(10);
      sheet.setFrozenRows(1); sheet.setColumnWidths(1,h.length,110); SpreadsheetApp.flush();
    }
    var tz=Session.getScriptTimeZone();
    var date=payload.date?new Date(payload.date):new Date();
    var dateStr=Utilities.formatDate(date,tz,"yyyyMMdd");
    var seq=0, ex=sheet.getDataRange().getValues();
    for(var i=1;i<ex.length;i++){ var b=String(ex[i][4]||''); if(b.indexOf('MB'+dateStr+'-')===0){var num=parseInt(b.split('-')[1])||0;if(num>seq)seq=num;} }
    var saved=0, entryTime=new Date(), batches=[];
    payload.rows.forEach(function(r){
      if(!r.opName||!r.opName.trim())return;
      if(!r.mediaType||n(r.filled)<=0)return;
      seq++;
      var batchNo='MB'+dateStr+'-'+seq;
      var sr=sheet.getLastRow();
      sheet.appendRow([sr,date,payload.shift,r.mediaType,batchNo,(r.container||''),n(r.filled),(r.acTemp||''),(r.acTime||''),(r.acDone||''),r.opName.trim(),(r.opCode||'').toString().toUpperCase(),payload.supervisorName||'',entryTime]);
      var nr=sheet.getLastRow();
      sheet.getRange(nr,2).setNumberFormat("dd-mmm-yy");
      sheet.getRange(nr,14).setNumberFormat("dd-mmm-yy HH:mm");
      if(nr%2===0)sheet.getRange(nr,1,1,14).setBackground("#E1F5EE");
      addMprTypeIfNew(r.mediaType,'');
      batches.push(batchNo); saved++;
    });
    SpreadsheetApp.flush();
    return {success:true,saved:saved,message:saved+" media batch save! ("+(batches.join(', ')||'-')+")"};
  } catch(e){ return {success:false,message:"Error: "+e.message}; }
  finally{ try{lock.releaseLock();}catch(e2){} }
}

// ── ISSUE (use hua) — LOCK ──
function saveMprIssue(payload) {
  var lock=LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var ss=SpreadsheetApp.openById(MPR_SS_ID);
    var sheet=ss.getSheetByName("MPR_ISSUE");
    if(!sheet){
      sheet=ss.insertSheet("MPR_ISSUE");
      var h=["Sr.No","Date","Media Type","Container","Qty Issued","Department","Issued To","Added By","Entry Time"];
      sheet.appendRow(h);
      sheet.getRange(1,1,1,h.length).setBackground("#2563EB").setFontColor("#fff").setFontWeight("bold").setFontSize(10);
      sheet.setFrozenRows(1); sheet.setColumnWidths(1,h.length,110); SpreadsheetApp.flush();
    }
    var date=payload.date?new Date(payload.date):new Date();
    var saved=0, entryTime=new Date();
    payload.rows.forEach(function(r){
      if(!r.mediaType||!r.department||n(r.qty)<=0)return;
      var sr=sheet.getLastRow();
      sheet.appendRow([sr,date,r.mediaType,(r.container||''),n(r.qty),r.department,(payload.issuedTo||''),(payload.addedBy||''),entryTime]);
      var nr=sheet.getLastRow();
      sheet.getRange(nr,2).setNumberFormat("dd-mmm-yy");
      sheet.getRange(nr,9).setNumberFormat("dd-mmm-yy HH:mm");
      if(nr%2===0)sheet.getRange(nr,1,1,9).setBackground("#DBEAFE");
      addMprTypeIfNew(r.mediaType,'');
      saved++;
    });
    SpreadsheetApp.flush();
    return {success:true,saved:saved,message:saved+" issue entries save!"};
  } catch(e){ return {success:false,message:"Error: "+e.message}; }
  finally{ try{lock.releaseLock();}catch(e2){} }
}

// ── OPENING — LOCK ──
function saveMprOpening(payload) {
  var lock=LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var ss=SpreadsheetApp.openById(MPR_SS_ID);
    var sheet=ss.getSheetByName("MPR_OPENING");
    if(!sheet){
      sheet=ss.insertSheet("MPR_OPENING");
      var h=["Sr.No","As On Date","Media Type","Container","Opening Qty","Added By","Entry Time"];
      sheet.appendRow(h);
      sheet.getRange(1,1,1,h.length).setBackground("#6D28D9").setFontColor("#fff").setFontWeight("bold").setFontSize(10);
      sheet.setFrozenRows(1); sheet.setColumnWidths(1,h.length,110); SpreadsheetApp.flush();
    }
    var date=payload.date?new Date(payload.date):new Date();
    var saved=0, entryTime=new Date();
    payload.rows.forEach(function(r){
      if(!r.mediaType||n(r.qty)<=0)return;
      var sr=sheet.getLastRow();
      sheet.appendRow([sr,date,r.mediaType,(r.container||''),n(r.qty),(payload.addedBy||''),entryTime]);
      var nr=sheet.getLastRow();
      sheet.getRange(nr,2).setNumberFormat("dd-mmm-yy");
      sheet.getRange(nr,7).setNumberFormat("dd-mmm-yy HH:mm");
      if(nr%2===0)sheet.getRange(nr,1,1,7).setBackground("#EDE9FE");
      addMprTypeIfNew(r.mediaType,'');
      saved++;
    });
    SpreadsheetApp.flush();
    return {success:true,saved:saved,message:saved+" opening entries save!"};
  } catch(e){ return {success:false,message:"Error: "+e.message}; }
  finally{ try{lock.releaseLock();}catch(e2){} }
}

// ── LOG (Production / Issue) ──
function getMprLog(type, from, to) {
  try {
    var ss=SpreadsheetApp.openById(MPR_SS_ID), tz=Session.getScriptTimeZone();
    var fromD=new Date(from);fromD.setHours(0,0,0,0);
    var toD=new Date(to);toD.setHours(23,59,59,999);
    var fromStr=Utilities.formatDate(fromD,tz,"yyyy-MM-dd"), toStr=Utilities.formatDate(toD,tz,"yyyy-MM-dd");
    var sheet=ss.getSheetByName(type==='issue'?"MPR_ISSUE":"MPR_PRODUCTION");
    if(!sheet||sheet.getLastRow()<=1)return {success:true,type:type,rows:[]};
    var data=sheet.getDataRange().getValues(), out=[];
    for(var i=1;i<data.length;i++){
      var cell=data[i][1];if(!cell)continue;
      var rd=cell instanceof Date?new Date(cell.getTime()):new Date(cell);
      if(isNaN(rd.getTime()))continue;
      var rdStr=Utilities.formatDate(rd,tz,"yyyy-MM-dd");
      if(rdStr<fromStr||rdStr>toStr)continue;
      if(type==='issue'){
        out.push({date:Utilities.formatDate(rd,tz,"dd-MMM-yy"),mediaType:String(data[i][2]||''),container:String(data[i][3]||''),qty:n(data[i][4]),department:String(data[i][5]||''),issuedTo:String(data[i][6]||''),addedBy:String(data[i][7]||'')});
      } else {
        out.push({date:Utilities.formatDate(rd,tz,"dd-MMM-yy"),shift:String(data[i][2]||''),mediaType:String(data[i][3]||''),batchNo:String(data[i][4]||''),container:String(data[i][5]||''),filled:n(data[i][6]),acTemp:String(data[i][7]||''),acTime:String(data[i][8]||''),acDone:String(data[i][9]||''),operator:String(data[i][10]||''),opCode:String(data[i][11]||''),supervisor:String(data[i][12]||'')});
      }
    }
    return {success:true,type:type,rows:out};
  } catch(e){ return {success:false,rows:[],message:e.message}; }
}

// ── STOCK REPORT (per Media Type + Container) ──
function getMprStockReport(params) {
  try {
    var tz=Session.getScriptTimeZone();
    var toD=new Date(params.dateTo);toD.setHours(23,59,59,999);
    var toStr=Utilities.formatDate(toD,tz,"yyyy-MM-dd");
    var fType=(params.mediaType||'').toLowerCase(), fCont=(params.container||'');
    var stock={};
    function key(t,c){return t+'|'+c;}
    function ensure(t,c){var k=key(t,c);if(!stock[k])stock[k]={mediaType:t,container:c,opening:0,produced:0,issued:0};return stock[k];}
    function dateOk(cell){if(!cell)return false;var rd=cell instanceof Date?new Date(cell.getTime()):new Date(cell);if(isNaN(rd.getTime()))return false;return Utilities.formatDate(rd,tz,"yyyy-MM-dd")<=toStr;}
    var ss=SpreadsheetApp.openById(MPR_SS_ID);
    var op=ss.getSheetByName("MPR_OPENING");
    if(op&&op.getLastRow()>1){var od=op.getDataRange().getValues();for(var i=1;i<od.length;i++){if(!dateOk(od[i][1]))continue;var t=String(od[i][2]||'').trim();if(!t)continue;ensure(t,String(od[i][3]||'').trim()).opening+=(parseFloat(od[i][4])||0);}}
    var pr=ss.getSheetByName("MPR_PRODUCTION");
    if(pr&&pr.getLastRow()>1){var pd=pr.getDataRange().getValues();for(var j=1;j<pd.length;j++){if(!dateOk(pd[j][1]))continue;var t2=String(pd[j][3]||'').trim();if(!t2)continue;ensure(t2,String(pd[j][5]||'').trim()).produced+=(parseFloat(pd[j][6])||0);}}
    var iss=ss.getSheetByName("MPR_ISSUE");
    if(iss&&iss.getLastRow()>1){var idd=iss.getDataRange().getValues();for(var k=1;k<idd.length;k++){if(!dateOk(idd[k][1]))continue;var t3=String(idd[k][2]||'').trim();if(!t3)continue;ensure(t3,String(idd[k][3]||'').trim()).issued+=(parseFloat(idd[k][4])||0);}}
    var rows=[];
    Object.keys(stock).forEach(function(kk){
      var s=stock[kk];
      if(fType&&s.mediaType.toLowerCase().indexOf(fType)<0)return;
      if(fCont&&s.container!==fCont)return;
      s.available=s.opening+s.produced-s.issued;
      if(s.opening===0&&s.produced===0&&s.issued===0)return;
      rows.push(s);
    });
    rows.sort(function(a,b){return (a.mediaType+a.container).localeCompare(b.mediaType+b.container);});
    return {success:true,rows:rows};
  } catch(e){ return {success:false,rows:[],message:e.message}; }
}

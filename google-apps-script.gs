/**
 * 根號拾｜彌月試算系統 × Google Sheet 同步後端（v2 通用寫入版）
 * 部署方式見 README.md「雲端同步設定」段落。
 *
 * Sheet 結構（程式會自動建立，不用手動開分頁）：
 *  - 「訂單總覽」：一位客人一列（名稱/更新時間/總數/總額/訂金/總尾款/進度 + 完整資料JSON）
 *  - 「取貨排程」：全部客人的批次彙總，一批一列
 *  - 「<客人名稱>」：一位媽咪一個分頁，格式仿原本 Excel（批次表＋訂購總數＋付款狀態＋話術全文）
 *    分頁內容由網頁端組好直接送過來（body.sheets），之後格式調整只改網頁、不用重貼此程式
 */

const TOKEN = 'genhouse2026';   // ← 同步密碼，可自行改，需與系統「⚙ 設定」內填的一致

const ORDER_HEADERS = ['客人名稱','最後更新','蛋糕總數','蛋糕總額','訂金','總尾款','已安排/總顆','完成批次','資料JSON(勿改)'];
const SCHED_HEADERS = ['客人名稱','方式','品項','蛋糕顆數','油飯盒數','收件人','電話','地址','到貨/取貨日','本島離島','箱數','完成'];

function doPost(e){
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);   // 兩個人同時上傳時排隊，避免互相蓋列
  try{
    const body = JSON.parse(e.postData.contents);
    if(body.token !== TOKEN) return json({error:'同步密碼(token)錯誤'});
    if(body.action === 'save')   return json(saveOrder(body));
    if(body.action === 'delete') return json(deleteOrder(body.name));
    return json({error:'未知動作'});
  }catch(err){
    return json({error:String(err)});
  }finally{
    lock.releaseLock();
  }
}

function doGet(e){
  try{
    const p = e.parameter || {};
    if(p.token !== TOKEN) return json({error:'同步密碼(token)錯誤'});
    if(p.action === 'list') return json(listOrders());
    if(p.action === 'load') return json(loadOrder(p.name));
    return json({error:'未知動作'});
  }catch(err){
    return json({error:String(err)});
  }
}

function json(obj){
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function ensureSheet(name, headers){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if(!sh){
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,headers.length).setFontWeight('bold');
  }
  return sh;
}
function sheetOrders(){ return ensureSheet('訂單總覽', ORDER_HEADERS); }
function sheetSched(){ return ensureSheet('取貨排程', SCHED_HEADERS); }

function findRow(sh, name){
  const last = sh.getLastRow();
  if(last < 2) return -1;
  const vals = sh.getRange(2,1,last-1,1).getValues();
  for(let i=0;i<vals.length;i++){
    if(String(vals[i][0]).trim() === name) return i+2;
  }
  return -1;
}

function deleteCustomerRows(sh, name){
  const last = sh.getLastRow();
  for(let r=last;r>=2;r--){
    if(String(sh.getRange(r,1).getValue()).trim() === name) sh.deleteRow(r);
  }
}

function saveOrder(body){
  const name = String(body.name||'').trim();
  if(!name) return {error:'缺少客人名稱'};
  const s = body.summary || {};
  const sh = sheetOrders();
  const row = [
    name, new Date(),
    s.cakeQty||0, s.cakeAmt||0, s.deposit||0, s.finalBalance||0,
    (s.scheduled||0) + '/' + ((s.cakeQty||0)+(s.giftQty||0)),
    (s.done||0) + '/' + (s.shipCnt||0),
    JSON.stringify(body.state||{}),
  ];
  const idx = findRow(sh, name);
  if(idx > 0) sh.getRange(idx,1,1,row.length).setValues([row]);
  else sh.appendRow(row);

  // 取貨排程：先清掉這位客人舊的列，再照目前排程重寫
  const sch = sheetSched();
  deleteCustomerRows(sch, name);
  const state = body.state || {};
  const P = state.params || {};
  (state.ships||[]).forEach(function(sp){
    const cake = +sp.cake||0, oil = +sp.oil||0;
    if(cake<=0 && oil<=0) return;
    const isDelivery = (sp.mode||'delivery')==='delivery';
    const boxes = isDelivery
      ? Math.ceil(cake/(P.boxCapCake||12)) + Math.ceil(oil/(P.boxCapOil||24))
      : '';
    sch.appendRow([
      name,
      isDelivery ? '宅配' : '自取',
      sp.label||'', cake, oil,
      sp.receiver||'',
      sp.phone ? "'"+sp.phone : '',   // 前置 ' 保住開頭的 0
      sp.addr||'', sp.date||'',
      isDelivery ? (sp.island?'離島':'本島') : '',
      boxes,
      sp.done ? 'V' : '',
    ]);
  });

  // 仿 Excel 的「一位媽咪一個分頁」：網頁端把資料＋格式指令(meta)組好送來，這裡整頁重寫
  (body.sheets||[]).forEach(function(spec){
    if(!spec || !spec.name || !spec.rows || !spec.rows.length) return;
    writeCustomSheet(safeSheetName(spec.name), spec.rows, spec.meta||{});
  });
  return {ok:true};
}

function safeSheetName(name){
  // Google Sheet 分頁名稱不能含 []/\*?: 且上限 100 字
  return String(name).replace(/[\[\]\/\\\*\?:]/g,'').slice(0,80) || '未命名';
}

function colIdx(letter){ return String(letter).toUpperCase().charCodeAt(0) - 64; }

function writeCustomSheet(name, rows, meta){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if(!sh) sh = ss.insertSheet(name);

  // 1) 先讀舊表，把手動勾的 KEY單/已完成 依「品項|數量|收件人」記下來
  const preserveMap = {};
  const pv = meta.preserve;
  if(pv && sh.getLastRow() > 0){
    const old = sh.getDataRange().getValues();
    const kc = colIdx(pv.keyCol)-1, dc = colIdx(pv.doneCol)-1;
    old.forEach(function(r){
      if(typeof r[kc] === 'boolean' || typeof r[dc] === 'boolean'){
        const key = pv.idCols.map(function(c){ return String(r[colIdx(c)-1]); }).join('|');
        preserveMap[key] = {key: r[kc] === true, done: r[dc] === true};
      }
    });
  }

  // 2) 整頁清空（含格式與下拉/核取方塊）
  sh.getRange(1,1,sh.getMaxRows(),sh.getMaxColumns()).clearDataValidations();
  sh.clear();

  // 3) 寫入資料
  const maxCols = rows.reduce(function(m,r){ return Math.max(m, (r||[]).length); }, 1);
  const grid = rows.map(function(r){
    const row = (r||[]).slice();
    while(row.length < maxCols) row.push('');
    return row;
  });
  sh.getRange(1,1,grid.length,maxCols).setValues(grid);
  sh.setFrozenRows(1);

  // 4) 欄寬
  if(meta.colWidths){
    for(var k in meta.colWidths){ sh.setColumnWidth(colIdx(k), meta.colWidths[k]); }
  }
  // 5) 樣式（粗體/顏色/置中/框線）
  (meta.styles||[]).forEach(function(st){
    const rng = sh.getRange(st.range);
    if(st.bold) rng.setFontWeight('bold');
    if(st.color) rng.setFontColor(st.color);
    if(st.align) rng.setHorizontalAlignment(st.align);
    if(st.border) rng.setBorder(true,true,true,true,true,true);
  });
  // 6) 下拉選單／日期驗證
  (meta.validations||[]).forEach(function(v){
    const rng = sh.getRange(v.range);
    if(v.list){
      rng.setDataValidation(SpreadsheetApp.newDataValidation()
        .requireValueInList(v.list, true).setAllowInvalid(true).build());
    }else if(v.date){
      rng.setDataValidation(SpreadsheetApp.newDataValidation()
        .requireDate().setAllowInvalid(true).build());
    }
  });
  // 7) 到貨日：轉成真正的日期值（雙擊出小日曆）＋ M/d 顯示
  (meta.dates||[]).forEach(function(d){
    const p = String(d.iso).split('-');
    if(p.length !== 3) return;
    const rng = sh.getRange(d.a1);
    rng.setValue(new Date(+p[0], +p[1]-1, +p[2]));
    rng.setNumberFormat('M/d');
  });
  // 8) KEY單/已完成 核取方塊
  (meta.checkboxes||[]).forEach(function(cb){
    const rng = sh.getRange(cb.a1);
    rng.insertCheckboxes();
    if(cb.value === true) rng.setValue(true);
  });
  // 9) 套回舊表手動勾選（同一批次以 Sheet 上的勾選為準）
  if(pv){
    (pv.rows||[]).forEach(function(r){
      const key = pv.idCols.map(function(c){ return String(grid[r-1][colIdx(c)-1]); }).join('|');
      const oldVal = preserveMap[key];
      if(!oldVal) return;
      sh.getRange(pv.keyCol + r).setValue(oldVal.key);
      sh.getRange(pv.doneCol + r).setValue(oldVal.done);
    });
  }
}

function listOrders(){
  const sh = sheetOrders();
  const last = sh.getLastRow();
  if(last < 2) return {orders:[]};
  const vals = sh.getRange(2,1,last-1,2).getValues();
  return {orders: vals
    .filter(function(v){ return String(v[0]).trim(); })
    .map(function(v){ return {name:String(v[0]).trim(), updated:v[1]}; })};
}

function loadOrder(name){
  const sh = sheetOrders();
  const idx = findRow(sh, String(name||'').trim());
  if(idx < 0) return {error:'雲端找不到這筆訂單'};
  const jsonStr = sh.getRange(idx, ORDER_HEADERS.length).getValue();
  try{
    return {ok:true, state: JSON.parse(jsonStr)};
  }catch(e){
    return {error:'這筆訂單的 JSON 欄位被改壞了，請勿手動編輯最後一欄'};
  }
}

function deleteOrder(name){
  const n = String(name||'').trim();
  const sh = sheetOrders();
  const idx = findRow(sh, n);
  if(idx > 0) sh.deleteRow(idx);
  deleteCustomerRows(sheetSched(), n);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const custSheet = ss.getSheetByName(safeSheetName(n));
  if(custSheet) ss.deleteSheet(custSheet);
  return {ok:true};
}

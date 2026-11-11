/** Home & Market Manager Nepal — Google Apps Script backend
 * Run setupProject() once from the Apps Script editor and approve permissions.
 * Deploy as Web app. This is a starter backend, not a regulated-finance platform.
 */
const TAB_HEADERS = {
  customers:['id','date','name','phone','email','address'],
  rooms:['id','date','room','rent','customer','status'],
  roomRent:['id','date','tenant','month','rentamt','duedate','status'],
  products:['id','date','pname','price','stock','unit','photo'],
  orders:['id','date','buyer','items','total','paymode','status'],
  payments:['id','date','payref','paycustomer','payamount','paymethod','paystatus'],
  savings:['id','date','scustomer','samount','snote','status'],
  withdrawals:['id','date','wcustomer','wamount','wnote','status'],
  loans:['id','date','lcustomer','principal','rate','months','method','interestMethod','lstatus','outstanding'],
  loanRepayments:['id','loanid','lrpayer','lramount','lrdate','interest','principal'],
  transactions:['id','date','type','reference','amount','note'],
  users:['id','username','password','role','useremail'],
  notifications:['id','date','to','subject','message','status'],
  settings:['key','value']
};
function setupProject(){
  const props=PropertiesService.getScriptProperties();
  let id=props.getProperty('SPREADSHEET_ID');
  let ss=id?SpreadsheetApp.openById(id):SpreadsheetApp.create('Home & Market Manager Nepal Database');
  props.setProperty('SPREADSHEET_ID',ss.getId());
  Object.keys(TAB_HEADERS).forEach(name=>{
    let sh=ss.getSheetByName(name)||ss.insertSheet(name);
    const headers=TAB_HEADERS[name];
    if(sh.getLastRow()===0) sh.getRange(1,1,1,headers.length).setValues([headers]);
    else {
      const existing=sh.getRange(1,1,1,Math.max(sh.getLastColumn(),headers.length)).getValues()[0].filter(String);
      headers.forEach(h=>{if(existing.indexOf(h)<0){sh.getRange(1,sh.getLastColumn()+1).setValue(h);}});
    }
    sh.setFrozenRows(1);
  });
  let settings=ss.getSheetByName('settings');
  if(settings.getLastRow()===1) settings.appendRow(['spreadsheetId',ss.getId()]);
  Logger.log('Database created: '+ss.getUrl());
  return {status:'OK',spreadsheetId:ss.getId(),url:ss.getUrl(),tabs:Object.keys(TAB_HEADERS)};
}
function doGet(e){
  try{
    const name=(e&&e.parameter&&e.parameter.sheet)||'';
    if(!TAB_HEADERS[name]) return json_({status:'ERROR',message:'Invalid sheet'});
    const sh=sheet_(name), values=sh.getDataRange().getValues();
    if(values.length<2) return json_({status:'OK',rows:[]});
    const headers=values.shift();
    const rows=values.filter(r=>r.some(v=>v!==''&&v!==null)).map(r=>{let o={};headers.forEach((h,i)=>o[h]=r[i]);return o;});
    return json_({status:'OK',rows:rows});
  }catch(err){return json_({status:'ERROR',message:String(err)});}
}
function doPost(e){
  const lock=LockService.getScriptLock();
  try{
    lock.waitLock(20000);
    const body=JSON.parse((e.postData&&e.postData.contents)||'{}');
    const action=body.action, name=body.sheet;
    if(action==='ping') return json_({status:'OK',message:'Backend ready'});
    if(!TAB_HEADERS[name]) return json_({status:'ERROR',message:'Invalid sheet'});
    const sh=sheet_(name), headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
    if(action==='upsert'){
      const row=body.row||{}; if(!row.id) row.id=Utilities.getUuid();
      const values=headers.map(h=>row[h]===undefined?'':row[h]);
      const idCol=headers.indexOf('id')+1;
      if(idCol){
        const last=sh.getLastRow();
        if(last>1){const ids=sh.getRange(2,idCol,last-1,1).getValues().flat();const idx=ids.findIndex(x=>String(x)===String(row.id));if(idx>=0){sh.getRange(idx+2,1,1,headers.length).setValues([values]);return json_({status:'OK',action:'updated',id:row.id});}}
      }
      sh.appendRow(values);
      if(name==='notifications') sendQueuedEmail_(row);
      return json_({status:'OK',action:'inserted',id:row.id});
    }
    if(action==='delete'){
      const id=String((body.row||{}).id||'');const idCol=headers.indexOf('id')+1;if(!id||!idCol)throw new Error('Missing row id');
      const last=sh.getLastRow();if(last>1){const ids=sh.getRange(2,idCol,last-1,1).getValues().flat();const idx=ids.findIndex(x=>String(x)===id);if(idx>=0){sh.deleteRow(idx+2);return json_({status:'OK',action:'deleted',id:id});}}
      return json_({status:'OK',action:'not_found',id:id});
    }
    return json_({status:'ERROR',message:'Unsupported action'});
  }catch(err){return json_({status:'ERROR',message:String(err)});}
  finally{try{lock.releaseLock();}catch(ignore){}}
}
function sendQueuedEmail_(row){
  const to=String(row.to||'').trim();if(!to||!row.subject)return;
  try{MailApp.sendEmail({to:to,subject:String(row.subject),body:String(row.message||'')});}
  catch(err){console.error('Email send failed: '+err);}
}
function sheet_(name){
  if(!TAB_HEADERS[name])throw new Error('Invalid sheet');
  const id=PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if(!id)throw new Error('Run setupProject() first');
  const ss=SpreadsheetApp.openById(id);let sh=ss.getSheetByName(name);
  if(!sh){sh=ss.insertSheet(name);sh.appendRow(TAB_HEADERS[name]);}
  return sh;
}
function json_(obj){return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);}
function testEmail(){MailApp.sendEmail(Session.getActiveUser().getEmail(),'Home & Market Manager test','Gmail notification test succeeded.');}
function getDatabaseUrl(){const id=PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');return id?SpreadsheetApp.openById(id).getUrl():'Run setupProject() first';}

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
  Logger.log('Database ready: '+ss.getUrl());
  return {status:'OK',message:'Spreadsheet and all tabs are ready',spreadsheetId:ss.getId(),url:ss.getUrl(),tabs:Object.keys(TAB_HEADERS)};
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
    // First setup call creates the spreadsheet and every required tab automatically.
    if(action==='setup' || action==='setupProject') return json_(setupProject());
    if(action==='ping') {
      const id=PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
      const database=id ? {status:'OK',spreadsheetId:id,url:SpreadsheetApp.openById(id).getUrl(),tabs:Object.keys(TAB_HEADERS)} : setupProject();
      return json_({status:'OK',message:'Backend ready',database:database});
    }
    if(action==='requestPasswordReset') return requestPasswordReset_(body);
    if(action==='confirmPasswordReset') return confirmPasswordReset_(body);
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

/**
 * Email OTP password reset.
 * Deploy the Apps Script as a Web App and authorize MailApp.
 * OTP is stored in Script Cache for 10 minutes and is never returned to the browser.
 */
function requestPasswordReset_(body){
  const username=String(body.username||'').trim();
  const email=String(body.email||'').trim().toLowerCase();
  if(!username||!email) return json_({status:'ERROR',message:'Username and registered email are required'});
  const sh=sheet_('users'), values=sh.getDataRange().getValues();
  if(values.length<2) return json_({status:'OK',message:'If the account matches, an OTP has been sent'});
  const headers=values[0].map(String);
  const u=headers.indexOf('username'), em=headers.indexOf('useremail'), id=headers.indexOf('id');
  if(u<0||em<0) return json_({status:'ERROR',message:'Users sheet needs username and useremail columns'});
  const match=values.slice(1).findIndex(r=>String(r[u]).trim()===username && String(r[em]).trim().toLowerCase()===email);
  // Use a generic success response to avoid revealing whether an account exists.
  if(match<0) return json_({status:'OK',message:'If the account matches, an OTP has been sent'});
  const cache=CacheService.getScriptCache();
  const throttleKey='reset_throttle_'+username.toLowerCase();
  if(cache.get(throttleKey)) return json_({status:'ERROR',message:'Please wait before requesting another OTP'});
  const otp=String(Math.floor(100000+Math.random()*900000));
  const tokenKey='reset_otp_'+username.toLowerCase();
  const payload={otp:otp,email:email,row:match+2,attempts:0,created:Date.now()};
  cache.put(tokenKey,JSON.stringify(payload),600);
  cache.put(throttleKey,'1',60);
  MailApp.sendEmail({
    to:email,
    subject:'Home & Market Manager Nepal - Password Reset OTP',
    body:'Your password reset OTP is: '+otp+'\n\nThis OTP expires in 10 minutes. If you did not request this, ignore this email.'
  });
  return json_({status:'OK',message:'If the account matches, an OTP has been sent'});
}
function confirmPasswordReset_(body){
  const username=String(body.username||'').trim();
  const email=String(body.email||'').trim().toLowerCase();
  const otp=String(body.otp||'').trim();
  const newPassword=String(body.newPassword||'');
  if(!username||!email||!otp||newPassword.length<8)
    return json_({status:'ERROR',message:'Invalid reset details; password must be at least 8 characters'});
  const cache=CacheService.getScriptCache();
  const key='reset_otp_'+username.toLowerCase();
  const raw=cache.get(key);
  if(!raw) return json_({status:'ERROR',message:'OTP expired or not found. Request a new OTP.'});
  const data=JSON.parse(raw);
  if(data.email!==email) return json_({status:'ERROR',message:'OTP verification failed'});
  data.attempts=(data.attempts||0)+1;
  if(data.attempts>5){cache.remove(key);return json_({status:'ERROR',message:'Too many attempts. Request a new OTP.'});}
  if(data.otp!==otp){cache.put(key,JSON.stringify(data),600);return json_({status:'ERROR',message:'Incorrect OTP'});}
  const sh=sheet_('users');
  const headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(String);
  const userCol=headers.indexOf('username')+1, emailCol=headers.indexOf('useremail')+1, passCol=headers.indexOf('password')+1;
  if(!userCol||!emailCol||!passCol) return json_({status:'ERROR',message:'Users sheet columns are incomplete'});
  const rows=sh.getDataRange().getValues();
  let foundRow=0;
  for(let i=1;i<rows.length;i++){
    if(String(rows[i][userCol-1]).trim()===username &&
       String(rows[i][emailCol-1]).trim().toLowerCase()===email){foundRow=i+1;break;}
  }
  if(!foundRow) return json_({status:'ERROR',message:'Account not found'});
  sh.getRange(foundRow,passCol).setValue(newPassword);
  cache.remove(key);
  return json_({status:'OK',message:'Password updated'});
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

/** Home & Market Manager Nepal — Google Apps Script backend
 * Run setupProject() once from the Apps Script editor and approve permissions.
 * Deploy as Web app. This is a starter backend, not a regulated-finance platform.
 */
const TAB_HEADERS = {
  customers:['id','date','name','phone','email','address','photo','ownerId','ownerUsername'],
  rooms:['id','date','room','rent','customer','status','photo','approvalStatus','rejectionReason','ownerId','ownerUsername'],
  roomRent:['id','date','tenant','month','rentamt','duedate','status','photo','approvalStatus','rejectionReason','ownerId','ownerUsername'],
  products:['id','date','pname','price','stock','unit','photo'],
  orders:['id','date','buyer','items','cartItems','total','paymode','status','approvalStatus','rejectionReason','ownerId','ownerUsername'],
  payments:['id','date','payref','paycustomer','payamount','paymethod','paystatus','ownerId','ownerUsername'],
  savings:['id','date','scustomer','samount','snote','savingType','status','ownerId','ownerUsername'],
  transfers:['id','date','fromCustomerId','fromCustomer','toCustomerId','toCustomer','amount','note','status','ownerId','ownerUsername'],
  withdrawals:['id','date','wcustomer','wamount','wnote','status','ownerId','ownerUsername'],
  loans:['id','date','lcustomer','principal','rate','months','method','interestMethod','lstatus','outstanding','status','approvalStatus','rejectionReason','ownerId','ownerUsername'],
  loanRepayments:['id','loanid','lrpayer','lramount','lrdate','interest','principal','ownerId','ownerUsername'],
  transactions:['id','date','type','reference','amount','note','ownerId','ownerUsername'],
  users:['id','username','password','role','useremail','photo','status','approvalStatus','rejectionReason','customerId'],
  notifications:['id','date','to','subject','message','status'],
  approvals:['id','date','type','sheet','recordid','customer','amount','status','reason','token','email','ownerId','ownerUsername','stage','assignedRole','history'],
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
  const users=ss.getSheetByName('users');
  const userHeaders=users.getRange(1,1,1,users.getLastColumn()).getValues()[0].map(String);
  const uCol=userHeaders.indexOf('username');
  const hasAdmin=users.getLastRow()>1 && users.getRange(2,uCol+1,users.getLastRow()-1,1).getValues().some(r=>String(r[0]).trim().toLowerCase()==='admin');
  if(!hasAdmin){const admin={id:'admin',username:'admin',password:'ChangeMe123!',role:'Admin',useremail:'',status:'Approved',approvalStatus:'Approved',rejectionReason:'',customerId:''};users.appendRow(userHeaders.map(h=>admin[h]===undefined?'':admin[h]));}
  let settings=ss.getSheetByName('settings');
  if(settings.getLastRow()===1) settings.appendRow(['spreadsheetId',ss.getId()]);
  Logger.log('Database ready: '+ss.getUrl());
  return {status:'OK',message:'Spreadsheet and all tabs are ready',spreadsheetId:ss.getId(),url:ss.getUrl(),tabs:Object.keys(TAB_HEADERS)};
}
function doGet(e){
  try{
    if(e&&e.parameter&&e.parameter.action==='approval') return handleApprovalLink_(e.parameter);
    if(e&&e.parameter&&e.parameter.action==='sync'){const session=authSession_(e.parameter.token);if(!session)return json_({status:'ERROR',message:'Login expired'});return json_({status:'OK',tables:readAllTables_(session),serverTime:new Date().toISOString()});}
    const name=(e&&e.parameter&&e.parameter.sheet)||'';
    const session=authSession_(e&&e.parameter&&e.parameter.token);if(!session)return json_({status:'ERROR',message:'Login expired'});
    if(!TAB_HEADERS[name]) return json_({status:'ERROR',message:'Invalid sheet'});
    const sh=sheet_(name), values=sh.getDataRange().getValues();
    if(values.length<2) return json_({status:'OK',rows:[]});
    const headers=values.shift();
    let rows=values.filter(r=>r.some(v=>v!==''&&v!==null)).map(r=>{let o={};headers.forEach((h,i)=>o[h]=r[i]);return o;});if(session.role==='Customer')rows=rows.filter(r=>isOwnRow_(name,r,session));if(name==='settings'){rows=session.role==='Admin'?rows.map(r=>String(r.key).toLowerCase()==='imgbbapikey'?{key:r.key,value:'••••••••'}:r):[];}
    return json_({status:'OK',rows:rows});
  }catch(err){return json_({status:'ERROR',message:String(err)});}
}
function readAllTables_(session){
  const out={}; Object.keys(TAB_HEADERS).forEach(name=>{const sh=sheet_(name), values=sh.getDataRange().getValues(); if(values.length<2){out[name]=[];return;} const headers=values.shift().map(String); let rows=values.filter(r=>r.some(v=>v!==''&&v!==null)).map(r=>{const o={};headers.forEach((h,i)=>o[h]=r[i]);return o;}); if(session.role==='Customer'){rows=rows.filter(r=>isOwnRow_(name,r,session));} else if(name==='users'&&session.role!=='Admin'){rows=rows.filter(r=>String(r.id)===String(session.id));} if(name==='approvals'&&session.role!=='Admin'&&session.role!=='Customer'){rows=rows.filter(r=>String(r.assignedRole||'')===String(session.role));} if(name==='settings'){rows=session.role==='Admin'?rows.map(r=>String(r.key).toLowerCase()==='imgbbapikey'?{key:r.key,value:'••••••••'}:r):[];} out[name]=rows;}); return out;
}
function authSession_(token){if(!token)return null;const raw=CacheService.getScriptCache().get('auth_'+token);return raw?JSON.parse(raw):null;}
function isOwnRow_(sheet,row,s){if(sheet==='approvals')return String(row.ownerId||'')===String(s.customerId||s.id)||String(row.ownerUsername||'').toLowerCase()===String(s.username||'').toLowerCase();if(sheet==='transfers')return String(row.fromCustomerId||'')===String(s.customerId||s.id)||String(row.toCustomerId||'')===String(s.customerId||s.id); if(['products'].includes(sheet))return true;if(sheet==='settings')return false;if(sheet==='users')return String(row.id)===String(s.id);const owner=String(row.ownerId||'');if(owner)return owner===String(s.customerId||s.id);const keys={customers:['id','email'],rooms:['customer'],roomRent:['tenant'],orders:['buyer'],payments:['paycustomer'],savings:['scustomer'],withdrawals:['wcustomer'],loans:['lcustomer'],loanRepayments:['lrpayer'],transactions:['reference']};const vals=(keys[sheet]||[]).map(k=>String(row[k]||'').toLowerCase());return vals.includes(String(s.customerId||'').toLowerCase())||(s.email&&vals.includes(String(s.email).toLowerCase()));}
function authenticate_(body){if(!PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'))setupProject();const u=String(body.username||'').trim(),p=String(body.password||'');const sh=sheet_('users'),v=sh.getDataRange().getValues(),h=v[0].map(String),ui=h.indexOf('username'),pi=h.indexOf('password');for(let i=1;i<v.length;i++){if(String(v[i][ui]).trim().toLowerCase()===u.toLowerCase()&&String(v[i][pi])===p){const o={};h.forEach((k,j)=>o[k]=v[i][j]);const st=String(o.approvalStatus||o.status||'Approved').toLowerCase();if(['pending','pending approval','rejected','blocked','inactive'].includes(st))return {status:'ERROR',message:'Account awaiting approval or inactive'};let cust=null;if(String(o.role).toLowerCase()==='customer'){const cs=sheet_('customers').getDataRange().getValues(),ch=cs[0].map(String),ci=ch.indexOf('id'),cn=ch.indexOf('name'),ce=ch.indexOf('email');for(let j=1;j<cs.length;j++){if((o.customerId&&String(cs[j][ci])===String(o.customerId))||(o.useremail&&ce>=0&&String(cs[j][ce]).toLowerCase()===String(o.useremail).toLowerCase())||String(cs[j][cn]).toLowerCase()===u.toLowerCase()){cust={id:cs[j][ci],name:cs[j][cn],email:ce>=0?cs[j][ce]:''};break;}}if(!cust)return {status:'ERROR',message:'Customer user is not linked. Set customerId or match the user email with the customer email.'};o.customerId=cust.id;o.customerName=cust.name;o.email=cust.email;}const token=Utilities.getUuid()+Utilities.getUuid();const normalizedRole=String(o.role||'Customer').toLowerCase().replace(/[_-]/g,' ').trim();const role=normalizedRole==='customer'?'Customer':normalizedRole==='admin'?'Admin':normalizedRole==='manager'?'Manager':(['field staff','fieldstaff','staff'].includes(normalizedRole)?'Field Staff':o.role||'Customer');const session={id:o.id,username:o.username,role,customerId:o.customerId||o.id,customerName:o.customerName||'',email:o.useremail||''};CacheService.getScriptCache().put('auth_'+token,JSON.stringify(session),21600);return {status:'OK',token:token,user:session};}}return {status:'ERROR',message:'Invalid username or password'};}

function doPost(e){
  const lock=LockService.getScriptLock();
  try{
    lock.waitLock(20000);
    const body=JSON.parse((e.postData&&e.postData.contents)||'{}');
    const action=body.action, name=body.sheet;
    // First setup call creates the spreadsheet and every required tab automatically.
    if(action==='setup' || action==='setupProject') return json_(setupProject());
    if(action==='authLogin') return json_(authenticate_(body));
    if(action==='ping') {
      const id=PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
      const database=id ? {status:'OK',spreadsheetId:id,url:SpreadsheetApp.openById(id).getUrl(),tabs:Object.keys(TAB_HEADERS)} : setupProject();
      return json_({status:'OK',message:'Backend ready',database:database});
    }
    if(action==='saveSetting'){const ssn=authSession_(body.token);if(!ssn||ssn.role!=='Admin')return json_({status:'ERROR',message:'Only Admin can change settings'});const k=String(body.key||'').trim(),v=String(body.value||'').trim();if(!['imgbbapikey','mainadminemail'].includes(k.toLowerCase())||!v)return json_({status:'ERROR',message:'Invalid setting'});return saveSetting_(k,v);}
    if(action==='uploadImage'){const ssn=authSession_(body.token);if(!ssn)return json_({status:'ERROR',message:'Login expired'});return uploadImageToImgBB_(body);}
    if(action==='requestPasswordReset') return requestPasswordReset_(body);
    if(action==='requestApproval'){const ssn=authSession_(body.token);if(!ssn)return json_({status:'ERROR',message:'Login expired'});if(ssn.role==='Customer'&&String((body.row||{}).ownerId)!==String(ssn.customerId))return json_({status:'ERROR',message:'Permission denied'});return requestApproval_(body);}
    if(action==='reviewLoanRequest'){const ssn=authSession_(body.token);if(!ssn)return json_({status:'ERROR',message:'Login expired'});return reviewLoanRequest_(body,ssn);}
    if(action==='transferOptionalSavings'){const ssn=authSession_(body.token);if(!ssn||ssn.role!=='Customer')return json_({status:'ERROR',message:'Only customers can transfer their optional savings'});return transferOptionalSavings_(body,ssn);}
    if(action==='decideApproval'){const ssn=authSession_(body.token);if(!ssn||!['Admin','Manager'].includes(ssn.role))return json_({status:'ERROR',message:'Permission denied'});return decideApproval_(body.id,body.decision,body.reason||'');}
    if(action==='confirmPasswordReset') return confirmPasswordReset_(body);
    if(!TAB_HEADERS[name]) return json_({status:'ERROR',message:'Invalid sheet'});
    const session=authSession_(body.token);if(!session)return json_({status:'ERROR',message:'Login expired'});
    if(session.role==='Customer'&&action==='delete')return json_({status:'ERROR',message:'Customers cannot delete records'});
    if(action==='delete'&&['savings','withdrawals','loanRepayments','transactions','transfers','payments'].includes(name))return json_({status:'ERROR',message:'Financial ledger records cannot be deleted. Use a correcting entry instead.'});
    if(action==='upsert'&&['savings','withdrawals','loanRepayments','transactions','transfers','payments'].includes(name)){const rr=body.row||{};const oldSh=sheet_(name),oldVals=oldSh.getDataRange().getValues(),oldHeaders=oldVals[0].map(String),oldId=oldHeaders.indexOf('id');if(rr.id&&oldId>=0&&oldVals.slice(1).some(r=>String(r[oldId])===String(rr.id)))return json_({status:'ERROR',message:'Financial ledger entries are immutable. Add a correction entry instead.'});}
    if(session.role==='Customer'&&action==='upsert'){const rid=String((body.row||{}).id||'');if(rid){const sh0=sheet_(name), vals0=sh0.getDataRange().getValues(), hh0=vals0[0].map(String), ic0=hh0.indexOf('id');if(ic0>=0&&vals0.slice(1).some(r=>String(r[ic0])===rid))return json_({status:'ERROR',message:'Customers cannot edit existing records'});}if(['savings','withdrawals','loanRepayments','transactions','transfers','users','customers','products'].includes(name))return json_({status:'ERROR',message:'This financial or master record can only be managed by staff'});}
    if(session.role==='Customer' && !['customers','rooms','roomRent','orders','payments','savings','withdrawals','loans','loanRepayments','transactions','products'].includes(name))return json_({status:'ERROR',message:'Permission denied'});if(session.role==='Customer'&&name==='products')return json_({status:'ERROR',message:'Product catalogue is read-only for customers'});
    const sh=sheet_(name), headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
    if(action==='upsert'){
      const row=body.row||{}; if(!row.id) row.id=Utilities.getUuid(); if(session.role==='Customer'){row.ownerId=session.customerId;row.ownerUsername=session.username;if(name==='customers')return json_({status:'ERROR',message:'Customer cannot create customer records'});}
      const values=headers.map(h=>row[h]===undefined?'':row[h]);
      const idCol=headers.indexOf('id')+1;
      if(idCol){
        const last=sh.getLastRow();
        if(last>1){const ids=sh.getRange(2,idCol,last-1,1).getValues().flat();const idx=ids.findIndex(x=>String(x)===String(row.id));if(idx>=0){const old={};headers.forEach((h,j)=>old[h]=sh.getRange(idx+2,j+1).getValue());if(session.role==='Customer'&&!isOwnRow_(name,old,session))return json_({status:'ERROR',message:'Permission denied'});sh.getRange(idx+2,1,1,headers.length).setValues([values]);return json_({status:'OK',action:'updated',id:row.id});}}
      }
      sh.appendRow(values);
      if(name==='notifications') sendQueuedEmail_(row);
      return json_({status:'OK',action:'inserted',id:row.id});
    }
    if(action==='delete'){
      const id=String((body.row||{}).id||'');const idCol=headers.indexOf('id')+1;if(!id||!idCol)throw new Error('Missing row id');
      const last=sh.getLastRow();if(last>1){const ids=sh.getRange(2,idCol,last-1,1).getValues().flat();const idx=ids.findIndex(x=>String(x)===id);if(idx>=0){const old={};headers.forEach((h,j)=>old[h]=sh.getRange(idx+2,j+1).getValue());if(session.role==='Customer'&&!isOwnRow_(name,old,session))return json_({status:'ERROR',message:'Permission denied'});sh.deleteRow(idx+2);return json_({status:'OK',action:'deleted',id:id});}}
      return json_({status:'OK',action:'not_found',id:id});
    }
    return json_({status:'ERROR',message:'Unsupported action'});
  }catch(err){return json_({status:'ERROR',message:String(err)});}
  finally{try{lock.releaseLock();}catch(ignore){}}
}

function saveSetting_(key,value){
  const sh=sheet_('settings'), data=sh.getDataRange().getValues(), h=data[0].map(String), kc=h.indexOf('key'), vc=h.indexOf('value');
  if(kc<0||vc<0)return json_({status:'ERROR',message:'settings tab must have key and value columns'});
  for(let i=1;i<data.length;i++)if(String(data[i][kc]).trim().toLowerCase()===String(key).toLowerCase()){sh.getRange(i+1,vc+1).setValue(value);return json_({status:'OK',message:'Setting updated'});}
  sh.appendRow(h.map(x=>x==='key'?key:x==='value'?value:''));return json_({status:'OK',message:'Setting saved'});
}
function uploadImageToImgBB_(body){
  const image=String(body.image||'').replace(/^data:image\/[^;]+;base64,/,'');
  if(!image)return json_({status:'ERROR',message:'Image data missing'});
  if(image.length>7*1024*1024)return json_({status:'ERROR',message:'Image is too large. Please use an image under 5 MB.'});
  const sh=sheet_('settings'), values=sh.getDataRange().getValues(), h=values[0].map(String), kc=h.indexOf('key'), vc=h.indexOf('value');let apiKey='';
  for(let i=1;i<values.length;i++)if(String(values[i][kc]||'').trim().toLowerCase()==='imgbbapikey'){apiKey=String(values[i][vc]||'').trim();break;}
  if(!apiKey)return json_({status:'ERROR',message:'ImgBB API key is missing in Google Sheets settings tab (key=imgbbApiKey).'});
  try{
    const response=UrlFetchApp.fetch('https://api.imgbb.com/1/upload?key='+encodeURIComponent(apiKey),{method:'post',payload:{image:image,name:String(body.fileName||'home-market-photo').replace(/[^a-zA-Z0-9._-]/g,'_')},muteHttpExceptions:true});
    const result=JSON.parse(response.getContentText());
    if(response.getResponseCode()<200||response.getResponseCode()>=300||!result.success||!result.data||!result.data.url)throw new Error(result.error&&result.error.message?result.error.message:'ImgBB upload rejected');
    return json_({status:'OK',url:result.data.url,displayUrl:result.data.display_url||result.data.url,deleteUrl:result.data.delete_url||''});
  }catch(err){return json_({status:'ERROR',message:'ImgBB upload failed: '+String(err)});}
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
  try {
    MailApp.sendEmail({
      to:email,
      subject:'Home & Market Manager Nepal - Password Reset OTP',
      body:'Your password reset OTP is: '+otp+'\\n\\nThis OTP expires in 10 minutes. If you did not request this, ignore this email.'
    });
  } catch(mailErr) {
    cache.remove(tokenKey);
    cache.remove(throttleKey);
    return json_({status:'ERROR',message:'Gmail OTP पठाउन सकेन। Apps Script मा MailApp permission authorize गर्नुहोस्: '+String(mailErr)});
  }
  return json_({status:'OK',message:'OTP sent'});
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


function getMainAdminApprovalEmail_(){
  const sh=sheet_('settings');
  const values=sh.getDataRange().getValues();
  if(values.length<2)return '';
  const headers=values[0].map(String), keyCol=headers.indexOf('key'), valueCol=headers.indexOf('value');
  if(keyCol<0||valueCol<0)return '';
  for(let i=1;i<values.length;i++){
    if(String(values[i][keyCol]||'').trim().toLowerCase()==='mainadminemail')
      return String(values[i][valueCol]||'').trim();
  }
  return '';
}

function requestApproval_(body){
  const sheetName=String(body.sheet||''), row=body.row||{};
  const allowed=['loans','users','orders','rooms','roomRent'];
  if(allowed.indexOf(sheetName)<0||!row.id) return json_({status:'ERROR',message:'Invalid approval request'});
  const sh=sheet_('approvals'), existing=sh.getDataRange().getValues(), h=existing[0].map(String);
  const idCol=h.indexOf('recordid'), statusCol=h.indexOf('status');
  if(existing.slice(1).some(r=>String(r[idCol])===String(row.id)&&['pending','returned'].includes(String(r[statusCol]).toLowerCase()))) return json_({status:'OK',message:'Approval already active'});
  const token=Utilities.getUuid()+Utilities.getUuid();
  const type=({loans:'Loan',users:'User',orders:'Cart / Order',rooms:'Room',roomRent:'Room Rent'})[sheetName];
  const customer=String(row.lcustomer||row.username||row.buyer||row.room||row.tenant||'');
  const amount=Number(row.principal||row.total||row.rentamt||0);
  const isLoan=sheetName==='loans';
  const stage=isLoan?'Field Staff':'Admin';
  const adminEmail=getMainAdminApprovalEmail_();
  const approval={id:Utilities.getUuid(),date:new Date(),type,sheet:sheetName,recordid:String(row.id),customer,amount,status:'Pending',reason:'',token,email:adminEmail,ownerId:String(row.ownerId||row.customerId||row.id),ownerUsername:String(row.ownerUsername||row.username||''),stage,assignedRole:stage,history:JSON.stringify([{at:new Date().toISOString(),by:row.ownerUsername||'Customer',action:'Submitted',to:stage}])};
  sh.appendRow(h.map(k=>approval[k]===undefined?'':approval[k]));
  if(isLoan){notifyRole_(stage,'New loan request from '+customer+' (NPR '+amount+')','Review loan request ID: '+row.id);return json_({status:'OK',message:'Loan sent to Field Staff for review',approvalId:approval.id});}
  if(!adminEmail||adminEmail.indexOf('@')<1)return json_({status:'ERROR',message:'Main Admin Gmail missing in settings tab (mainAdminEmail)'});
  const base=ScriptApp.getService().getUrl(), yes=base+'?action=approval&token='+encodeURIComponent(token)+'&decision=Approved', no=base+'?action=approval&token='+encodeURIComponent(token)+'&decision=Rejected';
  const html='<div style="font-family:Arial,sans-serif"><h2>Home & Market Manager Nepal</h2><p>'+type+' request: '+customer+' · NPR '+amount+'</p><p><a href="'+yes+'">YES — APPROVE</a> | <a href="'+no+'">NO — REJECT</a></p></div>';
  MailApp.sendEmail({to:adminEmail,subject:'Approval required: '+type+' - '+customer,body:'Approve: '+yes+'\nReject: '+no,htmlBody:html});
  return json_({status:'OK',message:'Approval email sent',approvalId:approval.id});
}
function notifyRole_(role,subject,message){
  try{const v=sheet_('users').getDataRange().getValues(),h=v[0].map(String),rc=h.indexOf('role'),ec=h.indexOf('useremail'),sc=h.indexOf('status');for(let i=1;i<v.length;i++){const r=String(v[i][rc]||'').toLowerCase().replace(/[_-]/g,' ').trim();const want=String(role).toLowerCase();if((r===want|| (want==='field staff'&&['staff','fieldstaff','field staff'].includes(r)))&&String(v[i][ec]||'').includes('@')&&(!String(v[i][sc]||'').toLowerCase()||String(v[i][sc]).toLowerCase()==='approved'))MailApp.sendEmail(String(v[i][ec]),subject,message);}}catch(e){console.error(e);}
}
function reviewLoanRequest_(body,session){
  const decision=String(body.decision||''), reason=String(body.reason||'').trim();if(!['Approved','Rejected'].includes(decision))return json_({status:'ERROR',message:'Invalid decision'});
  const sh=sheet_('approvals'),v=sh.getDataRange().getValues(),h=v[0].map(String),idc=h.indexOf('id');let i=v.findIndex((r,n)=>n>0&&String(r[idc])===String(body.id));if(i<1)return json_({status:'ERROR',message:'Request not found'});
  const a={};h.forEach((k,j)=>a[k]=v[i][j]);if(a.sheet!=='loans')return json_({status:'ERROR',message:'Not a loan request'});if(String(a.assignedRole)!==String(session.role))return json_({status:'ERROR',message:'यो request तपाईंको approval stage मा छैन'});if(String(a.status).toLowerCase()!=='pending'&&String(a.status).toLowerCase()!=='returned')return json_({status:'ERROR',message:'Request is not active'});if(decision==='Rejected'&&!reason)return json_({status:'ERROR',message:'Rejection reason is required'});
  const chain=['Field Staff','Manager','Admin'], pos=chain.indexOf(session.role);if(pos<0)return json_({status:'ERROR',message:'Invalid reviewer role'});
  let nextStatus,nextRole='',final=false;
  if(decision==='Approved'){if(pos===2){nextStatus='Approved';final=true;}else{nextRole=chain[pos+1];nextStatus='Pending';}}
  else if(pos===0){nextStatus='Rejected';final=true;}
  else{nextRole=chain[pos-1];nextStatus='Returned';}
  const history=JSON.parse(String(a.history||'[]'));history.push({at:new Date().toISOString(),by:session.username,role:session.role,action:decision,reason,to:nextRole|| (final?nextStatus:'')});
  const set=(col,val)=>{const c=h.indexOf(col);if(c>=0)sh.getRange(i+1,c+1).setValue(val);};
  set('status',nextStatus);set('reason',reason);set('stage',nextRole||session.role);set('assignedRole',nextRole);set('history',JSON.stringify(history));
  const src=sheet_('loans'),sv=src.getDataRange().getValues(),hh=sv[0].map(String),rid=hh.indexOf('id');for(let j=1;j<sv.length;j++)if(String(sv[j][rid])===String(a.recordid)){const put=(k,val)=>{const c=hh.indexOf(k);if(c>=0)src.getRange(j+1,c+1).setValue(val);};put('status',final?nextStatus:(decision==='Rejected'?'Returned to '+nextRole:'Pending '+nextRole));put('approvalStatus',nextStatus);put('lstatus',final?nextStatus:'Pending '+nextRole);put('rejectionReason',reason);break;}
  if(nextRole){notifyRole_(nextRole,'Loan request '+(decision==='Rejected'?'returned for correction':'ready for review'), 'Customer: '+a.customer+'\nAmount: NPR '+a.amount+'\nReason: '+reason+'\nApproval ID: '+a.id);}
  else if(nextStatus==='Rejected'){notifyApprovalOwner_('loans',String(a.recordid),'Rejected','Not successful: '+reason);}
  else if(nextStatus==='Approved'){notifyApprovalOwner_('loans',String(a.recordid),'Approved','');}
  if(session.role==='Admin'&&decision==='Approved'){const email=getMainAdminApprovalEmail_();if(email)MailApp.sendEmail(email,'Loan finally approved - '+a.customer,'Loan request '+a.recordid+' approved.');}
  return json_({status:'OK',message:nextStatus==='Approved'?'Loan approved':nextStatus==='Rejected'?'Not successful notification sent':'Request forwarded/returned to '+nextRole,nextStatus,nextRole});
}
function transferOptionalSavings_(body,session){
  const to=String(body.toCustomerId||''),amount=Number(body.amount||0),note=String(body.note||'');if(!to||!isFinite(amount)||amount<=0)return json_({status:'ERROR',message:'Recipient and valid amount are required'});if(to===String(session.customerId))return json_({status:'ERROR',message:'Cannot transfer to your own account'});
  const customers=sheet_('customers').getDataRange().getValues(),ch=customers[0].map(String),ci=ch.indexOf('id'),cn=ch.indexOf('name');let recipient=null;for(let i=1;i<customers.length;i++)if(String(customers[i][ci])===to){recipient={id:to,name:customers[i][cn]};break;}if(!recipient)return json_({status:'ERROR',message:'Recipient customer not found'});
  const savings=sheet_('savings').getDataRange().getValues(),sh=savings[0].map(String),owner=sh.indexOf('scustomer'),amt=sh.indexOf('samount'),typ=sh.indexOf('savingType');let available=0;for(let i=1;i<savings.length;i++)if(String(savings[i][owner])===String(session.customerId)&&(String(savings[i][typ]||'').toLowerCase()==='optional'))available+=Number(savings[i][amt]||0);
  const transfers=sheet_('transfers').getDataRange().getValues(),th=transfers[0].map(String),from=th.indexOf('fromCustomerId'),toCol=th.indexOf('toCustomerId'),ac=th.indexOf('amount');for(let i=1;i<transfers.length;i++){if(String(transfers[i][from])===String(session.customerId))available-=Number(transfers[i][ac]||0);if(String(transfers[i][toCol])===String(session.customerId))available+=Number(transfers[i][ac]||0);}
  if(amount>available)return json_({status:'ERROR',message:'Optional savings insufficient. Available: NPR '+available});
  const row={id:Utilities.getUuid(),date:new Date(),fromCustomerId:String(session.customerId),fromCustomer:session.customerName||session.username,toCustomerId:to,toCustomer:recipient.name,amount,note,status:'Completed',ownerId:String(session.customerId),ownerUsername:session.username};const tr=sheet_('transfers'),hh=tr.getRange(1,1,1,tr.getLastColumn()).getValues()[0].map(String);tr.appendRow(hh.map(k=>row[k]===undefined?'':row[k]));
  const emCol=ch.indexOf('email');for(let i=1;i<customers.length;i++)if(String(customers[i][ci])===to&&emCol>=0&&String(customers[i][emCol]).includes('@'))MailApp.sendEmail(String(customers[i][emCol]),'Savings received','You received NPR '+amount+' optional savings from '+row.fromCustomer+'.');
  return json_({status:'OK',message:'Optional savings transferred successfully',transfer:row,available:available-amount});
}

function handleApprovalLink_(params){
  const token=String(params.token||''), decision=String(params.decision||'');
  if(!token||['Approved','Rejected'].indexOf(decision)<0) return HtmlService.createHtmlOutput('Invalid approval link.');
  const sh=sheet_('approvals'), values=sh.getDataRange().getValues(), headers=values[0].map(String);
  const ti=headers.indexOf('token'), si=headers.indexOf('status'), ri=headers.indexOf('reason'), sheetCol=headers.indexOf('sheet'), idCol=headers.indexOf('recordid');
  const i=values.findIndex((r,n)=>n>0&&String(r[ti])===token);
  if(i<1) return HtmlService.createHtmlOutput('Approval link is invalid or already used.');
  if(String(values[i][si])!=='Pending') return HtmlService.createHtmlOutput('This request has already been '+values[i][si]+'.');
  sh.getRange(i+1,si+1).setValue(decision);
  const reason=decision==='Rejected'?'Rejected by Main Admin through Gmail approval link':'';
  if(ri>=0) sh.getRange(i+1,ri+1).setValue(reason);
  const sourceName=String(values[i][sheetCol]), recordId=String(values[i][idCol]);
  updateSourceApproval_(sourceName,recordId,decision,reason);
  notifyApprovalOwner_(sourceName,recordId,decision,reason);
  return HtmlService.createHtmlOutput('<div style="font-family:Arial;text-align:center;padding:35px"><h2>Home & Market Manager Nepal</h2><h3>Request '+decision+'</h3><p>The decision has been saved to Google Sheets.</p></div>');
}
function decideApproval_(id,decision,reason){
  if(['Approved','Rejected'].indexOf(String(decision))<0) return json_({status:'ERROR',message:'Invalid decision'});
  const sh=sheet_('approvals'), values=sh.getDataRange().getValues(), headers=values[0].map(String), idCol=headers.indexOf('id');
  const i=values.findIndex((r,n)=>n>0&&String(r[idCol])===String(id));
  if(i<1) return json_({status:'ERROR',message:'Approval not found'});
  const statusCol=headers.indexOf('status'),reasonCol=headers.indexOf('reason');
  sh.getRange(i+1,statusCol+1).setValue(decision);if(reasonCol>=0)sh.getRange(i+1,reasonCol+1).setValue(reason||'');
  updateSourceApproval_(String(values[i][headers.indexOf('sheet')]),String(values[i][headers.indexOf('recordid')]),decision,reason||'');
  notifyApprovalOwner_(String(values[i][headers.indexOf('sheet')]),String(values[i][headers.indexOf('recordid')]),decision,reason||'');
  return json_({status:'OK',message:'Decision saved'});
}
function updateSourceApproval_(name,id,decision,reason){
  if(!TAB_HEADERS[name]) throw new Error('Invalid source sheet');
  const sh=sheet_(name), data=sh.getDataRange().getValues(), h=data[0].map(String), idc=h.indexOf('id');
  for(let i=1;i<data.length;i++) if(String(data[i][idc])===String(id)){
    const sc=h.indexOf('status'), ac=h.indexOf('approvalStatus');
    const previous=sc>=0?String(data[i][sc]):'';
    if(name==='orders'&&decision==='Approved'&&previous!=='Approved'){
      const cartCol=h.indexOf('cartItems');
      if(cartCol>=0&&data[i][cartCol]){
        const lines=JSON.parse(String(data[i][cartCol]));
        const products=sheet_('products'), pd=products.getDataRange().getValues(), ph=pd[0].map(String), pid=ph.indexOf('id'), stock=ph.indexOf('stock');
        lines.forEach(line=>{for(let j=1;j<pd.length;j++)if(String(pd[j][pid])===String(line.id)){const remaining=Math.max(0,Number(pd[j][stock]||0)-Number(line.qty||0));products.getRange(j+1,stock+1).setValue(remaining);break;}});
      }
    }
    if(sc>=0)sh.getRange(i+1,sc+1).setValue(decision);
    const lc=h.indexOf('lstatus');if(lc>=0)sh.getRange(i+1,lc+1).setValue(decision);
    if(ac>=0)sh.getRange(i+1,ac+1).setValue(decision);
    const rc=h.indexOf('rejectionReason');if(rc>=0)sh.getRange(i+1,rc+1).setValue(reason||'');
    return;
  }
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

function notifyApprovalOwner_(name,id,decision,reason){
  try{
    const sh=sheet_(name), data=sh.getDataRange().getValues(); if(data.length<2)return;
    const h=data[0].map(String), idc=h.indexOf('id');
    for(let i=1;i<data.length;i++)if(String(data[i][idc])===String(id)){
      const row={};h.forEach((k,j)=>row[k]=data[i][j]);
      let email=String(row.useremail||row.email||'').trim();
      if(!email){const customer=String(row.ownerId||row.customerId||row.lcustomer||row.buyer||row.customer||row.tenant||'');const cs=sheet_('customers').getDataRange().getValues();if(cs.length>1){const ch=cs[0].map(String),idc2=ch.indexOf('id'),namec=ch.indexOf('name'),emc=ch.indexOf('email');for(let j=1;j<cs.length;j++)if(String(cs[j][idc2])===customer||String(cs[j][namec])===customer){if(emc>=0)email=String(cs[j][emc]||'');break;}}}
      if(email.indexOf('@')>0){const terminalFail=decision==='Rejected'&&String(reason||'').toLowerCase().includes('not successful');const subject=terminalFail?'Loan request: NOT SUCCESSFUL':'Your '+({loans:'Loan',users:'User registration',orders:'Order',rooms:'Room',roomRent:'Room rent'}[name]||name)+' request: '+decision;const msg=terminalFail?'NOT SUCCESSFUL: Field Staff did not recommend this loan request. Reason: '+reason:'Your request was returned for correction. Admin note: '+(reason||'Please review and correct your details.')+'\nOpen the app, edit the rejected request, and submit it again.';if(decision==='Approved')MailApp.sendEmail(email,subject,'Your request has been approved. You can continue in the app.');else MailApp.sendEmail(email,subject,msg);}
      return;
    }
  }catch(err){console.error('Approval notification failed: '+err);}
}

/* Google Apps Script collector. Deploy as a web app, executing as the owner.
   Results remain in a private Google Sheet; this endpoint never returns participant records. */
const SETTINGS = {
  sheetId: 'SET_YOUR_PRIVATE_SHEET_ID',
  origin: 'https://arisha-mughal.github.io',
  campaign: 'ebm-fire-safety-v1'
};
const KEY = [['electrical',0],['powder',1],['foam',2],['halotron',3],['inventory',0],['detectors',1],['alarm',2],['sprinklers',3],['suppression',0],['hoses',1],['pumps',2],['pressure',3],['location',0],['assemblycount',1],['assemblyplaces',2]];
function setup() {
  const sheet = SpreadsheetApp.openById(SETTINGS.sheetId).getSheetByName('Results');
  if (!sheet || sheet.getRange('A1').getValue() !== 'Received (UTC)') throw Error('Results template not found.');
  sheet.setFrozenRows(1);
  sheet.getRange('A1:L1').setBackground('#191b38').setFontColor('#ffffff').setFontWeight('bold').setWrap(true);
  sheet.setRowHeight(1,36);
  sheet.getRange('H2:H1001').setNumberFormat('0%');
  if(!sheet.getFilter())sheet.getRange('A1:L1001').createFilter();
  console.log('Private Results sheet connected. Ready to deploy.');
}
function doGet() {
  return ContentService.createTextOutput(JSON.stringify({service:'EBM Fire Quest results',status:'ready'})).setMimeType(ContentService.MimeType.JSON);
}
function cleanText(value,max) {
  if(typeof value!=='string'||!value.trim()||value.length>max)throw Error('Please check your player details.');
  const v=value.trim().replace(/[\u0000-\u001f]/g,' ');
  return /^[=+@-]/.test(v) ? "'"+v : v;
}
function validate(input) {
  if(!input||input.campaign!==SETTINGS.campaign||input.origin!==SETTINGS.origin)throw Error('This game is not configured for this collector.');
  if(!/^[a-f0-9-]{36}$/i.test(input.attemptId||'')||!/^[a-f0-9-]{36}$/i.test(input.nonce||''))throw Error('Invalid attempt.');
  const p=input.participant||{};
  const name=cleanText(p.name,80),department=cleanText(p.department,80);
  if(typeof p.email!=='string'||p.email.length>160||!/^\S+@\S+\.\S+$/.test(p.email))throw Error('Please check your email.');
  const email=p.email.trim().toLowerCase();
  if(!Array.isArray(input.answers)||input.answers.length!==KEY.length)throw Error('Complete all 15 questions first.');
  let correct=0;
  input.answers.forEach((a,i)=>{if(!a||a.id!==KEY[i][0]||!Number.isInteger(a.answer)||a.answer< -1||a.answer>3)throw Error('Invalid answers.');if(a.answer===KEY[i][1])correct++;});
  if(!Number.isFinite(input.durationSeconds)||input.durationSeconds<0||input.durationSeconds>86400)throw Error('Invalid duration.');
  return {name,department,email,correct,points:correct*100};
}
function recordResult(input) {
  const p=validate(input);
  const lock=LockService.getScriptLock();lock.waitLock(20000);
  try {
    const s=SpreadsheetApp.openById(SETTINGS.sheetId).getSheetByName('Results');
    if(!s)throw Error('The organiser needs to connect the results sheet.');
    const n=s.getLastRow();
    if(n>1000)throw Error('Result capacity reached. Contact the organiser.');
    if(n>1){
      const rows=s.getRange(2,1,n-1,12).getValues();
      const found=rows.some(r=>r[10]===SETTINGS.campaign&&(r[9]===input.attemptId||String(r[2]).toLowerCase()===p.email));
      if(found)return {ok:true,duplicate:true};
    }
    s.appendRow([new Date().toISOString(),p.name,p.email,p.department,p.points,p.correct,15,p.correct/15,Math.round(input.durationSeconds),input.attemptId,SETTINGS.campaign,JSON.stringify(input.answers.map(a=>a.answer))]);
    s.getRange(2,1,s.getLastRow()-1,12).sort([{column:5,ascending:false},{column:1,ascending:true}]);
    return {ok:true,duplicate:false};
  } finally {lock.releaseLock();}
}
function checkAttempt(input){
 if(input.campaign!==SETTINGS.campaign||input.origin!==SETTINGS.origin||!/^[a-f0-9-]{36}$/i.test(input.nonce||''))throw Error('Invalid request.');
 if(typeof input.email!=='string'||input.email.length>160||!/^\S+@\S+\.\S+$/.test(input.email))throw Error('Please check your email.');
 const email=input.email.trim().toLowerCase(),s=SpreadsheetApp.openById(SETTINGS.sheetId).getSheetByName('Results'),n=s.getLastRow();
 const used=n>1&&s.getRange(2,1,n-1,12).getValues().some(r=>r[10]===SETTINGS.campaign&&String(r[2]).toLowerCase()===email);
 return {ok:true,available:!used};
}
function doPost(e) {
  let input={},reply;
  try {
    const raw=e&&e.parameter&&e.parameter.payload;
    if(typeof raw!=='string'||raw.length>12000)throw Error('Invalid submission.');
    input=JSON.parse(raw);
    reply=input.action==='check'?checkAttempt(input):recordResult(input);
  } catch(err) {reply={ok:false,message:String(err.message||'Could not record result.').slice(0,180)};}
  const message={type:'fire-quest-result',nonce:input.nonce,...reply};
  const safe=JSON.stringify(message).replace(/</g,'\\u003c');
  const origin=JSON.stringify(SETTINGS.origin);
  return HtmlService.createHtmlOutput('<!doctype html><meta charset="utf-8"><p>Fire Quest result processed. Return to the game.</p><script>window.top.postMessage('+safe+','+origin+');<\/script>').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


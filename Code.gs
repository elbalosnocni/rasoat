/**
 * RASOAT V2 - Survey Engine
 * Google Apps Script backend
 *
 * Architecture:
 * Cloudflare Pages -> GAS /exec -> Google Sheet (opened by ID)
 *
 * Script Properties:
 *   SPREADSHEET_ID   = Google Sheet ID
 *   PASSWORD_PEPPER  = auto-created by setup()
 *
 * Run setup() once after setting SPREADSHEET_ID.
 */

const CFG = Object.freeze({
  EMPLOYEE_SHEET: 'DSCNV',
  ADMIN_SHEET: 'AdminUsers',
  SURVEY_SHEET: 'Surveys',
  QUESTION_SHEET: 'Questions',
  RESPONSE_SHEET: 'SurveyResponses',
  CONFIG_SHEET: 'SystemConfig',
  SPREADSHEET_PROPERTY: 'SPREADSHEET_ID',
  PEPPER_PROPERTY: 'PASSWORD_PEPPER',
  SESSION_SECONDS: 6 * 60 * 60,
  EMP_CACHE_SECONDS: 300,
  MAX_TEXT: 5000,
  MAX_CCCD: 12
});

function setup() {
  const ss = getSS_();

  ensureSheet_(ss, CFG.EMPLOYEE_SHEET, ['MaNV','HoTen','CCCD']);
  ensureSheet_(ss, CFG.ADMIN_SHEET, ['Username','PasswordHash','PasswordSalt','Active','CreatedAt','UpdatedAt']);
  ensureSheet_(ss, CFG.SURVEY_SHEET,
    ['SurveyId','Code','Name','Title','Description','ConsentText','Status','StartAt','EndAt','Version','UpdatedAt','UpdatedBy']);
  ensureSheet_(ss, CFG.QUESTION_SHEET,
    ['QuestionId','SurveyId','Section','Label','Type','Required','Options','Placeholder','HelpText','Active','SortOrder','UpdatedAt']);
  ensureSheet_(ss, CFG.RESPONSE_SHEET,
    ['ResponseId','SurveyId','SurveyVersion','MaNV','HoTen','CCCD','Consent','Status','StartedAt','SubmittedAt','AnswersJson','UpdatedAt']);
  ensureSheet_(ss, CFG.CONFIG_SHEET, ['Key','Value','UpdatedAt']);

  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty(CFG.PEPPER_PROPERTY)) {
    props.setProperty(CFG.PEPPER_PROPERTY, randomHex_(64));
  }

  const admin = getSheet_(CFG.ADMIN_SHEET);
  if (admin.getLastRow() <= 1) {
    const salt = randomHex_(32);
    admin.appendRow(['admin', hashPassword_('ChangeMe@123', salt), salt, true, new Date(), new Date()]);
    Logger.log('Default admin: admin / ChangeMe@123 - change immediately.');
  }

  const surveys = getSheet_(CFG.SURVEY_SHEET);
  if (surveys.getLastRow() <= 1) {
    const id = 'SV_' + Utilities.getUuid().slice(0,8).toUpperCase();
    surveys.appendRow([
      id, 'KSTT_2026', 'Khảo sát hiện tại',
      'KHẢO SÁT THU THẬP THÔNG TIN',
      'Nội dung khảo sát có thể chỉnh sửa hoàn toàn trong Admin.',
      'Tôi đồng ý cho Công ty xử lý dữ liệu theo nội dung thông báo.',
      'PUBLISHED', new Date(), new Date('2026-12-31T23:59:00'), 1, new Date(), 'setup'
    ]);
    const qs = [
      ['Q_NGAYSINH',id,'Thông tin cá nhân','Ngày, tháng, năm sinh','date',true,'','','',true,10],
      ['Q_GIOITINH',id,'Thông tin cá nhân','Giới tính','select',true,'Nam|Nữ|Khác','','',true,20],
      ['Q_DANTOC',id,'Thông tin cá nhân','Dân tộc','text',false,'','Nhập dân tộc','',true,30],
      ['Q_TONGIAO',id,'Thông tin cá nhân','Tôn giáo','text',false,'','Nhập tôn giáo','',true,40],
      ['Q_DC_THUONGTRU',id,'Địa chỉ','Địa chỉ thường trú','textarea',true,'','Nhập đầy đủ địa chỉ','',true,50],
      ['Q_DC_TAMTRU',id,'Địa chỉ','Địa chỉ tạm trú','textarea',false,'','Nếu có','',true,60],
      ['Q_TRINHDO_VH',id,'Trình độ','Trình độ văn hóa','text',false,'','','',true,70],
      ['Q_TRINHDO_CM',id,'Trình độ','Trình độ chuyên môn','text',false,'','','',true,80],
      ['Q_MIENGIAM',id,'Khác','Trường hợp miễn/giảm','multiselect',false,
       'Không thuộc trường hợp miễn/giảm|Trẻ em|Người cao tuổi|Người khuyết tật|Hộ nghèo/cận nghèo|Người có công với cách mạng|Thân nhân liệt sĩ|Khác','','',true,90]
    ];
    getSheet_(CFG.QUESTION_SHEET).getRange(2,1,qs.length,12).setValues(qs.map(r=>r.concat([new Date()])));
  }
  return {ok:true,message:'Setup completed.'};
}

function doGet(e) {
  return json_({ok:true,service:'rasoat-survey-api',version:'2.0',time:new Date().toISOString()});
}

function doPost(e) {
  try {
    const b = parse_(e), a = String(b.action||'').trim();
    switch(a) {
      case 'getSurvey': return getSurvey_(b);
      case 'employeeLogin': return employeeLogin_(b);
      case 'saveSurveyResponse': return saveSurveyResponse_(b);
      case 'adminLogin': return adminLogin_(b);
      case 'adminDashboard': return adminDashboard_(b);
      case 'adminGetSurvey': return adminGetSurvey_(b);
      case 'adminSaveSurvey': return adminSaveSurvey_(b);
      case 'adminAddQuestion': return adminAddQuestion_(b);
      case 'adminUpdateQuestion': return adminUpdateQuestion_(b);
      case 'adminDeleteQuestion': return adminDeleteQuestion_(b);
      case 'adminSetStatus': return adminSetStatus_(b);
      case 'adminLogout': destroySession_(b.token); return json_({ok:true});
      case 'adminChangePassword': return adminChangePassword_(b);
      default: return json_({ok:false,error:'INVALID_ACTION',message:'Yêu cầu không hợp lệ.'});
    }
  } catch(err) {
    console.error(err);
    return json_({ok:false,error:'SERVER_ERROR',message:safe_(err)});
  }
}

function getSurvey_(b) {
  const code = String(b.code||'KSTT_2026').trim();
  const survey = findSurveyByCode_(code);
  if (!survey) return json_({ok:false,error:'SURVEY_NOT_FOUND',message:'Không tìm thấy khảo sát.'});
  if (String(survey.Status).toUpperCase() !== 'PUBLISHED') {
    return json_({ok:false,error:'SURVEY_CLOSED',message:'Khảo sát chưa mở hoặc đã đóng.'});
  }
  return json_({ok:true,survey:publicSurvey_(survey)});
}

function employeeLogin_(b) {
  const cccd = normalizeCCCD_(b.cccd);
  if (!/^\d{9,12}$/.test(cccd)) return json_({ok:false,error:'INVALID_CCCD',message:'CCCD không hợp lệ.'});
  const emp = findEmployee_(cccd);
  if (!emp) return json_({ok:false,error:'NOT_FOUND',message:'Không tìm thấy nhân viên.'});
  const survey = findSurveyByCode_(String(b.code||'KSTT_2026').trim());
  if (!survey || String(survey.Status).toUpperCase()!=='PUBLISHED') {
    return json_({ok:false,error:'SURVEY_CLOSED',message:'Khảo sát chưa mở hoặc đã đóng.'});
  }
  const token = createSession_('employee',{maNV:emp.MaNV,hoTen:emp.HoTen,cccd});
  const prev = findLatestResponse_(survey.SurveyId, emp.MaNV);
  return json_({ok:true,token,employee:{maNV:emp.MaNV,hoTen:emp.HoTen,cccd},previous:prev||null,survey:publicSurvey_(survey)});
}

function saveSurveyResponse_(b) {
  const s = getSession_(b.token);
  if (!s || s.type!=='employee') return json_({ok:false,error:'UNAUTHORIZED',message:'Phiên đã hết hạn.'});
  const p = s.payload;
  const survey = findSurveyByCode_(String(b.code||'KSTT_2026').trim());
  if (!survey || String(survey.Status).toUpperCase()!=='PUBLISHED') {
    return json_({ok:false,error:'SURVEY_CLOSED',message:'Khảo sát đã đóng.'});
  }
  const consent = b.consent === true;
  const answers = (b.answers && typeof b.answers === 'object') ? b.answers : {};
  const questions = getQuestions_(survey.SurveyId);
  for (const q of questions) {
    if (!q.Active) continue;
    if (q.Required && q.QuestionId!=='Q_CONSENT' && isEmptyAnswer_(answers[q.QuestionId])) {
      return json_({ok:false,error:'MISSING_REQUIRED',questionId:q.QuestionId,message:'Vui lòng điền: '+q.Label});
    }
  }
  const existing = findLatestResponse_(survey.SurveyId,p.maNV);
  const now = new Date();
  const rowObj = {
    ResponseId: existing ? existing.ResponseId : Utilities.getUuid(),
    SurveyId: survey.SurveyId,
    SurveyVersion: survey.Version,
    MaNV: p.maNV, HoTen: p.hoTen, CCCD: p.cccd,
    Consent: consent ? 'Đồng ý' : 'Không đồng ý',
    Status: consent ? 'Đã hoàn tất' : 'Từ chối',
    StartedAt: existing ? existing.StartedAt : now,
    SubmittedAt: consent ? now : '',
    AnswersJson: JSON.stringify(answers),
    UpdatedAt: now
  };
  upsertResponse_(rowObj,existing);
  destroySession_(b.token);
  return json_({ok:true,message:consent?'Đã lưu khảo sát thành công.':'Đã ghi nhận lựa chọn.'});
}

function adminLogin_(b) {
  const u=clean_(b.username,100), pw=String(b.password||'');
  const a=findAdmin_(u);
  if (!a || !isActive_(a.active) || !secureEqual_(hashPassword_(pw,a.salt),a.hash))
    return json_({ok:false,error:'INVALID_LOGIN',message:'Tài khoản hoặc mật khẩu không đúng.'});
  return json_({ok:true,token:createSession_('admin',{username:u})});
}

function adminDashboard_(b) {
  if (!requireAdmin_(b.token)) return json_({ok:false,error:'UNAUTHORIZED',message:'Phiên Admin đã hết hạn.'});
  const survey=findSurveyByCode_(String(b.code||'KSTT_2026').trim());
  if (!survey) return json_({ok:false,error:'SURVEY_NOT_FOUND'});
  const emps=getEmployees_(), rs=getResponses_(survey.SurveyId), latest={};
  rs.forEach(r=>{const old=latest[r.MaNV];if(!old||dateMs_(r.UpdatedAt)>=dateMs_(old.UpdatedAt))latest[r.MaNV]=r;});
  const summary={total:emps.length,completed:0,inProgress:0,refused:0,notStarted:0};
  const rows=emps.map(e=>{
    const r=latest[e.MaNV], st=r?String(r.Status||'Đang kê khai'):'Chưa điền';
    if(st==='Đã hoàn tất')summary.completed++; else if(st==='Từ chối')summary.refused++; else if(st==='Đang kê khai')summary.inProgress++; else summary.notStarted++;
    return {maNV:e.MaNV,hoTen:e.HoTen,cccd:maskCCCD_(e.CCCD),status:st,updatedAt:r?formatDate_(r.UpdatedAt):''};
  });
  return json_({ok:true,survey:publicSurvey_(survey),summary,rows});
}

function adminGetSurvey_(b) {
  if (!requireAdmin_(b.token)) return json_({ok:false,error:'UNAUTHORIZED'});
  const survey=findSurveyByCode_(String(b.code||'KSTT_2026').trim());
  if(!survey)return json_({ok:false,error:'SURVEY_NOT_FOUND'});
  return json_({ok:true,survey:publicSurvey_(survey),questions:getQuestions_(survey.SurveyId)});
}

function adminSaveSurvey_(b) {
  const session=requireAdmin_(b.token); if(!session)return json_({ok:false,error:'UNAUTHORIZED'});
  const sh=getSheet_(CFG.SURVEY_SHEET), survey=findSurveyByCode_(String(b.code||'KSTT_2026').trim());
  if(!survey)return json_({ok:false,error:'SURVEY_NOT_FOUND'});
  const patch=b.patch||{}, headers=headerMap_(sh), row=survey._row;
  ['Name','Title','Description','ConsentText','StartAt','EndAt','Version'].forEach(k=>{
    if(Object.prototype.hasOwnProperty.call(patch,k)) sh.getRange(row,headers[k]+1).setValue(clean_(patch[k],10000));
  });
  sh.getRange(row,headers.UpdatedAt+1).setValue(new Date());
  sh.getRange(row,headers.UpdatedBy+1).setValue(session.username);
  return json_({ok:true});
}

function adminSetStatus_(b) {
  const session=requireAdmin_(b.token); if(!session)return json_({ok:false,error:'UNAUTHORIZED'});
  const survey=findSurveyByCode_(String(b.code||'KSTT_2026').trim()); if(!survey)return json_({ok:false,error:'SURVEY_NOT_FOUND'});
  const status=String(b.status||'').toUpperCase();
  if(['DRAFT','PUBLISHED','CLOSED'].indexOf(status)<0)return json_({ok:false,error:'INVALID_STATUS'});
  const sh=getSheet_(CFG.SURVEY_SHEET), hm=headerMap_(sh);
  sh.getRange(survey._row,hm.Status+1).setValue(status);
  sh.getRange(survey._row,hm.UpdatedAt+1).setValue(new Date());
  sh.getRange(survey._row,hm.UpdatedBy+1).setValue(session.username);
  return json_({ok:true});
}

function adminAddQuestion_(b) {
  const session=requireAdmin_(b.token); if(!session)return json_({ok:false,error:'UNAUTHORIZED'});
  const survey=findSurveyByCode_(String(b.code||'KSTT_2026').trim()); if(!survey)return json_({ok:false,error:'SURVEY_NOT_FOUND'});
  const q=b.question||{}, id='Q_'+Utilities.getUuid().slice(0,8).toUpperCase();
  const sh=getSheet_(CFG.QUESTION_SHEET);
  sh.appendRow([id,survey.SurveyId,clean_(q.Section,200),clean_(q.Label,1000),String(q.Type||'text'),
    !!q.Required,Array.isArray(q.Options)?q.Options.join('|'):clean_(q.Options,5000),
    clean_(q.Placeholder,500),clean_(q.HelpText,1000),true,Number(q.SortOrder)||999,new Date()]);
  return json_({ok:true,questionId:id});
}

function adminUpdateQuestion_(b) {
  const session=requireAdmin_(b.token); if(!session)return json_({ok:false,error:'UNAUTHORIZED'});
  const sh=getSheet_(CFG.QUESTION_SHEET), hm=headerMap_(sh), id=String(b.questionId||'');
  const row=findRowBy_(sh,hm.QuestionId,id); if(row<2)return json_({ok:false,error:'QUESTION_NOT_FOUND'});
  const q=b.question||{};
  const vals={Section:q.Section,Label:q.Label,Type:q.Type,Required:q.Required,Options:Array.isArray(q.Options)?q.Options.join('|'):q.Options,Placeholder:q.Placeholder,HelpText:q.HelpText,Active:q.Active,SortOrder:q.SortOrder};
  Object.keys(vals).forEach(k=>{if(vals[k]!==undefined)sh.getRange(row,hm[k]+1).setValue(vals[k]);});
  sh.getRange(row,hm.UpdatedAt+1).setValue(new Date());
  return json_({ok:true});
}

function adminDeleteQuestion_(b) {
  if(!requireAdmin_(b.token))return json_({ok:false,error:'UNAUTHORIZED'});
  const sh=getSheet_(CFG.QUESTION_SHEET),hm=headerMap_(sh),row=findRowBy_(sh,hm.QuestionId,String(b.questionId||''));
  if(row<2)return json_({ok:false,error:'QUESTION_NOT_FOUND'});
  sh.getRange(row,hm.Active+1).setValue(false);
  sh.getRange(row,hm.UpdatedAt+1).setValue(new Date());
  return json_({ok:true});
}

function adminChangePassword_(b){
  const s=requireAdmin_(b.token); if(!s)return json_({ok:false,error:'UNAUTHORIZED'});
  const old=String(b.oldPassword||''), nw=String(b.newPassword||''), cf=String(b.confirmPassword||'');
  if(nw.length<10)return json_({ok:false,error:'WEAK_PASSWORD',message:'Mật khẩu mới tối thiểu 10 ký tự.'});
  if(nw!==cf)return json_({ok:false,error:'PASSWORD_MISMATCH',message:'Xác nhận mật khẩu không khớp.'});
  const a=findAdmin_(s.username);
  if(!a||!secureEqual_(hashPassword_(old,a.salt),a.hash))return json_({ok:false,error:'INVALID_PASSWORD',message:'Mật khẩu hiện tại không đúng.'});
  const salt=randomHex_(32),sh=getSheet_(CFG.ADMIN_SHEET),hm=headerMap_(sh);
  sh.getRange(a.row,hm.PasswordHash+1,1,2).setValues([[hashPassword_(nw,salt),salt]]);
  sh.getRange(a.row,hm.UpdatedAt+1).setValue(new Date());
  return json_({ok:true});
}

function publicSurvey_(s){return {
  SurveyId:s.SurveyId,Code:s.Code,Name:s.Name,Title:s.Title,Description:s.Description,ConsentText:s.ConsentText,
  Status:s.Status,StartAt:dateISO_(s.StartAt),EndAt:dateISO_(s.EndAt),Version:Number(s.Version||1)
};}

function getQuestions_(surveyId){
  const sh=getSheet_(CFG.QUESTION_SHEET),rows=sh.getDataRange().getValues(); if(rows.length<2)return[];
  const hm=headerMap_(sh),out=[];
  for(let i=1;i<rows.length;i++){
    const r=rows[i]; if(String(r[hm.SurveyId]||'')!==String(surveyId)||!isActive_(r[hm.Active]))continue;
    out.push({_row:i+1,QuestionId:String(r[hm.QuestionId]||''),Section:String(r[hm.Section]||''),Label:String(r[hm.Label]||''),
      Type:String(r[hm.Type]||'text'),Required:isActive_(r[hm.Required]),Options:String(r[hm.Options]||'').split('|').filter(Boolean),
      Placeholder:String(r[hm.Placeholder]||''),HelpText:String(r[hm.HelpText]||''),Active:true,SortOrder:Number(r[hm.SortOrder]||999)});
  }
  return out.sort((a,b)=>a.SortOrder-b.SortOrder);
}

function findSurveyByCode_(code){
  const sh=getSheet_(CFG.SURVEY_SHEET),rows=sh.getDataRange().getValues();if(rows.length<2)return null;const hm=headerMap_(sh);
  for(let i=1;i<rows.length;i++){if(String(rows[i][hm.Code]||'').trim()===code){const o={_row:i+1};Object.keys(hm).forEach(k=>o[k]=rows[i][hm[k]]);return o;}}
  return null;
}

function findEmployee_(cccd){return getEmployees_().find(e=>normalizeCCCD_(e.CCCD)===cccd)||null;}
function getEmployees_(){
  const c=CacheService.getScriptCache(),hit=c.get('employees:v2');if(hit)try{return JSON.parse(hit)}catch(e){}
  const sh=getSheet_(CFG.EMPLOYEE_SHEET),rows=sh.getDataRange().getDisplayValues();if(rows.length<2)return[];
  const hm=employeeHeaderMap_(rows[0]),out=[];
  for(let i=1;i<rows.length;i++){const r=rows[i],e={_row:i+1,MaNV:r[hm.MaNV]||'',HoTen:r[hm.HoTen]||'',CCCD:normalizeCCCD_(r[hm.CCCD])};if(e.MaNV||e.HoTen||e.CCCD)out.push(e);}
  c.put('employees:v2',JSON.stringify(out),CFG.EMP_CACHE_SECONDS);return out;
}

function getResponses_(surveyId){
  const sh=getSheet_(CFG.RESPONSE_SHEET),rows=sh.getDataRange().getValues();if(rows.length<2)return[];const hm=headerMap_(sh),out=[];
  for(let i=1;i<rows.length;i++){if(String(rows[i][hm.SurveyId]||'')!==String(surveyId))continue;const o={_row:i+1};Object.keys(hm).forEach(k=>o[k]=rows[i][hm[k]]);out.push(o);}return out;
}
function findLatestResponse_(surveyId,maNV){let x=null;getResponses_(surveyId).forEach(r=>{if(String(r.MaNV)===String(maNV)&&(!x||dateMs_(r.UpdatedAt)>dateMs_(x.UpdatedAt)))x=r;});return x;}

function upsertResponse_(obj,existing){
  const sh=getSheet_(CFG.RESPONSE_SHEET),hm=headerMap_(sh),row=[];
  Object.keys(hm).forEach(k=>row[hm[k]]=obj[k]===undefined?'':obj[k] instanceof Date?obj[k]:obj[k]);
  if(existing)sh.getRange(existing._row,1,1,Object.keys(hm).length).setValues([row]);
  else sh.getRange(sh.getLastRow()+1,1,1,Object.keys(hm).length).setValues([row]);
}

function createSession_(type,payload){const t=Utilities.getUuid()+'-'+Utilities.getUuid();CacheService.getScriptCache().put('sess:'+t,JSON.stringify({type,payload,expiresAt:Date.now()+CFG.SESSION_SECONDS*1000}),CFG.SESSION_SECONDS);return t;}
function getSession_(t){if(!t)return null;const raw=CacheService.getScriptCache().get('sess:'+t);if(!raw)return null;try{const s=JSON.parse(raw);if(s.expiresAt<Date.now()){destroySession_(t);return null;}return s}catch(e){destroySession_(t);return null;}}
function destroySession_(t){if(t)CacheService.getScriptCache().remove('sess:'+t);}
function requireAdmin_(t){const s=getSession_(t);return s&&s.type==='admin'?s.payload:null;}

function findAdmin_(u){const sh=getSheet_(CFG.ADMIN_SHEET),rows=sh.getDataRange().getValues();if(rows.length<2)return null;const hm=headerMap_(sh);for(let i=1;i<rows.length;i++)if(String(rows[i][hm.Username]||'').trim()===u)return {row:i+1,username:u,hash:String(rows[i][hm.PasswordHash]||''),salt:String(rows[i][hm.PasswordSalt]||''),active:rows[i][hm.Active]};return null;}
function hashPassword_(p,s){const pepper=PropertiesService.getScriptProperties().getProperty(CFG.PEPPER_PROPERTY)||'';return bytesHex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(p)+':'+String(s)+':'+pepper,Utilities.Charset.UTF_8));}
function secureEqual_(a,b){a=String(a);b=String(b);if(a.length!==b.length)return false;let x=0;for(let i=0;i<a.length;i++)x|=a.charCodeAt(i)^b.charCodeAt(i);return x===0;}
function randomHex_(n){return bytesHex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,Utilities.getUuid()+Utilities.getUuid()+Date.now()+Math.random(),Utilities.Charset.UTF_8)).slice(0,n);}

function getSS_(){const id=PropertiesService.getScriptProperties().getProperty(CFG.SPREADSHEET_PROPERTY);if(!id)throw new Error('Chưa cấu hình SPREADSHEET_ID trong Script Properties.');return SpreadsheetApp.openById(id);}
function getSheet_(n){const s=getSS_().getSheetByName(n);if(!s)throw new Error('Không tìm thấy sheet: '+n);return s;}
function ensureSheet_(ss,n,h){let s=ss.getSheetByName(n);if(!s)s=ss.insertSheet(n);if(s.getLastRow()===0){s.getRange(1,1,1,h.length).setValues([h]);s.setFrozenRows(1);}}
function headerMap_(sh){const h=sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0],o={};h.forEach((v,i)=>{if(String(v).trim())o[String(v).trim()]=i});return o;}
function findRowBy_(sh,col,value){const vals=sh.getRange(2,col+1,Math.max(0,sh.getLastRow()-1),1).getDisplayValues();for(let i=0;i<vals.length;i++)if(String(vals[i][0])===value)return i+2;return -1;}
function employeeHeaderMap_(h){const o={};h.forEach((v,i)=>o[String(v).trim()]=i);if(o.MaNV===undefined&&o['Mã NV']!==undefined)o.MaNV=o['Mã NV'];if(o.HoTen===undefined&&o['Họ tên']!==undefined)o.HoTen=o['Họ tên'];if(o.HoTen===undefined&&o['Họ và tên']!==undefined)o.HoTen=o['Họ và tên'];if(o.CCCD===undefined&&o['Số CCCD']!==undefined)o.CCCD=o['Số CCCD'];if(o.CCCD===undefined&&o['CCCD']!==undefined)o.CCCD=o['CCCD'];['MaNV','HoTen','CCCD'].forEach(k=>{if(o[k]===undefined)throw new Error('DSCNV thiếu cột '+k)});return o;}
function normalizeCCCD_(v){return String(v??'').replace(/\D/g,'').slice(0,CFG.MAX_CCCD);}
function maskCCCD_(v){const s=normalizeCCCD_(v);return s.length>4?'*'.repeat(s.length-4)+s.slice(-4):'*'.repeat(s.length);}
function clean_(v,n){return String(v??'').trim().slice(0,n||CFG.MAX_TEXT);}
function isEmptyAnswer_(v){if(Array.isArray(v))return v.length===0;return String(v??'').trim()==='';}
function isActive_(v){return v===true||String(v).toLowerCase()==='true'||String(v).toLowerCase()==='yes';}
function dateMs_(v){const d=v instanceof Date?v:new Date(v);const t=d.getTime();return isNaN(t)?0:t;}
function formatDate_(v){if(!v)return'';const d=v instanceof Date?v:new Date(v);return isNaN(d)?String(v):Utilities.formatDate(d,Session.getScriptTimeZone(),'dd/MM/yyyy HH:mm:ss');}
function dateISO_(v){if(!v)return'';const d=v instanceof Date?v:new Date(v);return isNaN(d)?String(v):d.toISOString();}
function bytesHex_(b){return b.map(x=>('0'+(x<0?x+256:x).toString(16)).slice(-2)).join('');}
function safe_(e){return String(e&&e.message?e.message:e).slice(0,500);}
function parse_(e){const raw=e&&e.postData&&e.postData.contents||'{}';return JSON.parse(raw);}
function json_(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);}

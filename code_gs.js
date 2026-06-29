/* ================================================
   정읍지사 식수조사 - Apps Script API 모드
   GitHub Pages에서 fetch로 호출하는 방식
   ================================================ */

var SS = SpreadsheetApp.getActiveSpreadsheet();
function getSheet(name){ return SS.getSheetByName(name); }

/* ── JSON 응답 ── */
function jsonRes(data){
  var output = ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}

/* ================================================
   OPTIONS — CORS preflight 대응
   ================================================ */
function doOptions(e){
  return ContentService.createTextOutput('')
    .setMimeType(ContentService.MimeType.TEXT);
}

/* ================================================
   GET — 데이터 조회 API
   ================================================ */
function doGet(e){
  var action   = e.parameter.action   || '';
  var callback = e.parameter.callback || '';
  var postdata = e.parameter.postdata || '';

  /* POST 데이터가 GET으로 넘어온 경우 */
  if(postdata){
    try{
      var body   = JSON.parse(postdata);
      var act    = body.action || '';
      var result;
      if(act==='submitVote')       result = submitVote(body);
      else if(act==='saveMenu')    result = saveMenu(body);
      else if(act==='deleteMenu')  result = deleteMenu(body);
      else if(act==='addStaff')    result = addStaff(body);
      else if(act==='bulkStaff')   result = bulkStaff(body);
      else if(act==='toggleStaff') result = toggleStaff(body);
      else if(act==='changePin')   result = changePin(body);
      else result = {error:'unknown action: '+act};

      var out = callback
        ? ContentService.createTextOutput(callback+'('+JSON.stringify(result)+')')
            .setMimeType(ContentService.MimeType.JAVASCRIPT)
        : jsonRes(result);
      return out;
    } catch(err){
      var errRes = {error: err.toString()};
      return callback
        ? ContentService.createTextOutput(callback+'('+JSON.stringify(errRes)+')')
            .setMimeType(ContentService.MimeType.JAVASCRIPT)
        : jsonRes(errRes);
    }
  }

  /* 일반 GET 조회 */
  var data;
  if(action==='getInit'){
    data = {
      staff   : getStaff(),
      menu    : getMenu(today()),
      votes   : getVotes(today()),
      deadline: getDeadline(),
      pin     : getAdminPin()
    };
  } else if(action==='getAdmin'){
    var date = e.parameter.date || today();
    data = { staff: getStaff(), votes: getVotes(date) };
  } else if(action==='getMenu'){
    var date = e.parameter.date || today();
    data = { menu: getMenu(date) };
  } else {
    data = {error:'unknown action'};
  }

  if(callback){
    return ContentService
      .createTextOutput(callback+'('+JSON.stringify(data)+')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return jsonRes(data);
}

/* ================================================
   POST — 데이터 저장 API
   ================================================ */
function doPost(e){
  var body   = JSON.parse(e.postData.contents);
  var action = body.action || '';

  if(action === 'submitVote')  return jsonRes(submitVote(body));
  if(action === 'saveMenu')    return jsonRes(saveMenu(body));
  if(action === 'deleteMenu')  return jsonRes(deleteMenu(body));
  if(action === 'addStaff')    return jsonRes(addStaff(body));
  if(action === 'bulkStaff')   return jsonRes(bulkStaff(body));
  if(action === 'toggleStaff') return jsonRes(toggleStaff(body));
  if(action === 'changePin')   return jsonRes(changePin(body));

  return jsonRes({error:'unknown action'});
}

/* ================================================
   내부 헬퍼 함수
   ================================================ */
function today(){
  return Utilities.formatDate(new Date(),'Asia/Seoul','yyyy-MM-dd');
}

function getStaff(){
  var sh=getSheet('직원명단');
  var data=sh.getDataRange().getValues();
  var result=[];
  for(var i=1;i<data.length;i++){
    if(!data[i][0]) continue;
    result.push({id:String(data[i][0]),name:String(data[i][1]),active:String(data[i][2])==='Y'});
  }
  return result;
}

function getMenu(date){
  var sh=getSheet('식단표');
  var data=sh.getDataRange().getValues();
  for(var i=1;i<data.length;i++){
    if(!data[i][0]) continue;
    var d=Utilities.formatDate(new Date(data[i][0]),'Asia/Seoul','yyyy-MM-dd');
    if(d===date) return String(data[i][1]);
  }
  return '';
}

function getVotes(date){
  var sh=getSheet('투표응답');
  var data=sh.getDataRange().getValues();
  var result=[];
  for(var i=1;i<data.length;i++){
    if(!data[i][0]) continue;
    var d=Utilities.formatDate(new Date(data[i][0]),'Asia/Seoul','yyyy-MM-dd');
    if(d===date) result.push({id:String(data[i][1]),name:String(data[i][2]),vote:String(data[i][3])});
  }
  return result;
}

function getAdminPin(){
  var sh=getSheet('설정');
  var data=sh.getDataRange().getValues();
  for(var i=1;i<data.length;i++){
    if(data[i][0]==='admin_pin') return String(data[i][1]);
  }
  return '0000';
}

function getDeadline(){
  var sh=getSheet('설정');
  var data=sh.getDataRange().getValues();
  for(var i=1;i<data.length;i++){
    if(data[i][0]==='deadline') return String(data[i][1]);
  }
  return '09:30';
}

function submitVote(arg){
  var sh=getSheet('투표응답');
  var data=sh.getDataRange().getValues();
  var now=Utilities.formatDate(new Date(),'Asia/Seoul','yyyy-MM-dd HH:mm:ss');
  for(var i=1;i<data.length;i++){
    if(!data[i][0]) continue;
    var d=Utilities.formatDate(new Date(data[i][0]),'Asia/Seoul','yyyy-MM-dd');
    if(d===arg.date && String(data[i][1])===String(arg.id)){
      sh.getRange(i+1,4).setValue(arg.vote);
      sh.getRange(i+1,5).setValue(now);
      return {ok:true};
    }
  }
  sh.appendRow([arg.date,arg.id,arg.name,arg.vote,now]);
  return {ok:true};
}

function saveMenu(arg){
  var sh=getSheet('식단표');
  var data=sh.getDataRange().getValues();
  for(var i=1;i<data.length;i++){
    if(!data[i][0]) continue;
    var d=Utilities.formatDate(new Date(data[i][0]),'Asia/Seoul','yyyy-MM-dd');
    if(d===arg.date){ sh.getRange(i+1,2).setValue(arg.menu); return {ok:true}; }
  }
  sh.appendRow([arg.date,arg.menu]);
  return {ok:true};
}

function deleteMenu(arg){
  var sh=getSheet('식단표');
  var data=sh.getDataRange().getValues();
  for(var i=1;i<data.length;i++){
    if(!data[i][0]) continue;
    var d=Utilities.formatDate(new Date(data[i][0]),'Asia/Seoul','yyyy-MM-dd');
    if(d===arg.date){ sh.deleteRow(i+1); return {ok:true}; }
  }
  return {ok:false};
}

function addStaff(arg){
  var sh=getSheet('직원명단');
  var data=sh.getDataRange().getValues();
  for(var i=1;i<data.length;i++){
    if(String(data[i][0])===String(arg.id)) return {ok:false,msg:'이미 등록된 사번입니다.'};
  }
  sh.appendRow([arg.id,arg.name,'Y']);
  return {ok:true};
}

function bulkStaff(arg){
  var sh=getSheet('직원명단');
  var data=sh.getDataRange().getValues();
  var existing={};
  for(var i=1;i<data.length;i++) existing[String(data[i][0])]=true;
  var added=0;
  arg.list.forEach(function(s){
    if(!existing[String(s.id)]){ sh.appendRow([s.id,s.name,'Y']); added++; }
  });
  return {ok:true,added:added};
}

function toggleStaff(arg){
  var sh=getSheet('직원명단');
  var data=sh.getDataRange().getValues();
  for(var i=1;i<data.length;i++){
    if(String(data[i][0])===String(arg.id)){
      var cur=String(data[i][2])==='Y';
      sh.getRange(i+1,3).setValue(cur?'N':'Y');
      return {ok:true,active:!cur};
    }
  }
  return {ok:false};
}

function changePin(arg){
  var cur=getAdminPin();
  if(arg.current!==cur) return {ok:false,msg:'현재 비밀번호가 틀렸습니다.'};
  var sh=getSheet('설정');
  var data=sh.getDataRange().getValues();
  for(var i=1;i<data.length;i++){
    if(data[i][0]==='admin_pin'){
      sh.getRange(i+1,2).setValue(arg.newPin);
      return {ok:true};
    }
  }
  sh.appendRow(['admin_pin',arg.newPin]);
  return {ok:true};
}

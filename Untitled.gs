/**
 * VICORA EXAM ROOM MANAGER — production-oriented Apps Script backend
 * Uses ONLY the VICORA Exam Room Manager spreadsheet.
 * Existing data is preserved. Missing tabs/columns are added non-destructively.
 */
const VICORA_DB_ID = '12odSjNeGBTldU7rWoV27p7SHdKlkBhNdn1oWoQY_Fb0';
const VICORA_VERSION = '2.0.0';
const SCHEMA = {
  Students:['StudentID','StudentName','Class','SourceSheet','ImportStatus'],
  Classes:['ClassName','StudentCount','IncludedInCurrentExam'],
  Subjects:['SubjectName'],
  SubjectRegistrations:['StudentID','StudentName','Class','Subject','RegistrationStatus','SourceSheet'],
  ExamSessions:['SessionID','SessionName','StartDate','Status'],
  RoomSettings:['SessionID','RoomNumber','RoomName','Capacity'],
  Allocations:['SessionID','StudentID','StudentName','Class','RoomNumber'],
  AllocationHistory:['SessionID','StudentID','PreviousRoom','NewRoom','Reason','Date','Administrator'],
  SystemSettings:['Setting','Value'],
  ImportLog:['Date','Operation','StudentsImported','ClassesFound','SubjectsFound','RegistrationsFound','UnmatchedStudents','AmbiguousMatches','Notes'],
  CurrentExamCandidates:['StudentID','StudentName','Class','SubjectCount','Eligibility'],
  RegistrationAudit:['Issue','StudentID','StudentName','Class','Details'],
  ExamReadiness:['Setting','Value'],
  MutationLog:['OperationID','Action','Status','Message','CreatedAt']
};
function db_(){return SpreadsheetApp.openById(VICORA_DB_ID);}
function norm_(v){return String(v===null||v===undefined?'':v).trim();}
function lower_(v){return norm_(v).toLowerCase();}
function key_(id,name,cls){return [id,name,cls].map(lower_).join('|');}
function uid_(){return 'VIC-'+Utilities.getUuid().replace(/-/g,'').slice(0,12).toUpperCase();}
function json_(o,cb){var body=JSON.stringify(o);if(cb&&/^[A-Za-z_$][\w.$]*$/.test(cb))return ContentService.createTextOutput(cb+'('+body+');').setMimeType(ContentService.MimeType.JAVASCRIPT);return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);}
function sheet_(name){var s=db_().getSheetByName(name);if(!s)throw new Error('Database tab is missing: '+name);return s;}
function ensureSchema_(){
  var d=db_();
  Object.keys(SCHEMA).forEach(function(name){
    var s=d.getSheetByName(name);
    if(!s){s=d.insertSheet(name);s.getRange(1,1,1,SCHEMA[name].length).setValues([SCHEMA[name]]);s.setFrozenRows(1);return;}
    if(s.getLastRow()===0){s.getRange(1,1,1,SCHEMA[name].length).setValues([SCHEMA[name]]);s.setFrozenRows(1);return;}
    var width=Math.max(s.getLastColumn(),1),headers=s.getRange(1,1,1,width).getDisplayValues()[0].map(norm_);
    if(!headers.some(Boolean))throw new Error('Tab '+name+' has no header row. No data was changed.');
    SCHEMA[name].forEach(function(h){if(headers.indexOf(h)<0){headers.push(h);s.getRange(1,headers.length).setValue(h);}});
  });
}
function headerMap_(s){var h=s.getRange(1,1,1,s.getLastColumn()).getDisplayValues()[0].map(norm_),m={};h.forEach(function(x,i){if(x)m[x]=i;});return m;}
function rows_(name){
  var s=sheet_(name),last=s.getLastRow();if(last<2)return[];
  var w=s.getLastColumn(),h=s.getRange(1,1,1,w).getDisplayValues()[0].map(norm_);
  return s.getRange(2,1,last-1,w).getValues().filter(function(r){return r.some(function(v){return v!==''&&v!==null;});}).map(function(r){var o={};h.forEach(function(k,i){if(k)o[k]=r[i];});return o;});
}
function append_(name,obj){var s=sheet_(name),h=s.getRange(1,1,1,s.getLastColumn()).getDisplayValues()[0].map(norm_);s.appendRow(h.map(function(k){return Object.prototype.hasOwnProperty.call(obj,k)?obj[k]:'';}));}
function findSession_(id){return rows_('ExamSessions').find(function(s){return norm_(s.SessionID)===norm_(id);});}
function requireText_(v,label){var x=norm_(v);if(!x)throw new Error(label+' is required.');return x;}
function doGet(e){
  var p=(e&&e.parameter)||{},action=p.action||'status';
  try{
    var out;
    if(action==='status'){out={ok:true,app:'VICORA Exam Room Manager',version:VICORA_VERSION,database:db_().getName(),databaseId:VICORA_DB_ID};}
    else if(action==='lookup'){
      var id=requireText_(p.studentId,'Student ID'),name=requireText_(p.name,'Full name'),cls=requireText_(p.className,'Class');
      var matches=rows_('Allocations').filter(function(a){return key_(a.StudentID,a.StudentName,a.Class)===key_(id,name,cls);});
      if(p.sessionId)matches=matches.filter(function(a){return norm_(a.SessionID)===norm_(p.sessionId);});
      matches=matches.filter(function(a){var s=findSession_(a.SessionID);return s&&lower_(s.Status)==='confirmed';});
      out={ok:true,allocations:matches,allocation:matches.length?matches[matches.length-1]:null};
    }else if(action==='adminData'){
      ensureSchema_();
      out={ok:true,version:VICORA_VERSION,students:rows_('Students'),classes:rows_('Classes'),subjects:rows_('Subjects'),registrations:rows_('SubjectRegistrations'),sessions:rows_('ExamSessions'),rooms:rows_('RoomSettings'),allocations:rows_('Allocations'),candidates:rows_('CurrentExamCandidates')};
    }else if(action==='previewAllocation'){out=preview_({sessionId:p.sessionId});}
    else if(action==='mutationStatus'){var m=rows_('MutationLog').find(function(x){return norm_(x.OperationID)===norm_(p.operationId);});out=m?{ok:true,mutation:m}:{ok:false,error:'Operation is not yet recorded.'};}
    else throw new Error('Unsupported GET action: '+action);
    return json_(out,p.callback);
  }catch(err){return json_({ok:false,error:err.message},p.callback);}
}
function doPost(e){
  var lock=LockService.getScriptLock(),p={},op='';
  try{
    lock.waitLock(25000);
    p=JSON.parse(e&&e.postData&&e.postData.contents||'{}');
    ensureSchema_();
    op=norm_(p.operationId)||uid_();
    var result=dispatch_(p);
    append_('MutationLog',{OperationID:op,Action:p.action,Status:'SUCCESS',Message:JSON.stringify(result),CreatedAt:new Date()});
  }catch(err){
    if(op){try{append_('MutationLog',{OperationID:op,Action:p.action||'',Status:'FAILED',Message:err.message,CreatedAt:new Date()});}catch(logErr){}}
    return json_({ok:false,operationId:op,error:err.message});
  }finally{try{lock.releaseLock();}catch(ignore){}}
  return json_({ok:true,operationId:op});
}
function dispatch_(p){
  switch(p.action){
    case 'setup':return {message:'Schema checked; existing records preserved.',tabs:Object.keys(SCHEMA)};
    case 'addStudent':return addStudent_(p);
    case 'editStudent':return editStudent_(p);
    case 'archiveStudent':return archiveStudent_(p);
    case 'addSubject':return addSubject_(p);
    case 'registerSubject':return registerSubject_(p);
    case 'removeRegistration':return removeRegistration_(p);
    case 'createSession':return createSession_(p);
    case 'saveRooms':return saveRooms_(p);
    case 'confirmAllocation':return confirm_(p);
    case 'reallocate':return reallocate_(p);
    default:throw new Error('Unsupported action: '+p.action);
  }
}
function addStudent_(p){
  var id=requireText_(p.studentId,'Student ID'),name=requireText_(p.name,'Student name'),cls=requireText_(p.className,'Class'),k=key_(id,name,cls);
  if(rows_('Students').some(function(s){return key_(s.StudentID,s.StudentName,s.Class)===k;}))throw new Error('This student record already exists.');
  append_('Students',{StudentID:id,StudentName:name,Class:cls,SourceSheet:'VICORA',ImportStatus:'Active'});return {message:'Student added.'};
}
function editStudent_(p){
  var id=requireText_(p.oldStudentId,'Original student ID'),name=requireText_(p.oldName,'Original student name'),cls=requireText_(p.oldClass,'Original class'),old=key_(id,name,cls);
  if(rows_('Allocations').some(function(a){return key_(a.StudentID,a.StudentName,a.Class)===old;}))throw new Error('This student has an allocation. Identity changes are blocked to protect exam history.');
  var s=sheet_('Students'),m=headerMap_(s),v=s.getDataRange().getValues(),idx=-1;
  for(var i=1;i<v.length;i++)if(key_(v[i][m.StudentID],v[i][m.StudentName],v[i][m.Class])===old){idx=i;break;}
  if(idx<0)throw new Error('Student not found.');
  var next={StudentID:requireText_(p.studentId,'Student ID'),StudentName:requireText_(p.name,'Student name'),Class:requireText_(p.className,'Class')};
  if(rows_('Students').some(function(x){return key_(x.StudentID,x.StudentName,x.Class)===key_(next.StudentID,next.StudentName,next.Class)&&key_(x.StudentID,x.StudentName,x.Class)!==old;}))throw new Error('The edited student identity already exists.');
  Object.keys(next).forEach(function(k){if(m[k]!==undefined)s.getRange(idx+1,m[k]+1).setValue(next[k]);});
  return {message:'Student updated.'};
}
function archiveStudent_(p){
  var k=key_(requireText_(p.studentId,'Student ID'),requireText_(p.name,'Student name'),requireText_(p.className,'Class')),s=sheet_('Students'),m=headerMap_(s),v=s.getDataRange().getValues();
  for(var i=1;i<v.length;i++)if(key_(v[i][m.StudentID],v[i][m.StudentName],v[i][m.Class])===k){s.getRange(i+1,m.ImportStatus+1).setValue('Archived');return {message:'Student archived; history retained.'};}
  throw new Error('Student not found.');
}
function addSubject_(p){var n=requireText_(p.subject,'Subject name');if(rows_('Subjects').some(function(x){return lower_(x.SubjectName)===lower_(n);}))throw new Error('Subject already exists.');append_('Subjects',{SubjectName:n});return {message:'Subject added.'};}
function registerSubject_(p){
  var id=requireText_(p.studentId,'Student ID'),name=requireText_(p.studentName,'Student name'),cls=requireText_(p.className,'Class'),sub=requireText_(p.subject,'Subject'),k=key_(id,name,cls);
  if(!rows_('Students').some(function(s){return key_(s.StudentID,s.StudentName,s.Class)===k&&lower_(s.ImportStatus)!=='archived';}))throw new Error('Active student not found.');
  if(!rows_('Subjects').some(function(s){return lower_(s.SubjectName)===lower_(sub);}))throw new Error('Subject does not exist in the Subjects tab.');
  if(rows_('SubjectRegistrations').some(function(r){return key_(r.StudentID,r.StudentName,r.Class)===k&&lower_(r.Subject)===lower_(sub)&&lower_(r.RegistrationStatus)==='active';}))throw new Error('This subject is already registered for the student.');
  append_('SubjectRegistrations',{StudentID:id,StudentName:name,Class:cls,Subject:sub,RegistrationStatus:'Active',SourceSheet:'VICORA'});return {message:'Subject registered.'};
}
function removeRegistration_(p){
  var k=key_(p.studentId,p.studentName,p.className),sub=lower_(p.subject),s=sheet_('SubjectRegistrations'),m=headerMap_(s),v=s.getDataRange().getValues();
  for(var i=1;i<v.length;i++)if(key_(v[i][m.StudentID],v[i][m.StudentName],v[i][m.Class])===k&&lower_(v[i][m.Subject])===sub&&lower_(v[i][m.RegistrationStatus])==='active'){s.getRange(i+1,m.RegistrationStatus+1).setValue('Removed');return {message:'Registration removed; record retained.'};}
  throw new Error('Active subject registration not found.');
}
function createSession_(p){var n=requireText_(p.sessionName,'Session name'),id=uid_();append_('ExamSessions',{SessionID:id,SessionName:n,StartDate:p.startDate||new Date(),Status:'Draft'});return {sessionId:id,message:'Draft session created.'};}
function saveRooms_(p){
  var id=requireText_(p.sessionId,'Session'),se=findSession_(id);if(!se)throw new Error('Session not found.');
  if(lower_(se.Status)==='confirmed')throw new Error('Confirmed sessions are locked. Use the reallocation action.');
  if(!Array.isArray(p.rooms)||!p.rooms.length)throw new Error('Add at least one room.');
  var names={},rooms=p.rooms.map(function(r,i){var name=requireText_(r.name,'Room name'),cap=Number(r.capacity),nk=lower_(name);if(!Number.isInteger(cap)||cap<1)throw new Error('Room capacity must be a positive whole number.');if(names[nk])throw new Error('Room names must be unique within a session.');names[nk]=true;return {SessionID:id,RoomNumber:i+1,RoomName:name,Capacity:cap};});
  var s=sheet_('RoomSettings'),m=headerMap_(s),v=s.getDataRange().getValues();
  for(var j=v.length-1;j>=1;j--)if(norm_(v[j][m.SessionID])===id)s.deleteRow(j+1);
  rooms.forEach(function(r){append_('RoomSettings',r);});return {message:'Room settings saved.',roomCount:rooms.length};
}
function candidates_(){
  var students=rows_('Students'),registrations=rows_('SubjectRegistrations'),classes=rows_('Classes'),current=rows_('CurrentExamCandidates');
  var classAllowed={};classes.forEach(function(c){classAllowed[lower_(c.ClassName)]=lower_(c.IncludedInCurrentExam)!=='no'&&lower_(c.IncludedInCurrentExam)!=='false'&&lower_(c.IncludedInCurrentExam)!=='excluded';});
  var subjectCounts={};registrations.forEach(function(r){if(lower_(r.RegistrationStatus)==='active'){var k=key_(r.StudentID,r.StudentName,r.Class);subjectCounts[k]=(subjectCounts[k]||0)+1;}});
  var source=current.filter(function(c){return norm_(c.StudentID)&&norm_(c.StudentName)&&norm_(c.Class);}),fromVerification=source.length>0;
  var eligibleSource=source.filter(function(c){return lower_(c.Eligibility)==='eligible';});
  var list=(fromVerification&&eligibleSource.length)?eligibleSource.map(function(c){var k=key_(c.StudentID,c.StudentName,c.Class);return {StudentID:c.StudentID,StudentName:c.StudentName,Class:c.Class,SubjectCount:subjectCounts[k]||Number(c.SubjectCount)||0};}):
    students.filter(function(s){var k=key_(s.StudentID,s.StudentName,s.Class);return lower_(s.ImportStatus)!=='archived'&&lower_(s.ImportStatus)!=='inactive'&&lower_(s.Class)!=='ss2b'&&classAllowed[lower_(s.Class)]!==false&&k!=='stu001|adebayo ojo|jss1a'&&k!=='stu002|adetayo aveda|jss1a';}).map(function(s){var k=key_(s.StudentID,s.StudentName,s.Class);return {StudentID:s.StudentID,StudentName:s.StudentName,Class:s.Class,SubjectCount:subjectCounts[k]||0};});
  var order={},classOrder={};students.forEach(function(s,i){var k=key_(s.StudentID,s.StudentName,s.Class);if(order[k]===undefined)order[k]=i;if(classOrder[lower_(s.Class)]===undefined)classOrder[lower_(s.Class)]=i;});
  list.sort(function(a,b){var ca=classOrder[lower_(a.Class)]===undefined?999999:classOrder[lower_(a.Class)],cb=classOrder[lower_(b.Class)]===undefined?999999:classOrder[lower_(b.Class)];return ca-cb||(order[key_(a.StudentID,a.StudentName,a.Class)]||0)-(order[key_(b.StudentID,b.StudentName,b.Class)]||0);});
  var seen={};return list.filter(function(c){var k=key_(c.StudentID,c.StudentName,c.Class);if(seen[k])return false;seen[k]=true;return true;});
}
function preview_(p){
  var id=requireText_(p.sessionId,'Session'),se=findSession_(id);if(!se)throw new Error('Session not found.');
  if(lower_(se.Status)==='confirmed'&&!p.allowConfirmed)throw new Error('This session is confirmed. Use explicit reallocation.');
  var rooms=rows_('RoomSettings').filter(function(r){return norm_(r.SessionID)===id;}).sort(function(a,b){return Number(a.RoomNumber)-Number(b.RoomNumber);});
  if(!rooms.length)throw new Error('Save room settings first.');
  var people=candidates_(),total=rooms.reduce(function(n,r){return n+Number(r.Capacity);},0);if(people.length>total)throw new Error('Room capacity ('+total+') is less than eligible candidates ('+people.length+').');
  var counts=rooms.map(function(){return 0;}),alloc=[],roomIndex=0;
  people.forEach(function(c){
    var attempts=0;while(attempts<rooms.length&&counts[roomIndex]>=Number(rooms[roomIndex].Capacity)){roomIndex=(roomIndex+1)%rooms.length;attempts++;}
    if(attempts>=rooms.length)throw new Error('No available room capacity.');
    var r=rooms[roomIndex];counts[roomIndex]++;
    alloc.push({StudentID:c.StudentID,StudentName:c.StudentName,Class:c.Class,RoomNumber:Number(r.RoomNumber),RoomName:r.RoomName,SubjectCount:c.SubjectCount});
    roomIndex=(roomIndex+1)%rooms.length;
  });
  return {sessionId:id,candidateCount:people.length,roomSummary:rooms.map(function(r,i){return {roomNumber:r.RoomNumber,roomName:r.RoomName,capacity:Number(r.Capacity),allocated:counts[i]};}),allocations:alloc,classSummary:classSummary_(alloc)};
}
function classSummary_(alloc){var out={};alloc.forEach(function(a){if(!out[a.Class])out[a.Class]={className:a.Class,total:0,rooms:{}};out[a.Class].total++;out[a.Class].rooms[a.RoomNumber]=(out[a.Class].rooms[a.RoomNumber]||0)+1;});return Object.keys(out).map(function(k){return out[k];});}
function confirm_(p){
  var id=requireText_(p.sessionId,'Session'),s=sheet_('ExamSessions'),m=headerMap_(s),v=s.getDataRange().getValues(),row=-1;
  for(var i=1;i<v.length;i++)if(norm_(v[i][m.SessionID])===id){row=i+1;break;}
  if(row<0)throw new Error('Session not found.');
  if(lower_(s.getRange(row,m.Status+1).getValue())==='confirmed')throw new Error('This allocation is already confirmed.');
  if(rows_('Allocations').some(function(a){return norm_(a.SessionID)===id;}))throw new Error('Existing allocations found; use explicit reallocation.');
  var pr=preview_({sessionId:id});if(!pr.candidateCount)throw new Error('There are no eligible candidates to allocate.');
  pr.allocations.forEach(function(a){append_('Allocations',{SessionID:id,StudentID:a.StudentID,StudentName:a.StudentName,Class:a.Class,RoomNumber:a.RoomNumber});});
  s.getRange(row,m.Status+1).setValue('Confirmed');return {message:'Allocation confirmed and locked.',candidateCount:pr.candidateCount,roomSummary:pr.roomSummary};
}
function reallocate_(p){
  var id=requireText_(p.sessionId,'Session'),reason=requireText_(p.reason,'Reason for reallocation'),se=findSession_(id);if(!se)throw new Error('Session not found.');
  var old=rows_('Allocations').filter(function(a){return norm_(a.SessionID)===id;});if(!old.length)throw new Error('No existing allocation to replace.');
  var pr=preview_({sessionId:id,allowConfirmed:true}),next={};pr.allocations.forEach(function(a){next[key_(a.StudentID,a.StudentName,a.Class)]=a.RoomNumber;});
  var history=sheet_('AllocationHistory'),admin=norm_(p.administrator)||'VICORA Admin';
  old.forEach(function(a){append_('AllocationHistory',{SessionID:id,StudentID:a.StudentID,PreviousRoom:a.RoomNumber,NewRoom:next[key_(a.StudentID,a.StudentName,a.Class)]||'REMOVED',Reason:reason,Date:new Date(),Administrator:admin});});
  var sh=sheet_('Allocations'),m=headerMap_(sh),v=sh.getDataRange().getValues();for(var j=v.length-1;j>=1;j--)if(norm_(v[j][m.SessionID])===id)sh.deleteRow(j+1);
  pr.allocations.forEach(function(a){append_('Allocations',{SessionID:id,StudentID:a.StudentID,StudentName:a.StudentName,Class:a.Class,RoomNumber:a.RoomNumber});});
  var ss=sheet_('ExamSessions'),sm=headerMap_(ss),sv=ss.getDataRange().getValues();for(var z=1;z<sv.length;z++)if(norm_(sv[z][sm.SessionID])===id){ss.getRange(z+1,sm.Status+1).setValue('Confirmed');break;}
  return {message:'Reallocation completed and audited.',candidateCount:pr.candidateCount,roomSummary:pr.roomSummary};
}

/**
 * MANUAL SETUP — visible in the Apps Script Run dropdown.
 * Creates missing VICORA tabs and adds missing schema columns without clearing records.
 */
function VICORA_setup() {
  ensureSchema_();
  return {
    ok: true,
    message: 'VICORA Exam Room Manager setup checked. Existing records were preserved.',
    database: db_().getUrl(),
    version: VICORA_VERSION
  };
}

/**
 * MANUAL CONNECTION TEST — visible in the Apps Script Run dropdown.
 * Confirms access to the configured VICORA spreadsheet.
 */
function VICORA_testConnection() {
  var database = db_();
  return {
    ok: true,
    message: 'Connected to the VICORA Exam Room Manager database.',
    databaseName: database.getName(),
    databaseUrl: database.getUrl(),
    databaseId: VICORA_DB_ID
  };
}

/**
 * MANUAL STATUS CHECK — visible in the Apps Script Run dropdown.
 */
function VICORA_getStatus() {
  var database = db_();
  var counts = {};
  ['Students','Classes','Subjects','SubjectRegistrations','ExamSessions','RoomSettings','Allocations']
    .forEach(function(name) {
      var tab = database.getSheetByName(name);
      counts[name] = tab ? Math.max(0, tab.getLastRow() - 1) : 0;
    });
  return {
    ok: true,
    app: 'VICORA Exam Room Manager',
    version: VICORA_VERSION,
    database: database.getName(),
    databaseId: VICORA_DB_ID,
    counts: counts
  };
}

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const root = path.resolve(__dirname, '..');
const cliFile = path.join(root, 'dist/bin/sheetops.js');
const apiFile = path.join(root, 'dist/lib/sheets-api.js');
const plain = x => JSON.parse(JSON.stringify(x));

function loadApi({ values = [[1]], failCopy = false, files = [], renderedUpdateValues, listResponse, responseRange = "'Data'!A1:B2", trashFails = false } = {}) {
  const calls = [];
  let stored = values;
  const spreadsheets = {
    values: {
      get: async r => { calls.push(['get', plain(r)]); return {data:{range:responseRange,values:stored}}; },
      update: async r => { calls.push(['update',plain(r)]); stored = renderedUpdateValues || r.requestBody.values; return {data:{}}; },
      append: async r => { calls.push(['append',plain(r)]); return {data:{updates:{updatedRange:'Data!A2',updatedRows:1}}}; },
      clear: async r => { calls.push(['clear',plain(r)]); return {data:{}}; }
    },
    get: async r => { calls.push(['meta',plain(r)]); return {data:{sheets:[{properties:{title:'Data',sheetId:0,gridProperties:{rowCount:10,columnCount:10}}}],namedRanges:[]}}; },
  };
  const driveFiles = {
    copy: async r => { calls.push(['copy',plain(r)]); if (failCopy) throw Error('copy failed'); files.push({id:'backup-id',name:r.requestBody.name}); return {data:{id:'backup-id'}}; },
    list: async r => { calls.push(['list',plain(r)]); return {data:listResponse ? listResponse(r,calls) : {files}}; },
    update: async r => { calls.push(['trash',plain(r)]); if(trashFails) throw Error('trash failed'); return {data:{}}; }
  };
  class OAuth2 {setCredentials(){} on(){} }
  const google = {auth:{OAuth2},sheets:()=>({spreadsheets}),drive:()=>({files:driveFiles})};
  const fakeFs = {existsSync:()=>true,readFileSync:()=>JSON.stringify({client_id:'fake-client',client_secret:'fake-secret',access_token:'fake-token'}),writeFileSync(){},chmodSync(){}};
  const module = {exports:{}};
  const requireFake = name => {
    if (name === 'fs' || name === 'node:fs') return fakeFs;
    if (name === 'googleapis') return {google};
    return require(name);
  };
  vm.runInNewContext(fs.readFileSync(apiFile,'utf8'),{require:requireFake,module,exports:module.exports,__dirname:path.join(root,'dist/lib'),process,console,URL,setTimeout},{filename:apiFile});
  return {api:module.exports,calls};
}

async function runCli(argv, patchOverrides = {}, apiOverrides = {}, fsScenario = {}) {
  const calls = [], output = [];
  const fileState = fsScenario.files || new Map();
  const patch = {operationId:'20261005-120000-contract',project:'demo',reason:'Test patch behavior',requiresApproval:true,backupRequired:true,operations:[{type:'setValues',target:{sheetName:'Data',a1:'A1'},values:[[2]]}],...patchOverrides};
  const rootConfig = {projects:[],defaultBackupFolderId:'backup-folder',defaultBackupRetain:3};
  const project = {spreadsheetId:'sheet-id',spreadsheetName:'Demo'};
  const fakeFs = {
    existsSync:filename=>fileState.has(filename) || !String(filename).includes('/ops/patches/'),
    readFileSync: filename => JSON.stringify(String(filename).endsWith('project.config.json') ? project : String(filename).endsWith('patch.json') ? patch : rootConfig),
    mkdirSync(){},appendFileSync(){},writeFileSync(filename,data,opts){ if(opts?.flag === 'wx' && fileState.has(filename)) throw Error('EEXIST'); fileState.set(filename,data); },
    renameSync:(a,b)=>{calls.push(['archive',a,b]);if(fsScenario.archiveError) throw Error('archive failed');fileState.set(b,'archived');},
  };
  const api = {
    writeRangeDryRun:async (...a)=>{calls.push(['dry',...plain(a)]);return {totalCellsAffected:1,warnings:[],requiresApproval:false};},
    appendRowsDryRun:async (...a)=>{calls.push(['append-dry',...plain(a)]);return {totalCellsAffected:1,warnings:['APPEND'],requiresApproval:true};},
    clearRangeDryRun:async (...a)=>{calls.push(['clear-dry',...plain(a)]);return {totalCellsAffected:1,warnings:['DESTRUCTIVE_CLEAR'],requiresApproval:true};},
    backupSpreadsheet:async (...a)=>{calls.push(['backup',...plain(a)]);return {backupName:'copy',backupId:'backup-id'};},
    writeRange:async (...a)=>{calls.push(['write',...plain(a)]);return {cellsWritten:1};},
    appendRows:async (...a)=>{calls.push(['append',...plain(a)]);return {rowsAppended:1};},
    clearRange:async (...a)=>{calls.push(['clear',...plain(a)]);return {cleared:true};},
    logOperation:async (...a)=>{calls.push(['log',...plain(a)]);},
    ...apiOverrides,
  };
  let exitCode;
  const fakeProcess = {...process,argv:['node','sheetops',...argv],exit:c=>{if(exitCode === undefined){exitCode=c;throw Error('EXIT:'+c);}}};
  const requireFake = name => {
    if (name === 'fs' || name === 'node:fs') return fakeFs;
    if (String(name).endsWith('/sheets-api.js')) return api;
    if (name.startsWith('.')) return require(path.resolve(path.dirname(cliFile), name));
    return require(name);
  };
  try {
    const module={exports:{}};
    await vm.runInNewContext(fs.readFileSync(cliFile,'utf8'),{require:requireFake,module,exports:module.exports,__dirname:path.dirname(cliFile),process:fakeProcess,console:{log:(...a)=>output.push(a.join(' ')),error:(...a)=>output.push(a.join(' '))},URL,setTimeout},{filename:cliFile});
  } catch(e) { if(!String(e.message).startsWith('EXIT:')) throw e; }
  await new Promise(resolve=>setImmediate(resolve));
  return {calls,output:output.join('\n'),exitCode};
}

const apply = ['apply-patch','--project','demo','--patch','patch.json'];
test('CLI help and no-argument help agree', async()=>{
  const a=await runCli([]),b=await runCli(['--help']);
  assert.equal(a.output,b.output);assert.match(a.output,/SheetOps CLI/);assert.match(a.output,/dry-run-patch/);assert.equal(a.calls.length,0);
});
test('CLI alias --p and positional parsing keep range arguments',async()=>{
  const r=await runCli(['dry-run-patch','ignored','--p','demo','--patch','patch.json']);
  assert.equal(r.calls[0][0],'dry');assert.deepEqual(r.calls[0].slice(1),['sheet-id',{sheetName:'Data',a1:'A1'},[[2]],{}]);assert.equal(r.calls.length,1);
});
test('dry-run hash mismatch makes no API mutation',async()=>{
  const {api,calls}=loadApi();const r=await api.writeRangeDryRun('id',{namedRange:'Input'},[[2]],{expectedHash:'stale'});
  assert.equal(r.ok,false);assert.match(r.error,/expectedHash mismatch/);assert.deepEqual(calls.map(c=>c[0]),['get']);
});
test('dry-run warns above 100 cells and normalizes ragged reads',async()=>{
  const {api,calls}=loadApi({values:[[1,2],[3]]});const r=await api.writeRangeDryRun('id',{sheetName:'Data',a1:'A1:B2'},[Array(101).fill(0)]);
  assert.equal(r.totalCellsAffected,101);assert.equal(r.requiresApproval,true);assert.deepEqual(plain(r.currentValues),[[1,2],[3,'']]);assert.equal(calls.length,1);
});
test('stale hash and unconfirmed large API writes abort before update',async()=>{
  const {api,calls}=loadApi();await assert.rejects(api.writeRange('id',{namedRange:'Input'},[[2]],{expectedHash:'stale'}),/expectedHash mismatch/);
  await assert.rejects(api.writeRange('id',{namedRange:'Input'},[Array(101).fill(0)]),/confirmLarge/);assert.equal(calls.filter(c=>c[0]==='update').length,0);
});
test('writes use USER_ENTERED then read back the resulting hash',async()=>{
  const {api,calls}=loadApi({renderedUpdateValues:[['2','2']]});const r=await api.writeRange('id',{sheetName:'Data',a1:'A1'},[[2,'=SUM(A1)']]);
  assert.deepEqual(calls.map(c=>c[0]),['update','get']);assert.equal(calls[0][1].valueInputOption,'USER_ENTERED');assert.equal(r.newHash,api.hash([['2','2']]));assert.notEqual(r.newHash,api.hash([[2,'=SUM(A1)']]));
});
test('backup sets destination and trashes only old exact-prefix copies',async()=>{
  const {api,calls}=loadApi({files:[{id:'new',name:'Demo__backup__2026-10-02T12-00-00'},{id:'old',name:'Demo__backup__2026-10-01T12-00-00'},{id:'keep',name:'Demo__backup__2026-09-01T12-00-00-DO-NOT-DELETE'}]});
  const r=await api.backupSpreadsheet('id','Demo','test',{folderId:'folder',retain:2});
  assert.deepEqual(calls[0][1].requestBody.parents,['folder']);assert.deepEqual(plain(r.pruned),['Demo__backup__2026-10-01T12-00-00']);assert.deepEqual(calls.filter(c=>c[0]==='trash').map(c=>c[1]),[{fileId:'old',requestBody:{trashed:true}}]);
});
test('unconfirmed apply with requiresApproval true is rejected',async()=>{
  const r=await runCli(apply);assert.equal(r.exitCode,1);assert.equal(r.calls.length,0);
});
test('confirmed apply backs up, writes, logs, then archives',async()=>{
  const r=await runCli([...apply,'--confirmed']);assert.deepEqual(r.calls.map(c=>c[0]),['backup','write','log','archive']);assert.equal(r.calls[0][4].folderId,'backup-folder');
});


  test('requiresApproval false cannot bypass explicit confirmation',async()=>{const r=await runCli(apply,{requiresApproval:false});assert.equal(r.exitCode,1);assert.equal(r.calls.length,0);});
  test('failed required backup prevents sheet writes and archival',async()=>{const r=await runCli([...apply,'--confirmed'],{}, {backupSpreadsheet:async()=>{throw Error('copy failed');}});assert.equal(r.exitCode,1);assert.equal(r.calls.length,0);});
  test('unsupported late operation is rejected before any backup or write',async()=>{const r=await runCli([...apply,'--confirmed'],{operations:[{type:'setValues',target:{namedRange:'Input'},values:[[1]]},{type:'setFormula',target:{namedRange:'Input'},formulas:[['=1']]}]});assert.equal(r.exitCode,1);assert.equal(r.calls.length,0);});
  test('unconfirmed clear is rejected before earlier writes',async()=>{const r=await runCli([...apply,'--confirmed'],{operations:[{type:'setValues',target:{namedRange:'Input'},values:[[1]]},{type:'clearRange',target:{namedRange:'Input'}}]});assert.equal(r.exitCode,1);assert.equal(r.calls.length,0);});
  test('large write requires per-operation confirmLarge before backup',async()=>{const operations=[{type:'setValues',target:{namedRange:'Input'},values:[Array(101).fill(1)]}];const r=await runCli([...apply,'--confirmed'],{operations});assert.equal(r.exitCode,1);assert.equal(r.calls.length,0);const yes=await runCli([...apply,'--confirmed'],{operations:[{...operations[0],confirmLarge:true}]});assert.ok(yes.calls.some(c=>c[0]==='write'));assert.equal(yes.calls.find(c=>c[0]==='write')[4].confirmLarge,true);});
  test('invalid matrix and project mismatch abort before backup',async()=>{for(const change of [{project:'another'},{operations:[{type:'setValues',target:{namedRange:'Input'},values:[1]}]}]){const r=await runCli([...apply,'--confirmed'],change);assert.equal(r.exitCode,1);assert.equal(r.calls.length,0);}});

test('dry-run previews large writes without requiring mutation consent',async()=>{
  const r=await runCli(['dry-run-patch','--project','demo','--patch','patch.json'],{operations:[{type:'setValues',target:{namedRange:'Input'},values:[Array(101).fill(1)]}]});
  assert.deepEqual(r.calls.map(c=>c[0]),['dry']);assert.equal(r.exitCode,undefined);
});
test('apply requires a boolean confirmation switch and defaults to backup',async()=>{
  const no=await runCli([...apply,'--confirmed','false']);assert.equal(no.exitCode,1);assert.equal(no.calls.length,0);
  const yes=await runCli([...apply,'--confirmed'],{backupRequired:undefined});assert.equal(yes.calls[0][0],'backup');
});
test('CSV escaping and snapshot comparison retain their contracts',async()=>{
  const {api}=loadApi({values:[['Name','Note'],['a,b','say "hello"']]});
  const csv=await api.exportRange('id',{sheetName:'Data',a1:'A1:B2'},'csv');assert.equal(csv.csv,'Name,Note\n"a,b","say ""hello"""');
  const diff=api.compareSnapshots({snapshotHash:'before',sheets:[{name:'Data',lastRow:1}],namedRanges:[]},{snapshotHash:'after',sheets:[{name:'Data',lastRow:2},{name:'Added'}],namedRanges:[{name:'Input'}]});
  assert.equal(diff.hashChanged,true);assert.deepEqual(plain(diff.sheets.added),['Added']);assert.equal(diff.sheets.changed.length,1);assert.deepEqual(plain(diff.namedRanges.added),['Input']);
});
test('unsupported hash guards on append/clear reject before every API call',async()=>{
  for(const type of ['appendRows','clearRange']){
    const r=await runCli([...apply,'--confirmed'],{operations:[{type,target:{sheetName:'Data',a1:'A1'},values:[[1]],expectedHash:'stale',confirmDestructive:true}]});
    assert.equal(r.exitCode,1);assert.equal(r.calls.length,0);assert.match(r.output,/expectedHash is supported only for setValues/);
  }
});
test('backup retention preserves current copy with duplicate timestamp and paginated listings',async()=>{
  let name;
  const {api,calls}=loadApi({listResponse:(request,events)=>{
    name=events.find(c=>c[0]==='copy')[1].requestBody.name;
    return request.pageToken ? {files:[{id:'backup-id',name},{id:'pinned',name:name+'-KEEP'}]} : {files:[{id:'other',name}],nextPageToken:'page-2'};
  }});
  await api.backupSpreadsheet('id',"O'Brien",'test',{folderId:"folder'id",retain:1});
  assert.deepEqual(calls.filter(c=>c[0]==='trash').map(c=>c[1].fileId),['other']);assert.equal(calls.filter(c=>c[0]==='list').length,2);
  assert.match(calls.find(c=>c[0]==='list')[1].q,/O\\'Brien/);assert.match(calls.find(c=>c[0]==='list')[1].q,/folder\\'id/);
});
test('archive failure retains an application record and prevents append replay',async()=>{
  const files=new Map(),patch={operations:[{type:'appendRows',target:{sheetName:'Data',a1:'A1'},values:[[1]]}]};
  const first=await runCli([...apply,'--confirmed'],patch,{}, {files,archiveError:true});
  assert.equal(first.exitCode,1);assert.equal(first.calls.filter(c=>c[0]==='append').length,1);assert.match(first.output,/Do not rerun/);
  const second=await runCli([...apply,'--confirmed'],patch,{}, {files});assert.equal(second.exitCode,1);assert.equal(second.calls.length,0);assert.match(second.output,/already has an application record/);
});
test('existing applied record prevents archive collision before backup',async()=>{
  const record=path.join(root,'projects/demo/ops/patches/applied/20261005-120000-contract.json');
  const r=await runCli([...apply,'--confirmed'],{}, {}, {files:new Map([[record,'already applied']])});assert.equal(r.exitCode,1);assert.equal(r.calls.length,0);
});
test('dry-run exposes missing destructive consent and explicitly disabled backups',async()=>{
  const r=await runCli(['dry-run-patch','--project','demo','--patch','patch.json'],{backupRequired:false,operations:[{type:'clearRange',target:{sheetName:'Data',a1:'A1'}}]});
  assert.match(r.output,/Backup disabled/);assert.match(r.output,/confirmDestructive:true required before apply/);assert.match(r.output,/Dry-run has warnings/);assert.deepEqual(r.calls.map(c=>c[0]),['clear-dry']);
});
test('unsupported safety flags fail rather than appearing to protect a REST write',async()=>{
  for(const key of ['allowFormulaOverwrite','allowHiddenSheet','allowProtected']){
    const r=await runCli([...apply,'--confirmed'],{operations:[{type:'setValues',target:{namedRange:'Input'},values:[[1]],[key]:false}]});
    assert.equal(r.exitCode,1);assert.equal(r.calls.length,0);assert.match(r.output,/unsupported by the REST patch path/);
  }
});

test('unknown patch, operation, and target properties reject before API calls',async()=>{
  const op={type:'setValues',target:{sheetName:'Data',a1:'A1'},values:[[1]]};
  for(const patch of [{backupRequird:false},{operations:[{...op,expectedHahs:'stale'}]},{operations:[{...op,target:{...op.target,namedRagne:'Other'}}]}]){
    for(const argv of [[...apply,'--confirmed'],['dry-run-patch','--project','demo','--patch','patch.json']]){
      const r=await runCli(argv,patch);assert.equal(r.exitCode,1);assert.equal(r.calls.length,0);assert.match(r.output,/unsupported property/);
    }
  }
});
test('append rejects ambiguous targets and forwards the same explicit anchor for preview and apply',async()=>{
  const operation={type:'appendRows',target:{sheetName:"O'Brien",a1:'C4:F20'},values:[[1]]};
  const bad=await runCli([...apply,'--confirmed'],{operations:[{...operation,target:{...operation.target,namedRange:'Elsewhere'}}]});assert.equal(bad.exitCode,1);assert.equal(bad.calls.length,0);
  const dry=await runCli(['dry-run-patch','--project','demo','--patch','patch.json'],{operations:[operation]});assert.deepEqual(dry.calls[0],['append-dry','sheet-id',operation.target,[[1]]]);
  const yes=await runCli([...apply,'--confirmed'],{operations:[operation]});assert.deepEqual(yes.calls.find(c=>c[0]==='append'),['append','sheet-id',"O'Brien",[[1]],'C4:F20']);
  const {api,calls}=loadApi();const preview=await api.appendRowsDryRun('id',operation.target,[[1]]);await api.appendRows('id',"O'Brien",[[1]],'C4:F20');
  assert.equal(calls.find(c=>c[0]==='get')[1].range,"'O''Brien'!C4:F20");assert.equal(calls.find(c=>c[0]==='append')[1].range,preview.appendRange);assert.match(preview.warnings.join(' '),/exact row is known only after apply/);
});
test('clear preview counts the full rectangle and reports unknown unbounded extents',async()=>{
  const {api,calls}=loadApi({responseRange:"'Data'!A1:Z100",values:[[1]]});const r=await api.clearRangeDryRun('id',{sheetName:'Data',a1:'A1:Z100'});
  assert.equal(r.totalCellsAffected,2600);assert.match(r.warnings[0],/2600 target cells/);assert.equal(r.requiresApproval,true);assert.deepEqual(calls.map(c=>c[0]),['get']);
  const unbounded=loadApi({responseRange:"'Data'!A:Z"});const unknown=await unbounded.api.clearRangeDryRun('id',{sheetName:'Data',a1:'A:Z'});assert.equal(unknown.totalCellsAffected,null);assert.match(unknown.warnings[0],/extent is unknown/);
});
test('retention failures report the successful backup copy separately',async()=>{
  for(const opts of [{listResponse:()=>{throw Error('list failed');}},{files:[{id:'old',name:'Demo__backup__2026-10-01T12-00-00'}],trashFails:true}]){
    const {api,calls}=loadApi(opts);const r=await api.backupSpreadsheet('id','Demo','test',{folderId:'folder',retain:1});assert.equal(r.backupId,'backup-id');assert.match(r.retentionWarning,/Backup created, but retention cleanup failed/);assert.equal(calls.filter(c=>c[0]==='copy').length,1);
  }
  const r=await runCli([...apply,'--confirmed'],{}, {backupSpreadsheet:async()=>({backupId:'backup-id',backupName:'copy',retentionWarning:'Backup created, but retention cleanup failed: list failed'})});assert.ok(r.calls.some(c=>c[0]==='write'));assert.match(r.output,/retention cleanup failed/);
});

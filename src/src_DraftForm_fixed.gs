//DraftForm.gs

/**
 * Splatoon3 ドラフト会議 自動フォーム & サブラウンド抽選ロジック（バグ修正版）
 * Google スプレッドシート + Google フォーム + Apps Script
 *
 * ▼修正内容
 * 1. admin_buildAudienceSheet_streamGrid()の枠3の式を修正（INDEXの第2引数が2→3）
 * 2. isPickAlreadyRecorded_()の比較ロジックを修正（r[1]同士の比較→captainとの比較）
 * 3. resolveIfReady_()とonFormSubmit()で空のplayerを除外する処理を追加
 * 4. draft_getCurrentContests()で空のplayerを除外する処理を追加
 * 5. resolveIfReady_()で既に選ばれたプレイヤーを除外する処理を追加
 * 6. resolveIfReady_()とdraft_getCurrentContests()で同じキャプテンの重複提出を最新のみ使用する処理を追加
 * 7. draft_rollOne()の冗長な条件分岐を削除
 */

// ===== 設定 =====
const CFG = {
  SHEET: {
    CONFIG: 'Config',
    PLAYERS: 'Players',
    CAPTAINS: 'Captains',
    PICKS: 'Picks',
    TEAMS: 'Teams',
    REQUESTS: 'Requests',
    LOG: 'Log',
  },
  CFG_KEYS: {
    FORM_ID: 'FORM_ID',
    ROUND: 'ROUND',
    IS_OPEN: 'IS_OPEN',
    SUBROUND: 'SUBROUND',
    ELIGIBLE_CAPTAINS: 'ELIGIBLE_CAPTAINS_JSON',
    MANUAL_MODE: 'MANUAL_MODE',
    LAST_UPDATE_TIMESTAMP: 'LAST_UPDATE_TIMESTAMP',
    WEB_APP_URL: 'WEB_APP_URL',
    SETUP_COMPLETE: 'SETUP_COMPLETE', // ドラフト準備完了フラグ
  },

  FORM: {
    TITLE_PREFIX: 'Splatoon3 ドラフト会議',
    Q_CAPTAIN: 'キャプテン',
    Q_PLAYER: '指名する選手',
    Q_ROUND: 'Round',
    Q_SUB: 'Sub',
    DESC: '同時指名ラウンド。送信は取り消せません。',
  }
};

// =====　カスタムメニュー追加　======
function admin_openControlCenter(){
  const html = HtmlService.createHtmlOutputFromFile('src_ControlCenter')
    .setTitle('ドラフト運営センター')
    .setWidth(420);
  SpreadsheetApp.getUi().showSidebar(html);
}

function admin_openControlCenterDialog(){
  const html = HtmlService.createHtmlOutputFromFile('src_ControlCenter')
    .setTitle('ドラフト運営センター')
    .setWidth(800)
    .setHeight(600);
  // モーダルダイアログとして表示（中央表示、サイズ調整可能）
  SpreadsheetApp.getUi().showModalDialog(html, 'ドラフト運営センター');
}


function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('ドラフト運営')
    .addItem('運営センターを開く', 'admin_openControlCenter')
    .addSeparator()
    .addItem('ドラフト準備（フォーム作成＆トリガー再作成）', 'menu_setupDraft')
    .addSeparator()
    .addItem('Audienceを再構築', 'menu_buildAudience')
    .addSeparator()
    .addItem('全データリセット', 'admin_resetAllSafe')
    .addToUi();
}

function menu_setupDraft() {
  try {
    const result = setupDraft();
    // アラート表示は削除（「処理しています...」表示が消えない問題を回避）
    // 結果はLogシートに記録される
    logEvent('MENU_SETUP_DRAFT', result.isNewForm 
      ? `ドラフト準備完了（新規フォーム作成、キャプテン: ${result.captainsCount}名、プレイヤー: ${result.playersCount}名）`
      : `ドラフト準備完了（既存フォーム更新、キャプテン: ${result.captainsCount}名、プレイヤー: ${result.playersCount}名）`);
    // タイムスタンプ更新はsetupDraft_()内のputConfig()で自動的に実行される
  } catch(e) {
    // エラーもLogシートに記録
    logEvent('MENU_SETUP_DRAFT_ERROR', `ドラフト準備失敗: ${e}`);
  }
}

function menu_buildAudience() {
  admin_buildAudienceSheet_streamGrid();
  admin_polishAudienceLook_Final();
}

function buildRouletteArrayEqual_(names) {
  return names.slice();
}

// ===== ユーティリティ =====
const _ss = () => SpreadsheetApp.getActive();
const _sh = (name) => _ss().getSheetByName(name);

// 【追加】キャッシュサービス（30秒の有効期限）
const CACHE = CacheService.getScriptCache();
const CACHE_TTL = 30; // 30秒

// 【追加】キャッシュ付きデータ取得のヘルパー関数
function getCached_(key, fetchFn, ttl = CACHE_TTL) {
  const cached = CACHE.get(key);
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {
      // キャッシュのパースに失敗した場合は無視
    }
  }
  const data = fetchFn();
  try {
    CACHE.put(key, JSON.stringify(data), ttl);
  } catch (e) {
    // キャッシュの保存に失敗した場合は無視（データは返す）
  }
  return data;
}

// 【追加】キャッシュを無効化する関数
function invalidateCache_(keys) {
  if (Array.isArray(keys)) {
    keys.forEach(key => CACHE.remove(key));
  } else {
    CACHE.remove(keys);
  }
}

function ensureSheet_(name, headers) {
  let sh = _sh(name);
  if (!sh) sh = _ss().insertSheet(name);
  if (headers && sh.getLastRow() === 0) sh.appendRow(headers);
  return sh;
}

function ensureConfigSheet_() { return ensureSheet_(CFG.SHEET.CONFIG, ['Key','Value']); }

function getConfigMap() {
  return getCached_('config_map', () => {
    const sh = ensureConfigSheet_();
    const last = sh.getLastRow();
    const map = {};
    if (last >= 2) {
      const vs = sh.getRange(2,1,last-1,2).getValues();
      for (let i=0;i<vs.length;i++){ const k=vs[i][0]; if(k!==''&&k!=null) map[k]=vs[i][1]; }
    }
    return map;
  });
}
function putConfig(key, val) {
  const sh = ensureConfigSheet_();
  const last = sh.getLastRow();
  if (last >= 2) {
    const vs = sh.getRange(2,1,last-1,2).getValues();
    for (let i=0;i<vs.length;i++){ if (vs[i][0]===key){ sh.getRange(2+i,2).setValue(val); 
      // 【追加】キャッシュを無効化
      invalidateCache_('config_map');
      // UI更新をトリガーする重要な設定が変更された場合、タイムスタンプを自動更新
      const triggerKeys = [
        CFG.CFG_KEYS.ROUND,
        CFG.CFG_KEYS.SUBROUND,
        CFG.CFG_KEYS.IS_OPEN,
        CFG.CFG_KEYS.SETUP_COMPLETE,
        CFG.CFG_KEYS.ELIGIBLE_CAPTAINS,
        CFG.CFG_KEYS.MANUAL_MODE
      ];
      if (triggerKeys.includes(key) && key !== CFG.CFG_KEYS.LAST_UPDATE_TIMESTAMP) {
        updateDataTimestamp_();
      }
      return; 
    } }
  }
  sh.appendRow([key,val]);
  // 【追加】キャッシュを無効化
  invalidateCache_('config_map');
  // UI更新をトリガーする重要な設定が変更された場合、タイムスタンプを自動更新
  const triggerKeys = [
    CFG.CFG_KEYS.ROUND,
    CFG.CFG_KEYS.SUBROUND,
    CFG.CFG_KEYS.IS_OPEN,
    CFG.CFG_KEYS.SETUP_COMPLETE,
    CFG.CFG_KEYS.ELIGIBLE_CAPTAINS,
    CFG.CFG_KEYS.MANUAL_MODE
  ];
  if (triggerKeys.includes(key) && key !== CFG.CFG_KEYS.LAST_UPDATE_TIMESTAMP) {
    updateDataTimestamp_();
  }
}
function logEvent(event, detail){ ensureSheet_(CFG.SHEET.LOG,['Timestamp','Event','Detail']).appendRow([new Date(),event,detail||'']); }

function readCaptains(){
  return getCached_('captains', () => {
    const sh=ensureSheet_(CFG.SHEET.CAPTAINS,['Captain','Team','Prio','XP']);
    const vs=sh.getDataRange().getValues(); const rs=[];
    for(let i=1;i<vs.length;i++){ const name=vs[i][0]; if(!name) continue; rs.push({name,team:vs[i][1]||name,prio:Number(vs[i][2]||0),xp:Number(vs[i][3]||0)}); }
    return rs;
  });
}
function writeCaptains(caps){ 
  const sh=ensureSheet_(CFG.SHEET.CAPTAINS,['Captain','Team','Prio','XP']); 
  const lastRow = sh.getLastRow();
  const existingRows = lastRow > 1 ? lastRow - 1 : 0;
  
  // 【修正】clearContent()とappendRow()の組み合わせでは下にずれるため、setValues()で上書きする
  if (caps.length > 0) {
    // 既存のデータを上書き
    if (existingRows > 0) {
      // 既存の行数が新しいデータより多い場合は、余分な行をクリア
      if (existingRows > caps.length) {
        sh.getRange(2 + caps.length, 1, existingRows - caps.length, 4).clearContent();
      }
      // 新しいデータを既存の行に上書き
      const rowsToWrite = Math.min(existingRows, caps.length);
      if (rowsToWrite > 0) {
        const values = caps.slice(0, rowsToWrite).map(c => [c.name, c.team, c.prio || 0, c.xp || 0]);
        sh.getRange(2, 1, rowsToWrite, 4).setValues(values);
      }
    }
    
    // 【最適化】新しいデータが既存の行数より多い場合は、setValues()で一度に追加
    if (caps.length > existingRows) {
      const values = caps.slice(existingRows).map(c => [c.name, c.team, c.prio || 0, c.xp || 0]);
      sh.getRange(2 + existingRows, 1, values.length, 4).setValues(values);
    }
  } else {
    // データが空の場合は、既存のデータをクリア
    if (existingRows > 0) {
      sh.getRange(2, 1, existingRows, 4).clearContent();
    }
  }
  // 【追加】キャッシュを無効化
  invalidateCache_('captains');
}
function resetCaptainsPrio_(value){ 
  const sh=ensureSheet_(CFG.SHEET.CAPTAINS,['Captain','Team','Prio','XP']); 
  const last=sh.getLastRow(); 
  if(last>=2){ 
    const rows=last-1; 
    const vals=Array.from({length:rows},()=>[Number(value)||0]); 
    sh.getRange(2,3,rows,1).setValues(vals);
    // 【追加】キャッシュを無効化
    invalidateCache_('captains');
  }
}

function readPlayers(){
  return getCached_('players', () => {
    const sh=ensureSheet_(CFG.SHEET.PLAYERS,['Name','Status','エリアXP','持ちブキ','コメント']);
    const vs=sh.getDataRange().getValues(); const rs=[];
    for(let i=1;i<vs.length;i++){ const name=vs[i][0]; if(!name) continue; const active=(vs[i][1]||'active').toString().toLowerCase()==='active'; rs.push({name,active}); }
    return rs.filter(p=>p.active).map(p=>p.name);
  });
}
function readPicks(){ 
  return getCached_('picks', () => {
    const sh=ensureSheet_(CFG.SHEET.PICKS,['Round','Captain','Player','Method','Timestamp']); 
    if(sh.getLastRow()<2) return []; 
    return sh.getDataRange().getValues().slice(1).map(r=>({round:Number(r[0]),captain:r[1],player:r[2],method:r[3],ts:r[4]})); 
  });
}
function remainingPlayers(){ const all=readPlayers(); const picked=new Set(readPicks().map(p=>p.player)); return all.filter(n=>!picked.has(n)); }

function isPlayerPicked_(player, pickedSet) {
  // 【最適化】pickedSetが渡された場合はそれを使用（readPicks()を呼ばない）
  if (pickedSet) {
    return pickedSet.has(String(player));
  }
  return readPicks().some(p => String(p.player) === String(player));
}

function public_getState(){
  const round = getRound_();
  const sub   = getSubround_();
  const picks = readPicks();
  const teams = {};
  picks.forEach(p=>{
    if(!teams[p.captain]) teams[p.captain] = [];
    teams[p.captain].push(p.player);
  });
  // 【変更】Requestsシートではなく、フォームの回答シートから読み取る
  const respSh = getLinkedResponseSheet_();
  const nowReqs = [];
  if (respSh) {
    const header = respSh.getRange(1, 1, 1, respSh.getLastColumn()).getValues()[0].map(String);
    const colRound = header.findIndex(h => /^round$/i.test(h.trim())) + 1;
    const colSub = header.findIndex(h => /^sub$/i.test(h.trim())) + 1;
    const colCap = header.findIndex(h => /キャプテン|captain/i.test(h)) + 1;
    const colPl = header.findIndex(h => /指名する選手|player/i.test(h)) + 1;
    
    if (colRound && colSub && colCap && colPl) {
      const lastRow = respSh.getLastRow();
      if (lastRow > 1) {
        const vals = respSh.getRange(2, 1, lastRow - 1, respSh.getLastColumn()).getValues();
        const seen = new Map(); // captain -> { player, timestamp } (最新のみ)
        
        vals.forEach(row => {
          const rowRound = Number(row[colRound - 1]);
          const rowSub = Number(row[colSub - 1]);
          const rowCaptain = String(row[colCap - 1] || '').trim();
          const rowPlayer = String(row[colPl - 1] || '').trim();
          const colTs = header.findIndex(h => /timestamp|タイムスタンプ/i.test(h));
          const rowTs = colTs >= 0 ? (row[colTs] || new Date(0)) : new Date(0);
          
          if (rowRound === round && rowSub === sub && rowCaptain && rowPlayer) {
            if (!seen.has(rowCaptain) || rowTs > seen.get(rowCaptain).timestamp) {
              seen.set(rowCaptain, { player: rowPlayer, timestamp: rowTs });
            }
          }
        });
        
        seen.forEach((val, cap) => {
          nowReqs.push({ captain: cap, player: val.player });
        });
      }
    }
  }
  const map = {};
  nowReqs.forEach(r=>{
    if(!map[r.player]) map[r.player] = [];
    map[r.player].push(r.captain);
  });
  const contests = Object.keys(map)
    .filter(pl=>map[pl].length>=2)
    .map(pl=>({ player:pl, captains:map[pl] }));
  const singles = Object.keys(map)
    .filter(pl=>map[pl].length===1)
    .map(pl=>({ player:pl, captain:map[pl][0] }));
  const logSh = ensureSheet_(CFG.SHEET.LOG, ["Timestamp","Event","Detail"]);
  const lr = logSh.getLastRow();
  let logs = [];
  if(lr>=2){
    const take = Math.min(10, lr-1);
    logs = logSh.getRange(lr-take+1,1,take,3).getValues()
      .map(r=>({ ts:r[0], ev:String(r[1]), detail:String(r[2]) }));
  }
  let accepting = false, formUrl = '';
  try {
    const f = getForm_();
    accepting = f.isAcceptingResponses();
    if (typeof f.getPublishedUrl === 'function') formUrl = f.getPublishedUrl();
  } catch(_){}
  return {
    round, sub, accepting, formUrl,
    contests, singles,
    picks,
    teams,
    logs
  };
}

function doGet(e){
  // ControlCenterをWebアプリとして公開（認証チェックあり）
  try {
    SpreadsheetApp.getActiveSpreadsheet(); // アクセス権限チェック
  } catch(err) {
    return HtmlService.createHtmlOutput('<h1>アクセス権限がありません</h1><p>このスプレッドシートへのアクセス権限が必要です。<br>スプレッドシートの共有設定で、あなたのアカウントにアクセス権限を付与してください。</p>');
  }
  
  return HtmlService.createHtmlOutputFromFile('src_ControlCenter')
    .setTitle('ドラフト運営センター')
    .setWidth(800)
    .setHeight(600);
}

// 【削除】Broadcast.htmlが削除されたため、この関数は使用不可
// function admin_openBroadcast(){
//   const html = HtmlService.createHtmlOutputFromFile('Broadcast')
//     .setWidth(1000).setHeight(700);
//   SpreadsheetApp.getUi().showModalDialog(html, '配信用プレビュー');
// }

function listFormResponseSheets_(){
  const ss = SpreadsheetApp.getActive();
  const re = /^(Form Responses(?: \d+)?|フォームの回答(?: \d+)?)$/i;
  return ss.getSheets()
    .filter(sh => re.test(sh.getName()))
    .map(sh => {
      const m = sh.getName().match(/(\d+)$/);
      const idx = m ? Number(m[1]) : 1;
      return { sh, name: sh.getName(), idx };
    })
    .sort((a,b) => a.idx - b.idx);
}

// 【追加】古い回答シートを全て削除する関数
function deleteAllFormResponseSheets_(){
  const ss = SpreadsheetApp.getActive();
  const re = /^(Form Responses(?: \d+)?|フォームの回答(?: \d+)?)$/i;
  const sheetsToDelete = ss.getSheets()
    .filter(sh => re.test(sh.getName()));
  
  let deletedCount = 0;
  sheetsToDelete.forEach(sh => {
    try {
      ss.deleteSheet(sh);
      deletedCount++;
      logEvent('RESP_SHEET_DELETED', `回答シートを削除: ${sh.getName()}`);
    } catch (e) {
      logEvent('RESP_SHEET_DELETE_ERR', `回答シート削除失敗: ${sh.getName()} - ${e}`);
    }
  });
  
  return deletedCount;
}

// 【追加】フォームのリンクを解除してから回答シートを削除する関数
function unlinkAndDeleteFormSheets_(formId){
  if (!formId) return 0;
  const ss = SpreadsheetApp.getActive();
  let deletedCount = 0;
  
  try {
    // フォームオブジェクトを取得
    const form = FormApp.openById(formId);
    
    // フォームにリンクされているシートを検出
    const linkedSheets = [];
    ss.getSheets().forEach(sh => {
      try {
        const formUrl = sh.getFormUrl();
        if (formUrl && formUrl.includes(formId)) {
          linkedSheets.push(sh);
        }
      } catch (e) {
        // シートがフォームにリンクされていない場合は無視
      }
    });
    
    if (linkedSheets.length > 0) {
      // 一時的なスプレッドシートを作成して、フォームのリンク先を変更（リンク解除）
      const tempSs = SpreadsheetApp.create('Temp Form Unlink');
      let unlinkSuccess = false;
      try {
        // フォームのリンク先を一時スプレッドシートに変更（これにより元のスプレッドシートとのリンクが解除される）
        form.setDestination(FormApp.DestinationType.SPREADSHEET, tempSs.getId());
        unlinkSuccess = true;
        logEvent('FORM_UNLINKED', `フォームのリンクを解除: ${formId}`);
      } catch (e) {
        // setDestination()でエラーが発生した場合でも、既にリンクが解除されている可能性がある
        // エラーメッセージをログに記録するが、処理は続行する
        logEvent('FORM_UNLINK_WARN', `フォームリンク解除時に警告: ${formId} - ${e} (処理は続行します)`);
        // エラーが発生しても、リンク解除が成功している可能性があるため、unlinkSuccessはfalseのまま
        // ただし、シート削除は試みる（削除できない場合はエラーになる）
      }
      
      // リンク解除が成功した場合、またはエラーが発生したが処理を続行する場合、シート削除を試みる
      linkedSheets.forEach(sh => {
        try {
          // シートがまだフォームにリンクされているか確認
          let isStillLinked = false;
          try {
            const formUrl = sh.getFormUrl();
            if (formUrl && formUrl.includes(formId)) {
              isStillLinked = true;
            }
          } catch (e) {
            // getFormUrl()でエラーが発生した場合、リンクが解除されている可能性が高い
            isStillLinked = false;
          }
          
          if (!isStillLinked || unlinkSuccess) {
            // リンクが解除されている場合、またはリンク解除が成功した場合、シートを削除
            ss.deleteSheet(sh);
            deletedCount++;
            logEvent('FORM_LINKED_SHEET_DELETED', `フォームにリンクされていた回答シートを削除: ${sh.getName()}`);
          } else {
            logEvent('FORM_SHEET_STILL_LINKED', `回答シート「${sh.getName()}」はまだフォームにリンクされているため削除をスキップ`);
          }
        } catch (e) {
          logEvent('FORM_SHEET_DELETE_ERR', `回答シート削除失敗: ${sh.getName()} - ${e}`);
        }
      });
      
      // 一時的なスプレッドシートを削除
      try {
        DriveApp.getFileById(tempSs.getId()).setTrashed(true);
      } catch (e) {
        logEvent('TEMP_SS_DELETE_ERR', `一時スプレッドシート削除失敗: ${e}`);
      }
    }
  } catch (e) {
    // フォームが既に削除されている、またはアクセスできない場合はエラーを無視
    logEvent('FORM_UNLINK_ERR', `フォームリンク解除失敗: ${formId} - ${e}`);
  }
  
  return deletedCount;
}

// 【追加】フォームを削除する関数（DriveAppを使用）
// 削除前にリンクを解除してから回答シートも削除
function deleteForm_(formId){
  if (!formId) return false;
  try {
    // まず、フォームのリンクを解除してから回答シートを削除
    const deletedSheets = unlinkAndDeleteFormSheets_(formId);
    if (deletedSheets > 0) {
      logEvent('FORM_LINKED_SHEETS_DELETED', `フォームにリンクされていた回答シート ${deletedSheets} 件を削除`);
    }
    
    // その後、フォームを削除（ゴミ箱へ移動）
    const file = DriveApp.getFileById(formId);
    file.setTrashed(true); // ゴミ箱に入れる（完全削除はDriveAppでは不可）
    logEvent('FORM_DELETED', `フォームを削除（ゴミ箱へ移動）: ${formId}`);
    return true;
  } catch (e) {
    // フォームが既に削除されている、またはアクセス権限がない場合はエラーを無視
    logEvent('FORM_DELETE_ERR', `フォーム削除失敗: ${formId} - ${e}`);
    return false;
  }
}

function getLinkedResponseSheet_(){
  const ss = SpreadsheetApp.getActive();
  for (const sh of ss.getSheets()) {
    try {
      const url = sh.getFormUrl();
      if (url) return sh;
    } catch (e) {}
  }
  return null;
}

function getForm_(){ const id=getConfigMap()[CFG.CFG_KEYS.FORM_ID]; if(!id) throw new Error('FORM_IDがConfigにありません。setupDraft()を実行してください'); return FormApp.openById(id); }
function getRound_(){ return Number(getConfigMap()[CFG.CFG_KEYS.ROUND]||1); }
function setRound_(r){ putConfig(CFG.CFG_KEYS.ROUND, r); }
function setOpen_(f){ putConfig(CFG.CFG_KEYS.IS_OPEN, f?'1':'0'); }
function getSubround_(){ return Number(getConfigMap()[CFG.CFG_KEYS.SUBROUND]||1); }
function setSubround_(n){ putConfig(CFG.CFG_KEYS.SUBROUND, n); }
function getEligibleCaptains_(){ const txt=getConfigMap()[CFG.CFG_KEYS.ELIGIBLE_CAPTAINS]||'[]'; try{ const arr=JSON.parse(txt); return Array.isArray(arr)?arr:[]; }catch(e){ return []; } }
function setEligibleCaptains_(arr){ putConfig(CFG.CFG_KEYS.ELIGIBLE_CAPTAINS, JSON.stringify(arr||[])); }

function clearAndCompactSheet_(sh) {
  if (!sh) { logEvent('RESP_COMPACT_SKIP', 'sheet=null'); return; }
  try {
    logEvent('RESP_COMPACT_BEGIN', sh.getName());
    try {
      const filter = sh.getFilter && sh.getFilter();
      if (filter) filter.remove();
    } catch (_) {}
    try { sh.setFrozenRows(1); } catch (_) {}
    
    const lc = sh.getLastColumn();
    const lr = sh.getLastRow();
    if (lr > 1) {
      // 【改善】clearContent()の代わりにclear()を使用して、より確実にデータを削除
      // ただし、ヘッダー行（1行目）は保持する必要があるため、2行目以降を削除
      const dataRange = sh.getRange(2, 1, lr - 1, lc);
      dataRange.clear(); // clearContent()ではなくclear()を使用（フォーマットも含めて削除）
    }
    if (sh.getMaxRows() === 1) {
      sh.insertRowsAfter(1, 1);
    }
    let maxRows = sh.getMaxRows();
    if (maxRows > 2) {
      sh.deleteRows(3, maxRows - 2);
    }
    const finalCols = sh.getLastColumn();
    logEvent('RESP_COMPACT_DONE', sh.getName() + ` (列数: ${finalCols}, データ削除完了)`);
  } catch (err) {
    logEvent('RESP_COMPACT_ERR', sh.getName() + ': ' + err);
    throw err;
  }
}

function getCaptainListItem_(){ const form=getForm_(); const it=form.getItems().find(i=>i.getTitle()===CFG.FORM.Q_CAPTAIN); if(!it) throw new Error('フォームにキャプテン設問が見つかりません'); return it.asListItem(); }
function updateCaptainChoices_(eligible){ const item=getCaptainListItem_(); const names=(eligible&&eligible.length)?eligible:readCaptains().map(c=>c.name); item.setChoices(names.map(n=>item.createChoice(n))); }
function updatePlayerChoices_(){ const form=getForm_(); const rem=remainingPlayers(); form.getItems().forEach(it=>{ if(it.getType()===FormApp.ItemType.LIST && it.getTitle()===CFG.FORM.Q_PLAYER){ const li=it.asListItem(); if(rem.length===0){ li.setChoices([li.createChoice('（指名可能な選手がいません）')]); }else{ li.setChoices(rem.map(n=>li.createChoice(n))); } } }); }
function getManualMode_(){ return String(getConfigMap()[CFG.CFG_KEYS.MANUAL_MODE]||'0')==='1'; }
function setManualMode_(on){ putConfig(CFG.CFG_KEYS.MANUAL_MODE, on?'1':'0'); }
function updateDataTimestamp_(){ putConfig(CFG.CFG_KEYS.LAST_UPDATE_TIMESTAMP, new Date().toISOString()); }
function getDataTimestamp_(){ return String(getConfigMap()[CFG.CFG_KEYS.LAST_UPDATE_TIMESTAMP]||''); }

function listContestantsForPlayer_(round, sub, player) {
  const respSh = getLinkedResponseSheet_();
  if (!respSh) return [];
  
  const header = respSh.getRange(1, 1, 1, respSh.getLastColumn()).getValues()[0].map(String);
  const colRound = header.findIndex(h => /^round$/i.test(h.trim())) + 1;
  const colSub = header.findIndex(h => /^sub$/i.test(h.trim())) + 1;
  const colCap = header.findIndex(h => /キャプテン|captain/i.test(h)) + 1;
  const colPl = header.findIndex(h => /指名する選手|player/i.test(h)) + 1;
  
  if (!colRound || !colSub || !colCap || !colPl) return [];
  
  const lastRow = respSh.getLastRow();
  if (lastRow <= 1) return [];
  
  const vals = respSh.getRange(2, 1, lastRow - 1, respSh.getLastColumn()).getValues();
  const captains = [];
  const seen = new Set();
  
  vals.forEach(row => {
    const rowRound = Number(row[colRound - 1]);
    const rowSub = Number(row[colSub - 1]);
    const rowPlayer = String(row[colPl - 1] || '').trim();
    const rowCaptain = String(row[colCap - 1] || '').trim();
    
    if (rowRound === round && rowSub === sub && rowPlayer === String(player).trim() && rowCaptain && !seen.has(rowCaptain)) {
      captains.push(rowCaptain);
      seen.add(rowCaptain);
    }
  });
  
  return captains
    .filter(Boolean);
}

function setupDraft(){
  try {
    return setupDraft_();
  } catch(e) {
    logEvent('SETUP_ERROR', `ドラフト準備エラー: ${e}`);
    throw e;
  }
}

function setupDraft_(){
  const ss=_ss();
  const title=`${CFG.FORM.TITLE_PREFIX} - ${ss.getName()}`;
  
  // 【追加】ドラフト準備時にEntryシートからCaptains/Playersをインポート
  const sourceSheetName = 'Entry';
  const sourceSheet = ss.getSheetByName(sourceSheetName);
  
  if (!sourceSheet) {
    throw new Error(`シート「${sourceSheetName}」が見つかりません。Captains/Playersのインポートに必要です。`);
  }
  
  // Captainsシートをインポート
  const captainsResult = importCaptainsAuto_(sourceSheet, sourceSheetName);
  if (!captainsResult.success) {
    throw new Error(`Captainsシートのインポートに失敗しました: ${captainsResult.error}`);
  }
  logEvent('SETUP_IMPORT_CAPTAINS', `Captainsシートをインポート: ${captainsResult.count}件`);
  
  // Playersシートをインポート
  const playersResult = importPlayersAuto_(sourceSheet, sourceSheetName);
  if (!playersResult.success) {
    throw new Error(`Playersシートのインポートに失敗しました: ${playersResult.error}`);
  }
  logEvent('SETUP_IMPORT_PLAYERS', `Playersシートをインポート: ${playersResult.count}件`);
  
  // 【追加】ドラフト準備時にCaptainsのPrioを0にリセット
  resetCaptainsPrio_(0);
  
  const captains=readCaptains(); if(captains.length===0) throw new Error('Captainsシートが空です');
  const rem=remainingPlayers(); if(rem.length===0) throw new Error('Playersシートに有効な参加者がいません');
  
  // 【変更】フォームは常に新規作成（使い捨て方式）
  // 既存のFORM_IDがあればフォームを削除（リンク解除と回答シート削除も自動実行）
  const existingFormId = getConfigMap()[CFG.CFG_KEYS.FORM_ID];
  if (existingFormId) {
    try {
      deleteForm_(existingFormId); // フォーム削除時にリンクされた回答シートも自動削除
      logEvent('SETUP_CLEANUP', `既存のフォームを削除（ゴミ箱へ移動）: ${existingFormId}`);
    } catch (e) {
      logEvent('SETUP_CLEANUP_ERR', `フォーム削除エラー: ${existingFormId} - ${e}`);
    }
    putConfig(CFG.CFG_KEYS.FORM_ID, '');
    logEvent('SETUP_CLEANUP', `既存のFORM_IDを削除: ${existingFormId}`);
  }
  
  // 残っている古い回答シートを削除（念のため）
  try {
    const deletedCount = deleteAllFormResponseSheets_();
    if (deletedCount > 0) {
      logEvent('SETUP_CLEANUP', `残っていた古い回答シート ${deletedCount} 件を削除しました`);
    }
  } catch (e) {
    logEvent('SETUP_CLEANUP_ERR', '回答シート削除エラー: ' + e);
  }
  
  // 新規フォームを作成
  const form = FormApp.create(title);
  form.setDescription(CFG.FORM.DESC);
  form.setAllowResponseEdits(false);
  form.setCollectEmail(false);
  form.setLimitOneResponsePerUser(false);
  
  const capItem = form.addListItem();
  capItem.setTitle(CFG.FORM.Q_CAPTAIN).setChoices(captains.map(c=>capItem.createChoice(c.name))).setRequired(true);
  const plyItem = form.addListItem();
  plyItem.setTitle(CFG.FORM.Q_PLAYER).setChoices(rem.map(n=>plyItem.createChoice(n))).setRequired(true);
  
  // 【追加】フォームをリンクする直前に、全ての「フォームの回答」系シートを一時的に別名に変更
  // （Googleフォームが「フォームの回答」という名前が空いていると判断して、数字なしで作成するようにする）
  const tempRenamedSheets = [];
  try {
    const re = /^フォームの回答(\s+\d+)?$/;
    ss.getSheets().forEach(sh => {
      const sheetName = sh.getName();
      if (re.test(sheetName)) {
        try {
          // フォームにリンクされていない場合は削除可能
          const formUrl = sh.getFormUrl();
          if (!formUrl) {
            ss.deleteSheet(sh);
            logEvent('SETUP_CLEANUP', `既存の「${sheetName}」シートを削除（リンクなし）`);
          } else {
            // フォームにリンクされている場合は、一時的に別名に変更（リンク解除後に削除）
            const tempName = 'フォームの回答_一時_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
            sh.setName(tempName);
            tempRenamedSheets.push({ sheet: sh, originalName: sheetName, tempName: tempName });
            logEvent('SETUP_CLEANUP', `既存の「${sheetName}」シートを一時的に「${tempName}」にリネーム`);
          }
        } catch (e) {
          logEvent('SETUP_CLEANUP_ERR', `「${sheetName}」シート処理エラー: ${e}`);
        }
      }
    });
  } catch (e) {
    logEvent('SETUP_CLEANUP_ERR', `回答シート確認エラー: ${e}`);
  }
  
  // フォームをスプレッドシートにリンク（「フォームの回答」という名前が空いている状態でリンク）
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  
  // 【重要】フォームをリンクした直後に回答シート名を変更（Googleフォームのカウンターをリセット）
  // 少し待ってからシートを取得（フォームのリンク処理が完了するまで）
  Utilities.sleep(500);
  
  const newLinkedSheet = getLinkedResponseSheet_();
  if (newLinkedSheet) {
    // 【追加】回答シートの名前を必ず「フォームの回答」に変更（数字を付けない）
    const targetName = 'フォームの回答';
    const currentName = newLinkedSheet.getName();
    if (currentName !== targetName) {
      try {
        // 既に「フォームの回答」という名前のシートが存在する場合は、一時的に別名に変更
        const tempSheet = ss.getSheetByName(targetName);
        if (tempSheet && tempSheet.getSheetId() !== newLinkedSheet.getSheetId()) {
          tempSheet.setName('フォームの回答_削除予定_' + Date.now());
        }
        // 回答シート名を「フォームの回答」に変更
        newLinkedSheet.setName(targetName);
        logEvent('SETUP_SHEET', `回答シート名を変更: ${currentName} → ${targetName}`);
      } catch (e) {
        logEvent('SETUP_SHEET_RENAME_ERR', `回答シート名変更失敗: ${e}`);
      }
    } else {
      logEvent('SETUP_SHEET', `新しい回答シート作成: ${targetName}`);
    }
  }
  
  // 【追加】一時的にリネームしたシートを削除
  tempRenamedSheets.forEach(({ sheet, originalName, tempName }) => {
    try {
      // フォームにリンクされていないことを確認してから削除
      const formUrl = sheet.getFormUrl();
      if (!formUrl) {
        ss.deleteSheet(sheet);
        logEvent('SETUP_CLEANUP', `一時リネームした「${tempName}」シートを削除（元: ${originalName}）`);
      } else {
        logEvent('SETUP_CLEANUP', `「${tempName}」シートはまだフォームにリンクされているため削除スキップ`);
      }
    } catch (e) {
      logEvent('SETUP_CLEANUP_ERR', `「${tempName}」シート削除エラー: ${e}`);
    }
  });
  
  putConfig(CFG.CFG_KEYS.FORM_ID, form.getId());
  logEvent('SETUP_CREATE', `新規フォーム作成: ${form.getEditUrl()}`);
  
  setRound_(1); setSubround_(1); setEligibleCaptains_(captains.map(c=>c.name)); setOpen_(0);
  const shLog=ensureSheet_(CFG.SHEET.LOG,['Timestamp','Event','Detail']); 
  shLog.appendRow([new Date(),'SETUP','フォーム作成: '+form.getEditUrl()]);
  // 【修正】共通フォームURLをログに記録（各キャプテン用の事前入力URLは生成しない）
  try {
    const formUrl = form.getPublishedUrl();
    if (formUrl) {
      shLog.appendRow([new Date(),'FORM_URL','共通フォームURL: '+formUrl]);
    }
  } catch (e) {
    logEvent('SETUP_WARN', 'フォーム公開URLの取得に失敗: '+e);
  }
  createInstallableTrigger_();
  
  // ドラフト準備完了フラグを設定（putConfig内で自動的にタイムスタンプが更新される）
  putConfig(CFG.CFG_KEYS.SETUP_COMPLETE, '1');
  
  // 成功状態を返す
  return {
    success: true,
    isNewForm: true, // 常に新規作成
    formId: form.getId(),
    formUrl: form.getPublishedUrl() || '',
    editUrl: `https://docs.google.com/forms/d/${form.getId()}/edit`,
    captainsCount: captains.length,
    playersCount: rem.length
  };
}
// 【削除】各キャプテン用の事前入力URL生成関数は不要になったため削除
// function buildPrefilledUrls_(){ const form=getForm_(); const capItem=getCaptainListItem_(); const urls=[]; readCaptains().forEach(c=>{ const resp=form.createResponse(); resp.withItemResponse(capItem.createResponse(c.name)); urls.push({captain:c.name, url: resp.toPrefilledUrl()}); }); return urls; }
function createInstallableTrigger_(){ 
  const triggers=ScriptApp.getProjectTriggers(); 
  const formId = getConfigMap()[CFG.CFG_KEYS.FORM_ID];
  if (!formId) {
    logEvent('TRIGGER_ERR', 'FORM_IDが設定されていません');
    return;
  }
  
  // 既存のトリガーをチェック（同じフォームIDのトリガーが存在するか）
  const existingTrigger = triggers.find(t => {
    try {
      return t.getHandlerFunction() === 'onFormSubmit' && 
             t.getTriggerSourceId() === formId;
    } catch (e) {
      return false;
    }
  });
  
  if (!existingTrigger) {
    try {
      const form = getForm_();
      ScriptApp.newTrigger('onFormSubmit').forForm(form).onFormSubmit().create(); 
      logEvent('TRIGGER', `onFormSubmit を作成: FORM_ID=${formId}`);
    } catch (e) {
      logEvent('TRIGGER_ERR', `トリガー作成失敗: ${e}`);
    }
  } else {
    logEvent('TRIGGER_CHECK', `onFormSubmit トリガーは既に存在します: FORM_ID=${formId}`);
  }
}

function openRound(){ const form=getForm_(); const round=getRound_(); setSubround_(1); setEligibleCaptains_(readCaptains().map(c=>c.name)); updateCaptainChoices_(getEligibleCaptains_()); updatePlayerChoices_(); form.setAcceptingResponses(true); form.setDescription(`${CFG.FORM.DESC}\n現在のラウンド: ${round} / サブラウンド: ${getSubround_()}`); setOpen_(1); logEvent('ROUND_OPEN',`Round ${round} sub=${getSubround_()} opened`); }
function closeRound(){ const form=getForm_(); form.setAcceptingResponses(false); setOpen_(0); logEvent('ROUND_CLOSE',`Round ${getRound_()} sub=${getSubround_()} closed`); }

// ===== フォーム受信＆解決（サブラウンド） =====
function onFormSubmit(e){
  const lock=LockService.getScriptLock(); 
  if (!lock.tryLock(30000)) {
    logEvent('ONSUBMIT_LOCK_FAIL', 'ロック取得失敗');
    return;
  }
  try{
    const round=getRound_(); const sub=getSubround_(); const eligible=new Set(getEligibleCaptains_());
    const resp=e.response; 
    if (!resp) {
      logEvent('ONSUBMIT_ERR', 'responseがnullです');
      return;
    }
    const itemResponses=resp.getItemResponses(); 
    let captain='', player='';
    itemResponses.forEach(ir=>{ 
      const t=ir.getItem().getTitle(); 
      if(t===CFG.FORM.Q_CAPTAIN) captain=String(ir.getResponse()||''); 
      if(t===CFG.FORM.Q_PLAYER) player=String(ir.getResponse()||''); 
    });
    
    logEvent('ONSUBMIT_RECEIVED', `round=${round}, sub=${sub}, captain=${captain}, player=${player}`);
    
    if(captain && !eligible.has(captain)){ 
      logEvent('IGNORED_SUBMISSION',`round=${round}, sub=${sub}, captain=${captain} (not eligible)`); 
      return; 
    }
    // 【バグ修正】空のplayerを除外
    if(!player || player.trim()===''){ 
      logEvent('IGNORED_SUBMISSION',`round=${round}, sub=${sub}, captain=${captain} (empty player)`); 
      return; 
    }
    
    // 【変更】フォームの回答シートにRound/Subを追記
    // タイムスタンプを使って最新の回答（今回送信された回答）を特定し、その行にRound/Subを追記
    const respSh = getLinkedResponseSheet_();
    if (!respSh) {
      logEvent('ONSUBMIT_ERR', 'フォームの回答シートが見つかりません');
      return;
    }
    
    // ヘッダー行を確認し、Round/Sub列が存在しない場合は追加
    const headerRow = respSh.getRange(1, 1, 1, respSh.getLastColumn()).getValues()[0].map(String);
    let colRound = headerRow.findIndex(h => /^round$/i.test(h.trim()));
    let colSub = headerRow.findIndex(h => /^sub$/i.test(h.trim()));
    const colTs = headerRow.findIndex(h => /timestamp|タイムスタンプ/i.test(h));
    
    // Round/Sub列が存在しない場合は追加
    if (colRound < 0) {
      colRound = respSh.getLastColumn();
      respSh.getRange(1, colRound + 1).setValue('Round');
      colRound = colRound + 1; // 1ベースに変換
    } else {
      colRound = colRound + 1; // 1ベースに変換
    }
    
    if (colSub < 0) {
      colSub = respSh.getLastColumn();
      respSh.getRange(1, colSub + 1).setValue('Sub');
      colSub = colSub + 1; // 1ベースに変換
    } else {
      colSub = colSub + 1; // 1ベースに変換
    }
    
    // タイムスタンプを使って最新の回答（今回送信された回答）を特定
    const submitTime = resp.getTimestamp();
    let targetRow = -1;
    
    if (colTs >= 0 && submitTime) {
      // タイムスタンプ列がある場合、最も近いタイムスタンプの行を探す
      const lastRow = respSh.getLastRow();
      if (lastRow > 1) {
        const colTs1 = colTs + 1; // 1ベースに変換
        const timestamps = respSh.getRange(2, colTs1, lastRow - 1, 1).getValues();
        let minDiff = Infinity;
        
        for (let i = 0; i < timestamps.length; i++) {
          const rowTime = timestamps[i][0];
          if (rowTime instanceof Date) {
            const diff = Math.abs(rowTime.getTime() - submitTime.getTime());
            if (diff < minDiff) {
              minDiff = diff;
              targetRow = i + 2; // 1ベースの行番号（ヘッダー行を考慮）
            }
          }
        }
        
        // タイムスタンプが一致する行が見つからない場合、または1分以上の差がある場合は最後の行を使用
        if (targetRow < 0 || minDiff > 60000) {
          targetRow = lastRow;
        }
      }
    } else {
      // タイムスタンプ列がない場合は、最後の行を使用（フォールバック）
      targetRow = respSh.getLastRow();
    }
    
    // 特定した行のRound/Subを現在の値で追記
    if (targetRow > 1) {
      respSh.getRange(targetRow, colRound).setValue(round);
      respSh.getRange(targetRow, colSub).setValue(sub);
      logEvent('ONSUBMIT_ADDED', `round=${round}, sub=${sub}, captain=${captain}, player=${player} の提出を追加（回答シート行${targetRow}、Round/Subを追記）`);
    } else {
      logEvent('ONSUBMIT_WARN', `回答シートにデータ行が見つかりません（targetRow=${targetRow}）`);
    }

    // 【追加】データ更新タイムスタンプを更新
    updateDataTimestamp_();

    if (!getManualMode_()) {
      resolveIfReady_();
    } else {
      logEvent('WAIT_MANUAL', `Round=${round} Sub=${sub} 受付済み`);
    }

  } catch (err) {
    logEvent('ONSUBMIT_ERR', `エラー: ${err.toString()}, stack: ${err.stack}`);
  } finally { 
    lock.releaseLock(); 
  }
}

// === サブラウンド解決（提出が揃ったら抽選→次へ） ===
function resolveIfReady_(){
  const round=getRound_(); const sub=getSubround_();
  const captainsAll=readCaptains().map(c=>c.name);
  const eligible=getEligibleCaptains_(); const need=eligible.length;
  
  // 【変更】Requestsシートではなく、フォームの回答シートから読み取る
  const respSh = getLinkedResponseSheet_();
  if (!respSh) return;
  
  const header = respSh.getRange(1, 1, 1, respSh.getLastColumn()).getValues()[0].map(String);
  const colRound = header.findIndex(h => /^round$/i.test(h.trim())) + 1;
  const colSub = header.findIndex(h => /^sub$/i.test(h.trim())) + 1;
  const colCap = header.findIndex(h => /キャプテン|captain/i.test(h)) + 1;
  const colPl = header.findIndex(h => /指名する選手|player/i.test(h)) + 1;
  const colTs = header.findIndex(h => /timestamp|タイムスタンプ/i.test(h)) + 1;
  
  if (!colRound || !colSub || !colCap || !colPl) return;
  
  const lastRow = respSh.getLastRow();
  if (lastRow <= 1) return;
  
  const vs = respSh.getRange(2, 1, lastRow - 1, respSh.getLastColumn()).getValues();
  const rowsThis = [];
  const seen = new Map(); // captain -> { round, sub, player, timestamp } (最新のみ)
  
  vs.forEach(row => {
    const rowRound = Number(row[colRound - 1]);
    const rowSub = Number(row[colSub - 1]);
    const rowCaptain = String(row[colCap - 1] || '').trim();
    const rowPlayer = String(row[colPl - 1] || '').trim();
    const rowTs = colTs > 0 ? (row[colTs - 1] || new Date(0)) : new Date(0);
    
    if (rowRound === round && rowSub === sub && rowCaptain && rowPlayer) {
      if (!seen.has(rowCaptain) || rowTs > seen.get(rowCaptain).timestamp) {
        seen.set(rowCaptain, { round: rowRound, sub: rowSub, player: rowPlayer, timestamp: rowTs });
      }
    }
  });
  
  seen.forEach((val, cap) => {
    rowsThis.push([val.round, val.sub, cap, val.player, val.timestamp, '']);
  });
  
  if(rowsThis.length < need) return;

  const form=getForm_(); form.setAcceptingResponses(false); setOpen_(0); logEvent('SUBROUND_CLOSE',`Round ${round} sub=${sub} closed`);

  // 【追加】既に選ばれたプレイヤーのセットを作成
  const pickedSet = new Set(readPicks().map(p => String(p.player)));

  const picksThis=[]; const updatedCaptains=readCaptains(); const prioOf=(cn)=> (updatedCaptains.find(x=>x.name===cn)?.prio)||0;
  // 【バグ修正】空のplayerを除外、既に選ばれたプレイヤーも除外、重複は最新のみ
  const choiceByCap=new Map(); 
  // 同じキャプテンの最新の提出のみを使用するため、時系列でソート
  const sortedRows = rowsThis.slice().sort((a,b) => {
    const tsA = a[4] instanceof Date ? a[4].getTime() : (new Date(a[4]||0)).getTime();
    const tsB = b[4] instanceof Date ? b[4].getTime() : (new Date(b[4]||0)).getTime();
    return tsB - tsA; // 新しい順
  });
  sortedRows.forEach(r=>{ 
    const cap=String(r[2]||''); 
    const pl=String(r[3]||''); 
    // 空でなく、既に選ばれていない、かつこのキャプテンの最初の有効な提出のみ
    if(cap && pl && pl.trim()!=='' && !pickedSet.has(pl) && !choiceByCap.has(cap)) {
      choiceByCap.set(cap, pl);
    }
  });
  
  const wants=new Map(); 
  choiceByCap.forEach((player,cap)=>{ 
    if(!player || player.trim()==='' || pickedSet.has(player)) return; 
    if(!wants.has(player)) wants.set(player,[]); 
    wants.get(player).push(cap); 
  });

  const losers=[];
  wants.forEach((caps,player)=>{ 
    if(caps.length===1){ 
      const cap=caps[0]; 
      picksThis.push({captain:cap, player, method:`通常(第${sub}希望)`}); 
    } 
  });
  wants.forEach((caps,player)=>{ 
    if(caps.length>=2){ 
      const weights=caps.map(cn=>1+Number(prioOf(cn)||0)); 
      const winIdx=weightedChoice_(weights); 
      const winner=caps[winIdx]; 
      picksThis.push({captain:winner, player, method:`抽選(第${sub}希望)`}); 
      caps.forEach((cn,i)=>{ 
        if(i!==winIdx){ 
          const c=updatedCaptains.find(x=>x.name===cn); 
          if(c) c.prio=Number(c.prio||0)+1; 
          losers.push(cn);
        } 
      }); 
    } 
  });

  const shPicks=ensureSheet_(CFG.SHEET.PICKS,['Round','Captain','Player','Method','Timestamp']); 
  // 【最適化】バッチ処理で一度に追加（appendRow()を個別に呼ぶより高速）
  if (picksThis.length > 0) {
    const lastRow = shPicks.getLastRow();
    const rowsToAdd = picksThis.map(a => [round, a.captain, a.player, a.method, new Date()]);
    shPicks.getRange(lastRow + 1, 1, rowsToAdd.length, 5).setValues(rowsToAdd);
    // 【追加】キャッシュを無効化
    invalidateCache_('picks');
  }
  writeCaptains(updatedCaptains); rebuildTeams_(); logEvent('SUBROUND_RESOLVED', JSON.stringify(picksThis));

  const pickedCaps=new Set(picksThis.map(p=>p.captain)); const unassigned=eligible.filter(c=>!pickedCaps.has(c)); const stillPlayers=remainingPlayers();
  if(unassigned.length>0 && stillPlayers.length>0){ setSubround_(sub+1); setEligibleCaptains_(unassigned); updateCaptainChoices_(unassigned); updatePlayerChoices_(); form.setAcceptingResponses(true); form.setDescription(`${CFG.FORM.DESC}\n現在のラウンド: ${round} / サブラウンド: ${getSubround_()}`); setOpen_(1); logEvent('SUBROUND_OPEN',`Round ${round} sub=${getSubround_()} opened (eligible=${JSON.stringify(unassigned)})`); return; }

  setRound_(round+1); setSubround_(1); setEligibleCaptains_(captainsAll); updateCaptainChoices_(getEligibleCaptains_()); updatePlayerChoices_(); form.setAcceptingResponses(true); form.setDescription(`${CFG.FORM.DESC}\n現在のラウンド: ${getRound_()} / サブラウンド: ${getSubround_()}`); setOpen_(1); logEvent('ROUND_OPEN',`Round ${getRound_()} sub=${getSubround_()} opened`);
}

function draft_getCurrentContests() {
  const round = getRound_(), sub = getSubround_();
  const eligible = new Set(getEligibleCaptains_());

  // 【変更】Requestsシートではなく、フォームの回答シートから直接読み取る
  const respSh = getLinkedResponseSheet_();
  let rows = [];
  
  if (respSh) {
    const header = respSh.getRange(1, 1, 1, respSh.getLastColumn()).getValues()[0].map(String);
    const colRound = header.findIndex(h => /^round$/i.test(h.trim())) + 1;
    const colSub = header.findIndex(h => /^sub$/i.test(h.trim())) + 1;
    const colCap = header.findIndex(h => /キャプテン|captain/i.test(h)) + 1;
    const colPl = header.findIndex(h => /指名する選手|player/i.test(h)) + 1;
    const colTs = header.findIndex(h => /timestamp|タイムスタンプ/i.test(h)) + 1;
    
    if (colRound && colSub && colCap && colPl) {
      const lastRow = respSh.getLastRow();
      if (lastRow > 1) {
        const vs = respSh.getRange(2, 1, lastRow - 1, respSh.getLastColumn()).getValues();
        const seen = new Map(); // captain -> { round, sub, player, timestamp } (最新のみ)
        
        vs.forEach(row => {
          const rowRound = Number(row[colRound - 1]);
          const rowSub = Number(row[colSub - 1]);
          const rowCaptain = String(row[colCap - 1] || '').trim();
          const rowPlayer = String(row[colPl - 1] || '').trim();
          const rowTs = colTs > 0 ? (row[colTs - 1] || new Date(0)) : new Date(0);
          
          if (rowRound === round && rowSub === sub && rowCaptain && rowPlayer) {
            if (!seen.has(rowCaptain) || rowTs > seen.get(rowCaptain).timestamp) {
              seen.set(rowCaptain, { round: rowRound, sub: rowSub, player: rowPlayer, timestamp: rowTs });
            }
          }
        });
        
        seen.forEach((val, cap) => {
          rows.push([val.round, val.sub, cap, val.player, val.timestamp, '']);
        });
      }
    }
  }
  
  // 【削除】補完機能は不要（フォームの回答シートから直接読み取るため）

  // 【削除】デバッグログを削除（自動更新時にログが膨れ上がるため）
  // logEvent('DRAFT_DEBUG', `Round=${round} Sub=${sub} Eligible=${JSON.stringify([...eligible])} Requests=${rows.length}件`);

  const submittedBy = new Map();
  // 【バグ修正】空のplayerを除外、重複は最新のみ、eligibleチェックも追加
  // 同じキャプテンの最新の提出のみを使用するため、時系列でソート
  const sortedRows = rows.slice().sort((a,b) => {
    const tsA = a[4] instanceof Date ? a[4].getTime() : (new Date(a[4]||0)).getTime();
    const tsB = b[4] instanceof Date ? b[4].getTime() : (new Date(b[4]||0)).getTime();
    return tsB - tsA; // 新しい順
  });
  
  // 【デバッグ】提出データの詳細をログに記録
  const debugSubmissions = [];
  sortedRows.forEach(r => { 
    const cap = String(r[2] || '').trim(); 
    const pl = String(r[3] || '').trim(); 
    const isEligible = eligible.has(cap);
    const alreadySubmitted = submittedBy.has(cap);
    const isValid = cap && pl && pl !== '';
    
    debugSubmissions.push({
      cap, pl, isEligible, alreadySubmitted, isValid,
      willAdd: isValid && isEligible && !alreadySubmitted
    });
    
    // 空でなく、eligibleで、このキャプテンの最初の有効な提出のみ
    if(isValid && isEligible && !alreadySubmitted) {
      submittedBy.set(cap, pl); 
    }
  });
  
  // 【デバッグ】問題がある場合のみログに記録
  const hasIssues = debugSubmissions.some(d => !d.willAdd && d.isValid);
  if (hasIssues) {
    logEvent('DRAFT_DEBUG_SUBMISSIONS', JSON.stringify({
      round, sub,
      eligible: [...eligible],
      submissions: debugSubmissions,
      submittedBy: Object.fromEntries(submittedBy)
    }));
  }

  const waiting = [...eligible].filter(c => !submittedBy.has(c));
  // 【最適化】readPicks()を一度だけ呼び出し（既に取得済みの場合は再利用）
  const picks = readPicks();
  // 【修正】pickedSetのキーをString()で統一（比較の一貫性を保つ）
  const pickedSet = new Set(picks.map(p => String(p.player).trim()));

  const wants = new Map();
  submittedBy.forEach((player, cap) => {
    if (!player || player.trim()==='') return;
    // 【修正】pickedSetのキーはString().trim()で統一されているため、比較も統一
    const playerKey = String(player).trim();
    if (pickedSet.has(playerKey)) {
      // 【削除】デバッグログを削除（自動更新時にログが膨れ上がるため）
      // logEvent('DRAFT_DEBUG_SKIP_PICKED', `player=${playerKey} is already picked, round=${round}, sub=${sub}`);
      return;
    }
    if (!wants.has(playerKey)) wants.set(playerKey, []);
    wants.get(playerKey).push(cap);
  });
  
  // 【削除】デバッグログを削除（自動更新時にログが膨れ上がるため）
  // logEvent('DRAFT_DEBUG_WANTS', JSON.stringify({
  //   round, sub,
  //   submittedBy: Object.fromEntries(submittedBy),
  //   wants: Object.fromEntries(wants),
  //   pickedSet: [...pickedSet]
  // }));

  // 【最適化】readCaptains()を一度だけ呼び出し（既に取得済みの場合は再利用）
  const capsAll = readCaptains();
  // 【最適化】Mapで高速化
  const prioMap = new Map(capsAll.map(c => [c.name, Number(c.prio || 0)]));
  const prioOf = (cn) => prioMap.get(cn) || 0;

  const contests = [], singles = [];
  wants.forEach((capsArr, player) => {
    if (capsArr.length >= 2) {
      contests.push({
        player,
        captains: capsArr.map(cn => ({ name: cn, prio: Number(prioOf(cn) || 0) }))
      });
    } else if (capsArr.length === 1) {
      singles.push({ player, captain: capsArr[0] });
    }
  });
  
  // 【削除】デバッグログを削除（自動更新時にログが膨れ上がるため）
  // logEvent('DRAFT_DEBUG_RESULT', JSON.stringify({
  //   round, sub,
  //   singles: singles.map(s => ({ player: s.player, captain: s.captain })),
  //   contests: contests.map(c => ({ player: c.player, captains: c.captains.map(cap => cap.name) })),
  //   waiting
  // }));
  
  // 【削除】デバッグログを削除（自動更新時にログが膨れ上がるため）
  // if (submittedBy.size > 0 && singles.length === 0 && contests.length === 0) {
  //   logEvent('DRAFT_DEBUG_NO_SINGLES', JSON.stringify({
  //     round, sub,
  //     submittedBy: Object.fromEntries(submittedBy),
  //     wants: Object.fromEntries(wants),
  //     pickedSet: [...pickedSet]
  //   }));
  // }

  // 【削除】デバッグログを削除（自動更新時にログが膨れ上がるため）
  // logEvent('DRAFT_DEBUG', `結果: 単独=${singles.length}件 競合=${contests.length}件 未提出=${waiting.length}名`);

  return { round, sub, contests, singles, waiting };
}

function draft_rollOne(player) {
  const info = draft_getCurrentContests();
  const contest = info.contests.find(c => c.player === player);
  if (!contest) throw new Error('対象プレイヤーの競合が見つかりません: ' + player);

  const names = contest.captains.map(c => c.name);
  const prios = contest.captains.map(c => Number(c.prio || 0));

  const maxPrio = Math.max.apply(null, prios);
  const tiedIdx = prios.map((p,i)=> p===maxPrio ? i : -1).filter(i=>i>=0);
  let winnerIdx;

  if (tiedIdx.length === 1) {
    winnerIdx = tiedIdx[0];
  } else {
    const pick = Math.floor(Math.random() * tiedIdx.length);
    winnerIdx = tiedIdx[pick];
  }

  const winner = names[winnerIdx];
  const tiedNames = tiedIdx.map(i => names[i]);
  const wheelBase = (tiedIdx.length > 1) ? tiedNames : names;

  const wheel = buildRouletteArrayEqual_(wheelBase);
  // 【修正】冗長な条件分岐を削除
  const stopPos = wheel.indexOf(winner);
  const cycles = 3;
  const stopIndex = cycles * wheel.length + (stopPos >= 0 ? stopPos : 0);

  return {
    player,
    winner,
    names,
    prios,
    wheel,
    stopIndex,
    perItemPx: 120,
    cycles,
    isTie: tiedIdx.length > 1
  };
}

function draft_commitPick(player, winner, methodLabel) {
  if (isPlayerPicked_(player)) {
    logEvent('COMMIT_SKIP_ALREADY', `player=${player}`);
    return false;
  }
  const round = getRound_(), sub = getSubround_();
  const contestants = listContestantsForPlayer_(round, sub, player);
  const shP = ensureSheet_(CFG.SHEET.PICKS, ['Round','Captain','Player','Method','Timestamp']);
  shP.appendRow([round, winner, player, methodLabel || `抽選(第${sub}希望)`, new Date()]);
  // 【追加】キャッシュを無効化
  invalidateCache_('picks');

  if (contestants && contestants.length >= 2) {
    const caps = readCaptains();
    const losers = contestants.filter(name => name !== winner);
    const changed = [];
    losers.forEach(name => {
      const row = caps.find(x => x.name === name);
      if (row) {
        row.prio = Number(row.prio || 0) + 1;
        changed.push({ name, prio: row.prio });
      }
    });
    writeCaptains(caps);
    logEvent('PRIO_UP', JSON.stringify({ player, losers: changed }));
  }

  rebuildTeams_();
  logEvent('COMMIT_PICK', JSON.stringify({ round, sub, player, winner }));
  // 【追加】データ更新タイムスタンプを更新
  updateDataTimestamp_();
  
  // 【追加】手動モードでも、このサブラウンドの全ての競合が解決された場合は次に進む
  if (getManualMode_()) {
    const info = draft_getCurrentContests();
    // 単独指名も競合もない場合（全て解決された場合）、次のサブラウンド/ラウンドに進む
    if (info.singles.length === 0 && info.contests.length === 0) {
      draft_advanceAfterManual();
      updateDataTimestamp_(); // サブラウンド/ラウンドが進んだのでタイムスタンプを更新
    }
  }
  
  return true;
}

function draft_commitSingles() {
  const info = draft_getCurrentContests();
  const round = getRound_(), sub = getSubround_();
  const shP = ensureSheet_(CFG.SHEET.PICKS, ['Round','Captain','Player','Method','Timestamp']);

  // 【最適化】既に取得済みのpickedSetを使用（readPicks()を再呼び出ししない）
  const pickedSet = new Set(readPicks().map(p => String(p.player)));
  
  // 【最適化】バッチ処理で一度に追加（appendRow()を個別に呼ぶより高速）
  const picksToAdd = [];
  info.singles.forEach(s => {
    if (!pickedSet.has(String(s.player))) {
      picksToAdd.push([round, s.captain, s.player, `通常(第${sub}希望)`, new Date()]);
    }
  });

  if (picksToAdd.length > 0) {
    const lastRow = shP.getLastRow();
    shP.getRange(lastRow + 1, 1, picksToAdd.length, 5).setValues(picksToAdd);
    // 【追加】キャッシュを無効化
    invalidateCache_('picks');
  }

  // 【最適化】rebuildTeams_()を実行（バッチ処理で高速化済み）
  rebuildTeams_();

  logEvent('COMMIT_SINGLES', `count=${picksToAdd.length}`);
  // 【追加】データ更新タイムスタンプを更新
  if (picksToAdd.length > 0) {
    updateDataTimestamp_();
  }
  return picksToAdd.length;
}

function draft_advanceAfterManual() {
  const round=getRound_(), sub=getSubround_();
  const eligible=getEligibleCaptains_();
  const picksThis = readPicks().filter(p=>p.round===round);
  const pickedCaps = new Set(picksThis.map(p=>p.captain));

  const unassigned = eligible.filter(c=>!pickedCaps.has(c));
  const stillPlayers = remainingPlayers();
  const form = getForm_();

  // 【追加】全ての指名が終わった場合の処理
  if (stillPlayers.length === 0) {
    form.setAcceptingResponses(false);
    setOpen_(0);
    logEvent('DRAFT_COMPLETE', '全ての指名が終了しました');
    throw new Error('全ての指名が終了しました。ドラフトは完了です。');
  }

  if (unassigned.length>0 && stillPlayers.length>0) {
    setSubround_(sub+1);
    setEligibleCaptains_(unassigned);
    updateCaptainChoices_(unassigned);
    updatePlayerChoices_();
    form.setAcceptingResponses(true);
    form.setDescription(`${CFG.FORM.DESC}\nRound:${round} Sub:${getSubround_()}`);
    logEvent('SUBROUND_OPEN', `Round ${round} sub=${getSubround_()} opened (manual)`);
  } else {
    setRound_(round+1);
    setSubround_(1);
    setEligibleCaptains_(readCaptains().map(c=>c.name));
    updateCaptainChoices_(getEligibleCaptains_());
    updatePlayerChoices_();
    form.setAcceptingResponses(true);
    form.setDescription(`${CFG.FORM.DESC}\nRound:${getRound_()} Sub:${getSubround_()}`);
    logEvent('ROUND_OPEN', `Round ${getRound_()} sub=${getSubround_()} opened (manual)`);
  }
  // 【追加】サブラウンド/ラウンドが進んだのでタイムスタンプを更新
  updateDataTimestamp_();
}

function weightedChoice_(weights){ const total=weights.reduce((a,b)=>a+b,0); let r=Math.random()*total; for(let i=0;i<weights.length;i++){ if((r-=weights[i])<0) return i; } return weights.length-1; }
function rebuildTeams_(){ 
  const sh=ensureSheet_(CFG.SHEET.TEAMS,['Team','Member']); 
  if(sh.getLastRow()>1) sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).clearContent(); 
  const capRows=readCaptains(); 
  const teamByCaptain=new Map(capRows.map(c=>[c.name,c.team])); 
  
  // 【最適化】バッチ処理で一度に追加（appendRow()を個別に呼ぶより高速）
  const picks = readPicks();
  if (picks.length > 0) {
    const rows = picks.map(p => [teamByCaptain.get(p.captain) || p.captain, p.player]);
    sh.getRange(2, 1, rows.length, 2).setValues(rows);
  }
}

function admin_syncChoices(){ updatePlayerChoices_(); updateCaptainChoices_(getEligibleCaptains_()); }
function admin_forceResolve(){ 
  const round=getRound_(); const sub=getSubround_(); const eligible=getEligibleCaptains_(); 
  // 【変更】Requestsシートではなく、フォームの回答シートから読み取る
  const respSh = getLinkedResponseSheet_();
  const sentCaps = new Set();
  if (respSh) {
    const header = respSh.getRange(1, 1, 1, respSh.getLastColumn()).getValues()[0].map(String);
    const colRound = header.findIndex(h => /^round$/i.test(h.trim())) + 1;
    const colSub = header.findIndex(h => /^sub$/i.test(h.trim())) + 1;
    const colCap = header.findIndex(h => /キャプテン|captain/i.test(h)) + 1;
    if (colRound && colSub && colCap) {
      const lastRow = respSh.getLastRow();
      if (lastRow > 1) {
        const vs = respSh.getRange(2, 1, lastRow - 1, respSh.getLastColumn()).getValues();
        vs.forEach(row => {
          const rowRound = Number(row[colRound - 1]);
          const rowSub = Number(row[colSub - 1]);
          const rowCaptain = String(row[colCap - 1] || '').trim();
          if (rowRound === round && rowSub === sub && rowCaptain) {
            sentCaps.add(rowCaptain);
          }
        });
      }
    }
  }
  // 【削除】Requestsシートへの書き込みは不要（フォームの回答シートに直接記録されるため）
  // const missing=eligible.filter(c=>!sentCaps.has(c));
  // if(missing.length>0){ ... }
  resolveIfReady_(); 
}
function admin_setRound(n){ setRound_(Number(n)); logEvent('ADMIN','set round: '+n); }
function admin_setSubround(n){ setSubround_(Number(n)); logEvent('ADMIN','set subround: '+n); }
function admin_open(){ openRound(); }
function admin_close(){ closeRound(); }
function admin_setEligibleCaptains(csv){ const names=String(csv||'').split(',').map(s=>s.trim()).filter(Boolean); setEligibleCaptains_(names); updateCaptainChoices_(names); logEvent('ADMIN','set eligible captains: '+JSON.stringify(names)); }
function admin_resolveNow(){ try{ resolveIfReady_(); logEvent('ADMIN','resolveNow called'); }catch(e){ logEvent('ADMIN_ERR','resolveNow: '+e); throw e; } }
function admin_ping(){ return 'ok'; }

function admin_rebuildOnSubmitTrigger(){ ScriptApp.getProjectTriggers().forEach(t=>ScriptApp.deleteTrigger(t)); ScriptApp.newTrigger('onFormSubmit').forForm(getForm_()).onFormSubmit().create(); logEvent('TRIGGER_REBUILT','onFormSubmit re-attached to current FORM_ID'); }

// 【削除】Requestsシートがなくなったため、補完機能は不要
// function admin_backfillRequestsFromResponses(){
//   ... (関数全体を削除)
// }

function admin_debugWaiting(){ 
  const round=getRound_(); const sub=getSubround_(); const eligible=new Set(getEligibleCaptains_()); 
  // 【変更】Requestsシートではなく、フォームの回答シートから読み取る
  const respSh = getLinkedResponseSheet_();
  const submitted = new Set();
  if (respSh) {
    const header = respSh.getRange(1, 1, 1, respSh.getLastColumn()).getValues()[0].map(String);
    const colRound = header.findIndex(h => /^round$/i.test(h.trim())) + 1;
    const colSub = header.findIndex(h => /^sub$/i.test(h.trim())) + 1;
    const colCap = header.findIndex(h => /キャプテン|captain/i.test(h)) + 1;
    if (colRound && colSub && colCap) {
      const lastRow = respSh.getLastRow();
      if (lastRow > 1) {
        const vs = respSh.getRange(2, 1, lastRow - 1, respSh.getLastColumn()).getValues();
        vs.forEach(row => {
          const rowRound = Number(row[colRound - 1]);
          const rowSub = Number(row[colSub - 1]);
          const rowCaptain = String(row[colCap - 1] || '').trim();
          if (rowRound === round && rowSub === sub && rowCaptain) {
            submitted.add(rowCaptain);
          }
        });
      }
    }
  }
  const waiting=[...eligible].filter(c=>!submitted.has(c)); 
  logEvent('DEBUG_WAITING',`round=${round}, sub=${sub}, waiting=${JSON.stringify(waiting)}`); 
}
function admin_debugCounts(){ 
  const round=getRound_(); const sub=getSubround_(); const eligible=getEligibleCaptains_(); 
  // 【変更】Requestsシートではなく、フォームの回答シートから読み取る
  const respSh = getLinkedResponseSheet_();
  let got = 0;
  if (respSh) {
    const header = respSh.getRange(1, 1, 1, respSh.getLastColumn()).getValues()[0].map(String);
    const colRound = header.findIndex(h => /^round$/i.test(h.trim())) + 1;
    const colSub = header.findIndex(h => /^sub$/i.test(h.trim())) + 1;
    const colCap = header.findIndex(h => /キャプテン|captain/i.test(h)) + 1;
    if (colRound && colSub && colCap) {
      const lastRow = respSh.getLastRow();
      if (lastRow > 1) {
        const vs = respSh.getRange(2, 1, lastRow - 1, respSh.getLastColumn()).getValues();
        const seen = new Set();
        vs.forEach(row => {
          const rowRound = Number(row[colRound - 1]);
          const rowSub = Number(row[colSub - 1]);
          const rowCaptain = String(row[colCap - 1] || '').trim();
          if (rowRound === round && rowSub === sub && rowCaptain && !seen.has(rowCaptain)) {
            seen.add(rowCaptain);
            got++;
          }
        });
      }
    }
  }
  logEvent('DEBUG_COUNTS', JSON.stringify({round,sub,need:eligible.length,got,eligible})); 
}

// 【削除】Requestsシートがなくなったため、矛盾確認機能は不要
// function admin_compareResponsesAndRequests(){
//   ... (関数全体を削除)
// }

function admin_resetAll() {
  let form = null, wasOpen = false;
  try {
    form = getForm_();
    wasOpen = form.isAcceptingResponses();
    if (wasOpen) form.setAcceptingResponses(false);
  } catch (_) {}

  // 【追加】既存のフォームを削除（ゴミ箱へ移動）
  const existingFormId = getConfigMap()[CFG.CFG_KEYS.FORM_ID];
  if (existingFormId) {
    try {
      deleteForm_(existingFormId);
    } catch (e) {
      logEvent('ADMIN_RESET_ERR', `フォーム削除エラー: ${e}`);
    }
  }
  
  putConfig(CFG.CFG_KEYS.ROUND, 1);
  putConfig(CFG.CFG_KEYS.SUBROUND, 1);
  putConfig(CFG.CFG_KEYS.ELIGIBLE_CAPTAINS, '[]');
  putConfig(CFG.CFG_KEYS.IS_OPEN, 0);
  putConfig(CFG.CFG_KEYS.SETUP_COMPLETE, '0'); // ドラフト準備完了フラグをクリア
  putConfig(CFG.CFG_KEYS.FORM_ID, ''); // 【追加】FORM_IDを削除（フォームを使い捨てにする）
  
  // 【追加】LAST_BACKFILL_*キーを全て削除（無限増殖を防ぐ）
  const cfgSh = ensureConfigSheet_();
  if (cfgSh.getLastRow() > 1) {
    const cfgData = cfgSh.getDataRange().getValues();
    const cfgMap = new Map(cfgData.slice(1).map(r => [String(r[0]), r[1]]));
    const keysToDelete = [];
    cfgMap.forEach((value, key) => {
      if (key.startsWith('LAST_BACKFILL_')) {
        keysToDelete.push(key);
      }
    });
    if (keysToDelete.length > 0) {
      keysToDelete.forEach(key => cfgMap.delete(key));
      const newCfgData = [['Key', 'Value'], ...Array.from(cfgMap.entries())];
      cfgSh.getRange(1, 1, newCfgData.length, 2).setValues(newCfgData);
      // 削除した行をクリア
      if (cfgSh.getLastRow() > newCfgData.length) {
        cfgSh.getRange(newCfgData.length + 1, 1, cfgSh.getLastRow() - newCfgData.length, 2).clearContent();
      }
    }
  }

  // 各シートをクリア（Logシートはヘッダーを確保）
  // 【変更】Requestsシートは不要になったため削除
  [CFG.SHEET.PICKS, CFG.SHEET.TEAMS].forEach(name => {
    const sh = ensureSheet_(name);
    if (sh.getLastRow() > 1) {
      sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent();
    }
  });
  
  // Logシートは特別処理：全データをクリアしてヘッダーを確実に設定
  const logSh = ensureSheet_(CFG.SHEET.LOG, ['Timestamp','Event','Detail']);
  if (logSh.getLastRow() > 0) {
    // 全データをクリア（ヘッダー含む）
    logSh.clear();
  }
  // ヘッダーを設定
  logSh.getRange(1, 1, 1, 3).setValues([['Timestamp','Event','Detail']]);

  resetCaptainsPrio_(0);

  // 【追加】残っている古い回答シートを全て削除（全データリセットの一部として）
  // 注: deleteForm_()で既にリンク解除とシート削除が実行されているが、
  // 念のため残っているシートも削除
  try {
    const deletedCount = deleteAllFormResponseSheets_();
    if (deletedCount > 0) {
      logEvent('ADMIN_RESET', `残っていた古い回答シート ${deletedCount} 件を削除しました`);
    }
  } catch (e) {
    logEvent('ADMIN_RESET_ERR', '回答シート削除エラー: ' + e);
  }

  // 【修正】全リセット時はフォームを開かない（手動で開く必要がある）
  // try {
  //   if (form && wasOpen) {
  //     form.setAcceptingResponses(true);
  //     form.setDescription(`${CFG.FORM.DESC}\nRound:${getRound_()} Sub:${getSubround_()}`);
  //   }
  // } catch (_) {}

  logEvent('ADMIN_RESET', '初期化完了（Prio=0 / データクリア / フォーム削除（ゴミ箱へ移動） / フォームID削除 / 回答シート削除）');
  
  // タイムスタンプ更新はputConfig()で自動的に実行される（SETUP_COMPLETEの更新時に）
}

function admin_openDrawUI(){
  setManualMode_(true);
  const html = HtmlService.createHtmlOutputFromFile('DrawUI')
               .setTitle('抽選コントロール')
               .setWidth(420);
  SpreadsheetApp.getUi().showSidebar(html);
}

function admin_compactResponseSheet() {
  const sh = getLinkedResponseSheet_();
  clearAndCompactSheet_(sh);
}

function admin_openAdminUI(){
  const html = HtmlService.createHtmlOutputFromFile('AdminUI')
    .setTitle('運営パネル')
    .setWidth(420);
  SpreadsheetApp.getUi().showSidebar(html);
}

function admin_getStatus_fast(){
  const out = { round: 1, sub: 1, form: { accepting: false, formUrl:'', editUrl:'', respName:'', isReady: false }, lastUpdateTimestamp: '' };
  try { out.round = getRound_(); } catch(_) {}
  try { out.sub   = getSubround_(); } catch(_) {}
  
  // 【修正】FORM_IDが存在するかチェックしてからgetForm_()を呼び出す
  const formId = getConfigMap()[CFG.CFG_KEYS.FORM_ID];
  if (formId) {
    try {
      const f = getForm_();
      // フォームが取得でき、かつドラフト準備完了フラグが設定されている場合のみ完了
      const setupComplete = String(getConfigMap()[CFG.CFG_KEYS.SETUP_COMPLETE]||'0') === '1';
      out.form.isReady = setupComplete;
      // 【修正】IS_OPENの値も考慮（IS_OPENが0の場合は停止中と表示）
      const isOpen = String(getConfigMap()[CFG.CFG_KEYS.IS_OPEN]||'0') === '1';
      const formAccepting = !!f.isAcceptingResponses();
      out.form.accepting = isOpen && formAccepting; // 両方がtrueの場合のみ受付中
      try {
        const publishedUrl = f.getPublishedUrl();
        out.form.formUrl   = publishedUrl || '';
      } catch(urlErr) {
        logEvent('STATUS_FAST_URL_ERR', `公開URL取得エラー: ${urlErr}`);
      }
      // 編集URLはフォームIDから構築
      try {
        const formIdFromForm = f.getId();
        out.form.editUrl   = formIdFromForm ? `https://docs.google.com/forms/d/${formIdFromForm}/edit` : '';
      } catch(idErr) {
        logEvent('STATUS_FAST_ID_ERR', `フォームID取得エラー: ${idErr}`);
      }
    } catch(e) {
      // フォームが存在しない、またはアクセスできない場合のみエラーログを記録
      // （FORM_IDが存在するのに取得できない場合は異常）
      logEvent('STATUS_FAST_ERR', `フォーム取得エラー: ${e}`);
      out.form.isReady = false; // フォームが取得できない = ドラフト準備が未完了
    }
  } else {
    // FORM_IDが存在しない場合は正常な状態（全データリセット後など）
    // エラーログを記録しない
    out.form.isReady = false; // フォームが設定されていない = ドラフト準備が未完了
  }
  
  try { out.lastUpdateTimestamp = getDataTimestamp_(); } catch(_) {}
  return out;
}

function admin_getStatus(){
  const res = { round: 1, sub: 1, eligibleCount: 0, eligible: [], submittedCount: 0,
                form: { accepting:false, formUrl:'', editUrl:'', respName:'', isReady: false },
                manualMode: false, captains: [], errors: [] };

  try { res.round = getRound_(); } catch(e){ res.errors.push('round:'+e); }
  try { res.sub   = getSubround_(); } catch(e){ res.errors.push('sub:'+e); }

  try {
    const f = getForm_();
    // フォームが取得でき、かつドラフト準備完了フラグが設定されている場合のみ完了
    const setupComplete = String(getConfigMap()[CFG.CFG_KEYS.SETUP_COMPLETE]||'0') === '1';
    res.form.isReady = setupComplete;
    // 【修正】IS_OPENの値も考慮（IS_OPENが0の場合は停止中と表示）
    const isOpen = String(getConfigMap()[CFG.CFG_KEYS.IS_OPEN]||'0') === '1';
    const formAccepting = !!f.isAcceptingResponses();
    res.form.accepting = isOpen && formAccepting; // 両方がtrueの場合のみ受付中
    res.form.formUrl   = f.getPublishedUrl() || '';
    // 編集URLはフォームIDから構築
    const formId = f.getId();
    res.form.editUrl   = formId ? `https://docs.google.com/forms/d/${formId}/edit` : '';
  } catch(e){ 
    res.errors.push('form:'+e);
    res.form.isReady = false; // フォームが取得できない = ドラフト準備が未完了
  }

  try {
    const rsh = getLinkedResponseSheet_?.();
    res.form.respName = rsh ? rsh.getName() : '';
  } catch(e){ res.errors.push('resp:'+e); }

  try {
    const caps = readCaptains?.() || [];
    res.captains = caps.map(c=>({name:c.name, team:c.team, prio:Number(c.prio||0)}));
  } catch(e){ res.errors.push('caps:'+e); }

  try {
    const elig = getEligibleCaptains_?.() || [];
    res.eligible = elig;
    res.eligibleCount = elig.length;
  } catch(e){ res.errors.push('elig:'+e); }

  try {
    // 【変更】Requestsシートではなく、フォームの回答シートから読み取る
    const respSh = getLinkedResponseSheet_();
    res.submittedCount = 0;
    if (respSh) {
      const header = respSh.getRange(1, 1, 1, respSh.getLastColumn()).getValues()[0].map(String);
      const colRound = header.findIndex(h => /^round$/i.test(h.trim())) + 1;
      const colSub = header.findIndex(h => /^sub$/i.test(h.trim())) + 1;
      const colCap = header.findIndex(h => /キャプテン|captain/i.test(h)) + 1;
      if (colRound && colSub && colCap) {
        const lastRow = respSh.getLastRow();
        if (lastRow > 1) {
          const vs = respSh.getRange(2, 1, lastRow - 1, respSh.getLastColumn()).getValues();
          const seen = new Set();
          vs.forEach(row => {
            const rowRound = Number(row[colRound - 1]);
            const rowSub = Number(row[colSub - 1]);
            const rowCaptain = String(row[colCap - 1] || '').trim();
            if (rowRound === res.round && rowSub === res.sub && rowCaptain && !seen.has(rowCaptain)) {
              seen.add(rowCaptain);
              res.submittedCount++;
            }
          });
        }
      }
    }
  } catch(e){ res.errors.push('req:'+e); }

  try { res.manualMode = !!(typeof getManualMode_==='function' ? getManualMode_() : false); }
  catch(e){ res.errors.push('manual:'+e); }

  try { res.lastUpdateTimestamp = getDataTimestamp_(); } catch(e){ res.errors.push('lastUpdateTimestamp:'+e); }

  return res;
}

function admin_setManualMode(on){
  if (typeof setManualMode_ !== 'function') throw new Error('手動モード機能が未導入です');
  setManualMode_(!!on);
  logEvent('ADMIN','manualMode='+(on?'ON':'OFF'));
  return !!on;
}

function admin_setRoundSub(round, sub){
  setRound_(Number(round||1));
  setSubround_(Number(sub||1));
  logEvent('ADMIN', `set round=${getRound_()} sub=${getSubround_()}`);
  return {round:getRound_(), sub:getSubround_()};
}

function admin_setEligibleCaptainsCSV(csv){
  const names = String(csv||'').split(',').map(s=>s.trim()).filter(Boolean);
  setEligibleCaptains_(names);
  updateCaptainChoices_(names);
  logEvent('ADMIN','set eligible captains: '+JSON.stringify(names));
  return names;
}

function admin_diag(){
  const out = { ok:true, notes:[], form:{}, sheets:{}, triggers:[] };
  try{ out.form.id = getConfigMap()[CFG.CFG_KEYS.FORM_ID] || ''; }catch(e){ out.ok=false; out.notes.push('CONFIG/FORM_ID:'+e); }
  try{
    const f=getForm_();
    out.form.accepting = f.isAcceptingResponses();
    out.form.editUrl = f.getEditUrl && f.getEditUrl();
  }catch(e){ out.ok=false; out.notes.push('FORM:'+e); }
  try{
    const r = getLinkedResponseSheet_();
    out.sheets.response = r ? r.getName() : '(none)';
  }catch(e){ out.ok=false; out.notes.push('RESPONSE_SHEET:'+e); }
  try{
    const req = ensureSheet_(CFG.SHEET.REQUESTS);
    out.sheets.requestsRows = Math.max(0, req.getLastRow()-1);
  }catch(e){ out.ok=false; out.notes.push('REQUESTS:'+e); }
  try{
    out.triggers = ScriptApp.getProjectTriggers().map(t=>({handler:t.getHandlerFunction(), src:t.getTriggerSource()}));
  }catch(e){ out.notes.push('TRIGGERS:'+e); }
  return out;
}

function admin_openRoundSafe(){ openRound(); return true; }
function admin_closeRoundSafe(){ closeRound?.(); return true; }
function admin_rebuildTrigger(){ admin_rebuildOnSubmitTrigger(); return true; }
function admin_backfill(){ admin_backfillRequestsFromResponses(); return true; }
function admin_resolve(){ admin_resolveNow(); return true; }
function admin_compact(){ admin_compactResponseSheet(); return true; }
function admin_resetAllSafe(){ admin_resetAll(); return true; }
function admin_syncChoicesSafe(){ admin_syncChoices?.(); return true; }
// 【削除】Requestsシートがなくなったため、矛盾確認機能は不要
// function admin_compare(){ return admin_compareResponsesAndRequests(); }

// 【追加】Captainsシートを元のシートからインポート
function admin_importCaptains(){
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    return { success: false, error: 'ロック取得失敗（他の処理が実行中です）' };
  }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // シート名と列マッピングを固定値で使用
    const sourceSheetNameTrimmed = 'Entry';
    const sourceSheet = ss.getSheetByName(sourceSheetNameTrimmed);
    
    if (!sourceSheet) {
      lock.releaseLock();
      return { success: false, error: `シート「${sourceSheetNameTrimmed}」が見つかりません` };
    }
    
    // 自動検出モード（固定）
    lock.releaseLock(); // 自動検出関数内でロックを取得するため、ここで解放
    return importCaptainsAuto_(sourceSheet, sourceSheetNameTrimmed);
  } catch(e) {
    logEvent('IMPORT_CAPTAINS_ERROR', e.toString());
    lock.releaseLock();
    return { success: false, error: e.toString() };
  }
}

// 【追加】Captainsシートを自動検出モードでインポート（キャプテン列に値がある行をキャプテンとして扱う）
function importCaptainsAuto_(sourceSheet, sourceSheetName){
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    return { success: false, error: 'ロック取得失敗（他の処理が実行中です）' };
  }
  try {
    const dataRange = sourceSheet.getDataRange();
    const values = dataRange.getValues();
    
    if (values.length < 2) {
      return { success: false, error: 'データがありません（ヘッダー行のみ）' };
    }
    
    // 列インデックスを自動検出
    const playerNameColIdx = 0; // A列: プレイヤーネーム
    const xpColIdx = 2; // C列: エリアの最高XP
    const captainColIdx = 6; // G列: キャプテン（チーム名）
    
    // キャプテン列に値がある行を抽出
    const captainsMap = new Map(); // チーム名 -> {name, xp}のマップ
    
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const playerName = String(row[playerNameColIdx] || '').trim();
      const captainValue = String(row[captainColIdx] || '').trim();
      
      // プレイヤー名が空の場合はスキップ
      if (!playerName) continue;
      
      // キャプテン列に値がある場合、そのプレイヤーがキャプテンで、値がチーム名
      if (captainValue) {
        // 同じチーム名が既にある場合はスキップ（最初の1人をキャプテンとする）
        if (!captainsMap.has(captainValue)) {
          const xp = Number(row[xpColIdx] || 0);
          captainsMap.set(captainValue, { name: playerName, xp: xp });
        }
      }
    }
    
    if (captainsMap.size === 0) {
      return { success: false, error: 'キャプテン列（G列）に値がある行が見つかりません' };
    }
    
    // Captainsシートに書き込む
    const captains = [];
    for (const [teamName, captainData] of captainsMap) {
      captains.push({ name: captainData.name, team: teamName, prio: 0, xp: captainData.xp });
    }
    
    writeCaptains(captains);
    logEvent('IMPORT_CAPTAINS_AUTO', `シート「${sourceSheetName}」から${captains.length}件を自動検出でインポート（チーム名・XP含む）`);
    
    const result = { success: true, count: captains.length };
    lock.releaseLock();
    return result;
  } catch(e) {
    logEvent('IMPORT_CAPTAINS_AUTO_ERROR', e.toString());
    lock.releaseLock();
    return { success: false, error: e.toString() };
  }
}

// 【追加】Playersシートを元のシートからインポート
function admin_importPlayers(){
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    return { success: false, error: 'ロック取得失敗（他の処理が実行中です）' };
  }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // シート名と列マッピングを固定値で使用
    const sourceSheetNameTrimmed = 'Entry';
    const sourceSheet = ss.getSheetByName(sourceSheetNameTrimmed);
    
    if (!sourceSheet) {
      lock.releaseLock();
      return { success: false, error: `シート「${sourceSheetNameTrimmed}」が見つかりません` };
    }
    
    // 自動検出モード（固定）
    lock.releaseLock(); // 自動検出関数内でロックを取得するため、ここで解放
    return importPlayersAuto_(sourceSheet, sourceSheetNameTrimmed);
  } catch(e) {
    logEvent('IMPORT_PLAYERS_ERROR', e.toString());
    lock.releaseLock();
    return { success: false, error: e.toString() };
  }
}

// 【追加】Playersシートを自動検出モードでインポート（キャプテン列が空の場合のみ取得）
function importPlayersAuto_(sourceSheet, sourceSheetName){
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    return { success: false, error: 'ロック取得失敗（他の処理が実行中です）' };
  }
  try {
    const dataRange = sourceSheet.getDataRange();
    const values = dataRange.getValues();
    
    if (values.length < 2) {
      return { success: false, error: 'データがありません（ヘッダー行のみ）' };
    }
    
    // デバッグ用：最初の数行をログに記録
    try {
      const logSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Log');
      if (logSh && values.length > 1) {
        const sampleRow = values[1]; // 2行目（最初のデータ行）
        logSh.appendRow([new Date(), 'IMPORT_PLAYERS_DEBUG', `row[0]=${sampleRow[0]}, row[2]=${sampleRow[2]}, row[3]=${sampleRow[3]}, row[5]=${sampleRow[5]}, row[6]=${sampleRow[6]}`]);
      }
    } catch(_) {}
    
    // 列インデックスを自動検出（0ベース）
    // A列=0, B列=1, C列=2, D列=3, E列=4, F列=5, G列=6
    const playerNameColIdx = 0; // A列: プレイヤーネーム
    const xpColIdx = 2; // C列: エリアの最高XP
    const weaponColIdx = 3; // D列: 持ちブキ
    const commentColIdx = 5; // F列: 自己紹介
    const captainColIdx = 6; // G列: キャプテン
    
    // Playersシートに書き込む（キャプテン列が空の場合のみ）
    const players = [];
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      // 行の長さを確認（列が不足している場合の対策）
      if (row.length <= playerNameColIdx) continue;
      
      const name = String(row[playerNameColIdx] || '').trim();
      const captainValue = row.length > captainColIdx ? String(row[captainColIdx] || '').trim() : '';
      
      // プレイヤー名が空の場合はスキップ
      if (!name) continue;
      
      // キャプテン列に値がある場合はスキップ（キャプテンはPlayersシートに含めない）
      if (captainValue) continue;
      
      // 各列の値を取得（列が存在するか確認）
      const xp = row.length > xpColIdx ? String(row[xpColIdx] || '').trim() : '';
      const weapon = row.length > weaponColIdx ? String(row[weaponColIdx] || '').trim() : '';
      const comment = row.length > commentColIdx ? String(row[commentColIdx] || '').trim() : '';
      
      // Playersシートの列順: Name(0), Status(1), エリアXP(2), 持ちブキ(3), コメント(4)
      // Statusは'active'で初期化
      players.push([name, 'active', xp, weapon, comment]);
    }
    
    if (players.length === 0) {
      return { success: false, error: '有効なプレイヤーデータがありません（キャプテン列が空の行が見つかりません）' };
    }
    
    const playersSh = ensureSheet_(CFG.SHEET.PLAYERS, ['Name','Status','エリアXP','持ちブキ','コメント']);
    // 既存のデータをクリア（ヘッダーを除く）
    const lastRow = playersSh.getLastRow();
    if (lastRow > 1) {
      playersSh.getRange(2, 1, lastRow - 1, 5).clearContent();
    }
    // 新しいデータを書き込む
    if (players.length > 0) {
      playersSh.getRange(2, 1, players.length, 5).setValues(players);
      // 【追加】キャッシュを無効化
      invalidateCache_('players');
    }
    
    logEvent('IMPORT_PLAYERS_AUTO', `シート「${sourceSheetName}」から${players.length}件を自動検出でインポート（キャプテン列が空の行のみ、エリアXP/持ちブキ/コメント含む）`);
    
    const result = { success: true, count: players.length };
    lock.releaseLock();
    return result;
  } catch(e) {
    logEvent('IMPORT_PLAYERS_AUTO_ERROR', e.toString());
    lock.releaseLock();
    return { success: false, error: e.toString() };
  }
}

// 【追加】列マッピング文字列をパース（例: "Captain列=A, Team列=B" または "Captain=A, Team=B"）
function parseColumnMapping_(mappingStr){
  const result = {};
  const parts = mappingStr.split(',');
  
  for (const part of parts) {
    // "Captain列=A" または "Captain=A" の形式に対応
    const match = part.match(/(\w+)(?:列)?\s*=\s*([A-Z0-9]+)/i);
    if (match) {
      const key = match[1].toLowerCase();
      const col = match[2].toUpperCase();
      result[key] = col;
    }
  }
  
  return result;
}

// 【追加】列指定（A, B, C...または1, 2, 3...またはヘッダー名）をインデックスに変換
function parseColumnIndex_(colSpec, headers){
  if (!colSpec) return -1;
  
  const spec = String(colSpec).trim().toUpperCase();
  
  // 列名（A, B, C...）の場合
  if (/^[A-Z]+$/.test(spec)) {
    let col = 0;
    for (let i = 0; i < spec.length; i++) {
      const charCode = spec.charCodeAt(i);
      if (charCode >= 65 && charCode <= 90) { // A-Z
        col = col * 26 + (charCode - 64);
      } else {
        return -1;
      }
    }
    return col - 1; // 0ベースのインデックス
  }
  
  // 列番号（1, 2, 3...）の場合
  const colNum = Number(spec);
  if (!isNaN(colNum) && colNum > 0) {
    return colNum - 1; // 0ベースのインデックス
  }
  
  // ヘッダー名で検索
  const colName = spec.toLowerCase();
  for (let i = 0; i < headers.length; i++) {
    if (String(headers[i]).toLowerCase() === colName) {
      return i;
    }
  }
  
  return -1;
}

// 【追加】WebアプリのURLを取得/設定するヘルパー関数
function admin_getControlCenterUrl(){
  try {
    // Configシートに保存されたURLを取得
    let savedUrl = getConfigMap()[CFG.CFG_KEYS.WEB_APP_URL] || '';
    
    // URLから既存のパラメータを削除（クリーンなベースURLを取得）
    if (savedUrl) {
      savedUrl = savedUrl.split('?')[0]; // クエリパラメータを削除
    }
    
    if (savedUrl) {
      // 保存されたURLを表示
      const response = SpreadsheetApp.getUi().alert(
        'WebアプリのURL:\n\n' + savedUrl + '\n\nこのURLをコピーして、チームメンバーに共有してください。\n（スプレッドシートへのアクセス権限が必要です）\n\nURLを更新しますか？',
        SpreadsheetApp.getUi().ButtonSet.YES_NO_CANCEL
      );
      
      if (response === SpreadsheetApp.getUi().Button.YES) {
        // URLを更新
        const newUrl = SpreadsheetApp.getUi().prompt(
          'WebアプリのURLを入力してください:\n\n（デプロイ手順:\n1. Apps Scriptエディタで「デプロイ」→「新しいデプロイ」\n2. 種類: 「ウェブアプリ」\n3. 実行ユーザー: 「アクセス権限を持つユーザー」\n4. アクセスできるユーザー: 「全員」\n5. 「デプロイ」をクリック\n6. 表示されたURLをコピーしてここに貼り付けてください）',
          SpreadsheetApp.getUi().ButtonSet.OK_CANCEL
        );
        
        if (newUrl.getSelectedButton() === SpreadsheetApp.getUi().Button.OK) {
          let url = newUrl.getResponseText().trim();
          // クエリパラメータを削除
          url = url.split('?')[0];
          if (url) {
            putConfig(CFG.CFG_KEYS.WEB_APP_URL, url);
            SpreadsheetApp.getUi().alert('URLを保存しました。\n\n' + url);
            return url;
          }
        }
      }
      
      return savedUrl;
    } else {
      // URLが保存されていない場合、入力してもらう
      const response = SpreadsheetApp.getUi().prompt(
        'WebアプリのURLを入力してください:\n\n（デプロイ手順:\n1. Apps Scriptエディタで「デプロイ」→「新しいデプロイ」\n2. 種類: 「ウェブアプリ」\n3. 実行ユーザー: 「アクセス権限を持つユーザー」\n4. アクセスできるユーザー: 「全員」\n5. 「デプロイ」をクリック\n6. 表示されたURLをコピーしてここに貼り付けてください）',
        SpreadsheetApp.getUi().ButtonSet.OK_CANCEL
      );
      
      if (response.getSelectedButton() === SpreadsheetApp.getUi().Button.OK) {
        let url = response.getResponseText().trim();
        // クエリパラメータを削除
        url = url.split('?')[0];
        if (url) {
          putConfig(CFG.CFG_KEYS.WEB_APP_URL, url);
          SpreadsheetApp.getUi().alert('URLを保存しました。\n\n' + url);
          return url;
        }
      }
      
      return null;
    }
  } catch(e) {
    SpreadsheetApp.getUi().alert('エラー: ' + e.toString());
    return null;
  }
}

function FORM_ACCEPTING(){
  try{
    const f = getForm_();
    return f.isAcceptingResponses() ? "🟢 受付中" : "⏸ 停止中";
  }catch(e){
    return "—";
  }
}

/**
 * Audience（配信用）を再構築（完全版）
 * 【バグ修正】枠3の式を修正（INDEXの第2引数を2→3に変更）
 */
function admin_buildAudienceSheet_streamGrid(){
  const ss = SpreadsheetApp.getActive();
  const name = 'Audience';
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);

  sh.clear({contentsOnly:true});
  sh.clearFormats();
  sh.setFrozenRows(5);
  try { sh.setFrozenColumns(0); } catch(_) {}

  sh.setColumnWidths(1, 7, 120);
  sh.setColumnWidth(1, 200);
  sh.setColumnWidth(2, 170);
  sh.setColumnWidth(3, 170);
  sh.setColumnWidth(4, 170);
  sh.setColumnWidth(5, 170);
  sh.setColumnWidth(6, 110);
  sh.setColumnWidth(7, 20);

  sh.setColumnWidth(9,  230);
  sh.setColumnWidth(10, 110);
  sh.setColumnWidth(11, 160);
  sh.setColumnWidth(12, 380);

  sh.setColumnWidth(15, 1);
  sh.getRange('O5').setValue('picked_names_hidden').setFontColor('#999');
  sh.getRange('O6').setFormula('=FILTER(Picks!C2:C, Picks!C2:C<>"")');
  sh.hideColumn(sh.getRange('O:O'));

  sh.getRange('A1').setValue('配信用ビュー（左：チームボード / 右：選手リスト）')
    .setFontSize(22).setFontWeight('bold');
  sh.getRange('A3').setValue('※ Players: A=名前, B=Status, C=エリアXP, D=持ちブキ, E=コメント / Status=activeのみ表示')
    .setFontColor('#666').setFontSize(10);

  sh.getRange('A5').setValue('チームボード').setFontWeight('bold').setFontSize(14);
  sh.getRange('A6').setValue('Captain').setFontWeight('bold');
  sh.getRange('B6').setValue('枠1').setFontWeight('bold');
  sh.getRange('C6').setValue('枠2').setFontWeight('bold');
  sh.getRange('D6').setValue('枠3').setFontWeight('bold');
  sh.getRange('E6').setValue('枠4').setFontWeight('bold');
  sh.getRange('F6').setValue('XP合計').setFontWeight('bold');

  sh.getRange('A7').setFormula('=FILTER(Captains!A2:A, Captains!A2:A<>"")');

  const capSheet = ss.getSheetByName('Captains') || ensureSheet_('Captains',["Captain","Team","Prio","XP"]);
  const capCount = Math.max(1, capSheet.getLastRow()-1);

  for (let i = 0; i < capCount + 30; i++) {
    const r = 7 + i;

    // 【元に戻す】FILTERの結果をそのまま使用（Picksシートのデータはラウンド順に並んでいる前提）
    sh.getRange(r, 2).setFormula('=IF($A' + r + '="","", IFERROR(INDEX(FILTER(Picks!$C:$C, Picks!$B:$B=$A' + r + '), 1), ""))');
    sh.getRange(r, 3).setFormula('=IF($A' + r + '="","", IFERROR(INDEX(FILTER(Picks!$C:$C, Picks!$B:$B=$A' + r + '), 2), ""))');
    // 【バグ修正】枠3は3番目を参照する必要がある（INDEXの第2引数を2→3に修正）
    sh.getRange(r, 4).setFormula('=IF($A' + r + '="","", IFERROR(INDEX(FILTER(Picks!$C:$C, Picks!$B:$B=$A' + r + '), 3), ""))');
    sh.getRange(r, 5).setFormula('=IF($A' + r + '="","", IFERROR(INDEX(FILTER(Picks!$C:$C, Picks!$B:$B=$A' + r + '), 4), ""))');

    sh.getRange(r, 6).setFormula(
      '=IF($A' + r + '="","", ' +
        'IFERROR(VLOOKUP($A' + r + ', Captains!$A:$D, 4, FALSE), 0)' +
        '+' +
        'SUM(' +
          'IFERROR(VLOOKUP($B' + r + ', Players!$A:$C, 3, FALSE), 0),' +
          'IFERROR(VLOOKUP($C' + r + ', Players!$A:$C, 3, FALSE), 0),' +
          'IFERROR(VLOOKUP($D' + r + ', Players!$A:$C, 3, FALSE), 0),' +
          'IFERROR(VLOOKUP($E' + r + ', Players!$A:$C, 3, FALSE), 0)' +
        ')' +
      ')'
    );
  }

  sh.getRange('A5:F6').setBorder(true,true,true,true,true,true);
  sh.getRange('A5:F6').setFontSize(12).setFontWeight('bold');
  sh.getRange('A7:F1000').setFontSize(14);
  sh.getRange('B7:E1000').setWrap(false);
  sh.getRange('F7:F1000').setNumberFormat('0');
  sh.setRowHeights(1, 1, 36);
  sh.setRowHeights(5, 2, 24);

  sh.getRange('I5').setValue('選手リスト').setFontWeight('bold').setFontSize(14);
  sh.getRange('I6').setValue('名前').setFontWeight('bold');
  sh.getRange('J6').setValue('エリアXP').setFontWeight('bold');
  sh.getRange('K6').setValue('持ちブキ').setFontWeight('bold');
  sh.getRange('L6').setValue('コメント').setFontWeight('bold');

  sh.getRange('I7').setFormula(
    '=FILTER({' +
      'Players!A2:A,' +
      'Players!C2:C,' +
      'Players!D2:D,' +
      'ARRAYFORMULA(SUBSTITUTE(Players!E2:E, CHAR(10), " "))' +
    '}, Players!A2:A<>"", LOWER(Players!B2:B)="active")'
  );

  sh.getRange('I5:L6').setBorder(true,true,true,true,true,true);
  sh.getRange('I5:L6').setFontSize(12).setFontWeight('bold');
  sh.getRange('I7:L1000').setFontSize(14);
  sh.getRange('I7:L1000').setWrap(false);
  sh.getRange('L7:L1000').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  const rules = sh.getConditionalFormatRules() || [];
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .setRanges([sh.getRange('I7:L1000')])
      .whenFormulaSatisfied('=COUNTIF($O:$O, $I7)>0')
      .setBackground('#1f1f1f')
      .setFontColor('#aaaaaa')
      .build()
  );
  sh.setConditionalFormatRules(rules);

  admin_polishAudienceLook_Final();

  try { logEvent('AUDIENCE_STREAM_GRID_BUILT','Audience stream grid built (Captain XP included)'); } catch(_) {}
}

function admin_polishAudienceLook_Final() {
  const sh = SpreadsheetApp.getActive().getSheetByName('Audience');
  if (!sh) throw new Error('Audience シートが見つかりません');

  try { sh.setTabColor('#111827'); } catch(_) {}
  try { sh.setHiddenGridlines(true); } catch(_) {}
  const lastRow = Math.max(1000, sh.getMaxRows());
  const lastCol = Math.max(15, sh.getMaxColumns());
  sh.getRange(1,1,lastRow,lastCol)
    .setFontFamily('Noto Sans JP')
    .setFontSize(12);

  const headerLeft  = sh.getRange('A5:F6');
  const headerRight = sh.getRange('I5:L6');
  [headerLeft, headerRight].forEach(r => {
    r.setBackground('#111827')
     .setFontColor('#F9FAFB')
     .setFontWeight('bold')
     .setHorizontalAlignment('center')
     .setVerticalAlignment('middle');
    r.setBorder(true,true,true,true,true,true,'#111827',SpreadsheetApp.BorderStyle.SOLID_THICK);
  });
  sh.getRange('A1').setFontSize(22).setFontWeight('bold').setFontColor('#111827');
  sh.getRange('A3').setFontSize(10).setFontColor('#6B7280');

  const leftAll = sh.getRange('A7:F1000');
  leftAll.setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange('A7:A1000').setHorizontalAlignment('left');
  sh.getRange('B7:E1000').setWrap(false).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  sh.getRange('F7:F1000').setNumberFormat('0');
  leftAll.setBorder(true,true,true,true,false,false,'#1F2937',SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  const rightAll = sh.getRange('I7:L1000');
  rightAll.setVerticalAlignment('middle').setWrap(false).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  sh.getRange('I7:I1000').setHorizontalAlignment('left');
  sh.getRange('J7:J1000').setHorizontalAlignment('right');
  sh.getRange('K7:K1000').setHorizontalAlignment('center');
  sh.getRange('L7:L1000').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  rightAll.setBorder(true,true,true,true,false,false,'#D1D5DB',SpreadsheetApp.BorderStyle.SOLID);

  sh.setColumnWidth(1,  200);
  sh.setColumnWidths(2, 4, 170);
  sh.setColumnWidth(6,  110);
  sh.setColumnWidth(9,  230);
  sh.setColumnWidth(10, 110);
  sh.setColumnWidth(11, 160);
  sh.setColumnWidth(12, 380);

  for (let r = 1; r <= 6; r++) sh.setRowHeight(r, 28);
  sh.setRowHeights(7, 500, 24);

  try { sh.getBandings().forEach(b => b.remove()); } catch(_) {}
  let rules = sh.getConditionalFormatRules() || [];
  rules = rules.filter(rule => {
    const a1s = rule.getRanges().map(r => r.getA1Notation()).join(',');
    return !(a1s.includes('A7:F1000'));
  });
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .setRanges([sh.getRange('A7:F1000')])
      .whenFormulaSatisfied('=ISEVEN(ROW())')
      .setBackground('#F3F4F6')
      .build()
  );
  sh.setConditionalFormatRules(rules);

  try { sh.setFrozenColumns(0); } catch(_) {}
  try { sh.setFrozenRows(5); } catch(_) {}

  logEvent('AUDIENCE_POLISHED_FINAL','Audience sheet polished (final)');
}

function getCaptainsDataRange_() {
  const sh = ensureSheet_('Captains', ['Captain','Team','Prio','XP']);
  const lr = sh.getLastRow();
  const lc = Math.max(4, sh.getLastColumn());
  if (lr < 2) return { sh, lr, lc, rng: null, vals: [] };
  const rng = sh.getRange(2, 1, lr - 1, lc);
  const vals = rng.getValues();
  return { sh, lr, lc, rng, vals };
}

function findCaptainRowByName_(name) {
  const { sh, vals } = getCaptainsDataRange_();
  const key = String(name).trim();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === key) return 2 + i;
  }
  return -1;
}

function setCaptainPrio_(name, prioValue) {
  const row = findCaptainRowByName_(name);
  if (row < 0) throw new Error('Captain not found: ' + name);
  const sh = SpreadsheetApp.getActive().getSheetByName('Captains');
  sh.getRange(row, 3).setValue(Number(prioValue)||0);
}

function addCaptainPrio_(name, delta) {
  const row = findCaptainRowByName_(name);
  if (row < 0) throw new Error('Captain not found: ' + name);
  const sh = SpreadsheetApp.getActive().getSheetByName('Captains');
  const cur = Number(sh.getRange(row, 3).getValue()) || 0;
  sh.getRange(row, 3).setValue(cur + Number(delta||0));
}

function setCaptainXP_(name, xpValue) {
  const row = findCaptainRowByName_(name);
  if (row < 0) throw new Error('Captain not found: ' + name);
  const sh = SpreadsheetApp.getActive().getSheetByName('Captains');
  sh.getRange(row, 4).setValue(Number(xpValue)||0);
}

function compactCaptains_() {
  const { sh, vals, lc } = getCaptainsDataRange_();
  const rows = vals.filter(r => String(r[0]).trim() !== '');
  if (rows.length > 0) {
    sh.getRange(2, 1, rows.length, lc).setValues(rows);
    const toClear = (vals.length - rows.length);
    if (toClear > 0) {
      sh.getRange(2 + rows.length, 1, toClear, lc).clearContent().clearFormat();
    }
  } else {
    if (vals.length > 0) {
      sh.getRange(2, 1, vals.length, lc).clearContent().clearFormat();
    }
  }
  const last = sh.getLastRow();
  const dataLast = 1 + rows.length;
  if (last > dataLast + 100) {
    sh.deleteRows(dataLast + 1, last - (dataLast + 1) + 1);
  }
}

/** すでにその Captain がその Player を Picks に確定済みか？ */
function isPickAlreadyRecorded_(round, captain, player) {
  const sh = ensureSheet_('Picks', ['Round','Captain','Player','Method','Timestamp']);
  const vs = sh.getDataRange().getValues();
  for (let i = 1; i < vs.length; i++) {
    const r = vs[i];
    // 【バグ修正】r[1]同士の比較ではなく、captainとの比較に修正
    if (Number(r[0]) === Number(round) &&
        String(r[1]).trim() === String(captain).trim() &&
        String(r[2]).trim() === String(player).trim()) {
      return true;
    }
  }
  return false;
}

function isPlayerAlreadyTaken_(player) {
  const sh = ensureSheet_(CFG.SHEET.PICKS, ['Round','Captain','Player','Method','Timestamp']);
  const vs = sh.getRange(2,3,Math.max(0,sh.getLastRow()-1),1).getValues();
  const key = String(player).trim();
  return vs.some(r => String(r[0]).trim() === key);
}

function applyPickSafe_(round, captain, player, method) {
  if (!round) round = Number(getRound_() || 1);
  if (isPlayerAlreadyTaken_(player)) {
    logEvent('PICK_SKIPPED_ALREADY_TAKEN', `${player}`);
    return false;
  }
  if (isPickAlreadyRecorded_(round, captain, player)) {
    logEvent('PICK_SKIPPED_DUP', `${round}/${captain}/${player}`);
    return false;
  }
  const sh = ensureSheet_(CFG.SHEET.PICKS, ['Round','Captain','Player','Method','Timestamp']);
  sh.appendRow([Number(round), String(captain), String(player), String(method||'draw'), new Date()]);
  // 【追加】キャッシュを無効化
  invalidateCache_('picks');
  logEvent('PICK_APPLIED', `${round}/${captain}/${player}/${method||'draw'}`);
  return true;
}

function admin_drawConfirmSafe(player, winnerCaptain, loserCaptains, round) {
  try {
    if (!player || !winnerCaptain) throw new Error('player / winnerCaptain は必須です');
    if (!Array.isArray(loserCaptains)) loserCaptains = [];

    if (isPlayerAlreadyTaken_(player)) {
      return { ok:true, message:`${player} は既に確定済み` };
    }

    applyPickSafe_(round, winnerCaptain, player, 'draw');

    loserCaptains.forEach(lc => {
      try { addCaptainPrio_(lc, 1); }
      catch (e) { logEvent('WARN_LOSER_PRIO_ADD_FAIL', `${lc}: ${e}`); }
    });

    compactCaptains_();

    logEvent('DRAW_CONFIRMED_SAFE', `player=${player}, winner=${winnerCaptain}, losers=${loserCaptains.join(',')}`);
    return { ok:true, message:'確定しました' };
  } catch (err) {
    logEvent('ERR_DRAW_CONFIRM_SAFE', String(err));
    return { ok:false, message:String(err) };
  }
}


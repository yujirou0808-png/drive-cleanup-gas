/**
 * ドライブ整理ツール（重複ファイル 最新のみ残す）
 *
 * 使い方は3ステップ。いきなり消えることはありません。
 *   STEP1_棚卸しレポート作成()      … 重複を洗い出して一覧シートを作る（何も削除しない）
 *   STEP2_チェックした行をゴミ箱へ() … 一覧でチェックを付けた行だけゴミ箱へ移動
 *   STEP3_ゴミ箱から戻す()          … 直前のSTEP2を取り消す（保険）
 *
 * ※「削除」ではなく「ゴミ箱へ移動」です。30日間は Drive のゴミ箱から復元できます。
 */

// ===================== 設定 =====================
var CONFIG = {
  // 対象ファイル種別。不要なものは false に
  対象_スプレッドシート: true,
  対象_Excelファイル:   false,  // .xlsx
  対象_ドキュメント:     false,  // Google ドキュメント（引継ぎ資料.md など）

  // 自分が所有者のファイルだけを対象にする（他人から共有されたファイルは触らない）
  自分の所有ファイルのみ: true,

  // ゴミ箱にあるファイルは最初から無視
  ゴミ箱は除外: true,

  // レポートの保存先フォルダ名（マイドライブ直下に作られます）
  レポート保存フォルダ: '_ドライブ整理',

  // 1回の実行で走査する最大件数（保険）
  最大走査件数: 5000
};
// ================================================


// ===================== STEP 1 =====================
function STEP1_棚卸しレポート作成() {
  var 開始 = new Date().getTime();
  var files = 走査する_(開始);

  if (files.length === 0) {
    Logger.log('対象ファイルが見つかりませんでした。CONFIG の対象種別を確認してください。');
    return;
  }

  // 正規化名でグループ化
  var groups = {};
  files.forEach(function (f) {
    var key = 正規化_(f.name);
    (groups[key] = groups[key] || []).push(f);
  });

  // 日付サフィックスを外した名前でもグループ化（参考表示用）
  var loose = {};
  files.forEach(function (f) {
    var key = 日付を外す_(正規化_(f.name));
    (loose[key] = loose[key] || []).push(f);
  });

  var 行 = [];
  var グループ番号 = 0;
  var 削除候補数 = 0;
  var 出力済み = {};

  // --- レベルA：名前が完全一致（コピー表記のゆれは吸収）---
  Object.keys(groups).forEach(function (key) {
    var g = groups[key];
    if (g.length < 2) return;
    g.sort(function (a, b) { return b.updated - a.updated; });   // 更新日の新しい順
    グループ番号++;
    g.forEach(function (f, i) {
      出力済み[f.id] = true;
      var 残す = (i === 0);
      if (!残す) 削除候補数++;
      行.push([
        'A-' + グループ番号,
        残す ? '★最新（残す）' : '削除候補',
        残す ? false : true,          // チェックボックス
        f.name,
        f.updated,
        f.created,
        f.size,
        f.path,
        f.url,
        f.id,
        '名前が完全一致'
      ]);
    });
  });

  // --- レベルB：末尾の日付だけが違う（自動チェックはしない）---
  Object.keys(loose).forEach(function (key) {
    var g = loose[key].filter(function (f) { return !出力済み[f.id]; });
    if (g.length < 2) return;
    g.sort(function (a, b) { return b.updated - a.updated; });
    グループ番号++;
    g.forEach(function (f, i) {
      行.push([
        'B-' + グループ番号,
        i === 0 ? '★最新' : '要確認（別日の作業かも）',
        false,                        // ← 自動ではチェックしない
        f.name,
        f.updated,
        f.created,
        f.size,
        f.path,
        f.url,
        f.id,
        '日付部分だけ違う'
      ]);
    });
  });

  var ss = レポート作成_(行, files.length, 削除候補数);
  Logger.log('レポートを作成しました： ' + ss.getUrl());
  Logger.log('走査 ' + files.length + ' 件 / 削除候補 ' + 削除候補数 + ' 件');
}


// ===================== STEP 2 =====================
function STEP2_チェックした行をゴミ箱へ() {
  var sh = レポートシート取得_();
  var 最終行 = sh.getLastRow();
  if (最終行 < 2) { Logger.log('レポートが空です。先に STEP1 を実行してください。'); return; }

  var データ = sh.getRange(2, 1, 最終行 - 1, 11).getValues();
  var 成功 = 0, 失敗 = 0;
  var 実行時刻 = Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd HH:mm');

  for (var i = 0; i < データ.length; i++) {
    if (データ[i][2] !== true) continue;              // C列チェックなし → スキップ
    var id = データ[i][9];
    if (!id) continue;
    try {
      DriveApp.getFileById(id).setTrashed(true);
      sh.getRange(i + 2, 2).setValue('ゴミ箱へ移動済 ' + 実行時刻);
      sh.getRange(i + 2, 3).setValue(false);          // 二重実行防止
      sh.getRange(i + 2, 1, 1, 11).setFontLine('line-through').setFontColor('#999999');
      成功++;
    } catch (e) {
      sh.getRange(i + 2, 2).setValue('失敗: ' + e.message);
      失敗++;
    }
  }
  SpreadsheetApp.flush();
  Logger.log('ゴミ箱へ移動： ' + 成功 + ' 件 / 失敗 ' + 失敗 + ' 件');
}


// ===================== STEP 3（保険）=====================
function STEP3_ゴミ箱から戻す() {
  var sh = レポートシート取得_();
  var 最終行 = sh.getLastRow();
  if (最終行 < 2) return;

  var データ = sh.getRange(2, 1, 最終行 - 1, 11).getValues();
  var 戻した = 0;
  for (var i = 0; i < データ.length; i++) {
    if (String(データ[i][1]).indexOf('ゴミ箱へ移動済') !== 0) continue;
    try {
      DriveApp.getFileById(データ[i][9]).setTrashed(false);
      sh.getRange(i + 2, 2).setValue('復元済み');
      sh.getRange(i + 2, 1, 1, 11).setFontLine('none').setFontColor('#000000');
      戻した++;
    } catch (e) {
      sh.getRange(i + 2, 2).setValue('復元失敗: ' + e.message);
    }
  }
  Logger.log('復元： ' + 戻した + ' 件');
}


// ===================== 内部処理 =====================

function 走査する_(開始) {
  var types = [];
  if (CONFIG.対象_スプレッドシート) types.push(MimeType.GOOGLE_SHEETS);
  if (CONFIG.対象_ドキュメント)     types.push(MimeType.GOOGLE_DOCS);
  if (CONFIG.対象_Excelファイル)    types.push(MimeType.MICROSOFT_EXCEL);

  var 自分 = Session.getEffectiveUser().getEmail();
  var 結果 = [];
  var pathCache = {};

  types.forEach(function (t) {
    var it = DriveApp.getFilesByType(t);
    while (it.hasNext()) {
      if (結果.length >= CONFIG.最大走査件数) break;
      if (new Date().getTime() - 開始 > 270000) {           // 4分半で打ち切り
        Logger.log('※時間切れのため途中まででレポートを作ります。');
        break;
      }
      var f;
      try { f = it.next(); } catch (e) { continue; }
      try {
        if (CONFIG.ゴミ箱は除外 && f.isTrashed()) continue;
        if (CONFIG.自分の所有ファイルのみ) {
          var o = f.getOwner();
          if (!o || o.getEmail() !== 自分) continue;
        }
        結果.push({
          id:      f.getId(),
          name:    f.getName(),
          updated: f.getLastUpdated(),
          created: f.getDateCreated(),
          size:    f.getSize(),
          path:    パス取得_(f, pathCache),
          url:     f.getUrl()
        });
      } catch (e) { /* 権限エラーなどは無視 */ }
    }
  });
  return 結果;
}

function パス取得_(file, cache) {
  try {
    var ps = file.getParents();
    if (!ps.hasNext()) return 'マイドライブ直下';
    var p = ps.next();
    var id = p.getId();
    if (cache[id]) return cache[id];
    var 名前 = [];
    var cur = p;
    for (var i = 0; i < 6; i++) {
      名前.unshift(cur.getName());
      var up = cur.getParents();
      if (!up.hasNext()) break;
      cur = up.next();
    }
    var path = 名前.join(' / ');
    cache[id] = path;
    return path;
  } catch (e) { return '(不明)'; }
}

/** コピー表記のゆれを吸収して名前をそろえる */
function 正規化_(name) {
  var s = String(name);
  s = s.replace(/^【バックアップ】\s*/, '');
  s = s.replace(/^コピー\s*[~～]\s*/, '');
  s = s.replace(/^「(.*)」のコピー$/, '$1');
  s = s.replace(/\s*のコピー$/, '');
  s = s.replace(/\s*[（(]\d+[）)]$/, '');
  s = s.replace(/\s*[-–]\s*\d{1,2}月\d{1,2}日[、,]\s*\d{1,2}:\d{2}\s*$/, '');  // 「- 5月25日、17:54」
  s = s.replace(/\.(xlsx|xls|csv|md)$/i, '');
  s = s.replace(/[　\s]+/g, ' ').trim().toLowerCase();
  return s;
}

/** 末尾の日付（_20260720 / _2026-07-20 / _1433版 など）を落とす */
function 日付を外す_(name) {
  var s = String(name);
  s = s.replace(/[_\-\s]?\d{3,4}版$/, '');
  s = s.replace(/[_\-\s]?(20\d{6}|20\d{2}[-_\/.]\d{1,2}[-_\/.]\d{1,2})$/, '');
  return s.trim();
}

function レポート作成_(行, 走査数, 削除候補数) {
  var フォルダ = フォルダ確保_(CONFIG.レポート保存フォルダ);
  var 名前 = 'ドライブ整理レポート_' + Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmm');
  var ss = SpreadsheetApp.create(名前);
  var file = DriveApp.getFileById(ss.getId());
  フォルダ.addFile(file);
  DriveApp.getRootFolder().removeFile(file);

  var sh = ss.getSheets()[0];
  sh.setName('重複一覧');

  var ヘッダ = ['グループ', '判定', '削除する', 'ファイル名', '更新日', '作成日',
              'サイズ(byte)', '場所', 'リンク', 'ファイルID', '一致の根拠'];
  sh.getRange(1, 1, 1, ヘッダ.length).setValues([ヘッダ])
    .setFontWeight('bold').setBackground('#e8eaed');
  sh.setFrozenRows(1);

  if (行.length > 0) {
    sh.getRange(2, 1, 行.length, ヘッダ.length).setValues(行);
    sh.getRange(2, 3, 行.length, 1).insertCheckboxes();
    sh.getRange(2, 5, 行.length, 2).setNumberFormat('yyyy/mm/dd hh:mm');

    // 「残す」行を色分け
    var 規則 = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=LEFT($B2,1)="★"')
      .setBackground('#e6f4ea')
      .setRanges([sh.getRange(2, 1, 行.length, ヘッダ.length)])
      .build();
    sh.setConditionalFormatRules([規則]);
  }

  sh.setColumnWidth(4, 320);
  sh.setColumnWidth(8, 220);
  sh.setColumnWidth(9, 90);
  sh.autoResizeColumns(1, 3);

  // 説明シート
  var info = ss.insertSheet('はじめに', 1);
  info.getRange(1, 1, 12, 1).setValues([
    ['ドライブ整理レポート'],
    [''],
    ['走査したファイル数： ' + 走査数 + ' 件'],
    ['自動でチェックが付いた削除候補： ' + 削除候補数 + ' 件'],
    [''],
    ['【手順】'],
    ['1. 「重複一覧」シートを開き、C列「削除する」のチェックを確認してください。'],
    ['   ・A-◯ のグループ … 名前が完全に同じ重複。最新1件を残し、他は自動でチェック済み。'],
    ['   ・B-◯ のグループ … 末尾の日付だけが違うもの。別日の作業の可能性があるためチェックは付けていません。'],
    ['2. 残したいものはチェックを外す／消したいものはチェックを付ける。'],
    ['3. スクリプトエディタで STEP2_チェックした行をゴミ箱へ() を実行。'],
    ['※ 削除ではなくゴミ箱への移動です。30日以内なら STEP3_ゴミ箱から戻す() で元に戻せます。']
  ]);
  info.getRange(1, 1).setFontSize(14).setFontWeight('bold');
  info.setColumnWidth(1, 700);

  PropertiesService.getScriptProperties().setProperty('最新レポートID', ss.getId());
  return ss;
}

function レポートシート取得_() {
  var id = PropertiesService.getScriptProperties().getProperty('最新レポートID');
  if (!id) throw new Error('レポートが見つかりません。先に STEP1_棚卸しレポート作成() を実行してください。');
  return SpreadsheetApp.openById(id).getSheetByName('重複一覧');
}

function フォルダ確保_(名前) {
  var it = DriveApp.getRootFolder().getFoldersByName(名前);
  return it.hasNext() ? it.next() : DriveApp.getRootFolder().createFolder(名前);
}

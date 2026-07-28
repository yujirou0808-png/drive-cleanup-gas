/**
 * 作業ファイル片付けツール（会計整理で作った使い捨てシート用）
 *
 * 「ドライブ整理」プロジェクトに ファイル追加 → スクリプト で貼り付けて使います。
 *
 *   B1_作業ファイルを一覧にする()  … マイドライブ直下を仕分けして一覧化（何も消えません）
 *   → 一覧のチェックを確認 → STEP2_チェックした行をゴミ箱へ() で削除
 *
 * 重複ではなく「使い捨てかどうか」で判定します。
 */

// ===================== 設定 =====================
var 作業CONFIG = {
  // 対象範囲： 'ルート直下' … マイドライブ直下だけ（推奨） / '全体' … フォルダの中も
  対象範囲: 'ルート直下',

  // これより新しいファイルは対象から外す（日数）。0 にすると全部対象
  何日より前を対象にするか: 14,

  // 自分が所有者のファイルだけ
  自分の所有ファイルのみ: true,

  // 対象種別
  対象_スプレッドシート: true,
  対象_ドキュメント: true,      // 引継ぎ資料.md など
  対象_テキストや画像: true     // preview_*.png、*.txt など
};

// 「使い捨ての作業ファイル」と見なす言葉（名前に含まれていたら該当）
var 作業ワード = [
  '点検', '検証', 'チェック', '確認表', '修正', '一覧', 'リスト', '手順',
  '突合', '照合', '再判定', '漏れ', 'ズレ', '比較', '仕訳案', '整合性',
  '消込', '未計上', '二重計上', '要確認', '検討', '集計_', '_案'
];

// 消さずに残しておきたい言葉（該当したら自動チェックしない）
var 保護ワード = [
  '引継ぎ', 'MF取込用', 'マスタ', 'master', '納価', '売価', '日計', '月間収支',
  '出納帳', '請求書', '送付先', 'ルール', '計算表', '台帳', '残高表'
];
// ================================================


function B1_作業ファイルを一覧にする() {
  var 開始 = new Date().getTime();
  var files = 作業ファイルを集める_(開始);
  if (files.length === 0) { Logger.log('対象ファイルなし。設定を確認してください。'); return; }

  var pathCache = {};
  var 境界 = new Date().getTime() - 作業CONFIG.何日より前を対象にするか * 86400000;

  var 使い捨て = [], 要確認 = [], 常用 = [];
  files.forEach(function (f) {
    var 日付付き = /(20\d{6}|20\d{2}[-_\/.]\d{1,2}[-_\/.]\d{1,2})/.test(f.name);
    var 作業っぽい = 作業ワード.some(function (w) { return f.name.indexOf(w) >= 0; });
    var 保護 = 保護ワード.some(function (w) { return f.name.toLowerCase().indexOf(w.toLowerCase()) >= 0; });
    var 新しい = (作業CONFIG.何日より前を対象にするか > 0 && f.updated.getTime() > 境界);

    if (保護 || 新しい)            { f.理由 = 保護 ? '残す言葉を含む' : '最近さわった'; 要確認.push(f); }
    else if (日付付き && 作業っぽい) { f.理由 = '日付付き＋作業用の名前'; 使い捨て.push(f); }
    else if (日付付き || 作業っぽい) { f.理由 = 日付付き ? '日付付き' : '作業用の名前'; 要確認.push(f); }
    else                            { f.理由 = '常用マスタの可能性'; 常用.push(f); }
  });

  var 並べ替え = function (a, b) { return b.updated - a.updated; };
  使い捨て.sort(並べ替え); 要確認.sort(並べ替え); 常用.sort(並べ替え);

  var 行 = [];
  var 積む = function (グループ, ラベル, チェック) {
    グループ.forEach(function (f) {
      行.push([
        ラベル,
        チェック ? '削除候補' : '残す（要確認）',
        チェック,
        f.name,
        f.updated,
        f.created,
        f.size,
        パス取得_(DriveApp.getFileById(f.id), pathCache),
        f.url,
        f.id,
        f.理由
      ]);
    });
  };

  積む(使い捨て, 'C-使い捨て', true);
  積む(要確認,   'C-要確認',   false);
  積む(常用,     'C-常用',     false);

  var ss = レポート作成_(行, files.length, 使い捨て.length);

  // 説明を作業ファイル用に差し替え
  var info = ss.getSheetByName('はじめに');
  info.clear();
  info.getRange(1, 1, 14, 1).setValues([
    ['作業ファイル片付け レポート'],
    [''],
    ['走査： ' + files.length + ' 件（' + 作業CONFIG.対象範囲 + '）'],
    [''],
    ['C-使い捨て … ' + 使い捨て.length + ' 件 … 自動でチェック済み。'],
    ['              日付付き（_20260720 など）＋「点検」「検証」「修正リスト」等の作業用の名前。'],
    ['              一度使って終わりのシートです。'],
    [''],
    ['C-要確認 … ' + 要確認.length + ' 件 … チェックなし。'],
    ['            引継ぎ資料・MF取込用・出納帳など、後で根拠として要るかもしれないもの。'],
    ['            直近' + 作業CONFIG.何日より前を対象にするか + '日以内にさわったものもここへ。'],
    [''],
    ['C-常用 … ' + 常用.length + ' 件 … チェックなし。日常的に使っているマスタの可能性。'],
    ['確認できたら STEP2_チェックした行をゴミ箱へ() を実行してください（30日は復元可）。']
  ]);
  info.getRange(1, 1).setFontSize(14).setFontWeight('bold');
  info.setColumnWidth(1, 760);

  Logger.log('レポート: ' + ss.getUrl());
  Logger.log('使い捨て ' + 使い捨て.length + ' 件 / 要確認 ' + 要確認.length + ' 件 / 常用 ' + 常用.length + ' 件');
}


// ===================== 内部処理 =====================
function 作業ファイルを集める_(開始) {
  var 自分 = Session.getEffectiveUser().getEmail();
  var 結果 = [], 見た = {};

  var 拾う = function (f) {
    try {
      var id = f.getId();
      if (見た[id]) return;
      if (f.isTrashed()) return;
      var mt = f.getMimeType();
      if (mt === MimeType.GOOGLE_SHEETS && !作業CONFIG.対象_スプレッドシート) return;
      if (mt === MimeType.GOOGLE_DOCS   && !作業CONFIG.対象_ドキュメント) return;
      var その他 = (mt !== MimeType.GOOGLE_SHEETS && mt !== MimeType.GOOGLE_DOCS);
      if (その他 && !作業CONFIG.対象_テキストや画像) return;
      if (mt === MimeType.GOOGLE_FORMS) return;   // フォームは触らない
      if (作業CONFIG.自分の所有ファイルのみ) {
        var o = f.getOwner();
        if (!o || o.getEmail() !== 自分) return;
      }
      見た[id] = true;
      結果.push({
        id: id, name: f.getName(), size: f.getSize(),
        updated: f.getLastUpdated(), created: f.getDateCreated(), url: f.getUrl()
      });
    } catch (e) { /* 無視 */ }
  };

  if (作業CONFIG.対象範囲 === 'ルート直下') {
    var it = DriveApp.getRootFolder().getFiles();
    while (it.hasNext()) {
      if (new Date().getTime() - 開始 > 260000) { Logger.log('※時間切れ。途中までで作ります。'); break; }
      拾う(it.next());
    }
  } else {
    var it2 = DriveApp.getFiles();
    while (it2.hasNext()) {
      if (new Date().getTime() - 開始 > 260000) { Logger.log('※時間切れ。途中までで作ります。'); break; }
      拾う(it2.next());
    }
  }
  return 結果;
}

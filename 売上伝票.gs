/**
 * 売上伝票（xlsx）まとめツール
 *
 * 「ドライブ整理」プロジェクトに ファイル追加 → スクリプト で貼り付けて使います。
 * （ドライブ整理.gs と同じプロジェクトに入れてください）
 *
 *   A1_対象を数える()          … 何件・合計何MBあるか調べるだけ（何もしない）
 *   A2_ZIPにまとめる()         … 年ごとにZIPを作ってドライブに置く → PCへ一括ダウンロード
 *   A3_重複だけ一覧にする()    … 同名重複を洗い出して一覧化（削除は STEP2 で承認後）
 *
 * 削除は必ず A3 → 一覧を確認 → STEP2_チェックした行をゴミ箱へ() の順で行います。
 */

// ===================== 設定 =====================
var 伝票CONFIG = {
  // ファイル名にこの文字が含まれるものを対象にする
  キーワード: '共同生活援助',

  // 対象の拡張子（空文字にすると拡張子を問わない）
  拡張子: '.xlsx',

  // ZIPの分け方： '年' … 2024/2025/2026 ごと、 '事業所' … グループホーム名ごと、 'まとめて' … 1本
  ZIPの分け方: '年',

  // ZIP1本あたりの上限（MB）。超えると自動で -1, -2 と分割
  ZIP上限MB: 40,

  // ZIPの保存先（マイドライブ直下）
  ZIP保存フォルダ: '_ドライブ整理'
};
// ================================================


// ===================== A1：数えるだけ =====================
function A1_対象を数える() {
  var files = 伝票を集める_();
  if (files.length === 0) { Logger.log('該当ファイルなし。キーワードを確認してください。'); return; }

  var 合計 = 0;
  files.forEach(function (f) { 合計 += f.size; });

  var 事業所 = {};
  var 年月 = {};
  files.forEach(function (f) {
    事業所[事業所名_(f.name)] = (事業所[事業所名_(f.name)] || 0) + 1;
    年月[年月_(f.name)] = (年月[年月_(f.name)] || 0) + 1;
  });

  var 同名 = {};
  files.forEach(function (f) { (同名[f.name] = 同名[f.name] || []).push(f); });
  var 余分 = 0, 重複G = 0;
  Object.keys(同名).forEach(function (k) {
    if (同名[k].length > 1) { 重複G++; 余分 += 同名[k].length - 1; }
  });

  Logger.log('== 対象ファイル ==');
  Logger.log('件数: ' + files.length + ' 件 / 合計 ' + (合計 / 1048576).toFixed(1) + ' MB');
  Logger.log('');
  Logger.log('== 事業所別 ==');
  Object.keys(事業所).sort().forEach(function (k) { Logger.log('  ' + k + ' : ' + 事業所[k] + ' 件'); });
  Logger.log('');
  Logger.log('== 年月の範囲 ==');
  var ks = Object.keys(年月).sort();
  Logger.log('  ' + ks[0] + ' 〜 ' + ks[ks.length - 1] + '（' + ks.length + 'か月分）');
  Logger.log('');
  Logger.log('== 重複 ==');
  Logger.log('  同名グループ ' + 重複G + ' 件 / 余分な本数 ' + 余分 + ' 本');
}


// ===================== A2：ZIPにまとめる =====================
function A2_ZIPにまとめる() {
  var 開始 = new Date().getTime();
  var files = 伝票を集める_();
  if (files.length === 0) { Logger.log('該当ファイルなし。'); return; }

  var 保存先 = フォルダ確保_(伝票CONFIG.ZIP保存フォルダ);
  var 束 = {};
  files.forEach(function (f) {
    var key;
    if (伝票CONFIG.ZIPの分け方 === '事業所')      key = 事業所名_(f.name);
    else if (伝票CONFIG.ZIPの分け方 === 'まとめて') key = '売上伝票';
    else                                          key = (年月_(f.name) || '不明').substring(0, 4) + '年';
    (束[key] = 束[key] || []).push(f);
  });

  var 上限 = 伝票CONFIG.ZIP上限MB * 1048576;
  var 作成数 = 0, 時間切れ = false;

  Object.keys(束).sort().forEach(function (key) {
    if (時間切れ) return;
    var グループ = 束[key];
    グループ.sort(function (a, b) { return a.name < b.name ? -1 : 1; });

    var blobs = [], 累計 = 0, 連番 = 1, 使った = {};

    var 書き出す = function () {
      if (blobs.length === 0) return;
      var 名前 = 'エクセル_' + key + (連番 > 1 ? '-' + 連番 : '') + '.zip';
      // 既に同名ZIPがあれば作り直さない（再実行時の重複防止）
      var 既存 = 保存先.getFilesByName(名前);
      if (既存.hasNext()) { Logger.log('  既にあるので飛ばします: ' + 名前); blobs = []; 累計 = 0; 連番++; return; }
      保存先.createFile(Utilities.zip(blobs, 名前));
      Logger.log('  作成: ' + 名前 + '（' + blobs.length + ' ファイル / ' + (累計 / 1048576).toFixed(1) + ' MB）');
      作成数++;
      blobs = []; 累計 = 0; 連番++;
    };

    for (var i = 0; i < グループ.length; i++) {
      if (new Date().getTime() - 開始 > 260000) { 時間切れ = true; break; }
      var f = グループ[i];
      try {
        var blob = DriveApp.getFileById(f.id).getBlob();
        // ZIP内で名前がぶつからないようにする（重複ファイルも失わない）
        var 中の名前 = f.name;
        if (使った[中の名前]) {
          var n = ++使った[中の名前];
          中の名前 = f.name.replace(/(\.[^.]+)$/, '_' + n + '$1');
        } else {
          使った[f.name] = 1;
        }
        blob.setName(中の名前);
        if (累計 + f.size > 上限 && blobs.length > 0) 書き出す();
        blobs.push(blob);
        累計 += f.size;
      } catch (e) {
        Logger.log('  読めませんでした: ' + f.name + ' → ' + e.message);
      }
    }
    書き出す();
  });

  Logger.log('');
  Logger.log('ZIPを ' + 作成数 + ' 本 作りました。保存先: マイドライブ / ' + 伝票CONFIG.ZIP保存フォルダ);
  Logger.log('ドライブでZIPを右クリック →「ダウンロード」でPCに落とせます。');
  if (時間切れ) Logger.log('※時間切れで途中までです。もう一度 A2 を実行すると続きから作ります。');
}


// ===================== A3：重複だけ一覧にする =====================
function A3_重複だけ一覧にする() {
  var files = 伝票を集める_();
  if (files.length === 0) { Logger.log('該当ファイルなし。'); return; }

  var pathCache = {};
  var 同名 = {};
  files.forEach(function (f) { (同名[f.name] = 同名[f.name] || []).push(f); });

  var 行 = [], 番号 = 0, 削除候補数 = 0;
  Object.keys(同名).sort().forEach(function (k) {
    var g = 同名[k];
    if (g.length < 2) return;
    g.sort(function (a, b) { return b.updated - a.updated; });
    番号++;
    g.forEach(function (f, i) {
      var 残す = (i === 0);
      if (!残す) 削除候補数++;
      行.push([
        'D-' + 番号,
        残す ? '★最新（残す）' : '削除候補',
        残す ? false : true,
        f.name,
        f.updated,
        f.created,
        f.size,
        パス取得_(DriveApp.getFileById(f.id), pathCache),
        f.url,
        f.id,
        (f.size === g[0].size ? '同名・同サイズ' : '同名（サイズ違い）')
      ]);
    });
  });

  var ss = レポート作成_(行, files.length, 削除候補数);
  Logger.log('レポート: ' + ss.getUrl());
  Logger.log('対象 ' + files.length + ' 件 / 削除候補 ' + 削除候補数 + ' 件');
  Logger.log('中身を確認したら STEP2_チェックした行をゴミ箱へ() を実行してください。');
}


// ===================== 内部処理 =====================
function 伝票を集める_() {
  var 結果 = [], 見た = {};
  var it = DriveApp.searchFiles('title contains "' + 伝票CONFIG.キーワード + '" and trashed = false');
  while (it.hasNext()) {
    var f;
    try { f = it.next(); } catch (e) { continue; }
    try {
      var name = f.getName();
      if (伝票CONFIG.拡張子 && name.slice(-伝票CONFIG.拡張子.length).toLowerCase() !== 伝票CONFIG.拡張子.toLowerCase()) continue;
      var id = f.getId();
      if (見た[id]) continue;
      見た[id] = true;
      結果.push({
        id: id, name: name, size: f.getSize(),
        updated: f.getLastUpdated(), created: f.getDateCreated(), url: f.getUrl()
      });
    } catch (e) { /* 無視 */ }
  }
  return 結果;
}

/** ファイル名から事業所名を取り出す（例：グループホーム　しおん＋） */
function 事業所名_(name) {
  var m = String(name).match(/(グループホーム[　\s]*[^\s　]+)/);
  return m ? m[1] : 'その他';
}

/** ファイル名の末尾から年月を取り出す（例：2025-08） */
function 年月_(name) {
  var m = String(name).match(/(20\d{2})[-_](\d{2})/);
  return m ? m[1] + '-' + m[2] : '';
}

// @ts-nocheck
function getPrevMonthDates() {
  var today = new Date();
  var year = today.getFullYear();
  var prevMonth = today.getMonth(); // 0:1月, 1:2月 … 6=7月なので-1で前月

  // 前月1日
  var start = new Date(year, prevMonth - 1, 1);
  // 前月末日
  var end = new Date(year, prevMonth, 0);

  // フォーマット
  var pad = function(n) { return ('0'+n).slice(-2); };
  var startStr = start.getFullYear() + '-' + pad(start.getMonth() + 1) + '-' + pad(start.getDate());
  var endStr = end.getFullYear() + '-' + pad(end.getMonth() + 1) + '-' + pad(end.getDate());

  return { start: startStr, end: endStr };
}

function buildQuery(dates) {
  var query = `SELECT
      id as \`id\`, title as \`タイトル\`, body as \`本文\`, url as \`記事URL\`, DATE(created_at) as \`取得日\`
    FROM \`clipping-328811.articles_dataset.articles\`
    WHERE
      TIMESTAMP_TRUNC(created_at, DAY) BETWEEN TIMESTAMP("${dates.start}") AND TIMESTAMP("${dates.end}")
      AND (title LIKE '%株式会社PR TIMES%' OR body LIKE '%株式会社PR TIMES%')
      AND (pub_flag=0 OR pub_flag=2)
      AND NOT REGEXP_CONTAINS(title, '(（プレスリリース）|【プレスリリース】)')
    ORDER BY created_at ASC
  `.trim();
  return query;
}

function getBigQueryDataAndWriteToSheet() {
  // 1. BigQueryの設定
  var projectId = 'clipping-328811'; // GCPのプロジェクトIDに置き換え

  var dates = getPrevMonthDates();
  var query = buildQuery(dates);

  var request = {
    query: query,
    useLegacySql: false
  };

  // 2. BigQueryクエリ実行と取得処理
  var queryResults = BigQuery.Jobs.query(request, projectId);
  var jobId = queryResults.jobReference.jobId;

  var sleepTimeMs = 500;
  while (!queryResults.jobComplete) {
    Utilities.sleep(sleepTimeMs);
    queryResults = BigQuery.Jobs.getQueryResults(projectId, jobId);
  }

  // 3. ヘッダーとデータ構造整形
  var rows = queryResults.rows;
  if (!rows) {
    Logger.log('No data found.');
    return;
  }

  var cols = queryResults.schema.fields.map(function(field) {
    return field.name;
  });

  var data = rows.map(function(row) {
    return row.f.map(function(cell) { return cell.v; });
  });

  // 先頭にヘッダーを追加
  data.unshift(cols);

  // 4. スプレッドシートに書き込み
  //   ※既存のスプレッドシートIDを指定する場合は下記でOK
  var ss = SpreadsheetApp.openById('1onlosAouX8AFhHtbV47xeF0Qvq4mkaS0SJ6MlMIcRuA'); // ←ご自分のIDに書き換えて！
  //   ※新規作成する場合：var ss = SpreadsheetApp.create('BigQuery Results');
  
  var sheet = ss.getSheetByName('毎月');; // 既存のシートを使う場合
  // もし常に1行目から書き込み直したい場合は、以下でシートをクリア

  // データ貼り付け
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);

  Logger.log('書き込み完了！');
}
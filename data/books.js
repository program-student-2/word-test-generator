/* レジストリ（サンプル）。scripts/convert_excel.py で単語帳を追加すると自動で上書きされます。 */
(function (global) {
  var ns = global.WTG = global.WTG || {};
  ns.registry = {
    "books": [
      {
        "id": "sample-en",
        "title": "サンプル英単語",
        "file": "sample-en.js",
        "count": 14,
        "kind": "en"
      },
      {
        "id": "sample-kobun",
        "title": "サンプル古文単語",
        "file": "sample-kobun.js",
        "count": 8,
        "kind": "kobun"
      }
    ]
  };
})(window);

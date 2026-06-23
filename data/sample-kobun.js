/* SAMPLE DATA — 動作確認用のサンプル古文単語（縦書きデモ用）。自由に差し替え/削除してください。 */
(function (global) {
  var ns = global.WTG = global.WTG || {};
  ns.books = ns.books || {};
  ns.books["sample-kobun"] = {
    "id": "sample-kobun",
    "title": "サンプル古文単語",
    "count": 8,
    "kind": "kobun",
    "words": [
      { "no": 1, "word": "あはれなり", "meaning": "しみじみと心を動かされる" },
      { "no": 2, "word": "をかし",     "meaning": "趣がある、興味深い" },
      { "no": 3, "word": "うつくし",   "meaning": "かわいらしい" },
      { "no": 4, "word": "ありがたし", "meaning": "めったにない、珍しい" },
      { "no": 5, "word": "やうやう",   "meaning": "だんだん、次第に" },
      { "no": 6, "word": "つとめて",   "meaning": "早朝、翌朝" },
      { "no": 7, "word": "いと",       "meaning": "たいそう、とても" },
      { "no": 8, "word": "おどろく",   "meaning": "はっと気づく、目を覚ます" }
    ]
  };
})(window);

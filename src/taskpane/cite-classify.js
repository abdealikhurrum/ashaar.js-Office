(function (root, factory) {
  if (typeof module !== "undefined" && module.exports) {
    module.exports = factory();
  } else {
    root.CiteClassify = factory();
  }
}(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  var BUCKET_ORDER = ["primary.fatemi", "primary.other", "secondary.fatemi", "secondary.other"];

  var HEADINGS = {
    en: {
      "primary.fatemi": "Primary Sources — Fatemi",
      "primary.other": "Primary Sources — Other",
      "secondary.fatemi": "Secondary Sources — Fatemi",
      "secondary.other": "Secondary Sources — Other"
    },
    ar: {
      "primary.fatemi": "المصادر الأساسية — الفاطمية",
      "primary.other": "المصادر الأساسية — أخرى",
      "secondary.fatemi": "المصادر الثانوية — الفاطمية",
      "secondary.other": "المصادر الثانوية — أخرى"
    }
  };

  function bucketForTags(tags) {
    var list = tags || [];
    var corpus = "other";
    var cls = "primary";
    for (var i = 0; i < list.length; i++) {
      if (list[i] === "corpus:fatemi") { corpus = "fatemi"; }
      else if (list[i] === "class:secondary") { cls = "secondary"; }
      // class:primary and everything else leave the defaults in place
    }
    return { corpus: corpus, cls: cls, key: cls + "." + corpus };
  }

  function orderedBuckets(citekeys, tagsByCitekey) {
    var map = tagsByCitekey || {};
    var groups = {};
    (citekeys || []).forEach(function (ck) {
      var tags = Object.prototype.hasOwnProperty.call(map, ck) ? map[ck] : [];
      var key = bucketForTags(tags).key;
      if (!groups[key]) { groups[key] = []; }
      groups[key].push(ck);
    });
    var out = [];
    BUCKET_ORDER.forEach(function (key) {
      if (groups[key] && groups[key].length) { out.push({ key: key, citekeys: groups[key] }); }
    });
    return out;
  }

  function headingFor(bucketKey, lang) {
    var table = (/^ar/i.test(lang || "")) ? HEADINGS.ar : HEADINGS.en;
    return table[bucketKey] || bucketKey;
  }

  function planBibliographySections(citekeys, tagsByCitekey, opts) {
    var o = opts || {};
    var keys = (citekeys || []).slice();
    if (o.sectioned) {
      var buckets = orderedBuckets(keys, tagsByCitekey);
      if (buckets.length >= 2) {
        return buckets.map(function (b) {
          return { key: b.key, heading: headingFor(b.key, o.lang), citekeys: b.citekeys };
        });
      }
      // exactly one (or zero) non-empty bucket => collapse to flat
      if (buckets.length === 1) {
        return [{ key: buckets[0].key, heading: null, citekeys: buckets[0].citekeys }];
      }
    }
    return [{ key: null, heading: null, citekeys: keys }];
  }

  return {
    BUCKET_ORDER: BUCKET_ORDER,
    bucketForTags: bucketForTags,
    orderedBuckets: orderedBuckets,
    headingFor: headingFor,
    planBibliographySections: planBibliographySections
  };
}));

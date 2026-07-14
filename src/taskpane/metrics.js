(function (root, factory) {
  if (typeof module !== "undefined" && module.exports) {
    module.exports = factory();
  } else {
    root.AshaarMetrics = factory();
  }
}(typeof globalThis !== "undefined" ? globalThis : this, function () {
  /**
   * Start a performance run with named phases.
   *
   * Semantics: phase(name) labels the segment that just ENDED — from the
   * previous boundary (startRun(), or the last phase() call) up to now. Every
   * call site names the work it just finished, so this matches the natural
   * "I just did X, mark it" usage throughout taskpane.js. end() closes the
   * run; if any time elapsed since the last phase() boundary (or since start,
   * if phase() was never called), that trailing span is reported as a final
   * "(tail)" segment so no time silently goes unattributed.
   *
   * @param {string} label - Name of the run (e.g. "apply poem")
   * @param {function} nowFn - Optional clock function; defaults to Date.now
   * @returns {object} Run object with phase(name), end(), and report() methods
   */
  function startRun(label, nowFn) {
    nowFn = nowFn || Date.now;
    var startTime = nowFn();
    var lastBoundary = startTime;
    var phases = [];   // [{name, ms}, ...]
    var endTime = null;

    return {
      /**
       * Close out the segment since the last boundary under this name.
       * @param {string} name - Name of the segment that just finished
       */
      phase: function (name) {
        var t = nowFn();
        phases.push({ name: name, ms: t - lastBoundary });
        lastBoundary = t;
      },

      /**
       * Close the run. If time elapsed since the last boundary, report it
       * as a final "(tail)" segment.
       */
      end: function () {
        endTime = nowFn();
        if (endTime > lastBoundary) {
          phases.push({ name: "(tail)", ms: endTime - lastBoundary });
          lastBoundary = endTime;
        }
      },

      /**
       * Get the final report.
       * @returns {object} {label, totalMs, phases: [{name, ms}]}
       */
      report: function () {
        var total = endTime !== null ? (endTime - startTime) : (nowFn() - startTime);
        return {
          label: label,
          totalMs: total,
          phases: phases.slice()
        };
      }
    };
  }

  return {
    startRun: startRun
  };
}));

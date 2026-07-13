(function (root, factory) {
  if (typeof module !== "undefined" && module.exports) {
    module.exports = factory();
  } else {
    root.AshaarMetrics = factory();
  }
}(typeof globalThis !== "undefined" ? globalThis : this, function () {
  /**
   * Start a performance run with named phases.
   * @param {string} label - Name of the run (e.g. "apply poem")
   * @param {function} nowFn - Optional clock function; defaults to Date.now
   * @returns {object} Run object with phase(name), end(), and report() methods
   */
  function startRun(label, nowFn) {
    nowFn = nowFn || Date.now;
    var startTime = nowFn();
    var boundaries = [startTime];  // [start, phase1_boundary, phase2_boundary, ..., end]
    var phaseNames = ["start"];    // ["start", phase1_name, phase2_name, ...]
    var endTime = null;

    return {
      /**
       * Close the previous phase and open a new one.
       * @param {string} name - Phase name
       */
      phase: function (name) {
        boundaries.push(nowFn());
        phaseNames.push(name);
      },

      /**
       * Close the last phase.
       */
      end: function () {
        endTime = nowFn();
        boundaries.push(endTime);
      },

      /**
       * Get the final report.
       * @returns {object} {label, totalMs, phases: [{name, ms}]}
       */
      report: function () {
        var phases = [];
        for (var i = 0; i < phaseNames.length; i++) {
          var ms = boundaries[i + 1] - boundaries[i];
          phases.push({ name: phaseNames[i], ms: ms });
        }
        var total = endTime !== null ? (endTime - startTime) : (nowFn() - startTime);
        return {
          label: label,
          totalMs: total,
          phases: phases
        };
      }
    };
  }

  return {
    startRun: startRun
  };
}));

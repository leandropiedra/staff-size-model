/**
 * Anaplan FY26 NA monthly seasonality (Traditional Core).
 * Source: FY26 plan screenshot — Arrivals, Departures, Elections Out (Promotions).
 * Each array is NA monthly totals (Jan–Dec); weights = month / FY26 annual total.
 */
(function (global) {
  "use strict";

  const MONTH_LABELS = [
    "Jan", "Feb", "Mar", "Apr", "May", "Jun",
    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
  ];

  /** @type {{ arrivals: number[]; departures: number[]; promotions: number[] }} */
  const SEASONALITY_RAW_FY26 = {
    // FY26 total 1,779 — peak Jul–Oct (ASC/BA hiring season)
    arrivals: [45, 42, 48, 52, 58, 65, 95, 435, 448, 400, 48, 43],
    // FY26 total 1,512 — elevated Jan, Apr, Jul
    departures: [255, 95, 110, 150, 95, 110, 250, 105, 95, 85, 82, 80],
    // FY26 Elections Out total 1,461 — Jan & Jul promotion cycles
    promotions: [528, 43, 43, 43, 43, 43, 503, 43, 43, 43, 43, 43],
  };

  function normalizeWeights(monthly) {
    const total = monthly.reduce((a, b) => a + b, 0);
    if (total <= 0) return monthly.map(() => 1 / 12);
    return monthly.map((v) => v / total);
  }

  const SEASONALITY_WEIGHTS = {
    arrivals: normalizeWeights(SEASONALITY_RAW_FY26.arrivals),
    departures: normalizeWeights(SEASONALITY_RAW_FY26.departures),
    promotions: normalizeWeights(SEASONALITY_RAW_FY26.promotions),
  };

  /** Distribute an annual total across 12 months using movement-specific weights. */
  function distributeAnnual(total, movement) {
    const annual = Math.round(total);
    if (annual === 0) return Array(12).fill(0);
    const w = SEASONALITY_WEIGHTS[movement] || SEASONALITY_WEIGHTS.arrivals;
    const raw = w.map((x) => annual * x);
    const rounded = raw.map((x) => Math.round(x));
    const drift = annual - rounded.reduce((a, b) => a + b, 0);
    rounded[11] += drift;
    return rounded;
  }

  global.SEASONALITY_FY26 = {
    MONTH_LABELS,
    SEASONALITY_RAW_FY26,
    SEASONALITY_WEIGHTS,
    distributeAnnual,
    normalizeWeights,
  };
})(typeof window !== "undefined" ? window : globalThis);

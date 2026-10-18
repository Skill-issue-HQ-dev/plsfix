// The IRR of a cash-flow vector for the template suites' formula evaluator, by
// Excel's own documented method: Newton from the guess (10 % unless one is given),
// 20 tries at most, done once a step is under 0.00001 %. Test support only; the
// invariant is that it finds no root Excel would miss, so it has no fallback search.

const DEFAULT_GUESS = 0.1;
const MAX_TRIES = 20;
const ACCURACY = 1e-7;

function npv(rate: number, flows: number[]): number {
  return flows.reduce(
    (sum, flow, period) => sum + flow / (1 + rate) ** period,
    0,
  );
}

// The derivative of the net present value with respect to the rate.
function slope(rate: number, flows: number[]): number {
  return flows.reduce(
    (sum, flow, period) => sum - (period * flow) / (1 + rate) ** (period + 1),
    0,
  );
}

/**
 * The rate that makes the flows' net present value zero, or null where Excel
 * answers #NUM!: no outflow or no inflow, a step that leaves the real numbers,
 * or no step under 0.00001 % within 20 tries.
 */
export function irr(flows: number[], guess = DEFAULT_GUESS): number | null {
  const hasOutflow = flows.some((flow) => flow < 0);
  const hasInflow = flows.some((flow) => flow > 0);
  if (!hasOutflow || !hasInflow) return null;
  let rate = guess;
  for (let attempt = 0; attempt < MAX_TRIES; attempt += 1) {
    const next = rate - npv(rate, flows) / slope(rate, flows);
    if (!Number.isFinite(next)) return null;
    if (Math.abs(next - rate) < ACCURACY) return next;
    rate = next;
  }
  return null;
}

/** Bounded secant solver shared by desktop and host-owned what-if evaluation. */
export interface GoalSeekResult {
  readonly found: boolean
  readonly reached: number
  readonly solution: number
  readonly iterations: number
}

export async function solveGoalSeekSamples(
  evaluate: (value: number) => Promise<number>,
  initial: number,
  target: number,
): Promise<GoalSeekResult> {
  if (!Number.isFinite(initial) || !Number.isFinite(target))
    throw new Error('Goal Seek requires a finite initial value and target.')
  const tolerance = 0.001
  let x0 = initial
  let f0 = await evaluate(x0)
  if (!Number.isFinite(f0)) throw new Error('goal_seek_target_not_number')
  if (Math.abs(f0 - target) <= tolerance)
    return { found: true, reached: f0, solution: x0, iterations: 1 }
  let x1 = x0 + Math.max(Math.abs(x0) * 0.01, 1)
  if (!Number.isFinite(x1)) return { found: false, reached: f0, solution: x0, iterations: 1 }
  let f1 = await evaluate(x1)
  let iterations = 2
  while (iterations < 100) {
    if (Number.isFinite(f1) && Math.abs(f1 - target) <= tolerance)
      return { found: true, reached: f1, solution: x1, iterations }
    const slope = f1 - f0
    if (slope === 0 || !Number.isFinite(slope)) break
    const next = x1 - ((f1 - target) * (x1 - x0)) / slope
    if (!Number.isFinite(next)) break
    x0 = x1
    f0 = f1
    x1 = next
    f1 = await evaluate(x1)
    iterations++
  }
  return {
    found: Number.isFinite(f1) && Math.abs(f1 - target) <= tolerance,
    reached: f1,
    solution: x1,
    iterations,
  }
}

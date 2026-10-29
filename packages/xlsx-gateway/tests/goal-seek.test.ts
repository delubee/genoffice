import { expect, it } from 'vitest'
import { solveGoalSeekSamples } from '../src/domain/goal-seek'

it('solves a linear target and returns an already satisfied initial input', async () => {
  const solved = await solveGoalSeekSamples(async (x) => x * 3 + 2, 0, 17)
  expect(solved.found).toBe(true)
  expect(solved.solution).toBeCloseTo(5)
  expect(Math.abs(solved.reached - 17)).toBeLessThanOrEqual(0.001)
  expect(await solveGoalSeekSamples(async (x) => x * 3 + 2, 5, 17)).toMatchObject({
    found: true,
    iterations: 1,
    solution: 5,
  })
})
it('converges on a nonlinear formula without publishing any workbook writes', async () => {
  const result = await solveGoalSeekSamples(async (x) => x * x, 1, 25)
  expect(result.found).toBe(true)
  expect(result.solution).toBeCloseTo(5, 3)
})
it('stops on a constant function or error-valued guess', async () => {
  expect(await solveGoalSeekSamples(async () => 4, 0, 10)).toMatchObject({
    found: false,
    iterations: 2,
  })
  expect(await solveGoalSeekSamples(async (x) => (x === 0 ? 1 : NaN), 0, 10)).toMatchObject({
    found: false,
    iterations: 2,
  })
})
it('bounds sampling and propagates cancellation from the evaluator', async () => {
  let calls = 0
  const result = await solveGoalSeekSamples(
    async (x) => {
      calls++
      return Math.atan(x)
    },
    0,
    2,
  )
  expect(result.found).toBe(false)
  expect(calls).toBeLessThanOrEqual(100)
  const cancelled = new Error('cancelled')
  await expect(
    solveGoalSeekSamples(
      async () => {
        throw cancelled
      },
      0,
      1,
    ),
  ).rejects.toBe(cancelled)
})
it('rejects nonfinite initial parameters and first formula results', async () => {
  await expect(solveGoalSeekSamples(async (x) => x, Infinity, 1)).rejects.toThrow('finite')
  await expect(solveGoalSeekSamples(async () => Infinity, 0, 1)).rejects.toThrow(
    'target_not_number',
  )
})

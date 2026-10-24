import { suite, test } from 'mocha'
import { runBenchmark } from '../../scripts/benchmark'
suite('Workbook performance samples', () => {
  const benchmark = process.env.GENOFFICE_BENCHMARK === '1' ? test : test.skip
  benchmark('records 20 opening and 20 complete UI batch samples', async function () {
    this.timeout(1_200_000)
    await runBenchmark()
  })
})

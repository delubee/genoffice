/** Positions bound by LET/LAMBDA, distinct from document defined names.
 * Token offsets refer to the original string, preserving author whitespace,
 * table tokens and quoted strings instead of rebuilding an authored formula.
 */
export function formulaVariableOffsets(formula: string): Set<number> {
  const tokens: { text: string; start: number; word: boolean }[] = []
  const identifier = /[\p{L}_\\][\p{L}\p{N}_.\\]*/uy
  for (let at = 0; at < formula.length;) {
    if (/\s/.test(formula[at]!)) {
      at++
      continue
    }
    const start = at,
      char = formula[at]!
    if (char === '"' || char === "'") {
      at++
      while (at < formula.length) {
        if (formula[at++] !== char) continue
        if (formula[at] === char) {
          at++
          continue
        }
        break
      }
      tokens.push({ text: formula.slice(start, at), start, word: false })
      continue
    }
    if (char === '[') {
      let depth = 0
      do {
        const c = formula[at++]
        if (c === '[') depth++
        else if (c === ']') depth--
      } while (at < formula.length && depth)
      tokens.push({ text: formula.slice(start, at), start, word: false })
      continue
    }
    identifier.lastIndex = at
    const word = identifier.exec(formula)?.[0]
    if (word) {
      tokens.push({ text: word, start, word: true })
      at += word.length
    } else {
      tokens.push({ text: char, start, word: false })
      at++
    }
  }
  const matching = new Map<number, number>(),
    stack: number[] = []
  tokens.forEach((token, index) => {
    if (token.text === '(') stack.push(index)
    else if (token.text === ')') {
      const open = stack.pop()
      if (open !== undefined) matching.set(open, index)
    }
  })
  const protectedOffsets = new Set<number>()
  const argsOf = (from: number, to: number) => {
    const args: [number, number][] = []
    let start = from,
      depth = 0
    for (let at = from; at < to; at++) {
      const text = tokens[at]!.text
      if (text === '(' || text === '{') depth++
      else if (text === ')' || text === '}') depth--
      else if (text === ',' && depth === 0) {
        args.push([start, at])
        start = at + 1
      }
    }
    args.push([start, to])
    return args
  }
  const visit = (from: number, to: number, bound: Set<string>) => {
    for (let at = from; at < to; at++) {
      const token = tokens[at]!
      const unqualified = tokens[at - 1]?.text !== '!' && tokens[at + 1]?.text !== '!'
      if (token.word && unqualified && bound.has(token.text.toLowerCase()))
        protectedOffsets.add(token.start)
      const open =
        token.word && tokens[at + 1]?.text === '(' ? at + 1 : token.text === '(' ? at : undefined
      const close = open === undefined ? undefined : matching.get(open)
      if (open === undefined || close === undefined || close >= to) continue
      const functionName =
        token.word && unqualified ? token.text.replace(/^_xlfn\./i, '').toUpperCase() : ''
      const args = argsOf(open + 1, close),
        env = new Set(bound)
      for (let index = 0; index < args.length; index++) {
        const [a, b] = args[index]!
        const declaration =
          (functionName === 'LET' && index < args.length - 1 && index % 2 === 0) ||
          (functionName === 'LAMBDA' && index < args.length - 1)
        if (declaration && b === a + 1 && tokens[a]?.word) {
          protectedOffsets.add(tokens[a]!.start)
          if (functionName === 'LAMBDA') env.add(tokens[a]!.text.toLowerCase())
        } else {
          visit(a, b, env)
          if (functionName === 'LET' && index % 2 === 1) {
            const [x, y] = args[index - 1]!
            if (y === x + 1 && tokens[x]?.word) env.add(tokens[x]!.text.toLowerCase())
          }
        }
      }
      at = close
    }
  }
  visit(0, tokens.length, new Set())
  return protectedOffsets
}

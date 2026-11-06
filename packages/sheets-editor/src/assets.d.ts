// Image assets reached through the shared UI package's barrel (typecheck only; bundlers resolve them).
declare module '*.png' {
  const src: string
  export default src
}

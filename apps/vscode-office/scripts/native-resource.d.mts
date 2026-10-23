export const targets: string[]
export function nativePath(root: string, target: string): string
export function validateHeader(bytes: Buffer, target: string): void
export function checkNative(root: string, target: string): Promise<string>

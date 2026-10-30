import type { DesktopApi } from '../shared/ipc'

/**
 * Privileged capabilities the shared Docs ribbon and editor helpers use. The
 * desktop shell answers through its preload bridge; an embedding host (the VS
 * Code view) installs its own typed port instead of a desktop preload object.
 */
export type DocsHostPort = Pick<
  DesktopApi,
  'pickImage' | 'openNewTab' | 'listDocsTabs' | 'focusDocsTab' | 'zoteroCommand' | 'print'
> & {
  /** absent: copies keep the default HTML clipboard payload */
  copyImageToClipboard?: DesktopApi['copyImageToClipboard']
}

const unavailable = (feature: string) => () =>
  Promise.reject(new Error(`${feature} is not available in this host.`))

/** A host without a port reports each capability instead of failing on a missing bridge. */
const NO_HOST: DocsHostPort = {
  pickImage: unavailable('Picture picking'),
  openNewTab: unavailable('New tab'),
  listDocsTabs: unavailable('Tab switching'),
  focusDocsTab: unavailable('Tab switching'),
  zoteroCommand: unavailable('Zotero'),
  print: unavailable('Printing'),
}

let installed: DocsHostPort | null = null

export function setDocsHost(port: DocsHostPort | null): void {
  installed = port
}

export function docsHost(): DocsHostPort {
  return installed ?? (window as { desktop?: DocsHostPort }).desktop ?? NO_HOST
}

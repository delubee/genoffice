/**
 * Folder token the Markdown references pictures under; the writer swaps it
 * for the real `<file>_files` folder once the save path is known. User text
 * can never spell it: `_` is escaped on the way out.
 */
export const MARKDOWN_MEDIA_DIR_TOKEN = '__genoffice_media__'

/// Days between the 1900 and 1904 date-system epochs. numfmt is 1900-only,
/// so 1904 workbooks shift date serials before formatting (display-only; the
/// model keeps the file's original serials so saving stays lossless).
export const DATE_1904_OFFSET = 1462

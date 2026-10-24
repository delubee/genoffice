/** Excel column widths use 1/256 of the Normal font's maximum digit width.
 * Keep import and export conversions inverse at the chosen pixel precision. */
export function columnWidthToPixels(width: number, maximumDigitWidth = 7) {
  return Math.floor((width + Math.floor(128 / maximumDigitWidth) / 256) * maximumDigitWidth)
}
export function pixelsToColumnWidth(pixels: number, maximumDigitWidth = 7) {
  return Math.ceil((pixels / maximumDigitWidth) * 256) / 256
}

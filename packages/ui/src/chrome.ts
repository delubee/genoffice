/**
 * Browser-only office chrome. Keep this entry independent of the AI composer,
 * providers and desktop bridges so native hosts can supply their own chat UI.
 */
export {
  ColorPicker,
  THEME_COLORS,
  THEME_COLOR_SHADES,
  STANDARD_COLORS,
  type ColorPickerProps,
  type ColorPickerStrings,
  type ColorSwatch,
} from './color-picker'
export { Dropdown, type DropdownOption } from './dropdown'
export { installScreenTips } from './screentip'
export {
  installPopoverDismiss,
  subscribeChromePressed,
  useDismissablePopover,
  type PopoverDismissOptions,
} from './popover-dismiss'
export {
  useRibbonCollapse,
  RibbonCollapseButton,
  RibbonExpandButton,
  isRibbonCompactShortcut,
  isRibbonToggleShortcut,
  readRibbonCollapsed,
  readRibbonDensity,
  RIBBON_COMPACT_SHORTCUT,
  RIBBON_TOGGLE_SHORTCUT,
  type RibbonCollapse,
  type RibbonCollapseLabels,
  type RibbonDensity,
} from './ribbon-collapse'
export { isSymbolFontFamily } from './symbol-fonts'
export {
  BUILTIN_FONT_FAMILIES,
  fontFamiliesFor,
  partitionFontFamilies,
  systemFamiliesBesidesCandidates,
} from './font-list'
export { SHAPE_GALLERY_GROUPS, ShapePreview } from './shape-gallery'
export {
  WORDART_PRESETS,
  wordArtSolidColor,
  wordArtStrokePx,
  type WordArtPreset,
} from './wordart-presets'
export { clampZoom, notchStep, type ZoomWheelIntent } from './wheel-zoom'
export { trapTab, useModalKeys } from './modal-keys'
export { OfficeFileActions } from './OfficeFileActions'
export { OFFICE_ACTION_LABELS, type OfficeActionLabels } from './strings-office-actions'

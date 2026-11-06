import { BehaviorSubject } from 'rxjs'

/// True while nothing is on the visual clipboard; the cell context menu's
/// Paste Picture/Shape entry disables on it.
export const visualPasteDisabled$ = new BehaviorSubject(true)

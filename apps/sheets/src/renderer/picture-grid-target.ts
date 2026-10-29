import { IRenderManagerService, Vector2 } from '@univerjs/engine-render'
import { SheetSkeletonManagerService } from '@univerjs/preset-sheets-core'
import type { UniverRuntime } from './univer-state'
import { pointInRect, type PictureAnchor } from './picture-paste'

/// Grid cell under a viewport point (drop target), mirroring Univer's own
/// pointer hit-test; null off the sheet canvas.
export function cellAtClientPoint(
  runtime: UniverRuntime,
  clientX: number,
  clientY: number,
): PictureAnchor | null {
  const workbook = runtime.univerAPI.getActiveWorkbook()
  if (!workbook) return null
  const render = runtime.univer
    .__getInjector()
    .get(IRenderManagerService)
    .getRenderById(workbook.getId())
  const skeleton = render?.with(SheetSkeletonManagerService).getCurrentSkeleton()
  if (!render || !skeleton) return null
  const bounds = render.engine.getCanvasElement().getBoundingClientRect()
  if (!pointInRect(clientX, clientY, bounds)) return null
  const scene = render.scene
  const relative = scene.getCoordRelativeToViewport(
    Vector2.FromArray([clientX - bounds.left, clientY - bounds.top]),
  )
  const scrollXY = scene.getScrollXYInfoByViewport(relative)
  const { scaleX, scaleY } = scene.getAncestorScale()
  const { row, column } = skeleton.getCellIndexByOffset(
    relative.x,
    relative.y,
    scaleX,
    scaleY,
    scrollXY,
  )
  return row < 0 || column < 0 ? null : { row, column }
}

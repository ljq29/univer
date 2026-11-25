export interface SelectionPosition {
  actualRow: number;
  actualColumn: number;
}

export function shouldBlockArrowKey(key: string, cell: SelectionPosition | null | undefined | void): boolean {
  if (!cell) return false;
  return (key === 'ArrowUp' && cell.actualRow === 0)
    || (key === 'ArrowLeft' && cell.actualColumn === 0);
}

export function isCellEditingTarget(target: EventTarget | null): boolean {
  if (!(target instanceof Element)) return false;
  if (target.closest('input, textarea')) return true;
  const editable = target.closest('[contenteditable]:not([contenteditable="false"])');
  if (!editable) return false;
  if (editable.getAttribute('data-u-comp') === 'editor') {
    return editable.id.includes('FORMULA_BAR') || Boolean(editable.textContent);
  }
  return true;
}

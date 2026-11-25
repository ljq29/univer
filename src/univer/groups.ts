import type { CustomData, IRange } from '@univerjs/core';

export const GROUPS_CUSTOM_KEY = 'siyuanUniverGroups';

export type GroupDimension = 'row' | 'column';

export interface SheetGroup {
  sheetId: string;
  dimension: GroupDimension;
  start: number;
  end: number;
  collapsed: boolean;
}

export function readGroups(custom: CustomData): SheetGroup[] {
  const value = custom?.[GROUPS_CUSTOM_KEY];
  if (!Array.isArray(value)) return [];

  return value.filter((group): group is SheetGroup => {
    return group
      && typeof group.sheetId === 'string'
      && (group.dimension === 'row' || group.dimension === 'column')
      && Number.isInteger(group.start)
      && Number.isInteger(group.end)
      && group.start >= 0
      && group.end >= group.start
      && typeof group.collapsed === 'boolean';
  }).map((group) => ({ ...group }));
}

export function writeGroups(custom: CustomData, groups: SheetGroup[]): CustomData {
  return {
    ...custom,
    [GROUPS_CUSTOM_KEY]: groups,
  };
}

export function addGroup(groups: SheetGroup[], group: SheetGroup): SheetGroup[] {
  const overlaps = groups.some((item) => {
    return item.sheetId === group.sheetId
      && item.dimension === group.dimension
      && item.start <= group.end
      && item.end >= group.start;
  });
  if (overlaps) {
    throw new Error('所选范围与已有分组重叠，请先取消原分组；暂不支持嵌套分组');
  }
  return [...groups, group].sort((a, b) => a.start - b.start);
}

export function findSelectedGroups(
  groups: SheetGroup[],
  sheetId: string,
  dimension: GroupDimension,
  range: IRange,
): SheetGroup[] {
  const start = dimension === 'row' ? range.startRow : range.startColumn;
  const end = dimension === 'row' ? range.endRow : range.endColumn;
  return groups.filter((group) => {
    return group.sheetId === sheetId
      && group.dimension === dimension
      && group.start <= end
      && group.end >= start;
  });
}

export function groupToRange(group: SheetGroup, rowCount: number, columnCount: number): IRange {
  if (group.dimension === 'row') {
    return {
      startRow: group.start,
      endRow: group.end,
      startColumn: 0,
      endColumn: Math.max(0, columnCount - 1),
    };
  }
  return {
    startRow: 0,
    endRow: Math.max(0, rowCount - 1),
    startColumn: group.start,
    endColumn: group.end,
  };
}

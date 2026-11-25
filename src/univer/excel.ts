import ExcelJS from 'exceljs';
import {
  BooleanNumber,
  CellValueType,
  HorizontalAlign,
  VerticalAlign,
  WrapStrategy,
  type CellValue,
  type ICellData,
  type IColorStyle,
  type IStyleData,
  type IWorkbookData,
  type IWorksheetData,
} from '@univerjs/core';
import { readGroups, writeGroups, type GroupDimension, type SheetGroup } from './groups';

const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const DAY_MS = 24 * 60 * 60 * 1000;
const EXCEL_EPOCH = Date.UTC(1899, 11, 30);

function colorToHex(color?: Partial<ExcelJS.Color>): string | undefined {
  const argb = color?.argb;
  if (!argb) return undefined;
  if (argb.length === 8) return `#${argb.slice(2)}`;
  if (argb.length === 6) return `#${argb}`;
  return undefined;
}

function colorToArgb(color?: IColorStyle | null | void): string | undefined {
  const rgb = color?.rgb;
  if (!rgb) return undefined;
  const hex = rgb.match(/^#([\da-f]{6})$/i);
  if (hex) return `FF${hex[1].toUpperCase()}`;
  const css = rgb.match(/^rgb\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)$/i);
  if (!css) return undefined;
  return `FF${css.slice(1).map((part) => Number(part).toString(16).padStart(2, '0')).join('').toUpperCase()}`;
}

function resolveStyle(cell: ICellData, snapshot: IWorkbookData): IStyleData {
  if (typeof cell.s === 'string') return snapshot.styles[cell.s] ?? {};
  return cell.s ?? {};
}

function applyStyleToExcel(cell: ExcelJS.Cell, style: IStyleData): void {
  const font: Partial<ExcelJS.Font> = {};
  if (style.ff) font.name = style.ff;
  if (style.fs) font.size = style.fs;
  if (style.it !== undefined) font.italic = Boolean(style.it);
  if (style.bl !== undefined) font.bold = Boolean(style.bl);
  if (style.ul?.s !== undefined) font.underline = Boolean(style.ul.s);
  if (style.st?.s !== undefined) font.strike = Boolean(style.st.s);
  const fontColor = colorToArgb(style.cl);
  if (fontColor) font.color = { argb: fontColor };
  if (Object.keys(font).length) cell.font = font;

  const background = colorToArgb(style.bg);
  if (background) {
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: background },
    };
  }

  const alignment: Partial<ExcelJS.Alignment> = {};
  const horizontal: Partial<Record<HorizontalAlign, ExcelJS.Alignment['horizontal']>> = {
    [HorizontalAlign.LEFT]: 'left',
    [HorizontalAlign.CENTER]: 'center',
    [HorizontalAlign.RIGHT]: 'right',
    [HorizontalAlign.JUSTIFIED]: 'justify',
    [HorizontalAlign.DISTRIBUTED]: 'distributed',
  };
  const vertical: Partial<Record<VerticalAlign, ExcelJS.Alignment['vertical']>> = {
    [VerticalAlign.TOP]: 'top',
    [VerticalAlign.MIDDLE]: 'middle',
    [VerticalAlign.BOTTOM]: 'bottom',
  };
  if (style.ht && horizontal[style.ht]) alignment.horizontal = horizontal[style.ht];
  if (style.vt && vertical[style.vt]) alignment.vertical = vertical[style.vt];
  if (style.tb !== undefined) alignment.wrapText = style.tb === WrapStrategy.WRAP;
  if (Object.keys(alignment).length) cell.alignment = alignment;
  if (style.n?.pattern) cell.numFmt = style.n.pattern;
}

function excelStyleToUniver(cell: ExcelJS.Cell): IStyleData | undefined {
  const style: IStyleData = {};
  const font = cell.font;
  if (font?.name) style.ff = font.name;
  if (font?.size) style.fs = font.size;
  if (font?.italic !== undefined) style.it = Number(font.italic);
  if (font?.bold !== undefined) style.bl = Number(font.bold);
  if (font?.underline !== undefined) style.ul = { s: Number(Boolean(font.underline)) };
  if (font?.strike !== undefined) style.st = { s: Number(font.strike) };
  const fontColor = colorToHex(font?.color);
  if (fontColor) style.cl = { rgb: fontColor };

  if (cell.fill?.type === 'pattern' && cell.fill.pattern === 'solid') {
    const background = colorToHex(cell.fill.fgColor);
    if (background) style.bg = { rgb: background };
  }

  const horizontal: Partial<Record<ExcelJS.Alignment['horizontal'], HorizontalAlign>> = {
    left: HorizontalAlign.LEFT,
    center: HorizontalAlign.CENTER,
    right: HorizontalAlign.RIGHT,
    justify: HorizontalAlign.JUSTIFIED,
    distributed: HorizontalAlign.DISTRIBUTED,
  };
  const vertical: Partial<Record<ExcelJS.Alignment['vertical'], VerticalAlign>> = {
    top: VerticalAlign.TOP,
    middle: VerticalAlign.MIDDLE,
    bottom: VerticalAlign.BOTTOM,
  };
  if (cell.alignment?.horizontal && horizontal[cell.alignment.horizontal]) {
    style.ht = horizontal[cell.alignment.horizontal];
  }
  if (cell.alignment?.vertical && vertical[cell.alignment.vertical]) {
    style.vt = vertical[cell.alignment.vertical];
  }
  if (cell.alignment?.wrapText !== undefined) {
    style.tb = cell.alignment.wrapText ? WrapStrategy.WRAP : WrapStrategy.OVERFLOW;
  }
  if (cell.numFmt && cell.numFmt !== 'General') style.n = { pattern: cell.numFmt };

  return Object.keys(style).length ? style : undefined;
}

function primitiveValue(value: ExcelJS.CellValue): { value: CellValue | null; date?: Date } {
  if (value === null || value === undefined) return { value: null };
  if (value instanceof Date) return { value: (value.getTime() - EXCEL_EPOCH) / DAY_MS, date: value };
  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
    return { value };
  }
  if ('richText' in value) return { value: value.richText.map((part) => part.text).join('') };
  if ('text' in value) return { value: value.text };
  if ('error' in value) return { value: value.error };
  if ('formula' in value || 'sharedFormula' in value) {
    return primitiveValue(value.result ?? null);
  }
  return { value: String(value) };
}

function valueType(value: CellValue | null): CellValueType | undefined {
  if (typeof value === 'string') return CellValueType.STRING;
  if (typeof value === 'number') return CellValueType.NUMBER;
  if (typeof value === 'boolean') return CellValueType.BOOLEAN;
  return undefined;
}

function columnNameToNumber(name: string): number {
  let result = 0;
  for (const char of name.toUpperCase()) result = result * 26 + char.charCodeAt(0) - 64;
  return result;
}

function parseMerge(reference: string) {
  const match = reference.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i);
  if (!match) return undefined;
  return {
    startRow: Number(match[2]) - 1,
    startColumn: columnNameToNumber(match[1]) - 1,
    endRow: Number(match[4]) - 1,
    endColumn: columnNameToNumber(match[3]) - 1,
  };
}

function outlineGroups(
  worksheet: ExcelJS.Worksheet,
  sheetId: string,
  dimension: GroupDimension,
): SheetGroup[] {
  const groups: SheetGroup[] = [];
  const count = dimension === 'row' ? worksheet.rowCount : worksheet.columnCount;
  let current: SheetGroup | undefined;
  let hiddenCount = 0;

  const finish = () => {
    if (current) {
      current.collapsed = hiddenCount === current.end - current.start + 1;
      groups.push(current);
    }
    current = undefined;
    hiddenCount = 0;
  };

  for (let index = 1; index <= count; index++) {
    const item = dimension === 'row' ? worksheet.getRow(index) : worksheet.getColumn(index);
    const level = item.outlineLevel ?? 0;
    if (level <= 0) {
      finish();
      continue;
    }
    if (!current || current.end !== index - 2) {
      finish();
      current = { sheetId, dimension, start: index - 1, end: index - 1, collapsed: false };
    } else {
      current.end = index - 1;
    }
    if (item.hidden) hiddenCount++;
  }
  finish();
  return groups;
}

export async function importXlsx(file: File, base: IWorkbookData): Promise<IWorkbookData> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(new Uint8Array(await file.arrayBuffer()) as unknown as ExcelJS.Buffer);
  const source = workbook.worksheets[0];
  if (!source) throw new Error('Excel 文件中没有工作表');

  const sheetId = base.sheetOrder[0] ?? Object.keys(base.sheets)[0] ?? 'sheet-1';
  const currentSheet = base.sheets[sheetId];
  const cellData: IWorksheetData['cellData'] = {};
  const rowData: IWorksheetData['rowData'] = {};
  const columnData: IWorksheetData['columnData'] = {};

  source.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    row.eachCell({ includeEmpty: true }, (cell, columnNumber) => {
      const formula = typeof cell.value === 'object' && cell.value !== null && 'formula' in cell.value
        ? cell.value.formula
        : undefined;
      const converted = primitiveValue(cell.value);
      const style = excelStyleToUniver(cell);
      if (converted.value === null && !formula && !style) return;
      const target: ICellData = {
        v: converted.value,
        t: valueType(converted.value),
      };
      if (formula) target.f = formula.startsWith('=') ? formula : `=${formula}`;
      if (converted.date && !target.s) target.t = CellValueType.NUMBER;
      if (style || converted.date) {
        target.s = {
          ...style,
          ...(converted.date && !style?.n ? { n: { pattern: 'yyyy-mm-dd hh:mm:ss' } } : {}),
        };
      }
      (cellData[rowNumber - 1] ??= {})[columnNumber - 1] = target;
    });

    if (row.height || row.hidden) {
      rowData[rowNumber - 1] = {
        ...(row.height ? { h: row.height * 96 / 72 } : {}),
        ...(row.hidden ? { hd: BooleanNumber.TRUE } : {}),
      };
    }
  });

  for (let index = 1; index <= source.columnCount; index++) {
    const column = source.getColumn(index);
    if (column.width || column.hidden) {
      columnData[index - 1] = {
        ...(column.width ? { w: column.width * 7 + 5 } : {}),
        ...(column.hidden ? { hd: BooleanNumber.TRUE } : {}),
      };
    }
  }

  const mergeData = (source.model.merges ?? [])
    .map(parseMerge)
    .filter((range): range is NonNullable<ReturnType<typeof parseMerge>> => Boolean(range));
  const groups = [
    ...outlineGroups(source, sheetId, 'row'),
    ...outlineGroups(source, sheetId, 'column'),
  ];
  const sheet: Partial<IWorksheetData> = {
    ...currentSheet,
    id: sheetId,
    name: source.name,
    hidden: BooleanNumber.FALSE,
    rowCount: Math.max(currentSheet?.rowCount ?? 1000, source.rowCount, 100),
    columnCount: Math.max(currentSheet?.columnCount ?? 20, source.columnCount, 20),
    cellData,
    rowData,
    columnData,
    mergeData,
    custom: undefined,
  };

  return {
    ...base,
    sheetOrder: [sheetId],
    sheets: { [sheetId]: sheet },
    styles: {},
    resources: [],
    custom: writeGroups(base.custom, groups),
  };
}

export async function exportXlsx(snapshot: IWorkbookData): Promise<Blob> {
  const sheetId = snapshot.sheetOrder[0] ?? Object.keys(snapshot.sheets)[0];
  const source = sheetId ? snapshot.sheets[sheetId] : undefined;
  if (!sheetId || !source) throw new Error('当前工作簿没有可导出的工作表');

  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'univer';
  workbook.created = new Date();
  const worksheet = workbook.addWorksheet(source.name || 'Sheet1');

  Object.entries(source.cellData ?? {}).forEach(([rowKey, row]) => {
    const cells = Object.entries(row ?? {}) as Array<[string, ICellData | null | undefined]>;
    cells.forEach(([columnKey, data]) => {
      if (!data) return;
      const cell = worksheet.getCell(Number(rowKey) + 1, Number(columnKey) + 1);
      if (data.f) {
        const result = data.v ?? undefined;
        cell.value = {
          formula: data.f.replace(/^=/, ''),
          ...(result !== undefined ? { result } : {}),
        };
      } else if (data.v !== null && data.v !== undefined) {
        cell.value = data.v;
      }
      applyStyleToExcel(cell, resolveStyle(data, snapshot));
    });
  });

  Object.entries(source.rowData ?? {}).forEach(([key, data]) => {
    if (!data) return;
    const row = worksheet.getRow(Number(key) + 1);
    if (data.h) row.height = data.h * 72 / 96;
    if (data.hd) row.hidden = true;
  });
  Object.entries(source.columnData ?? {}).forEach(([key, data]) => {
    if (!data) return;
    const column = worksheet.getColumn(Number(key) + 1);
    if (data.w) column.width = Math.max(1, (data.w - 5) / 7);
    if (data.hd) column.hidden = true;
  });
  (source.mergeData ?? []).forEach((range) => {
    worksheet.mergeCells(range.startRow + 1, range.startColumn + 1, range.endRow + 1, range.endColumn + 1);
  });

  readGroups(snapshot.custom).filter((group) => group.sheetId === sheetId).forEach((group) => {
    for (let index = group.start; index <= group.end; index++) {
      const item = group.dimension === 'row'
        ? worksheet.getRow(index + 1)
        : worksheet.getColumn(index + 1);
      item.outlineLevel = 1;
      if (group.collapsed) item.hidden = true;
    }
  });

  const buffer = await workbook.xlsx.writeBuffer();
  return new Blob([buffer as unknown as BlobPart], { type: XLSX_MIME });
}

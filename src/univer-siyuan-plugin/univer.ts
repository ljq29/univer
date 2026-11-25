// 参考 https://docs.univer.ai/zh-CN/guides/sheets/getting-started/installation#usage 配置
import '@univerjs/design/lib/index.css';
import '@univerjs/docs-ui/lib/index.css';
import '@univerjs/sheets-filter-ui/lib/index.css';
import '@univerjs/sheets-ui/lib/index.css';
import '@univerjs/ui/lib/index.css';

import { CommandType, LocaleType, merge, Univer, UniverInstanceType } from '@univerjs/core';
import { defaultTheme } from '@univerjs/themes';

import { UniverFormulaEnginePlugin } from '@univerjs/engine-formula';
import { UniverRenderEnginePlugin } from '@univerjs/engine-render';

import { UniverUIPlugin } from '@univerjs/ui';

import { UniverDocsPlugin } from '@univerjs/docs';
import { UniverDocsUIPlugin } from '@univerjs/docs-ui';

import { UniverSheetsFilterPlugin } from '@univerjs/sheets-filter';
import { UniverSheetsFilterUIPlugin } from '@univerjs/sheets-filter-ui';

import {
  SetColHiddenCommand,
  SetRowHiddenCommand,
  SetSpecificColsVisibleCommand,
  SetSpecificRowsVisibleCommand,
  UniverSheetsPlugin,
} from '@univerjs/sheets';
import { UniverSheetsUIPlugin } from '@univerjs/sheets-ui';
//#region 数学公式
import { UniverSheetsFormulaPlugin, CalculationMode } from '@univerjs/sheets-formula';
import { UniverSheetsFormulaUIPlugin } from '@univerjs/sheets-formula-ui'; // 新增
import SheetsFormulaUIZhCN from '@univerjs/sheets-formula-ui/locale/zh-CN';
//#endregion 数学公式
//#region 数字格式
import { UniverSheetsNumfmtPlugin } from '@univerjs/sheets-numfmt';
import { UniverSheetsNumfmtUIPlugin } from '@univerjs/sheets-numfmt-ui';
import SheetsNumfmtUIZhCN from '@univerjs/sheets-numfmt-ui/locale/zh-CN';
//#endregion 数字格式

//#region 排序功能
import { UniverSheetsSortPlugin } from '@univerjs/sheets-sort';
import { UniverSheetsSortUIPlugin } from '@univerjs/sheets-sort-ui';
import '@univerjs/sheets-sort-ui/lib/index.css';
import SheetsSortUIZhCN from '@univerjs/sheets-sort-ui/locale/zh-CN';
import '@univerjs/sheets-sort/facade';
//#endregion 排序功能

//#region hyperlink
// https://docs.univer.ai/zh-CN/guides/sheets/features/hyperlink#%E6%89%8B%E5%8A%A8%E7%BB%84%E5%90%88%E5%AE%89%E8%A3%85
import { UniverSheetsHyperLinkPlugin } from '@univerjs/sheets-hyper-link';
import { UniverSheetsHyperLinkUIPlugin } from '@univerjs/sheets-hyper-link-ui';
import SheetsHyperLinkUIZhCN from '@univerjs/sheets-hyper-link-ui/locale/zh-CN';
import '@univerjs/sheets-hyper-link-ui/lib/index.css';
import '@univerjs/sheets-hyper-link/facade';
import '@univerjs/sheets-hyper-link-ui/facade';
//#endregion hyperlink

//crosshair
// https://docs.univer.ai/zh-CN/guides/sheets/features/crosshair#%E6%89%8B%E5%8A%A8%E7%BB%84%E5%90%88%E5%AE%89%E8%A3%85
// import { UniverSheetsCrosshairHighlightPlugin } from '@univerjs/sheets-crosshair-highlight'
// import SheetsCrosshairHighlightZhCN from '@univerjs/sheets-crosshair-highlight/locale/zh-CN'
// import '@univerjs/sheets-crosshair-highlight/lib/index.css'
// import '@univerjs/sheets-crosshair-highlight/facade'

// find-replace
// https://docs.univer.ai/zh-CN/guides/sheets/features/find-replace#%E6%89%8B%E5%8A%A8%E7%BB%84%E5%90%88%E5%AE%89%E8%A3%85
import { UniverFindReplacePlugin } from '@univerjs/find-replace';
import { UniverSheetsFindReplacePlugin } from '@univerjs/sheets-find-replace';
import FindReplaceZhCN from '@univerjs/find-replace/locale/zh-CN';
import SheetsFindReplaceZhCN from '@univerjs/sheets-find-replace/locale/zh-CN';
import '@univerjs/find-replace/lib/index.css';
import '@univerjs/sheets-find-replace/facade';

// notes
// https://docs.univer.ai/zh-CN/guides/sheets/features/note#%E6%89%8B%E5%8A%A8%E7%BB%84%E5%90%88%E5%AE%89%E8%A3%85
import { UniverSheetsNotePlugin } from '@univerjs/sheets-note';
import { UniverSheetsNoteUIPlugin } from '@univerjs/sheets-note-ui';
import SheetsNoteUIZhCN from '@univerjs/sheets-note-ui/locale/zh-CN';
import '@univerjs/sheets-note-ui/lib/index.css';
import '@univerjs/sheets-note/facade';

// thread-comment
// https://docs.univer.ai/zh-CN/guides/sheets/features/thread-comment#%E6%89%8B%E5%8A%A8%E7%BB%84%E5%90%88%E5%AE%89%E8%A3%85
import { UniverThreadCommentPlugin } from '@univerjs/thread-comment';
import { UniverThreadCommentUIPlugin } from '@univerjs/thread-comment-ui';
import { UniverSheetsThreadCommentPlugin } from '@univerjs/sheets-thread-comment';
import { UniverSheetsThreadCommentUIPlugin } from '@univerjs/sheets-thread-comment-ui';
import ThreadCommentUIZhCN from '@univerjs/thread-comment-ui/locale/zh-CN';
import SheetsThreadCommentUIZhCN from '@univerjs/sheets-thread-comment-ui/locale/zh-CN';
import '@univerjs/thread-comment-ui/lib/index.css';
import '@univerjs/sheets-thread-comment/facade';


import DesignZhCN from '@univerjs/design/locale/zh-CN';
import DocsUIZhCN from '@univerjs/docs-ui/locale/zh-CN';
import SheetsFilterUIZhCN from '@univerjs/sheets-filter-ui/locale/zh-CN';
import SheetsFormulaZhCN from '@univerjs/sheets-formula/locale/zh-CN';

import SheetsUIZhCN from '@univerjs/sheets-ui/locale/zh-CN';
import SheetsZhCN from '@univerjs/sheets/locale/zh-CN';
import UIZhCN from '@univerjs/ui/locale/zh-CN';
import { apiProxy, sendMsg } from '~/libs/apiProxy';
import { functionZhCN } from './custom-function';
import type { checkId } from './msg';
import pkg from './plugin.json';
import { FUniver } from '@univerjs/core/facade';
import { exportXlsx, importXlsx } from './excel';
import {
  addGroup,
  findSelectedGroups,
  groupToRange,
  readGroups,
  writeGroups,
  type GroupDimension,
  type SheetGroup,
} from './groups';
import { isCellEditingTarget, shouldBlockArrowKey } from './interactions';

window.addEventListener('keydown', (event) => {
  if (event.isComposing
    || event.ctrlKey || event.metaKey || event.altKey || event.shiftKey
    || isCellEditingTarget(event.target)
    || !(event.target instanceof Element)) return;
  const app = document.getElementById('univer-siyuan-plugin--app');
  if ((!app?.contains(event.target) && event.target.getAttribute('data-u-comp') !== 'editor')) return;
  const cell = univerAPI.getActiveWorkbook()?.getActiveSheet().getSelection()?.getCurrentCell();
  if (!shouldBlockArrowKey(event.key, cell)) return;
  event.preventDefault();
  event.stopImmediatePropagation();
}, true);

const univer = new Univer({
  theme: defaultTheme,
  locale: LocaleType.ZH_CN,
  locales: {
    [LocaleType.ZH_CN]: merge(
      SheetsZhCN,
      DocsUIZhCN,
      SheetsUIZhCN,
      SheetsFormulaZhCN,
      SheetsNumfmtUIZhCN,
      DesignZhCN,
      UIZhCN,
      functionZhCN,
      SheetsFilterUIZhCN,
      SheetsSortUIZhCN,
      SheetsFormulaUIZhCN,

      SheetsHyperLinkUIZhCN,

      // SheetsCrosshairHighlightZhCN,

      FindReplaceZhCN,
      SheetsFindReplaceZhCN,

      SheetsNoteUIZhCN,

      ThreadCommentUIZhCN,
      SheetsThreadCommentUIZhCN,
    ),
  },
});

univer.registerPlugin(UniverRenderEnginePlugin);
univer.registerPlugin(UniverFormulaEnginePlugin, {
  // function: functionUser,
});

univer.registerPlugin(UniverUIPlugin, {
  container: 'univer-siyuan-plugin--app',
});

univer.registerPlugin(UniverDocsPlugin, {
  hasScroll: false,
});
univer.registerPlugin(UniverDocsUIPlugin);

univer.registerPlugin(UniverSheetsPlugin);
univer.registerPlugin(UniverSheetsUIPlugin);
univer.registerPlugin(UniverSheetsNumfmtPlugin);
univer.registerPlugin(UniverSheetsNumfmtUIPlugin);
univer.registerPlugin(UniverSheetsFormulaPlugin, {
  notExecuteFormula: false,
  description: [],
  function: [],
  initialFormulaComputing: CalculationMode.WHEN_EMPTY,
});
univer.registerPlugin(UniverSheetsFormulaUIPlugin);
univer.registerPlugin(UniverSheetsFilterPlugin);
univer.registerPlugin(UniverSheetsFilterUIPlugin);
univer.registerPlugin(UniverSheetsSortPlugin);
univer.registerPlugin(UniverSheetsSortUIPlugin);

univer.registerPlugin(UniverSheetsHyperLinkPlugin);
univer.registerPlugin(UniverSheetsHyperLinkUIPlugin);

// univer.registerPlugin(UniverSheetsCrosshairHighlightPlugin);

const univerAPI = FUniver.newAPI(univer);

// univerAPI.setCrosshairHighlightEnabled(true);

univer.registerPlugin(UniverFindReplacePlugin);
univer.registerPlugin(UniverSheetsFindReplacePlugin);

univer.registerPlugin(UniverSheetsNotePlugin);
univer.registerPlugin(UniverSheetsNoteUIPlugin);

univer.registerPlugin(UniverThreadCommentPlugin);
univer.registerPlugin(UniverThreadCommentUIPlugin);
univer.registerPlugin(UniverSheetsThreadCommentPlugin);
univer.registerPlugin(UniverSheetsThreadCommentUIPlugin);

const urlParams = new URLSearchParams(window.location.search);
const id = urlParams.get('id')!;
const copy = urlParams.get('copy');
const api = apiProxy(id);

function requireElement<T extends Element>(selector: string): T {
  const element = document.querySelector<T>(selector);
  if (!element) throw new Error(`缺少界面元素：${selector}`);
  return element;
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

async function main() {
  const dataPath = `/data/storage/petal/${pkg.name}/univer-${id}.json`;
  const checkPar: checkId = {
    type: 'llej-plugin-rpc-univer-check-id',
    blockId: id,
  };

  sendMsg(checkPar);
  const oldData = await api.getFile(dataPath);

  let data = {};
  if (oldData.sheets === undefined) {
    if (copy) data = await api.getFile(`/data/storage/petal/${pkg.name}/univer-${copy}.json`);
  } else {
    data = oldData;
  }

  univer.createUnit(UniverInstanceType.UNIVER_SHEET, data);

  let saveTimer: ReturnType<typeof setTimeout> | null = null;
  let savePromise: Promise<void> = Promise.resolve();
  let hasPendingChanges = false;
  let lastSaveError: unknown = null;

  const save = () => {
    if (!hasPendingChanges) return savePromise;
    hasPendingChanges = false;
    savePromise = savePromise.then(async () => {
      const fWorkbook = univerAPI.getActiveWorkbook();
      if (!fWorkbook) throw new Error('找不到当前工作簿');
      const snapshot = fWorkbook.save();
      await api.putFile(
        dataPath,
        false,
        new Blob([JSON.stringify(snapshot)], { type: 'text/plain' }),
      );
      lastSaveError = null;
    }).catch(async (error) => {
      hasPendingChanges = true;
      lastSaveError = error;
      console.error('[univer-siyuan-plugin] Failed to save workbook', error);
      try {
        await api.pushErrMsg(`Univer 表格保存失败：${errorMessage(error)}`);
      } catch (notifyError) {
        console.error('[univer-siyuan-plugin] Failed to show save error', notifyError);
      }
    });
    return savePromise;
  };
  const scheduleSave = () => {
    hasPendingChanges = true;
    if (saveTimer) clearTimeout(saveTimer);
    saveTimer = setTimeout(() => {
      saveTimer = null;
      void save();
    }, 500);
  };
  const flushSave = () => {
    if (saveTimer) {
      clearTimeout(saveTimer);
      saveTimer = null;
    }
    return save();
  };

  const status = requireElement<HTMLSpanElement>('#univer-siyuan-plugin--status');
  const font = requireElement<HTMLSelectElement>('#univer-siyuan-plugin--font');
  const fileInput = requireElement<HTMLInputElement>('#univer-siyuan-plugin--file');
  const groupMenu = requireElement<HTMLDetailsElement>('#univer-siyuan-plugin--group-menu');
  const app = requireElement<HTMLDivElement>('#univer-siyuan-plugin--app');
  const setBusy = (busy: boolean) => {
    app.inert = busy;
    document.querySelectorAll<HTMLButtonElement | HTMLSelectElement>(
      '#univer-siyuan-plugin--tools button, #univer-siyuan-plugin--tools select',
    ).forEach((control) => {
      control.disabled = busy;
    });
  };
  const runAction = async (pendingText: string, action: () => Promise<string>) => {
    setBusy(true);
    status.textContent = pendingText;
    try {
      status.textContent = await action();
    } catch (error) {
      status.textContent = '操作失败';
      console.error('[univer-siyuan-plugin] Action failed', error);
      try {
        await api.pushErrMsg(`Univer 操作失败：${errorMessage(error)}`);
      } catch (notifyError) {
        console.error('[univer-siyuan-plugin] Failed to show action error', notifyError);
      }
    } finally {
      setBusy(false);
    }
  };
  const activeWorkbook = () => {
    const workbook = univerAPI.getActiveWorkbook();
    if (!workbook) throw new Error('找不到当前工作簿');
    return workbook;
  };
  const persistGroups = (groups: SheetGroup[]) => {
    const workbook = activeWorkbook();
    workbook.setCustomMetadata(writeGroups(workbook.getCustomMetadata(), groups));
    scheduleSave();
  };
  const setGroupVisibility = async (groups: SheetGroup[], hidden: boolean) => {
    const workbook = activeWorkbook();
    const worksheet = workbook.getActiveSheet();
    const unitId = workbook.getId();
    const subUnitId = worksheet.getSheetId();
    for (const dimension of ['row', 'column'] as const) {
      const ranges = groups
        .filter((group) => group.sheetId === subUnitId && group.dimension === dimension)
        .map((group) => groupToRange(group, worksheet.getMaxRows(), worksheet.getMaxColumns()));
      if (!ranges.length) continue;
      const command = dimension === 'row'
        ? (hidden ? SetRowHiddenCommand : SetSpecificRowsVisibleCommand)
        : (hidden ? SetColHiddenCommand : SetSpecificColsVisibleCommand);
      const succeeded = await univerAPI.executeCommand(command.id, { unitId, subUnitId, ranges });
      if (!succeeded) throw new Error(`${dimension === 'row' ? '行' : '列'}可见性设置失败`);
    }
  };
  const runGroupAction = async (action: string): Promise<string> => {
    const workbook = activeWorkbook();
    const worksheet = workbook.getActiveSheet();
    const sheetId = worksheet.getSheetId();
    const groups = readGroups(workbook.getCustomMetadata());

    if (action === 'expand-all') {
      const sheetGroups = groups.filter((group) => group.sheetId === sheetId);
      if (!sheetGroups.length) return '当前表没有分组';
      await setGroupVisibility(sheetGroups, false);
      persistGroups(groups.map((group) => group.sheetId === sheetId
        ? { ...group, collapsed: false }
        : group));
      return '已展开当前表全部分组';
    }

    const range = workbook.getActiveRange()?.getRange();
    if (!range) throw new Error('请先选择表格范围');
    const dimension: GroupDimension = action.endsWith('row') ? 'row' : 'column';

    if (action.startsWith('create-')) {
      const start = dimension === 'row' ? range.startRow : range.startColumn;
      const end = dimension === 'row' ? range.endRow : range.endColumn;
      persistGroups(addGroup(groups, { sheetId, dimension, start, end, collapsed: false }));
      return `已创建${dimension === 'row' ? '行' : '列'}分组 ${start + 1}–${end + 1}`;
    }

    const selected = findSelectedGroups(groups, sheetId, dimension, range);
    if (!selected.length) throw new Error('当前选区没有对应分组');

    if (action.startsWith('remove-')) {
      await setGroupVisibility(selected, false);
      const removed = new Set(selected);
      persistGroups(groups.filter((group) => !removed.has(group)));
      return '已取消分组并展开内容';
    }

    const collapsed = action.startsWith('collapse-');
    await setGroupVisibility(selected, collapsed);
    const changed = new Set(selected);
    persistGroups(groups.map((group) => changed.has(group)
      ? { ...group, collapsed }
      : group));
    return collapsed ? '已折叠所选分组' : '已展开所选分组';
  };

  font.addEventListener('change', () => {
    if (!font.value) return;
    const range = activeWorkbook().getActiveRange();
    if (!range) {
      void api.pushErrMsg('请先选择要设置字体的表格范围');
      return;
    }
    range.setFontFamily(font.value);
    status.textContent = `已设置选区字体：${font.selectedOptions[0].textContent}`;
    font.value = '';
    scheduleSave();
  });
  requireElement<HTMLButtonElement>('#univer-siyuan-plugin--import').addEventListener('click', () => {
    fileInput.value = '';
    fileInput.click();
  });
  fileInput.addEventListener('change', () => {
    const file = fileInput.files?.[0];
    if (!file) return;
    if (!window.confirm('导入会替换当前表格，并且只保留 Excel 的第一个工作表。是否继续？')) return;
    void runAction('正在导入 Excel…', async () => {
      const workbook = activeWorkbook();
      const imported = await importXlsx(file, workbook.save());
      await flushSave();
      if (lastSaveError) {
        throw new Error('当前表格修改尚未保存，已取消导入，请先解决保存问题');
      }
      univerAPI.disposeUnit(workbook.getId());
      univer.createUnit(UniverInstanceType.UNIVER_SHEET, imported);
      lastSaveError = null;
      scheduleSave();
      await flushSave();
      if (lastSaveError) throw new Error('Excel 已载入，但保存失败，请保留此页面并重试');
      return 'Excel 导入并保存完成（仅第一个工作表）';
    });
  });
  requireElement<HTMLButtonElement>('#univer-siyuan-plugin--export').addEventListener('click', () => {
    void runAction('正在导出 Excel…', async () => {
      const workbook = activeWorkbook();
      const snapshot = workbook.save();
      const blob = await exportXlsx(snapshot);
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      const name = (snapshot.name || 'univer').replace(/[\\/:*?"<>|]/g, '_');
      link.href = url;
      link.download = `${name}.xlsx`;
      document.body.appendChild(link);
      link.click();
      link.remove();
      setTimeout(() => URL.revokeObjectURL(url), 1000);
      return '已导出第一个工作表';
    });
  });
  groupMenu.addEventListener('click', (event) => {
    const button = (event.target as Element).closest<HTMLButtonElement>('button[data-group]');
    if (!button) return;
    groupMenu.open = false;
    void runAction('正在更新分组…', () => runGroupAction(button.dataset.group!));
  });

  // @ts-ignore - Deprecated API but still functional
  univerAPI.onCommandExecuted((command) => {
    if (command.type !== CommandType.MUTATION) return;
    scheduleSave();
  });
  window.addEventListener('pagehide', () => {
    void flushSave();
  });
  document.addEventListener('visibilitychange', () => {
    if (document.visibilityState === 'hidden') void flushSave();
  });
}
void main().catch(async (error) => {
  console.error('[univer-siyuan-plugin] Failed to initialize', error);
  try {
    await api.pushErrMsg(`Univer 初始化失败：${errorMessage(error)}`);
  } catch (notifyError) {
    console.error('[univer-siyuan-plugin] Failed to show initialization error', notifyError);
  }
});

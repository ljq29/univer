import type { apiRpcData, apiRpcReply } from "~/libs/apiProxy";
import { SiyuanPlugin } from "~/libs/siyuanPlugin";
// 引入这个变量后 vite 会自动注入 hot
import.meta.hot;
import * as apis from "~/libs/api";
import { Menu } from "siyuan";
import { generateTimestamp, generateUniqueId } from "~/libs/js_util";
import type { checkId } from "./msg";
export default class UniverPlugin extends SiyuanPlugin {
  private resizeObserver?: MutationObserver;
  private resizeScanFrame = 0;

  onload(): void {
    this.eventBus.on("click-blockicon", (e) => {
      const menu = window.siyuan.menus.menu as Menu;
      menu.addItem({
        label: "插入 Univer表格",
        icon: ``,
        click: () => {
          const id = generateUniqueId();
          const updated = generateTimestamp();
          apis.insertBlock(
            "markdown",
            kramdownIframe({
              updated,
              id,
              src: `/plugins/univer-siyuan-plugin/univer.html?id=${id}&type=sheet`,
            }),
            undefined,
            e.detail.blockElements[0]?.dataset.nodeId,
          );
        },
      });
    });
    this.protyleSlash.push({
      id: "univer-add-sheet",
      filter: ["excel", "sheet", "表格", "univer", "univer表格"],
      callback(protyle) {
        console.log("[protyle]", protyle);
        const id = generateUniqueId();
        const updated = generateTimestamp();
        protyle.insert(
          `<div data-node-id="${id}" data-type="NodeIFrame" class="iframe" updated="${updated}">\
<div class="iframe-content">\
<iframe sandbox="allow-forms allow-presentation allow-same-origin allow-scripts allow-modals allow-popups allow-downloads" src="/plugins/univer-siyuan-plugin/univer.html?id=${id}&amp;type=sheet" data-src="" border="0" frameborder="no" framespacing="0" allowfullscreen="true" style="width: 1028px; height: 438px;"></iframe><span class="protyle-action__drag" contenteditable="false"></span></div><div class="protyle-attr" contenteditable="false">&ZeroWidthSpace;</div></div>`,
          true,
          true,
        );
        // const id = generateUniqueId();
        // const updated = generateTimestamp();
        // apis.insertBlock(
        //   "markdown",
        //   kramdownIframe({
        //     updated,
        //     id,
        //     src: `/plugins/univer-siyuan-plugin/univer.html?id=${id}&type=sheet`,
        //   }),
        //   undefined,
        //   ,
        // );
      },
      html: `Univer表格`,
    });
  }

  onLayoutReady(): void {
    // @ts-expect-error
    window["univerPlugin"] = this;
    const fn = async (ev: MessageEvent<apiRpcData | checkId>) => {
      const event = ev.data;
      // 检测复制粘贴等情况
      if (event.type === "llej-plugin-rpc-univer-check-id") {
        const iframeList = document.querySelectorAll<HTMLIFrameElement>(
          `iframe[src*="univer-siyuan-plugin/univer.html?id=${event.blockId}"]`,
        );
        for (const iframe of iframeList) {
          const block = iframe.closest<HTMLDivElement>('[data-type="NodeIFrame"]')!;
          const blockId = block.dataset.nodeId!;
          if (blockId !== event.blockId) {
            await new Promise((r) => setTimeout(r, 500));

            let newSrc = iframe.src.replace(
              `id=${event.blockId}`,
              `id=${blockId}&copy=${event.blockId}`,
            );
            if (newSrc.startsWith(location.origin)) {
              newSrc = newSrc.slice(location.origin.length);
            }

            await apis.updateBlock(
              "markdown",
              kramdownIframe({
                updated: generateTimestamp(),
                id: blockId,
                src: newSrc,
                width: iframe.style.width || "794px",
                height: iframe.style.height || "636px",
              }),
              blockId,
            );
          }
        }
      } else if (event.type === "llej-plugin-rpc") {
        const payload = event.payload;
        const iframeList = [
          ...document.querySelectorAll(
            `iframe[src*="univer-siyuan-plugin/univer.html?id=${payload.blockId}"]`,
          ),
        ] as HTMLIFrameElement[];
        let reply: apiRpcReply;
        try {
          //@ts-expect-error
          const res = await apis[payload.apiName](...payload.args);
          reply = {
            type: "llej-plugin-rpc-reply",
            msgID: event.msgID,
            payload: res,
          };
        } catch (error) {
          reply = {
            type: "llej-plugin-rpc-reply",
            msgID: event.msgID,
            payload: null,
            error: error instanceof Error ? error.message : String(error),
          };
        }
        iframeList.forEach((iframe) => {
          iframe.contentWindow?.postMessage(reply, "*");
        });
      }
    };
    window.addEventListener("message", fn);

    const scanForIframes = () => {
      this.resizeScanFrame = 0;
      document
        .querySelectorAll<HTMLIFrameElement>('iframe[src*="univer-siyuan-plugin/univer.html"]')
        .forEach(installResizeHandle);
    };
    this.resizeObserver?.disconnect();
    this.resizeObserver = new MutationObserver(() => {
      if (this.resizeScanFrame) return;
      this.resizeScanFrame = requestAnimationFrame(scanForIframes);
    });
    this.resizeObserver.observe(document.body, { childList: true, subtree: true });
    scanForIframes();
  }

  onunload(): void {
    this.resizeObserver?.disconnect();
    if (this.resizeScanFrame) cancelAnimationFrame(this.resizeScanFrame);
    document.querySelectorAll<HTMLElement>(`.${resizeHandleClass}`).forEach((handle) => {
      if (!handle.classList.contains("protyle-block-resize")) {
        handle.remove();
        return;
      }
      handle.classList.remove(resizeHandleClass);
      handle.removeAttribute("title");
      handle.removeAttribute("aria-label");
      Object.assign(handle.style, {
        bottom: "",
        height: "",
        left: "",
        opacity: "",
        right: "",
        transform: "",
        width: "",
      });
    });
  }
}

const resizeHandleClass = "univer-siyuan-plugin--resize-handle";

const installResizeHandle = (iframe: HTMLIFrameElement) => {
  const block = iframe.closest<HTMLElement>('[data-type="NodeIFrame"]');
  const container = iframe.parentElement;
  if (!block || !container) return;
  container.querySelector(`:scope > .${resizeHandleClass}:not(.protyle-block-resize)`)?.remove();
  const handle = container.querySelector<HTMLElement>(
    ':scope > .protyle-block-resize[data-resize-axis="height"]',
  );
  if (!handle || handle.classList.contains(resizeHandleClass)) return;

  handle.classList.add(resizeHandleClass);
  handle.title = "拖拽调整表格高度";
  handle.setAttribute("aria-label", "拖拽调整表格高度");
  Object.assign(handle.style, {
    bottom: "-14px",
    height: "14px",
    left: "0",
    opacity: "0.45",
    right: "0",
    transform: "none",
    width: "auto",
  });
};

const kramdownIframe = (par: {
  updated: string;
  id: string;
  src: string;
  width?: string;
  height?: string;
}) => {
  const width = par.width || "794px";
  const height = par.height || "636px";
  return `<iframe sandbox="allow-forms allow-presentation allow-same-origin allow-scripts allow-modals allow-popups allow-downloads" src="${par.src}" data-src="" border="0" frameborder="no" framespacing="0" allowfullscreen="true" style="width: ${width}; height: ${height};"></iframe>

{: updated="${par.updated}" id="${par.id}"}`;
};

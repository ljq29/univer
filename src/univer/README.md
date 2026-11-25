# Univer表格

在思源笔记中嵌入可同步的 Univer 电子表格，支持 Excel 导入导出、公式、筛选、排序、批注、行列分组和表格高度调整。

本仓库是从 [OceanPress_siyuan_plugin](https://github.com/2234839/OceanPress_siyuan_plugin) 独立出来的定制版本，保留原作者崮生的信息与赞助链接。原项目特别鸣谢 [iqeqsq](https://ld246.com/member/iqeqsq) 的赞助支持。

## 使用方法

在思源编辑器中输入 `/Univer表格`，选择对应命令即可插入表格。表格数据保存在思源文档中，随思源正常同步。

## 安装

从 [GitHub Releases](https://github.com/ljq29/univer/releases/latest) 下载 `univer.zip`，将其解压到思源工作空间的 `data/plugins/univer/`，然后重新加载思源并在“设置 → 插件”中启用“Univer表格”。每个需要使用表格的设备都要安装并启用插件。

## 从旧定制版迁移

旧版 `univer-siyuan-plugin` 创建的文档块仍引用 `/plugins/univer-siyuan-plugin/univer.html`。当前改名版本不会自动重写这些地址，迁移前请保留旧插件，不要直接停用。

## 本地构建

需要 Node.js 22 和 pnpm 10：

```bash
pnpm install --frozen-lockfile
pnpm build
```

产物位于 `dist/univer/`。

## 自动构建与发布

- 推送到 `main` 分支或提交 Pull Request 时，自动安装依赖并构建。
- 推送与 `package.json`、`plugin.json` 版本一致的 `v*` 标签时，自动构建、打包并发布 GitHub Release。

```bash
git tag v0.0.92
git push origin v0.0.92
```

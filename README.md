# Excel 对账比对器

在浏览器中比对两份 Excel，按配对键找出缺失行、多余行和已配对行中的单元格差异。项目随附本地解析库，文件的读取、比对和导出都在当前浏览器内完成。

## 快速开始

1. 保留完整目录，用 Chrome 或 Edge 打开 `index.html`。页面可直接双击打开。
2. 把两份 `.xlsx`、`.xlsm` 或 `.xls` 文件分别拖到「客户文件」和「我方文件」。也可以点击各自的区域选择文件。
3. 在「比对设置」中确认规则、工作表和表头行。检查自动推荐的配对键，必要时勾选一组能唯一标识行的共同列。
4. 用列级统计、搜索、筛选和排序查看差异，或点「导出 Excel」保存完整结果。

原入口 `tools/reconcile-viewer.html` 会打开同一页面并预选「生产物料」。

### 选择规则

| 规则 | 自动表头 | 自动配对键 |
| --- | --- | --- |
| 通用（默认） | 第一条非空行。表格有标题或说明时可手动指定表头行。 | 按共同列的非空率和唯一值比例推荐，必要时回退到第一个共同列。 |
| 生产物料 | 按生产物料关键词在前 8 行识别。 | 优先使用原工具的业务组合，例如「设计批号 + 任务代码 + 料品」，也支持「任务代码 + BOM号 + 物料代码」等预设。 |

### 比对与导出口径

- 按列名对齐，支持两侧列顺序不同。单侧列显示在统计与明细中，单元格差异只统计已配对行的共同列。
- 重复键按出现顺序配对并提示复核。所有键列都为空的行计入缺失或多余。组合键的部分列为空时，仍使用其余值配对。
- 比较时去除文本首尾空白。数字与等值的数字文本可视为一致。两侧都是文本时保留前导零，配对键也保留 `00123` 这类编码。
- 导出的工作簿包含「概览」「差异单元格明细」「缺失与多余行」。搜索、筛选、排序和分页只影响页面。每张明细工作表最多导出 200,000 条记录，截断时概览会给出提示。差异值和单侧行的值按文本保存。

使用单行表头。合并单元格读取左上角的值，公式读取文件中的缓存值。加密工作簿需要先解密。解析大型工作簿可能占用较多内存和时间，建议使用 64 位浏览器。

## 项目结构

| 路径 | 职责 |
| --- | --- |
| `index.html`、`web/` | 页面入口、样式与浏览器交互 |
| `src/core.js` | 表头和行模型、配对键推荐、行列比较 |
| `src/profiles.js` | 通用与生产物料规则 |
| `src/workbook.js` | Excel 解析、工作表选择与二维表格转换 |
| `src/report.js` | 摘要、行复制文本与三表工作簿 |
| `test/` | 合成表格、自测与 Excel 读写回归 |
| `vendor/sheetjs/` | 固定版本解析引擎及其许可 |

## 开发和复用

用 Node.js 18 或更高版本在项目根目录执行 `npm test`。测试会在内存中生成 Excel 样本，直接调用页面使用的模块。项目开发无需安装 npm 依赖。

可在项目内通过 CommonJS 入口复用核心逻辑。`buildSheetModel` 的表头行号从 1 开始。

```js
const { buildSheetModel, guessMatchKey, compareModels } = require('excel-reconcile');

const customer = buildSheetModel([['编号', '数量'], ['001', 5]], 1);
const ours = buildSheetModel([['编号', '数量'], ['001', 6]], 1);
const key = guessMatchKey(customer, ours);
const result = compareModels(customer, ours, key.names);
console.log(result.metrics.diffCells); // 1
```

`excel-reconcile/workbook` 接受 `ArrayBuffer` 或 `Uint8Array`（包括 Node.js 的 `Buffer`）。`excel-reconcile/report` 的 `buildReportWorkbook(result, meta)` 返回可交给 SheetJS 写出的工作簿。

新增业务规则时，可在 `src/profiles.js` 定义规则并加入导出对象。为规则设置唯一 `id` 和显示名称 `label`，用 `headerStrategy` 选择 `first-nonempty` 或 `keywords`。关键词策略读取 `headerKeywords` 和 `strongKeywords`。`keyPresets` 按优先级排列，每个预设的 `groups` 包含每个键字段的候选列名。核心 API 可在 `detectHeaderRow`、`listKeyPresets`、`guessMatchKey` 的第三个参数接收该规则。

本地 Excel 输入与导出可放在 `data/`。常见 Excel 文件扩展名和该目录都已加入 `.gitignore`。

## 第三方依赖

项目随附 [SheetJS CE 0.20.3 独立构建](https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js)，其 Apache 2.0 许可见 `vendor/sheetjs/LICENSE`。本地脚本的 SHA-256 为 `cc015130aa8521e7f088f88898eba949ccdcbfb38df0bd129b44b7273c3a6f41`。

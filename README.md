> **自用声明**：本仓是 [DSH-OrcaMind](https://github.com/HometownVoyager/DSH-OrcaMind) 个人工作区的自研插件独立发布仓，与 deepseek-ai 官方无关。路径按本机硬编码（Windows，工作区根 `D:\Tools\DeepSeek Harness`），不保证开箱即用；`*.local.json` 凭据文件不入库，缺失时按 README 说明重建。

# @dsh-external/dsh-office-reader

Office 文档查阅工具集：概览（`office_inspect`）、正文提取（`office_read`）与关键词搜索（`office_find`），解析 `.docx` / `.xlsx` / `.pptx`（`.xls` / `.xlsm` 经 SheetJS 支持），与 `dsh-pdf-reader` 的三件套工具形态对齐。

## 工具一览

| 工具 | 用途 |
|---|---|
| `office_inspect` | 文档概览：格式、元数据、docx 标题大纲 / xlsx 工作表列表 / pptx 幻灯片列表与单元内容量预览 |
| `office_read` | 正文提取——docx 全文 / 章节（`section`）/ 段落范围（`paragraphs`）；xlsx 全部 / 指定工作表（`sheet`）/ 单元格区域（`range`）；pptx 全文 / 幻灯片范围（`slides`）；超长自动截断 |
| `office_find` | 关键词搜索（NFKC + 大小写折叠 + 空白不敏感），命中单元为 docx 段落 / xlsx 数据行（`Sheet!第N行`）/ pptx 幻灯片，附上下文片段 |

## 输入与缓存

- 输入统一为 `source`：本地路径（绝对或相对会话工作目录）、`file://` 链接或 http(s) 网址（URL 下载仅入内存，不落盘）。
- 解析结果经进程内缓存复用：本地文件按 `size + mtime` 校验，命中不读盘不解析（上限见 `Config.cacheMaxEntries` / `cacheMaxChars`）。
- 旧版二进制格式（`.doc` / `.ppt`）与加密文档以明确中文错误报告（`.xls` / `.xlsm` 例外，经 SheetJS 支持）。

## 解析引擎

docx / pptx 为内置轻量 OOXML 遍历（fflate 解包 + 结构化 XML 提取，见 `src/ooxml.ts` / `src/docx.ts` / `src/pptx.ts`），xlsx / xls 为 SheetJS；`src/audit.ts` 提供文档审计辅助，`src/cache.ts` 为解析缓存层。

## 构建

```bash
npm install        # devDependencies 仅 typescript 与 @types/node
npm run build      # bash scripts/build.sh，产物写入 lib/
npm run typecheck  # tsc --noEmit
```

## 安装进 DSH

本插件面向 [DeepSeek Harness](https://github.com/deepseek-ai/deepseek-harness)（DSH）web profile：把本目录作为 `link:` 依赖登记到 `~/.dsh/profiles/web/package.json`，或经宿主层 `cordis.patch.yml` 以 `file:` URL 插入挂载（本工作区采用后者全局挂载，全部预设可用）。改动后须重启 DSH 生效。完整接线方式见主仓 [DSH-OrcaMind](https://github.com/HometownVoyager/DSH-OrcaMind) README 与 `Plugins/docs/插件开发指南.md`。

## 相关插件

- [`dsh-office-editor`](https://github.com/HometownVoyager/dsh-office-editor)：Office 文档编辑与制造（坐标语言与本插件对齐）
- [`dsh-pdf-reader`](https://github.com/HometownVoyager/dsh-pdf-reader) / [`dsh-pdf-editor`](https://github.com/HometownVoyager/dsh-pdf-editor)：PDF 查阅与编辑

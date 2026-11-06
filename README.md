# dsh-onenote

让 DSH 里的 agent 直接读写本机 **Microsoft OneNote 桌面版**的插件。走 OneNote 的 COM 自动化接口，操作的是真实的本地笔记本（包含 `.one` 分区文件），不经过 Graph/云端副本。

- 插件名：`dsh-onenote`
- 工具：`onenote_list` / `onenote_read` / `onenote_search` / `onenote_create` / `onenote_update` / `onenote_append` / `onenote_publish` / `onenote_delete` / `onenote_create_section` / `onenote_rename_section`
- 依赖：无 npm 依赖；运行时只需要 Windows PowerShell 5.1（`powershell.exe`）+ 本机 OneNote

## 为什么是 COM 而不是 MCP

现成的 OneNote MCP server 要么只能建页不能改页（`wgthomas/onenote-mcp`），要么走 Graph API 碰不到本地 `.one`（npm 上那几个）。本机 Office 2024 是 64 位，64 位 PowerShell 能直接实例化 `OneNote.Application`，没有 32/64 位断层，所以这里直接用 COM，不引入 Python 桥接层。

## 安装

插件是「配置型 bundle」：`package.json`（带 `dsh.bundle.patch`）+ `cordis.patch.yml` + `index.js`。

**先确认 DSH_HOME**：桌面端读的是 `$env:DSH_HOME` 指向的 home，不是 `%USERPROFILE%\.dsh`。两边都可能存在、且默认 home 里也可能有一个 desktop profile —— 装错地方的现象是「重启后工具还是不出现」。`install.ps1` 优先用 `DSH_HOME`，并在第一行打印实际使用的 home。

```powershell
# 一条命令：同步到 <DSH_HOME>\plugins-src\dsh-onenote → 复制进
# profiles\desktop\node_modules → 写 package.json 的 dependency 与
# dsh.profile.bundles → 跑 dsh plugin add 把 pnpm-lock.yaml 补齐
powershell -ExecutionPolicy Bypass -File .\install.ps1
# 需要挂本机独立分区时再加两个参数（路径由本机决定，不进仓库）：
#   -ExtraSectionName "快速笔记" -ExtraSectionPath "<你的笔记目录>\快速笔记.one"
# 找不到 dsh CLI 时它只在 PATH 里找；也可以 -DshCli <你的 dsh.cmd 绝对路径>

# 重启 DSH 桌面端，工具出现后名字是 onenote_*
```

手动安装（等价）：

```powershell
$home = $env:DSH_HOME; if (-not $home) { $home = "$env:USERPROFILE\.dsh" }
robocopy . "$home\plugins-src\dsh-onenote" /MIR /XD node_modules .git
dsh plugin --profile desktop add --registry <registry-url> "file:$($home -replace '\\','/')/plugins-src/dsh-onenote"
# 再把 "dsh-onenote" 追进 <home>\profiles\desktop\package.json 的 dsh.profile.bundles
```

`--registry` 不能省：插件包在私有 registry（`<registry-url>`，可用 `-Registry` 覆盖），不带的话 pnpm 会去 registry.npmjs.org 打 404。

如果 `node_modules` 是从别的 store 链过来的，`pnpm add` 会报 `ERR_PNPM_UNEXPECTED_STORE` / `UNEXPECTED_VIRTUAL_STORE`。`install.ps1` 第 2 步的文件复制已经能保证插件生效，第 4 步失败只会警告，不影响使用。

**桥接脚本改动不需要重启**：`scripts/onenote-bridge.ps1` 是每次工具调用时才 spawn 的，改完直接同步到 `profiles\desktop\node_modules\dsh-onenote\scripts\` 就生效；只有 `index.js`/`lib/*.js`（工具注册与参数）才需要重启。

## 配置

写在 `cordis.patch.yml` 的 `insert[0].config` 里，也可以在 profile 的 `cordis.patch.yml` 里按 `id: dsh-onenote` 覆盖：

| 配置 | 默认 | 说明 |
|---|---|---|
| `powershellPath` | `powershell.exe` | 必须是 Windows PowerShell 5.1，pwsh 7 未验证 |
| `timeoutMs` | `180000` | 单次调用超时；冷启动会拉起 OneNote，第一次可能较慢 |
| `extraSections` | 快速笔记 | **独立分区**（不属于任何笔记本的本地分区），按 `path` 指定：形如 `[{ name, path }]`。**必须用 path 而不是 id**——OneNote 每次重启都会重新分配分区 ID，路径才是稳定的；插件会在每次调用时用 `OpenHierarchy(path)` 现查 ID。`OpenHierarchy` **只认反斜杠**，正斜杠路径会解析失败；0.2.2 起插件会自动把 `/` 规范成 `\`，两种写法都能用 |
| `pageColor` | `#C7EDCC` | 新建/改写页面时写入的页面背景色（`#RRGGBB`、`automatic` 或 `none`；留空 = 跟随 OneNote 默认）。当前值 `#C7EDCC` = RGB(199,237,204) 豆沙绿，显式写死，换机器/换 OneNote 默认色也不会变 |
| `bulletValue` | `1` | OneNote 的 bullet 字形索引（`one:Bullet bullet="N"`）。实测导出 PDF：`0`=⚫大黑圆、`1`/`2`=•实心点、`3`=○空心、`7`=◊菱形、`4/5/6/8/9`=无字形。**0.5.4 起默认 1**（之前是 2，界面上可能显示成 Ø） |
| `alignCommentMax` | `80` | `alignCommentColumn="auto"` 时列宽上限；超过上限的超长行只保留 1 个空格 |
| `pageSize` | **不设置**（Automatic） | 页面尺寸（单位 pt，1pt = 1/72 in）。**默认不设置**，页面沿用 OneNote 的 `Automatic`（本机 A4 竖排）。需要更宽的正文时才配：`{ width, height, landscape, margins: { top, bottom, left, right } }`（如 A3 横排 1190×842 + 36pt 边距 → 正文宽 1118pt ≈ 150 字符/行）。OneNote 的 schema 只接受 `Automatic` 或 `Orientation + Dimensions + Margins` 三件套，宽高必须同时给；配 `"automatic"` 可把页面刷回自动尺寸 |
| `codePrefix` | `"│ "` | 围栏代码块**每个源代码行**的前缀。默认竖线，读回时据此把行识别成 `code` 块；设成 `""` 就不画竖线，代价是这些行读回时变成普通 `para` |
| `alignCommentColumn` | `0` | 代码块的**行尾注释对齐**：`0` 关闭（默认，输出与旧版逐字节一致）；`"auto"` 按块内最长代码段 + 2 对齐（受 `alignCommentMax` 限制）；`N > 0` 对齐到第 N 列。只处理「最后一个、前面是空白的 `//`」，整行注释与 `scheme://` 跳过；列宽按显示宽度算（ASCII=1，中文=2）；对齐幂等 |

> 实测参考（默认 A4 竖排、12pt 微软雅黑）：页面 21×29.7cm，左右边距 72pt，正文宽 468pt，**每行约 61 个半角字符 / 30 个汉字**；超过就折行，ASCII 图会错位。想放长代码行就调大 `pageSize`，或者减小 `margins`。
>
> 另外：OneNote 的 PDF 导出**不带页面底色**（导出永远白底），底色只在 OneNote 界面里生效。

## 工具

| 工具 | 关键参数 | 说明 |
|---|---|---|
| `onenote_list` | `notebook?`、`section?`、`includePages?`、`includeRecycleBin?`、`maxSections?`、`maxPages?` | 笔记本 → 分区 → 页面树，返回所有 ID；`maxSections`/`maxPages` 限量（0=全部），返回里带 `truncated` |
| `onenote_read` | `pageId` 或 `sectionId`/`sectionName` + `pageTitle`、`maxBlocks?` | 读页面，转成 Markdown；`maxBlocks` 限量，返回带 `blockCount`/`truncated` |
| `onenote_search` | `query`、`scopeId?`、`limit?`、`includeContent?`、`maxScan?`、`includeRecycleBin?` | 插件自己扫层级：默认按**标题**子串匹配（快，~0.5s）；`includeContent=true` 时再逐页读正文（约 50ms/页，受 `maxScan` 限制，默认 200 页）。**不用 OneNote 的 FindPages**——这台机器上它无视 query、直接返回整棵树 |
| `onenote_create` | `sectionId`/`sectionName`、`title`、`markdown` | 建页；标题也可由 markdown 的 `# 标题` 提供 |
| `onenote_update` | 页面定位 + `title?` + `markdown` | **改写已有页**：保留 pageId，删掉原有文本 Outline 后写新内容 |
| `onenote_append` | 页面定位 + `markdown` | 在页面末尾追加一个 Outline |
| `onenote_edit` | 页面定位 + `mode` + `objectId` + 模式参数 | **只改一个块**，其余内容不动。`mode` 见下节；`objectId` 从 `onenote_read` 的每个 block 里取 |
| `onenote_publish` | 页面定位 + `path` + `format?` | 导出 `.one` / `.pdf` / `.docx` / `.xps` / `.mht` / `.html` / `.onepkg` / `.emf`，目标必须是绝对路径 |
| `onenote_delete` | 页面定位 + `permanent?` | 删页，默认进 OneNote 回收站 |
| `onenote_create_section` | `path`（绝对路径 `.one`）、`keepBlankPage?` | 建/开独立分区：父目录不存在会自动创建；`OpenHierarchy(..., cft=3)` 幂等。返回 `created` 标明是新建还是打开；新建时若分区里有且仅有一页且该页无正文（OneNote 偶尔自动生成的空白页），默认删掉，`keepBlankPage=true` 可保留 |
| `onenote_rename_section` | `sectionId` 或 `path`、`name` | 改名。**笔记本内的分区可以改；独立 `.one` 分区是只读的**，会被 OneNote 拒绝并映射成 `SECTION_READ_ONLY`（`0x8004200B`）。名字里含 `\ / : * ? " < > |` 会直接 `BAD_ARGUMENT` |

页面定位一律遵循：`pageId` 优先；否则用 `sectionId`/`sectionName` + `pageTitle`（先精确匹配，再子串匹配）；只给 `pageTitle` 时全库查找。

## 局部修改（onenote_edit）

`onenote_update` 是整页重写，`onenote_append` 只能加在末尾；要改**已有内容里的一块**用 `onenote_edit`。

| mode | 参数 | 作用 |
|---|---|---|
| `replaceText` | `search`、`replacement?`、`all?` | 在该块里替换文字（默认第一处，`all=true` 全部） |
| `replaceBlock` | `text` | 把该块内容整体换成新文本（可空串清空） |
| `insertAfter` | `text`、`kind?` | 在该块**后面**插入一段（`kind`: `para`/`h2`/`h3`/`bullet`/`number`） |
| `deleteBlock` | — | 删掉该块（该 Outline 空了会整块移除） |

典型流程：

```
onenote_read(pageId)                → blocks[i].objectId
onenote_edit(pageId, mode="replaceText", objectId=..., search="旧", replacement="新")
onenote_read(pageId)                → 确认只有目标块变了
```

**实现原理**：`UpdatePageContent` 是增量的——带 `objectID` 的元素只替换该对象。所以插件只把「目标块所属的那一个 Outline」读出来、改完、原样回写，页面里其它 Outline、标题、背景、页面尺寸都不受影响。

**边界**：

- 粒度是 **Outline 内的一个 OE（段落）**，不是任意字符区间；跨段落的修改要多次调用
- 只对**文本**有效；表格单元格、图片、墨迹不在支持范围内（表格块能定位到，但不支持编辑）
- 回写会让 OneNote 重排该 Outline 的 QuickStyle 索引，但内容与已识别的 `h2`/`h3` 字号不受影响（实测 14 项检查全过）

## Markdown 子集与符号约定

`one:Style` / `QuickStyles` 这类**样式表**写法会被拒绝，正文也写不进换行（`<br/>` 不是合法子元素）；但**逐段落的 `style="font-size:…"` 加 `spaceBefore/spaceAfter` 是可以写的**（`spaceBefore/spaceAfter` 单位是磅，落在 `one:OE` 上）。所以标题写成真格式段落，其余层级仍用符号：

| Markdown | 写入 OneNote 的形态 | 读回 |
|---|---|---|
| `# 标题` | 页面标题（真正的 OneNote 标题） | `# 标题` |
| `## 标题` | **正文+4pt** 段落，段前 8 / 段后 5 | `## 标题` |
| `### 标题` | **正文+2pt** 段落，段前 5 / 段后 5 | `### 标题` |
| `> 引用` | `▎ 引用` 段落 | `> 引用` |
| ```` ``` ```` 代码块 | 每行一个 `│ 代码行` 段落 | 代码块（语言标记会丢） |
| `- 项目` / 嵌套缩进 | 真正的项目符号列表，缩进即层级 | `- 项目`，缩进还原 |
| `1. 条目` | 项目符号 + 文本自带 `1. ` 前缀 | `1. 条目` |
| `\| 表格 \|` | 真正的 OneNote 表格（带边框、首行表头） | Markdown 表格 |

**字号不是写死的**，而是相对正文大小算出来的：`###` = 正文 + 2，`##` = 正文 + 4、段前/段后间距固定 8/5 与 5/5。正文大小按这个顺序解析（`Resolve-BaseFontSize`）：

1. 目标页面自己的正文样式（`QuickStyleDef name="p"`，其次取段落里出现最多的字号）
2. 同分区已有页面取样（最多 5 页）——这最接近 OneNote 实际渲染用的默认值
3. 注册表 `HKCU\Software\Microsoft\Office\<版本>\OneNote\Options\Editing\DefaultFontSize`
4. 都没有才回落 12pt

> 之所以不直接用注册表：OneNote **启动时缓存**默认字体，改了注册表不重启它不认，写出来的标题会和正文同号、被 OneNote 当成普通段落合并掉。

读回时按「相对正文的字号差」识别：≥ 正文+3 认 `##`，≥ 正文+1 认 `###`；旧版本写的 `━━━ 标题 ━━━`、`▌ 标题` 仍然能读成 `##`/`###`。`onenote_create` 的返回里带 `baseFontSize`，方便确认这次用的是哪个基准。

不支持：字体、颜色、粗斜体、行内链接、图片/墨迹写入、页面级样式（`pageSize` 除外，见配置表）。

## 已知限制

1. **不要把 `.one` 放在网络路径上**（UNC `\\host\share`，含 WSL 的映射盘符如 `Z:`）：OneNote 能"打开"但读到 **0 页**，写入只落在它自己的缓存里，关掉就丢。文件必须放本地盘（`C:`/`D:`）。这是实测结论——同一个文件，网络路径 0 页、本地副本 11 页；一次 OneNote 重启后网络分区直接不可用。
2. **编号列表用的是「圆点 + 序号」**：`one:Number` 在这版 OneNote 上不论 `numberFormat` 取什么值都报 `0x80042001`；`bullet="0"`（无符号）又被 `0x80042000` 拒绝。所以编号项是带序号的圆点项，读回时按 `1. ` 前缀还原成有序列表。
3. **`onenote_update` 会删掉页面里所有文本 Outline**（要局部改用 `onenote_edit`）（包括嵌在 Outline 里的图片/墨迹），页面级的独立图片会保留。
4. **代码块的缩进保留在文本里，但语言标记（```js 的 js）无法保存**。
5. **首次调用慢**：OneNote 没运行时，COM 会拉起它并做一次同步，`GetHierarchy` 可能要 1–3 分钟，超过 `timeoutMs` 会失败，重试即可。
6. **COM 与 DSH 沙箱**：OneNote 的 COM 服务器在 DSH 的文件沙箱里会以 `0x80080005 (CO_E_SERVER_EXEC_FAILURE)` 失败。插件运行在宿主进程里不受沙箱限制；但**在 DSH 的 pwsh 工具里手动调 COM 需要非沙箱（完全权限）会话**。
7. **PowerShell 5.1 的两个坑**（已在本插件中修掉，二次开发注意）：
   - `.ps1` 里带中文必须存成 UTF-8 **带 BOM**；本插件的桥接脚本保持纯 ASCII，中文只走 JSON。
   - `param([string]$X)` 的变量有 `[string]` 类型约束，把解析出来的对象赋回同名变量会被转成字符串；另外 `XmlDocument` / `XmlNamespaceManager` 都是 `IEnumerable`，函数 `return` 时要写 `return ,$obj` 阻止被展开。
8. **工具参数不是任意 JSON Schema**：DSH 的 `defineTool.parameters` 用的是自己的 value schema DSL（只认 `type`/`description`/`enum` 和嵌套对象数组），写 `minimum`、`maximum`、`pattern` 之类会在**激活时**抛 `JsonSchemaError` —— 表现是重启后 `onenote_*` 一个都不出现，日志里也看不到。改参数后必须跑一遍 `node dev/host-contract.mjs`。
9. **`FindPages` 不可用**（`0x80042019` 或无视 query 返回整棵树）：它依赖 Windows Search 索引，这台机器上没装/没启。所以搜索是插件自己扫层级；正文搜索要开 `includeContent`，代价是逐页 `GetPageContent`。
10. **`onenote_publish` 的目标路径要用反斜杠**：OneNote 的 `Publish` 对正斜杠路径直接返回 `0x80042006`（看起来像「无效层级」，实际是路径分隔符问题）。桥接里已统一把 `/` 换成 `\`。
11. **长的 ASCII 图会被页宽折行**：每行代码是一个段落，默认 A4 竖排约 61 个半角字符/行就折行，`├──` 这类树形对齐会被打断。要避免只能调大 `pageSize`（见配置表），或把宽图拆成多段。
12. **`one:T` 里的文本是 HTML 形态**：OneNote 会把 `< > & "` 存成 `&lt; &gt; &amp; &quot;`，并在需要换字体时插入 `<span style=...>`（例如 `►` 落到 Arial）。**0.1.1 起读取时会自动剥离 span 并解码实体**；在此之前读回会是 `&gt;`、`<span ...>`，看起来像内容损坏——但 `onenote_read` 之外的工具（直接看 XML）仍会看到实体形态。
13. **建分区的两个坑**（`OpenHierarchy(path, '', [ref]$id, 3)` 的 `cftSection=3`）：同一次实测里我们**没有**观察到自动生成的空白页（新建分区 0 页），但对方环境报告过，所以 `onenote_create_section` 会做防御性清理——仅当新分区恰好 1 页且该页无正文时删除，`keepBlankPage=true` 可关闭；另外用 `UpdateHierarchy` 给**独立分区**改名会被拒 `0x8004200B SECTION_READ_ONLY`，只有笔记本内的分区能改。
14. **COM 错误映射曾经无效（0.3.0 修）**：PowerShell 把 COM 失败包在 `MethodInvocationException` 里，外层 `HResult` 是 `0x80131501`，真正的 OneNote 码在 `InnerException` 上；而且 `[uint32]` 转**负数** HResult 会抛异常，被 `catch` 吞掉后 `$hex` 恒为 `00000000`。结果 0.3.0 之前**所有** OneNote 错误都落到 `ONENOTE_ERROR`。现在沿 `InnerException` 链查找并先把 HRESULT 转成 Int64，`SECTION_READ_ONLY` / `PAGE_NOT_FOUND` 这些才开始真正返回。

## 开发与测试

```powershell
# 纯函数单测（markdown ⇄ block 模型 + 工具参数校验），不需要 OneNote
cd dsh-onenote; node --test

# 宿主契约检查：用 app.asar 里真实的 @deepseek-ai/dsh-tools.defineTool
# 跑一遍注册流程，改过 TOOL_PARAMETERS 后必须跑（不需要重启 DSH）
node dev/host-contract.mjs

# 端到端冒烟（会真的建页/改页/导出/删页，跑完自动清理）
node dev/live-smoke.mjs
```

桥接脚本 `scripts/onenote-bridge.ps1` 只做三件事：收 JSON 请求 → 调 COM → 写 JSON 响应。它不依赖插件代码，可以单独手工调用排查：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File scripts\onenote-bridge.ps1 -Request req.json -Response res.json
```

请求形如 `{"action":"read","params":{"pageId":"..."}}`，响应是 `{"ok":true,"data":{...}}` 或 `{"ok":false,"code":"...","message":"..."}`。

## 安装脚本

`install.ps1` 做五件事：同步源码到 `%USERPROFILE%\.dsh\plugins-src\dsh-onenote` → 复制进 `profiles\desktop\node_modules` → 在 profile 的 `package.json` 里登记 dependency 与 `dsh.profile.bundles` 条目（改动前备份成 `package.json.bak-dsh-onenote`）→ 跑 `dsh plugin add --registry ...` 让 `pnpm-lock.yaml` 与 `package.json` 一致 → 提示重启。

改完插件重新执行一次即可（幂等）；`-ProfileName` 默认 `desktop`，`-DshCli` 不传时会尝试从 PATH 找 `dsh`。

# 功耗续航工作台

支持 Supabase 邮箱账号登录、GitHub OAuth 管理员登录、云端模板和账号隔离历史记录的 PWA 功耗分析工具。未登录用户可使用游客模式和内置模板，但不会读取团队模板或云端历史。

## 功能模块

- **器件功耗数据库**：按一级 / 二级分类维护器件功耗数据，管理员可维护分类目录并上传 `.xlsx` / `.csv` 预览；分类结构与数据保存在浏览器 `localStorage`，不依赖 Supabase。
- **功耗拆解分析**：统一入口，进入后以二级列表展示三个子功能——
  - 项目模板管理：模板库管理，支持在线编辑「采样通道梳理」并生成拆解模板；
    - 采样通道梳理：本地 Excel / CSV 文档导入与飞书云文档（电子表格）链接导入（均自动识别表头并映射 15 个通道字段）、在线填写、通道列表预览与优先级统计（必测 / 合并 / 可选 / 不测）；
    - 拆解模板自动生成：基于梳理结果，用 SheetJS 在前端生成 `项目名_功耗拆解数据处理模板_日期.xlsx`（8 个 Sheet，结构对齐 `DoraG_功耗拆解数据处理模板V1.0.xlsx`），无需后端；
  - 拆解数据对比分析：导入 CSV / Excel，生成结果并完成多项目对比；
  - 拆解数据导出：选择模板、导入数据并按需导出 Excel 或云文档。
- **续航/功耗预估**、**续航/功耗现状**：开发中占位模块。

## Supabase 初始化

1. 在目标 Supabase 项目的 SQL Editor 执行 `supabase-schema.sql`。
2. 在 Authentication → URL Configuration 中加入生产地址（例如 `https://power-analysis-alpha.vercel.app`）。
3. 在 Authentication → Sign In / Providers → GitHub 中填入 GitHub OAuth App 的 Client ID / Client Secret，并启用 GitHub 登录。
4. 在 `config.js` 的 `adminEmails` / `adminGithubLogins` 中维护管理员邮箱和 GitHub 用户名白名单，并同步修改 SQL 中的管理员策略。

## 云文档导入（飞书电子表格）

「云文档导入」通过 Supabase Edge Function `feishu-sheet-proxy` 代理取数：浏览器只传飞书表格链接，
函数在服务端用飞书应用凭证换取 `tenant_access_token` 后读取表格，返回二维数组给前端复用本地导入的解析逻辑。
前端不会也不应该保存飞书 App Secret。

### 1. 在飞书开放平台准备应用

1. 创建自建应用，获取 **App ID** / **App Secret**。
2. 开通权限：`sheets:spreadsheet:readonly`（读取电子表格）、`sheets:spreadsheet`（读取工作表列表，如接口要求）。
3. 发布应用版本，并把该应用（机器人）添加为目标文档的**协作者（可阅读）**——云文档只能读取应用被授权的文档。

### 2. 配置并部署 Edge Function

```bash
supabase login
supabase link --project-ref osopccfscvxcukspwkzv

# 注入飞书应用凭证（不要写进仓库或前端代码）
supabase secrets set FEISHU_APP_ID=<app_id> FEISHU_APP_SECRET=<app_secret> \
  --project-ref osopccfscvxcukspwkzv

supabase functions deploy feishu-sheet-proxy --project-ref osopccfscvxcukspwkzv
```

可选环境变量：

| 变量 | 说明 |
| --- | --- |
| `LARK_API_BASE` | 默认 `https://open.feishu.cn/open-apis`；Lark 国际版改为 `https://open.larksuite.com/open-apis` |
| `FEISHU_PROXY_TOKEN` | 设置后前端须在 `config.js` 的 `feishuSheetProxyToken` 填入同一口令，请求头 `x-proxy-token` 校验 |

`supabase/config.toml` 中该函数已声明 `verify_jwt = false`（游客模式也能导入），因此它只提供**只读**接口，
且读取范围受限于飞书应用的文档授权。

### 3. 前端配置

`config.js`：

```js
feishuSheetProxyEndpoint: 'https://<project-ref>.supabase.co/functions/v1/feishu-sheet-proxy',
feishuSheetProxyToken: '', // 仅在 Edge Function 配置了 FEISHU_PROXY_TOKEN 时填写
```

### 4. 使用方式与降级方案

粘贴形如 `https://xxx.larkoffice.com/sheets/XXXXXXXX?sheet=xxxxxx` 的电子表格链接后点「导入」：
自动选中链接里 `sheet=` 指定的工作表（未指定时取第一个可见工作表），也可在「选择工作表」中切换后重读。
已有数据时会询问「覆盖 / 追加」。

若代理未部署或未授权（状态栏会给出具体原因），可在飞书中「··· → 下载 → Excel」导出后用「本地文档导入」完成导入。

健康检查：`GET <feishuSheetProxyEndpoint>` 返回 `{ ok, configured }`，`configured=false` 表示凭证未配置。

## 本地预览

```bash
python3 -m http.server 4173
```

浏览器访问 `http://127.0.0.1:4173`。

## 部署

项目为纯静态站点，可直接导入 Vercel；根目录保持为仓库根目录。

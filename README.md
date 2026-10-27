# 功耗续航工作台

支持 Supabase 邮箱账号登录、GitHub OAuth 管理员登录、云端模板和账号隔离历史记录的 PWA 功耗分析工具。未登录用户可使用游客模式和内置模板，但不会读取团队模板或云端历史。

## 功能模块

- **项目模板管理**：模板库管理，支持在线编辑「采样通道梳理」并生成拆解模板。
- **拆解数据对比分析**：导入 CSV / Excel，生成结果并完成多项目对比。
- **拆解数据导出**：选择模板、导入数据并按需导出 Excel 或云文档。
- **器件功耗数据库**：按一级 / 二级分类维护器件功耗数据，管理员可维护分类目录并上传 `.xlsx` / `.csv` 预览；分类结构与数据保存在浏览器 `localStorage`，不依赖 Supabase。
- **续航/功耗预估**、**续航/功耗现状**：开发中占位模块。

## Supabase 初始化

1. 在目标 Supabase 项目的 SQL Editor 执行 `supabase-schema.sql`。
2. 在 Authentication → URL Configuration 中加入生产地址（例如 `https://power-analysis-alpha.vercel.app`）。
3. 在 Authentication → Sign In / Providers → GitHub 中填入 GitHub OAuth App 的 Client ID / Client Secret，并启用 GitHub 登录。
4. 在 `config.js` 的 `adminEmails` / `adminGithubLogins` 中维护管理员邮箱和 GitHub 用户名白名单，并同步修改 SQL 中的管理员策略。

## 本地预览

```bash
python3 -m http.server 4173
```

浏览器访问 `http://127.0.0.1:4173`。

## 部署

项目为纯静态站点，可直接导入 Vercel；根目录保持为仓库根目录。

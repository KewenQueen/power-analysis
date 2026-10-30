/* global window */

window.POWER_ANALYSIS_CONFIG = {
  supabaseUrl: 'https://osopccfscvxcukspwkzv.supabase.co',
  supabasePublishableKey: 'sb_publishable_33RiwTuLTuN9bjUTgZ5m0g_GLCfVwW8',
  cloudExportEndpoint: 'https://osopccfscvxcukspwkzv.supabase.co/functions/v1/lark-sheet-export',
  // 云文档导入代理（Supabase Edge Function），由服务端用飞书应用凭证取数。
  feishuSheetProxyEndpoint: 'https://osopccfscvxcukspwkzv.supabase.co/functions/v1/feishu-sheet-proxy',
  // 若 Edge Function 配置了 FEISHU_PROXY_TOKEN，可在此填入同一个值。
  feishuSheetProxyToken: '',
  // 管理员邮箱与 GitHub 用户名白名单。GitHub 登录必须经 OAuth 验证。
  adminEmails: ['hukehuan@bytedance.com'],
  adminGithubLogins: ['KewenQueen'],
};

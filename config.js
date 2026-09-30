/* global window */

window.POWER_ANALYSIS_CONFIG = {
  supabaseUrl: 'https://osopccfscvxcukspwkzv.supabase.co',
  supabasePublishableKey: 'sb_publishable_33RiwTuLTuN9bjUTgZ5m0g_GLCfVwW8',
  cloudExportEndpoint: 'https://osopccfscvxcukspwkzv.supabase.co/functions/v1/lark-sheet-export',
  // 飞书 OAuth 导出云文档（纯前端 PKCE 流程，不涉及 App Secret）
  feishuAppId: 'cli_aa3672d381789bd1',
  feishuOAuthRedirectUri: 'https://kewenqueen.github.io/power-analysis/',
  // 管理员邮箱与 GitHub 用户名白名单。GitHub 登录必须经 OAuth 验证。
  adminEmails: ['hukehuan@bytedance.com'],
  adminGithubLogins: ['KewenQueen'],
};

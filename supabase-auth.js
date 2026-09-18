/* global supabase */
(() => {
  const config = window.POWER_ANALYSIS_CONFIG || {};
  const url = config.supabaseUrl;
  const key = config.supabasePublishableKey;
  const adminEmails = (config.adminEmails || []).map((email) => String(email).trim().toLowerCase()).filter(Boolean);

  let client = null;
  let currentUser = null;
  let resolveReady;
  const ready = new Promise((resolve) => { resolveReady = resolve; });

  function isAdmin(user = currentUser) {
    return Boolean(user && adminEmails.includes(String(user.email || '').toLowerCase()));
  }

  function setMessage(message, tone = 'info') {
    const el = document.getElementById('authMessage');
    if (!el) return;
    el.textContent = message || '';
    el.dataset.tone = tone;
    el.classList.toggle('hidden', !message);
  }

  function render(session) {
    currentUser = session && session.user ? session.user : null;
    const gate = document.getElementById('authGate');
    const shell = document.getElementById('appShell');
    const email = document.getElementById('currentUserEmail');
    if (gate) gate.classList.toggle('hidden', Boolean(currentUser));
    if (shell) shell.classList.toggle('hidden', !currentUser);
    if (email) email.textContent = currentUser ? currentUser.email : '';
    window.dispatchEvent(new CustomEvent('power-auth-change', { detail: { user: currentUser, isAdmin: isAdmin() } }));
  }

  async function submit(mode) {
    const emailInput = document.getElementById('authEmail');
    const passwordInput = document.getElementById('authPassword');
    const submitBtn = document.getElementById('authSubmitBtn');
    const email = String(emailInput && emailInput.value || '').trim();
    const password = String(passwordInput && passwordInput.value || '');
    if (!email || !password) {
      setMessage('请输入邮箱和密码。', 'error');
      return;
    }
    if (password.length < 6) {
      setMessage('密码至少需要 6 位。', 'error');
      return;
    }
    submitBtn.disabled = true;
    submitBtn.textContent = mode === 'signup' ? '正在注册…' : '正在登录…';
    setMessage('');
    try {
      const result = mode === 'signup'
        ? await client.auth.signUp({ email, password })
        : await client.auth.signInWithPassword({ email, password });
      if (result.error) throw result.error;
      if (mode === 'signup' && !result.data.session) {
        setMessage('注册成功，请前往邮箱完成验证后再登录。', 'success');
      } else {
        setMessage('登录成功。', 'success');
      }
    } catch (error) {
      setMessage(error && error.message ? error.message : '操作失败，请稍后重试。', 'error');
    } finally {
      submitBtn.disabled = false;
      submitBtn.textContent = mode === 'signup' ? '注册账号' : '登录';
    }
  }

  async function initialize() {
    if (!url || !key || !window.supabase) {
      setMessage('Supabase 配置缺失，暂时无法登录。', 'error');
      resolveReady(null);
      return;
    }
    client = window.supabase.createClient(url, key, {
      auth: { persistSession: true, autoRefreshToken: true, detectSessionInUrl: true },
    });
    window.powerAuth.client = client;

    const loginTab = document.getElementById('authLoginTab');
    const signupTab = document.getElementById('authSignupTab');
    const submitBtn = document.getElementById('authSubmitBtn');
    let mode = 'login';
    const setMode = (next) => {
      mode = next;
      loginTab.classList.toggle('active', mode === 'login');
      signupTab.classList.toggle('active', mode === 'signup');
      submitBtn.textContent = mode === 'signup' ? '注册账号' : '登录';
      setMessage('');
    };
    loginTab.addEventListener('click', () => setMode('login'));
    signupTab.addEventListener('click', () => setMode('signup'));
    submitBtn.addEventListener('click', () => submit(mode));
    document.getElementById('authPassword').addEventListener('keydown', (event) => {
      if (event.key === 'Enter') submit(mode);
    });
    document.getElementById('logoutBtn').addEventListener('click', () => client.auth.signOut());

    const { data } = await client.auth.getSession();
    render(data.session);
    resolveReady(data.session);
    client.auth.onAuthStateChange((_event, session) => render(session));
  }

  window.powerAuth = {
    ready,
    client,
    getUser: () => currentUser,
    isAdmin,
    signOut: () => client ? client.auth.signOut() : Promise.resolve(),
  };

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', initialize, { once: true });
  else initialize();
})();

(() => {
  const config = window.POWER_ANALYSIS_CONFIG || {};
  const url = config.supabaseUrl;
  const key = config.supabasePublishableKey;
  const adminEmails = (config.adminEmails || []).map((email) => String(email).trim().toLowerCase()).filter(Boolean);
  const adminGithubLogins = (config.adminGithubLogins || []).map((login) => String(login).trim().toLowerCase()).filter(Boolean);

  let client = null;
  let currentUser = null;
  let guestMode = false;
  let resolveReady;
  const ready = new Promise((resolve) => { resolveReady = resolve; });

  function isAdmin(user = currentUser) {
    if (!user) return false;
    const email = String(user.email || '').trim().toLowerCase();
    const provider = String(user.app_metadata && user.app_metadata.provider || '').toLowerCase();
    const metadata = user.user_metadata || {};
    const githubLogin = String(metadata.user_name || metadata.preferred_username || '').trim().toLowerCase();
    return adminEmails.includes(email) || (provider === 'github' && adminGithubLogins.includes(githubLogin));
  }

  function setMessage(message, tone = 'info') {
    const el = document.getElementById('authMessage');
    if (!el) return;
    el.textContent = message || '';
    el.dataset.tone = tone;
    el.classList.toggle('hidden', !message);
  }

  function setRecoveryMessage(message, tone = 'error') {
    const el = document.getElementById('recoveryPasswordMessage');
    if (!el) return;
    el.textContent = message || '';
    el.classList.toggle('hidden', !message);
    el.classList.toggle('text-red-600', tone === 'error');
    el.classList.toggle('text-emerald-600', tone === 'success');
  }

  function openPasswordRecoveryModal() {
    const modal = document.getElementById('passwordRecoveryModal');
    const input = document.getElementById('recoveryPassword');
    if (modal) modal.classList.remove('hidden');
    setRecoveryMessage('');
    if (input) requestAnimationFrame(() => input.focus({ preventScroll: true }));
  }

  async function saveRecoveryPassword() {
    const passwordInput = document.getElementById('recoveryPassword');
    const confirmInput = document.getElementById('recoveryPasswordConfirm');
    const button = document.getElementById('recoveryPasswordSaveBtn');
    const password = String(passwordInput && passwordInput.value || '');
    const confirmation = String(confirmInput && confirmInput.value || '');
    if (password.length < 6) {
      setRecoveryMessage('新密码至少需要 6 位。');
      passwordInput.focus({ preventScroll: true });
      return;
    }
    if (password !== confirmation) {
      setRecoveryMessage('两次输入的密码不一致。');
      confirmInput.focus({ preventScroll: true });
      return;
    }
    button.disabled = true;
    button.textContent = '正在保存…';
    setRecoveryMessage('');
    try {
      const { error } = await client.auth.updateUser({ password });
      if (error) throw error;
      passwordInput.value = '';
      confirmInput.value = '';
      setRecoveryMessage('密码设置成功，正在进入管理员页面…', 'success');
      window.history.replaceState({}, document.title, `${window.location.origin}${window.location.pathname}`);
      setTimeout(() => document.getElementById('passwordRecoveryModal').classList.add('hidden'), 900);
    } catch (error) {
      setRecoveryMessage(error && error.message ? error.message : '密码保存失败，请重新打开重置链接。');
    } finally {
      button.disabled = false;
      button.textContent = '保存新密码';
    }
  }

  function render(session, options = {}) {
    currentUser = session && session.user ? session.user : null;
    if (typeof options.guest === 'boolean') guestMode = options.guest;
    if (currentUser) guestMode = false;
    const hasAccess = Boolean(currentUser || guestMode);
    const gate = document.getElementById('authGate');
    const shell = document.getElementById('appShell');
    const email = document.getElementById('currentUserEmail');
    if (gate) gate.classList.toggle('hidden', hasAccess);
    if (shell) shell.classList.toggle('hidden', !hasAccess);
    if (email) {
      const metadata = currentUser && currentUser.user_metadata || {};
      const identity = currentUser && (currentUser.email || metadata.user_name || metadata.preferred_username);
      email.textContent = currentUser ? identity : (guestMode ? '游客模式 · 公共数据不可见' : '');
    }
    window.dispatchEvent(new CustomEvent('power-auth-change', { detail: { user: currentUser, guest: guestMode, isAdmin: isAdmin() } }));
  }

  function enterGuestMode() {
    setMessage('');
    render(null, { guest: true });
  }

  async function signInWithGitHub() {
    const button = document.getElementById('githubLoginBtn');
    if (!client || !button) return;
    button.disabled = true;
    setMessage('正在跳转到 GitHub 安全登录…');
    const redirectTo = `${window.location.origin}${window.location.pathname}`;
    const { error } = await client.auth.signInWithOAuth({
      provider: 'github',
      options: {
        redirectTo,
        scopes: 'read:user user:email',
      },
    });
    if (error) {
      setMessage(error.message || 'GitHub 登录失败，请稍后重试。', 'error');
      button.disabled = false;
    }
  }

  async function signInWithAdminPassword() {
    const passwordInput = document.getElementById('adminPasswordOnly');
    const button = document.getElementById('adminPasswordLoginBtn');
    const password = String(passwordInput && passwordInput.value || '');
    const adminEmail = adminEmails[0] || '';
    if (!password) {
      setMessage('请输入管理员密码。', 'error');
      if (passwordInput) passwordInput.focus({ preventScroll: true });
      return;
    }
    if (!adminEmail) {
      setMessage('管理员账号尚未配置，请使用其他管理员登录方式。', 'error');
      return;
    }
    button.disabled = true;
    button.innerHTML = '<span class="material-symbols-outlined">progress_activity</span> 正在登录…';
    setMessage('');
    try {
      const { data, error } = await client.auth.signInWithPassword({ email: adminEmail, password });
      if (error) throw error;
      if (!data || !data.user || !isAdmin(data.user)) {
        await client.auth.signOut();
        throw new Error('当前账号没有管理员权限。');
      }
      passwordInput.value = '';
      setMessage('管理员登录成功。', 'success');
    } catch (error) {
      const message = error && error.message === '当前账号没有管理员权限。'
        ? error.message
        : '管理员密码不正确，请重试。';
      setMessage(message, 'error');
      passwordInput.select();
    } finally {
      button.disabled = false;
      button.innerHTML = '<span class="material-symbols-outlined">key</span> 管理员密码登录';
    }
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
    const emailInput = document.getElementById('authEmail');
    const passwordInput = document.getElementById('authPassword');
    const passwordToggleBtn = document.getElementById('passwordToggleBtn');
    const adminLoginToggleBtn = document.getElementById('adminLoginToggleBtn');
    const adminLoginOptions = document.getElementById('adminLoginOptions');
    const adminPasswordOnly = document.getElementById('adminPasswordOnly');
    const adminPasswordLoginBtn = document.getElementById('adminPasswordLoginBtn');
    const recoveryPasswordInput = document.getElementById('recoveryPassword');
    const recoveryPasswordConfirm = document.getElementById('recoveryPasswordConfirm');
    const recoveryPasswordSaveBtn = document.getElementById('recoveryPasswordSaveBtn');
    let mode = 'login';
    const setMode = (next) => {
      mode = next;
      loginTab.classList.toggle('active', mode === 'login');
      signupTab.classList.toggle('active', mode === 'signup');
      submitBtn.textContent = mode === 'signup' ? '注册账号' : '登录';
      setMessage('');
    };
    const setAdminOptions = (expanded) => {
      adminLoginOptions.classList.toggle('hidden', !expanded);
      adminLoginToggleBtn.classList.toggle('expanded', expanded);
      adminLoginToggleBtn.setAttribute('aria-expanded', String(expanded));
    };
    adminLoginToggleBtn.addEventListener('click', () => {
      const expanded = adminLoginToggleBtn.getAttribute('aria-expanded') !== 'true';
      setAdminOptions(expanded);
      if (expanded) requestAnimationFrame(() => adminPasswordOnly.focus({ preventScroll: true }));
    });
    adminPasswordLoginBtn.addEventListener('click', signInWithAdminPassword);
    adminPasswordOnly.addEventListener('keydown', (event) => {
      if (event.key === 'Enter') signInWithAdminPassword();
    });
    recoveryPasswordSaveBtn.addEventListener('click', saveRecoveryPassword);
    [recoveryPasswordInput, recoveryPasswordConfirm].forEach((input) => {
      input.addEventListener('keydown', (event) => {
        if (event.key === 'Enter') saveRecoveryPassword();
      });
    });
    document.getElementById('adminEmailLoginBtn').addEventListener('click', () => {
      setMode('login');
      setMessage('请输入管理员邮箱和密码，登录后将自动校验管理员权限。');
      emailInput.focus({ preventScroll: true });
    });
    loginTab.addEventListener('click', () => setMode('login'));
    signupTab.addEventListener('click', () => setMode('signup'));
    submitBtn.addEventListener('click', () => submit(mode));
    document.getElementById('githubLoginBtn').addEventListener('click', signInWithGitHub);
    passwordToggleBtn.addEventListener('click', () => {
      const willShow = passwordInput.type === 'password';
      passwordInput.type = willShow ? 'text' : 'password';
      passwordToggleBtn.textContent = willShow ? '隐藏' : '显示';
      passwordToggleBtn.setAttribute('aria-label', willShow ? '隐藏密码' : '显示密码');
      passwordToggleBtn.setAttribute('aria-pressed', String(willShow));
      passwordInput.focus({ preventScroll: true });
    });
    passwordInput.addEventListener('keydown', (event) => {
      if (event.key === 'Enter') submit(mode);
    });
    document.getElementById('guestModeBtn').addEventListener('click', enterGuestMode);
    document.getElementById('logoutBtn').addEventListener('click', async () => {
      if (guestMode) {
        render(null, { guest: false });
        return;
      }
      await client.auth.signOut();
    });

    const recoveryUrl = window.location.hash.includes('type=recovery') || window.location.search.includes('type=recovery');
    const { data } = await client.auth.getSession();
    render(data.session, { guest: false });
    if (data.session && recoveryUrl) openPasswordRecoveryModal();
    resolveReady(data.session);
    client.auth.onAuthStateChange((event, session) => {
      render(session, { guest: guestMode && !session });
      if (event === 'PASSWORD_RECOVERY') openPasswordRecoveryModal();
    });
  }

  window.powerAuth = {
    ready,
    client,
    getUser: () => currentUser,
    isGuest: () => guestMode,
    isAdmin,
    signInWithGitHub,
    enterGuestMode,
    signOut: async () => {
      if (guestMode) return render(null, { guest: false });
      return client ? client.auth.signOut() : Promise.resolve();
    },
  };

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', initialize, { once: true });
  else initialize();
})();

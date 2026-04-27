/* =========================================================
   Control OC Mantenimiento - Login actual
   Login local temporal. Más adelante se puede reemplazar por Firebase/Auth real.
   ========================================================= */

(function () {
  const AUTH_KEY = 'control_oc_auth';
  const AUTH_ATTEMPTS_KEY = 'control_oc_auth_attempts';
  const allowedUsers = [
    { email: 'juan@d1.com', password: '123' },
    { email: 'mtto@d1.com', password: 'D1SAS.CST01' }
  ];
  const MAX_AUTH_ATTEMPTS = 9999;
  const AUTH_LOCK_MINUTES = 0;
  const authOverlay = document.getElementById('authOverlay');
  const authForm = document.getElementById('authForm');
  const authCard = document.querySelector('.auth-card');
  const authEmail = document.getElementById('authEmail');
  const authPassword = document.getElementById('authPassword');
  const authRemember = document.getElementById('authRemember');
  const authSubmitBtn = document.getElementById('authSubmitBtn');
  const authError = document.getElementById('authError');
  const authHint = document.getElementById('authHint');
  const authThemeBtn = document.getElementById('authThemeBtn');
  const authUserPill = document.getElementById('authUserPill');
  const logoutBtn = document.getElementById('logoutBtn');
  const authEyeShell = document.getElementById('authEyeShell');
  const authEye = document.getElementById('authEye');
  const authPupil = document.getElementById('authPupil');
  const authMascotText = document.getElementById('authMascotText');
  const authEmailShell = document.getElementById('authEmailShell');
  const authPasswordShell = document.getElementById('authPasswordShell');

  function readAuthAttempts() {
    try {
      return JSON.parse(localStorage.getItem(AUTH_ATTEMPTS_KEY)) || { count: 0, lockUntil: null };
    } catch (_) {
      return { count: 0, lockUntil: null };
    }
  }

  function writeAuthAttempts(data) {
    localStorage.setItem(AUTH_ATTEMPTS_KEY, JSON.stringify(data));
  }

  function clearAuthAttempts() {
    writeAuthAttempts({ count: 0, lockUntil: null });
  }

  function lockRemainingMs() {
    const data = readAuthAttempts();
    if (!data.lockUntil) return 0;
    return Math.max(0, data.lockUntil - Date.now());
  }

  function formatRemaining(ms) {
    const mins = Math.ceil(ms / 60000);
    return `${mins} minuto${mins === 1 ? '' : 's'}`;
  }

  function setFieldError(fieldShell, isError) {
    fieldShell?.classList.toggle('error', !!isError);
  }

  function pulseCard(type) {
    if (!authCard) return;
    authCard.classList.remove('login-error', 'login-success');
    void authCard.offsetWidth;
    authCard.classList.add(type === 'success' ? 'login-success' : 'login-error');
    setTimeout(() => authCard.classList.remove('login-error', 'login-success'), 550);
  }

  function showHint(message) {
    if (!authHint) return;
    authHint.textContent = message;
    authHint.style.display = 'block';
  }

  let angryTimer = null;
  function triggerAngryEye(duration = 1200) {
    if (!authEyeShell) return;
    authEyeShell.classList.add('angry');
    clearTimeout(angryTimer);
    angryTimer = setTimeout(() => authEyeShell?.classList.remove('angry'), duration);
  }

  function clearHint() {
    if (!authHint) return;
    authHint.textContent = '';
    authHint.style.display = 'none';
  }

  function showAuthError(message, opts = {}) {
    if (!authError) return;
    authError.textContent = message;
    authError.dataset.text = message;
    authError.classList.remove('success', 'glitch-warning');
    if (opts.success) authError.classList.add('success');
    if (opts.glitch && !opts.success) authError.classList.add('glitch-warning');
    authError.classList.add('show');
    if (!opts.success) {
      pulseCard('error');
      triggerAngryEye(opts.angryDuration || 1400);
      triggerEyeBlink(160, true);
    }
  }

  function clearAuthError() {
    if (!authError) return;
    authError.textContent = '';
    authError.classList.remove('show', 'success');
    setFieldError(authEmailShell, false);
    setFieldError(authPasswordShell, false);
  }

  let blinkTimer = null;

  function triggerEyeBlink(duration = 180, force = false) {
    if (!authEyeShell) return;
    if (!force && authEyeShell.classList.contains('password-focus')) return;
    authEyeShell.classList.add('blink');
    clearTimeout(blinkTimer);
    blinkTimer = setTimeout(() => authEyeShell?.classList.remove('blink'), duration);
  }

  function triggerAuthCharge() {
    if (!authSubmitBtn) return;
    authSubmitBtn.classList.remove('auth-submit-blink');
    void authSubmitBtn.offsetWidth;
    authSubmitBtn.classList.add('auth-submit-blink');
    setTimeout(() => authSubmitBtn?.classList.remove('auth-submit-blink'), 1200);
  }

  function setAuthFocus(target) {
    authEmailShell?.classList.toggle('active', target === 'email');
    authPasswordShell?.classList.toggle('active', target === 'password');
    if (authEyeShell) {
      authEyeShell.classList.toggle('email-focus', target === 'email');
      authEyeShell.classList.toggle('password-focus', target === 'password');
      if (target === 'password') {
        authEyeShell.classList.remove('blink');
      } else if (target === 'email') {
        triggerEyeBlink(150);
      }
    }
    if (authMascotText) {
      authMascotText.textContent = target === 'email' ? 'El ojo te mira mientras escribes 👀' : '';
    }
  }

  function movePupil(clientX, clientY) {
    if (!authEye || !authPupil || authEyeShell?.classList.contains('password-focus')) return;
    const rect = authEye.getBoundingClientRect();
    const cx = rect.left + rect.width / 2;
    const cy = rect.top + rect.height / 2;
    const dx = clientX - cx;
    const dy = clientY - cy;
    const max = 7;
    const distance = Math.max(Math.hypot(dx, dy), 1);
    const x = (dx / distance) * Math.min(max, Math.abs(dx) / 12);
    const y = (dy / distance) * Math.min(max * 0.75, Math.abs(dy) / 14);
    authPupil.style.transform = `translate(calc(-50% + ${x}px), calc(-50% + ${y}px))`;
  }

  function centerPupil() {
    if (authPupil) authPupil.style.transform = 'translate(-50%, -50%)';
  }

  setInterval(() => {
    if (!authOverlay || authOverlay.style.display === 'none') return;
    if (document.activeElement === authPassword || authEyeShell?.classList.contains('password-focus')) return;
    triggerEyeBlink(170);
  }, 3800);

  function applyTheme(mode) {
    document.body.classList.remove('dark', 'light');
    document.body.classList.add(mode);
    try { localStorage.setItem('control_oc_theme', mode); } catch (e) {}
    if (authThemeBtn) {
      authThemeBtn.textContent = mode === 'dark' ? '☀️' : '🌙';
      authThemeBtn.title = mode === 'dark' ? 'Cambiar a modo claro' : 'Cambiar a modo oscuro';
    }
  }

  function setAuthenticated(email) {
    document.body.classList.remove('auth-locked');
    if (authOverlay) authOverlay.style.display = 'none';
    if (authUserPill) {
      authUserPill.style.display = 'inline-flex';
      authUserPill.textContent = '👤 ' + email;
    }
    if (logoutBtn) logoutBtn.style.display = 'inline-flex';
    setAuthFocus(null);
    centerPupil();
    clearAuthError();
    clearHint();
  }

  function persistAuth(email) {
    const payload = JSON.stringify({ email, ts: Date.now() });
    const storage = authRemember.checked ? localStorage : sessionStorage;
    storage.setItem(AUTH_KEY, payload);
    const other = authRemember.checked ? sessionStorage : localStorage;
    other.removeItem(AUTH_KEY);
  }

  function clearAuth() {
    localStorage.removeItem(AUTH_KEY);
    sessionStorage.removeItem(AUTH_KEY);
    document.body.classList.add('auth-locked');
    if (authOverlay) authOverlay.style.display = 'flex';
    if (authUserPill) {
      authUserPill.style.display = 'none';
      authUserPill.textContent = '';
    }
    if (logoutBtn) logoutBtn.style.display = 'none';
    authPassword.value = '';
    clearAuthError();
    showHint('Usuarios autorizados: juan@d1.com / 123 · mtto@d1.com / D1SAS.CST01');
    setAuthFocus('password');
    authPassword.focus();
  }


  function checkLockState(show = true) {
    const remaining = lockRemainingMs();
    if (remaining > 0) {
      if (show) {
        setFieldError(authEmailShell, true);
        setFieldError(authPasswordShell, true);
        showAuthError(`🛑 ALERTA: acceso bloqueado temporalmente. Intenta de nuevo en ${formatRemaining(remaining)}.`, { glitch: true, angryDuration: 1800 });
        showHint('Demasiados intentos fallidos. Espera un momento antes de volver a intentar.');
      }
      return true;
    }
    const data = readAuthAttempts();
    if (data.lockUntil && remaining === 0) clearAuthAttempts();
    return false;
  }

  function handleFailedLogin(emailValid) {
    const data = readAuthAttempts();
    const nextCount = (data.count || 0) + 1;
    if (nextCount >= MAX_AUTH_ATTEMPTS) {
      const lockUntil = Date.now() + AUTH_LOCK_MINUTES * 60 * 1000;
      writeAuthAttempts({ count: nextCount, lockUntil });
      setFieldError(authEmailShell, true);
      setFieldError(authPasswordShell, true);
      showAuthError(`🛑 ALERTA CRÍTICA: acceso bloqueado por ${AUTH_LOCK_MINUTES} minutos por demasiados intentos.`, { glitch: true, angryDuration: 2200 });
      showHint('Tip: revisa bien el correo y la contraseña antes de volver a probar.');
      return;
    }
    writeAuthAttempts({ count: nextCount, lockUntil: null });
    setFieldError(authEmailShell, !emailValid);
    setFieldError(authPasswordShell, true);
    const remaining = MAX_AUTH_ATTEMPTS - nextCount;
    showAuthError(emailValid
      ? `❌ ADVERTENCIA: contraseña incorrecta. Intentos restantes: ${remaining}.`
      : `❌ ADVERTENCIA: correo no autorizado o incorrecto. Intentos restantes: ${remaining}.`,
      { glitch: true, angryDuration: 1600 });
    showHint('Usuarios autorizados: juan@d1.com / 123 · mtto@d1.com / D1SAS.CST01');
  }

  authThemeBtn?.addEventListener('click', () => {
    const next = document.body.classList.contains('dark') ? 'light' : 'dark';
    applyTheme(next);
  });

  authEmail?.addEventListener('input', () => {
    clearAuthError();
    clearHint();
    checkLockState(false);
  });
  authPassword?.addEventListener('input', () => {
    clearAuthError();
    clearHint();
    checkLockState(false);
  });

  authPassword?.addEventListener('keydown', (event) => {
    if (event.getModifierState && event.getModifierState('CapsLock')) {
      showHint('⚠️ Caps Lock está activado. Revisa la contraseña antes de enviar.');
    }
  });

  authEmail?.addEventListener('keydown', () => triggerEyeBlink(140));
  authEyeShell?.addEventListener('click', () => {
    triggerEyeBlink(260, true);
    triggerAuthCharge();
    if (authMascotText && document.activeElement === authEmail) {
      authMascotText.textContent = 'El ojo te mira mientras escribes 👀';
    }
  });

  authEmail?.addEventListener('focus', () => setAuthFocus('email'));
  authPassword?.addEventListener('focus', () => setAuthFocus('password'));
  authEmail?.addEventListener('blur', () => setTimeout(() => {
    if (document.activeElement !== authPassword) setAuthFocus(null);
  }, 0));
  authPassword?.addEventListener('blur', () => setTimeout(() => {
    if (document.activeElement !== authEmail) setAuthFocus(null);
  }, 0));

  authOverlay?.addEventListener('mousemove', (event) => movePupil(event.clientX, event.clientY));
  authOverlay?.addEventListener('mouseleave', centerPupil);

  authForm?.addEventListener('submit', (event) => {
    event.preventDefault();
    clearAuthError();
    clearHint();

    if (checkLockState()) return;

    const email = String(authEmail.value || '').trim().toLowerCase();
    const password = String(authPassword.value || '');

    if (!email && !password) {
      setFieldError(authEmailShell, true);
      setFieldError(authPasswordShell, true);
      showAuthError('⚠️ ADVERTENCIA: debes ingresar correo y contraseña antes de continuar.', { glitch: true, angryDuration: 1000 });
      showHint('Usuarios autorizados: juan@d1.com / 123 · mtto@d1.com / D1SAS.CST01');
      return;
    }
    if (!email) {
      setFieldError(authEmailShell, true);
      setFieldError(authPasswordShell, false);
      showAuthError('⚠️ ADVERTENCIA: falta el correo de acceso.', { glitch: true, angryDuration: 900 });
      return;
    }
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      setFieldError(authEmailShell, true);
      showAuthError('⚠️ ADVERTENCIA: el formato del correo no es válido.', { glitch: true, angryDuration: 900 });
      return;
    }
    if (!password) {
      setFieldError(authPasswordShell, true);
      setFieldError(authEmailShell, false);
      showAuthError('⚠️ ADVERTENCIA: falta la contraseña.', { glitch: true, angryDuration: 900 });
      return;
    }

    const userMatch = allowedUsers.find(u => u.email.toLowerCase() === email && u.password === password);
    const emailValid = allowedUsers.some(u => u.email.toLowerCase() === email);

    if (!userMatch) {
      handleFailedLogin(emailValid);
      return;
    }

    clearAuthAttempts();
    persistAuth(email);
    pulseCard('success');
    showAuthError('Acceso concedido. Abriendo el aplicativo…', { success: true });
    showHint('Ingreso correcto.');
    setTimeout(() => {
      setAuthenticated(email);
      if (typeof toast === 'function') toast('Acceso concedido ✅', 'success', 1400);
    }, 520);
  });

  logoutBtn?.addEventListener('click', () => {
    clearAuth();
    if (typeof toast === 'function') toast('Sesión cerrada 🔒', 'info', 1400);
  });

  try {
    const savedTheme = localStorage.getItem('control_oc_theme') || (document.body.classList.contains('light') ? 'light' : 'dark');
    applyTheme(savedTheme);
  } catch (e) {
    applyTheme('dark');
  }

  document.body.classList.add('auth-locked');
  if (authOverlay) authOverlay.style.display = 'flex';
  if (authUserPill) { authUserPill.style.display = 'none'; authUserPill.textContent = ''; }
  if (logoutBtn) logoutBtn.style.display = 'none';
  if (authEmail) authEmail.value = '';
  if (authPassword) authPassword.value = '';
  clearAuthError();
  showHint('Usuarios autorizados: juan@d1.com / 123 · mtto@d1.com / D1SAS.CST01');
  checkLockState(false);
  setAuthFocus('email');
  authEmail.focus();
})();

/**
 * 認証モジュール
 * セッション管理、ログイン、パスワード変更
 */

// セッションはCacheServiceで管理（6時間有効）
const SESSION_DURATION = 21600; // 6時間（秒）

/**
 * ログイン処理
 */
function login(employeeId, password) {
  const empIdStr = String(employeeId);
  const pwStr = String(password);

  const user = findUserByEmployeeId(empIdStr);
  if (!user) {
    return { success: false, message: '社員番号またはパスワードが正しくありません' };
  }
  if (!user.active) {
    return { success: false, message: 'このアカウントは無効です' };
  }

  const inputHash = hashPassword(pwStr);
  if (inputHash !== String(user.password)) {
    return { success: false, message: '社員番号またはパスワードが正しくありません' };
  }

  // セッショントークン生成
  const token = Utilities.getUuid();
  const sessionData = JSON.stringify({
    employeeId: String(user.employeeId),
    name: user.name,
    role: user.role,
    category: user.category,
    loginAt: new Date().toISOString()
  });

  CacheService.getScriptCache().put('session_' + token, sessionData, SESSION_DURATION);

  return {
    success: true,
    token: token,
    user: {
      employeeId: user.employeeId,
      name: user.name,
      role: user.role,
      category: user.category
    }
  };
}

/**
 * セッション検証
 */
function validateSession(token) {
  if (!token) return null;
  const cache = CacheService.getScriptCache();
  const sessionData = cache.get('session_' + token);
  if (!sessionData) return null;

  // セッション更新（アクティブなうちは延長）
  cache.put('session_' + token, sessionData, SESSION_DURATION);

  return JSON.parse(sessionData);
}

/**
 * ログアウト
 */
function logout(token) {
  if (token) {
    CacheService.getScriptCache().remove('session_' + token);
  }
  return { success: true };
}

/**
 * パスワード変更
 */
function changePassword(token, currentPassword, newPassword) {
  const session = validateSession(token);
  if (!session) {
    return { success: false, message: 'セッションが無効です。再ログインしてください。' };
  }

  if (!newPassword || newPassword.length < 4) {
    return { success: false, message: 'パスワードは4文字以上で設定してください。' };
  }

  const user = findUserByEmployeeId(session.employeeId);
  if (!user) {
    return { success: false, message: 'ユーザーが見つかりません。' };
  }

  const currentHash = hashPassword(String(currentPassword));
  if (currentHash !== String(user.password)) {
    return { success: false, message: '現在のパスワードが正しくありません。' };
  }

  const newHash = hashPassword(newPassword);
  updateUserPassword(session.employeeId, newHash);

  return { success: true, message: 'パスワードを変更しました。' };
}

/**
 * 管理者権限チェック
 */
function requireAdmin(token) {
  const session = validateSession(token);
  if (!session) {
    return { valid: false, message: 'セッションが無効です' };
  }
  if (session.role !== 'admin') {
    return { valid: false, message: '管理者権限が必要です' };
  }
  return { valid: true, session: session };
}

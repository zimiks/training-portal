function verifyOTP(email, inputOtp) {
  const ss = getOrCreateSpreadsheet_();
  const sh = ss.getSheetByName('OTPRequests');
  const data = sh.getDataRange().getValues();

  for (let i = data.length - 1; i > 0; i--) {
    const rowEmail = data[i][1];
    const otp = data[i][2];
    const expiry = new Date(data[i][3]);
    const status = data[i][4];

    if (rowEmail === email && status === 'SENT') {
      if (new Date() > expiry) return { success: false, message: 'OTP expired' };
      if (otp === inputOtp) {
        sh.getRange(i + 1, 5).setValue('VERIFIED');
        sh.getRange(i + 1, 6).setValue(new Date());
        return { success: true };
      }
      return { success: false, message: 'Invalid OTP' };
    }
  }

  return { success: false, message: 'OTP not found' };
}
function checkSession_(email) {
  const ss = getOrCreateSpreadsheet_();
  const sh = ss.getSheetByName('AuthSessions');
  const data = sh.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email) {
      const lastActivity = new Date(data[i][2]);
      const lastOTP = new Date(data[i][3]);

      const diffMinutes = (Date.now() - lastActivity.getTime()) / 60000;

      if (diffMinutes > 60) {
        return { requireOTP: true };
      }

      return { requireOTP: false };
    }
  }

  return { requireOTP: true }; // first time
}
function getAuthBootstrap_(ss, user) {
  const email = String((user || {}).email || '').trim();
  if (!email || email === 'Not available') {
    return {
      requireOtp: true,
      mode: 'first_login',
      message: 'Your account could not be verified. Please sign in with your company Google account.',
      reverifyMinutes: 60
    };
  }

  const sessionState = getSessionState_(ss, email);

  return {
    requireOtp: sessionState.requireOtp,
    mode: sessionState.mode,
    message: sessionState.message,
    reverifyMinutes: sessionState.reverifyMinutes
  };
}
function sendOtpEmail_(toEmail, subject, body) {
  GmailApp.sendEmail(toEmail, subject, body, {
    from: CONFIG.NO_REPLY_EMAIL,
    name: 'First Connect Health',
    replyTo: CONFIG.SUPPORT_EMAIL
  });
}
function sendLoginOtp() {
  const ctx = requireEnabledPortalUser_();
  const email = String(ctx.user.email || '').trim().toLowerCase();
  const cache = CacheService.getScriptCache();
  const lock = LockService.getScriptLock();
  const otp = String(Math.floor(100000 + Math.random() * 900000));
  const preferredAlias = String(CONFIG.NO_REPLY_EMAIL || '').trim().toLowerCase();

  if (!email || email === 'not available') {
    throw new Error('User email is not available.');
  }

  lock.waitLock(3000);
  try {
    const cooldownKey = buildPortalOtpCooldownCacheKey_(email);
    const cooldownActive = cache.get(cooldownKey);
    if (cooldownActive) {
      return {
        success: false,
        message: 'Please wait 60 seconds before requesting another OTP.'
      };
    }

    cache.put(buildPortalOtpCacheKey_(email), otp, Number(CONFIG.OTP_EXPIRY_SECONDS || 300));
    cache.put(buildPortalOtpAttemptsCacheKey_(email), '0', Number(CONFIG.OTP_EXPIRY_SECONDS || 300));
    cache.put(cooldownKey, 'Y', Number(CONFIG.OTP_SEND_COOLDOWN_SECONDS || 60));
    cache.remove(buildPortalOtpVerifiedCacheKey_(email));
  } finally {
    lock.releaseLock();
  }

  const ss = getOrCreateSpreadsheet_();
  const sh = ss.getSheetByName('OTPRequests');
  const now = new Date();
  const expiry = new Date(now.getTime() + (Number(CONFIG.OTP_EXPIRY_SECONDS || 300) * 1000));

  sh.appendRow([
    now,
    email,
    otp,
    expiry,
    'SENT',
    ''
  ]);

  const subject = 'Training Portal OTP';
  const body = [
    'Hello ' + (ctx.user.name || 'User') + ',',
    '',
    'Your OTP is: ' + otp,
    '',
    'This OTP will expire in 5 minutes.',
    'If you did not request this, please ignore this email.',
    '',
    'Training Portal'
  ].join('\n');

  GmailApp.sendEmail(email, subject, body, {
    from: preferredAlias,
    name: 'First Connect Health',
    replyTo: CONFIG.SUPPORT_EMAIL
  });

  return {
    success: true,
    message: 'OTP sent to your enabled email address.'
  };
}
function debugCurrentPortalUser() {
  const user = getUserInfo_();
  const access = evaluateUserAccess_(user);

  return {
    email: user.email,
    domain: user.domain,
    name: user.name,
    domainAllowed: access.domainAllowed,
    userFound: access.userFound,
    enabled: access.enabled,
    allowed: access.allowed,
    role: access.role
  };
}
function verifyLoginOtp(otpInput) {
  const user = getUserInfo_();
  const access = evaluateUserAccess_(user);
  if (!access.domainAllowed) throw new Error('Access denied.');

  const ss = getOrCreateSpreadsheet_();
  const email = String(user.email || '').trim();
  const otp = String(otpInput || '').trim();

  if (!otp) {
    return { success: false, message: 'OTP is required.' };
  }

  const sh = ss.getSheetByName('OTPRequests');
  const values = sh.getDataRange().getValues();

  for (let i = values.length - 1; i >= 1; i--) {
    const rowEmail = String(values[i][1] || '').trim();
    const rowOtp = String(values[i][2] || '').trim();
    const expiry = new Date(values[i][3]);
    const status = String(values[i][4] || '').trim();

    if (rowEmail === email && status === 'SENT') {
      if (new Date() > expiry) {
        sh.getRange(i + 1, 5).setValue('EXPIRED');
        writeAuditLog_(ss, email, 'OTP_VERIFY', 'OTP expired', 'FAILED');
        return { success: false, message: 'OTP expired. Please request a new OTP.' };
      }

      if (rowOtp !== otp) {
        writeAuditLog_(ss, email, 'OTP_VERIFY', 'Invalid OTP entered', 'FAILED');
        return { success: false, message: 'Invalid OTP.' };
      }

      sh.getRange(i + 1, 5).setValue('VERIFIED');
      sh.getRange(i + 1, 6).setValue(new Date());

      upsertSession_(ss, email, true);
      writeAuditLog_(ss, email, 'OTP_VERIFY', 'OTP verified', 'SUCCESS');

      return { success: true, message: 'OTP verified successfully.' };
    }
  }

  return { success: false, message: 'No active OTP found. Please request a new OTP.' };
}

function updateUserActivity() {
  const user = getUserInfo_();
  const access = evaluateUserAccess_(user);
  if (!access.domainAllowed) throw new Error('Access denied.');

  const ss = getOrCreateSpreadsheet_();
  const email = String(user.email || '').trim();
  touchSession_(ss, email);
  return true;
}

function markSessionReverificationRequired() {
  const user = getUserInfo_();
  const access = evaluateUserAccess_(user);
  if (!access.domainAllowed) throw new Error('Access denied.');

  const ss = getOrCreateSpreadsheet_();
  const email = String(user.email || '').trim();

  setSessionStatus_(ss, email, 'REVERIFY_REQUIRED');
  writeAuditLog_(ss, email, 'SESSION_EXPIRED', 'Reverification required after inactivity', 'INFO');
  return true;
}

function getSessionState_(ss, email) {
  const sh = ss.getSheetByName('AuthSessions');
  const values = sh.getDataRange().getValues();
  const reverifyMinutes = 60;

  for (let i = 1; i < values.length; i++) {
    const rowEmail = String(values[i][0] || '').trim();
    const lastActivity = values[i][2] ? new Date(values[i][2]) : null;
    const status = String(values[i][4] || '').trim();

    if (rowEmail === email) {
      if (status === 'REVERIFY_REQUIRED') {
        return {
          requireOtp: true,
          mode: 'reauth',
          message: 'You were inactive for more than 60 minutes. Please verify OTP to continue.',
          reverifyMinutes: reverifyMinutes
        };
      }

      if (lastActivity) {
        const diffMinutes = (Date.now() - lastActivity.getTime()) / 60000;
        if (diffMinutes > reverifyMinutes) {
          return {
            requireOtp: true,
            mode: 'reauth',
            message: 'You were inactive for more than 60 minutes. Please verify OTP to continue.',
            reverifyMinutes: reverifyMinutes
          };
        }
      }

      return {
        requireOtp: false,
        mode: '',
        message: '',
        reverifyMinutes: reverifyMinutes
      };
    }
  }

  return {
    requireOtp: true,
    mode: 'first_login',
    message: 'First-time verification required. Please verify OTP to continue.',
    reverifyMinutes: reverifyMinutes
  };
}

function upsertSession_(ss, email, otpVerified) {
  const sh = ss.getSheetByName('AuthSessions');
  const values = sh.getDataRange().getValues();
  const now = new Date();

  for (let i = 1; i < values.length; i++) {
    const rowEmail = String(values[i][0] || '').trim();
    if (rowEmail === email) {
      sh.getRange(i + 1, 2).setValue(values[i][1] || now);
      sh.getRange(i + 1, 3).setValue(now);
      sh.getRange(i + 1, 4).setValue(otpVerified ? now : values[i][3] || '');
      sh.getRange(i + 1, 5).setValue('ACTIVE');
      return;
    }
  }

  sh.appendRow([email, now, now, otpVerified ? now : '', 'ACTIVE']);
}

function touchSession_(ss, email) {
  const sh = ss.getSheetByName('AuthSessions');
  const values = sh.getDataRange().getValues();
  const now = new Date();

  for (let i = 1; i < values.length; i++) {
    const rowEmail = String(values[i][0] || '').trim();
    if (rowEmail === email) {
      sh.getRange(i + 1, 3).setValue(now);
      return;
    }
  }

  sh.appendRow([email, now, now, '', 'ACTIVE']);
}

function setSessionStatus_(ss, email, status) {
  const sh = ss.getSheetByName('AuthSessions');
  const values = sh.getDataRange().getValues();
  const now = new Date();

  for (let i = 1; i < values.length; i++) {
    const rowEmail = String(values[i][0] || '').trim();
    if (rowEmail === email) {
      sh.getRange(i + 1, 3).setValue(now);
      sh.getRange(i + 1, 5).setValue(status);
      return;
    }
  }

  sh.appendRow([email, now, now, '', status]);
}

function writeAuditLog_(ss, email, eventName, details, result) {
  const sh = ss.getSheetByName('AccessAudit');
  sh.appendRow([new Date(), email, eventName, details, result]);
}
function updateUserActivity_(email) {
  const ss = getOrCreateSpreadsheet_();
  const sh = ss.getSheetByName('AuthSessions');
  const data = sh.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email) {
      sh.getRange(i + 1, 3).setValue(new Date());
      return;
    }
  }

  sh.appendRow([email, new Date(), new Date(), new Date(), 'ACTIVE']);
}

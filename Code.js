const CONFIG = {
  APP_NAME: 'First Connect Health - Training Portal',
  DRIVE_FOLDER_NAME: 'Training Portal',
  SPREADSHEET_NAME: 'Training Portal Data',
  SUPPORT_EMAIL: 'hr@firstconnecthealth.com',
  COMPANY_WEBSITE: 'https://firstconnecthealth.com',
  ALLOW_DOMAIN_ONLY: true,
  ALLOWED_DOMAIN: 'firstconnecthealth.com',
  DEFAULT_CERTIFICATE_FOLDER_NAME: 'Training Portal Certificates',
  PASS_PERCENT: 70,
  NO_REPLY_EMAIL: 'no-reply@firstconnecthealth.com',
  OTP_EXPIRY_SECONDS: 300,
  OTP_SEND_COOLDOWN_SECONDS: 60,
  OTP_MAX_VERIFY_ATTEMPTS: 5
};

function doGet() {
  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle(CONFIG.APP_NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * RUN ONCE manually after pasting code
 */
function initializePortalData() {
  const ss = getOrCreateSpreadsheet_();
  seedUsersSheet_(ss);
  seedModulesSheet_(ss);
  seedProgressSheet_(ss);
  seedQuizzesSheet_(ss);
  seedQuizResultsSheet_(ss);
  seedCertificatesSheet_(ss);
  seedDailyLearningSheet_(ss);
  return 'Training Portal data initialized: ' + ss.getUrl();
}

/**
 * Main bootstrap for UI
 */

function getPortalBootstrap() { 
  const user = getUserInfo_();
  const access = evaluateUserAccess_(user);
  const ss = getOrCreateSpreadsheet_();

  const portalData = {
    appName: CONFIG.APP_NAME,
    supportEmail: CONFIG.SUPPORT_EMAIL,
    companyWebsite: CONFIG.COMPANY_WEBSITE,
    user: user,
    access: access,
    settings: getSettingsCached_(ss),
    auth: {
      requireOtp: false,
      mode: '',
      message: '',
      reverifyMinutes: 60
    },

    dailyLearning: {
      items: [],
      todayCard: null,
      progress: { seenCardIds: [], learnedCardIds: [] },
      streak: { currentStreak: 0, bestStreak: 0, totalXp: 0, lastLearnedDate: '' },
      meta: { streak: 0, todayCompleted: false, totalXp: 0, todayXp: 0 },
      weeklyStatus: [],
      leaderboard: []
    },

    updates: [],
    progressSummary: {
      progressPercent: 0,
      completedLessons: 0,
      inProgressLessons: 0,
      totalLessons: 0,
      currentWeek: '',
      nextTask: ''
    },
    quizSummary: {
      quizzesTaken: 0,
      avgQuizScore: 0,
      certificatesEarned: 0,
      certificates: []
    },
    dashboardStats: {
      moduleCount: 0,
      resourceCount: 0,
      updateCount: 0,
      progressPercent: 0,
      completedLessons: 0,
      totalLessons: 0,
      currentWeek: '',
      nextTask: '',
      quizzesTaken: 0,
      avgQuizScore: 0,
      certificatesEarned: 0
    },

    modules: [],
    progressMap: {},
    resources: [],
    adminSnapshot: access.isAdmin ? getAdminSnapshot_(user.email) : null
  };

  portalData.recommendedAction = getRecommendedAction_(portalData);

  if (!access.allowed) {
    portalData.restricted = {
      title: 'Access Denied',
      message: 'This website is not available for your account.',
      detail: 'Please contact the administrator if you believe this is an error.'
    };
    return portalData;
  }

  const authState = getAuthBootstrap_(ss, user);
  const dailyLearningItems = getDailyLearningForUser_(ss, user);
  const updates = getUpdatesCached_(ss);
  const todayCard = pickDailyLearningCard_(dailyLearningItems);

  const progressSummary = getProgressSummary_(ss, user.email);
  const quizSummary = getQuizSummary_(ss, user.email);
  const dashboardStats = getDashboardStatsLight_(ss, user.email, progressSummary, quizSummary);
  const modules = getModulesCached_(ss);
  const progressMap = getProgressMap_(ss, user.email);

  portalData.auth = authState;
  portalData.dailyLearning = {
    items: dailyLearningItems,
    todayCard: todayCard,
    progress: getDailyLearningUserProgress_(ss, user.email),
    streak: getUserStreak_(ss, user.email),
    meta: getDailyLearningMeta_(user.email, todayCard),
    weeklyStatus: getWeeklyDailyStatus_(user.email),
    leaderboard: getDailyLearningLeaderboard_(ss, user.email, 5)
  };
  portalData.updates = updates;
  portalData.progressSummary = progressSummary;
  portalData.quizSummary = quizSummary;
  portalData.dashboardStats = dashboardStats;
  portalData.modules = modules;
  portalData.progressMap = progressMap;
  portalData.recommendedAction = getRecommendedAction_(portalData);

  return portalData;
}
function getInitialPortalBootstrap() {
  const user = getUserInfo_();
  const access = evaluateUserAccess_(user);
  const ss = getOrCreateSpreadsheet_();
  const authState = getAuthBootstrap_(ss, user);

  const portalData = {
    appName: CONFIG.APP_NAME,
    supportEmail: CONFIG.SUPPORT_EMAIL,
    companyWebsite: CONFIG.COMPANY_WEBSITE,
    user: user,
    access: access,
    settings: getSettingsCached_(ss),
    auth: authState,

    dailyLearning: {
      items: [],
      todayCard: null,
      progress: { seenCardIds: [], learnedCardIds: [] },
      streak: { currentStreak: 0, bestStreak: 0, totalXp: 0, lastLearnedDate: '' },
      meta: { streak: 0, todayCompleted: false, totalXp: 0, todayXp: 0 },
      weeklyStatus: [],
      leaderboard: []
    },

    updates: [],
    progressSummary: {
      progressPercent: 0,
      completedLessons: 0,
      inProgressLessons: 0,
      totalLessons: 0,
      currentWeek: '',
      nextTask: ''
    },
    quizSummary: {
      quizzesTaken: 0,
      avgQuizScore: 0,
      certificatesEarned: 0,
      certificates: []
    },
    dashboardStats: {
      moduleCount: 0,
      resourceCount: 0,
      updateCount: 0,
      progressPercent: 0,
      completedLessons: 0,
      totalLessons: 0,
      currentWeek: '',
      nextTask: '',
      quizzesTaken: 0,
      avgQuizScore: 0,
      certificatesEarned: 0
    },

    modules: [],
    progressMap: {},
    resources: [],
    adminSnapshot: null,
    recommendedAction: {
      type: 'explore',
      label: 'Explore your training portal',
      page: 'dashboard'
    }
  };

  if (!access.allowed) {
    portalData.restricted = {
      title: 'Access Denied',
      message: 'This website is not available for your account.',
      detail: 'Please contact the administrator if you believe this is an error.'
    };
    return portalData;
  }

  return portalData;
}
function getDeferredPortalBootstrapData() {
  const user = getUserInfo_();
  const access = evaluateUserAccess_(user);
  const ss = getOrCreateSpreadsheet_();

  if (!access.allowed) {
    return {
      dailyLearning: {
        items: [],
        todayCard: null,
        progress: { seenCardIds: [], learnedCardIds: [] },
        streak: { currentStreak: 0, bestStreak: 0, totalXp: 0, lastLearnedDate: '' },
        meta: { streak: 0, todayCompleted: false, totalXp: 0, todayXp: 0 },
        weeklyStatus: [],
        leaderboard: []
      },
      updates: [],
      progressSummary: {
        progressPercent: 0,
        completedLessons: 0,
        inProgressLessons: 0,
        totalLessons: 0,
        currentWeek: '',
        nextTask: ''
      },
      quizSummary: {
        quizzesTaken: 0,
        avgQuizScore: 0,
        certificatesEarned: 0,
        certificates: []
      },
      dashboardStats: {
        moduleCount: 0,
        resourceCount: 0,
        updateCount: 0,
        progressPercent: 0,
        completedLessons: 0,
        totalLessons: 0,
        currentWeek: '',
        nextTask: '',
        quizzesTaken: 0,
        avgQuizScore: 0,
        certificatesEarned: 0
      },
      modules: [],
      progressMap: {},
      resources: [],
      adminSnapshot: null,
      recommendedAction: {
        type: 'explore',
        label: 'Explore your training portal',
        page: 'dashboard'
      }
    };
  }

  const dailyLearningItems = getDailyLearningForUser_(ss, user);
  const todayCard = pickDailyLearningCard_(dailyLearningItems);
  const updates = getUpdatesCached_(ss);
  const progressSummary = getProgressSummary_(ss, user.email);
  const quizSummary = getQuizSummary_(ss, user.email);
  const dashboardStats = getDashboardStatsLight_(ss, user.email, progressSummary, quizSummary);
  const modules = getModulesCached_(ss);
  const progressMap = getProgressMap_(ss, user.email);

  const payload = {
    dailyLearning: {
      items: dailyLearningItems,
      todayCard: todayCard,
      progress: getDailyLearningUserProgress_(ss, user.email),
      streak: getUserStreak_(ss, user.email),
      meta: getDailyLearningMeta_(user.email, todayCard),
      weeklyStatus: getWeeklyDailyStatus_(user.email),
      leaderboard: getDailyLearningLeaderboard_(ss, user.email, 5)
    },
    updates: updates,
    progressSummary: progressSummary,
    quizSummary: quizSummary,
    dashboardStats: dashboardStats,
    modules: modules,
    progressMap: progressMap,
    resources: [],
    adminSnapshot: access.isAdmin ? getAdminSnapshot_(user.email) : null
  };

  payload.recommendedAction = getRecommendedAction_({
    dailyLearning: payload.dailyLearning,
    progressSummary: payload.progressSummary,
    updates: payload.updates
  });

  return payload;
}
function getRecommendedAction_(bootstrap) {
  const dailyMeta = bootstrap && bootstrap.dailyLearning ? bootstrap.dailyLearning.meta : null;
  const progressSummary = bootstrap && bootstrap.progressSummary ? bootstrap.progressSummary : null;
  const updates = bootstrap && bootstrap.updates ? bootstrap.updates : [];

  if (dailyMeta && !dailyMeta.todayCompleted) {
    return {
      type: 'daily_learning',
      label: 'Complete today’s Daily Learning',
      page: 'dashboard'
    };
  }

  if (progressSummary && Number(progressSummary.progressPercent || 0) < 100) {
    return {
      type: 'continue_training',
      label: 'Continue your current training',
      page: 'training'
    };
  }

  if (updates.length) {
    return {
      type: 'updates',
      label: 'Check the latest updates',
      page: 'updates'
    };
  }

  return {
    type: 'explore',
    label: 'Explore your training portal',
    page: 'dashboard'
  };
}

function buildDailyLearningCompletionResult_(email, cardId) {
  var ss = getOrCreateSpreadsheet_();
  var card = getDailyLearningCardById_(cardId);
  var meta = getDailyLearningMeta_(email, card);
  var weekly = getWeeklyDailyStatus_(email);
  var leaderboard = getDailyLearningLeaderboard_(ss, email, 5);
  var streakInfo = getUserStreak_(ss, email);

  return {
    success: true,
    cardId: String(cardId || ''),
    xpEarned: Number(card && card.xp ? card.xp : 0),
    totalXp: Number(meta.totalXp || 0),
    todayXp: Number(meta.todayXp || 0),
    streak: Number(meta.streak || 0),
    bestStreak: Number(streakInfo.bestStreak || 0),
    todayCompleted: !!meta.todayCompleted,
    weeklyDailyStatus: weekly,
    leaderboard: leaderboard
  };
}
function getDailyLearningLeaderboard_(ss, currentEmail, limit) {
  const streakSh = ss.getSheetByName('UserStreaks');
  const usersSh = ss.getSheetByName('Users');
  const maxItems = Number(limit || 5);

  if (!streakSh || streakSh.getLastRow() < 2) return [];

  const userNameMap = {};
  if (usersSh && usersSh.getLastRow() > 1) {
    const userValues = usersSh.getDataRange().getValues();
    for (let i = 1; i < userValues.length; i++) {
      const email = String(userValues[i][0] || '').trim().toLowerCase();
      const name = String(userValues[i][1] || '').trim();
      if (email) {
        userNameMap[email] = name || email;
      }
    }
  }

  const values = streakSh.getDataRange().getValues();
  const targetCurrent = String(currentEmail || '').trim().toLowerCase();
  const rows = [];

  for (let i = 1; i < values.length; i++) {
    const email = String(values[i][0] || '').trim().toLowerCase();
    if (!email) continue;

    rows.push({
      email: email,
      name: userNameMap[email] || email,
      currentStreak: Number(values[i][1] || 0),
      bestStreak: Number(values[i][2] || 0),
      totalXp: Number(values[i][3] || 0),
      lastLearnedDate: values[i][4] ? formatDateSafe_(values[i][4]) : '',
      isCurrentUser: email === targetCurrent
    });
  }

  rows.sort(function(a, b) {
    if (b.totalXp !== a.totalXp) return b.totalXp - a.totalXp;
    if (b.currentStreak !== a.currentStreak) return b.currentStreak - a.currentStreak;
    if (b.bestStreak !== a.bestStreak) return b.bestStreak - a.bestStreak;
    return String(a.name || '').localeCompare(String(b.name || ''));
  });

  return rows.slice(0, maxItems).map(function(row, index) {
    return {
      rank: index + 1,
      email: row.email,
      name: row.name,
      currentStreak: row.currentStreak,
      bestStreak: row.bestStreak,
      totalXp: row.totalXp,
      lastLearnedDate: row.lastLearnedDate,
      isCurrentUser: row.isCurrentUser
    };
  });
}
function getHeaderMap_(headers) {
  var map = {};
  for (var i = 0; i < headers.length; i++) {
    map[String(headers[i] || '').trim()] = i;
  }
  return map;
}

function markDailyLearningComplete(cardId) {
  const ctx = requireEnabledPortalUser_();
  const email = String(ctx.user.email || '').trim().toLowerCase();

  if (!email) {
    throw new Error('Unable to determine current user email.');
  }

  const card = getDailyLearningCardById_(cardId);
  if (!card) {
    throw new Error('Daily Learning card not found.');
  }

  const progressSheet = getSheetByName_('DailyLearningProgress');
  if (!progressSheet) {
    throw new Error('DailyLearningProgress sheet not found.');
  }

  ensureDailyLearningProgressHeaders_(progressSheet);

  const values = progressSheet.getDataRange().getValues();
  const headers = values[0];
  const index = getHeaderMap_(headers);
  const timezone = Session.getScriptTimeZone();
  const now = new Date();
  const todayKey = Utilities.formatDate(now, timezone, 'yyyy-MM-dd');
  const normalizedEmail = String(email || '').trim().toLowerCase();
  const normalizedCardId = String(cardId || '').trim();

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rowEmail = String(row[index['Email']] || '').trim().toLowerCase();
    const rowCardId = String(row[index['Card ID']] || '').trim();
    const learnedAt = row[index['Learned At']] || '';
    const rowDateKey = learnedAt ? Utilities.formatDate(new Date(learnedAt), timezone, 'yyyy-MM-dd') : '';

    if (rowEmail === normalizedEmail && rowCardId === normalizedCardId && rowDateKey === todayKey) {
      return buildDailyLearningCompletionResult_(email, cardId);
    }
  }

  progressSheet.appendRow([
    email,
    card.type || '',
    card.title || '',
    card.content || '',
    card.tag || '',
    normalizedCardId,
    Number(card.xp || 0),
    now
  ]);

  updateUserStreakForToday_(email, Number(card.xp || 0));
  clearDailyLearningUserCache_(email);

  return buildDailyLearningCompletionResult_(email, cardId);
}
function getSheetByName_(name) {
  var ss = getOrCreateSpreadsheet_();
  return ss ? ss.getSheetByName(name) : null;
}

function getCurrentUserEmail_() {
  return String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
}
function ensureDailyLearningProgressHeaders_(sheet) {
  var requiredHeaders = [
    'Email',
    'Type',
    'Title',
    'Content',
    'Tag',
    'Card ID',
    'XP Earned',
    'Learned At'
  ];

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(requiredHeaders);
    return;
  }

  var existing = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), requiredHeaders.length)).getValues()[0];
  var needsReset = false;

  for (var i = 0; i < requiredHeaders.length; i++) {
    if (String(existing[i] || '').trim() !== requiredHeaders[i]) {
      needsReset = true;
      break;
    }
  }

  if (needsReset) {
    sheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
  }
}

function getDailyLearningCardById_(cardId) {
  var ss = getOrCreateSpreadsheet_();
  var items = getDailyLearningCached_(ss);
  var target = String(cardId || '').trim();

  for (var i = 0; i < items.length; i++) {
    if (String(items[i].cardId || '').trim() === target) {
      return items[i];
    }
  }

  return null;
}
function clearDailyLearningUserCache_(email) {
  return email;
}
function updateUserStreakForToday_(email, xpEarned) {
  var ss = getOrCreateSpreadsheet_();
  updateUserStreak_(ss, email, Number(xpEarned || 0));
}

function getWeeklyDailyStatus_(email) {
  var progressSheet = getSheetByName_('DailyLearningProgress');
  var timezone = Session.getScriptTimeZone();
  var today = new Date();
  var statusMap = {};
  var result = [];

  if (progressSheet && progressSheet.getLastRow() > 1) {
    var values = progressSheet.getDataRange().getValues();
    var headers = values[0];
    var index = getHeaderMap_(headers);

    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      var rowEmail = String(row[index['Email']] || '').trim().toLowerCase();
      if (rowEmail !== String(email || '').trim().toLowerCase()) continue;

      var learnedAt = row[index['Learned At']] || row[index['Completed At']] || row[index['Date']] || '';
      if (!learnedAt) continue;

      var dateKey = Utilities.formatDate(new Date(learnedAt), timezone, 'yyyy-MM-dd');
      statusMap[dateKey] = true;
    }
  }

  for (var d = 6; d >= 0; d--) {
    var dateObj = new Date(today);
    dateObj.setDate(today.getDate() - d);

    var dateKeyOut = Utilities.formatDate(dateObj, timezone, 'yyyy-MM-dd');
    var label = Utilities.formatDate(dateObj, timezone, 'E').charAt(0);

    result.push({
      label: label,
      completed: !!statusMap[dateKeyOut],
      dateKey: dateKeyOut
    });
  }

  return result;
}
function getDailyLearningMeta_(email, card) {
  var progressSheet = getSheetByName_('DailyLearningProgress');
  var streakSheet = getSheetByName_('UserStreaks');

  var todayKey = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var todayCompleted = false;
  var todayXp = 0;
  var totalXp = 0;
  var streak = 0;

  if (progressSheet && progressSheet.getLastRow() > 1) {
    var progressValues = progressSheet.getDataRange().getValues();
    var progressHeaders = progressValues[0];
    var pIndex = getHeaderMap_(progressHeaders);

    for (var i = 1; i < progressValues.length; i++) {
      var row = progressValues[i];
      var rowEmail = String(row[pIndex['Email']] || '').trim().toLowerCase();
      if (rowEmail !== String(email || '').trim().toLowerCase()) continue;

      var learnedAt = row[pIndex['Learned At']] || row[pIndex['Completed At']] || row[pIndex['Date']] || '';
      var xpValue = Number(row[pIndex['XP Earned']] || 0);
      var rowCardId = String(row[pIndex['Card ID']] || '').trim();

      totalXp += xpValue;

      if (learnedAt) {
        var rowDateKey = Utilities.formatDate(new Date(learnedAt), Session.getScriptTimeZone(), 'yyyy-MM-dd');
        if (rowDateKey === todayKey) {
          todayXp += xpValue;
          if (card && rowCardId === String(card.cardId || '').trim()) {
            todayCompleted = true;
          }
        }
      }
    }
  }

  if (streakSheet && streakSheet.getLastRow() > 1) {
    var streakValues = streakSheet.getDataRange().getValues();
    var streakHeaders = streakValues[0];
    var sIndex = getHeaderMap_(streakHeaders);

    for (var j = 1; j < streakValues.length; j++) {
      var streakRow = streakValues[j];
      var streakEmail = String(streakRow[sIndex['Email']] || '').trim().toLowerCase();
      if (streakEmail !== String(email || '').trim().toLowerCase()) continue;

      streak = Number(streakRow[sIndex['Current Streak']] || 0);
      break;
    }
  }

  return {
    streak: streak,
    todayCompleted: todayCompleted,
    totalXp: totalXp,
    todayXp: todayXp
  };
}

function getDailyLearningUserProgress_(ss, email) {
  const sh = ss.getSheetByName('DailyLearningProgress');
  if (!sh) return {
    seenCardIds: [],
    learnedCardIds: []
  };

  ensureDailyLearningProgressHeaders_(sh);

  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const index = getHeaderMap_(headers);
  const target = String(email || '').trim().toLowerCase();

  const seen = {};
  const learned = {};

  for (let i = 1; i < values.length; i++) {
    const rowEmail = String(values[i][index['Email']] || '').trim().toLowerCase();
    const cardId = String(values[i][index['Card ID']] || '').trim();
    const learnedAt = values[i][index['Learned At']];

    if (rowEmail !== target || !cardId) continue;

    seen[cardId] = true;
    if (learnedAt) learned[cardId] = true;
  }

  return {
    seenCardIds: Object.keys(seen),
    learnedCardIds: Object.keys(learned)
  };
}

function markDailyLearningSeen(cardId) {
  const ctx = requireEnabledPortalUser_();
  const email = String(ctx.user.email || '').trim();

  if (!email || email === 'Not available') throw new Error('User email is not available.');
  if (!cardId) throw new Error('Card ID is required.');

  const ss = ctx.ss;
  const sh = ss.getSheetByName('DailyLearningProgress');
  if (!sh) throw new Error('DailyLearningProgress sheet not found.');

  ensureDailyLearningProgressHeaders_(sh);

  return {
    progress: getDailyLearningUserProgress_(ss, email),
    streak: getUserStreak_(ss, email)
  };
}
function markDailyLearningLearned(cardId) {
  return markDailyLearningComplete(cardId);
}

function getUserStreak_(ss, email) {
  const sh = ss.getSheetByName('UserStreaks');
  if (!sh) {
    return {
      currentStreak: 0,
      bestStreak: 0,
      totalXp: 0,
      lastLearnedDate: ''
    };
  }

  const values = sh.getDataRange().getValues();
  const target = String(email || '').trim().toLowerCase();

  for (let i = 1; i < values.length; i++) {
    const rowEmail = String(values[i][0] || '').trim().toLowerCase();
    if (rowEmail !== target) continue;

    return {
      currentStreak: Number(values[i][1] || 0),
      bestStreak: Number(values[i][2] || 0),
      totalXp: Number(values[i][3] || 0),
      lastLearnedDate: values[i][4] ? formatDateSafe_(values[i][4]) : ''
    };
  }

  return {
    currentStreak: 0,
    bestStreak: 0,
    totalXp: 0,
    lastLearnedDate: ''
  };
}

function updateUserStreak_(ss, email, xpEarned) {
  const sh = ss.getSheetByName('UserStreaks');
  const values = sh.getDataRange().getValues();
  const target = String(email || '').trim().toLowerCase();

  const tz = Session.getScriptTimeZone() || 'Asia/Kolkata';
  const today = new Date();
  const todayKey = Utilities.formatDate(today, tz, 'yyyy-MM-dd');

  for (let i = 1; i < values.length; i++) {
    const rowEmail = String(values[i][0] || '').trim().toLowerCase();
    if (rowEmail !== target) continue;

    const currentStreak = Number(values[i][1] || 0);
    const bestStreak = Number(values[i][2] || 0);
    const totalXp = Number(values[i][3] || 0);
    const lastLearned = values[i][4];

    let newStreak = 1;

    if (lastLearned) {
      const lastDate = new Date(lastLearned);
      const diffDays = Math.floor((new Date(todayKey) - new Date(Utilities.formatDate(lastDate, tz, 'yyyy-MM-dd'))) / (1000 * 60 * 60 * 24));

      if (diffDays === 0) {
        newStreak = currentStreak;
      } else if (diffDays === 1) {
        newStreak = currentStreak + 1;
      } else {
        newStreak = 1;
      }
    }

    sh.getRange(i + 1, 2).setValue(newStreak);
    sh.getRange(i + 1, 3).setValue(Math.max(bestStreak, newStreak));
    sh.getRange(i + 1, 4).setValue(totalXp + Number(xpEarned || 0));
    sh.getRange(i + 1, 5).setValue(today);
    return;
  }

  sh.appendRow([email, 1, 1, Number(xpEarned || 0), today]);
}

function getDashboardStatsLight_(ss, email, progressSummary, quizSummary) {
  const updates = getUpdatesCached_(ss);
  const effectiveProgressSummary = progressSummary || getProgressSummary_(ss, email);
  const effectiveQuizSummary = quizSummary || getQuizSummary_(ss, email);
  const modules = getModulesCached_(ss);

  return {
    moduleCount: modules.length,
    resourceCount: 0,
    updateCount: updates.length,
    progressPercent: effectiveProgressSummary.progressPercent,
    completedLessons: effectiveProgressSummary.completedLessons,
    totalLessons: effectiveProgressSummary.totalLessons,
    currentWeek: effectiveProgressSummary.currentWeek,
    nextTask: effectiveProgressSummary.nextTask,
    quizzesTaken: effectiveQuizSummary.quizzesTaken,
    avgQuizScore: effectiveQuizSummary.avgQuizScore,
    certificatesEarned: effectiveQuizSummary.certificatesEarned
  };
}

function getSettingsCached_(ss) {
  const cache = CacheService.getScriptCache();
  const key = 'settings_v1';
  const cached = cache.get(key);

  if (cached) {
    return JSON.parse(cached);
  }

  const data = getSettings_(ss);
  cache.put(key, JSON.stringify(data), 300);
  return data;
}

function getUpdatesCached_(ss) {
  const cache = CacheService.getScriptCache();
  const key = 'updates_v1';
  const cached = cache.get(key);

  if (cached) {
    return JSON.parse(cached);
  }

  const data = getUpdates_(ss);
  cache.put(key, JSON.stringify(data), 300);
  return data;
}

function getDailyLearningCached_(ss) {
  const cache = CacheService.getScriptCache();
  const key = 'dailyLearning_v1';
  const cached = cache.get(key);

  if (cached) {
    return JSON.parse(cached);
  }

  const data = getDailyLearning_(ss);
  cache.put(key, JSON.stringify(data), 300);
  return data;
}

function getModulesCached_(ss) {
  const cache = CacheService.getScriptCache();
  const key = 'modules_v1';
  const cached = cache.get(key);

  if (cached) {
    return JSON.parse(cached);
  }

  const data = getModules_(ss);
  cache.put(key, JSON.stringify(data), 300);
  return data;
}

function clearPortalCaches() {
  const cache = CacheService.getScriptCache();
  [
    'settings_v1',
    'updates_v1',
    'dailyLearning_v1',
    'modules_v1'
  ].forEach(function(key) {
    cache.remove(key);
  });

  return 'Portal caches cleared.';
}
function getCompanyPageData() {
  const ctx = requireEnabledPortalUser_();
  const ss = ctx.ss;

  return {
    company: getCompanyBlock_(),
    faq: getFaq_(ss)
  };
}

function getTrainingPageData() {
  const ctx = requireEnabledPortalUser_();
  const ss = ctx.ss;

  return {
    modules: getModulesCached_(ss),
    progressMap: getProgressMap_(ss, ctx.user.email)
  };
}

function getUpdatesPageData() {
  const ctx = requireEnabledPortalUser_();
  const ss = ctx.ss;

  return {
    updates: getUpdatesCached_(ss)
  };
}
function getAdminPageData() {
  const user = getUserInfo_();
  const access = evaluateUserAccess_(user);
  if (!access.isAdmin) throw new Error('Only admins can access admin data.');

  return {
    adminSnapshot: getAdminSnapshot_(user.email)
  };
}
function getDailyLearning_(ss) {
  const sh = ss.getSheetByName('DailyLearning');
  if (!sh) return [];

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];

  const items = [];

  for (let i = 1; i < values.length; i++) {
    const enabled = String(values[i][0] || '').trim().toUpperCase();
    if (enabled !== 'Y') continue;

    items.push({
      type: String(values[i][1] || '').trim() || 'Tip',
      title: String(values[i][2] || '').trim(),
      content: String(values[i][3] || '').trim(),
      subtext: String(values[i][4] || '').trim(),
      tag: String(values[i][5] || '').trim(),
      optionA: String(values[i][6] || '').trim(),
      optionB: String(values[i][7] || '').trim(),
      optionC: String(values[i][8] || '').trim(),
      optionD: String(values[i][9] || '').trim(),
      correctOption: String(values[i][10] || '').trim().toUpperCase(),
      displayOrder: Number(values[i][11] || 9999),
      cardId: String(values[i][12] || '').trim(),
      audience: String(values[i][13] || 'All').trim(),
      xp: Number(values[i][14] || 5)
    });
  }

  items.sort(function(a, b) {
    return a.displayOrder - b.displayOrder;
  });

  return items;
}
function getDailyLearningForUser_(ss, user) {
  const allItems = getDailyLearningCached_(ss);
  const role = String((evaluateUserAccess_(user) || {}).role || '').trim().toLowerCase();

  return allItems.filter(function(item) {
    const audience = String(item.audience || 'All').trim().toLowerCase();

    if (audience === 'all') return true;
    if (audience === 'recruiter') {
      return role === 'trainee' || role === 'trainer' || role === 'admin' || role === 'super admin';
    }
    if (audience === 'hr') {
      return role === 'hr';
    }

    return true;
  });
}

function pickDailyLearningCard_(items) {
  if (!items || !items.length) {
    return null;
  }

  const tz = Session.getScriptTimeZone() || 'Asia/Kolkata';
  const todayKey = Utilities.formatDate(new Date(), tz, 'yyyyMMdd');
  const dayNumber = Number(todayKey.replace(/[^\d]/g, '')) || 0;
  const index = dayNumber % items.length;

  return items[index];
}

function seedDailyLearningSheet_(ss) {
  const sh = ss.getSheetByName('DailyLearning');
  if (sh.getLastRow() > 1) return;

  const rows = [
    ['Y', 'Vocabulary', 'Submission', 'A candidate profile sent to the client for review.', 'Common staffing term used in recruiter workflow.', 'Staffing Basics', '', '', '', '', '', 1],
    ['Y', 'Vocabulary', 'Screening', 'The process of evaluating whether a candidate fits a role before submission.', 'A key first-level recruiter responsibility.', 'Recruiter Basics', '', '', '', '', '', 2],
    ['Y', 'Vocabulary', 'C2H', 'C2H stands for Contract to Hire.', 'Important hiring model used in staffing.', 'Hiring Models', '', '', '', '', '', 3],
    ['Y', 'Tip', 'Daily Recruiter Tip', 'Always confirm notice period, work authorization, preferred location, and shift flexibility before submission.', 'This reduces avoidable rejections later.', 'Recruiter Tip', '', '', '', '', '', 4],
    ['Y', 'Quote', 'Daily Motivation', 'Accuracy builds trust faster than speed without quality.', 'Strong recruiters are known for reliable execution.', 'Motivation', '', '', '', '', '', 5],
    ['Y', 'Quiz', 'Mini Quiz', 'What does C2H stand for?', 'Test your staffing vocabulary.', 'Quiz', 'Contract to Hire', 'Candidate to Hire', 'Client to Hire', 'Contract to Hold', 'A', 6]
  ];

  sh.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}
function refreshPortalData() {
  return getPortalBootstrap();
}

/**
 * USER / ACCESS
 */
function getUserInfo_() {
  let email = '';
  let name = 'Trainee';
  let domain = '';

  try {
    email = Session.getActiveUser().getEmail() || '';
  } catch (err) {}

  if (email) {
    domain = email.split('@')[1] || '';

    // ✅ Get real name from Users sheet
    const ss = getOrCreateSpreadsheet_();
    const sh = ss.getSheetByName('Users');
    const values = sh.getDataRange().getValues();
    const target = String(email).trim().toLowerCase();

    for (let i = 1; i < values.length; i++) {
      const rowEmail = String(values[i][0] || '').trim().toLowerCase();
      const rowName = String(values[i][1] || '').trim();

      if (rowEmail === target && rowName) {
        name = rowName;
        break;
      }
    }

    // fallback (ONLY if name not found)
    if (!name || name === 'Trainee') {
      name = formatNameFromEmail_(email);
    }
  }

  return {
    email: email || 'Not available',
    name: name,
    domain: domain
  };
}
function evaluateUserAccess_(user) {
  const ss = getOrCreateSpreadsheet_();
  const email = String((user && user.email) || '').trim().toLowerCase();
  const domain = String((user && user.domain) || '').trim().toLowerCase();
  const allowedDomain = String(CONFIG.ALLOWED_DOMAIN || '').trim().toLowerCase();

  const domainAllowed = !CONFIG.ALLOW_DOMAIN_ONLY || domain === allowedDomain;
  const record = getUserRecordByEmail_(ss, email);
  const enabled = !!(record && String(record.enabled || '').trim().toUpperCase() === 'Y');
  const role = enabled ? (String(record.role || 'Trainee').trim() || 'Trainee') : '';

  const normalizedRole = role.toLowerCase();

  const adminRoles = {
    'admin': true,
    'director': true,
    'hr': true,
    'trainer': true,
    'qa': true
  };

  return {
    email: email,
    domainAllowed: domainAllowed,
    userFound: !!record,
    enabled: enabled,
    allowed: domainAllowed && !!record && enabled,
    role: role,
    isAdmin: !!adminRoles[normalizedRole],
    isTrainer: false,
    message: 'Access denied.'
  };
}
function getUserRole_(ss, email) {
  const record = getUserRecordByEmail_(ss, email);
  if (!record) return '';

  const enabled = String(record.enabled || '').trim().toUpperCase();
  if (enabled !== 'Y') return '';

  return String(record.role || 'Trainee').trim() || 'Trainee';
}
/**
 * DASHBOARD
 */
function getDashboardStats_(ss, email) {
  const modules = getModules_(ss);
  const resources = getDriveFiles_();
  const updates = getUpdates_(ss);
  const progressSummary = getProgressSummary_(ss, email);
  const quizSummary = getQuizSummary_(ss, email);

  return {
    moduleCount: modules.length,
    resourceCount: resources.length,
    updateCount: updates.length,
    progressPercent: progressSummary.progressPercent,
    completedLessons: progressSummary.completedLessons,
    totalLessons: progressSummary.totalLessons,
    currentWeek: progressSummary.currentWeek,
    nextTask: progressSummary.nextTask,
    quizzesTaken: quizSummary.quizzesTaken,
    avgQuizScore: quizSummary.avgQuizScore,
    certificatesEarned: quizSummary.certificatesEarned
  };
}

function getCompanyBlock_() {
  return {
    title: 'Know Your Company',
    about: `First Connect Health (FCH) is a Joint Commission (TJC) certified healthcare staffing agency, providing permanent, travel, and per diem staffing solutions across the U.S. Based in Newark, USA, with a significant office in Noida, India, the company specializes in nursing, therapy, and technical roles, with employee reviews indicating a positive work culture (4.6/5 stars on Glassdoor) and strong supportive leadership.

Core Services & Specialties

• Staffing Solutions: Offers travel nursing, travel therapy, per diem staffing, and permanent placement.
• Specialties: Covers various positions, including registered nurses, therapy, radiation oncology, radiology professionals, and laboratory specialists.
• Facilities Served: Places professionals in acute care, nursing homes, assisted living, psychiatric, and correctional facilities.`,

values: [
  {
    title: 'Accuracy over speed',
    desc: 'Take time to get the details right. Correct work builds trust and prevents avoidable mistakes.'
  },
  {
    title: 'No assumptions',
    desc: 'Always verify information before acting. Never move forward based on guesswork or incomplete understanding.'
  },
  {
    title: 'Verify every important detail',
    desc: 'Check requirements, credentials, dates, and submissions carefully before finalizing any step.'
  },
  {
    title: 'Quality over quantity',
    desc: 'Focus on meaningful, well-qualified output instead of rushing through high volumes with low accuracy.'
  },
  {
    title: 'Continuous learning',
    desc: 'Keep improving through feedback, updates, practice, and a better understanding of the process.'
  },
  {
    title: 'Professional communication',
    desc: 'Communicate clearly, respectfully, and promptly with candidates, teammates, and stakeholders.'
  }
],
    contacts: [
      {
        role: 'HR Support',
        name: 'HR Team',
        email: 'hr@firstconnecthealth.com',
        note: 'Company Policy, Attendance, leave, HR documents, onboarding support'
      },
      {
        role: 'Training Support',
        name: 'Training Team',
        email: 'training@firstconnecthealth.com',
        note: 'Module questions, quiz issues, training support'
      },
      {
        role: 'Manager Support',
        name: 'Reporting Manager / Team Lead',
        email: 'manager@firstconnecthealth.com',
        note: 'Performance expectations, practical guidance, operational help'
      }
    ]
  };
}
/**
 * SETTINGS
 */
function getSettings_(ss) {
  return {
    certificateFolderName: CONFIG.DEFAULT_CERTIFICATE_FOLDER_NAME,
    chatbotMessage: 'AI chatbot can be connected later to Drive files, FAQs, and sheet-based training knowledge.',
    welcomeBanner: 'Welcome to the First Connect Health Training Portal'
  };
}

/**
 * MODULES
 * Sheet Columns:
 * Enabled | Module ID | Module Title | Module Subtitle | Lesson Order | Lesson Name | Lesson Type | Lesson Link | Video Embed URL
 */
function getModules_(ss) {
  const sh = ss.getSheetByName('Modules');
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];

  const grouped = {};

  for (let i = 1; i < values.length; i++) {
    const enabled = String(values[i][0] || '').trim().toUpperCase();
    if (enabled !== 'Y') continue;

    const moduleId = String(values[i][1] || '').trim();
    const moduleTitle = String(values[i][2] || '').trim();
    const moduleStage = String(values[i][3] || '').trim();
    const lessonOrder = Number(values[i][4] || 0);
    const lessonName = String(values[i][5] || '').trim();
    const lessonType = String(values[i][6] || '').trim();
    const lessonLink = String(values[i][7] || '').trim();
    const videoEmbedUrl = String(values[i][8] || '').trim();
    const trainingVisualFileId = String(values[i][9] || '').trim();

    if (!grouped[moduleId]) {
      grouped[moduleId] = {
        id: moduleId,
        title: moduleTitle,
        subtitle: moduleStage,
        visualFileId: '',
        lessons: []
      };
    }

    grouped[moduleId].lessons.push({
      order: lessonOrder,
      name: lessonName,
      type: lessonType || 'Lesson',
      link: lessonLink,
      videoEmbedUrl: videoEmbedUrl,
      trainingVisualFileId: trainingVisualFileId
    });
  }

  const modules = Object.keys(grouped).map(function(key) {
    grouped[key].lessons.sort(function(a, b) {
      return a.order - b.order;
    });

    const firstVisualLesson = grouped[key].lessons.find(function(lesson) {
      return String(lesson.trainingVisualFileId || '').trim();
    });

    grouped[key].visualFileId = firstVisualLesson
      ? String(firstVisualLesson.trainingVisualFileId || '').trim()
      : '';

    return grouped[key];
  });

  modules.sort(function(a, b) {
    return getNumberFromModuleId_(a.id) - getNumberFromModuleId_(b.id);
  });

  return modules;
}
function getDefaultTrainingVisualForModule(module) {
  if (!module || !Array.isArray(module.lessons)) return '';

  const lessonWithVisual = module.lessons.find(function(lesson) {
    return String(lesson.trainingVisualFileId || '').trim();
  });

  return lessonWithVisual ? String(lessonWithVisual.trainingVisualFileId || '').trim() : '';
}

function setActiveTrainingVisual(moduleId, visualFileId, lessonName) {
  const targetModuleId = String(moduleId || '').trim();
  const targetVisualFileId = String(visualFileId || '').trim();

  if (!targetModuleId) return;

  state.trainingActiveVisuals = state.trainingActiveVisuals || {};

  if (!targetVisualFileId) {
    return;
  }

  state.trainingActiveVisuals[targetModuleId] = {
    visualFileId: targetVisualFileId,
    lessonName: String(lessonName || '').trim()
  };

  const img = document.getElementById('trainingVisualImg_' + targetModuleId);
  const empty = document.getElementById('trainingVisualEmpty_' + targetModuleId);
  const title = document.getElementById('trainingVisualTitle_' + targetModuleId);

  if (img) {
    img.src = getDriveImageUrlFromFileId(targetVisualFileId, 1600);
    img.style.display = 'block';
  }

  if (empty) {
    empty.style.display = 'none';
  }

  if (title) {
    title.textContent = (lessonName ? lessonName + ' Visual' : 'Training Visual');
  }
}
function getSmartSuggestions_(question, match) {
  const suggestions = [];

  if (match && match.mainQuestion) {
    suggestions.push('Explain this step by step');
    suggestions.push('Show related module');
    suggestions.push('Give real example');
  }

  if (question.includes('module')) {
    suggestions.push('Open this module');
    suggestions.push('Show assessment');
  }

  if (question.includes('leave') || question.includes('attendance')) {
    suggestions.push('Who is HR contact?');
    suggestions.push('Where to raise request?');
  }

  return suggestions.slice(0, 3);
}
function getNumberFromModuleId_(value) {
  const n = Number(String(value || '').replace(/[^\d]/g, ''));
  return isNaN(n) ? 9999 : n;
}

/**
 * FAQ / ANNOUNCEMENTS
 */
function getFaq_(ss) {
  return [];
}
function getUpdates_(ss) {
  const sh = ss.getSheetByName('Updates');
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = values[0];
  const index = getHeaderMap_(headers);
  const items = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const enabled = String(row[index['Enabled']] || '').trim().toUpperCase();
    const updateCategory = String(row[index['Update Category']] || '').trim();

    if (enabled !== 'Y') continue;
    if (updateCategory === 'AdminLog') continue;

    items.push({
      rowNumber: i + 1,
      title: String(row[index['Title']] || '').trim(),
      message: String(row[index['Message']] || '').trim(),
      type: String(row[index['Type']] || 'info').trim() || 'info',
      updatedAt: row[index['Updated At']] ? formatDateSafe_(row[index['Updated At']]) : '',
      updatedBy: String(row[index['Updated By']] || '').trim(),
      updateTopic: String(row[index['Update Topic']] || '').trim(),
      updateCategory: updateCategory
    });
  }

  return items;
}
/**
 * DRIVE RESOURCES
 */
function getDriveFiles_() {
  try {
    const folders = DriveApp.getFoldersByName(CONFIG.DRIVE_FOLDER_NAME);
    if (!folders.hasNext()) return [];

    const folder = folders.next();
    const files = folder.getFiles();
    const result = [];

    while (files.hasNext()) {
      const file = files.next();
      result.push({
        name: file.getName(),
        url: file.getUrl(),
        type: getFileTypeLabel_(file.getMimeType()),
        updatedAt: Utilities.formatDate(file.getLastUpdated(), Session.getScriptTimeZone(), 'dd MMM yyyy')
      });
    }

    result.sort(function(a, b) {
      return a.name.localeCompare(b.name);
    });

    return result;
  } catch (err) {
    return [{
      name: 'Could not load Drive files',
      url: '',
      type: 'Error',
      updatedAt: err.message
    }];
  }
}

function getFileTypeLabel_(mimeType) {
  const map = {
    'application/pdf': 'PDF',
    'application/vnd.google-apps.document': 'Google Doc',
    'application/vnd.google-apps.spreadsheet': 'Google Sheet',
    'application/vnd.google-apps.presentation': 'Google Slides',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document': 'Word',
    'application/msword': 'Word',
    'video/mp4': 'Video',
    'image/png': 'Image',
    'image/jpeg': 'Image'
  };
  return map[mimeType] || 'File';
}

/**
 * PROGRESS
 * Sheet Columns:
 * Email | Module ID | Lesson Name | Status | Updated At
 */
function getProgressMap_(ss, email) {
  const sh = ss.getSheetByName('Progress');
  const values = sh.getDataRange().getValues();
  const target = String(email || '').trim().toLowerCase();
  const map = {};

  for (let i = 1; i < values.length; i++) {
    const rowEmail = String(values[i][0] || '').trim().toLowerCase();
    if (rowEmail !== target) continue;

    const moduleId = String(values[i][1] || '').trim();
    const lessonName = String(values[i][2] || '').trim();
    const status = String(values[i][3] || 'Not Started').trim() || 'Not Started';

    map[buildProgressKey_(moduleId, lessonName)] = status;
  }

  return map;
}
function getSmartSuggestions_(question) {
  const q = String(question || '').toLowerCase();

  if (q.includes('leave')) {
    return ['Who is HR contact?', 'How to apply leave step by step?'];
  }

  if (q.includes('attendance')) {
    return ['Who handles attendance issues?', 'How to fix attendance?'];
  }

  if (q.includes('module')) {
    return ['Open this module', 'Show assessment'];
  }

  return ['Ask another question', 'Show training modules'];
}
function getProgressSummary_(ss, email) {
  const modules = getModules_(ss);
  const progressMap = getProgressMap_(ss, email);

  let totalLessons = 0;
  let completedLessons = 0;
  let inProgressLessons = 0;
  let currentWeek = 'Not Started';
  let nextTask = 'Open your first module';

  for (let i = 0; i < modules.length; i++) {
    const module = modules[i];
    for (let j = 0; j < module.lessons.length; j++) {
      totalLessons++;
      const lesson = module.lessons[j];
      const key = buildProgressKey_(module.id, lesson.name);
      const status = progressMap[key] || 'Not Started';

      if (status === 'Completed') completedLessons++;
      if (status === 'In Progress') inProgressLessons++;
    }
  }

  const progressPercent = totalLessons > 0
    ? Math.round((completedLessons / totalLessons) * 100)
    : 0;

  for (let i = 0; i < modules.length; i++) {
    const module = modules[i];
    let found = false;

    for (let j = 0; j < module.lessons.length; j++) {
      const lesson = module.lessons[j];
      const status = progressMap[buildProgressKey_(module.id, lesson.name)] || 'Not Started';

      if (status !== 'Completed') {
        currentWeek = module.title;
        nextTask = lesson.name;
        found = true;
        break;
      }
    }

    if (found) break;
  }

  if (completedLessons === totalLessons && totalLessons > 0) {
    currentWeek = 'All core lessons completed';
    nextTask = 'Complete quizzes and download certificates';
  }

  return {
    progressPercent: progressPercent,
    completedLessons: completedLessons,
    inProgressLessons: inProgressLessons,
    totalLessons: totalLessons,
    currentWeek: currentWeek,
    nextTask: nextTask
  };
}

function buildProgressKey_(moduleId, lessonName) {
  return moduleId + '||' + lessonName;
}

function updateLessonProgress(payload) {
  if (!payload) throw new Error('Missing payload.');

  const ctx = requireEnabledPortalUser_();
  const email = String(ctx.user.email || '').trim();
  const moduleId = String(payload.moduleId || '').trim();
  const lessonName = String(payload.lessonName || '').trim();
  const status = String(payload.status || '').trim();

  if (!email || email === 'Not available') throw new Error('User email is not available.');
  if (!moduleId) throw new Error('Module ID is required.');
  if (!lessonName) throw new Error('Lesson name is required.');

  const validStatuses = ['Not Started', 'In Progress', 'Completed'];
  if (validStatuses.indexOf(status) === -1) throw new Error('Invalid status.');

  const ss = ctx.ss;
  const sh = ss.getSheetByName('Progress');
  const values = sh.getDataRange().getValues();
  let rowToUpdate = -1;

  for (let i = 1; i < values.length; i++) {
    const rowEmail = String(values[i][0] || '').trim().toLowerCase();
    const rowModuleId = String(values[i][1] || '').trim();
    const rowLessonName = String(values[i][2] || '').trim();

    if (
      rowEmail === email.toLowerCase() &&
      rowModuleId === moduleId &&
      rowLessonName === lessonName
    ) {
      rowToUpdate = i + 1;
      break;
    }
  }

  const now = new Date();

  if (rowToUpdate > -1) {
    sh.getRange(rowToUpdate, 4).setValue(status);
    sh.getRange(rowToUpdate, 5).setValue(now);
  } else {
    sh.appendRow([email, moduleId, lessonName, status, now]);
  }

  return {
    progressSummary: getProgressSummary_(ss, email),
    progressMap: getProgressMap_(ss, email),
    dashboardStats: getDashboardStats_(ss, email)
  };
}

/**
 * QUIZZES
 * Sheet Columns:
 * Enabled | Module ID | Question ID | Question | Option A | Option B | Option C | Option D | Correct Option | Explanation
 */
function getQuizQuestionsByModule(moduleId) {
  requireEnabledPortalUser_();

  const ss = getOrCreateSpreadsheet_();
  const sh = ss.getSheetByName('Quizzes');
  const data = sh.getDataRange().getValues();
  const targetModuleId = String(moduleId || '').trim();
  // Each module can store 40+ questions in Quizzes.
  // For every attempt, only 25 random enabled questions are served.
  const QUIZ_PICK_COUNT = 25;

  if (!targetModuleId) return [];

  const rows = data.slice(1).filter(function(r) {
    const enabled = String(r[0] || '').trim().toUpperCase();
    const rowModuleId = String(r[1] || '').trim();
    return enabled === 'Y' && rowModuleId === targetModuleId;
  });

  for (let i = rows.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    const temp = rows[i];
    rows[i] = rows[j];
    rows[j] = temp;
  }

  const selectedRows = rows.slice(0, QUIZ_PICK_COUNT);
  if (!selectedRows.length) return [];
  return selectedRows.map(function(r) {
    return {
      moduleId: String(r[1] || '').trim(),
      questionId: String(r[2] || '').trim(),
      question: String(r[3] || '').trim(),
      options: [
        String(r[4] || '').trim(),
        String(r[5] || '').trim(),
        String(r[6] || '').trim(),
        String(r[7] || '').trim()
      ],
      explanation: String(r[9] || '').trim()
    };
  });
}
function submitQuiz(payload) {
  if (!payload) throw new Error('Missing payload.');

  const ctx = requireEnabledPortalUser_();
  const email = String(ctx.user.email || '').trim();

  if (!email || email === 'Not available') throw new Error('User email is not available.');

  const moduleId = String(payload.moduleId || '').trim();
  const answers = payload.answers || {};
  const questionIds = Array.isArray(payload.questionIds) ? payload.questionIds : [];

  if (!moduleId) throw new Error('Module ID is required.');
  if (!questionIds.length) throw new Error('No quiz questions were submitted.');

  const allowedQuestionIds = {};
  questionIds.forEach(function(id) {
    const qid = String(id || '').trim();
    if (qid) allowedQuestionIds[qid] = true;
  });

  const ss = ctx.ss;
  const sh = ss.getSheetByName('Quizzes');
  const values = sh.getDataRange().getValues();

  let total = 0;
  let correct = 0;
  const details = [];

  for (let i = 1; i < values.length; i++) {
    const enabled = String(values[i][0] || '').trim().toUpperCase();
    const rowModuleId = String(values[i][1] || '').trim();
    const questionId = String(values[i][2] || '').trim();

    if (enabled !== 'Y' || rowModuleId !== moduleId || !allowedQuestionIds[questionId]) {
      continue;
    }

    total++;
    const question = String(values[i][3] || '').trim();
    const correctOption = String(values[i][8] || '').trim().toUpperCase();
    const explanation = String(values[i][9] || '').trim();
    const selected = String(answers[questionId] || '').trim().toUpperCase();

    const isCorrect = selected && selected === correctOption;
    if (isCorrect) correct++;

    details.push({
      questionId: questionId,
      question: question,
      selected: selected,
      correctOption: correctOption,
      isCorrect: isCorrect,
      explanation: explanation
    });
  }
details.sort(function(a, b) {
  return questionIds.indexOf(a.questionId) - questionIds.indexOf(b.questionId);
});
  if (total === 0) throw new Error('No quiz found for this module.');
  if (total !== questionIds.length) throw new Error('Submitted quiz set does not match the loaded questions.');

  const scorePercent = Math.round((correct / total) * 100);
  const passed = scorePercent >= CONFIG.PASS_PERCENT;

  const resultsSh = ss.getSheetByName('QuizResults');
  resultsSh.appendRow([
    email,
    moduleId,
    scorePercent,
    passed ? 'Passed' : 'Failed',
    JSON.stringify({
      questionIds: questionIds,
      answers: answers
    }),
    new Date()
  ]);

  let certificateInfo = null;
  if (passed) {
    certificateInfo = generateCertificateForModule_(ss, email, moduleId);
  }

  return {
    moduleId: moduleId,
    totalQuestions: total,
    correctAnswers: correct,
    scorePercent: scorePercent,
    passed: passed,
    passPercent: CONFIG.PASS_PERCENT,
    details: details,
    certificate: certificateInfo,
    quizSummary: getQuizSummary_(ss, email)
  };
}

function getQuizSummary_(ss, email) {
  const sh = ss.getSheetByName('QuizResults');
  const values = sh.getDataRange().getValues();
  const target = String(email || '').trim().toLowerCase();
  let quizzesTaken = 0;
  let scoreTotal = 0;
  let passCount = 0;

  for (let i = 1; i < values.length; i++) {
    const rowEmail = String(values[i][0] || '').trim().toLowerCase();
    if (rowEmail !== target) continue;

    quizzesTaken++;
    const score = Number(values[i][2] || 0);
    if (!isNaN(score)) scoreTotal += score;

    if (String(values[i][3] || '').trim() === 'Passed') passCount++;
  }

  const certificates = getCertificatesForUser_(ss, email);

  return {
    quizzesTaken: quizzesTaken,
    avgQuizScore: quizzesTaken > 0 ? Math.round(scoreTotal / quizzesTaken) : 0,
    passCount: passCount,
    certificatesEarned: certificates.length,
    certificates: certificates
  };
}

/**
 * CERTIFICATES
 * Certificates sheet:
 * Email | Module ID | Certificate Name | File URL | Issued At
 */
function generateCertificateForModule_(ss, email, moduleId) {
  const existing = getCertificateByModule_(ss, email, moduleId);
  if (existing) return existing;

  const folderName = getSettings_(ss).certificateFolderName || CONFIG.DEFAULT_CERTIFICATE_FOLDER_NAME;
  const folder = getOrCreateFolderByName_(folderName);

  let userName = '';
  const users = getUsers_(ss);
  const targetEmail = String(email || '').trim().toLowerCase();

  for (let i = 0; i < users.length; i++) {
    if (String(users[i].email || '').trim().toLowerCase() === targetEmail) {
      userName = String(users[i].name || '').trim();
      break;
    }
  }

  if (!userName) {
    userName = formatNameFromEmail_(email);
  }

  const module = findModuleById_(ss, moduleId);
  const moduleTitle = module ? String(module.title || '').trim() : String(moduleId || '').trim();

  const issuedAt = new Date();
  const certificateId = buildCertificateId_(moduleId, email, issuedAt);

  const html = buildCertificateHtml_(userName, moduleTitle, certificateId, issuedAt);
  const pdfBlob = HtmlService
    .createHtmlOutput(html)
    .getBlob()
    .getAs(MimeType.PDF);

  const pdfFileName = sanitizeFileName_(
    'Certificate - ' + userName + ' - ' + moduleTitle + '.pdf'
  );

  pdfBlob.setName(pdfFileName);

  const file = folder.createFile(pdfBlob);
  const certSh = ss.getSheetByName('Certificates');
  if (!certSh) throw new Error('Certificates sheet not found.');

  ensureCertificatesSheetHeaders_(certSh);

  certSh.appendRow([
    email,
    moduleId,
    file.getName(),
    file.getUrl(),
    issuedAt,
    file.getId(),
    certificateId
  ]);

  return {
    name: file.getName(),
    url: file.getUrl(),
    moduleId: String(moduleId || '').trim(),
    moduleTitle: moduleTitle,
    issuedAt: formatDateSafe_(issuedAt),
    fileId: file.getId(),
    certificateId: certificateId
  };
}
function buildCertificateHtml_(userName, moduleTitle, certificateId, issuedAt) {
  const safeUserName = escapeHtmlServer_(String(userName || '').trim());
  const safeModuleTitle = escapeHtmlServer_(String(moduleTitle || '').trim());
  const safeCertificateId = escapeHtmlServer_(String(certificateId || '').trim());

  const issuedDate = Utilities.formatDate(
    issuedAt || new Date(),
    Session.getScriptTimeZone(),
    'MMMM d, yyyy'
  );
  const safeIssuedDate = escapeHtmlServer_(issuedDate);

  const logoUrl = 'https://drive.google.com/uc?export=view&id=1qGHveKbBlPXHo44lx_KwngmIYhZHWjPp';

  return [
    '<!DOCTYPE html>',
    '<html>',
    '<head>',
    '<meta charset="utf-8">',
    '<title>Certificate of Completion</title>',
    '<style>',
    '@page { size: A4 landscape; margin: 0; }',
    'html, body { margin:0; padding:0; width:100%; height:100%; background:#ffffff; font-family:Arial, Helvetica, sans-serif; }',
    'body { -webkit-print-color-adjust: exact; print-color-adjust: exact; }',

    '.page {',
    '  position: relative;',
    '  width: 1123px;',
    '  height: 794px;',
    '  margin: 0 auto;',
    '  background: #ffffff;',
    '  overflow: hidden;',
    '  color: #1f1f1f;',
    '}',

    '.top-right-arc-1, .top-right-arc-2, .top-right-arc-3, .top-right-arc-4, .top-right-arc-5,',
    '.mid-left-wave-1, .mid-left-wave-2, .mid-left-wave-3, .mid-left-wave-4, .mid-left-wave-5 {',
    '  position: absolute;',
    '  border-top: 2px solid #b9d2ee;',
    '  border-radius: 50%;',
    '  opacity: 1;',
    '}',

    '.top-right-arc-1 { top: -115px; right: -40px; width: 420px; height: 220px; transform: rotate(12deg); }',
    '.top-right-arc-2 { top: -104px; right: -26px; width: 400px; height: 210px; transform: rotate(12deg); }',
    '.top-right-arc-3 { top: -93px; right: -12px; width: 380px; height: 200px; transform: rotate(12deg); }',
    '.top-right-arc-4 { top: -82px; right: 2px; width: 360px; height: 190px; transform: rotate(12deg); }',
    '.top-right-arc-5 { top: -71px; right: 16px; width: 340px; height: 180px; transform: rotate(12deg); }',

    '.mid-left-wave-1 { top: 150px; left: -120px; width: 420px; height: 150px; transform: rotate(12deg); }',
    '.mid-left-wave-2 { top: 162px; left: -102px; width: 405px; height: 145px; transform: rotate(12deg); }',
    '.mid-left-wave-3 { top: 174px; left: -84px; width: 390px; height: 140px; transform: rotate(12deg); }',
    '.mid-left-wave-4 { top: 186px; left: -66px; width: 375px; height: 135px; transform: rotate(12deg); }',
    '.mid-left-wave-5 { top: 198px; left: -48px; width: 360px; height: 130px; transform: rotate(12deg); }',

    '.bottom-dark {',
    '  position:absolute;',
    '  left:-40px;',
    '  bottom:-90px;',
    '  width:520px;',
    '  height:210px;',
    '  background:#0f2e72;',
    '  border-top-left-radius: 0;',
    '  border-top-right-radius: 260px 150px;',
    '}',

    '.bottom-light {',
    '  position:absolute;',
    '  left:300px;',
    '  bottom:-58px;',
    '  width:360px;',
    '  height:120px;',
    '  background:#6b93d8;',
    '  border-top-left-radius: 220px 110px;',
    '  border-top-right-radius: 220px 110px;',
    '}',

    '.bottom-mid {',
    '  position:absolute;',
    '  left:345px;',
    '  bottom:-76px;',
    '  width:420px;',
    '  height:145px;',
    '  background:#0f4da3;',
    '  border-top-left-radius: 260px 130px;',
    '  border-top-right-radius: 260px 130px;',
    '}',

    '.header-left {',
    '  position:absolute;',
    '  top:28px;',
    '  left:26px;',
    '  width:180px;',
    '  text-align:left;',
    '}',

    '.logo {',
    '  width:140px;',
    '  height:auto;',
    '  display:block;',
    '}',

    '.top-center {',
    '  position:absolute;',
    '  top:55px;',
    '  left:0;',
    '  right:0;',
    '  text-align:center;',
    '}',

    '.company-name {',
    '  margin:0;',
    '  color:#0c4a9a;',
    '  font-size:19px;',
    '  font-weight:800;',
    '  letter-spacing:0.5px;',
    '}',

    '.title {',
    '  margin:28px 0 0;',
    '  color:#0b57c2;',
    '  font-size:72px;',
    '  line-height:0.95;',
    '  font-weight:900;',
    '  letter-spacing:1px;',
    '}',

    '.subtitle {',
    '  margin:6px 0 0;',
    '  color:#0d4ea4;',
    '  font-size:28px;',
    '  font-weight:700;',
    '  letter-spacing:8px;',
    '}',

    '.badge-wrap {',
    '  position:absolute;',
    '  top:104px;',
    '  right:92px;',
    '  width:110px;',
    '  height:140px;',
    '}',

    '.badge-circle {',
    '  position:absolute;',
    '  top:0;',
    '  left:18px;',
    '  width:72px;',
    '  height:72px;',
    '  border-radius:50%;',
    '  background:#f3c93c;',
    '  box-shadow: inset 0 0 0 5px #fff3b5, inset 0 0 0 10px #f3c93c;',
    '}',

    '.badge-ribbon-left, .badge-ribbon-right {',
    '  position:absolute;',
    '  top:62px;',
    '  width:0;',
    '  height:0;',
    '  border-left:10px solid transparent;',
    '  border-right:10px solid transparent;',
    '  border-top:44px solid #0d4ea4;',
    '}',
    '.badge-ribbon-left { left:30px; transform: rotate(8deg); }',
    '.badge-ribbon-right { left:54px; transform: rotate(-8deg); }',

    '.body {',
    '  position:absolute;',
    '  left:0;',
    '  right:0;',
    '  top:210px;',
    '  text-align:center;',
    '  padding:0 90px;',
    '  box-sizing:border-box;',
    '}',

    '.intro {',
    '  margin:0;',
    '  font-size:23px;',
    '  color:#2d2d2d;',
    '  font-weight:400;',
    '}',

    '.name {',
    '  margin:24px 0 8px;',
    '  color:#0b57c2;',
    '  font-size:68px;',
    '  line-height:1;',
    '  font-weight:400;',
    '  font-family:"Brush Script MT","Lucida Handwriting","Segoe Script",cursive;',
    '}',

    '.name-line {',
    '  width:440px;',
    '  max-width:100%;',
    '  margin:0 auto 12px;',
    '  border-top:2px solid #8b8b8b;',
    '}',

    '.completion-line {',
    '  margin:0;',
    '  font-size:18px;',
    '  color:#262626;',
    '}',

    '.module-line {',
    '  margin:12px 0 0;',
    '  font-size:20px;',
    '  color:#262626;',
    '  font-weight:400;',
    '}',

    '.module-name {',
    '  color:#0b57c2;',
    '  font-weight:800;',
    '}',

    '.date-line {',
    '  margin:4px 0 0;',
    '  font-size:18px;',
    '  color:#262626;',
    '}',

    '.footer-id {',
    '  position:absolute;',
    '  left:42px;',
    '  bottom:26px;',
    '  font-size:12px;',
    '  color:#5f6d7d;',
    '}',

    '</style>',
    '</head>',
    '<body>',
    '<div class="page">',

    '<div class="top-right-arc-1"></div>',
    '<div class="top-right-arc-2"></div>',
    '<div class="top-right-arc-3"></div>',
    '<div class="top-right-arc-4"></div>',
    '<div class="top-right-arc-5"></div>',

    '<div class="mid-left-wave-1"></div>',
    '<div class="mid-left-wave-2"></div>',
    '<div class="mid-left-wave-3"></div>',
    '<div class="mid-left-wave-4"></div>',
    '<div class="mid-left-wave-5"></div>',

    '<div class="bottom-dark"></div>',
    '<div class="bottom-light"></div>',
    '<div class="bottom-mid"></div>',

    '<div class="header-left">',
    '<img class="logo" src="' + logoUrl + '" alt="First Connect Health Logo">',
    '</div>',

    '<div class="top-center">',
    '<div class="company-name">FIRST CONNECT HEALTH</div>',
    '<div class="title">CERTIFICATE</div>',
    '<div class="subtitle">OF COMPLETION</div>',
    '</div>',

    '<div class="badge-wrap">',
    '<div class="badge-circle"></div>',
    '<div class="badge-ribbon-left"></div>',
    '<div class="badge-ribbon-right"></div>',
    '</div>',

    '<div class="body">',
    '<p class="intro">This certificate is proudly presented to</p>',
    '<div class="name">' + safeUserName + '</div>',
    '<div class="name-line"></div>',
    '<p class="completion-line">for successfully completing the training module</p>',
    '<p class="module-line"><span class="module-name">' + safeModuleTitle + '</span></p>',
    '<p class="date-line">on ' + safeIssuedDate + '</p>',
    '</div>',

    '<div class="footer-id">Certificate ID: ' + safeCertificateId + '</div>',

    '</div>',
    '</body>',
    '</html>'
  ].join('');
}
function getCertificatesForUser() {
  const ctx = requireEnabledPortalUser_();
  const ss = ctx.ss;
  return getCertificatesForUser_(ss, ctx.user.email);
}
function getUserRecordByEmail_(ss, email) {
  const sh = ss.getSheetByName('Users');
  if (!sh) return null;

  const values = sh.getDataRange().getValues();
  const target = String(email || '').trim().toLowerCase();
  if (!target) return null;

  for (let i = 1; i < values.length; i++) {
    const rowEmail = String(values[i][0] || '').trim().toLowerCase();
    if (rowEmail !== target) continue;

    return {
      email: String(values[i][0] || '').trim(),
      name: String(values[i][1] || '').trim(),
      role: String(values[i][2] || '').trim(),
      enabled: String(values[i][3] || '').trim().toUpperCase()
    };
  }

  return null;
}
function requireEnabledPortalUser_() {
  const user = getUserInfo_();
  const ss = getOrCreateSpreadsheet_();
  const access = evaluateUserAccess_(user);

  if (!access.allowed) {
    throw new Error(access.message || 'Access denied.');
  }

  return {
    user: user,
    ss: ss,
    access: access
  };
}
function buildPortalOtpCacheKey_(email) {
  return 'portal_otp_code_' + String(email || '').trim().toLowerCase();
}
function buildPortalOtpVerifiedCacheKey_(email) {
  return 'portal_otp_verified_' + String(email || '').trim().toLowerCase();
}
function buildPortalOtpReverifyCacheKey_(email) {
  return 'portal_otp_reverify_' + String(email || '').trim().toLowerCase();
}
function buildPortalOtpAttemptsCacheKey_(email) {
  return 'portal_otp_attempts_' + String(email || '').trim().toLowerCase();
}
function buildPortalOtpCooldownCacheKey_(email) {
  return 'portal_otp_cooldown_' + String(email || '').trim().toLowerCase();
}
function isPortalOtpVerified_(email) {
  const cache = CacheService.getScriptCache();
  const value = cache.get(buildPortalOtpVerifiedCacheKey_(email));
  return value === 'Y';
}
function setPortalOtpVerified_(email, isVerified) {
  const cache = CacheService.getScriptCache();
  const key = buildPortalOtpVerifiedCacheKey_(email);

  if (isVerified) {
    cache.put(key, 'Y', 21600);
  } else {
    cache.remove(key);
  }
}
function clearPortalOtpVerified_(email) {
  const cache = CacheService.getScriptCache();
  cache.remove(buildPortalOtpVerifiedCacheKey_(email));
  cache.remove(buildPortalOtpReverifyCacheKey_(email));
}
function clearPortalOtpCode_(email) {
  const cache = CacheService.getScriptCache();
  cache.remove(buildPortalOtpCacheKey_(email));
  cache.remove(buildPortalOtpAttemptsCacheKey_(email));
}
function getAuthBootstrap_(ss, user) {
  const access = evaluateUserAccess_(user);
  const email = String((user && user.email) || '').trim().toLowerCase();
  const cache = CacheService.getScriptCache();
  const reverifyKey = buildPortalOtpReverifyCacheKey_(email);
  const reverifyRequired = cache.get(reverifyKey) === 'Y';
  const verified = access.allowed ? isPortalOtpVerified_(email) : false;

  if (!access.allowed) {
    return {
      requireOtp: false,
      mode: '',
      message: '',
      reverifyMinutes: 60
    };
  }

  if (verified && !reverifyRequired) {
    return {
      requireOtp: false,
      mode: '',
      message: '',
      reverifyMinutes: 60
    };
  }

  return {
    requireOtp: true,
    mode: reverifyRequired ? 'reauth' : 'login',
    message: reverifyRequired
      ? 'Your session expired. Please enter the OTP sent to your enabled company email.'
      : 'Enter the OTP sent to your enabled company email to continue.',
    reverifyMinutes: 60
  };
}

function verifyLoginOtp(otp) {
  const ctx = requireEnabledPortalUser_();
  const email = String(ctx.user.email || '').trim().toLowerCase();
  const inputOtp = String(otp || '').trim();
  const cache = CacheService.getScriptCache();
  const savedOtp = String(cache.get(buildPortalOtpCacheKey_(email)) || '').trim();
  const attemptsKey = buildPortalOtpAttemptsCacheKey_(email);
  const maxAttempts = Number(CONFIG.OTP_MAX_VERIFY_ATTEMPTS || 5);
  const attemptsUsed = Number(cache.get(attemptsKey) || 0);

  if (!/^\d{6}$/.test(inputOtp)) {
    return {
      success: false,
      message: 'Please enter a valid 6-digit OTP.'
    };
  }

  if (!savedOtp) {
    return {
      success: false,
      message: 'OTP expired or not found. Please send OTP again.'
    };
  }

  if (savedOtp !== inputOtp) {
    const updatedAttempts = attemptsUsed + 1;
    cache.put(attemptsKey, String(updatedAttempts), Number(CONFIG.OTP_EXPIRY_SECONDS || 300));

    if (updatedAttempts >= maxAttempts) {
      clearPortalOtpCode_(email);
      return {
        success: false,
        message: 'Too many invalid attempts. Please request a new OTP.'
      };
    }

    return {
      success: false,
      message: 'Invalid OTP. Attempts left: ' + Math.max(0, maxAttempts - updatedAttempts)
    };
  }

  setPortalOtpVerified_(email, true);
  clearPortalOtpCode_(email);
  cache.remove(buildPortalOtpReverifyCacheKey_(email));

  return {
    success: true,
    message: 'OTP verified successfully.'
  };
}
function markSessionReverificationRequired() {
  const ctx = requireEnabledPortalUser_();
  const email = String(ctx.user.email || '').trim().toLowerCase();
  const cache = CacheService.getScriptCache();

  clearPortalOtpVerified_(email);
  clearPortalOtpCode_(email);
  cache.put(buildPortalOtpReverifyCacheKey_(email), 'Y', 21600);

  return {
    success: true
  };
}
function updateUserActivity() {
  const ctx = requireEnabledPortalUser_();
  return {
    success: true,
    email: String(ctx.user.email || '').trim().toLowerCase(),
    updatedAt: new Date()
  };
}
function getCertificatesForUser_(ss, email) {
  const sh = ss.getSheetByName('Certificates');
  if (!sh) return [];

  ensureCertificatesSheetHeaders_(sh);

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = values[0];
  const index = getHeaderMap_(headers);
  const target = String(email || '').trim().toLowerCase();
  const items = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rowEmail = String(row[index['Email']] || '').trim().toLowerCase();
    if (rowEmail !== target) continue;

    const moduleId = String(row[index['Module ID']] || '').trim();

    items.push({
      moduleId: moduleId,
      moduleTitle: (findModuleById_(ss, moduleId) || {}).title || moduleId,
      name: String(row[index['Certificate Name']] || '').trim(),
      url: String(row[index['File URL']] || '').trim(),
      issuedAt: formatDateSafe_(row[index['Issued At']]),
      fileId: index['File ID'] !== undefined ? String(row[index['File ID']] || '').trim() : '',
      certificateId: index['Certificate ID'] !== undefined ? String(row[index['Certificate ID']] || '').trim() : ''
    });
  }

  return items;
}
function getCertificateByModule_(ss, email, moduleId) {
  const sh = ss.getSheetByName('Certificates');
  if (!sh) return null;

  ensureCertificatesSheetHeaders_(sh);

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return null;

  const headers = values[0];
  const index = getHeaderMap_(headers);
  const targetEmail = String(email || '').trim().toLowerCase();
  const targetModuleId = String(moduleId || '').trim();

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rowEmail = String(row[index['Email']] || '').trim().toLowerCase();
    const rowModuleId = String(row[index['Module ID']] || '').trim();

    if (rowEmail === targetEmail && rowModuleId === targetModuleId) {
      return {
        moduleId: rowModuleId,
        moduleTitle: (findModuleById_(ss, rowModuleId) || {}).title || rowModuleId,
        name: String(row[index['Certificate Name']] || '').trim(),
        url: String(row[index['File URL']] || '').trim(),
        issuedAt: formatDateSafe_(row[index['Issued At']]),
        fileId: index['File ID'] !== undefined ? String(row[index['File ID']] || '').trim() : '',
        certificateId: index['Certificate ID'] !== undefined ? String(row[index['Certificate ID']] || '').trim() : ''
      };
    }
  }

  return null;
}
function ensureCertificatesSheetHeaders_(sheet) {
  const requiredHeaders = [
    'Email',
    'Module ID',
    'Certificate Name',
    'File URL',
    'Issued At',
    'File ID',
    'Certificate ID'
  ];

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(requiredHeaders);
    return;
  }

  const currentLastColumn = Math.max(sheet.getLastColumn(), requiredHeaders.length);
  const existingHeaders = sheet.getRange(1, 1, 1, currentLastColumn).getValues()[0];
  const normalizedExisting = existingHeaders.map(function(value) {
    return String(value || '').trim();
  });

  let nextColumn = normalizedExisting.length + 1;

  for (let i = 0; i < requiredHeaders.length; i++) {
    if (normalizedExisting.indexOf(requiredHeaders[i]) === -1) {
      sheet.getRange(1, nextColumn).setValue(requiredHeaders[i]);
      normalizedExisting.push(requiredHeaders[i]);
      nextColumn++;
    }
  }
}
function buildCertificateId_(moduleId, email, issuedAt) {
  const baseEmail = String(email || '').trim().toLowerCase().split('@')[0] || 'user';
  const cleanEmail = baseEmail.replace(/[^a-z0-9]/gi, '').toUpperCase().slice(0, 6) || 'USER';
  const cleanModule = String(moduleId || '').trim().replace(/[^a-z0-9]/gi, '').toUpperCase() || 'MODULE';
  const datePart = Utilities.formatDate(
    issuedAt || new Date(),
    Session.getScriptTimeZone(),
    'yyyyMMdd'
  );

  return 'FCH-' + cleanModule + '-' + cleanEmail + '-' + datePart;
}
function sanitizeFileName_(value) {
  return String(value || '')
    .replace(/[\\\/:\*\?"<>\|#%&\{\}\$!'@`=+]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}
/**
 * ADMIN
 */
function getAdminSnapshot() {
  const user = getUserInfo_();
  const access = evaluateUserAccess_(user);
  if (!access.isAdmin) throw new Error('Only admins can access admin data.');

  return getAdminSnapshot_(user.email);
}

function getAdminSnapshot_(viewerEmail) {
  const ss = getOrCreateSpreadsheet_();
  const modules = getModules_(ss);
  const updates = getUpdates_(ss);
  const faq = getFaq_(ss);
  const users = getUsers_(ss);
  const progressRecords = getProgressRecordCount_(ss);
  const quizResultsCount = getQuizResultsCount_(ss);
  const certificateCount = getCertificateCount_(ss);

  return {
    spreadsheetUrl: ss.getUrl(),
    moduleCount: modules.length,
    updateCount: updates.length,
    faqCount: faq.length,
    userCount: users.length,
    progressRecords: progressRecords,
    quizResultsCount: quizResultsCount,
    certificateCount: certificateCount,
    users: users,
    viewedBy: viewerEmail || ''
  };
}

function getUsers_(ss) {
  const sh = ss.getSheetByName('Users');
  const values = sh.getDataRange().getValues();
  const result = [];

  for (let i = 1; i < values.length; i++) {
    result.push({
      email: String(values[i][0] || '').trim(),
      name: String(values[i][1] || '').trim(),
      role: String(values[i][2] || '').trim(),
      enabled: String(values[i][3] || '').trim()
    });
  }

  return result;
}

function getProgressRecordCount_(ss) {
  const sh = ss.getSheetByName('Progress');
  return Math.max(0, sh.getLastRow() - 1);
}

function getQuizResultsCount_(ss) {
  const sh = ss.getSheetByName('QuizResults');
  return Math.max(0, sh.getLastRow() - 1);
}

function getCertificateCount_(ss) {
  const sh = ss.getSheetByName('Certificates');
  return Math.max(0, sh.getLastRow() - 1);
}

function addUpdate(payload) {
  const user = getUserInfo_();
  const access = evaluateUserAccess_(user);
  if (!access.isAdmin) throw new Error('Only admins can add updates.');

  const title = String(payload.title || '').trim();
  const message = String(payload.message || '').trim();
  const type = String(payload.type || 'info').trim().toLowerCase();
  const updateTopic = String(payload.updateTopic || 'Update').trim();
  const updateCategory = String(payload.updateCategory || 'Portal').trim();

  if (!title) throw new Error('Title is required.');
  if (!message) throw new Error('Message is required.');

  const validTypes = ['info', 'success', 'warning', 'announcement'];
  if (validTypes.indexOf(type) === -1) throw new Error('Invalid update type.');

  const ss = getOrCreateSpreadsheet_();
  const sh = ss.getSheetByName('Updates');
  const now = new Date();
  const updatedBy = String(user.name || user.email || '').trim();

  sh.appendRow([
    'Y',
    title,
    message,
    type,
    now,
    updatedBy,
    updateTopic,
    updateCategory
  ]);

  return getUpdates_(ss);
}

function addUser(payload) {
  const user = getUserInfo_();
  const access = evaluateUserAccess_(user);
  if (!access.isAdmin) throw new Error('Only admins can add users.');

  const email = String(payload.email || '').trim().toLowerCase();
  const name = String(payload.name || '').trim();
  const role = String(payload.role || 'Trainee').trim();

  if (!email) throw new Error('Email is required.');
  if (!name) throw new Error('Name is required.');

  const validRoles = ['Trainee', 'Trainer', 'Admin', 'Super Admin'];
  if (validRoles.indexOf(role) === -1) throw new Error('Invalid role.');

  const ss = getOrCreateSpreadsheet_();
  const sh = ss.getSheetByName('Users');
  const values = sh.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    const rowEmail = String(values[i][0] || '').trim().toLowerCase();
    if (rowEmail === email) throw new Error('User already exists.');
  }

  sh.appendRow([email, name, role, 'Y']);
  logAdminAction_(ss, user.email, 'ADD_USER', {
    email: email,
    role: role
  });
  return getUsers_(ss);
}

function logAdminAction_(ss, actorEmail, actionName, details) {
  const sh = ss.getSheetByName('AdminActionLog');
  if (!sh) return;

  sh.appendRow([
    new Date(),
    String(actorEmail || '').trim().toLowerCase(),
    String(actionName || '').trim(),
    JSON.stringify(details || {})
  ]);
}

function getAdminUserProgressReport() {
  const user = getUserInfo_();
  const access = evaluateUserAccess_(user);
  if (!access.isAdmin) throw new Error('Only admins can access reports.');

  const ss = getOrCreateSpreadsheet_();
  const users = getUsers_(ss);
  const report = [];

  for (let i = 0; i < users.length; i++) {
    if (String(users[i].enabled || '').toUpperCase() !== 'Y') continue;

    const email = users[i].email;
    const progressSummary = getProgressSummary_(ss, email);
    const quizSummary = getQuizSummary_(ss, email);

    report.push({
      email: email,
      name: users[i].name,
      role: users[i].role,
      progressPercent: progressSummary.progressPercent,
      completedLessons: progressSummary.completedLessons,
      totalLessons: progressSummary.totalLessons,
      avgQuizScore: quizSummary.avgQuizScore,
      certificatesEarned: quizSummary.certificatesEarned
    });
  }

  report.sort(function(a, b) {
    return String(a.name || '').localeCompare(String(b.name || ''));
  });

  return report;
}

/**
 * HELPERS
 */
function findModuleById_(ss, moduleId) {
  const modules = getModules_(ss);
  for (let i = 0; i < modules.length; i++) {
    if (modules[i].id === moduleId) return modules[i];
  }
  return null;
}

function addDailyLearningItem(payload) { 
  const user = getUserInfo_();
  const access = evaluateUserAccess_(user);
  if (!access.isAdmin) throw new Error('Only admins can add daily learning items.');

  const type = String(payload.type || 'Tip').trim();
  const title = String(payload.title || '').trim();
  const content = String(payload.content || '').trim();
  const subtext = String(payload.subtext || '').trim();
  const tag = String(payload.tag || '').trim();
  const optionA = String(payload.optionA || '').trim();
  const optionB = String(payload.optionB || '').trim();
  const optionC = String(payload.optionC || '').trim();
  const optionD = String(payload.optionD || '').trim();
  const correctOption = String(payload.correctOption || '').trim().toUpperCase();
  const audience = String(payload.audience || 'All').trim() || 'All';
  const xp = Number(payload.xp || 5);
  let cardId = String(payload.cardId || '').trim();

  if (!title) throw new Error('Title is required.');
  if (!content) throw new Error('Content is required.');

  const validTypes = ['Tip', 'Vocabulary', 'Quote', 'Quiz'];
  if (validTypes.indexOf(type) === -1) throw new Error('Invalid daily learning type.');

  if (type === 'Quiz') {
    if (!optionA || !optionB || !optionC || !optionD) {
      throw new Error('All 4 quiz options are required for Quiz type.');
    }
    if (['A', 'B', 'C', 'D'].indexOf(correctOption) === -1) {
      throw new Error('Correct option must be A, B, C, or D for Quiz type.');
    }
  }

  const ss = getOrCreateSpreadsheet_();
  const dailySh = ss.getSheetByName('DailyLearning');
  if (!dailySh) throw new Error('DailyLearning sheet not found.');

  const values = dailySh.getDataRange().getValues();
  let maxDisplayOrder = 0;

  for (let i = 1; i < values.length; i++) {
    const rowOrder = Number(values[i][11] || 0);
    if (rowOrder > maxDisplayOrder) maxDisplayOrder = rowOrder;
  }

  const nextDisplayOrder = maxDisplayOrder + 1;

  if (!cardId) {
    cardId = 'DL_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  }

  dailySh.appendRow([
    'Y',
    type,
    title,
    content,
    subtext,
    tag,
    optionA,
    optionB,
    optionC,
    optionD,
    correctOption,
    nextDisplayOrder,
    cardId,
    audience,
    isNaN(xp) ? 5 : xp
  ]);

  const updatesSh = ss.getSheetByName('Updates');
  if (updatesSh) {
    updatesSh.appendRow([
      'Y',
      'Daily Learning Added',
      title + ' was added to Daily Learning.',
      'info',
      new Date(),
      String(user.name || user.email || '').trim(),
      'Learning',
      'AdminLog'
    ]);
  }

  clearPortalCaches();
  return {
    success: true,
    cardId: cardId
  };
}
function getOrCreateSpreadsheet_() {
  const ss = SpreadsheetApp.openById('11hc9yxH9F6P8SMQ62ry42WgPYmOPuSOj02U18XD9R1M');

  ensureSheet_(ss, 'Users', ['Email', 'Name', 'Role', 'Enabled']);
  ensureSheet_(ss, 'Modules', ['Enabled', 'Module ID', 'Module Title', 'Module Subtitle', 'Lesson Order', 'Lesson Name', 'Lesson Type', 'Lesson Link', 'Video Embed URL', 'Training Visual File ID']);
  ensureSheet_(ss, 'Updates', ['Enabled', 'Title', 'Message', 'Type', 'Updated At', 'Updated By', 'Update Topic', 'Update Category']);
  ensureSheet_(ss, 'Progress', ['Email', 'Module ID', 'Lesson Name', 'Status', 'Updated At']);
  ensureSheet_(ss, 'Quizzes', ['Enabled', 'Module ID', 'Question ID', 'Question', 'Option A', 'Option B', 'Option C', 'Option D', 'Correct Option', 'Explanation']);
  ensureSheet_(ss, 'QuizResults', ['Email', 'Module ID', 'Score Percent', 'Result', 'Answers JSON', 'Submitted At']);
  ensureSheet_(ss, 'Certificates', ['Email', 'Module ID', 'Certificate Name', 'File URL', 'Issued At', 'File ID', 'Certificate ID']);
  ensureSheet_(ss, 'DailyLearning', ['Enabled', 'Type', 'Title', 'Content', 'Subtext', 'Tag', 'Option A', 'Option B', 'Option C', 'Option D', 'Correct Option', 'Display Order', 'Card ID', 'Audience', 'XP']);
  ensureSheet_(ss, 'DailyLearningProgress', ['Email', 'Type', 'Title', 'Content', 'Tag', 'Card ID', 'XP Earned', 'Learned At']);
  ensureSheet_(ss, 'UserStreaks', ['Email', 'Current Streak', 'Best Streak', 'Total XP', 'Last Learned Date']);
  ensureSheet_(ss, 'AdminActionLog', ['Timestamp', 'Actor Email', 'Action', 'Details JSON']);
  return ss;
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);

  const firstRow = sh.getRange(1, 1, 1, headers.length).getValues()[0];
  const isEmpty = firstRow.every(function(cell) { return cell === ''; });

  if (isEmpty) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  }
}

function seedUsersSheet_(ss) {
  const sh = ss.getSheetByName('Users');
  if (sh.getLastRow() > 1) return;

  const rows = [
    ['admin@firstconnecthealth.com', 'Portal Admin', 'Admin', 'Y'],
    ['zimik.s@firstconnecthealth.com', 'Portal Admin', 'Admin', 'Y'],
    ['jane.s@firstconnecthealth.com', 'Training Lead', 'Trainer', 'Y'],
    ['trainer@firstconnecthealth.com', 'Training Lead', 'Trainer', 'Y']
  ];
  sh.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

function seedModulesSheet_(ss) {
  const sh = ss.getSheetByName('Modules');
  if (sh.getLastRow() > 1) return;

  const rows = [
    ['Y', 'M0', 'Pre-Week Orientation', 'Practical Knowledge', 1, 'US Basics', 'Video', '', ''],
    ['Y', 'M0', 'Pre-Week Orientation', 'Practical Knowledge', 2, 'States', 'Video', '', ''],
    ['Y', 'M0', 'Pre-Week Orientation', 'Practical Knowledge', 3, 'SSN', 'Video', '', ''],
    ['Y', 'M0', 'Pre-Week Orientation', 'Practical Knowledge', 4, 'Visa', 'Video', '', ''],

    ['Y', 'M1', 'Week 1 Foundation', 'Recruiter Basics', 1, 'Recruiter Role and Responsibilities', 'Video', '', ''],
    ['Y', 'M1', 'Week 1 Foundation', 'Recruiter Basics', 2, 'Job Types and Shifts', 'Video', '', ''],
    ['Y', 'M1', 'Week 1 Foundation', 'Recruiter Basics', 3, 'Pay Structure', 'Video', '', ''],
    ['Y', 'M1', 'Week 1 Foundation', 'Recruiter Basics', 4, 'Licensing', 'Video', '', ''],
    ['Y', 'M1', 'Week 1 Foundation', 'Recruiter Basics', 5, 'Compliance Basics', 'Video', '', ''],

    ['Y', 'M2', 'Week 2 Industry Understanding', 'Healthcare Ecosystem', 1, 'Healthcare System and Facilities', 'Video', '', ''],
    ['Y', 'M2', 'Week 2 Industry Understanding', 'Healthcare Ecosystem', 2, 'Healthcare Roles', 'Video', '', ''],
    ['Y', 'M2', 'Week 2 Industry Understanding', 'Healthcare Ecosystem', 3, 'Specialties and Sub-specialties', 'Video', '', ''],

    ['Y', 'M3', 'Week 3 Execution Skills', 'Practical Recruiting', 1, 'Candidate Screening', 'Video', '', ''],
    ['Y', 'M3', 'Week 3 Execution Skills', 'Practical Recruiting', 2, 'Sourcing and Boolean Search', 'Video', '', ''],
    ['Y', 'M3', 'Week 3 Execution Skills', 'Practical Recruiting', 3, 'Submission Logic', 'Video', '', ''],
    ['Y', 'M3', 'Week 3 Execution Skills', 'Practical Recruiting', 4, 'Follow-up and Candidate Management', 'Video', '', '']
  ];

  sh.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

function seedAnnouncementsSheet_(ss) {
  const sh = ss.getSheetByName('Announcements');
  if (sh.getLastRow() > 1) return;

  const rows = [
    ['Y', 'Welcome to the Training Portal', 'Use this portal to track lessons, complete quizzes, and access company resources.', 'info'],
    ['Y', 'Phase 3 Enabled', 'Role-based admin tools, quizzes, and certificates are now active.', 'success']
  ];

  sh.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

function seedProgressSheet_(ss) {
  const sh = ss.getSheetByName('Progress');
  if (sh.getLastRow() > 1) return;
}

function seedQuizzesSheet_(ss) {
  const sh = ss.getSheetByName('Quizzes');
  if (sh.getLastRow() > 1) return;

  const rows = [
    ['Y', 'M0', 'M0Q1', 'What does SSN stand for?', 'Social Security Number', 'State Security Number', 'System Serial Number', 'Social State Number', 'A', 'SSN stands for Social Security Number.'],
    ['Y', 'M0', 'M0Q2', 'Visa topic belongs to which module?', 'Week 3', 'Pre-Week Orientation', 'Admin Module', 'FAQ', 'B', 'Visa is part of the pre-orientation foundation.'],

    ['Y', 'M1', 'M1Q1', 'Which topic is part of Week 1 Foundation?', 'Healthcare Facilities', 'Submission Logic', 'Licensing', 'States', 'C', 'Licensing is part of Week 1 Foundation.'],
    ['Y', 'M1', 'M1Q2', 'Which one is directly connected to recruiter basics?', 'Pay Structure', 'Specialties', 'SSN', 'Visa', 'A', 'Pay Structure belongs to recruiter basics.'],

    ['Y', 'M2', 'M2Q1', 'Which module covers Healthcare Roles?', 'Week 2 Industry Understanding', 'Week 1 Foundation', 'Pre-Week Orientation', 'Admin Snapshot', 'A', 'Healthcare Roles is in Week 2 Industry Understanding.'],
    ['Y', 'M2', 'M2Q2', 'Specialties and Sub-specialties belong to?', 'Week 3', 'Week 2', 'Week 1', 'FAQ', 'B', 'This topic belongs to Week 2.'],

    ['Y', 'M3', 'M3Q1', 'Boolean Search belongs to which week?', 'Pre-Week', 'Week 1', 'Week 3', 'FAQ', 'C', 'Boolean Search is part of Week 3 Execution Skills.'],
    ['Y', 'M3', 'M3Q2', 'Candidate Screening is mainly a?', 'Practical recruiting topic', 'Company policy only', 'Payroll topic', 'Leave process', 'A', 'Candidate Screening is a practical recruiting skill.']
  ];

  sh.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

function seedQuizResultsSheet_(ss) {
  const sh = ss.getSheetByName('QuizResults');
  if (sh.getLastRow() > 1) return;
}

function seedCertificatesSheet_(ss) {
  const sh = ss.getSheetByName('Certificates');
  if (sh.getLastRow() > 1) return;
}

function addQuizQuestion(payload) {
  const user = getUserInfo_();
  const access = evaluateUserAccess_(user);
  if (!access.isAdmin) throw new Error('Only admins can add quiz questions.');

  payload = payload || {};

  const moduleId = String(payload.moduleId || '').trim();
  const questionId = String(payload.questionId || '').trim();
  const question = String(payload.question || '').trim();
  const optionA = String(payload.optionA || '').trim();
  const optionB = String(payload.optionB || '').trim();
  const optionC = String(payload.optionC || '').trim();
  const optionD = String(payload.optionD || '').trim();
  const correctOption = String(payload.correctOption || '').trim().toUpperCase();
  const explanation = String(payload.explanation || '').trim();

  if (!moduleId) throw new Error('Module ID is required.');
  if (!questionId) throw new Error('Question ID is required.');
  if (!question) throw new Error('Question is required.');
  if (!optionA || !optionB || !optionC || !optionD) {
    throw new Error('All 4 options are required.');
  }
  if (['A', 'B', 'C', 'D'].indexOf(correctOption) === -1) {
    throw new Error('Correct option must be A, B, C, or D.');
  }

  const ss = getOrCreateSpreadsheet_();
  const sh = ss.getSheetByName('Quizzes');
  if (!sh) throw new Error('Quizzes sheet not found.');

  const values = sh.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    const enabled = String(values[i][0] || '').trim().toUpperCase();
    const rowModuleId = String(values[i][1] || '').trim();
    const rowQuestionId = String(values[i][2] || '').trim();

    if (enabled === 'Y' && rowModuleId === moduleId && rowQuestionId === questionId) {
      throw new Error('A quiz question with this Module ID and Question ID already exists.');
    }
  }

  sh.appendRow([
    'Y',
    moduleId,
    questionId,
    question,
    optionA,
    optionB,
    optionC,
    optionD,
    correctOption,
    explanation
  ]);

  const updatesSh = ss.getSheetByName('Updates');
  if (updatesSh) {
    updatesSh.appendRow([
      'Y',
      'Quiz Question Added',
      'Quiz question "' + question + '" was added for module ' + moduleId + '.',
      'info',
      new Date(),
      String(user.name || user.email || '').trim(),
      'Quiz',
      'AdminLog'
    ]);
  }

  clearPortalCaches();

  return {
    success: true,
    moduleId: moduleId,
    questionId: questionId
  };
}
function getOrCreateFolderByName_(name) {
  const folders = DriveApp.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(name);
}

function formatNameFromEmail_(email) {
  return String(email || '')
    .split('@')[0]
    .replace(/[._-]/g, ' ')
    .replace(/\b\w/g, function(c) { return c.toUpperCase(); });
}

function formatDateSafe_(value) {
  try {
    if (!value) return '';
    return Utilities.formatDate(new Date(value), Session.getScriptTimeZone(), 'dd MMM yyyy');
  } catch (err) {
    return String(value || '');
  }
}

function escapeHtmlServer_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

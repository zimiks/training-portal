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
  seedSettingsSheet_(ss);
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
      weeklyStatus: []
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

  portalData.auth = authState;
  portalData.dailyLearning = {
    items: dailyLearningItems,
    todayCard: todayCard,
    progress: getDailyLearningUserProgress_(ss, user.email),
    streak: getUserStreak_(ss, user.email),
    meta: getDailyLearningMeta_(user.email, todayCard),
    weeklyStatus: getWeeklyDailyStatus_(user.email)
  };
  portalData.updates = updates;
  portalData.progressSummary = progressSummary;
  portalData.quizSummary = quizSummary;
  portalData.dashboardStats = dashboardStats;
  portalData.recommendedAction = getRecommendedAction_(portalData);

  return portalData;
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
  var card = getDailyLearningCardById_(cardId);
  var meta = getDailyLearningMeta_(email, card);
  var weekly = getWeeklyDailyStatus_(email);

  return {
    success: true,
    cardId: String(cardId || ''),
    xpEarned: Number(card && card.xp ? card.xp : 0),
    totalXp: meta.totalXp,
    todayXp: meta.todayXp,
    streak: meta.streak,
    todayCompleted: meta.todayCompleted,
    weeklyDailyStatus: weekly
  };
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
  const sh = ss.getSheetByName('Settings');
  const values = sh.getDataRange().getValues();
  const map = {};

  for (let i = 1; i < values.length; i++) {
    const key = String(values[i][0] || '').trim();
    const value = String(values[i][1] || '').trim();
    if (key) map[key] = value;
  }

  return {
    certificateFolderName: map.certificateFolderName || CONFIG.DEFAULT_CERTIFICATE_FOLDER_NAME,
    chatbotMessage: map.chatbotMessage || 'AI chatbot can be connected later to Drive docs, FAQs, and Sheets-based training knowledge.',
    welcomeBanner: map.welcomeBanner || 'Welcome to the First Connect Health Training Portal'
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
        visualFileId: trainingVisualFileId,
        lessons: []
      };
    }
if (!grouped[moduleId].visualFileId && trainingVisualFileId) {
  grouped[moduleId].visualFileId = trainingVisualFileId;
}
    grouped[moduleId].lessons.push({
      order: lessonOrder,
      name: lessonName,
      type: lessonType || 'Lesson',
      link: lessonLink,
      videoEmbedUrl: videoEmbedUrl
    });
  }

  const modules = Object.keys(grouped).map(function(key) {
    grouped[key].lessons.sort(function(a, b) {
      return a.order - b.order;
    });
    return grouped[key];
  });

  modules.sort(function(a, b) {
    return getNumberFromModuleId_(a.id) - getNumberFromModuleId_(b.id);
  });

  return modules;
}

function getNumberFromModuleId_(value) {
  const n = Number(String(value || '').replace(/[^\d]/g, ''));
  return isNaN(n) ? 9999 : n;
}

/**
 * FAQ / ANNOUNCEMENTS
 */
function getFaq_(ss) {
  const sh = ss.getSheetByName('FAQ');
  const values = sh.getDataRange().getValues();
  const items = [];

  for (let i = 1; i < values.length; i++) {
    const enabled = String(values[i][0] || '').trim().toUpperCase();
    if (enabled !== 'Y') continue;

    items.push({
      question: String(values[i][1] || '').trim(),
      answer: String(values[i][2] || '').trim()
    });
  }

  return items;
}
function getUpdates_(ss) {
  const sh = ss.getSheetByName('Updates');
  const values = sh.getDataRange().getValues();
  const items = [];

  for (let i = 1; i < values.length; i++) {
    const enabled = String(values[i][0] || '').trim().toUpperCase();
    if (enabled !== 'Y') continue;

    items.push({
      rowNumber: i + 1,
      title: String(values[i][1] || '').trim(),
      message: String(values[i][2] || '').trim(),
      type: String(values[i][3] || 'info').trim() || 'info'
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

  return rows.map(function(r) {
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

  // ✅ Get real name from Users sheet
  let userName = '';
  const users = getUsers_(ss);
  const target = String(email || '').trim().toLowerCase();

  for (let i = 0; i < users.length; i++) {
    if (String(users[i].email || '').trim().toLowerCase() === target) {
      userName = users[i].name;
      break;
    }
  }

  // fallback
  if (!userName) {
    userName = formatNameFromEmail_(email);
  }

  const module = findModuleById_(ss, moduleId);
  const moduleTitle = module ? module.title : moduleId;

  const html = buildCertificateHtml_(userName, moduleTitle);
  const blob = Utilities.newBlob(html, 'text/html', 'certificate.html');

  const file = folder.createFile(blob)
    .setName('Certificate - ' + userName + ' - ' + moduleTitle + '.html');

  const certSh = ss.getSheetByName('Certificates');
  certSh.appendRow([
    email,
    moduleId,
    file.getName(),
    file.getUrl(),
    new Date()
  ]);

  return {
    name: file.getName(),
    url: file.getUrl(),
    moduleId: moduleId,
    moduleTitle: moduleTitle
  };
}

function buildCertificateHtml_(userName, moduleTitle) {
  const issuedDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy');
  const logoUrl = 'https://drive.google.com/uc?export=view&id=1qGHveKbBlPXHo44lx_KwngmIYhZHWjPp';

  return [
    '<!DOCTYPE html>',
    '<html><head><meta charset="utf-8"><title>Certificate</title>',
    '<style>',
    '@page{size:A4 landscape;margin:0;}',
    'html,body{margin:0;padding:0;background:#eef3f8;font-family:Arial,sans-serif;color:#1d2c3c;}',
    'body{padding:18px;box-sizing:border-box;}',
    '.sheet{width:1123px;min-height:794px;margin:0 auto;background:linear-gradient(135deg,#fafbfd 0%,#f3f6fa 38%,#eef2f7 100%);position:relative;overflow:hidden;}',
    '.sheet:before{content:"";position:absolute;left:-120px;top:-80px;width:420px;height:420px;background:radial-gradient(circle at center, rgba(34,76,124,0.08) 0%, rgba(34,76,124,0.03) 45%, rgba(34,76,124,0) 72%);border-radius:50%;}',
    '.sheet:after{content:"";position:absolute;right:-180px;bottom:-180px;width:620px;height:620px;background:radial-gradient(circle at center, rgba(34,76,124,0.08) 0%, rgba(34,76,124,0.03) 42%, rgba(34,76,124,0) 72%);border-radius:50%;}',
    '.frame{position:absolute;top:24px;left:24px;right:24px;bottom:24px;border:6px solid #304b73;border-radius:18px;box-sizing:border-box;}',
    '.frame-inner{position:absolute;top:12px;left:12px;right:12px;bottom:12px;border:2px solid #304b73;border-radius:12px;box-sizing:border-box;}',
    '.corner{position:absolute;width:54px;height:54px;border:4px solid #304b73;border-radius:0 0 28px 0;background:transparent;}',
    '.corner.tl{top:-6px;left:-6px;border-right:none;border-bottom:none;border-top-left-radius:18px;border-bottom-right-radius:0;}',
    '.corner.tr{top:-6px;right:-6px;border-left:none;border-bottom:none;border-top-right-radius:18px;border-bottom-right-radius:0;}',
    '.corner.bl{bottom:-6px;left:-6px;border-right:none;border-top:none;border-bottom-left-radius:18px;border-bottom-right-radius:0;}',
    '.corner.br{bottom:-6px;right:-6px;border-left:none;border-top:none;border-bottom-right-radius:18px;border-bottom-left-radius:0;}',
    '.content{position:relative;z-index:2;padding:74px 92px 56px;text-align:center;}',
    '.logo-wrap{position:absolute;top:46px;right:68px;text-align:right;z-index:3;max-width:220px;}',
    '.logo{max-width:170px;max-height:64px;object-fit:contain;display:block;margin-left:auto;}',
    '.logo-name{margin-top:8px;font-size:11px;letter-spacing:3px;text-transform:uppercase;color:#2f5f98;font-weight:700;}',
    '.title{font-family:Georgia,"Times New Roman",serif;font-size:78px;line-height:1.02;font-weight:400;color:#25385a;margin:26px 0 10px;}',
    '.subtitle{font-size:24px;letter-spacing:3px;text-transform:uppercase;color:#111;margin:0 0 42px;}',
    '.line1{font-size:22px;color:#2a2a2a;margin:0 0 14px;}',
    '.name{font-family:"Times New Roman",Georgia,serif;font-size:58px;line-height:1.15;font-style:italic;color:#222;margin:0 0 26px;}',
    '.rule{width:64%;height:1px;background:#555;margin:0 auto 36px;}',
    '.line2{font-size:18px;letter-spacing:4px;color:#444;font-style:italic;margin:0 0 18px;text-transform:none;}',
    '.module{max-width:760px;margin:0 auto;font-size:34px;line-height:1.25;color:#1c4f89;font-weight:700;}',
    '.meta-row{margin-top:72px;display:flex;justify-content:space-between;align-items:flex-end;gap:36px;}',
    '.sign-block{flex:1;text-align:center;}',
    '.sign-line{width:78%;height:1px;background:#666;margin:0 auto 14px;}',
    '.sign-name{font-size:16px;letter-spacing:3px;text-transform:uppercase;color:#333;}',
    '.sign-role{font-size:14px;letter-spacing:3px;text-transform:lowercase;color:#555;margin-top:4px;}',
    '.seal{width:118px;height:118px;border-radius:50%;margin:0 auto;background:radial-gradient(circle at 35% 30%, #f8e7a1 0%, #e6c565 45%, #c99b34 70%, #b98522 100%);border:7px solid #d7b25a;box-shadow:0 0 0 5px #2c446b, 0 10px 20px rgba(0,0,0,0.15);position:relative;}',
    '.seal:before{content:"";position:absolute;left:50%;transform:translateX(-50%);bottom:-38px;width:0;height:0;border-left:22px solid transparent;border-right:22px solid transparent;border-top:42px solid #243f6a;}',
    '.seal:after{content:"";position:absolute;left:50%;transform:translateX(-50%) translateX(34px);bottom:-38px;width:0;height:0;border-left:22px solid transparent;border-right:22px solid transparent;border-top:42px solid #243f6a;}',
    '.seal-center{position:absolute;inset:18px;border-radius:50%;border:3px solid rgba(255,255,255,0.55);display:flex;align-items:center;justify-content:center;font-size:12px;font-weight:700;letter-spacing:1px;color:#5a3c06;text-transform:uppercase;text-align:center;line-height:1.25;}',
    '.issued{margin-top:40px;font-size:16px;color:#4c5d70;}',
    '.issued strong{color:#243f6a;}',
    '@media print{body{padding:0;background:#fff;}.sheet{margin:0;}.frame{top:10px;left:10px;right:10px;bottom:10px;}}',
    '</style></head><body>',
    '<div class="sheet">',
    '<div class="frame">',
    '<div class="frame-inner"></div>',
    '<div class="corner tl"></div>',
    '<div class="corner tr"></div>',
    '<div class="corner bl"></div>',
    '<div class="corner br"></div>',
    '</div>',
    '<div class="logo-wrap">',
    '<img class="logo" src="' + logoUrl + '" alt="First Connect Health Logo">',
    '<div class="logo-name">First Connect Health</div>',
    '</div>',
    '<div class="content">',
    '<div class="title">Certificate</div>',
    '<div class="subtitle">of Completion</div>',
    '<div class="line1">This certificate is proudly presented to</div>',
    '<div class="name">' + escapeHtmlServer_(userName) + '</div>',
    '<div class="rule"></div>',
    '<div class="line2">for successfully completing the training module</div>',
    '<div class="module">' + escapeHtmlServer_(moduleTitle) + '</div>',
    '<div class="issued">Issued on <strong>' + escapeHtmlServer_(issuedDate) + '</strong></div>',
    '<div class="meta-row">',
    '<div class="sign-block">',
    '<div class="sign-line"></div>',
    '<div class="sign-name">Training Team</div>',
    '<div class="sign-role">First Connect Health</div>',
    '</div>',
    '<div class="sign-block" style="flex:0 0 180px;">',
    '<div class="seal"><div class="seal-center">Certified<br>Learning</div></div>',
    '</div>',
    '<div class="sign-block">',
    '<div class="sign-line"></div>',
    '<div class="sign-name">Organization</div>',
    '<div class="sign-role">Training Portal</div>',
    '</div>',
    '</div>',
    '</div>',
    '</div>',
    '</body></html>'
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

  const aliases = GmailApp.getAliases().map(function(alias) {
    return String(alias || '').trim().toLowerCase();
  });

  if (aliases.indexOf(preferredAlias) > -1) {
    GmailApp.sendEmail(email, subject, body, {
      from: preferredAlias,
      name: 'Training Portal',
      replyTo: CONFIG.SUPPORT_EMAIL
    });
  } else {
    MailApp.sendEmail({
      to: email,
      subject: subject,
      body: body,
      name: 'Training Portal',
      noReply: true
    });
  }

  return {
    success: true,
    message: 'OTP sent to your enabled email address.'
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
  const values = sh.getDataRange().getValues();
  const target = String(email || '').trim().toLowerCase();
  const items = [];

  for (let i = 1; i < values.length; i++) {
    const rowEmail = String(values[i][0] || '').trim().toLowerCase();
    if (rowEmail !== target) continue;

    items.push({
      moduleId: String(values[i][1] || '').trim(),
      name: String(values[i][2] || '').trim(),
      url: String(values[i][3] || '').trim(),
      issuedAt: formatDateSafe_(values[i][4])
    });
  }

  return items;
}

function getCertificateByModule_(ss, email, moduleId) {
  const sh = ss.getSheetByName('Certificates');
  const values = sh.getDataRange().getValues();
  const targetEmail = String(email || '').trim().toLowerCase();
  const targetModuleId = String(moduleId || '').trim();

  for (let i = 1; i < values.length; i++) {
    const rowEmail = String(values[i][0] || '').trim().toLowerCase();
    const rowModuleId = String(values[i][1] || '').trim();
    if (rowEmail === targetEmail && rowModuleId === targetModuleId) {
      return {
        moduleId: rowModuleId,
        name: String(values[i][2] || '').trim(),
        url: String(values[i][3] || '').trim(),
        issuedAt: formatDateSafe_(values[i][4]),
        moduleTitle: (findModuleById_(ss, rowModuleId) || {}).title || rowModuleId
      };
    }
  }

  return null;
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

  if (!title) throw new Error('Title is required.');
  if (!message) throw new Error('Message is required.');

  const validTypes = ['info', 'success', 'warning', 'announcement'];
  if (validTypes.indexOf(type) === -1) throw new Error('Invalid update type.');

  const ss = getOrCreateSpreadsheet_();
  const sh = ss.getSheetByName('Updates');
  sh.appendRow(['Y', title, message, type]);
  logAdminAction_(ss, user.email, 'ADD_UPDATE', {
    title: title,
    type: type
  });

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


function getOrCreateSpreadsheet_() {
  const ss = SpreadsheetApp.openById('11hc9yxH9F6P8SMQ62ry42WgPYmOPuSOj02U18XD9R1M');

  ensureSheet_(ss, 'Users', ['Email', 'Name', 'Role', 'Enabled']);
  ensureSheet_(ss, 'Modules', ['Enabled', 'Module ID', 'Module Title', 'Module Subtitle', 'Lesson Order', 'Lesson Name', 'Lesson Type', 'Lesson Link', 'Video Embed URL', 'Training Visual File ID']);
  ensureSheet_(ss, 'FAQ', ['Enabled', 'Question', 'Answer']);
  ensureSheet_(ss, 'Updates', ['Enabled', 'Title', 'Message', 'Type']);
  ensureSheet_(ss, 'Progress', ['Email', 'Module ID', 'Lesson Name', 'Status', 'Updated At']);
  ensureSheet_(ss, 'Quizzes', ['Enabled', 'Module ID', 'Question ID', 'Question', 'Option A', 'Option B', 'Option C', 'Option D', 'Correct Option', 'Explanation']);
  ensureSheet_(ss, 'QuizResults', ['Email', 'Module ID', 'Score Percent', 'Result', 'Answers JSON', 'Submitted At']);
  ensureSheet_(ss, 'Certificates', ['Email', 'Module ID', 'Certificate Name', 'File URL', 'Issued At']);
  ensureSheet_(ss, 'Settings', ['Key', 'Value']);
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

function seedFaqSheet_(ss) {
  const sh = ss.getSheetByName('FAQ');
  if (sh.getLastRow() > 1) return;

  const rows = [
    ['Y', 'How do I apply for leave?', 'Please contact HR or follow the internal leave request process.'],
    ['Y', 'Who do I contact for attendance issues?', 'Please contact HR and your reporting manager.'],
    ['Y', 'Where are training files stored?', 'Training resources are available inside the portal and Drive resources section.'],
    ['Y', 'What should I do if I miss a session?', 'Contact the training team and review the relevant module and files.']
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

function seedSettingsSheet_(ss) {
  const sh = ss.getSheetByName('Settings');
  if (sh.getLastRow() > 1) return;

  const rows = [
    ['certificateFolderName', CONFIG.DEFAULT_CERTIFICATE_FOLDER_NAME],
    ['chatbotMessage', 'AI chatbot can later be connected to Drive files, FAQs, and sheet-based training knowledge.'],
    ['welcomeBanner', 'Welcome to the First Connect Health Training Portal']
  ];

  sh.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
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

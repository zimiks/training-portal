const KNOWLEDGE_SPREADSHEET_ID = '11hc9yxH9F6P8SMQ62ry42WgPYmOPuSOj02U18XD9R1M';
const ASSISTANT_SHEET_NAME = 'AI_Chatbot_Knowledge';
const ASSISTANT_UNANSWERED_SHEET_NAME = 'AI_Unanswered_Log';

/**
 * Grounded assistant backend
 * Reads only from AI_Chatbot_Knowledge
 */
function askAssistant(question) {
  const rawQuestion = String(question || '').trim();
  const cleanQuestion = normalizeAssistantText_(rawQuestion);

  try {
    Logger.log('QUESTION RECEIVED: ' + question);

    Logger.log('RAW QUESTION: ' + rawQuestion);
    Logger.log('CLEAN QUESTION: ' + cleanQuestion);

    if (!cleanQuestion) {
      return {
        ok: false,
        answer: 'Please type a question so I can help.',
        links: [],
        suggestions: getSmartSuggestions_(cleanQuestion),
        moduleId: ''
      };
    }

    const debugInfo = debugAskAssistant_(cleanQuestion);
    Logger.log('DEBUG INFO: ' + JSON.stringify(debugInfo));

    const matchResult = findBestAssistantAnswer_(cleanQuestion);
    const match = matchResult.match;

    Logger.log('MATCH RESULT: ' + JSON.stringify(matchResult));

    if (!match) {
      logUnansweredAssistantQuestion_(rawQuestion, cleanQuestion);
      return {
        ok: true,
        answer: 'I could not find a matching grounded answer in the chatbot knowledge sheet.',
        links: [],
        suggestions: getSmartSuggestions_(cleanQuestion),
        moduleId: ''
      };
    }

    var detectedModuleId =
      extractAssistantModuleId_(match.mainQuestion) ||
      extractAssistantModuleId_(match.answer);

    return {
      ok: true,
      answer: buildAssistantDisplayText_(match.answer),
      links: match.links || [],
      matchedQuestion: match.mainQuestion || '',
      rowId: match.id || '',
      confidence: match.score || 0,
      suggestions: getSmartSuggestions_(cleanQuestion),
      moduleId: detectedModuleId || ''
    };
  } catch (error) {
    Logger.log('askAssistant error: ' + error);
    Logger.log('askAssistant stack: ' + (error && error.stack ? error.stack : 'no stack'));

    return {
      ok: false,
      answer: 'Backend error: ' + (error && error.message ? error.message : String(error)),
      links: [],
      suggestions: getSmartSuggestions_(cleanQuestion),
      moduleId: ''
    };
  }
}
function extractAssistantModuleId_(text) {
  var value = String(text || '').trim();

  if (!value) {
    return '';
  }

  var directMatch = value.match(/\bM\s*([0-9]+)\b/i);
  if (directMatch && directMatch[1]) {
    return 'M' + directMatch[1];
  }

  var moduleMatch = value.match(/\bmodule\s*([0-9]+)\b/i);
  if (moduleMatch && moduleMatch[1]) {
    return 'M' + moduleMatch[1];
  }

  return '';
}
function buildAssistantDisplayText_(answer) {
  var text = String(answer || '').trim();

  if (!text) {
    return '';
  }

  text = text
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();

  var parts = text.split(/\n+/).map(function(part) {
    return String(part || '').trim();
  }).filter(Boolean);

  if (!parts.length) {
    return text;
  }

  return parts.join('\n\n');
}
function normalizeAssistantSuggestions_(suggestions, rawQuestion, match) {
  var cleaned = [];
  var seen = {};
  var raw = String(rawQuestion || '').trim().toLowerCase();
  var matchedQuestion = match && match.mainQuestion
    ? String(match.mainQuestion).trim().toLowerCase()
    : '';

  (Array.isArray(suggestions) ? suggestions : []).forEach(function(item) {
    var value = String(item || '').trim();
    var key = value.toLowerCase();

    if (!value) return;
    if (key === raw) return;
    if (matchedQuestion && key === matchedQuestion) return;
    if (seen[key]) return;

    seen[key] = true;
    cleaned.push(value);
  });

  if (!cleaned.length && match && match.mainQuestion) {
    [
      'Can you explain this step by step?',
      'Is there a related module for this?',
      'Do you want the exact process?'
    ].forEach(function(item) {
      var key = String(item).toLowerCase();
      if (!seen[key]) {
        seen[key] = true;
        cleaned.push(item);
      }
    });
  }

  if (!cleaned.length) {
    [
      'Show me the related module',
      'Explain it step by step',
      'Ask another training question'
    ].forEach(function(item) {
      var key = String(item).toLowerCase();
      if (!seen[key]) {
        seen[key] = true;
        cleaned.push(item);
      }
    });
  }

  return cleaned.slice(0, 3);
}
function debugAskAssistant_(cleanQuestion) {
  const spreadsheet = getAssistantSpreadsheet_();
  const sheet = spreadsheet.getSheetByName(ASSISTANT_SHEET_NAME);

  if (!sheet) {
    throw new Error('Sheet not found: ' + ASSISTANT_SHEET_NAME);
  }

  const values = sheet.getDataRange().getDisplayValues();

  if (!values || !values.length) {
    throw new Error('Sheet is empty: ' + ASSISTANT_SHEET_NAME);
  }

  const headers = values[0] || [];
  const totalRows = values.length;

  return {
    spreadsheetId: KNOWLEDGE_SPREADSHEET_ID,
    sheetName: ASSISTANT_SHEET_NAME,
    totalRows: totalRows,
    headers: headers,
    firstDataRow: values.length > 1 ? values[1] : [],
    cleanQuestion: cleanQuestion
  };
}
function getAssistantSpreadsheet_() {
  return SpreadsheetApp.openById(KNOWLEDGE_SPREADSHEET_ID);
}

function getAssistantKnowledgeRows_() {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'assistant_knowledge_rows_v4';
  const cached = cache.get(cacheKey);

  if (cached) {
    return JSON.parse(cached);
  }

  const spreadsheet = getAssistantSpreadsheet_();
  const sheet = spreadsheet.getSheetByName(ASSISTANT_SHEET_NAME);

  if (!sheet) {
    throw new Error('Sheet not found: ' + ASSISTANT_SHEET_NAME);
  }

  const values = sheet.getDataRange().getDisplayValues();
  if (!values || values.length < 2) {
    return { rows: [], headerMap: {} };
  }

  const headers = values[0];
  const rows = values.slice(1);
  const headerMap = buildAssistantHeaderMap_(headers);

  const payload = {
    rows: rows,
    headerMap: headerMap
  };

  try {
    const serialized = JSON.stringify(payload);

    if (serialized.length <= 90000) {
      cache.put(cacheKey, serialized, 300);
    }
  } catch (error) {
    Logger.log('Assistant cache skipped: ' + error);
  }

  return payload;
}
function findBestAssistantAnswer_(cleanQuestion) {
  const knowledge = getAssistantKnowledgeRows_();
  const rows = knowledge.rows || [];
  const headerMap = knowledge.headerMap || {};

  if (!rows.length) {
    return { match: null, suggestions: [] };
  }

  const exactMatch = findExactAssistantMatch_(rows, headerMap, cleanQuestion);
  if (exactMatch) {
    return { match: exactMatch, suggestions: [] };
  }

  const acronymMatch = findAssistantAcronymMatch_(rows, headerMap, cleanQuestion);
  if (acronymMatch) {
    return { match: acronymMatch, suggestions: [] };
  }

  const phraseMatch = findAssistantPhraseMatch_(rows, headerMap, cleanQuestion);
  if (phraseMatch) {
    return { match: phraseMatch, suggestions: [] };
  }

  let bestRow = null;
  let bestScore = 0;

  rows.forEach(function(row) {
    const answer = safeCell_(row, headerMap, 'Answer').trim();
    if (!answer) return;

    const mainQuestionRaw = safeCell_(row, headerMap, 'Main Question');
    const alternateQuestionsRaw = safeCell_(row, headerMap, 'Alternate Questions');
    const keywordsRaw = safeCell_(row, headerMap, 'Keywords');

    const mainQuestion = normalizeAssistantText_(mainQuestionRaw);
    const alternateParts = splitAssistantTerms_(alternateQuestionsRaw);
    const keywordParts = splitAssistantTerms_(keywordsRaw);

    let score = 0;

    score += scoreAssistantField_(cleanQuestion, mainQuestion, 10);
    score += scoreAssistantAlternateQuestions_(cleanQuestion, alternateQuestionsRaw, 12);
    score += scoreAssistantKeywords_(cleanQuestion, keywordsRaw, 10);

    if (mainQuestion && cleanQuestion.indexOf(mainQuestion) !== -1) {
      score += 6;
    }

    if (mainQuestion && mainQuestion.indexOf(cleanQuestion) !== -1) {
      score += 5;
    }

    alternateParts.forEach(function(alt) {
      if (!alt) return;

      if (alt === cleanQuestion) {
        score += 40;
      } else if (alt.indexOf(cleanQuestion) !== -1 || cleanQuestion.indexOf(alt) !== -1) {
        score += 12;
      }
    });

    keywordParts.forEach(function(keyword) {
      if (!keyword) return;

      if (keyword === cleanQuestion) {
        score += 30;
      } else if (keyword.indexOf(cleanQuestion) !== -1 || cleanQuestion.indexOf(keyword) !== -1) {
        score += 10;
      }
    });

    const qTokens = tokenizeAssistantText_(cleanQuestion);
    const mainTokens = tokenizeAssistantText_(mainQuestion);
    const sharedMainTokens = qTokens.filter(function(token) {
      return mainTokens.indexOf(token) !== -1;
    }).length;

    if (sharedMainTokens >= 2) {
      score += sharedMainTokens * 2;
    }

    alternateParts.forEach(function(alt) {
      const altTokens = tokenizeAssistantText_(alt);
      const sharedAltTokens = qTokens.filter(function(token) {
        return altTokens.indexOf(token) !== -1;
      }).length;

      if (sharedAltTokens >= 2) {
        score += sharedAltTokens * 3;
      } else if (sharedAltTokens === 1 && qTokens.length === 1) {
        score += 6;
      }
    });

    keywordParts.forEach(function(keyword) {
      const keywordTokens = tokenizeAssistantText_(keyword);
      const sharedKeywordTokens = qTokens.filter(function(token) {
        return keywordTokens.indexOf(token) !== -1;
      }).length;

      if (sharedKeywordTokens >= 2) {
        score += sharedKeywordTokens * 2;
      } else if (sharedKeywordTokens === 1 && qTokens.length === 1) {
        score += 5;
      }
    });

    if (score > bestScore) {
      bestScore = score;
      bestRow = row;
    }
  });

  if (!bestRow || bestScore < 8) {
    return {
      match: null,
      suggestions: []
    };
  }

  return {
    match: buildAssistantResult_(bestRow, headerMap, bestScore),
    suggestions: []
  };
}
function normalizeAssistantTokenForIntent_(token) {
  var value = String(token || '').trim().toLowerCase();

  if (!value) {
    return '';
  }

  if (value.length > 4 && /ies$/.test(value)) {
    return value.replace(/ies$/, 'y');
  }

  if (value.length > 4 && /sses$/.test(value)) {
    return value.replace(/es$/, '');
  }

  if (value.length > 3 && /s$/.test(value) && !/ss$/.test(value)) {
    return value.replace(/s$/, '');
  }

  return value;
}

function reduceAssistantIntentText_(value) {
  var text = normalizeAssistantText_(value);

  if (!text) {
    return '';
  }

  text = text
    .replace(/\b(can you|could you|would you|please|pls)\b/g, ' ')
    .replace(/\b(tell me|show me|help me|guide me|explain to me)\b/g, ' ')
    .replace(/\b(i want to know|i need to know|i want|i need)\b/g, ' ')
    .replace(/\b(what is|what are|how do i|how to|where is|where are|when is|when are|why is|why are)\b/g, ' ')
    .replace(/\b(the|a|an|my|your|about)\b/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  if (!text) {
    return '';
  }

  return text
    .split(' ')
    .map(function(token) {
      return normalizeAssistantTokenForIntent_(token);
    })
    .filter(Boolean)
    .join(' ');
}
function logUnansweredAssistantQuestion_(rawQuestion, cleanQuestion) {
  const spreadsheet = getAssistantSpreadsheet_();
  const sheet = ensureAssistantUnansweredSheet_(spreadsheet);
  const user = typeof getUserInfo_ === 'function' ? getUserInfo_() : { email: '' };

  sheet.appendRow([
    new Date(),
    String((user || {}).email || '').trim().toLowerCase(),
    rawQuestion,
    cleanQuestion
  ]);
}

function ensureAssistantUnansweredSheet_(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(ASSISTANT_UNANSWERED_SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(ASSISTANT_UNANSWERED_SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Timestamp', 'User Email', 'Original Question', 'Normalized Question']);
  }

  return sheet;
}

function findExactAssistantMatch_(rows, headerMap, cleanQuestion) {
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var answer = safeCell_(row, headerMap, 'Answer').trim();
    if (!answer) continue;

    var mainQuestion = normalizeAssistantText_(safeCell_(row, headerMap, 'Main Question'));
    var alternateParts = splitAssistantTerms_(safeCell_(row, headerMap, 'Alternate Questions'));
    var keywordParts = splitAssistantTerms_(safeCell_(row, headerMap, 'Keywords'));

    if (mainQuestion && mainQuestion === cleanQuestion) {
      return buildAssistantResult_(row, headerMap, 1000);
    }

    for (var j = 0; j < alternateParts.length; j++) {
      if (alternateParts[j] === cleanQuestion) {
        return buildAssistantResult_(row, headerMap, 950);
      }
    }

    for (var k = 0; k < keywordParts.length; k++) {
      if (keywordParts[k] === cleanQuestion) {
        return buildAssistantResult_(row, headerMap, 900);
      }
    }
  }

  return null;
}

function findAssistantAcronymMatch_(rows, headerMap, cleanQuestion) {
  if (!isLikelyAcronym_(cleanQuestion)) {
    return null;
  }

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var answer = safeCell_(row, headerMap, 'Answer').trim();
    if (!answer) continue;

    var candidates = [safeCell_(row, headerMap, 'Main Question')]
      .concat(splitAssistantTermsRaw_(safeCell_(row, headerMap, 'Alternate Questions')))
      .concat(splitAssistantTermsRaw_(safeCell_(row, headerMap, 'Keywords')));

    for (var j = 0; j < candidates.length; j++) {
      var candidate = normalizeAssistantText_(candidates[j]);
      if (!candidate) continue;

      if (candidate === cleanQuestion || extractAcronym_(candidate) === cleanQuestion) {
        return buildAssistantResult_(row, headerMap, 860);
      }
    }
  }

  return null;
}

function findAssistantPhraseMatch_(rows, headerMap, cleanQuestion) {
  if (!cleanQuestion || cleanQuestion.length < 4) {
    return null;
  }

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var answer = safeCell_(row, headerMap, 'Answer').trim();
    if (!answer) continue;

    var mainQuestion = normalizeAssistantText_(safeCell_(row, headerMap, 'Main Question'));
    var alternates = splitAssistantTerms_(safeCell_(row, headerMap, 'Alternate Questions'));
    var keywords = splitAssistantTerms_(safeCell_(row, headerMap, 'Keywords'));

    if (mainQuestion && (mainQuestion.indexOf(cleanQuestion) === 0 || cleanQuestion.indexOf(mainQuestion) === 0)) {
      return buildAssistantResult_(row, headerMap, 700);
    }

    for (var j = 0; j < alternates.length; j++) {
      var alt = alternates[j];
      if (alt.indexOf(cleanQuestion) === 0 || cleanQuestion.indexOf(alt) === 0) {
        return buildAssistantResult_(row, headerMap, 680);
      }
    }

    for (var k = 0; k < keywords.length; k++) {
      var keyword = keywords[k];
      if (keyword.indexOf(cleanQuestion) === 0 || cleanQuestion.indexOf(keyword) === 0) {
        return buildAssistantResult_(row, headerMap, 650);
      }
    }
  }

  return null;
}

function buildAssistantResult_(row, headerMap, score) {
  return {
    id: safeCell_(row, headerMap, 'ID'),
    mainQuestion: safeCell_(row, headerMap, 'Main Question'),
    answer: safeCell_(row, headerMap, 'Answer').trim(),
    links: extractAssistantLinks_(row, headerMap),
    score: score || 0
  };
}

function splitAssistantTerms_(value) {
  return String(value || '')
    .split(/[\n,;|]/)
    .map(function(part) {
      return normalizeAssistantText_(part);
    })
    .filter(Boolean);
}

function splitAssistantTermsRaw_(value) {
  return String(value || '')
    .split(/[\n,;|]/)
    .map(function(part) {
      return String(part || '').trim();
    })
    .filter(Boolean);
}

function isLikelyAcronym_(value) {
  var raw = String(value || '').replace(/\s+/g, '').trim();
  return /^[a-z0-9]{2,6}$/i.test(raw);
}

function extractAcronym_(text) {
  var tokens = normalizeAssistantText_(text).split(' ').filter(Boolean);

  if (!tokens.length) {
    return '';
  }

  if (tokens.length === 1) {
    return tokens[0];
  }

  return tokens.map(function(token) {
    return token.charAt(0);
  }).join('');
}

function scoreAssistantAlternateQuestions_(questionText, alternateQuestionsText, weight) {
  if (!questionText || !alternateQuestionsText) {
    return 0;
  }

  const alternates = splitAssistantTerms_(alternateQuestionsText);
  let score = 0;

  alternates.forEach(function(alt) {
    score += scoreAssistantField_(questionText, alt, weight);
  });

  return score;
}

function extractAssistantLinks_(row, headerMap) {
  const links = [];

  const title1 = safeCell_(row, headerMap, 'Link Title 1').trim();
  const url1 = safeCell_(row, headerMap, 'Link URL 1').trim();
  const title2 = safeCell_(row, headerMap, 'Link Title 2').trim();
  const url2 = safeCell_(row, headerMap, 'Link URL 2').trim();

  if (title1 && url1) links.push({ title: title1, url: url1 });
  if (title2 && url2) links.push({ title: title2, url: url2 });

  return links;
}

function buildAssistantHeaderMap_(headers) {
  const map = {};

  map['ID'] = findAssistantHeaderIndex_(headers, [
    'ID',
    'Id',
    'Row ID',
    'RowID'
  ]);

  map['Main Question'] = findAssistantHeaderIndex_(headers, [
    'Main Question',
    'Question',
    'Primary Question',
    'User Question'
  ]);

  map['Alternate Questions'] = findAssistantHeaderIndex_(headers, [
    'Alternate Questions',
    'Alternate Question',
    'Alt Questions',
    'Alternative Questions',
    'Variations',
    'Question Variations'
  ]);

  map['Keywords'] = findAssistantHeaderIndex_(headers, [
    'Keywords',
    'Keyword',
    'Tags',
    'Key Terms'
  ]);

  map['Answer'] = findAssistantHeaderIndex_(headers, [
    'Answer',
    'Response',
    'Grounded Answer',
    'Bot Answer'
  ]);

  map['Link Title 1'] = findAssistantHeaderIndex_(headers, [
    'Link Title 1',
    'Link 1 Title',
    'Resource Title 1',
    'CTA Title 1'
  ], true);

  map['Link URL 1'] = findAssistantHeaderIndex_(headers, [
    'Link URL 1',
    'Link 1 URL',
    'Resource URL 1',
    'CTA URL 1'
  ], true);

  map['Link Title 2'] = findAssistantHeaderIndex_(headers, [
    'Link Title 2',
    'Link 2 Title',
    'Resource Title 2',
    'CTA Title 2'
  ], true);

  map['Link URL 2'] = findAssistantHeaderIndex_(headers, [
    'Link URL 2',
    'Link 2 URL',
    'Resource URL 2',
    'CTA URL 2'
  ], true);

  map['Last Updated'] = findAssistantHeaderIndex_(headers, [
    'Last Updated',
    'Updated At',
    'Last Modified',
    'Modified At'
  ], true);

  if (map['Main Question'] === -1) {
    throw new Error('Missing required column in AI_Chatbot_Knowledge: Main Question');
  }

  if (map['Answer'] === -1) {
    throw new Error('Missing required column in AI_Chatbot_Knowledge: Answer');
  }

  if (map['ID'] === -1) {
    map['ID'] = '';
  }

  if (map['Alternate Questions'] === -1) {
    map['Alternate Questions'] = '';
  }

  if (map['Keywords'] === -1) {
    map['Keywords'] = '';
  }

  return map;
}
function debugAssistantSheet() {
  const ss = SpreadsheetApp.openById(KNOWLEDGE_SPREADSHEET_ID);
  const sh = ss.getSheetByName(ASSISTANT_SHEET_NAME);

  if (!sh) {
    return '❌ Sheet not found';
  }

  const data = sh.getDataRange().getValues();

  return {
    sheetName: sh.getName(),
    totalRows: data.length,
    headers: data[0],
    firstRow: data[1]
  };
}
function safeCell_(row, headerMap, headerName) {
  const index = headerMap[headerName];

  if (index === undefined || index === null || index === '' || index < 0) {
    return '';
  }

  return String(row[index] || '');
}
function findAssistantHeaderIndex_(headers, aliases, optional) {
  const normalizedHeaders = (headers || []).map(function(header) {
    return normalizeAssistantText_(header);
  });

  for (var i = 0; i < aliases.length; i++) {
    var target = normalizeAssistantText_(aliases[i]);
    var index = normalizedHeaders.indexOf(target);

    if (index !== -1) {
      return index;
    }
  }

  return optional ? -1 : -1;
}

function normalizeAssistantText_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/&/g, ' and ')
    .replace(/[^\w\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function tokenizeAssistantText_(value) {
  const stopWords = {
    a: true, an: true, and: true, are: true, as: true, at: true, be: true, by: true,
    do: true, for: true, from: true, how: true, i: true, in: true, is: true, it: true,
    of: true, on: true, or: true, the: true, to: true, what: true, when: true,
    where: true, which: true, who: true, why: true, with: true, your: true
  };

  return normalizeAssistantText_(value)
    .split(' ')
    .filter(function(token) {
      return token && token.length >= 2 && !stopWords[token];
    });
}

function scoreAssistantField_(questionText, fieldText, weight) {
  if (!questionText || !fieldText) {
    return 0;
  }

  const qTokens = tokenizeAssistantText_(questionText);
  const fTokens = tokenizeAssistantText_(fieldText);

  if (!qTokens.length || !fTokens.length) {
    return 0;
  }

  let score = 0;
  const qJoined = qTokens.join(' ');
  const fJoined = fTokens.join(' ');

  if (fJoined === qJoined) {
    score += weight + 12;
  } else if (fJoined.indexOf(qJoined) !== -1 && qJoined.length >= 5) {
    score += weight + 4;
  } else if (qJoined.indexOf(fJoined) !== -1 && fJoined.length >= 5) {
    score += weight + 2;
  }

  let overlap = 0;
  qTokens.forEach(function(token) {
    if (fTokens.indexOf(token) !== -1) overlap++;
  });

  if (overlap === qTokens.length && overlap === fTokens.length) {
    score += overlap * weight + 8;
  } else if (overlap === qTokens.length) {
    score += overlap * weight + 4;
  } else if (overlap >= 2) {
    score += overlap * (weight - 1);
  } else if (overlap === 1 && qTokens.length === 1 && fTokens.length === 1) {
    score += weight;
  }

  return score;
}

function scoreAssistantKeywords_(questionText, keywordsText, weight) {
  if (!questionText || !keywordsText) {
    return 0;
  }

  const qTokens = tokenizeAssistantText_(questionText);
  if (!qTokens.length) {
    return 0;
  }

  const keywordParts = splitAssistantTerms_(keywordsText);
  let score = 0;

  keywordParts.forEach(function(keyword) {
    const kTokens = tokenizeAssistantText_(keyword);
    if (!kTokens.length) return;

    if (keyword === questionText) {
      score += weight + 12;
      return;
    }

    if (keyword.length >= 5 && questionText.indexOf(keyword) !== -1) {
      score += weight + 4;
      return;
    }

    if (keyword.length >= 5 && keyword.indexOf(questionText) !== -1) {
      score += weight + 2;
      return;
    }

    let overlap = 0;
    qTokens.forEach(function(token) {
      if (kTokens.indexOf(token) !== -1) overlap++;
    });

    if (overlap >= 2) {
      score += overlap * (weight - 1);
    } else if (overlap === 1 && qTokens.length === 1) {
      score += weight;
    }
  });

  return score;
}

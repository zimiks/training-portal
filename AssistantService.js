const KNOWLEDGE_SPREADSHEET_ID = '11hc9yxH9F6P8SMQ62ry42WgPYmOPuSOj02U18XD9R1M';
const ASSISTANT_SHEET_NAME = 'AI_Chatbot_Knowledge';

/**
 * Grounded assistant backend
 * Reads only from AI_Chatbot_Knowledge
 */
function askAssistant(question) {
  try {
    const cleanQuestion = normalizeAssistantText_(question);

    if (!cleanQuestion) {
      return {
        ok: false,
        answer: 'Please type a question so I can help.',
        links: []
      };
    }

    const match = findBestAssistantAnswer_(cleanQuestion);

    if (!match) {
      return {
        ok: true,
        answer: 'I could not find a grounded answer in AI_Chatbot_Knowledge for that question. Please try rephrasing it or update the knowledge sheet.',
        links: []
      };
    }

    return {
      ok: true,
      answer: match.answer,
      links: match.links,
      matchedQuestion: match.mainQuestion,
      rowId: match.id || ''
    };
  } catch (error) {
    Logger.log('askAssistant error: ' + error);
    return {
      ok: false,
      answer: 'Something went wrong while checking AI_Chatbot_Knowledge. Please try again.',
      links: []
    };
  }
}

function getAssistantSpreadsheet_() {
  return SpreadsheetApp.openById('11hc9yxH9F6P8SMQ62ry42WgPYmOPuSOj02U18XD9R1M');
}

function findBestAssistantAnswer_(cleanQuestion) {
  const spreadsheet = getAssistantSpreadsheet_();
  const sheet = spreadsheet.getSheetByName(ASSISTANT_SHEET_NAME);

  if (!sheet) {
    throw new Error('Sheet not found: ' + ASSISTANT_SHEET_NAME);
  }

  const values = sheet.getDataRange().getDisplayValues();
  if (!values || values.length < 2) {
    return null;
  }

  const headers = values[0];
  const rows = values.slice(1);
  const headerMap = buildAssistantHeaderMap_(headers);

  // 1. Exact full match
  const exactMatch = findExactAssistantMatch_(rows, headerMap, cleanQuestion);
  if (exactMatch) {
    return exactMatch;
  }

  // 2. Acronym-first match
  const acronymMatch = findAssistantAcronymMatch_(rows, headerMap, cleanQuestion);
  if (acronymMatch) {
    return acronymMatch;
  }

  // 3. Phrase-first match
  const phraseMatch = findAssistantPhraseMatch_(rows, headerMap, cleanQuestion);
  if (phraseMatch) {
    return phraseMatch;
  }

  // 4. Scored fallback
  let bestRow = null;
  let bestScore = 0;

  rows.forEach(function(row) {
    const answer = safeCell_(row, headerMap, 'Answer').trim();
    if (!answer) return;

    const mainQuestion = normalizeAssistantText_(safeCell_(row, headerMap, 'Main Question'));
    const alternateQuestions = safeCell_(row, headerMap, 'Alternate Questions');
    const keywords = safeCell_(row, headerMap, 'Keywords');

    let score = 0;
    score += scoreAssistantField_(cleanQuestion, mainQuestion, 8);
    score += scoreAssistantAlternateQuestions_(cleanQuestion, alternateQuestions, 7);
    score += scoreAssistantKeywords_(cleanQuestion, keywords, 6);

    if (score > bestScore) {
      bestScore = score;
      bestRow = row;
    }
  });

  if (!bestRow || bestScore < 8) {
    return null;
  }

  return buildAssistantResult_(bestRow, headerMap, bestScore);
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

    var mainQuestionRaw = safeCell_(row, headerMap, 'Main Question');
    var alternateRaw = safeCell_(row, headerMap, 'Alternate Questions');
    var keywordsRaw = safeCell_(row, headerMap, 'Keywords');

    var candidates = [mainQuestionRaw]
      .concat(splitAssistantTermsRaw_(alternateRaw))
      .concat(splitAssistantTermsRaw_(keywordsRaw));

    for (var j = 0; j < candidates.length; j++) {
      var candidate = normalizeAssistantText_(candidates[j]);
      if (!candidate) continue;

      if (candidate === cleanQuestion) {
        return buildAssistantResult_(row, headerMap, 880);
      }

      if (extractAcronym_(candidate) === cleanQuestion) {
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

    if (mainQuestion && (
      mainQuestion.indexOf(cleanQuestion) === 0 ||
      cleanQuestion.indexOf(mainQuestion) === 0
    )) {
      return buildAssistantResult_(row, headerMap, 700);
    }

    for (var j = 0; j < alternates.length; j++) {
      var alt = alternates[j];
      if (
        alt.indexOf(cleanQuestion) === 0 ||
        cleanQuestion.indexOf(alt) === 0
      ) {
        return buildAssistantResult_(row, headerMap, 680);
      }
    }

    for (var k = 0; k < keywords.length; k++) {
      var keyword = keywords[k];
      if (
        keyword.indexOf(cleanQuestion) === 0 ||
        cleanQuestion.indexOf(keyword) === 0
      ) {
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
  var tokens = normalizeAssistantText_(text)
    .split(' ')
    .filter(Boolean);

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
function findExactAssistantMatch_(rows, headerMap, cleanQuestion) {
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var answer = safeCell_(row, headerMap, 'Answer').trim();
    if (!answer) continue;

    var mainQuestion = normalizeAssistantText_(safeCell_(row, headerMap, 'Main Question'));
    var alternateQuestionsRaw = safeCell_(row, headerMap, 'Alternate Questions');
    var keywordsRaw = safeCell_(row, headerMap, 'Keywords');

    // Exact main question match
    if (mainQuestion && mainQuestion === cleanQuestion) {
      return buildAssistantResult_(row, headerMap, 1000);
    }

    // Exact alternate question match
    var alternateParts = splitAssistantTerms_(alternateQuestionsRaw);
    for (var j = 0; j < alternateParts.length; j++) {
      if (alternateParts[j] === cleanQuestion) {
        return buildAssistantResult_(row, headerMap, 950);
      }
    }

    // Exact keyword match
    var keywordParts = splitAssistantTerms_(keywordsRaw);
    for (var k = 0; k < keywordParts.length; k++) {
      if (keywordParts[k] === cleanQuestion) {
        return buildAssistantResult_(row, headerMap, 900);
      }
    }
  }

  return null;
}
function splitAssistantTerms_(value) {
  return String(value || '')
    .split(/[\n,;|]/)
    .map(function(part) {
      return normalizeAssistantText_(part);
    })
    .filter(Boolean);
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
function extractAssistantLinks_(row, headerMap) {
  const links = [];

  const title1 = safeCell_(row, headerMap, 'Link Title 1').trim();
  const url1 = safeCell_(row, headerMap, 'Link URL 1').trim();
  const title2 = safeCell_(row, headerMap, 'Link Title 2').trim();
  const url2 = safeCell_(row, headerMap, 'Link URL 2').trim();

  if (title1 && url1) {
    links.push({ title: title1, url: url1 });
  }

  if (title2 && url2) {
    links.push({ title: title2, url: url2 });
  }

  return links;
}

function buildAssistantHeaderMap_(headers) {
  const requiredHeaders = [
    'ID',
    'Main Question',
    'Alternate Questions',
    'Keywords',
    'Answer',
    'Link Title 1',
    'Link URL 1',
    'Link Title 2',
    'Link URL 2',
    'Last Updated'
  ];

  const map = {};

  requiredHeaders.forEach(function(name) {
    const index = headers.indexOf(name);
    if (index === -1) {
      throw new Error('Missing required column in AI_Chatbot_Knowledge: ' + name);
    }
    map[name] = index;
  });

  return map;
}

function safeCell_(row, headerMap, headerName) {
  const index = headerMap[headerName];
  return index === undefined ? '' : String(row[index] || '');
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
    a: true,
    an: true,
    and: true,
    are: true,
    as: true,
    at: true,
    be: true,
    by: true,
    do: true,
    for: true,
    from: true,
    how: true,
    i: true,
    in: true,
    is: true,
    it: true,
    of: true,
    on: true,
    or: true,
    the: true,
    to: true,
    what: true,
    when: true,
    where: true,
    which: true,
    who: true,
    why: true,
    with: true,
    your: true
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
    if (fTokens.indexOf(token) !== -1) {
      overlap++;
    }
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
      if (kTokens.indexOf(token) !== -1) {
        overlap++;
      }
    });

    if (overlap === qTokens.length && overlap > 0) {
      score += overlap * weight + 4;
    } else if (overlap >= 2) {
      score += overlap * (weight - 1);
    } else if (overlap === 1 && qTokens.length === 1 && kTokens.length === 1) {
      score += weight;
    }
  });

  return score;
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

function testAskAssistant() {
  const tests = [
    'What is W2?',
    'What is ATS?',
    'What is a compact license?',
    'Explain PACU',
    'How does submission readiness work?'
  ];

  tests.forEach(function(question) {
    const result = askAssistant(question);
    Logger.log('QUESTION: ' + question);
    Logger.log(JSON.stringify(result, null, 2));
  });
}
const SHEET_NAMES = {
  products: '商品候補一覧',
  categories: 'カテゴリ設定',
  logs: '収集ログ',
  excludes: '除外リスト',
};

const PRODUCT_HEADERS = [
  '収集日',
  'カテゴリ',
  '検索キーワード',
  'ブランド名',
  '商品名',
  '商品概要',
  '商品画像URL',
  '公式URL',
  '商品ページURL',
  '販売ページURL',
  '海外販売価格',
  '通貨',
  '為替レート',
  '卸掛率',
  '仮仕入れ価格',
  '手入力仕入れ価格',
  '採用仕入れ価格',
  '商品原価円',
  '国際送料',
  '関税率',
  '関税額',
  '輸入消費税率',
  '輸入消費税額',
  '着地原価',
  '国内配送料',
  '販路手数料率',
  '販路手数料',
  '想定販売価格',
  '推奨販売価格',
  '利益額',
  '利益率',
  '利益判定',
  '仕入れ判断',
  'ハードル',
  'ハードル分類理由',
  '販売前確認事項',
  '注意フラグ',
  'ステータス',
  'メモ',
];

const CATEGORY_HEADERS = [
  'カテゴリ',
  '有効',
  '1日収集件数',
  '優先度',
  '検索キーワード',
  '除外キーワード',
];

const LOG_HEADERS = [
  '実行日時',
  '成功件数',
  '重複件数',
  'エラー件数',
  '検索キーワード',
  'メッセージ',
];

const EXCLUDE_HEADERS = ['除外キーワード', 'メモ'];

const STATUS_OPTIONS = [
  '未確認',
  '確認中',
  '深掘り対象',
  '問い合わせ候補',
  '保留',
  '除外',
  '日本販売あり',
  '規制リスク高',
  '価格確認待ち',
];

const DEFAULT_SETTINGS = {
  dailyLimit: 5,
  scheduleHour: 9,
  scheduleMinute: 5,
  fxRateUsd: 160,
  wholesaleRate: 0.45,
  internationalShipping: 500,
  dutyRate: 0,
  importConsumptionTaxRate: 0.10,
  domesticShipping: 800,
  ecFeeRate: 0.15,
  targetProfitRate: 0.20,
  recommendedProfitRate: 0.30,
  retailMarkup: 3.5,
};

const DEFAULT_CATEGORIES = [
  ['掃除道具', true, 1, 1, 'innovative cleaning tools D2C brand', 'used,replica,fake'],
  ['洗濯便利グッズ', true, 1, 1, 'laundry helper product brand', 'detergent,cream,lotion'],
  ['収納用品', true, 1, 2, 'home storage organizer product brand', 'used,replica,fake'],
  ['デスクワーク便利グッズ', true, 1, 2, 'desk work productivity gadget D2C', 'software,app'],
  ['ランニング小物', true, 1, 3, 'running accessory brand reflective pouch', 'supplement,food'],
];

const LOW_HURDLE_TERMS = [
  '掃除道具',
  '洗濯便利グッズ',
  '収納用品',
  'ファイリング用品',
  'ブラシ',
  'スポンジ',
  'クロス',
  '手袋',
  '作業用グローブ',
  '反射グッズ',
  'バッグ',
  'ポーチ',
  'アクセサリー',
  'デスクワーク便利グッズ',
  'ランニング小物',
  '肌に塗らない手荒れ対策グッズ',
  'cleaning tool',
  'cleaning brush',
  'sponge',
  'cloth',
  'glove',
  'storage',
  'organizer',
  'file organizer',
  'bag',
  'pouch',
  'accessory',
  'reflective',
  'running accessory',
  'desk work',
];

const MEDIUM_HURDLE_TERMS = [
  '洗剤',
  '洗濯用石けん',
  '台所用石けん',
  '住宅用洗浄剤',
  '漂白剤',
  '除菌スプレー',
  '消臭スプレー',
  '肌に長時間触れる粘着系商品',
  'インソール',
  '靴ずれ防止パッド',
  'シリコンパッド',
  'detergent',
  'laundry soap',
  'dish soap',
  'household cleaner',
  'bleach',
  'disinfecting spray',
  'deodorizing spray',
  'adhesive',
  'insole',
  'silicone pad',
  'heel protector',
];

const HIGH_HURDLE_TERMS = [
  '顔に塗るもの',
  '体に塗るもの',
  'ハンドクリーム',
  'フットクリーム',
  'スキンケア',
  '化粧品',
  '身体洗浄用石けん',
  '食品',
  'サプリ',
  '口に入れるもの',
  '医療機器っぽいもの',
  '電源を使う美容機器',
  '治る',
  '効く',
  '改善',
  'アトピー',
  '湿疹',
  'あかぎれ',
  '炎症',
  'face cream',
  'facial',
  'skincare',
  'skin care',
  'cream',
  'lotion',
  'serum',
  'soap',
  'food',
  'supplement',
  'edible',
  'medical device',
  'electric beauty',
  'eczema',
  'dermatitis',
  'heal',
  'cure',
  'treatment',
  'inflammation',
];

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const productsSheet = getOrCreateSheet_(ss, SHEET_NAMES.products);
  const categoriesSheet = getOrCreateSheet_(ss, SHEET_NAMES.categories);
  const logsSheet = getOrCreateSheet_(ss, SHEET_NAMES.logs);
  const excludesSheet = getOrCreateSheet_(ss, SHEET_NAMES.excludes);

  setupHeader_(productsSheet, PRODUCT_HEADERS);
  setupHeader_(categoriesSheet, CATEGORY_HEADERS);
  setupHeader_(logsSheet, LOG_HEADERS);
  setupHeader_(excludesSheet, EXCLUDE_HEADERS);

  if (categoriesSheet.getLastRow() < 2) {
    categoriesSheet.getRange(2, 1, DEFAULT_CATEGORIES.length, CATEGORY_HEADERS.length).setValues(DEFAULT_CATEGORIES);
  }

  applyProductSheetFormatting_(productsSheet);
  applyCategorySheetFormatting_(categoriesSheet);
  applyLogSheetFormatting_(logsSheet);
  applyExcludeSheetFormatting_(excludesSheet);
  setProfitFormulas();
}

function collectDailyProducts() {
  setupSheets();

  const settings = getSettings_();
  const categories = getEnabledCategories_();
  const excludeWords = getExcludeWords_();
  const targetLimit = settings.dailyLimit;
  let successCount = 0;
  let duplicateCount = 0;
  let errorCount = 0;
  const searchedKeywords = [];
  const messages = [];

  try {
    for (const category of categories) {
      if (successCount >= targetLimit) {
        break;
      }

      const remaining = targetLimit - successCount;
      const perCategoryLimit = Math.min(category.dailyLimit, remaining);
      if (perCategoryLimit <= 0) {
        continue;
      }

      const candidates = searchProductsBySerpApi(category.searchKeyword, category.category, perCategoryLimit, category.excludeKeywords.concat(excludeWords));
      searchedKeywords.push(category.searchKeyword);
      const result = appendProductCandidates(candidates, perCategoryLimit);
      successCount += result.added;
      duplicateCount += result.duplicates;
    }
  } catch (error) {
    errorCount += 1;
    messages.push(error.message || String(error));
  }

  logRun(successCount, duplicateCount, errorCount, searchedKeywords.join(', '), messages.join(' / ') || '実行完了');
}

function searchProductsBySerpApi(searchKeyword, category, limit, excludeKeywords) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('SERPAPI_API_KEY');
  if (!apiKey) {
    throw new Error('Script Properties に SERPAPI_API_KEY が設定されていません。');
  }

  const queryParts = [searchKeyword, 'product official brand'];
  if (excludeKeywords && excludeKeywords.length) {
    excludeKeywords.filter(Boolean).forEach((word) => queryParts.push('-' + word));
  }

  const params = {
    engine: 'google',
    q: queryParts.join(' '),
    api_key: apiKey,
    num: Math.min(10, Math.max(5, limit * 2)),
    hl: 'en',
    gl: 'us',
  };
  const url = 'https://serpapi.com/search.json?' + Object.keys(params).map((key) => {
    return encodeURIComponent(key) + '=' + encodeURIComponent(params[key]);
  }).join('&');

  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    muteHttpExceptions: true,
  });
  const statusCode = response.getResponseCode();
  if (statusCode < 200 || statusCode >= 300) {
    throw new Error('SerpAPI request failed: HTTP ' + statusCode + ' ' + response.getContentText().slice(0, 300));
  }

  const payload = JSON.parse(response.getContentText());
  const organicResults = payload.organic_results || [];
  const candidates = [];

  for (const result of organicResults) {
    if (candidates.length >= limit) {
      break;
    }

    const title = result.title || '';
    const link = result.link || '';
    const snippet = result.snippet || '';
    if (!title || !link || isExcluded_(title + ' ' + snippet + ' ' + link, excludeKeywords)) {
      continue;
    }

    const priceInfo = extractPrice_(result);
    const brandAndProduct = splitBrandAndProduct_(title);
    const imageUrl = extractImageUrl_(result);
    const hurdle = classifyHurdle(category, brandAndProduct.productName, snippet, searchKeyword);

    candidates.push({
      collectedDate: new Date(),
      category: category,
      searchKeyword: searchKeyword,
      brandName: brandAndProduct.brandName,
      productName: brandAndProduct.productName,
      summary: snippet,
      imageUrl: imageUrl,
      officialUrl: getSiteRoot_(link),
      productUrl: link,
      salesUrl: link,
      overseasPrice: priceInfo.price,
      currency: priceInfo.currency || 'USD',
      fxRate: DEFAULT_SETTINGS.fxRateUsd,
      wholesaleRate: DEFAULT_SETTINGS.wholesaleRate,
      internationalShipping: DEFAULT_SETTINGS.internationalShipping,
      dutyRate: DEFAULT_SETTINGS.dutyRate,
      importConsumptionTaxRate: DEFAULT_SETTINGS.importConsumptionTaxRate,
      domesticShipping: DEFAULT_SETTINGS.domesticShipping,
      ecFeeRate: DEFAULT_SETTINGS.ecFeeRate,
      hurdle: hurdle.level,
      hurdleReason: hurdle.reason,
      preSaleChecks: hurdle.checks,
      warningFlag: hurdle.warningFlag,
      status: priceInfo.price ? '未確認' : '価格確認待ち',
      memo: '',
    });
  }

  return candidates;
}

function appendProductCandidates(candidates, limit) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.products);
  const headers = getHeaderMap_(sheet);
  const existingKeys = getExistingProductKeys_(sheet, headers);
  const rows = [];
  let added = 0;
  let duplicates = 0;

  for (const candidate of candidates) {
    if (added >= limit) {
      break;
    }

    const key = makeProductKey_(candidate);
    if (existingKeys.has(key)) {
      duplicates += 1;
      continue;
    }

    rows.push(buildProductRow_(candidate));
    existingKeys.add(key);
    added += 1;
  }

  if (rows.length) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rows.length, PRODUCT_HEADERS.length).setValues(rows);
    for (let offset = 0; offset < rows.length; offset += 1) {
      setProfitFormulas(startRow + offset);
    }
    applyProductSheetFormatting_(sheet);
  }

  return {added: added, duplicates: duplicates};
}

function classifyHurdle(category, productName, summary, searchKeyword) {
  const text = [category, productName, summary, searchKeyword].join(' ').toLowerCase();
  const highMatches = findMatches_(text, HIGH_HURDLE_TERMS);
  if (highMatches.length) {
    return {
      level: '高',
      reason: '高ハードル語句に一致: ' + highMatches.join(', '),
      checks: '薬機法確認, 食品衛生法確認, PSE確認, 医療機器該当性確認, 広告表現注意, 初回仕入れ非推奨',
      warningFlag: highMatches.join(', '),
    };
  }

  const mediumMatches = findMatches_(text, MEDIUM_HURDLE_TERMS);
  if (mediumMatches.length) {
    return {
      level: '中',
      reason: '中ハードル語句に一致: ' + mediumMatches.join(', '),
      checks: '成分表示確認, 日本語ラベル確認, 安全表示確認, 広告表現確認',
      warningFlag: mediumMatches.join(', '),
    };
  }

  const lowMatches = findMatches_(text, LOW_HURDLE_TERMS);
  return {
    level: '低',
    reason: lowMatches.length ? '低ハードル語句に一致: ' + lowMatches.join(', ') : '高・中ハードル語句に一致なし',
    checks: '通常の輸入・表示・安全確認',
    warningFlag: '',
  };
}

function setProfitFormulas(targetRow) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.products);
  if (!sheet) {
    return;
  }

  const lastRow = targetRow || sheet.getLastRow();
  if (lastRow < 2) {
    return;
  }

  const startRow = targetRow || 2;
  for (let row = startRow; row <= lastRow; row += 1) {
    sheet.getRange(row, 15).setFormula(`=IF(K${row}="","",K${row}*N${row})`);
    sheet.getRange(row, 17).setFormula(`=IF(P${row}<>"",P${row},O${row})`);
    sheet.getRange(row, 18).setFormula(`=IF(Q${row}="","",Q${row}*M${row})`);
    sheet.getRange(row, 21).setFormula(`=IF(R${row}="","",(R${row}+S${row})*T${row})`);
    sheet.getRange(row, 23).setFormula(`=IF(R${row}="","",(R${row}+S${row}+U${row})*V${row})`);
    sheet.getRange(row, 24).setFormula(`=IF(R${row}="","",R${row}+S${row}+U${row}+W${row})`);
    sheet.getRange(row, 27).setFormula(`=IF(AB${row}<>"",AB${row}*Z${row},AC${row}*Z${row})`);
    sheet.getRange(row, 29).setFormula(`=IF(R${row}="","",CEILING(MAX(R${row}*${DEFAULT_SETTINGS.retailMarkup},(X${row}+Y${row})/(1-Z${row}-${DEFAULT_SETTINGS.recommendedProfitRate})),100))`);
    sheet.getRange(row, 30).setFormula(`=IF(R${row}="","",IF(AB${row}<>"",AB${row},AC${row})-X${row}-Y${row}-AA${row})`);
    sheet.getRange(row, 31).setFormula(`=IF(OR(R${row}="",IF(AB${row}<>"",AB${row},AC${row})=0),"",AD${row}/IF(AB${row}<>"",AB${row},AC${row}))`);
    sheet.getRange(row, 32).setFormula(`=IF(AE${row}="","",IFS(AE${row}>=0.3,"◎優秀",AE${row}>=0.2,"○候補",AE${row}>=0.1,"△要改善",AE${row}>=0,"▲低利益",AE${row}<0,"✕赤字リスク"))`);
    sheet.getRange(row, 33).setFormula(`=IFS(AL${row}="価格確認待ち","価格確認待ち",AE${row}<0,"除外候補",AH${row}="高","初回非推奨",AH${row}="中","要確認",AND(AH${row}="低",AE${row}>=0.3),"有力候補",AND(AH${row}="低",AE${row}>=0.2),"候補",TRUE,"要確認")`);
  }
}

function createDailyTrigger() {
  deleteTriggers();
  ScriptApp.newTrigger('collectDailyProducts')
    .timeBased()
    .atHour(DEFAULT_SETTINGS.scheduleHour)
    .nearMinute(DEFAULT_SETTINGS.scheduleMinute)
    .everyDays(1)
    .inTimezone('Asia/Tokyo')
    .create();
}

function deleteTriggers() {
  ScriptApp.getProjectTriggers().forEach((trigger) => {
    if (trigger.getHandlerFunction() === 'collectDailyProducts') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function logRun(successCount, duplicateCount, errorCount, searchKeyword, message) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet_(ss, SHEET_NAMES.logs);
  setupHeader_(sheet, LOG_HEADERS);
  sheet.appendRow([
    new Date(),
    successCount || 0,
    duplicateCount || 0,
    errorCount || 0,
    searchKeyword || '',
    message || '',
  ]);
  applyLogSheetFormatting_(sheet);
}

function manualCollectDailyProducts() {
  collectDailyProducts();
}

function setSerpApiKey(apiKey) {
  if (!apiKey) {
    throw new Error('apiKey を指定してください。');
  }
  PropertiesService.getScriptProperties().setProperty('SERPAPI_API_KEY', apiKey);
}

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function setupHeader_(sheet, headers) {
  const current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const needsHeader = current.join('') === '' || headers.some((header, index) => current[index] !== header);
  if (needsHeader) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#1f4e78')
    .setFontColor('#ffffff');
  if (!sheet.getFilter()) {
    sheet.getRange(1, 1, Math.max(2, sheet.getMaxRows()), headers.length).createFilter();
  }
}

function applyProductSheetFormatting_(sheet) {
  setupHeader_(sheet, PRODUCT_HEADERS);
  sheet.autoResizeColumns(1, PRODUCT_HEADERS.length);
  sheet.setColumnWidths(6, 4, 220);
  sheet.setColumnWidths(30, 4, 110);
  sheet.getRange(2, 11, Math.max(1, sheet.getMaxRows() - 1), 1).setNumberFormat('$#,##0.00');
  sheet.getRange(2, 13, Math.max(1, sheet.getMaxRows() - 1), 1).setNumberFormat('#,##0');
  sheet.getRange(2, 14, Math.max(1, sheet.getMaxRows() - 1), 1).setNumberFormat('0%');
  sheet.getRange(2, 18, Math.max(1, sheet.getMaxRows() - 1), 1).setNumberFormat('#,##0');
  sheet.getRange(2, 19, Math.max(1, sheet.getMaxRows() - 1), 1).setNumberFormat('#,##0');
  sheet.getRange(2, 20, Math.max(1, sheet.getMaxRows() - 1), 1).setNumberFormat('0%');
  sheet.getRange(2, 22, Math.max(1, sheet.getMaxRows() - 1), 1).setNumberFormat('0%');
  sheet.getRange(2, 24, Math.max(1, sheet.getMaxRows() - 1), 1).setNumberFormat('#,##0');
  sheet.getRange(2, 26, Math.max(1, sheet.getMaxRows() - 1), 1).setNumberFormat('0%');
  sheet.getRange(2, 27, Math.max(1, sheet.getMaxRows() - 1), 4).setNumberFormat('#,##0');
  sheet.getRange(2, 31, Math.max(1, sheet.getMaxRows() - 1), 1).setNumberFormat('0.0%');

  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(STATUS_OPTIONS, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 38, Math.max(1, sheet.getMaxRows() - 1), 1).setDataValidation(statusRule);

  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('低')
      .setBackground('#d9ead3')
      .setRanges([sheet.getRange(2, 34, Math.max(1, sheet.getMaxRows() - 1), 1)])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('中')
      .setBackground('#fff2cc')
      .setRanges([sheet.getRange(2, 34, Math.max(1, sheet.getMaxRows() - 1), 1)])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('高')
      .setBackground('#f4cccc')
      .setRanges([sheet.getRange(2, 34, Math.max(1, sheet.getMaxRows() - 1), 1)])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(0.3)
      .setBackground('#d9ead3')
      .setRanges([sheet.getRange(2, 31, Math.max(1, sheet.getMaxRows() - 1), 1)])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(0.2, 0.299999)
      .setBackground('#fff2cc')
      .setRanges([sheet.getRange(2, 31, Math.max(1, sheet.getMaxRows() - 1), 1)])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setBackground('#f4cccc')
      .setRanges([sheet.getRange(2, 31, Math.max(1, sheet.getMaxRows() - 1), 1)])
      .build(),
  ];
  sheet.setConditionalFormatRules(rules);
}

function applyCategorySheetFormatting_(sheet) {
  setupHeader_(sheet, CATEGORY_HEADERS);
  sheet.autoResizeColumns(1, CATEGORY_HEADERS.length);
  const boolRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['TRUE', 'FALSE'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 2, Math.max(1, sheet.getMaxRows() - 1), 1).setDataValidation(boolRule);
}

function applyLogSheetFormatting_(sheet) {
  setupHeader_(sheet, LOG_HEADERS);
  sheet.autoResizeColumns(1, LOG_HEADERS.length);
}

function applyExcludeSheetFormatting_(sheet) {
  setupHeader_(sheet, EXCLUDE_HEADERS);
  sheet.autoResizeColumns(1, EXCLUDE_HEADERS.length);
}

function getSettings_() {
  return DEFAULT_SETTINGS;
}

function getEnabledCategories_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.categories);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }

  return sheet.getRange(2, 1, lastRow - 1, CATEGORY_HEADERS.length).getValues()
    .filter((row) => row[0] && row[1] !== false && String(row[1]).toUpperCase() !== 'FALSE')
    .map((row) => ({
      category: String(row[0]),
      dailyLimit: Number(row[2]) || 1,
      priority: Number(row[3]) || 999,
      searchKeyword: String(row[4] || row[0]),
      excludeKeywords: splitWords_(row[5]),
    }))
    .sort((a, b) => a.priority - b.priority);
}

function getExcludeWords_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.excludes);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }
  return sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().filter(Boolean).map(String);
}

function getHeaderMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((header, index) => {
    if (header) {
      map[header] = index + 1;
    }
  });
  return map;
}

function getExistingProductKeys_(sheet, headers) {
  const keys = new Set();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return keys;
  }

  const values = sheet.getRange(2, 1, lastRow - 1, PRODUCT_HEADERS.length).getDisplayValues();
  values.forEach((row) => {
    const officialUrl = normalizeUrl_(row[headers['公式URL'] - 1]);
    const productUrl = normalizeUrl_(row[headers['商品ページURL'] - 1]);
    const brandProduct = normalizeText_(row[headers['ブランド名'] - 1] + ' ' + row[headers['商品名'] - 1]);
    keys.add([officialUrl, productUrl, brandProduct].join('|'));
  });
  return keys;
}

function makeProductKey_(candidate) {
  return [
    normalizeUrl_(candidate.officialUrl),
    normalizeUrl_(candidate.productUrl),
    normalizeText_(candidate.brandName + ' ' + candidate.productName),
  ].join('|');
}

function buildProductRow_(candidate) {
  return [
    candidate.collectedDate,
    candidate.category,
    candidate.searchKeyword,
    candidate.brandName,
    candidate.productName,
    candidate.summary,
    hyperlinkFormula_(candidate.imageUrl),
    hyperlinkFormula_(candidate.officialUrl),
    hyperlinkFormula_(candidate.productUrl),
    hyperlinkFormula_(candidate.salesUrl),
    candidate.overseasPrice || '',
    candidate.currency || 'USD',
    candidate.fxRate,
    candidate.wholesaleRate,
    '',
    '',
    '',
    '',
    candidate.internationalShipping,
    candidate.dutyRate,
    '',
    candidate.importConsumptionTaxRate,
    '',
    '',
    candidate.domesticShipping,
    candidate.ecFeeRate,
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    candidate.hurdle,
    candidate.hurdleReason,
    candidate.preSaleChecks,
    candidate.warningFlag,
    candidate.status,
    candidate.memo,
  ];
}

function extractPrice_(result) {
  const candidates = [
    result.rich_snippet && result.rich_snippet.top && result.rich_snippet.top.detected_extensions && result.rich_snippet.top.detected_extensions.price,
    result.rich_snippet && result.rich_snippet.bottom && result.rich_snippet.bottom.detected_extensions && result.rich_snippet.bottom.detected_extensions.price,
    result.price,
    result.extracted_price,
    result.snippet,
  ];

  for (const value of candidates) {
    if (typeof value === 'number' && value > 0) {
      return {price: value, currency: 'USD'};
    }
    if (typeof value === 'string') {
      const match = value.match(/([$€£])\s?([0-9][0-9,]*(?:\.[0-9]+)?)/);
      if (match) {
        const currency = match[1] === '$' ? 'USD' : match[1] === '€' ? 'EUR' : 'GBP';
        return {price: Number(match[2].replace(/,/g, '')), currency: currency};
      }
    }
  }
  return {price: '', currency: 'USD'};
}

function extractImageUrl_(result) {
  return result.thumbnail || result.image || '';
}

function splitBrandAndProduct_(title) {
  const cleaned = String(title).replace(/\s+/g, ' ').trim();
  const parts = cleaned.split(/\s[-|–:]\s/).filter(Boolean);
  if (parts.length >= 2) {
    return {
      productName: parts[0].slice(0, 120),
      brandName: parts[parts.length - 1].slice(0, 80),
    };
  }
  const words = cleaned.split(' ');
  return {
    brandName: words.slice(0, 2).join(' ').slice(0, 80) || 'Unknown',
    productName: cleaned.slice(0, 120),
  };
}

function getSiteRoot_(url) {
  const match = String(url || '').match(/^(https?:\/\/[^/]+)\/?/i);
  return match ? match[1] + '/' : (url || '');
}

function isExcluded_(text, excludeKeywords) {
  const haystack = String(text || '').toLowerCase();
  return (excludeKeywords || []).some((word) => word && haystack.includes(String(word).toLowerCase()));
}

function findMatches_(text, terms) {
  const haystack = String(text || '').toLowerCase();
  return terms.filter((term) => haystack.includes(String(term).toLowerCase())).slice(0, 5);
}

function splitWords_(value) {
  if (!value) {
    return [];
  }
  return String(value).split(/[,\n、]/).map((word) => word.trim()).filter(Boolean);
}

function normalizeUrl_(url) {
  return String(url || '')
    .replace(/^=HYPERLINK\("([^"]+)".*$/i, '$1')
    .trim()
    .replace(/[?#].*$/, '')
    .replace(/\/$/, '')
    .toLowerCase();
}

function normalizeText_(value) {
  return String(value || '').trim().replace(/\s+/g, ' ').toLowerCase();
}

function hyperlinkFormula_(url) {
  if (!url) {
    return '';
  }
  const escaped = String(url).replace(/"/g, '""');
  return '=HYPERLINK("' + escaped + '","' + escaped + '")';
}

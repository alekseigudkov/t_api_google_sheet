/**
 * Выгрузка счетов и позиций Т-Инвестиций (Invest API 2.0) в Google Sheets.
 */

const TINKOFF_API_BASE = 'https://invest-public-api.tinkoff.ru/rest';
const AUTO_SYNC_PROPERTY = 'AUTO_SYNC_ON_OPEN';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Т-Инвестиции')
    .addItem('Обновить данные', 'menuRefreshData_')
    .addToUi();

  const autoSyncEnabled =
    String(PropertiesService.getScriptProperties().getProperty(AUTO_SYNC_PROPERTY) || '').toLowerCase() === 'true';

  if (autoSyncEnabled) {
    menuRefreshData_();
  }
}

function menuRefreshData_() {
  syncTinkoffToSheet();
}

function syncTinkoffToSheet() {
  const token = PropertiesService.getScriptProperties().getProperty('TINKOFF_TOKEN');
  if (!token) {
    throw new Error('Не задан Script Property: TINKOFF_TOKEN');
  }

  const futuresMeta = loadFuturesMeta_(token);
  const fullNameCache = {};

  const accountsResponse = callTinkoffApi_(
    '/tinkoff.public.invest.api.contract.v1.UsersService/GetAccounts',
    {},
    token
  );

  const accounts = (accountsResponse.accounts || []).map((acc) => ({
    accountId: acc.id || '',
    name: acc.name || '',
    type: normalizeEnum_(acc.type),
    status: normalizeEnum_(acc.status),
    openedDate: toIsoDate_(acc.openedDate),
    closedDate: toIsoDate_(acc.closedDate),
    accessLevel: normalizeEnum_(acc.accessLevel),
  }));

  const positionRows = [];
  accounts.forEach((account) => {
    const positionsResponse = callTinkoffApi_(
      '/tinkoff.public.invest.api.contract.v1.OperationsService/GetPositions',
      { accountId: account.accountId },
      token
    );

    const blockedIndex = buildBlockedIndex_(positionsResponse);

    (positionsResponse.money || []).forEach((money) => {
      positionRows.push([
        account.accountId,
        account.name,
        'money',
        '',
        'Валюта ' + (money.currency || ''),
        'Валюта ' + (money.currency || ''),
        'currency',
        moneyToNumber_(money),
        0,
        0,
        0,
        0,
        money.currency || '',
        null,
        null,
        null,
        null,
      ]);
    });

    const portfolioResponse = callTinkoffApi_(
      '/tinkoff.public.invest.api.contract.v1.OperationsService/GetPortfolio',
      { accountId: account.accountId },
      token
    );

    (portfolioResponse.positions || []).forEach((position) => {
      const instrumentType = (position.instrumentType || '').toLowerCase();
      const meta =
        futuresMeta.byFigi[position.figi || ''] ||
        futuresMeta.byUid[position.instrumentUid || ''] ||
        null;

      const lot = instrumentType === 'futures' ? numberOrNull_(meta?.lot) : null;
      const minPriceIncrementRaw = quotationToNumber_(meta?.minPriceIncrement);
      const minPriceIncrement = instrumentType === 'futures' ? numberOrNull_(minPriceIncrementRaw) : null;

      const minPriceIncrementAmountRaw = moneyToNumber_(meta?.minPriceIncrementAmount);
      const pointPriceRaw = safeDivide_(minPriceIncrementAmountRaw, minPriceIncrementRaw);
      const pointPrice = instrumentType === 'futures' ? numberOrNull_(pointPriceRaw) : null;

      const qty = parseNumericField_(position.quantity);
      const currentPrice = parseNumericField_(position.currentPrice);
      const currentPositionValue = computeCurrentPositionValue_(instrumentType, qty, currentPrice, pointPrice);

      const blockedValue = resolveBlocked_(position, blockedIndex);

      const shortName = readInstrumentName_(position);
      const fullName = resolveInstrumentFullName_(position, token, futuresMeta, fullNameCache) || shortName;

      positionRows.push([
        account.accountId,
        account.name,
        'portfolio',
        position.figi || '',
        shortName,
        fullName,
        position.instrumentType || '',
        qty,
        blockedValue,
        parseNumericField_(position.averagePositionPrice),
        parseNumericField_(position.expectedYield),
        currentPrice,
        readCurrency_(position),
        lot,
        minPriceIncrement,
        pointPrice,
        numberOrNull_(currentPositionValue),
      ]);
    });
  });

  const accountRows = accounts.map((a) => [
    a.accountId,
    a.name,
    a.type,
    a.status,
    a.openedDate,
    a.closedDate,
    a.accessLevel,
  ]);

  writeSheet_('accounts', [
    ['account_id', 'name', 'type', 'status', 'opened_date', 'closed_date', 'access_level'],
    ...accountRows,
  ]);

  writeSheet_('positions', [
    [
      'account_id',
      'account_name',
      'source',
      'figi',
      'instrument_name',
      'instrument_full_name',
      'instrument_type',
      'quantity',
      'blocked',
      'average_price',
      'expected_yield',
      'current_price',
      'currency',
      'lot',
      'min_price_increment',
      'point_price',
      'current_position_value',
    ],
    ...positionRows,
  ]);

  postProcessPositionsSheet_();

  Logger.log('Готово: выгружено счетов=%s, строк позиций=%s', accounts.length, positionRows.length);
}

function postProcessPositionsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('positions');
  if (!sheet) return;

  // Скрываем технические колонки, но оставляем в таблице.
  // C:source, D:figi, N:lot, O:min_price_increment
  sheet.hideColumns(3, 2);
  sheet.hideColumns(14, 2);

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return;

  // Оформление: заголовок, фильтр и чередование цветов.
  const headerRange = sheet.getRange(1, 1, 1, lastCol);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#d9ead3');
  headerRange.setWrap(true);

  if (sheet.getFilter()) {
    sheet.getFilter().remove();
  }
  if (lastRow >= 2) {
    sheet.getRange(1, 1, lastRow, lastCol).createFilter();
  }

  const bandingRange = sheet.getRange(1, 1, Math.max(lastRow, 2), lastCol);
  const existingBandings = sheet.getBandings();
  existingBandings.forEach((b) => b.remove());
  bandingRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);

  if (lastRow < 2) return;

  // Преобразуем потенциально текстовые числа в number и выставляем числовой формат.
  // H:quantity I:blocked J:average_price K:expected_yield L:current_price
  // N:lot O:min_price_increment P:point_price Q:current_position_value
  const numericCols = [8, 9, 10, 11, 12, 14, 15, 16, 17];
  numericCols.forEach((col) => {
    const rng = sheet.getRange(2, col, lastRow - 1, 1);
    const values = rng.getValues().map((row) => [coerceNumberCell_(row[0])]);
    rng.setValues(values);
    rng.setNumberFormat('0.##########');
  });
}

function coerceNumberCell_(value) {
  if (value === null || value === undefined || value === '') return '';
  if (typeof value === 'number') return value;

  const normalized = String(value).replace(/\s+/g, '').replace(',', '.');
  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : value;
}

function buildBlockedIndex_(positionsResponse) {
  const blockedByFigi = {};
  const blockedByUid = {};

  const add = (item) => {
    const blockedCandidates = [
      parseNumericField_(item?.blocked),
      parseNumericField_(item?.blockedLots),
      parseNumericField_(item?.blockedQuantity),
    ];
    const blocked = firstNonZero_(blockedCandidates);

    if (item?.figi) blockedByFigi[item.figi] = blocked;
    if (item?.instrumentUid) blockedByUid[item.instrumentUid] = blocked;
  };

  (positionsResponse.securities || []).forEach(add);
  (positionsResponse.futures || []).forEach(add);
  (positionsResponse.options || []).forEach(add);

  return { blockedByFigi, blockedByUid };
}

function resolveBlocked_(position, blockedIndex) {
  const portfolioCandidates = [
    parseNumericField_(position?.blocked),
    parseNumericField_(position?.blockedLots),
    parseNumericField_(position?.blockedQuantity),
  ];
  const fromPortfolio = firstNonZero_(portfolioCandidates);
  if (fromPortfolio !== 0) return fromPortfolio;

  const byFigi = blockedIndex.blockedByFigi[position.figi || ''];
  if (byFigi !== undefined) return byFigi;

  const byUid = blockedIndex.blockedByUid[position.instrumentUid || ''];
  if (byUid !== undefined) return byUid;

  return 0;
}


function computeCurrentPositionValue_(instrumentType, qty, currentPrice, pointPrice) {
  if (!qty || !currentPrice) return null;

  if (instrumentType === 'futures') {
    if (pointPrice === null || pointPrice === undefined) return null;
    return qty * currentPrice * pointPrice;
  }

  return qty * currentPrice;
}

function resolveInstrumentFullName_(position, token, futuresMeta, cache) {
  const uid = position.instrumentUid || '';
  const figi = position.figi || '';

  if (uid && cache[uid]) return cache[uid];
  if (figi && cache[figi]) return cache[figi];

  const futureMeta = futuresMeta.byUid[uid] || futuresMeta.byFigi[figi];
  if (futureMeta?.name) {
    if (uid) cache[uid] = futureMeta.name;
    if (figi) cache[figi] = futureMeta.name;
    return futureMeta.name;
  }

  const byUidResponse = uid
    ? callTinkoffApiSafe_(
        '/tinkoff.public.invest.api.contract.v1.InstrumentsService/GetInstrumentBy',
        { idType: 'INSTRUMENT_ID_TYPE_UID', id: uid },
        token
      )
    : null;

  const byUidName = byUidResponse?.instrument?.name || '';
  if (byUidName) {
    if (uid) cache[uid] = byUidName;
    if (figi) cache[figi] = byUidName;
    return byUidName;
  }

  const byFigiResponse = figi
    ? callTinkoffApiSafe_(
        '/tinkoff.public.invest.api.contract.v1.InstrumentsService/GetInstrumentBy',
        { idType: 'INSTRUMENT_ID_TYPE_FIGI', id: figi },
        token
      )
    : null;

  const byFigiName = byFigiResponse?.instrument?.name || '';
  if (byFigiName) {
    if (uid) cache[uid] = byFigiName;
    if (figi) cache[figi] = byFigiName;
    return byFigiName;
  }

  return '';
}

function loadFuturesMeta_(token) {
  const response = callTinkoffApi_(
    '/tinkoff.public.invest.api.contract.v1.InstrumentsService/Futures',
    { instrumentStatus: 'INSTRUMENT_STATUS_BASE' },
    token
  );

  const byFigi = {};
  const byUid = {};
  (response.instruments || []).forEach((f) => {
    if (f.figi) byFigi[f.figi] = f;
    if (f.uid) byUid[f.uid] = f;
  });

  return { byFigi, byUid };
}

function callTinkoffApiSafe_(path, payload, token) {
  try {
    return callTinkoffApi_(path, payload, token);
  } catch (e) {
    return null;
  }
}

function callTinkoffApi_(path, payload, token) {
  const response = UrlFetchApp.fetch(TINKOFF_API_BASE + path, {
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    headers: {
      Authorization: 'Bearer ' + token,
    },
    payload: JSON.stringify(payload || {}),
  });

  const statusCode = response.getResponseCode();
  const bodyText = response.getContentText();
  if (statusCode < 200 || statusCode >= 300) {
    throw new Error('Ошибка Tinkoff API [' + statusCode + ']: ' + bodyText);
  }

  return JSON.parse(bodyText);
}

function writeSheet_(sheetName, values) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  sheet.clear();
  sheet.getRange(1, 1, values.length, values[0].length).setValues(values);
}

function quotationToNumber_(q) {
  if (!q) return 0;
  return Number(q.units || 0) + Number(q.nano || 0) / 1e9;
}

function moneyToNumber_(m) {
  if (!m) return 0;
  return quotationToNumber_(m);
}

function parseNumericField_(value) {
  if (!value) return 0;
  if (typeof value === 'number') return value;
  if (typeof value === 'string') {
    const normalized = value.replace(/\s+/g, '').replace(',', '.');
    const parsed = Number(normalized);
    return Number.isFinite(parsed) ? parsed : 0;
  }

  if (value.units !== undefined || value.nano !== undefined) {
    return Number(value.units || 0) + Number(value.nano || 0) / 1e9;
  }

  return 0;
}

function firstNonZero_(values) {
  for (let i = 0; i < values.length; i += 1) {
    if (values[i] !== 0) return values[i];
  }
  return 0;
}

function safeDivide_(a, b) {
  if (!a || !b) return 0;
  return a / b;
}

function numberOrNull_(value) {
  if (value === null || value === undefined || value === '') return null;
  const num = Number(value);
  return Number.isFinite(num) ? num : null;
}

function readCurrency_(position) {
  return (
    position.currentPrice?.currency ||
    position.averagePositionPrice?.currency ||
    position.expectedYield?.currency ||
    ''
  );
}

function readInstrumentName_(position) {
  return (
    position.ticker ||
    position.instrumentName ||
    position.name ||
    position.title ||
    position.figi ||
    ''
  );
}

function normalizeEnum_(value) {
  return value || '';
}

function toIsoDate_(value) {
  if (!value) return '';
  const d = new Date(value);
  if (isNaN(d.getTime())) return value;
  return d.toISOString();
}

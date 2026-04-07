function onOpen() {
  try {
    ensureAccessControlSheets_();
  } catch (e) {
    // Не блокируем открытие файла для пользователей без прав изменения служебных листов.
  }

  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Склад');
  let access;
  try {
    access = getCurrentUserAccess_();
  } catch (e) {
    access = {
      permissions: {
        refreshAll: false,
        inventory: false,
        search: false,
        storekeeperDashboard: false,
        managementDashboard: false,
        fleetDashboard: false,
        manageAccess: false
      }
    };
  }

  menu.addItem('Меню', 'showMainMenuPanel');
  if (access.permissions.refreshAll) menu.addItem('1. Обновить всё', 'refreshAll');
  menu.addSeparator();
  if (access.permissions.inventory) menu.addItem('2. Инвентаризация', 'showInventoryForm');
  if (access.permissions.search) menu.addItem('3. Поиск товара', 'showSearchForm');
  if (access.permissions.storekeeperDashboard) menu.addItem('4. Дашборд кладовщика', 'showStorekeeperDashboard');
  if (access.permissions.managementDashboard) menu.addItem('5. Дашборд руководителя', 'showManagementDashboard');
  if (access.permissions.fleetDashboard) menu.addItem('6. Автопарк и поездки', 'showFleetTripsDashboard');
  if (access.permissions.manageAccess) {
    menu.addSeparator();
    menu.addItem('6. Управление доступом', 'showAccessAdminPanel');
    menu.addItem('7. Защитить листы', 'syncSheetProtections');
  }
  menu.addToUi();
}
const ACCESS_ROLE_SHEET = 'Роли доступа';
const ACCESS_USERS_SHEET = 'Пользователи и роли';
const ACCESS_SNAPSHOT_KEY = 'ACCESS_CONTROL_SNAPSHOT_V1';
const ACCESS_PERMISSION_FIELDS = [
  'menu',
  'inventory',
  'search',
  'storekeeperDashboard',
  'managementDashboard',
  'fleetDashboard',
  'manageAccess',
  'refreshAll'
];
const ACCESS_PERMISSION_LABELS = {
  menu: 'Меню',
  inventory: 'Инвентаризация',
  search: 'Поиск',
  storekeeperDashboard: 'Дашборд кладовщика',
  managementDashboard: 'Дашборд руководителя',
  fleetDashboard: 'Автопарк и поездки',
  manageAccess: 'Управление доступом',
  refreshAll: 'Обновить всё'
};
const ACCESS_ROLE_COMMENT_HEADER = 'Комментарий';

function normalizeEmail_(value) {
  return String(value || '').trim().toLowerCase();
}

function getRoleSheetHeaders_() {
  return ['Роль'].concat(ACCESS_PERMISSION_FIELDS.map(function (key) {
    return ACCESS_PERMISSION_LABELS[key];
  })).concat([ACCESS_ROLE_COMMENT_HEADER]);
}

function getRoleColumnIndexes_(sheet, allowRepair) {
  if (allowRepair === undefined) allowRepair = true;
  const headers = getRoleSheetHeaders_();
  const currentLastCol = Math.max(sheet.getLastColumn(), 1);
  const currentHeader = sheet.getRange(1, 1, 1, currentLastCol).getValues()[0];
  const missing = headers.some(function (h) { return currentHeader.indexOf(h) === -1; });

  if (missing || currentHeader[0] !== 'Роль') {
    if (!allowRepair) {
      throw new Error('Некорректная структура листа "' + ACCESS_ROLE_SHEET + '". Откройте "Управление доступом" под администратором и восстановите заголовки.');
    }
    const existing = sheet.getLastRow() >= 2 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, currentLastCol).getValues() : [];
    const byHeader = {};
    currentHeader.forEach(function (h, i) {
      byHeader[String(h || '').trim()] = i;
    });
    const migrated = existing.map(function (row) {
      const out = headers.map(function () { return ''; });
      headers.forEach(function (h, idx) {
        const oldIdx = byHeader[h];
        if (typeof oldIdx === 'number') out[idx] = row[oldIdx];
      });
      return out;
    });

    sheet.clear();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    if (migrated.length) {
      sheet.getRange(2, 1, migrated.length, headers.length).setValues(migrated);
    }
    formatHeader_(sheet, 1, 1, 1, headers.length);
    sheet.setFrozenRows(1);
    autoResize_(sheet, 1, headers.length);
  }

  const finalHeader = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const map = {};
  finalHeader.forEach(function (h, i) {
    map[String(h || '').trim()] = i;
  });
  return {
    header: finalHeader,
    indexes: map,
    totalColumns: headers.length
  };
}

function ensureAccessControlSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let roleSheet = ss.getSheetByName(ACCESS_ROLE_SHEET);
  if (!roleSheet) roleSheet = ss.insertSheet(ACCESS_ROLE_SHEET);
  if (roleSheet.getLastRow() === 0) {
    const headers = [getRoleSheetHeaders_()];
    const rows = [
      ['Руководитель', 'Да', 'Да', 'Да', 'Да', 'Да', 'Да', 'Да', 'Да', 'Полный доступ'],
      ['Прораб', 'Да', 'Да', 'Да', 'Нет', 'Нет', 'Нет', 'Нет', 'Нет', 'Инвентаризация и поиск'],
      ['Инженер по закупкам', 'Да', 'Нет', 'Да', 'Да', 'Нет', 'Нет', 'Нет', 'Нет', 'Поиск и дашборд кладовщика'],
      ['Инженер по снабжению', 'Да', 'Да', 'Да', 'Да', 'Да', 'Да', 'Да', 'Да', 'Полный доступ']
    ];
    roleSheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
    roleSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    formatHeader_(roleSheet, 1, 1, 1, headers[0].length);
    roleSheet.setFrozenRows(1);
    autoResize_(roleSheet, 1, headers[0].length);
  }
  getRoleColumnIndexes_(roleSheet);

  let userSheet = ss.getSheetByName(ACCESS_USERS_SHEET);
  if (!userSheet) userSheet = ss.insertSheet(ACCESS_USERS_SHEET);
  if (userSheet.getLastRow() === 0) {
    const headers = [['Email', 'Роль', 'Активен', 'Комментарий']];
    userSheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
    formatHeader_(userSheet, 1, 1, 1, headers[0].length);
    userSheet.setFrozenRows(1);
    const currentEmail = getCurrentUserEmail_();
    if (currentEmail) {
      userSheet.getRange(2, 1, 1, 4).setValues([[currentEmail, 'Инженер по снабжению', 'Да', 'Добавлен автоматически как первый администратор']]);
    }
    autoResize_(userSheet, 1, headers[0].length);
  }
  persistAccessSnapshot_();
}

function ensureAccessControlSheetsReadable_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const roleSheet = ss.getSheetByName(ACCESS_ROLE_SHEET);
  const userSheet = ss.getSheetByName(ACCESS_USERS_SHEET);

  if (!roleSheet || !userSheet) {
    throw new Error('Служебные листы доступа не найдены. Администратор должен открыть "Управление доступом" и инициализировать таблицу.');
  }
}

function isPermissionDeniedError_(error) {
  const message = String(error && error.message || error || '');
  return message.indexOf('PERMISSION_DENIED') !== -1 || message.indexOf('Недостаточно прав') !== -1;
}

function getAccessSnapshot_() {
  const raw = PropertiesService.getScriptProperties().getProperty(ACCESS_SNAPSHOT_KEY);
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch (e) {
    return null;
  }
}

function persistAccessSnapshot_() {
  const payload = {
    rolesMap: readAccessRolesFromSheet_(),
    users: readUserRoleEntriesFromSheet_(),
    updatedAt: new Date().toISOString()
  };
  PropertiesService.getScriptProperties().setProperty(ACCESS_SNAPSHOT_KEY, JSON.stringify(payload));
}

function ensureAccessControlSheetsReadable_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const roleSheet = ss.getSheetByName(ACCESS_ROLE_SHEET);
  const userSheet = ss.getSheetByName(ACCESS_USERS_SHEET);

  if (!roleSheet || !userSheet) {
    throw new Error('Служебные листы доступа не найдены. Администратор должен открыть "Управление доступом" и инициализировать таблицу.');
  }
}

function getCurrentUserEmail_() {
  return normalizeEmail_(Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail());
}

function toBoolAccess_(value) {
  const v = normalizeText_(String(value || ''));
  return ['да', 'true', '1', 'yes', 'y'].indexOf(v) !== -1;
}

function readAccessRolesFromSheet_() {
  ensureAccessControlSheetsReadable_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(ACCESS_ROLE_SHEET);
  const cols = getRoleColumnIndexes_(sheet, false);
  const lastRow = sheet.getLastRow();
  const map = {};
  if (lastRow < 2) return map;
  const data = sheet.getRange(2, 1, lastRow - 1, cols.totalColumns).getValues();
  data.forEach(function (r) {
    const role = String(r[cols.indexes['Роль']] || '').trim();
    if (!role) return;
    const permissions = {};
    ACCESS_PERMISSION_FIELDS.forEach(function (key) {
      const colName = ACCESS_PERMISSION_LABELS[key];
      permissions[key] = toBoolAccess_(r[cols.indexes[colName]]);
    });
    map[role] = {
      role: role,
      permissions: permissions,
      comment: String(r[cols.indexes[ACCESS_ROLE_COMMENT_HEADER]] || '').trim()
    };
  });
  return map;
}

function getAccessRolesMap_() {
  try {
    return readAccessRolesFromSheet_();
  } catch (e) {
    if (!isPermissionDeniedError_(e)) throw e;
    const snapshot = getAccessSnapshot_();
    if (snapshot && snapshot.rolesMap) return snapshot.rolesMap;
    throw new Error('Нет прав для чтения листа "' + ACCESS_ROLE_SHEET + '". Попросите администратора открыть "Управление доступом" и сохранить настройки.');
  }
}

function readUserRoleEntriesFromSheet_() {
  ensureAccessControlSheetsReadable_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(ACCESS_USERS_SHEET);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, 4).getValues().map(function (r, idx) {
    return {
      sheetRow: idx + 2,
      email: normalizeEmail_(r[0]),
      role: String(r[1] || '').trim(),
      active: toBoolAccess_(r[2]),
      comment: String(r[3] || '').trim()
    };
  }).filter(function (row) { return !!row.email; });
}

function getUserRoleEntries_() {
  try {
    return readUserRoleEntriesFromSheet_();
  } catch (e) {
    if (!isPermissionDeniedError_(e)) throw e;
    const snapshot = getAccessSnapshot_();
    if (snapshot && snapshot.users) return snapshot.users;
    throw new Error('Нет прав для чтения листа "' + ACCESS_USERS_SHEET + '". Попросите администратора открыть "Управление доступом" и сохранить настройки.');
  }
}

function getActiveUsersWithPermission_(permissionKey) {
  const rolesMap = getAccessRolesMap_();
  return getUserRoleEntries_()
    .filter(function (row) {
      if (!row.active) return false;
      const roleInfo = rolesMap[row.role];
      return !!(roleInfo && roleInfo.permissions && roleInfo.permissions[permissionKey]);
    })
    .map(function (row) { return row.email; });
}

function getCurrentUserAccess_() {
  const email = getCurrentUserEmail_();
  const rolesMap = getAccessRolesMap_();
  const users = getUserRoleEntries_();
  const userRow = users.find(function (row) { return row.email === email && row.active; }) || null;
  const roleName = userRow ? userRow.role : '';
  const role = rolesMap[roleName] || null;
  const permissions = {
    menu: false,
    inventory: false,
    search: false,
    storekeeperDashboard: false,
    managementDashboard: false,
    fleetDashboard: false,
    manageAccess: false,
    refreshAll: false
  };
  if (role && role.permissions) {
    ACCESS_PERMISSION_FIELDS.forEach(function (key) { permissions[key] = !!role.permissions[key]; });
  }
  return {
    email: email,
    role: roleName,
    roleComment: role ? role.comment : '',
    permissions: permissions,
    hasAnyAccess: ACCESS_PERMISSION_FIELDS.some(function (key) { return !!permissions[key]; })
  };
}

function requirePermission_(permissionKey, actionName) {
  const access = getCurrentUserAccess_();
  if (access.permissions[permissionKey]) return access;
  throw new Error('Нет доступа: ' + actionName + '. Текущая роль: ' + (access.role || 'не назначена') + '.');
}

function getCurrentUserRoleInfo() {
  const access = getCurrentUserAccess_();
  return {
    email: access.email,
    role: access.role,
    roleComment: access.roleComment,
    permissions: access.permissions,
    permissionLabels: ACCESS_PERMISSION_LABELS,
    hasAnyAccess: access.hasAnyAccess
  };
}

function getAccessControlData() {
  requirePermission_('manageAccess', 'управление доступом');
  const rolesMap = getAccessRolesMap_();
  const users = getUserRoleEntries_();
  const roles = Object.keys(rolesMap).map(function (name) {
    return {
      role: name,
      permissions: rolesMap[name].permissions,
      comment: rolesMap[name].comment
    };
  }).sort(function (a, b) { return a.role.localeCompare(b.role, 'ru'); });

  return {
    currentUser: getCurrentUserRoleInfo(),
    roles: roles,
    users: users
  };
}

function saveAccessRole(payload) {
  requirePermission_('manageAccess', 'управление доступом');
  ensureAccessControlSheets_();

  const roleName = String(payload && payload.role || '').trim();
  const comment = String(payload && payload.comment || '').trim();
  const permissionsPayload = payload && payload.permissions || {};

  if (!roleName) throw new Error('Укажи название роли.');

  const rolesMap = getAccessRolesMap_();
  if (rolesMap[roleName]) throw new Error('Роль с таким названием уже существует.');

  const permissionValues = ACCESS_PERMISSION_FIELDS.map(function (key) {
    return permissionsPayload[key] ? 'Да' : 'Нет';
  });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(ACCESS_ROLE_SHEET);
  const row = [roleName].concat(permissionValues).concat([comment]);
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, row.length).setValues([row]);
  autoResize_(sheet, 1, row.length);
  persistAccessSnapshot_();
  return 'Роль добавлена: ' + roleName;
}

function saveAccessControlData(payload) {
  requirePermission_('manageAccess', 'управление доступом');
  ensureAccessControlSheets_();

  const users = Array.isArray(payload && payload.users) ? payload.users : [];
  const rolesMap = getAccessRolesMap_();
  const currentEmail = getCurrentUserEmail_();
  const prepared = [];
  const seen = {};

  users.forEach(function (row) {
    const email = normalizeEmail_(row.email);
    const role = String(row.role || '').trim();
    const active = row.active === false ? false : true;
    const comment = String(row.comment || '').trim();
    if (!email) return;
    if (!rolesMap[role]) throw new Error('Неизвестная роль: ' + role + '. Сначала добавь роль в лист "' + ACCESS_ROLE_SHEET + '".');
    if (seen[email]) throw new Error('Пользователь ' + email + ' указан несколько раз.');
    seen[email] = true;
    prepared.push([email, role, active ? 'Да' : 'Нет', comment]);
  });

  if (!prepared.length) throw new Error('Список пользователей пуст.');

  const currentEntry = prepared.find(function (r) { return r[0] === currentEmail; });
  if (!currentEntry) throw new Error('Нельзя удалить себе доступ. Оставь свою учетную запись в списке.');
  if (!toBoolAccess_(currentEntry[2])) throw new Error('Нельзя отключить свою учетную запись.');
  const currentRoleInfo = rolesMap[currentEntry[1]];
  if (!currentRoleInfo || !currentRoleInfo.permissions.manageAccess) {
    throw new Error('Нельзя снять у себя право управления доступом в этом действии.');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(ACCESS_USERS_SHEET);
  sheet.getRange(2, 1, Math.max(sheet.getMaxRows() - 1, 1), 4).clearContent();
  if (prepared.length) {
    if (sheet.getMaxRows() < prepared.length + 1) {
      sheet.insertRowsAfter(sheet.getMaxRows(), prepared.length + 1 - sheet.getMaxRows());
    }
    sheet.getRange(2, 1, prepared.length, 4).setValues(prepared);
  }
  autoResize_(sheet, 1, 4);
  persistAccessSnapshot_();
  syncSheetProtections();
  return 'Доступы обновлены. Пользователей: ' + prepared.length;
}

function syncSheetProtections() {
  requirePermission_('manageAccess', 'защита листов');
  ensureAccessControlSheets_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const adminEmails = getActiveUsersWithPermission_('manageAccess');
  const inventoryEmails = getActiveUsersWithPermission_('inventory');
  const storekeeperEmails = getActiveUsersWithPermission_('storekeeperDashboard');
  const fleetEmails = getActiveUsersWithPermission_('fleetDashboard');

  if (!adminEmails.length) throw new Error('Нет ни одного администратора с правом управления доступом.');

  const storekeeperSheetNames = [
    'Номенклатура',
    'Журнал движения',
    'Справочник',
    'Запросы номенклатуры',
    'Остатки',
    'Дашборд',
    'Закреплено',
    'Ответственные',
    'История ответственности'
  ];
  const fleetSheetNames = ['Автопарк', 'ТО и ремонты', 'Поездки автопарк', 'Списания накоплений'];
  const adminOnlySheets = [ACCESS_ROLE_SHEET, ACCESS_USERS_SHEET];

  const extraEditorsBySheet = {};
  const inventorySheetNames = ['Журнал движения', 'Справочник', 'Запросы номенклатуры', 'Остатки', 'Дашборд', 'Закреплено'];
  inventorySheetNames.forEach(function (name) {
    extraEditorsBySheet[name] = (extraEditorsBySheet[name] || []).concat(inventoryEmails);
  });
  storekeeperSheetNames.forEach(function (name) {
    extraEditorsBySheet[name] = (extraEditorsBySheet[name] || []).concat(storekeeperEmails);
  });
  fleetSheetNames.forEach(function (name) { extraEditorsBySheet[name] = fleetEmails; });

  ss.getSheets().forEach(function (sheet) {
    const sheetName = sheet.getName();
    let protection = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0];
    if (!protection) {
      protection = sheet.protect();
    }

    const extraEditors = adminOnlySheets.indexOf(sheetName) === -1 ? (extraEditorsBySheet[sheetName] || []) : [];
    const editors = Array.from(new Set(adminEmails.concat(extraEditors))).filter(function (email) { return !!email; });

    protection.setDescription('Автозащита: редактирование по роли доступа');
    protection.setWarningOnly(false);
    try { protection.removeEditors(protection.getEditors()); } catch (e) {}
    try { if (editors.length) protection.addEditors(editors); } catch (e) {}
    if (protection.canDomainEdit && protection.canDomainEdit()) {
      protection.setDomainEdit(false);
    }
  });

  return 'Листы защищены по ролям. Администраторов: ' + adminEmails.length + ', инвентаризация: ' + inventoryEmails.length + ', кладовщик/закупки: ' + storekeeperEmails.length + ', автопарк: ' + fleetEmails.length + '.';
}

function showAccessAdminPanel() {
  requirePermission_('manageAccess', 'управление доступом');
  const html = HtmlService.createHtmlOutputFromFile('AccessAdminPanel')
    .setWidth(1200)
    .setHeight(820);
  SpreadsheetApp.getUi().showModalDialog(html, 'Управление доступом');
}

function getItemCardData(article) {
  const articleNorm = String(article || '').trim().toLowerCase();
  if (!articleNorm) {
    throw new Error('Не выбран товар.');
  }

  const item = getCatalogItemByArticle_(articleNorm);
  if (!item) {
    throw new Error('Товар не найден.');
  }

  const balances = getBalancesData_()
    .filter(function (row) {
      return String(row.article || '').trim().toLowerCase() === articleNorm;
    })
    .map(function (row) {
      return {
        objectName: row.objectName,
        qty: round3_(Number(row.qty) || 0),
        price: round2_(Number(row.price) || 0),
        sum: round2_(Number(row.sum) || 0)
      };
    })
    .sort(function (a, b) {
      return String(a.objectName).localeCompare(String(b.objectName), 'ru');
    });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const journal = ss.getSheetByName('Журнал движения');
  if (!journal) throw new Error('Не найден лист "Журнал движения".');

  const journalLastRow = journal.getLastRow();
  const journalData = journalLastRow >= 2
    ? journal.getRange(2, 1, journalLastRow - 1, 16).getValues()
    : [];

  const movements = journalData
    .filter(function (r) {
      return String(r[2] || '').trim().toLowerCase() === articleNorm;
    })
    .map(function (r) {
      return {
        date: formatDateTimeRu_(r[0]),
        operation: String(r[1] || '').trim(),
        qty: round3_(Number(r[7]) || 0),
        fromObj: String(r[8] || '').trim(),
        toObj: String(r[9] || '').trim(),
        basis: String(r[10] || '').trim(),
        employee: String(r[11] || '').trim(),
        comment: String(r[12] || '').trim(),
        user: String(r[13] || '').trim(),
        price: round2_(Number(r[15]) || 0)
      };
    })
    .sort(function (a, b) {
      const da = a.date || '';
      const db = b.date || '';
      return db.localeCompare(da);
    })
    .slice(0, 30);

  const assigned = getAssignedInstrumentRows({})
    .filter(function (row) {
      return String(row.article || '').trim().toLowerCase() === articleNorm;
    })
    .map(function (row) {
      return {
        employee: row.employee,
        objectName: row.objectName,
        qty: round3_(Number(row.qty) || 0),
        lastIssueDate: row.assignedAt || ''
      };
    });

  return {
    item: {
      article: item.article,
      name: item.name,
      type: item.type,
      category: item.category,
      unit: item.unit,
      price: round2_(Number(item.price) || 0),
      active: item.active,
      comment: item.comment
    },
    balances: balances,
    movements: movements,
    assigned: assigned
  };
}
/**
 * =========================
 * REQUESTS FOR NEW ITEMS
 * =========================
 */

function ensureRequestsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Запросы номенклатуры');

  if (!sheet) {
    sheet = ss.insertSheet('Запросы номенклатуры');
  }

  const headers = [[
    'ID запроса',
    'Дата запроса',
    'Email автора',
    'Объект',
    'Общее основание',
    'Название товара',
    'Ед. изм.',
    'Количество',
    'Цена за ед.',
    'Статус',
    'Тип',
    'Категория',
    'Итоговое название',
    'Комментарий кладовщика',
    'Артикул созданного товара',
    'Дата обработки',
    'Обработал'
  ]];

  const currentHeader = sheet.getRange(1, 1, 1, headers[0].length).getValues()[0];
  const isEmptyHeader = currentHeader.every(function (v) { return String(v || '').trim() === ''; });

  if (isEmptyHeader) {
    sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
    formatHeader_(sheet, 1, 1, 1, headers[0].length);
    sheet.setFrozenRows(1);
    autoResize_(sheet, 1, headers[0].length);
  }

  return sheet;
}

function generateRequestId_() {
  return 'REQ-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss') + '-' + Math.floor(Math.random() * 1000);
}

function getRequestFormInitData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dir = ss.getSheetByName('Справочник');
  if (!dir) throw new Error('Не найден лист "Справочник".');

  return {
    objects: getColumnValues(dir, 1, 2),
    units: getColumnValues(dir, 6, 2)
  };
}

function saveNewItemRequest(payload) {
  ensureRequestsSheet_();

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Запросы номенклатуры');

    const objectName = String(payload.objectName || '').trim();
    const commonBasis = String(payload.commonBasis || '').trim();
    const rows = Array.isArray(payload.rows) ? payload.rows : [];

    if (!objectName) throw new Error('Не выбран объект.');
    if (!commonBasis) throw new Error('Укажи общее основание.');
    if (!rows.length) throw new Error('Нет строк для запроса.');
    if (rows.length > 15) throw new Error('В одном запросе можно отправить не более 15 наименований.');

    const errors = [];
    const preparedRows = [];

    rows.forEach(function (row, index) {
      const line = index + 1;
      const name = String(row.name || '').trim();
      const unit = String(row.unit || '').trim();
      const qty = Number(row.qty);
      const price = Number(row.price);

      if (!name) errors.push('Строка ' + line + ': не заполнено название.');
      if (!unit) errors.push('Строка ' + line + ': не заполнена ед. изм.');
      if (!isFinite(qty) || qty <= 0) errors.push('Строка ' + line + ': количество должно быть больше нуля.');
      if (!isFinite(price) || price < 0) errors.push('Строка ' + line + ': цена некорректна.');

      preparedRows.push({
        name: name,
        unit: unit,
        qty: isFinite(qty) ? qty : 0,
        price: isFinite(price) ? price : 0
      });
    });

    if (errors.length) {
      throw new Error(errors.join('\n'));
    }

    const requestId = generateRequestId_();
    const now = new Date();
    const user = Session.getActiveUser().getEmail() || '';

    const values = preparedRows.map(function (row) {
      return [
        requestId,
        now,
        user,
        objectName,
        commonBasis,
        row.name,
        row.unit,
        row.qty,
        row.price,
        'Новый',
        '',
        '',
        '',
        '',
        '',
        '',
        ''
      ];
    });

    sheet.getRange(sheet.getLastRow() + 1, 1, values.length, 17).setValues(values);
    sheet.getRange(sheet.getLastRow() - values.length + 1, 2, values.length, 1).setNumberFormat('dd.MM.yyyy HH:mm');
    sheet.getRange(sheet.getLastRow() - values.length + 1, 8, values.length, 2).setNumberFormat('#,##0.00');

    autoResize_(sheet, 1, 17);

    return 'Запрос отправлен. ID: ' + requestId + '. Строк: ' + values.length;
  } finally {
    lock.releaseLock();
  }
}

function getPendingNewItemRequests() {
  ensureRequestsSheet_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Запросы номенклатуры');
  const dir = ss.getSheetByName('Справочник');

  if (!sheet) throw new Error('Не найден лист "Запросы номенклатуры".');
  if (!dir) throw new Error('Не найден лист "Справочник".');

  const lastRow = sheet.getLastRow();
  const types = getColumnValues(dir, 2, 2);
  const categoriesByType = getCategoriesByType_();

  if (lastRow < 2) {
    return {
      types: types,
      categoriesByType: categoriesByType,
      rows: []
    };
  }

  const data = sheet.getRange(2, 1, lastRow - 1, 17).getValues();

  const rows = data
    .map(function (r, index) {
      return {
        sheetRow: index + 2,
        requestId: String(r[0] || '').trim(),
        requestDate: formatDateTimeRu_(r[1]),
        requesterEmail: String(r[2] || '').trim(),
        objectName: String(r[3] || '').trim(),
        commonBasis: String(r[4] || '').trim(),
        name: String(r[5] || '').trim(),
        unit: String(r[6] || '').trim(),
        qty: Number(r[7]) || 0,
        price: Number(r[8]) || 0,
        status: String(r[9] || '').trim(),
        type: String(r[10] || '').trim(),
        category: String(r[11] || '').trim(),
        finalName: String(r[12] || '').trim(),
        storekeeperComment: String(r[13] || '').trim(),
        createdArticle: String(r[14] || '').trim(),
        processedAt: formatDateTimeRu_(r[15]),
        processedBy: String(r[16] || '').trim()
      };
    })
    .filter(function (row) {
      return row.status === 'Новый' || row.status === 'В работе';
    });

  return {
    types: types,
    categoriesByType: categoriesByType,
    rows: rows
  };
}

function ensureJournalPriceColumn_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const journal = ss.getSheetByName('Журнал движения');
  if (!journal) throw new Error('Не найден лист "Журнал движения".');

  const header = String(journal.getRange(1, 16).getValue() || '').trim();
  if (!header) {
    journal.getRange(1, 16).setValue('Цена операции');
  }
}

function updateCatalogPriceByArticle_(article, newPrice) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const catalog = ss.getSheetByName('Номенклатура');
  if (!catalog) throw new Error('Не найден лист "Номенклатура".');

  const articleNorm = String(article || '').trim().toLowerCase();
  const price = Number(newPrice);

  if (!articleNorm || !isFinite(price) || price < 0) return;

  const lastRow = catalog.getLastRow();
  if (lastRow < 2) return;

  const articles = catalog.getRange(2, 1, lastRow - 1, 1).getValues().flat();

  for (var i = 0; i < articles.length; i++) {
    if (String(articles[i] || '').trim().toLowerCase() === articleNorm) {
      catalog.getRange(i + 2, 6).setValue(price);
      return;
    }
  }
}

function showMainMenuPanel() {
  const html = HtmlService.createHtmlOutputFromFile('QuickAccessPanel')
    .setTitle('Меню');
  SpreadsheetApp.getUi().showSidebar(html);
}

function showQuickAccessPanel() {
  showMainMenuPanel();
}

function openQuickInventory() {
  showInventoryForm();
}

function openQuickSearch() {
  showSearchForm();
}

function openQuickStorekeeperDashboard() {
  showStorekeeperDashboard();
}

function openQuickManagementDashboard() {
  showManagementDashboard();
}

function openQuickFleetDashboard() {
  showFleetTripsDashboard();
}

function openQuickRefreshAll() {
  requirePermission_('refreshAll', 'обновление данных');
  refreshAll();
}
/**
 * =========================
 * UI / FORMS
 * =========================
 */
function showManagementDashboard() {
  requirePermission_('managementDashboard', 'дашборд руководителя');
  const html = HtmlService.createHtmlOutputFromFile('ManagementDashboard')
    .setWidth(1400)
    .setHeight(900);
  SpreadsheetApp.getUi().showModalDialog(html, 'Дашборд руководителя');
}

function showFleetTripsDashboard() {
  requirePermission_('fleetDashboard', 'панель автопарка и поездок');
  const html = HtmlService.createHtmlOutputFromFile('FleetTripsDashboard')
    .setWidth(1480)
    .setHeight(920);
  SpreadsheetApp.getUi().showModalDialog(html, 'Автопарк и поездки');
}

function showInventoryForm() {
  requirePermission_('inventory', 'инвентаризация');
  const html = HtmlService.createHtmlOutputFromFile('InventoryForm')
    .setWidth(1400)
    .setHeight(900);
  SpreadsheetApp.getUi().showModalDialog(html, 'Инвентаризация');
}

function showSearchForm() {
  requirePermission_('search', 'поиск товара');
  const html = HtmlService.createHtmlOutputFromFile('SearchForm')
    .setWidth(1300)
    .setHeight(850);
  SpreadsheetApp.getUi().showModalDialog(html, 'Поиск товара');
}

function showImportForm() {
  const html = HtmlService.createHtmlOutputFromFile('ImportForm')
    .setWidth(1400)
    .setHeight(900);
  SpreadsheetApp.getUi().showModalDialog(html, 'Импорт товаров');
}

/**
 * =========================
 * DATA FOR HTML FORMS
 * =========================
 */
function getInventoryFormInitData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dir = ss.getSheetByName('Справочник');
  const balances = ss.getSheetByName('Остатки');

  if (!dir) throw new Error('Не найден лист "Справочник".');
  if (!balances) throw new Error('Не найден лист "Остатки".');

  const objects = getColumnValues(dir, 1, 2);
  const items = getBalancesData_();

  return {
    objects: objects,
    items: items
  };
}

function getSearchFormData() {
  return getBalancesData_();
}

function getImportFormInitData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dir = ss.getSheetByName('Справочник');
  const catalog = ss.getSheetByName('Номенклатура');

  if (!dir) throw new Error('Не найден лист "Справочник".');
  if (!catalog) throw new Error('Не найден лист "Номенклатура".');

  const objects = getColumnValues(dir, 1, 2);
  const types = getColumnValues(dir, 2, 2);
  const units = getColumnValues(dir, 6, 2);

  const lastDirRow = dir.getLastRow();
  const typeCategoryPairs = lastDirRow >= 2
    ? dir.getRange(2, 3, lastDirRow - 1, 2).getValues()
    : [];

  const categoriesByType = {};
  typeCategoryPairs.forEach(function (row) {
    const type = String(row[0] || '').trim();
    const category = String(row[1] || '').trim();
    if (!type || !category) return;

    if (!categoriesByType[type]) categoriesByType[type] = [];
    if (categoriesByType[type].indexOf(category) === -1) {
      categoriesByType[type].push(category);
    }
  });

  const lastCatalogRow = catalog.getLastRow();
  const existingItems = lastCatalogRow >= 2
    ? catalog.getRange(2, 1, lastCatalogRow - 1, 8).getValues().map(function (r) {
        return {
          article: String(r[0] || '').trim(),
          type: String(r[1] || '').trim(),
          category: String(r[2] || '').trim(),
          name: String(r[3] || '').trim(),
          unit: String(r[4] || '').trim(),
          price: Number(r[5]) || 0,
          active: String(r[6] || '').trim(),
          comment: String(r[7] || '').trim()
        };
      })
    : [];

  return {
    objects: objects,
    types: types,
    units: units,
    categoriesByType: categoriesByType,
    existingItems: existingItems
  };
}

function getManagementDashboardInitData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dir = ss.getSheetByName('Справочник');
  if (!dir) throw new Error('Не найден лист "Справочник".');

  return {
    objects: getColumnValues(dir, 1, 2),
    types: getColumnValues(dir, 2, 2),
    employees: getEmployees_()
  };
}

/**
 * =========================
 * DASHBOARD
 * =========================
 */
function getManagementDashboardData(filters) {
  filters = filters || {};

  const objectFilter = String(filters.objectName || '').trim();
  const typeFilter = String(filters.typeName || '').trim();
  const employeeFilter = String(filters.employee || '').trim();
  const q = normalizeText_(filters.query || '');
  const dateFrom = parseDashboardDate_(filters.dateFrom, false);
  const dateTo = parseDashboardDate_(filters.dateTo, true);

  const balances = getBalancesData_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const journal = ss.getSheetByName('Журнал движения');
  if (!journal) throw new Error('Не найден лист "Журнал движения".');

  const journalLastRow = journal.getLastRow();
  const journalData = journalLastRow >= 2
    ? journal.getRange(2, 1, journalLastRow - 1, 16).getValues()
    : [];

  let filteredBalances = balances.slice();

  if (objectFilter) {
    filteredBalances = filteredBalances.filter(function (r) {
      return String(r.objectName || '').trim() === objectFilter;
    });
  }

  if (typeFilter) {
    filteredBalances = filteredBalances.filter(function (r) {
      return String(r.type || '').trim() === typeFilter;
    });
  }

  if (q) {
    filteredBalances = filteredBalances.filter(function (r) {
      return normalizeText_([r.article, r.name, r.type, r.category, r.objectName].join(' ')).indexOf(q) !== -1;
    });
  }

  const filteredJournal = journalData
    .map(function (r) {
      return {
        opDate: r[0] instanceof Date ? r[0] : null,
        operation: String(r[1] || '').trim(),
        article: String(r[2] || '').trim(),
        name: String(r[3] || '').trim(),
        type: String(r[4] || '').trim(),
        category: String(r[5] || '').trim(),
        unit: String(r[6] || '').trim(),
        qty: round3_(Number(r[7]) || 0),
        fromObj: String(r[8] || '').trim(),
        toObj: String(r[9] || '').trim(),
        basis: String(r[10] || '').trim(),
        employee: String(r[11] || '').trim(),
        comment: String(r[12] || '').trim(),
        user: String(r[13] || '').trim(),
        sortTime: r[14] instanceof Date ? r[14].getTime() : (r[0] instanceof Date ? r[0].getTime() : 0),
        price: round2_(Number(r[15]) || 0)
      };
    })
    .filter(function (row) {
      if (typeFilter && row.type !== typeFilter) return false;
      if (employeeFilter && row.employee !== employeeFilter) return false;

      if (objectFilter) {
        const hitObject = row.fromObj === objectFilter || row.toObj === objectFilter;
        if (!hitObject) return false;
      }

      if (dateFrom && row.opDate && row.opDate < dateFrom) return false;
      if (dateTo && row.opDate && row.opDate > dateTo) return false;

      if (q) {
        const txt = normalizeText_([row.article, row.name, row.type, row.category, row.fromObj, row.toObj, row.basis, row.comment, row.employee].join(' '));
        if (txt.indexOf(q) === -1) return false;
      }

      return true;
    });

  const totalValue = filteredBalances.reduce(function (sum, r) { return sum + (Number(r.sum) || 0); }, 0);
  const totalQty = filteredBalances.reduce(function (sum, r) { return sum + (Number(r.qty) || 0); }, 0);
  const positionsCount = filteredBalances.length;
  const objectsCount = new Set(filteredBalances.map(function (r) { return String(r.objectName || '').trim(); }).filter(Boolean)).size;

  const valueByObjectMap = {};
  filteredBalances.forEach(function (r) {
    const objectName = String(r.objectName || '').trim();
    if (!objectName) return;
    valueByObjectMap[objectName] = (valueByObjectMap[objectName] || 0) + (Number(r.sum) || 0);
  });

  const valueByObject = Object.keys(valueByObjectMap)
    .map(function (name) {
      return { name: name, value: round2_(valueByObjectMap[name]) };
    })
    .sort(function (a, b) { return b.value - a.value; });

  const qtyByTypeMap = {};
  const valueByTypeMap = {};
  filteredBalances.forEach(function (r) {
    const typeName = String(r.type || '').trim() || 'Без типа';
    qtyByTypeMap[typeName] = (qtyByTypeMap[typeName] || 0) + (Number(r.qty) || 0);
    valueByTypeMap[typeName] = (valueByTypeMap[typeName] || 0) + (Number(r.sum) || 0);
  });

  const summaryByType = Object.keys(qtyByTypeMap)
    .map(function (name) {
      return { name: name, qty: round3_(qtyByTypeMap[name]), value: round2_(valueByTypeMap[name] || 0) };
    })
    .sort(function (a, b) { return b.value - a.value; });

  const allProducts = filteredBalances
    .map(function (r) {
      return {
        objectName: String(r.objectName || '').trim(),
        article: String(r.article || '').trim(),
        name: String(r.name || '').trim(),
        type: String(r.type || '').trim(),
        category: String(r.category || '').trim(),
        unit: String(r.unit || '').trim(),
        qty: round3_(Number(r.qty) || 0),
        price: round2_(Number(r.price) || 0),
        value: round2_(Number(r.sum) || 0)
      };
    })
    .sort(function (a, b) { return b.value - a.value; });

  const topProducts = allProducts.slice(0, 20);

  const problematic = filteredBalances
    .filter(function (r) {
      const qty = Number(r.qty) || 0;
      const price = Number(r.price) || 0;
      return qty <= 0 || price <= 0;
    })
    .map(function (r) {
      const qty = Number(r.qty) || 0;
      const price = Number(r.price) || 0;
      let problemType = '';
      if (qty < 0) problemType = 'Отрицательный остаток';
      else if (qty === 0 && price <= 0) problemType = 'Нулевой остаток и цена';
      else if (qty === 0) problemType = 'Нулевой остаток';
      else if (price <= 0) problemType = 'Не заполнена цена';

      return {
        objectName: String(r.objectName || '').trim(),
        article: String(r.article || '').trim(),
        name: String(r.name || '').trim(),
        type: String(r.type || '').trim(),
        category: String(r.category || '').trim(),
        unit: String(r.unit || '').trim(),
        qty: round3_(qty),
        price: round2_(price),
        value: round2_(Number(r.sum) || 0),
        problemType: problemType
      };
    })
    .slice(0, 100);

  const recentMovements = filteredJournal
    .sort(function (a, b) { return b.sortTime - a.sortTime; })
    .slice(0, 50)
    .map(function (r) {
      return {
        date: r.opDate ? formatDateTimeRu_(r.opDate) : '',
        operation: r.operation,
        article: r.article,
        name: r.name,
        type: r.type,
        qty: r.qty,
        fromObj: r.fromObj,
        toObj: r.toObj,
        basis: r.basis,
        employee: r.employee,
        comment: r.comment
      };
    });

  const responsibilityFilters = {
    query: filters.query || '',
    objectName: objectFilter,
    typeName: typeFilter,
    employee: employeeFilter
  };

  return {
    filters: {
      objectName: objectFilter,
      typeName: typeFilter,
      dateFrom: filters.dateFrom || '',
      dateTo: filters.dateTo || '',
      employee: employeeFilter,
      query: filters.query || ''
    },
    kpi: {
      totalValue: round2_(totalValue),
      totalQty: round3_(totalQty),
      positionsCount: positionsCount,
      objectsCount: objectsCount,
      movementsCount: filteredJournal.length,
      responsibleCount: getAssignedInstrumentRows(responsibilityFilters).length
    },
    valueByObject: valueByObject,
    summaryByType: summaryByType,
    allProducts: allProducts,
    topProducts: topProducts,
    problematic: problematic,
    recentMovements: recentMovements,
    currentResponsibilities: getAssignedInstrumentRows(responsibilityFilters),
    history: getResponsibilityHistoryData_(responsibilityFilters)
  };
}

/**
 * =========================
 * INVENTORY
 * =========================
 */
function saveInventoryForm(payload) {
  requirePermission_('inventory', 'проведение инвентаризации');
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const journal = ss.getSheetByName('Журнал движения');

    if (!journal) throw new Error('Не найден лист "Журнал движения".');

    const objectName = String(payload.objectName || '').trim();
    const rows = Array.isArray(payload.rows) ? payload.rows : [];

    if (!objectName) {
      throw new Error('Не выбран объект.');
    }

    if (!rows.length) {
      throw new Error('Нет данных для проведения.');
    }

    refreshBalances();

    const journalRows = [];
    const errors = [];
    const now = new Date();
    const user = Session.getActiveUser().getEmail() || '';

    rows.forEach(function (r, index) {
      const line = index + 1;
      const article = String(r.article || '').trim();
      const name = String(r.name || '').trim();
      const type = String(r.type || '').trim();
      const category = String(r.category || '').trim();
      const unit = String(r.unit || '').trim();
      const action = String(r.action || '').trim();
      const basis = String(r.basis || '').trim();
      const destination = String(r.destination || '').trim();
      const qty = Number(r.changeQty);

      if (!action) return;

      if (!article) {
        errors.push('Строка ' + line + ': не найден артикул.');
        return;
      }

      if (!isFinite(qty) || qty <= 0) {
        errors.push('Строка ' + line + ': укажи количество больше нуля.');
        return;
      }

      if (action === 'Списание' && !basis) {
        errors.push('Строка ' + line + ': причина списания не указана.');
        return;
      }

      if (action === 'Приход' && !basis) {
        errors.push('Строка ' + line + ': укажи откуда приход.');
        return;
      }

      if (action === 'Перемещение' && !destination) {
        errors.push('Строка ' + line + ': укажи куда перемещать.');
        return;
      }

      if (action === 'Перемещение' && destination === objectName) {
        errors.push('Строка ' + line + ': объект назначения совпадает с текущим.');
        return;
      }

      if (action === 'Списание' || action === 'Перемещение') {
        const stock = getCurrentStock(objectName, article);
        if (qty > stock) {
          errors.push(
            'Строка ' + line + ': недостаточно остатка. В наличии ' +
            formatQty_(stock) + ', указано ' + formatQty_(qty) + '.'
          );
          return;
        }
      }

      if (action === 'Списание') {
        journalRows.push([
          new Date(),
          'Списание',
          article,
          name,
          type,
          category,
          unit,
          qty,
          objectName,
          '',
          basis,
          '',
          'Инвентаризация',
          user,
          now
        ]);
      }

      if (action === 'Приход') {
        journalRows.push([
          new Date(),
          'Приход',
          article,
          name,
          type,
          category,
          unit,
          qty,
          '',
          objectName,
          basis,
          '',
          'Инвентаризация',
          user,
          now
        ]);
      }

      if (action === 'Перемещение') {
        journalRows.push([
          new Date(),
          'Перемещение',
          article,
          name,
          type,
          category,
          unit,
          qty,
          objectName,
          destination,
          basis || 'Инвентаризация',
          '',
          'Инвентаризация',
          user,
          now
        ]);
      }
    });

    if (errors.length) {
      throw new Error(errors.join('\n'));
    }

    if (!journalRows.length) {
      throw new Error('Нет заполненных действий для проведения.');
    }

    journal
      .getRange(journal.getLastRow() + 1, 1, journalRows.length, 15)
      .setValues(journalRows);

    refreshAll();

    return 'Инвентаризация проведена. Добавлено строк: ' + journalRows.length;
  } finally {
    lock.releaseLock();
  }
}

/**
 * =========================
 * REFRESH
 * =========================
 */
function refreshBalances() {
  ensureJournalPriceColumn_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const journal = ss.getSheetByName('Журнал движения');
  const catalog = ss.getSheetByName('Номенклатура');
  const balances = ss.getSheetByName('Остатки');

  if (!journal) throw new Error('Не найден лист "Журнал движения".');
  if (!catalog) throw new Error('Не найден лист "Номенклатура".');
  if (!balances) throw new Error('Не найден лист "Остатки".');

  balances.clearContents();

  const headers = [[
    'Объект',
    'Артикул',
    'Название товара',
    'Тип',
    'Категория',
    'Ед. изм.',
    'Количество',
    'Цена за ед.',
    'Сумма'
  ]];

  balances.getRange(1, 1, 1, headers[0].length).setValues(headers);
  formatHeader_(balances, 1, 1, 1, headers[0].length);
  balances.setFrozenRows(1);

  const journalLastRow = journal.getLastRow();
  const catalogData = getCatalogData_();

  if (journalLastRow < 2) {
    autoResize_(balances, 1, 9);
    return;
  }

  const journalData = journal.getRange(2, 1, journalLastRow - 1, 16).getValues();

  const catalogPriceMap = {};
  const metaMap = {};

  catalogData.forEach(function (r) {
    const article = String(r.article || '').trim();
    if (!article) return;

    catalogPriceMap[article] = Number(r.price) || 0;
    metaMap[article] = {
      name: r.name || '',
      type: r.type || '',
      category: r.category || '',
      unit: r.unit || ''
    };
  });

  const stockMap = {};

  function getKey_(objectName, article) {
    return objectName + '||' + article;
  }

  function ensureState_(objectName, article, fallbackMeta) {
    const key = getKey_(objectName, article);

    if (!stockMap[key]) {
      stockMap[key] = {
        qty: 0,
        value: 0
      };
    }

    if (!metaMap[article]) {
      metaMap[article] = {
        name: fallbackMeta.name || '',
        type: fallbackMeta.type || '',
        category: fallbackMeta.category || '',
        unit: fallbackMeta.unit || ''
      };
    }

    return stockMap[key];
  }

  journalData.forEach(function (r) {
    const operation = String(r[1] || '').trim();
    const article = String(r[2] || '').trim();
    const name = String(r[3] || '').trim();
    const type = String(r[4] || '').trim();
    const category = String(r[5] || '').trim();
    const unit = String(r[6] || '').trim();
    const qty = Number(r[7]) || 0;
    const fromObj = String(r[8] || '').trim();
    const toObj = String(r[9] || '').trim();

    const opPriceCell = r[15];
    const hasOperationPrice =
      opPriceCell !== '' &&
      opPriceCell !== null &&
      opPriceCell !== undefined &&
      isFinite(Number(opPriceCell));

    const opPrice = hasOperationPrice ? Number(opPriceCell) : null;

    if (!article || !qty) return;

    const fallbackMeta = {
      name: name,
      type: type,
      category: category,
      unit: unit
    };

    if (operation === 'Приход') {
      if (!toObj) return;

      const stateIn = ensureState_(toObj, article, fallbackMeta);

      const incomingPrice = (opPrice !== null && opPrice >= 0)
        ? opPrice
        : (Number(catalogPriceMap[article]) || 0);

      stateIn.qty += qty;
      stateIn.value += qty * incomingPrice;
      return;
    }

    if (operation === 'Списание') {
      if (!fromObj) return;

      const stateOut = ensureState_(fromObj, article, fallbackMeta);
      const avgPrice = stateOut.qty > 0 ? stateOut.value / stateOut.qty : 0;

      stateOut.qty -= qty;
      stateOut.value -= qty * avgPrice;

      if (Math.abs(stateOut.qty) < 0.0000001) stateOut.qty = 0;
      if (Math.abs(stateOut.value) < 0.0000001) stateOut.value = 0;

      return;
    }

    if (operation === 'Перемещение') {
      if (!fromObj || !toObj) return;

      const stateFrom = ensureState_(fromObj, article, fallbackMeta);
      const transferPrice = stateFrom.qty > 0 ? stateFrom.value / stateFrom.qty : 0;

      stateFrom.qty -= qty;
      stateFrom.value -= qty * transferPrice;

      if (Math.abs(stateFrom.qty) < 0.0000001) stateFrom.qty = 0;
      if (Math.abs(stateFrom.value) < 0.0000001) stateFrom.value = 0;

      const stateTo = ensureState_(toObj, article, fallbackMeta);
      stateTo.qty += qty;
      stateTo.value += qty * transferPrice;

      return;
    }
  });

  const result = [];

  Object.keys(stockMap).forEach(function (key) {
    const state = stockMap[key];
    const qty = round3_(state.qty);

    if (qty === 0) return;

    const parts = key.split('||');
    const objectName = parts[0];
    const article = parts[1];
    const meta = metaMap[article] || {};

    const avgPrice = state.qty !== 0 ? state.value / state.qty : 0;
    const price = round2_(avgPrice);
    const sum = round2_(state.value);

    result.push([
      objectName,
      article,
      meta.name || '',
      meta.type || '',
      meta.category || '',
      meta.unit || '',
      qty,
      price,
      sum
    ]);
  });

  result.sort(function (a, b) {
    if (a[0] === b[0]) {
      return String(a[2]).localeCompare(String(b[2]), 'ru');
    }
    return String(a[0]).localeCompare(String(b[0]), 'ru');
  });

  if (result.length) {
    balances.getRange(2, 1, result.length, 9).setValues(result);
    balances.getRange(2, 7, result.length, 1).setNumberFormat('#,##0.###');
    balances.getRange(2, 8, result.length, 2).setNumberFormat('#,##0.00');
  }

  autoResize_(balances, 1, 9);
}

function refreshDashboard() {
  SpreadsheetApp.flush();
}

/**
 * =========================
 * HELPERS
 * =========================
 */
function getBalancesData_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const balances = ss.getSheetByName('Остатки');
  if (!balances) throw new Error('Не найден лист "Остатки".');

  const lastRow = balances.getLastRow();
  const data = lastRow >= 2
    ? balances.getRange(2, 1, lastRow - 1, 9).getValues()
    : [];

  return data.map(function (r) {
    return {
      objectName: String(r[0] || '').trim(),
      article: String(r[1] || '').trim(),
      name: String(r[2] || '').trim(),
      type: String(r[3] || '').trim(),
      category: String(r[4] || '').trim(),
      unit: String(r[5] || '').trim(),
      qty: Number(r[6]) || 0,
      price: Number(r[7]) || 0,
      sum: Number(r[8]) || 0
    };
  });
}

function getCatalogData_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const catalog = ss.getSheetByName('Номенклатура');
  if (!catalog) throw new Error('Не найден лист "Номенклатура".');

  const lastRow = catalog.getLastRow();
  if (lastRow < 2) return [];

  return catalog.getRange(2, 1, lastRow - 1, 8).getValues().map(function (r) {
    return {
      article: String(r[0] || '').trim(),
      type: String(r[1] || '').trim(),
      category: String(r[2] || '').trim(),
      name: String(r[3] || '').trim(),
      unit: String(r[4] || '').trim(),
      price: Number(r[5]) || 0,
      active: String(r[6] || '').trim(),
      comment: String(r[7] || '').trim()
    };
  });
}

function getCurrentStock(objectName, article) {
  const data = getBalancesData_();
  const objectNorm = String(objectName || '').trim();
  const articleNorm = String(article || '').trim().toLowerCase();

  for (var i = 0; i < data.length; i++) {
    if (
      String(data[i].objectName).trim() === objectNorm &&
      String(data[i].article).trim().toLowerCase() === articleNorm
    ) {
      return Number(data[i].qty) || 0;
    }
  }

  return 0;
}

function getProductNameByArticle(article) {
  const catalog = getCatalogData_();
  const articleNorm = String(article || '').trim().toLowerCase();

  for (var i = 0; i < catalog.length; i++) {
    if (String(catalog[i].article).trim().toLowerCase() === articleNorm) {
      return catalog[i].name || '';
    }
  }

  return '';
}

function ensureDirectoryValues_(rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dir = ss.getSheetByName('Справочник');
  if (!dir) throw new Error('Не найден лист "Справочник".');

  const existingTypes = getColumnValues(dir, 2, 2);
  const existingUnits = getColumnValues(dir, 6, 2);

  const lastRow = dir.getLastRow();
  const existingPairs = lastRow >= 2
    ? dir.getRange(2, 3, lastRow - 1, 2).getValues()
    : [];

  const typeSet = {};
  existingTypes.forEach(function (value) {
    typeSet[normalizeText_(value)] = true;
  });

  const unitSet = {};
  existingUnits.forEach(function (value) {
    unitSet[normalizeText_(value)] = true;
  });

  const pairSet = {};
  existingPairs.forEach(function (pair) {
    const type = normalizeText_(pair[0]);
    const category = normalizeText_(pair[1]);
    if (!type || !category) return;
    pairSet[type + '||' + category] = true;
  });

  const typesToAdd = [];
  const unitsToAdd = [];
  const pairsToAdd = [];

  rows.forEach(function (row) {
    const type = String(row.type || '').trim();
    const category = String(row.category || '').trim();
    const unit = String(row.unit || '').trim();

    const typeKey = normalizeText_(type);
    const categoryKey = normalizeText_(category);
    const unitKey = normalizeText_(unit);
    const pairKey = typeKey + '||' + categoryKey;

    if (type && !typeSet[typeKey]) {
      typeSet[typeKey] = true;
      typesToAdd.push([type]);
    }

    if (unit && !unitSet[unitKey]) {
      unitSet[unitKey] = true;
      unitsToAdd.push([unit]);
    }

    if (type && category && !pairSet[pairKey]) {
      pairSet[pairKey] = true;
      pairsToAdd.push([type, category]);
    }
  });

  if (typesToAdd.length) {
    const startRowTypes = getNextRowInColumn_(dir, 2, 2);
    dir.getRange(startRowTypes, 2, typesToAdd.length, 1).setValues(typesToAdd);
  }

  if (pairsToAdd.length) {
    const startRowPairs = getNextRowInColumn_(dir, 3, 2);
    dir.getRange(startRowPairs, 3, pairsToAdd.length, 2).setValues(pairsToAdd);
  }

  if (unitsToAdd.length) {
    const startRowUnits = getNextRowInColumn_(dir, 6, 2);
    dir.getRange(startRowUnits, 6, unitsToAdd.length, 1).setValues(unitsToAdd);
  }
}

function getColumnValues(sheet, columnNumber, startRow) {
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return [];

  return sheet
    .getRange(startRow, columnNumber, lastRow - startRow + 1, 1)
    .getValues()
    .flat()
    .map(function (value) {
      return String(value || '').trim();
    })
    .filter(Boolean);
}

function getNextRowInColumn_(sheet, columnNumber, startRow) {
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return startRow;

  const values = sheet
    .getRange(startRow, columnNumber, lastRow - startRow + 1, 1)
    .getValues()
    .flat();

  for (var i = values.length - 1; i >= 0; i--) {
    if (String(values[i] || '').trim() !== '') {
      return startRow + i + 1;
    }
  }

  return startRow;
}

function normalizeText_(value) {
  return String(value || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function formatHeader_(sheet, row, col, numRows, numCols) {
  sheet.getRange(row, col, numRows, numCols)
    .setFontWeight('bold')
    .setBackground('#d9eaf7')
    .setHorizontalAlignment('center');
}

function autoResize_(sheet, startCol, numCols) {
  for (var i = startCol; i < startCol + numCols; i++) {
    sheet.autoResizeColumn(i);
  }
}

function round2_(value) {
  return Math.round((Number(value) || 0) * 100) / 100;
}

function round3_(value) {
  return Math.round((Number(value) || 0) * 1000) / 1000;
}

function formatDate_(value) {
  if (!value) return '';
  return Utilities.formatDate(new Date(value), Session.getScriptTimeZone(), 'dd.MM.yyyy');
}

function parseDashboardDate_(value, endOfDay) {
  if (!value) return null;

  const d = new Date(value);
  if (isNaN(d.getTime())) return null;

  if (endOfDay) {
    d.setHours(23, 59, 59, 999);
  } else {
    d.setHours(0, 0, 0, 0);
  }

  return d;
}

function formatDateTimeRu_(value) {
  if (!(value instanceof Date)) return '';
  return Utilities.formatDate(
    value,
    Session.getScriptTimeZone(),
    'dd.MM.yyyy HH:mm'
  );
}

function formatQty_(value) {
  const num = Number(value) || 0;
  return Utilities.formatString('%s', round3_(num));
}
/**
 * =========================
 * STOREKEEPER DASHBOARD
 * =========================
 */

function showStorekeeperDashboard() {
  requirePermission_('storekeeperDashboard', 'дашборд кладовщика');
  const html = HtmlService.createHtmlOutputFromFile('StorekeeperDashboard')
    .setWidth(1500)
    .setHeight(950);
  SpreadsheetApp.getUi().showModalDialog(html, 'Дашборд кладовщика');
}

function getAssignedToolsData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const journal = ss.getSheetByName('Журнал движения');

  if (!journal) throw new Error('Не найден лист "Журнал движения".');

  const catalog = getCatalogData_();
  const catalogMap = {};

  catalog.forEach(function (item) {
    catalogMap[String(item.article || '').trim().toLowerCase()] = item;
  });

  const lastRow = journal.getLastRow();
  if (lastRow < 2) return [];

  const data = journal.getRange(2, 1, lastRow - 1, 15).getValues();
  const map = {};

  data.forEach(function (r) {
    const operation = String(r[1] || '').trim();
    const article = String(r[2] || '').trim();
    const qty = Number(r[7]) || 0;
    const fromObj = String(r[8] || '').trim();
    const toObj = String(r[9] || '').trim();
    const employee = String(r[11] || '').trim();
    const opDate = r[0] instanceof Date ? r[0] : null;

    if (!article || !qty || !employee) return;

    const isReturnOperation =
      operation === 'Возврат' ||
      operation === 'Снятие закрепления';

    const objectName = toObj || fromObj || '';
    const key = employee + '||' + article + '||' + objectName;

    if (!map[key]) {
      map[key] = {
        employee: employee,
        article: article,
        objectName: objectName,
        qty: 0,
        lastIssueDate: null
      };
    }

    if (isReturnOperation) {
      map[key].qty -= qty;
    } else {
      map[key].qty += qty;

      if (opDate) {
        if (!map[key].lastIssueDate || opDate > map[key].lastIssueDate) {
          map[key].lastIssueDate = opDate;
        }
      }
    }
  });

  return Object.keys(map)
    .map(function (key) {
      const row = map[key];
      const item = catalogMap[String(row.article || '').trim().toLowerCase()] || {};

      return {
        employee: row.employee,
        article: row.article,
        name: String(item.name || ''),
        type: String(item.type || ''),
        category: String(item.category || ''),
        objectName: row.objectName,
        qty: round3_(row.qty),
        lastIssueDate: row.lastIssueDate ? formatDateTimeRu_(row.lastIssueDate) : ''
      };
    })
    .filter(function (row) {
      return row.qty > 0;
    })
    .sort(function (a, b) {
      if (a.employee === b.employee) {
        return String(a.name).localeCompare(String(b.name), 'ru');
      }
      return String(a.employee).localeCompare(String(b.employee), 'ru');
    });
}

function saveStorekeeperOperation(payload) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    ensureJournalPriceColumn_();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const catalog = ss.getSheetByName('Номенклатура');
    const journal = ss.getSheetByName('Журнал движения');

    if (!catalog) throw new Error('Не найден лист "Номенклатура".');
    if (!journal) throw new Error('Не найден лист "Журнал движения".');

    const mode = String(payload.mode || 'existing').trim();
    const operation = String(payload.operation || '').trim();
    const objectName = String(payload.objectName || '').trim();
    const destination = String(payload.destination || '').trim();
    const basis = String(payload.basis || '').trim();
    const comment = String(payload.comment || '').trim();
    const qty = Number(payload.qty);

    const rawPrice = payload.price;
    const hasManualPrice =
      rawPrice !== '' &&
      rawPrice !== null &&
      rawPrice !== undefined &&
      isFinite(Number(rawPrice));

    const operationPrice = hasManualPrice ? Number(rawPrice) : null;

    const now = new Date();
    const user = Session.getActiveUser().getEmail() || '';

    if (!operation) throw new Error('Не выбрано действие.');
    if (!isFinite(qty) || qty <= 0) {
      throw new Error('Количество должно быть больше нуля.');
    }

    if (operation === 'Приход' && !objectName) {
      throw new Error('Укажи объект прихода.');
    }

    if ((operation === 'Списание' || operation === 'Перемещение') && !objectName) {
      throw new Error('Укажи объект.');
    }

    if (operation === 'Перемещение' && !destination) {
      throw new Error('Укажи куда перемещать.');
    }

    if (operation === 'Перемещение' && destination === objectName) {
      throw new Error('Объект назначения совпадает с текущим.');
    }

    let article = '';
    let name = '';
    let type = '';
    let category = '';
    let unit = '';
    let priceForJournal = '';

    if (mode === 'new') {
      if (operation !== 'Приход') {
        throw new Error('Для нового товара доступно только действие "Приход".');
      }

      name = String(payload.name || '').trim();
      type = String(payload.type || '').trim();
      category = String(payload.category || '').trim();
      unit = String(payload.unit || '').trim();

      if (!name) throw new Error('Укажи название нового товара.');
      if (!type) throw new Error('Укажи тип.');
      if (!category) throw new Error('Укажи категорию.');
      if (!unit) throw new Error('Укажи ед. изм.');

      if (operationPrice === null || operationPrice < 0) {
        throw new Error('Для нового товара обязательно укажи цену.');
      }

      const existingCatalog = getCatalogData_();
      const duplicate = existingCatalog.some(function (item) {
        return normalizeText_(item.name) === normalizeText_(name);
      });
      if (duplicate) {
        throw new Error('Такой товар уже есть в номенклатуре.');
      }

      ensureDirectoryValues_([{
        type: type,
        category: category,
        unit: unit
      }]);

      article = generateArticle(type);
      priceForJournal = operationPrice;

      catalog.getRange(catalog.getLastRow() + 1, 1, 1, 8).setValues([[
        article,
        type,
        category,
        name,
        unit,
        operationPrice,
        'Да',
        'Создано из дашборда кладовщика'
      ]]);
    } else {
      article = String(payload.article || '').trim();
      if (!article) throw new Error('Не выбран товар.');

      const item = getCatalogItemByArticle_(article);
      if (!item) throw new Error('Товар не найден в номенклатуре.');

      name = item.name;
      type = item.type;
      category = item.category;
      unit = item.unit;

      if (operation === 'Приход') {
        const catalogPrice = Number(item.price) || 0;
        const finalPrice = hasManualPrice ? operationPrice : catalogPrice;

        if (!isFinite(finalPrice) || finalPrice < 0) {
          throw new Error('Не удалось определить цену прихода.');
        }

        priceForJournal = finalPrice;

        if (hasManualPrice && Math.abs(finalPrice - catalogPrice) > 0.000001) {
          updateCatalogPriceByArticle_(article, finalPrice);
        }
      }

      if (operation === 'Списание' || operation === 'Перемещение') {
        refreshBalances();
        const stock = getCurrentStock(objectName, article);
        if (qty > stock) {
          throw new Error(
            'Недостаточно остатка. В наличии ' + formatQty_(stock) + ', указано ' + formatQty_(qty) + '.'
          );
        }
      }
    }

    let fromObj = '';
    let toObj = '';

    if (operation === 'Приход') {
      toObj = objectName;
    }

    if (operation === 'Списание') {
      fromObj = objectName;
    }

    if (operation === 'Перемещение') {
      fromObj = objectName;
      toObj = destination;
    }

    journal.getRange(journal.getLastRow() + 1, 1, 1, 16).setValues([[
      new Date(),
      operation,
      article,
      name,
      type,
      category,
      unit,
      qty,
      fromObj,
      toObj,
      basis || defaultBasisByOperation_(operation),
      '',
      comment || 'Дашборд кладовщика',
      user,
      now,
      priceForJournal
    ]]);

    refreshAll();

    return {
      ok: true,
      message: mode === 'new'
        ? 'Новый товар создан и операция проведена. Артикул: ' + article
        : 'Операция проведена.',
      article: article
    };
  } finally {
    lock.releaseLock();
  }
}

function updateCatalogItem(payload) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const catalog = ss.getSheetByName('Номенклатура');

    if (!catalog) throw new Error('Не найден лист "Номенклатура".');

    const article = String(payload.article || '').trim();
    const name = String(payload.name || '').trim();
    const type = String(payload.type || '').trim();
    const category = String(payload.category || '').trim();
    const unit = String(payload.unit || '').trim();
    const active = String(payload.active || 'Да').trim() || 'Да';
    const comment = String(payload.comment || '').trim();
    const price = Number(payload.price);

    if (!article) throw new Error('Не выбран товар.');
    if (!name) throw new Error('Укажи название.');
    if (!type) throw new Error('Укажи тип.');
    if (!category) throw new Error('Укажи категорию.');
    if (!unit) throw new Error('Укажи ед. изм.');
    if (!isFinite(price) || price < 0) throw new Error('Некорректная цена.');

    ensureDirectoryValues_([{
      type: type,
      category: category,
      unit: unit
    }]);

    const lastRow = catalog.getLastRow();
    if (lastRow < 2) throw new Error('Номенклатура пуста.');

    const data = catalog.getRange(2, 1, lastRow - 1, 8).getValues();
    let targetRow = 0;

    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0] || '').trim().toLowerCase() === article.toLowerCase()) {
        targetRow = i + 2;
        break;
      }
    }

    if (!targetRow) throw new Error('Товар не найден в номенклатуре.');

    catalog.getRange(targetRow, 2, 1, 7).setValues([[
      type,
      category,
      name,
      unit,
      price,
      active,
      comment
    ]]);

    refreshAll();
    return 'Карточка товара обновлена.';
  } finally {
    lock.releaseLock();
  }
}

function getCategoriesByType_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dir = ss.getSheetByName('Справочник');
  if (!dir) throw new Error('Не найден лист "Справочник".');

  const lastRow = dir.getLastRow();
  const pairs = lastRow >= 2
    ? dir.getRange(2, 3, lastRow - 1, 2).getValues()
    : [];

  const result = {};
  pairs.forEach(function (row) {
    const type = String(row[0] || '').trim();
    const category = String(row[1] || '').trim();
    if (!type || !category) return;
    if (!result[type]) result[type] = [];
    if (result[type].indexOf(category) === -1) {
      result[type].push(category);
    }
  });

  return result;
}

function getCatalogItemByArticle_(article) {
  const items = getCatalogData_();
  const articleNorm = String(article || '').trim().toLowerCase();

  for (var i = 0; i < items.length; i++) {
    if (String(items[i].article || '').trim().toLowerCase() === articleNorm) {
      return items[i];
    }
  }

  return null;
}

function defaultBasisByOperation_(operation) {
  if (operation === 'Приход') return 'Приход кладовщика';
  if (operation === 'Списание') return 'Списание кладовщика';
  if (operation === 'Перемещение') return 'Перемещение кладовщика';
  return 'Операция кладовщика';
}

/***********************
 * PATCH: ответственность, массовое перемещение, справочники
 ***********************/

function buildArticleCounters_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const catalog = ss.getSheetByName('Номенклатура');
  if (!catalog) throw new Error('Не найден лист "Номенклатура".');

  const counters = {};
  const lastRow = catalog.getLastRow();
  if (lastRow < 2) return counters;

  const articles = catalog.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  articles.forEach(function (value) {
    const article = String(value || '').trim().toLowerCase();
    if (!article) return;
    const parts = article.split('-');
    if (parts.length < 2) return;
    const prefix = parts[0];
    const match = article.match(/(\d+)$/);
    if (!match) return;
    const num = Number(match[1]) || 0;
    counters[prefix] = Math.max(counters[prefix] || 0, num);
  });
  return counters;
}

function getArticlePrefix_(type) {
  const prefixMap = {
    'Материал': 'мат',
    'Инструмент': 'инс',
    'Расходник': 'рас',
    'Оборудование': 'обо'
  };
  return prefixMap[String(type || '').trim()] || 'тов';
}

function createArticleGenerator_() {
  const counters = buildArticleCounters_();
  return function (type) {
    const prefix = getArticlePrefix_(type);
    counters[prefix] = (counters[prefix] || 0) + 1;
    return prefix + '-' + Utilities.formatString('%04d', counters[prefix]);
  };
}

function generateArticle(type) {
  const nextArticle = createArticleGenerator_();
  return nextArticle(type);
}

function ensureResponsibilitySheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let current = ss.getSheetByName('Ответственные');
  if (!current) {
    current = ss.insertSheet('Ответственные');
  }
  if (current.getLastRow() === 0) {
    current.getRange(1, 1, 1, 6).setValues([['Артикул', 'Ответственный', 'Назначено', 'Кем', 'Комментарий', 'Объект']]);
    formatHeader_(current, 1, 1, 1, 6);
    current.setFrozenRows(1);
  } else {
    const header = current.getRange(1, 1, 1, Math.max(current.getLastColumn(), 6)).getValues()[0];
    const objectHeader = String(header[5] || '').trim();
    if (!objectHeader) {
      current.getRange(1, 6).setValue('Объект');
    }
  }

  let history = ss.getSheetByName('История ответственности');
  if (!history) {
    history = ss.insertSheet('История ответственности');
  }
  if (history.getLastRow() === 0) {
    history.getRange(1, 1, 1, 11).setValues([[
      'Дата', 'Действие', 'Артикул', 'Название', 'Тип', 'Категория',
      'Объект', 'Количество', 'Был ответственный', 'Новый ответственный', 'Кем'
    ]]);
    formatHeader_(history, 1, 1, 1, 11);
    history.setFrozenRows(1);
  }
}

function getResponsibilityKey_(article, objectName) {
  return normalizeText_(String(article || '').trim()) + '||' + normalizeText_(String(objectName || '').trim());
}

function getEmployees_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dir = ss.getSheetByName('Справочник');
  if (!dir) throw new Error('Не найден лист "Справочник".');

  const lastRow = dir.getLastRow();
  if (lastRow < 2) return [];

  const employeesE = dir.getRange(2, 5, lastRow - 1, 1).getValues().flat();
  const employeesG = dir.getRange(2, 7, lastRow - 1, 1).getValues().flat();
  const map = {};
  const result = [];

  employeesG.concat(employeesE).forEach(function (value) {
    const name = String(value || '').trim();
    const key = normalizeText_(name);
    if (!key || map[key]) return;
    map[key] = true;
    result.push(name);
  });

  result.sort(function (a, b) {
    return String(a).localeCompare(String(b), 'ru');
  });

  return result;
}

function syncEmployeesDirectory_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dir = ss.getSheetByName('Справочник');
  if (!dir) throw new Error('Не найден лист "Справочник".');

  const lastRow = Math.max(dir.getLastRow(), 2);
  const employees = getEmployees_();

  if (String(dir.getRange(1, 7).getValue() || '').trim() === '') {
    dir.getRange(1, 7).setValue('Сотрудники');
  }

  if (lastRow >= 2) {
    dir.getRange(2, 7, lastRow - 1, 1).clearContent();
    dir.getRange(2, 5, lastRow - 1, 1).clearContent();
  }

  if (employees.length) {
    dir.getRange(2, 7, employees.length, 1).setValues(
      employees.map(function (name) { return [name]; })
    );
  }
}

function getCurrentResponsibilityMap_() {
  ensureResponsibilitySheets_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Ответственные');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { byObject: {}, byArticle: {} };

  const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  const map = {};
  const legacyByArticle = {};
  data.forEach(function (r) {
    const article = String(r[0] || '').trim().toLowerCase();
    const employee = String(r[1] || '').trim();
    const objectName = String(r[5] || '').trim();
    if (!article || !employee) return;
    const item = {
      article: String(r[0] || '').trim(),
      objectName: objectName,
      employee: employee,
      assignedAt: formatDateTimeRu_(r[2]),
      assignedBy: String(r[3] || '').trim(),
      comment: String(r[4] || '').trim()
    };
    map[getResponsibilityKey_(article, objectName)] = item;
    if (!objectName && !legacyByArticle[article]) {
      legacyByArticle[article] = item;
    }
  });
  return {
    byObject: map,
    byArticle: legacyByArticle
  };
}

function appendResponsibilityHistory_(items, action, oldEmployee, newEmployee, user, comment) {
  ensureResponsibilitySheets_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('История ответственности');
  const rows = items.map(function (item) {
    return [
      new Date(),
      action,
      item.article,
      item.name || getProductNameByArticle(item.article),
      item.type || '',
      item.category || '',
      item.objectName || '',
      Number(item.qty) || 0,
      oldEmployee || '',
      newEmployee || '',
      user || ''
    ];
  });
  if (rows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 11).setValues(rows);
  }
}

function getResponsibilityHistoryData_(filters) {
  ensureResponsibilitySheets_();
  filters = filters || {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('История ответственности');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const q = normalizeText_(filters.query || '');
  const objectFilter = String(filters.objectName || '').trim();
  const typeFilter = String(filters.typeName || '').trim();
  const employeeFilter = String(filters.employee || '').trim();

  const data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
  return data.map(function (r) {
    const oldEmployee = String(r[8] || '').trim();
    const newEmployee = String(r[9] || '').trim();
    return {
      date: formatDateTimeRu_(r[0]),
      action: String(r[1] || '').trim(),
      article: String(r[2] || '').trim(),
      name: String(r[3] || '').trim(),
      type: String(r[4] || '').trim(),
      category: String(r[5] || '').trim(),
      objectName: String(r[6] || '').trim(),
      qty: round3_(Number(r[7]) || 0),
      oldEmployee: oldEmployee,
      newEmployee: newEmployee,
      employee: newEmployee || oldEmployee,
      comment: '',
      user: String(r[10] || '').trim()
    };
  }).filter(function (row) {
    if (objectFilter && row.objectName !== objectFilter) return false;
    if (typeFilter && row.type !== typeFilter) return false;
    if (employeeFilter && row.employee !== employeeFilter && row.newEmployee !== employeeFilter && row.oldEmployee !== employeeFilter) return false;
    if (q) {
      const text = normalizeText_([
        row.article, row.name, row.type, row.category, row.objectName, row.employee, row.oldEmployee, row.newEmployee, row.user
      ].join(' '));
      if (text.indexOf(q) === -1) return false;
    }
    return true;
  });
}

function getCurrentResponsibilityRows_(filters) {
  filters = filters || {};
  const q = normalizeText_(filters.query || '');
  const objectFilter = String(filters.objectName || '').trim();
  const typeFilter = String(filters.typeName || '').trim();
  const employeeFilter = String(filters.employee || '').trim();
  const onlyAssigned = filters.onlyAssigned === true;
  const onlyUnassigned = filters.onlyUnassigned === true;

  const respMap = getCurrentResponsibilityMap_();

  let rows = getBalancesData_()
    .filter(function (item) {
      return Number(item.qty) > 0;
    })
    .map(function (item) {
      const article = String(item.article || '').trim();
      const objectName = String(item.objectName || '').trim();
      const resp = respMap.byObject[getResponsibilityKey_(article, objectName)] || null;
      return {
        objectName: String(item.objectName || '').trim(),
        article: String(item.article || '').trim(),
        name: String(item.name || '').trim(),
        type: String(item.type || '').trim(),
        category: String(item.category || '').trim(),
        unit: String(item.unit || '').trim(),
        qty: round3_(Number(item.qty) || 0),
        employee: resp ? resp.employee : '',
        assignedAt: resp ? resp.assignedAt : '',
        assignedBy: resp ? resp.assignedBy : '',
        comment: resp ? String(resp.comment || '') : ''
      };
    });

  if (objectFilter) rows = rows.filter(function (r) { return r.objectName === objectFilter; });
  if (typeFilter) rows = rows.filter(function (r) { return r.type === typeFilter; });
  if (employeeFilter) rows = rows.filter(function (r) { return r.employee === employeeFilter; });
  if (onlyAssigned) rows = rows.filter(function (r) { return !!r.employee; });
  if (onlyUnassigned) rows = rows.filter(function (r) { return !r.employee; });

  if (q) {
    rows = rows.filter(function (r) {
      return normalizeText_([r.article, r.name, r.type, r.category, r.objectName, r.employee].join(' ')).indexOf(q) !== -1;
    });
  }

  rows.sort(function (a, b) {
    if (a.objectName === b.objectName) {
      return String(a.name).localeCompare(String(b.name), 'ru');
    }
    return String(a.objectName).localeCompare(String(b.objectName), 'ru');
  });

  return rows;
}

function getUnassignedInstrumentRows(filters) {
  return getCurrentResponsibilityRows_(Object.assign({}, filters || {}, { onlyUnassigned: true }));
}

function getAssignedInstrumentRows(filters) {
  return getCurrentResponsibilityRows_(Object.assign({}, filters || {}, { onlyAssigned: true }));
}

function saveResponsibilityAssignments(payload) {
  ensureResponsibilitySheets_();

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    const employee = String(payload.employee || '').trim();
    const rows = Array.isArray(payload.rows) ? payload.rows : [];
    const comment = String(payload.comment || '').trim();
    if (!employee) throw new Error('Выбери сотрудника.');
    if (!rows.length) throw new Error('Не выбраны позиции для назначения.');

    const employees = getEmployees_().map(normalizeText_);
    if (employees.indexOf(normalizeText_(employee)) === -1) {
      throw new Error('Сотрудник не найден в справочнике.');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Ответственные');
    const user = Session.getActiveUser().getEmail() || '';
    const currentMap = getCurrentResponsibilityMap_();

    const rowsToAppend = [];
    const historyItems = [];

    rows.forEach(function (row) {
      const article = String(row.article || '').trim();
      const objectName = String(row.objectName || '').trim();
      if (!article) return;
      if (!objectName) {
        throw new Error('Не найден объект для товара ' + article + '.');
      }
      const current = currentMap.byObject[getResponsibilityKey_(article, objectName)] || null;
      if (current && current.employee) {
        throw new Error('У товара ' + article + ' на объекте "' + objectName + '" уже назначен ответственный.');
      }
      rowsToAppend.push([article, employee, new Date(), user, comment, objectName]);
      historyItems.push(Object.assign({}, row, { objectName: objectName }));
    });

    if (!rowsToAppend.length) throw new Error('Нет строк для назначения.');
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, 6).setValues(rowsToAppend);
    appendResponsibilityHistory_(historyItems, 'Назначение', '', employee, user, comment);
    refreshAssigned();
    return 'Назначено позиций: ' + rowsToAppend.length;
  } finally {
    lock.releaseLock();
  }
}

function saveResponsibilityReassignments(payload) {
  ensureResponsibilitySheets_();

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    const employee = String(payload.employee || '').trim();
    const rows = Array.isArray(payload.rows) ? payload.rows : [];
    const comment = String(payload.comment || '').trim();
    if (!employee) throw new Error('Выбери нового ответственного.');
    if (!rows.length) throw new Error('Не выбраны позиции для переназначения.');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Ответственные');
    const currentMap = getCurrentResponsibilityMap_();
    const user = Session.getActiveUser().getEmail() || '';

    const data = sheet.getLastRow() >= 2 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues() : [];
    const historyItems = [];

    rows.forEach(function (row) {
      const article = String(row.article || '').trim();
      const objectName = String(row.objectName || '').trim();
      if (!objectName) {
        throw new Error('Не найден объект для товара ' + article + '.');
      }
      const norm = article.toLowerCase();
      const current = currentMap.byObject[getResponsibilityKey_(article, objectName)] || null;
      if (!current || !current.employee) {
        throw new Error('У товара ' + article + ' на объекте "' + objectName + '" нет текущего ответственного.');
      }

      let targetRow = 0;
      for (var i = 0; i < data.length; i++) {
        const rowArticle = String(data[i][0] || '').trim().toLowerCase();
        const rowObject = String(data[i][5] || '').trim();
        if (rowArticle === norm && normalizeText_(rowObject) === normalizeText_(objectName)) {
          targetRow = i + 2;
          break;
        }
      }
      if (!targetRow) throw new Error('Не найдена строка ответственного для ' + article + ' на объекте "' + objectName + '"');

      sheet.getRange(targetRow, 2, 1, 4).setValues([[employee, new Date(), user, comment]]);
      historyItems.push(Object.assign({}, row, { objectName: objectName, oldEmployee: current.employee }));
    });

    appendResponsibilityHistory_(
      historyItems.map(function (r) { return r; }),
      'Переназначение',
      '',
      employee,
      user,
      comment
    );

    refreshAssigned();
    return 'Переназначено позиций: ' + rows.length;
  } finally {
    lock.releaseLock();
  }
}

function saveStorekeeperMassTransfer(payload) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    ensureJournalPriceColumn_();
    ensureResponsibilitySheets_();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const journal = ss.getSheetByName('Журнал движения');
    if (!journal) throw new Error('Не найден лист "Журнал движения".');

    const sourceObject = String(payload.sourceObject || '').trim();
    const destinationObject = String(payload.destinationObject || '').trim();
    const basis = String(payload.basis || '').trim();
    const comment = String(payload.comment || '').trim();
    const newEmployee = String(payload.newEmployee || '').trim();
    const rows = Array.isArray(payload.rows) ? payload.rows : [];

    if (!sourceObject) throw new Error('Укажи объект-источник.');
    if (!destinationObject) throw new Error('Укажи объект назначения.');
    if (sourceObject === destinationObject) throw new Error('Объекты совпадают.');
    if (!rows.length) throw new Error('Не выбраны товары для перемещения.');

    refreshBalances();

    const user = Session.getActiveUser().getEmail() || '';
    const now = new Date();
    const journalRows = [];
    const reassignedItems = [];

    rows.forEach(function (row, idx) {
      const article = String(row.article || '').trim();
      const qty = Number(row.qty);
      if (!article) throw new Error('Строка ' + (idx + 1) + ': не найден артикул.');
      if (!isFinite(qty) || qty <= 0) throw new Error('Строка ' + (idx + 1) + ': количество должно быть больше нуля.');

      const item = getCatalogItemByArticle_(article);
      if (!item) throw new Error('Товар ' + article + ' не найден в номенклатуре.');

      const stock = getCurrentStock(sourceObject, article);
      if (qty > stock) {
        throw new Error(
          'Строка ' + (idx + 1) + ': нельзя переместить больше чем есть. В наличии ' +
          formatQty_(stock) + ', указано ' + formatQty_(qty) + '.'
        );
      }

      journalRows.push([
        new Date(),
        'Перемещение',
        article,
        item.name,
        item.type,
        item.category,
        item.unit,
        qty,
        sourceObject,
        destinationObject,
        basis || 'Массовое перемещение',
        '',
        comment || 'Дашборд кладовщика / массовое перемещение',
        user,
        now,
        ''
      ]);

      reassignedItems.push({
        article: article,
        name: item.name,
        type: item.type,
        category: item.category,
        objectName: destinationObject,
        qty: qty
      });
    });

    journal.getRange(journal.getLastRow() + 1, 1, journalRows.length, 16).setValues(journalRows);

    if (newEmployee) {
      const currentMap = getCurrentResponsibilityMap_();
      const assignRows = [];
      const reassignRows = [];

      reassignedItems.forEach(function (item) {
        const current = currentMap.byObject[getResponsibilityKey_(item.article, item.objectName)] || null;
        if (current && current.employee) {
          reassignRows.push(item);
        } else {
          assignRows.push(item);
        }
      });

      if (assignRows.length) {
        saveResponsibilityAssignments({ employee: newEmployee, rows: assignRows, comment: 'Назначено при массовом перемещении' });
      }
      if (reassignRows.length) {
        saveResponsibilityReassignments({ employee: newEmployee, rows: reassignRows, comment: 'Переназначено при массовом перемещении' });
      }
    }

    refreshAll();
    return 'Перемещено позиций: ' + rows.length;
  } finally {
    lock.releaseLock();
  }
}

function deleteNewItemRequests(payload) {
  ensureRequestsSheet_();
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const requestsSheet = ss.getSheetByName('Запросы номенклатуры');
    if (!requestsSheet) throw new Error('Не найден лист "Запросы номенклатуры".');

    const rows = Array.isArray(payload.rows) ? payload.rows : [];
    if (!rows.length) throw new Error('Нет строк для удаления.');

    const lastRow = requestsSheet.getLastRow();
    if (lastRow < 2) throw new Error('Лист запросов пуст.');

    const data = requestsSheet.getRange(2, 1, lastRow - 1, 17).getValues();
    const allowedRows = {};
    data.forEach(function (r, index) {
      const status = String(r[9] || '').trim();
      if (status === 'Новый' || status === 'В работе') {
        allowedRows[index + 2] = true;
      }
    });

    const rowsToDelete = rows
      .map(function (row) { return Number(row.sheetRow) || 0; })
      .filter(function (sheetRow) { return sheetRow && allowedRows[sheetRow]; })
      .sort(function (a, b) { return b - a; });

    if (!rowsToDelete.length) throw new Error('Нет подходящих строк для удаления.');

    rowsToDelete.forEach(function (sheetRow) { requestsSheet.deleteRow(sheetRow); });
    return 'Удалено запросов: ' + rowsToDelete.length;
  } finally {
    lock.releaseLock();
  }
}

function approveNewItemRequests(payload) {
  ensureRequestsSheet_();

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    ensureJournalPriceColumn_();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const requestsSheet = ss.getSheetByName('Запросы номенклатуры');
    const catalog = ss.getSheetByName('Номенклатура');
    const journal = ss.getSheetByName('Журнал движения');

    if (!requestsSheet) throw new Error('Не найден лист "Запросы номенклатуры".');
    if (!catalog) throw new Error('Не найден лист "Номенклатура".');
    if (!journal) throw new Error('Не найден лист "Журнал движения".');

    const rows = Array.isArray(payload.rows) ? payload.rows : [];
    if (!rows.length) throw new Error('Нет строк для подтверждения.');

    const errors = [];
    const catalogRows = [];
    const journalRows = [];
    const updates = [];
    const now = new Date();
    const user = Session.getActiveUser().getEmail() || '';

    const existingCatalog = getCatalogData_();
    const existingNameMap = {};
    const batchNameMap = {};

    existingCatalog.forEach(function (item) {
      existingNameMap[normalizeText_(item.name)] = item;
    });

    rows.forEach(function (row, index) {
      const line = index + 1;
      const sheetRow = Number(row.sheetRow);
      const requestId = String(row.requestId || '').trim();
      const objectName = String(row.objectName || '').trim();
      const commonBasis = String(row.commonBasis || '').trim();
      const originalName = String(row.name || '').trim();
      const finalName = String(row.finalName || originalName).trim();
      const unit = String(row.unit || '').trim();
      const type = String(row.type || '').trim();
      const category = String(row.category || '').trim();
      const qty = Number(row.qty);
      const price = Number(row.price);
      const storekeeperComment = String(row.storekeeperComment || '').trim();

      if (!sheetRow) errors.push('Строка ' + line + ': не найден номер строки заявки.');
      if (!requestId) errors.push('Строка ' + line + ': не найден ID заявки.');
      if (!objectName) errors.push('Строка ' + line + ': не найден объект.');
      if (!finalName) errors.push('Строка ' + line + ': не заполнено итоговое название.');
      if (!unit) errors.push('Строка ' + line + ': не заполнена ед. изм.');
      if (!type) errors.push('Строка ' + line + ': не заполнен тип.');
      if (!category) errors.push('Строка ' + line + ': не заполнена категория.');
      if (!isFinite(qty) || qty <= 0) errors.push('Строка ' + line + ': количество должно быть больше нуля.');
      if (!isFinite(price) || price < 0) errors.push('Строка ' + line + ': цена некорректна.');

      const finalNameKey = normalizeText_(finalName);
      if (existingNameMap[finalNameKey]) errors.push('Строка ' + line + ': товар уже есть в номенклатуре — "' + finalName + '".');
      if (batchNameMap[finalNameKey]) errors.push('Строка ' + line + ': дубль внутри подтверждения — "' + finalName + '".');
      batchNameMap[finalNameKey] = true;

      updates.push({
        sheetRow: sheetRow,
        requestId: requestId,
        originalName: originalName,
        finalName: finalName,
        objectName: objectName,
        commonBasis: commonBasis,
        unit: unit,
        type: type,
        category: category,
        qty: qty,
        price: price,
        storekeeperComment: storekeeperComment
      });
    });

    if (errors.length) throw new Error(errors.join('\n'));

    ensureDirectoryValues_(updates.map(function (r) {
      return { type: r.type, category: r.category, unit: r.unit };
    }));

    const nextArticle = createArticleGenerator_();

    updates.forEach(function (row) {
      const article = nextArticle(row.type);

      catalogRows.push([
        article, row.type, row.category, row.finalName, row.unit, row.price, 'Да',
        'Создано из запроса ' + row.requestId
      ]);

      journalRows.push([
        new Date(), 'Приход', article, row.finalName, row.type, row.category, row.unit, row.qty,
        '', row.objectName, row.commonBasis || ('Запрос новой номенклатуры ' + row.requestId), '',
        row.storekeeperComment || 'Подтверждено из запроса', user, now, row.price
      ]);

      requestsSheet.getRange(row.sheetRow, 10, 1, 8).setValues([[
        'Подтвержден', row.type, row.category, row.finalName, row.storekeeperComment, article, now, user
      ]]);
    });

    if (catalogRows.length) {
      catalog.getRange(catalog.getLastRow() + 1, 1, catalogRows.length, 8).setValues(catalogRows);
    }
    if (journalRows.length) {
      journal.getRange(journal.getLastRow() + 1, 1, journalRows.length, 16).setValues(journalRows);
    }

    refreshAll();
    return 'Подтверждено заявок: ' + updates.length;
  } finally {
    lock.releaseLock();
  }
}

function saveImportForm(payload) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    ensureJournalPriceColumn_();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const catalog = ss.getSheetByName('Номенклатура');
    const journal = ss.getSheetByName('Журнал движения');
    const dir = ss.getSheetByName('Справочник');

    if (!catalog) throw new Error('Не найден лист "Номенклатура".');
    if (!journal) throw new Error('Не найден лист "Журнал движения".');
    if (!dir) throw new Error('Не найден лист "Справочник".');

    const rows = Array.isArray(payload.rows) ? payload.rows : [];
    const createReceipt = Boolean(payload.createReceipt);
    const objectName = String(payload.objectName || '').trim();
    const commonBasis = String(payload.commonBasis || '').trim();

    if (!rows.length) throw new Error('Нет строк для загрузки.');
    if (createReceipt && !objectName) throw new Error('Укажи объект для прихода.');

    const existingCatalog = getCatalogData_();
    const existingNameMap = {};
    existingCatalog.forEach(function (item) {
      existingNameMap[normalizeText_(item.name)] = item;
    });

    const errors = [];
    const duplicatesInsideImport = {};
    const preparedRows = [];

    rows.forEach(function (row, index) {
      const line = index + 1;
      const name = String(row.name || '').trim();
      const type = String(row.type || '').trim();
      const category = String(row.category || '').trim();
      const unit = String(row.unit || '').trim();
      const price = Number(row.price);
      const qty = Number(row.qty);

      if (!name) errors.push('Строка ' + line + ': не заполнено название товара.');
      if (!type) errors.push('Строка ' + line + ': не заполнен тип.');
      if (!category) errors.push('Строка ' + line + ': не заполнена категория.');
      if (!unit) errors.push('Строка ' + line + ': не заполнена ед. изм.');
      if (!isFinite(price) || price < 0) errors.push('Строка ' + line + ': некорректная цена.');
      if (createReceipt && (!isFinite(qty) || qty <= 0)) errors.push('Строка ' + line + ': количество должно быть больше нуля.');

      const normalizedName = normalizeText_(name);
      if (normalizedName) {
        if (duplicatesInsideImport[normalizedName]) {
          errors.push('Строка ' + line + ': дубль внутри импорта — "' + name + '".');
        } else {
          duplicatesInsideImport[normalizedName] = true;
        }
        if (existingNameMap[normalizedName]) {
          errors.push('Строка ' + line + ': товар уже есть в номенклатуре — "' + name + '".');
        }
      }

      preparedRows.push({
        line: line, name: name, type: type, category: category, unit: unit,
        price: isFinite(price) ? price : 0, qty: isFinite(qty) ? qty : 0
      });
    });

    if (errors.length) throw new Error(errors.join('\n'));

    ensureDirectoryValues_(preparedRows);

    const now = new Date();
    const user = Session.getActiveUser().getEmail() || '';
    const catalogRows = [];
    const journalRows = [];
    const nextArticle = createArticleGenerator_();

    preparedRows.forEach(function (item) {
      const article = nextArticle(item.type);

      catalogRows.push([article, item.type, item.category, item.name, item.unit, item.price, 'Да', 'Импорт']);

      if (createReceipt) {
        journalRows.push([
          new Date(), 'Приход', article, item.name, item.type, item.category, item.unit, item.qty, '',
          objectName, commonBasis || 'Импорт номенклатуры', '', 'Новый товар (импорт)', user, now, item.price
        ]);
      }
    });

    if (catalogRows.length) {
      catalog.getRange(catalog.getLastRow() + 1, 1, catalogRows.length, 8).setValues(catalogRows);
    }
    if (journalRows.length) {
      journal.getRange(journal.getLastRow() + 1, 1, journalRows.length, 16).setValues(journalRows);
    }

    refreshAll();
    return createReceipt
      ? 'Импорт завершён. Добавлено в номенклатуру: ' + catalogRows.length + '. Создано приходов: ' + journalRows.length + '.'
      : 'Импорт завершён. Добавлено в номенклатуру: ' + catalogRows.length + '.';
  } finally {
    lock.releaseLock();
  }
}

function saveStorekeeperBulkNewItems(payload) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    ensureJournalPriceColumn_();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const catalog = ss.getSheetByName('Номенклатура');
    const journal = ss.getSheetByName('Журнал движения');

    if (!catalog) throw new Error('Не найден лист "Номенклатура".');
    if (!journal) throw new Error('Не найден лист "Журнал движения".');

    const objectName = String(payload.objectName || '').trim();
    const commonBasis = String(payload.commonBasis || '').trim();
    const rows = Array.isArray(payload.rows) ? payload.rows : [];

    if (!objectName) throw new Error('Укажи объект прихода.');
    if (!rows.length) throw new Error('Нет строк для внесения.');

    const existingCatalog = getCatalogData_();
    const existingNameMap = {};
    existingCatalog.forEach(function (item) {
      existingNameMap[normalizeText_(item.name)] = item;
    });

    const errors = [];
    const duplicatesInsideBatch = {};
    const preparedRows = [];

    rows.forEach(function (row, index) {
      const line = index + 1;
      const name = String(row.name || '').trim();
      const type = String(row.type || '').trim();
      const category = String(row.category || '').trim();
      const unit = String(row.unit || '').trim();
      const qty = Number(row.qty);
      const price = Number(row.price);

      if (!name) errors.push('Строка ' + line + ': не заполнено название.');
      if (!type) errors.push('Строка ' + line + ': не заполнен тип.');
      if (!category) errors.push('Строка ' + line + ': не заполнена категория.');
      if (!unit) errors.push('Строка ' + line + ': не заполнена ед. изм.');
      if (!isFinite(qty) || qty <= 0) errors.push('Строка ' + line + ': количество должно быть больше нуля.');
      if (!isFinite(price) || price < 0) errors.push('Строка ' + line + ': некорректная цена.');

      const nameKey = normalizeText_(name);
      if (nameKey) {
        if (existingNameMap[nameKey]) errors.push('Строка ' + line + ': товар уже есть в номенклатуре — "' + name + '".');
        if (duplicatesInsideBatch[nameKey]) errors.push('Строка ' + line + ': дубль внутри внесения — "' + name + '".');
        duplicatesInsideBatch[nameKey] = true;
      }

      preparedRows.push({ name: name, type: type, category: category, unit: unit, qty: qty, price: price });
    });

    if (errors.length) throw new Error(errors.join('\n'));

    ensureDirectoryValues_(preparedRows);

    const user = Session.getActiveUser().getEmail() || '';
    const now = new Date();
    const nextArticle = createArticleGenerator_();
    const catalogRows = [];
    const journalRows = [];

    preparedRows.forEach(function (row) {
      const article = nextArticle(row.type);

      catalogRows.push([
        article, row.type, row.category, row.name, row.unit, row.price, 'Да',
        'Создано из массового внесения кладовщика'
      ]);

      journalRows.push([
        new Date(), 'Приход', article, row.name, row.type, row.category, row.unit, row.qty, '',
        objectName, commonBasis || 'Массовое внесение кладовщиком', '',
        'Дашборд кладовщика / массовое внесение', user, now, row.price
      ]);
    });

    if (catalogRows.length) {
      catalog.getRange(catalog.getLastRow() + 1, 1, catalogRows.length, 8).setValues(catalogRows);
    }
    if (journalRows.length) {
      journal.getRange(journal.getLastRow() + 1, 1, journalRows.length, 16).setValues(journalRows);
    }

    refreshAll();
    return 'Добавлено новых позиций: ' + catalogRows.length + '. Приходы созданы.';
  } finally {
    lock.releaseLock();
  }
}

function getStorekeeperDashboardInitData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dir = ss.getSheetByName('Справочник');
  if (!dir) throw new Error('Не найден лист "Справочник".');

  ensureResponsibilitySheets_();
  syncEmployeesDirectory_();

  const currentResponsibilities = getAssignedInstrumentRows({});
  const balances = getBalancesData_();

  return {
    objects: getColumnValues(dir, 1, 2),
    types: getColumnValues(dir, 2, 2),
    units: getColumnValues(dir, 6, 2),
    employees: getEmployees_(),
    categoriesByType: getCategoriesByType_(),
    catalogItems: getCatalogData_(),
    balances: balances,
    currentResponsibilities: currentResponsibilities,
    assignedCount: currentResponsibilities.length
  };
}

function getResponsibilityDashboardData(filters) {
  return {
    current: getAssignedInstrumentRows(filters || {}),
    available: getUnassignedInstrumentRows(filters || {}),
    history: getResponsibilityHistoryData_(filters || {})
  };
}

function saveStorekeeperDirectoryItem(payload) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    const kind = String(payload.kind || '').trim();
    const typeName = String(payload.typeName || '').trim();
    const categoryName = String(payload.categoryName || '').trim();
    const unitName = String(payload.unitName || '').trim();
    const objectName = String(payload.objectName || '').trim();
    const employeeName = String(payload.employeeName || '').trim();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dir = ss.getSheetByName('Справочник');
    if (!dir) throw new Error('Не найден лист "Справочник".');

    if (kind === 'object') {
      if (!objectName) throw new Error('Укажи название объекта.');
      const existing = getColumnValues(dir, 1, 2).map(normalizeText_);
      if (existing.indexOf(normalizeText_(objectName)) !== -1) throw new Error('Такой объект уже есть.');
      const row = getNextRowInColumn_(dir, 1, 2);
      dir.getRange(row, 1).setValue(objectName);
      return 'Объект добавлен.';
    }

    if (kind === 'employee') {
      if (!employeeName) throw new Error('Укажи сотрудника.');
      const existingEmployees = getEmployees_().map(normalizeText_);
      if (existingEmployees.indexOf(normalizeText_(employeeName)) !== -1) throw new Error('Такой сотрудник уже есть.');
      const row = getNextRowInColumn_(dir, 7, 2);
      dir.getRange(row, 7).setValue(employeeName);
      syncEmployeesDirectory_();
      return 'Сотрудник добавлен.';
    }

    if (kind === 'type') {
      if (!typeName) throw new Error('Укажи название типа.');
      const existingTypes = getColumnValues(dir, 2, 2).map(normalizeText_);
      if (existingTypes.indexOf(normalizeText_(typeName)) !== -1) throw new Error('Такой тип уже есть.');
      const row = getNextRowInColumn_(dir, 2, 2);
      dir.getRange(row, 2).setValue(typeName);
      return 'Тип добавлен.';
    }

    if (kind === 'category') {
      if (!typeName) throw new Error('Выбери тип.');
      if (!categoryName) throw new Error('Укажи категорию.');

      const pairs = dir.getLastRow() >= 2 ? dir.getRange(2, 3, dir.getLastRow() - 1, 2).getValues() : [];
      const exists = pairs.some(function (r) {
        return normalizeText_(r[0]) === normalizeText_(typeName) &&
               normalizeText_(r[1]) === normalizeText_(categoryName);
      });

      if (exists) throw new Error('Такая категория уже привязана к этому типу.');

      const row = getNextRowInColumn_(dir, 3, 2);
      dir.getRange(row, 3, 1, 2).setValues([[typeName, categoryName]]);
      return 'Категория добавлена.';
    }

    if (kind === 'unit') {
      if (!unitName) throw new Error('Укажи ед. изм.');
      const existingUnits = getColumnValues(dir, 6, 2).map(normalizeText_);
      if (existingUnits.indexOf(normalizeText_(unitName)) !== -1) throw new Error('Такая ед. изм. уже есть.');
      const row = getNextRowInColumn_(dir, 6, 2);
      dir.getRange(row, 6).setValue(unitName);
      return 'Ед. изм. добавлена.';
    }

    throw new Error('Неизвестный вид справочника.');
  } finally {
    lock.releaseLock();
  }
}

function deleteStorekeeperDirectoryItem(payload) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    const kind = String(payload.kind || '').trim();
    const value = String(payload.value || '').trim();
    if (!value) throw new Error('Не выбрано значение для удаления.');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dir = ss.getSheetByName('Справочник');
    if (!dir) throw new Error('Не найден лист "Справочник".');

    if (kind === 'employee') {
      const lastRow = dir.getLastRow();
      if (lastRow >= 2) {
        const dataE = dir.getRange(2, 5, lastRow - 1, 1).getValues();
        const dataG = dir.getRange(2, 7, lastRow - 1, 1).getValues();

        for (var i = dataE.length - 1; i >= 0; i--) {
          if (normalizeText_(dataE[i][0]) === normalizeText_(value)) {
            dir.getRange(i + 2, 5).clearContent();
          }
          if (normalizeText_(dataG[i][0]) === normalizeText_(value)) {
            dir.getRange(i + 2, 7).clearContent();
          }
        }
      }

      syncEmployeesDirectory_();
      return 'Сотрудник удалён.';
    }

    let startCol = 0;
    if (kind === 'object') startCol = 1;
    if (kind === 'type') startCol = 2;
    if (kind === 'unit') startCol = 6;

    if (kind === 'category') {
      const typeName = String(payload.typeName || '').trim();
      if (!typeName) throw new Error('Для удаления категории выбери тип.');
      const lastRow = dir.getLastRow();
      const data = lastRow >= 2 ? dir.getRange(2, 3, lastRow - 1, 2).getValues() : [];

      for (var j = data.length - 1; j >= 0; j--) {
        if (normalizeText_(data[j][0]) === normalizeText_(typeName) &&
            normalizeText_(data[j][1]) === normalizeText_(value)) {
          dir.getRange(j + 2, 3, 1, 2).clearContent();
          return 'Категория удалена.';
        }
      }
      throw new Error('Категория не найдена.');
    }

    if (!startCol) throw new Error('Неизвестный вид справочника.');

    const lastRow = dir.getLastRow();
    const data = lastRow >= 2 ? dir.getRange(2, startCol, lastRow - 1, 1).getValues() : [];

    for (var k = data.length - 1; k >= 0; k--) {
      if (normalizeText_(data[k][0]) === normalizeText_(value)) {
        dir.getRange(k + 2, startCol).clearContent();
        return 'Значение удалено.';
      }
    }

    throw new Error('Значение не найдено.');
  } finally {
    lock.releaseLock();
  }
}

function refreshAssigned() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const assigned = ss.getSheetByName('Закреплено');
  if (!assigned) return;

  const rows = getAssignedInstrumentRows({});
  assigned.clearContents();

  const headers = [['Ответственный', 'Артикул', 'Название товара', 'Объект', 'Количество', 'Тип', 'Категория']];
  assigned.getRange(1, 1, 1, headers[0].length).setValues(headers);
  formatHeader_(assigned, 1, 1, 1, headers[0].length);
  assigned.setFrozenRows(1);

  if (rows.length) {
    assigned.getRange(2, 1, rows.length, 7).setValues(
      rows.map(function (r) {
        return [r.employee, r.article, r.name, r.objectName, r.qty, r.type, r.category];
      })
    );
    assigned.getRange(2, 5, rows.length, 1).setNumberFormat('#,##0.###');
  }

  autoResize_(assigned, 1, 7);
}

function refreshAll() {
  ensureResponsibilitySheets_();
  syncEmployeesDirectory_();
  refreshBalances();
  refreshAssigned();
  refreshDashboard();
}

/**
 * =========================
 * FLEET & TRIPS DASHBOARD
 * =========================
 */
function ensureFleetTripsSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const specs = [
    {
      name: 'Автопарк',
      headers: ['ID', 'Автомобиль', 'Госномер', 'Водитель', 'Пробег', 'Ближайшее ТО', 'Комментарий', 'Создано', 'Кем', 'Тип авто']
    },
    {
      name: 'ТО и ремонты',
      headers: ['ID', 'Дата', 'Автомобиль', 'Тип работы', 'Пробег', 'Что делалось', 'Сумма', 'Комментарий', 'Следующее ТО', 'Создано', 'Кем']
    },
    {
      name: 'Поездки автопарк',
      headers: ['ID', 'Дата', 'Автомобиль', 'Сотрудник', 'Тип поездки', 'Объект', 'Маршрут', 'Цель', 'Километраж', 'Расходы', 'Комментарий', 'Создано', 'Кем']
    },
    {
      name: 'Списания накоплений',
      headers: ['ID', 'Дата', 'Сотрудник', 'Сумма', 'Основание', 'Комментарий', 'Создано', 'Кем']
    }
  ];

  specs.forEach(function (spec) {
    let sh = ss.getSheetByName(spec.name);
    if (!sh) sh = ss.insertSheet(spec.name);
    const lastCol = spec.headers.length;
    const header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    const isHeaderEmpty = header.every(function (v) { return String(v || '').trim() === ''; });
    if (isHeaderEmpty) {
      sh.getRange(1, 1, 1, lastCol).setValues([spec.headers]);
      formatHeader_(sh, 1, 1, 1, lastCol);
      sh.setFrozenRows(1);
      autoResize_(sh, 1, lastCol);
    } else if (spec.name === 'Автопарк' && String(header[9] || '').trim() !== 'Тип авто') {
      sh.getRange(1, 10).setValue('Тип авто');
    }
  });
}

function toIsoDate_(value) {
  if (!value) return '';
  const d = value instanceof Date ? value : new Date(value);
  if (!d || isNaN(d.getTime())) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function getFleetTripsInitData() {
  requirePermission_('fleetDashboard', 'данные панели автопарка и поездок');
  ensureFleetTripsSheets_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fleetSh = ss.getSheetByName('Автопарк');
  const serviceSh = ss.getSheetByName('ТО и ремонты');
  const tripSh = ss.getSheetByName('Поездки автопарк');
  const writeoffSh = ss.getSheetByName('Списания накоплений');
  const dir = ss.getSheetByName('Справочник');

  const fleetRows = fleetSh.getLastRow() >= 2 ? fleetSh.getRange(2, 1, fleetSh.getLastRow() - 1, 10).getValues() : [];
  const serviceRows = serviceSh.getLastRow() >= 2 ? serviceSh.getRange(2, 1, serviceSh.getLastRow() - 1, 11).getValues() : [];
  const tripRows = tripSh.getLastRow() >= 2 ? tripSh.getRange(2, 1, tripSh.getLastRow() - 1, 13).getValues() : [];
  const writeoffRows = writeoffSh.getLastRow() >= 2 ? writeoffSh.getRange(2, 1, writeoffSh.getLastRow() - 1, 8).getValues() : [];

  const objects = dir ? getColumnValues(dir, 1, 2) : [];
  const employees = dir ? getEmployees_() : [];

  const servicesByCar = {};
  serviceRows.forEach(function (r) {
    const car = String(r[2] || '').trim();
    if (!car) return;
    if (!servicesByCar[car]) servicesByCar[car] = [];
    servicesByCar[car].push({
      id: String(r[0] || ''),
      date: toIsoDate_(r[1]),
      car: car,
      workType: String(r[3] || '').trim(),
      mileage: Number(r[4]) || 0,
      works: String(r[5] || '').trim(),
      amount: Number(r[6]) || 0,
      comment: String(r[7] || '').trim(),
      nextMaintenance: Number(r[8]) || 0
    });
  });

  const trips = tripRows.map(function (r) {
    return {
      id: String(r[0] || ''),
      date: toIsoDate_(r[1]),
      car: String(r[2] || '').trim(),
      employee: String(r[3] || '').trim(),
      tripType: String(r[4] || '').trim(),
      objectName: String(r[5] || '').trim(),
      route: String(r[6] || '').trim(),
      purpose: String(r[7] || '').trim(),
      mileage: Number(r[8]) || 0,
      cost: Number(r[9]) || 0,
      comment: String(r[10] || '').trim()
    };
  });

  const accumulationMap = {};
  trips.forEach(function (trip) {
    const key = String(trip.employee || '').trim() || 'Без сотрудника';
    if (!accumulationMap[key]) {
      accumulationMap[key] = { employee: key, accrued: 0, writtenOff: 0, available: 0 };
    }
    accumulationMap[key].accrued += Number(trip.cost) || 0;
  });

  const writeoffs = writeoffRows.map(function (r) {
    return {
      id: String(r[0] || '').trim(),
      date: toIsoDate_(r[1]),
      employee: String(r[2] || '').trim(),
      amount: Number(r[3]) || 0,
      reason: String(r[4] || '').trim(),
      comment: String(r[5] || '').trim()
    };
  });

  writeoffs.forEach(function (row) {
    const key = String(row.employee || '').trim() || 'Без сотрудника';
    if (!accumulationMap[key]) {
      accumulationMap[key] = { employee: key, accrued: 0, writtenOff: 0, available: 0 };
    }
    accumulationMap[key].writtenOff += Number(row.amount) || 0;
  });

  employees.forEach(function (employee) {
    const key = String(employee || '').trim();
    if (!key) return;
    if (!accumulationMap[key]) {
      accumulationMap[key] = { employee: key, accrued: 0, writtenOff: 0, available: 0 };
    }
  });

  Object.keys(accumulationMap).forEach(function (employee) {
    const entry = accumulationMap[employee];
    entry.accrued = round2_(entry.accrued);
    entry.writtenOff = round2_(entry.writtenOff);
    entry.available = round2_(Math.max(0, entry.accrued - entry.writtenOff));
  });

  const cars = fleetRows.map(function (r) {
    const carName = String(r[1] || '').trim();
    const list = servicesByCar[carName] || [];
    list.sort(function (a, b) { return String(b.date).localeCompare(String(a.date)); });
    const lastService = list[0] || null;
    const totalAmount = list.reduce(function (sum, row) { return sum + (Number(row.amount) || 0); }, 0);
    const mileage = Number(r[4]) || 0;
    const nextTO = Number(r[5]) || 0;
    let status = 'В норме';
    if (nextTO && mileage >= nextTO) status = 'Просрочено';
    else if (nextTO && nextTO - mileage <= 3000) status = 'Скоро ТО';

    return {
      id: String(r[0] || ''),
      name: carName,
      plate: String(r[2] || '').trim(),
      driver: String(r[3] || '').trim(),
      mileage: mileage,
      nextMaintenance: nextTO,
      comment: String(r[6] || '').trim(),
      isPersonal: String(r[9] || '').trim() === 'personal',
      lastServiceDate: lastService ? lastService.date : '',
      totalMaintenanceCost: totalAmount,
      worksDone: lastService ? lastService.works : '',
      status: status
    };
  });

  const history = serviceRows.map(function (r) {
    return {
      id: String(r[0] || ''),
      date: toIsoDate_(r[1]),
      car: String(r[2] || '').trim(),
      workType: String(r[3] || '').trim(),
      works: String(r[5] || '').trim(),
      mileage: Number(r[4]) || 0,
      amount: Number(r[6]) || 0,
      comment: String(r[7] || '').trim()
    };
  });

  return {
    employees: employees,
    objects: objects,
    cars: cars,
    services: serviceRows.length ? Object.keys(servicesByCar).reduce(function (acc, car) { return acc.concat(servicesByCar[car]); }, []) : [],
    trips: trips,
    history: history,
    accumulations: Object.keys(accumulationMap).map(function (key) { return accumulationMap[key]; }),
    writeoffs: writeoffs
  };
}

function saveFleetVehicle(payload) {
  requirePermission_('fleetDashboard', 'добавление автомобиля');
  ensureFleetTripsSheets_();

  const car = String(payload && payload.name || '').trim();
  const plate = String(payload && payload.plate || '').trim();
  const driver = String(payload && payload.driver || '').trim();
  const mileage = Number(payload && payload.mileage);
  const nextTO = Number(payload && payload.nextMaintenance);
  const comment = String(payload && payload.comment || '').trim();
  const ownershipType = String(payload && payload.ownershipType || 'company').trim() === 'personal' ? 'personal' : 'company';

  if (!car) throw new Error('Укажи название автомобиля.');
  if (!plate) throw new Error('Укажи госномер.');
  if (!isFinite(mileage) || mileage < 0) throw new Error('Пробег должен быть числом 0 или больше.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Автопарк');
  const now = new Date();
  const user = getCurrentUserEmail_();
  const id = 'CAR-' + Utilities.getUuid().slice(0, 8).toUpperCase();

  sh.appendRow([id, car, plate, driver, mileage, isFinite(nextTO) ? nextTO : '', comment, now, user, ownershipType]);
  return 'Автомобиль добавлен: ' + car;
}

function saveFleetService(payload) {
  requirePermission_('fleetDashboard', 'добавление ТО/ремонта');
  ensureFleetTripsSheets_();

  const car = String(payload && payload.car || '').trim();
  const date = payload && payload.date ? new Date(payload.date) : new Date();
  const workType = String(payload && payload.workType || '').trim();
  const mileage = Number(payload && payload.mileage);
  const works = String(payload && payload.works || '').trim();
  const amount = Number(payload && payload.amount);
  const comment = String(payload && payload.comment || '').trim();
  const nextTO = Number(payload && payload.nextMaintenance);

  if (!car) throw new Error('Выбери автомобиль.');
  if (!workType) throw new Error('Укажи тип работы.');
  if (!isFinite(mileage) || mileage < 0) throw new Error('Пробег должен быть числом 0 или больше.');
  if (!works) throw new Error('Опиши, что делалось.');
  if (!isFinite(amount) || amount < 0) throw new Error('Сумма должна быть числом 0 или больше.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('ТО и ремонты');
  const fleet = ss.getSheetByName('Автопарк');
  const now = new Date();
  const user = getCurrentUserEmail_();
  const id = 'SRV-' + Utilities.getUuid().slice(0, 8).toUpperCase();

  let carExists = false;
  if (fleet.getLastRow() >= 2) {
    const fleetData = fleet.getRange(2, 1, fleet.getLastRow() - 1, 2).getValues();
    for (var j = 0; j < fleetData.length; j++) {
      if (String(fleetData[j][1] || '').trim() === car) {
        carExists = true;
        break;
      }
    }
  }
  if (!carExists) throw new Error('Автомобиль для ТО/ремонта должен быть из списка добавленных автомобилей.');

  sh.appendRow([id, date, car, workType, mileage, works, amount, comment, isFinite(nextTO) ? nextTO : '', now, user]);

  if (fleet.getLastRow() >= 2) {
    const fleetData = fleet.getRange(2, 1, fleet.getLastRow() - 1, 9).getValues();
    for (var i = 0; i < fleetData.length; i++) {
      if (String(fleetData[i][1] || '').trim() !== car) continue;
      const row = i + 2;
      const currentMileage = Number(fleetData[i][4]) || 0;
      if (mileage > currentMileage) fleet.getRange(row, 5).setValue(mileage);
      if (isFinite(nextTO) && nextTO > 0) fleet.getRange(row, 6).setValue(nextTO);
      break;
    }
  }

  return 'ТО/ремонт сохранен для: ' + car;
}

function saveFleetTrip(payload) {
  requirePermission_('fleetDashboard', 'добавление поездки');
  ensureFleetTripsSheets_();

  const date = payload && payload.date ? new Date(payload.date) : new Date();
  const car = String(payload && payload.car || '').trim();
  const employee = String(payload && payload.employee || '').trim();
  const route = String(payload && payload.route || '').trim();
  const purpose = String(payload && payload.purpose || '').trim();
  const mileage = Number(payload && payload.mileage);
  const cost = mileage * 1.5;

  if (!car) throw new Error('Выбери автомобиль.');
  if (!employee) throw new Error('Укажи сотрудника.');
  if (!route) throw new Error('Укажи маршрут.');
  if (!isFinite(mileage) || mileage <= 0) throw new Error('Километраж должен быть больше 0.');
  if (!isFinite(cost) || cost < 0) throw new Error('Не удалось рассчитать сумму поездки.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Поездки автопарк');
  const now = new Date();
  const user = getCurrentUserEmail_();

  const id = 'TRP-' + Utilities.getUuid().slice(0, 8).toUpperCase();

  sh.appendRow([id, date, car, employee, 'Личный автомобиль', '', route, purpose, mileage, cost, '', now, user]);

  return 'Поездка сохранена.';
}

function deleteFleetTrip(tripId) {
  requirePermission_('fleetDashboard', 'удаление поездки');
  ensureFleetTripsSheets_();

  const id = String(tripId || '').trim();
  if (!id) throw new Error('Не передан идентификатор поездки.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Поездки автопарк');
  if (sh.getLastRow() < 2) throw new Error('Журнал поездок пуст.');

  const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0] || '').trim() !== id) continue;
    sh.deleteRow(i + 2);
    return 'Поездка удалена.';
  }

  throw new Error('Поездка не найдена.');
}

function getEmployeeAccumulationsMap_() {
  const data = getFleetTripsInitData();
  const map = {};
  (data.accumulations || []).forEach(function (row) {
    map[row.employee] = row;
  });
  return map;
}

function saveFleetAccumulationWriteoff(payload) {
  requirePermission_('fleetDashboard', 'списание накоплений сотрудников');
  ensureFleetTripsSheets_();

  const employee = String(payload && payload.employee || '').trim();
  const amount = Number(payload && payload.amount);
  const reason = String(payload && payload.reason || '').trim();
  const comment = String(payload && payload.comment || '').trim();
  const writeoffDate = payload && payload.date ? new Date(payload.date) : new Date();

  if (!employee) throw new Error('Укажи сотрудника.');
  if (!isFinite(amount) || amount <= 0) throw new Error('Сумма списания должна быть больше 0.');
  if (!reason) throw new Error('Укажи основание списания (например: автомагазин, замена масла).');

  const accumulations = getEmployeeAccumulationsMap_();
  const available = Number(accumulations[employee] && accumulations[employee].available) || 0;
  if (amount > available) {
    throw new Error('Недостаточно накоплений. Доступно: ' + fmtMoney_(available) + ', запрошено: ' + fmtMoney_(amount) + '.');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Списания накоплений');
  const now = new Date();
  const user = getCurrentUserEmail_();
  const id = 'WOF-' + Utilities.getUuid().slice(0, 8).toUpperCase();

  sh.appendRow([id, writeoffDate, employee, round2_(amount), reason, comment, now, user]);
  return 'Списание сохранено: ' + employee + ' — ' + fmtMoney_(amount);
}

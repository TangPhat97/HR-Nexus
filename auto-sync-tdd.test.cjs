const assert = require('assert');
const crypto = require('crypto');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

function createUtilitiesStub() {
  return {
    DigestAlgorithm: {
      SHA_256: 'sha256',
    },
    formatDate(date, timezone, format) {
      const value = new Date(date);
      if (Number.isNaN(value.getTime())) {
        return '';
      }
      const pad = (input) => String(input).padStart(2, '0');
      const weekdayNames = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
      switch (format) {
        case 'EEE':
          return weekdayNames[value.getDay()];
        case 'H':
          return String(value.getHours());
        case 'm':
          return String(value.getMinutes());
        case 'yyyy-MM-dd':
          return [value.getFullYear(), pad(value.getMonth() + 1), pad(value.getDate())].join('-');
        default:
          return [value.getFullYear(), pad(value.getMonth() + 1), pad(value.getDate())].join('-');
      }
    },
    computeDigest(algorithm, payload) {
      const buffer = Buffer.isBuffer(payload) ? payload : Buffer.from(String(payload));
      return Array.from(crypto.createHash(algorithm).update(buffer).digest());
    },
  };
}

function createContext() {
  const documentPropertyStore = {};
  const menuItems = [];
  const alerts = [];
  const context = {
    console,
    require,
    Buffer,
    Date,
    JSON,
    Math,
    __menuItems: menuItems,
    __alerts: alerts,
    Logger: { log() {} },
    Utilities: createUtilitiesStub(),
    SpreadsheetApp: {
      getUi() {
        return {
          alert(title, message) {
            alerts.push({ title, message });
          },
          prompt() {
            return {
              getSelectedButton() { return 'OK'; },
              getResponseText() { return ''; },
            };
          },
          ButtonSet: { OK: 'OK', OK_CANCEL: 'OK_CANCEL' },
          Button: { OK: 'OK' },
          createMenu() {
            return {
              addItem(label) { menuItems.push(label); return this; },
              addSeparator() { return this; },
              addToUi() { return this; },
            };
          },
        };
      },
      getActiveSpreadsheet() {
        return { getId() { return 'SPREADSHEET001'; } };
      },
      ProtectionType: { SHEET: 'SHEET' },
    },
    HtmlService: {
      SandboxMode: { IFRAME: 'IFRAME' },
      createTemplateFromFile() {
        return {
          evaluate() {
            return {
              setTitle() { return this; },
              setSandboxMode() { return this; },
            };
          },
        };
      },
    },
    Session: {
      getScriptTimeZone() { return 'Asia/Bangkok'; },
      getActiveUser() { return { getEmail() { return 'tester@example.com'; } }; },
      getEffectiveUser() { return { getEmail() { return 'tester@example.com'; } }; },
    },
    PropertiesService: {
      getDocumentProperties() {
        return {
          getProperty(key) { return documentPropertyStore[key] || ''; },
          setProperty(key, value) { documentPropertyStore[key] = String(value); },
          deleteProperty(key) { delete documentPropertyStore[key]; },
        };
      },
    },
    LockService: {
      getDocumentLock() {
        return {
          waitLock() {},
          releaseLock() {},
        };
      },
    },
    ScriptApp: {
      getProjectTriggers() { return []; },
      newTrigger() {
        return {
          timeBased() { return this; },
          after() { return this; },
          everyMinutes() { return this; },
          create() {},
        };
      },
      deleteTrigger() {},
    },
    Drive: {},
    DriveApp: {},
    MimeType: {},
  };

  context.global = context;
  return vm.createContext(context);
}

function loadGasFiles(context, fileNames) {
  fileNames.forEach((fileName) => {
    const filePath = path.join(process.cwd(), 'gas', fileName);
    const code = fs.readFileSync(filePath, 'utf8');
    vm.runInContext(code, context, { filename: filePath });
  });
}

function runTests() {
  const context = createContext();
  loadGasFiles(context, [
    'Constants.gs',
    'Utils.gs',
    'ErrorHandler.gs',
    'Repository.gs',
    'RuntimeConfigService.gs',
    'MenuService.gs',
    'AppController.gs',
  ]);

  const APP_CONFIG = vm.runInContext('APP_CONFIG', context);
  const defaultConfigKeys = APP_CONFIG.DEFAULT_CONFIG.map((entry) => entry.config_key);

  [
    'AUTO_SYNC_ENABLED',
    'AUTO_SYNC_WEEKDAYS',
    'AUTO_SYNC_HOUR',
    'AUTO_SYNC_MINUTE',
    'AUTO_SYNC_TIMEZONE',
    'AUTO_SYNC_MODE',
    'AUTO_SYNC_NOTIFY_EMAIL',
    'AUTO_SYNC_LAST_RUN_AT',
    'AUTO_SYNC_LAST_STATUS',
    'AUTO_SYNC_LAST_MESSAGE',
    'AUTO_SYNC_LAST_SLOT_KEY',
    'AUTO_SYNC_SCHEDULER_ENABLED',
    'AUTO_SYNC_SCHEDULER_INTERVAL_MINUTES',
  ].forEach((key) => {
    assert.ok(defaultConfigKeys.includes(key), 'DEFAULT_CONFIG should seed ' + key);
  });

  assert.strictEqual(typeof context.getAutoSyncDefaultConfigEntries_, 'function', 'Auto sync should expose default config entries');
  assert.strictEqual(typeof context.ensureAutoSyncConfigRows_, 'function', 'Auto sync should expose config row seeding helper');
  assert.strictEqual(typeof context.getAutoSyncRuntimeConfig_, 'function', 'Auto sync should expose runtime config parser');
  assert.strictEqual(typeof context.runAutoSyncScheduler, 'function', 'AppController should expose the auto sync scheduler entrypoint');
  assert.strictEqual(typeof context.showAutoSyncStatus, 'function', 'AppController should expose a user-facing auto sync status action');
  assert.strictEqual(typeof context.buildAutoSyncStatusLines_, 'function', 'AppController should expose a formatter for auto sync status lines');

  const appendedRows = [];
  context.getConfigMap_ = function () {
    return {
      RAW_SYNC_DIRECT_THRESHOLD: '75',
      RAW_SYNC_CHUNK_SIZE: '420',
      RAW_SYNC_RESUME_ENABLED: 'true',
      RAW_SYNC_QUEUE_ENABLED: 'true',
      RAW_SYNC_QUEUE_DELAY_SECONDS: '15',
    };
  };
  context.appendObjects_ = function (sheetKey, rows) {
    appendedRows.push({ sheetKey, rows });
  };

  context.ensureAutoSyncConfigRows_();

  assert.strictEqual(appendedRows.length, 1, 'Auto sync config seeding should append missing rows once');
  assert.strictEqual(appendedRows[0].sheetKey, 'CONFIG_SYSTEM', 'Auto sync config seeding should target CONFIG_SYSTEM');
  assert.ok(
    appendedRows[0].rows.some((row) => row.config_key === 'AUTO_SYNC_WEEKDAYS' && row.config_value === 'MON'),
    'Auto sync config seeding should provide a safe default weekday'
  );
  assert.ok(
    appendedRows[0].rows.some((row) => row.config_key === 'AUTO_SYNC_ENABLED' && row.config_value === 'false'),
    'Auto sync config seeding should default auto sync to disabled on live systems'
  );

  const seedWrites = [];
  context.ensureSheet_ = function () {
    return {
      getLastRow() { return 1; },
      getRange() {
        return {
          setValues(values) { seedWrites.push(values); },
        };
      },
    };
  };
  context.readObjectsFromSheet_ = function (sheetKey) {
    if (sheetKey === 'CONFIG_SYSTEM' || sheetKey === 'CONFIG_USERS') {
      return [];
    }
    return [];
  };
  context.getConfigMap_ = function () { return {}; };
  context.getCurrentUserEmail_ = function () { return 'tester@example.com'; };
  context.appendObjects_ = function () {};
  context.seedSystemDefaults_();

  const seededConfigKeys = seedWrites.flat().map((row) => row[0]);
  assert.ok(
    seededConfigKeys.includes('AUTO_SYNC_ENABLED'),
    'System setup should seed AUTO_SYNC_* rows into CONFIG_SYSTEM during initialization'
  );

  const setupOrder = [];
  context.runAction_ = function (actionName, executor) {
    try {
      return executor();
    } catch (error) {
      throw error;
    }
  };
  context.seedSystemDefaults_ = function () { setupOrder.push('seed-defaults'); };
  context.ensureAutoSyncConfigRows_ = function () { setupOrder.push('seed-auto-sync'); };
  context.ensureAutoSyncSchedulerInstalled_ = function () { setupOrder.push('install-scheduler'); };
  context.setupWorkbookTopology_ = function () {
    setupOrder.push('topology');
    throw new Error('Sheet Bắt đầu không khớp schema v3.');
  };
  context.refreshAnalyticsCore_ = function () { setupOrder.push('refresh-analytics'); };
  context.showAlert_ = function () {};

  context.setupSystem();

  assert.deepStrictEqual(
    setupOrder.slice(0, 3),
    ['seed-defaults', 'seed-auto-sync', 'topology'],
    'setupSystem should seed runtime config before topology validation can fail on system sheets'
  );

  context.getConfigMap_ = function () {
    return {
      AUTO_SYNC_ENABLED: 'true',
      AUTO_SYNC_WEEKDAYS: 'MON,WED,FRI',
      AUTO_SYNC_HOUR: '22',
      AUTO_SYNC_MINUTE: '15',
      AUTO_SYNC_TIMEZONE: 'Asia/Bangkok',
      AUTO_SYNC_MODE: 'quick',
      AUTO_SYNC_NOTIFY_EMAIL: 'ops@example.com',
      AUTO_SYNC_LAST_RUN_AT: '2026-03-30T22:15:00',
      AUTO_SYNC_LAST_STATUS: 'QUEUED',
      AUTO_SYNC_LAST_MESSAGE: 'Da dua job auto sync vao queue.',
      AUTO_SYNC_LAST_SLOT_KEY: '2026-03-30|22|15|quick',
      AUTO_SYNC_SCHEDULER_ENABLED: 'true',
      AUTO_SYNC_SCHEDULER_INTERVAL_MINUTES: '5',
    };
  };

  const runtimeConfig = context.getAutoSyncRuntimeConfig_();
  assert.strictEqual(runtimeConfig.enabled, true, 'Auto sync runtime config should parse enabled flag');
  assert.strictEqual(runtimeConfig.mode, 'quick', 'Auto sync runtime config should normalize mode');
  assert.deepStrictEqual(Array.from(runtimeConfig.weekdays), ['MON', 'WED', 'FRI'], 'Auto sync runtime config should parse multiple weekdays');
  assert.strictEqual(runtimeConfig.hour, 22, 'Auto sync runtime config should parse scheduled hour');
  assert.strictEqual(runtimeConfig.minute, 15, 'Auto sync runtime config should parse scheduled minute');
  assert.strictEqual(runtimeConfig.timezone, 'Asia/Bangkok', 'Auto sync runtime config should keep configured timezone');
  assert.strictEqual(runtimeConfig.schedulerEnabled, true, 'Auto sync runtime config should parse scheduler enable flag');
  assert.strictEqual(runtimeConfig.schedulerIntervalMinutes, 5, 'Auto sync runtime config should parse scheduler polling interval');

  assert.strictEqual(
    context.isAutoSyncWeekdayEnabled_('WED', runtimeConfig),
    true,
    'Auto sync should detect enabled weekdays from a comma-separated config'
  );
  assert.strictEqual(
    context.isAutoSyncWeekdayEnabled_('SUN', runtimeConfig),
    false,
    'Auto sync should reject weekdays outside the configured list'
  );

  const statusLines = context.buildAutoSyncStatusLines_(runtimeConfig, 3);
  assert.ok(statusLines.some((line) => /MON, WED, FRI/.test(line)), 'Auto sync status should show configured weekdays');
  assert.ok(statusLines.some((line) => /22:15/.test(line)), 'Auto sync status should show configured time');
  assert.ok(statusLines.some((line) => /3/.test(line)), 'Auto sync status should show queue backlog');

  context.buildMenu_();
  assert.ok(context.__menuItems.includes('Kiểm tra auto sync'), 'Menu should expose a dedicated auto sync status action');
}

try {
  runTests();
  console.log('PASS auto-sync-tdd');
} catch (error) {
  console.error('FAIL auto-sync-tdd');
  console.error(error.stack || error.message);
  process.exit(1);
}

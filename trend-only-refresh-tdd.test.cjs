const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

function createUtilitiesStub() {
  return {
    formatDate(date, timezone, format) {
      const value = new Date(date);
      if (Number.isNaN(value.getTime())) {
        return '';
      }
      const pad = (input) => String(input).padStart(2, '0');
      switch (format) {
        case "yyyy-MM-dd'T'HH:mm:ss":
          return [
            value.getFullYear(),
            pad(value.getMonth() + 1),
            pad(value.getDate()),
          ].join('-') + 'T' + [pad(value.getHours()), pad(value.getMinutes()), pad(value.getSeconds())].join(':');
        case 'yyyy-MM':
          return [value.getFullYear(), pad(value.getMonth() + 1)].join('-');
        case 'yyyy-MM-dd':
        default:
          return [value.getFullYear(), pad(value.getMonth() + 1), pad(value.getDate())].join('-');
      }
    },
  };
}

function createSheetStub(name) {
  const sheet = {
    name,
    filterRange: null,
    rangeCalls: [],
    maxRows: 200,
    maxColumns: 30,
    getName() { return this.name; },
    getLastRow() { return 0; },
    getLastColumn() { return 0; },
    getMaxRows() { return this.maxRows; },
    getMaxColumns() { return this.maxColumns; },
    clearContents() { return this; },
    clearFormats() { return this; },
    setTabColor() { return this; },
    setFrozenRows() { return this; },
    setColumnWidth() { return this; },
    getFilter() {
      if (!this.filterRange) {
        return null;
      }
      const sheetRef = this;
      return { remove() { sheetRef.filterRange = null; } };
    },
    getRange(row, column, numRows, numColumns) {
      const call = { row, column, numRows, numColumns, values: null };
      this.rangeCalls.push(call);
      const sheetRef = this;
      const chain = {
        breakApart() { return chain; },
        clearContent() { return chain; },
        setValues(values) { call.values = values; return chain; },
        setBackground() { return chain; },
        setFontColor() { return chain; },
        setFontWeight() { return chain; },
        setHorizontalAlignment() { return chain; },
        setVerticalAlignment() { return chain; },
        setNumberFormat() { return chain; },
        setWrap() { return chain; },
        setWrapStrategy() { return chain; },
        setBorder() { return chain; },
        mergeAcross() { call.merged = true; return chain; },
        setFontSize() { return chain; },
        setBackgrounds() { return chain; },
        createFilter() { sheetRef.filterRange = { row, column, numRows, numColumns }; return chain; },
      };
      return chain;
    },
  };
  return sheet;
}

function createContext() {
  const menuItems = [];
  const alerts = [];
  const sheets = {};

  const context = {
    console,
    require,
    Buffer,
    Date,
    JSON,
    Math,
    Logger: { log() {} },
    Utilities: createUtilitiesStub(),
    SpreadsheetApp: {
      getActiveSpreadsheet() {
        return { getId() { return 'SPREADSHEET001'; } };
      },
      ProtectionType: { SHEET: 'SHEET' },
      getUi() {
        return {
          alert(title, message) {
            alerts.push({ title, message });
          },
          createMenu() {
            return {
              addItem(label) {
                menuItems.push(label);
                return this;
              },
              addSeparator() { return this; },
              addToUi() { return this; },
            };
          },
        };
      },
    },
    Session: {
      getActiveUser() { return { getEmail() { return 'tester@example.com'; } }; },
      getEffectiveUser() { return { getEmail() { return 'tester@example.com'; } }; },
      getScriptTimeZone() { return 'Asia/Bangkok'; },
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
    PropertiesService: {
      getDocumentProperties() {
        return {
          getProperty() { return ''; },
          setProperty() {},
          deleteProperty() {},
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
    __menuItems: menuItems,
    __alerts: alerts,
    __sheets: sheets,
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

function testRawFullRebuildDoesNotReferenceOutOfScopeFinalRecords() {
  const context = createContext();
  loadGasFiles(context, [
    'Constants.gs',
    'Utils.gs',
    'TrainingRecordService.gs',
    'RawSyncService.gs',
  ]);

  const APP_CONFIG = vm.runInContext('APP_CONFIG', context);
  let replacedWrite = null;

  context.getCurrentUserEmail_ = function () { return 'tester@example.com'; };
  context.nowIsoString_ = function () { return '2026-03-31T09:00:00'; };
  context.getEmployeeMasterRecords_ = function () { return []; };
  context.getAllEmployeeMasterRecords_ = function () { return []; };
  context.buildEmployeeLookup_ = function () { return {}; };
  context.getCourseMasterRecords_ = function () { return []; };
  context.buildCourseLookup_ = function () { return {}; };
  context.buildTrainingSessionLookupForRawSync_ = function () { return {}; };
  context.getTrainingRawParticipantRows_ = function () { return []; };
  context.readObjectsFromSheet_ = function (sheetKey) {
    if (sheetKey === 'TRAINING_RECORDS') {
      return [
        { __rowNumber: 2, record_id: 'MANUAL-001', source_type: 'MANUAL', row_status: APP_CONFIG.ROW_STATUS.ACTIVE },
        { __rowNumber: 3, record_id: 'RAW-001', source_type: APP_CONFIG.SOURCE_TYPES.RAW_PARTICIPANT, row_status: APP_CONFIG.ROW_STATUS.ACTIVE },
      ];
    }
    return [];
  };
  context.batchUpdateSheetRows_ = function () {};
  context.appendObjects_ = function () {
    throw new Error('Full rebuild should not append incrementally.');
  };
  context.writeObjectsReplacingBody_ = function (sheetKey, objects) {
    replacedWrite = { sheetKey, objects };
  };
  context.setConfigValue_ = function () {};

  const result = context.syncTrainingRecordsFromRawDataWithMode_({ mode: 'full' });

  assert.ok(replacedWrite, 'Full rebuild should rewrite TRAINING_RECORDS');
  assert.strictEqual(replacedWrite.sheetKey, 'TRAINING_RECORDS', 'Full rebuild should rewrite only TRAINING_RECORDS');
  assert.deepStrictEqual(
    Array.from(result.finalRecords.map(function (row) { return row.record_id; })),
    ['MANUAL-001'],
    'Full rebuild should return finalRecords without removed raw-derived rows'
  );
}

function testTrendOnlyRefreshWritesCustomTrendReportAndExposesMenuAction() {
  const context = createContext();
  loadGasFiles(context, [
    'Constants.gs',
    'Utils.gs',
    'ErrorHandler.gs',
    'TrainingSessionService.gs',
    'AnalyticsService.gs',
    'AnalyticsPresentationService.gs',
    'AnalyticsCustomReportService.gs',
    'AppController.gs',
    'MenuService.gs',
  ]);

  context.getActiveTrainingRecords_ = function () {
    return [
      {
        row_status: 'ACTIVE',
        emp_id: 'E001',
        attendance_status: 'Có mặt',
        duration_hours: 2,
        cost_per_pax: 100,
        training_month: '2026-03',
        training_date: '2026-03-12',
        session_id: 'SES-001',
        session_code: 'HOINHAP26-202603-01',
      },
      {
        row_status: 'ACTIVE',
        emp_id: 'E002',
        attendance_status: 'Có mặt',
        duration_hours: 3,
        cost_per_pax: 120,
        training_month: new Date('2026-01-01T00:00:00+07:00'),
        training_date: '2026-01-18',
        session_id: 'SES-002',
        session_code: 'GDPGSP26-202601-01',
      },
    ];
  };
  context.setConfigValue_ = function () {};
  context.runAction_ = function (actionName, executor) {
    return executor();
  };
  context.requireRole_ = function () {};
  context.ensureSheet_ = function (sheetKey) {
    if (!context.__sheets[sheetKey]) {
      context.__sheets[sheetKey] = createSheetStub(sheetKey);
    }
    return context.__sheets[sheetKey];
  };
  context.showAlert_ = function (title, message) {
    context.__alerts.push({ title, message });
  };

  assert.strictEqual(typeof context.refreshTrendAnalyticsOnly, 'function', 'AppController should expose a trend-only refresh action');

  const result = context.refreshTrendAnalyticsOnly();
  const trendSheet = context.ensureSheet_('ANALYTICS_TREND');

  assert.ok(
    trendSheet.rangeCalls.length > 0,
    'Trend-only refresh should write a custom report to ANALYTICS_TREND'
  );
  assert.ok(
    trendSheet.filterRange,
    'Trend-only refresh should create a filter for the detail block'
  );
  assert.strictEqual(result.data.total_records, 2, 'Trend-only refresh should report the number of active training records used');
  assert.ok(
    context.__alerts.some(function (entry) { return entry.title === 'Làm mới Phân tích xu hướng'; }),
    'Trend-only refresh should show a dedicated success alert'
  );

  context.buildMenu_();
  assert.ok(
    context.__menuItems.includes('Làm mới Phân tích xu hướng'),
    'Menu should expose a dedicated trend-only refresh action'
  );
}

try {
  testRawFullRebuildDoesNotReferenceOutOfScopeFinalRecords();
  testTrendOnlyRefreshWritesCustomTrendReportAndExposesMenuAction();
  console.log('PASS trend-only-refresh-tdd');
} catch (error) {
  console.error(error.stack || error.message || String(error));
  process.exitCode = 1;
}

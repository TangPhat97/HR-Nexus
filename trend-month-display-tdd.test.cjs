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
        case 'yyyy-MM-dd':
          return [value.getFullYear(), pad(value.getMonth() + 1), pad(value.getDate())].join('-');
        default:
          return [value.getFullYear(), pad(value.getMonth() + 1), pad(value.getDate())].join('-');
      }
    },
  };
}

function createContext() {
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
          alert() {},
          createMenu() {
            return {
              addItem() { return this; },
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
    'TrainingSessionService.gs',
    'AnalyticsService.gs',
    'AnalyticsPresentationService.gs',
    'AnalyticsCustomReportService.gs',
  ]);

  context.syncTrainingRecordsFromRawData_ = function () {
    return {
      failed_rows: 0,
      finalRecords: [
        {
          row_status: 'ACTIVE',
          emp_id: 'E001',
          course_id: 'HOINHAP26',
          course_name: 'Hoi nhap',
          department: 'Phong A',
          attendance_status: 'Có mặt',
          applied_on_job: 'Có',
          duration_hours: 2,
          cost_per_pax: 100,
          training_month: new Date('2026-03-01T00:00:00+07:00'),
          training_date: '2026-03-12',
          session_id: 'SES-001',
          session_code: 'HOINHAP26-202603-01',
        },
        {
          row_status: 'ACTIVE',
          emp_id: 'E002',
          course_id: 'GDPGSP26',
          course_name: 'GDP',
          department: 'Phong B',
          attendance_status: 'Có mặt',
          applied_on_job: 'Có',
          duration_hours: 3,
          cost_per_pax: 120,
          training_month: '2026-01',
          training_date: '2026-01-18',
          session_id: 'SES-002',
          session_code: 'GDPGSP26-202601-01',
        },
      ],
    };
  };

  context.readObjectsFromSheet_ = function (sheetKey) {
    if (sheetKey === 'QUEUE_JOBS' || sheetKey === 'QA_RESULTS') {
      return [];
    }
    if (sheetKey === 'CONFIG_SYSTEM') {
      return [];
    }
    if (sheetKey === 'MASTER_EMPLOYEES') {
      return [
        { emp_id: 'E001', department: 'Phong A' },
        { emp_id: 'E002', department: 'Phong B' },
      ];
    }
    return [];
  };

  context.writeDashboardRows_ = function () {};
  context.writeObjectsReplacingBody_ = function () {};
  context.batchUpsertObjects_ = function () {};
  context.setConfigValue_ = function () {};
  context.getQueueBacklogCount_ = function () { return 0; };
  context.getStagingSheetKeys_ = function () { return []; };
  context.getEmployeeMasterRecords_ = function () { return [{ emp_id: 'E001', department: 'Phong A' }, { emp_id: 'E002', department: 'Phong B' }]; };
  context.buildManagedSyncDashboardRows_ = function () { return []; };
  let capturedTrendRows = null;
  context.refreshDepartmentAnalyticsReportCore_ = function () { return { data: {} }; };
  context.refreshDepartmentCourseAnalyticsReportCore_ = function () { return { data: {} }; };
  context.refreshExternalAssignmentReportCore_ = function () { return { data: {} }; };
  context.refreshSessionReconciliationReportCore_ = function () { return { data: {} }; };
  context.refreshTrendAnalyticsReportCore_ = function (options) {
    capturedTrendRows = options.rows;
    return { data: {} };
  };

  context.refreshAnalyticsCore_({ mode: 'quick' });

  assert.ok(capturedTrendRows, 'refreshAnalyticsCore_ should rebuild trend analytics rows');
  const monthKeys = Array.from(capturedTrendRows.map((row) => row.month_key));
  assert.deepStrictEqual(
    monthKeys,
    ['2026-01', '2026-03'],
    'Trend month_key should always be normalized to YYYY-MM even when training_month is a Date value'
  );
}

runTests();

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
          return [value.getFullYear(), pad(value.getMonth() + 1), pad(value.getDate())].join('-');
        case 'dd/MM/yyyy':
          return [pad(value.getDate()), pad(value.getMonth() + 1), value.getFullYear()].join('/');
        case 'dd/MM/yyyy HH:mm:ss':
          return [
            pad(value.getDate()),
            pad(value.getMonth() + 1),
            value.getFullYear(),
          ].join('/') + ' ' + [pad(value.getHours()), pad(value.getMinutes()), pad(value.getSeconds())].join(':');
        default:
          return [value.getFullYear(), pad(value.getMonth() + 1), pad(value.getDate())].join('-');
      }
    },
  };
}

function createSheetStub(name) {
  const sheet = {
    name,
    tabColor: '',
    frozenRows: 0,
    columnWidths: [],
    rangeCalls: [],
    maxRows: 400,
    maxColumns: 40,
    filterRange: null,
    getName() { return this.name; },
    setName(nextName) { this.name = nextName; return this; },
    lastRow: 0,
    lastColumn: 0,
    getLastRow() { return this.lastRow; },
    getLastColumn() { return this.lastColumn; },
    getMaxRows() { return this.maxRows; },
    getMaxColumns() { return this.maxColumns; },
    clearContents() { this.clearedContents = true; return this; },
    clearFormats() { this.clearedFormats = true; return this; },
    autoResizeColumns() { return this; },
    setTabColor(color) { this.tabColor = color; return this; },
    setFrozenRows(count) { this.frozenRows = count; return this; },
    setColumnWidth(index, width) { this.columnWidths.push([index, width]); return this; },
    setColumnWidths(start, howMany, width) { this.columnWidths.push([start, howMany, width]); return this; },
    getFilter() {
      if (!this.filterRange) {
        return null;
      }
      const sheetRef = this;
      return {
        remove() { sheetRef.filterRange = null; },
      };
    },
    getRange(row, column, numRows, numColumns) {
      const call = { row, column, numRows, numColumns, values: null };
      this.rangeCalls.push(call);
      const sheetRef = this;
      const chain = {
        breakApart() { call.breakApart = true; return chain; },
        clearContent() { call.clearContent = true; return chain; },
        clearFormat() { call.clearFormat = true; return chain; },
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
        createFilter() {
          sheetRef.filterRange = { row, column, numRows, numColumns };
          call.filterCreated = true;
          return chain;
        },
      };
      return chain;
    },
  };
  return sheet;
}

function createContext() {
  const menuItems = [];
  const alerts = [];
  const writes = [];
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
    __writes: writes,
    __sheets: sheets,
  };
  context.global = context;
  return vm.createContext(context);
}

function ensureSheetStub(context, sheetKey) {
  if (!context.__sheets[sheetKey]) {
    context.__sheets[sheetKey] = createSheetStub(sheetKey);
  }
  return context.__sheets[sheetKey];
}

function loadGasFiles(context, fileNames) {
  fileNames.forEach((fileName) => {
    const filePath = path.join(process.cwd(), 'gas', fileName);
    const code = fs.readFileSync(filePath, 'utf8');
    vm.runInContext(code, context, { filename: filePath });
  });
}

function buildSampleTrainingRecords() {
  return [
    {
      row_status: 'ACTIVE',
      emp_id: 'E001',
      full_name: 'Nguyen Van A',
      department: 'Kho',
      division: 'Van hanh',
      job_title: 'Nhan vien',
      session_id: 'SES-001',
      session_code: 'EXT-202601-01',
      course_id: 'COURSE-EXT-01',
      course_name: 'Xuat nhap khau',
      training_date: '2026-01-25',
      training_month: '2026-01',
      delivery_type: 'Đào tạo bên ngoài',
      training_format: 'Cử nhân sự đi học',
      training_unit: 'Don vi A',
      location: 'HCM',
      attendance_status: 'Có mặt',
      plan_scope: 'Trong kế hoạch',
      program_type: 'Đào tạo tuân thủ',
      cost_per_pax: 1000000,
      estimated_cost: 1000000,
      duration_hours: 8,
      score: 88,
      satisfaction: 4,
      applied_on_job: 'Có',
      registered_count: 1,
      actual_count: 1,
    },
    {
      row_status: 'ACTIVE',
      emp_id: 'E002',
      full_name: 'Le Thi B',
      department: 'Kho',
      division: 'Van hanh',
      job_title: 'Nhan vien',
      session_id: 'SES-002',
      session_code: 'EXT-202603-01',
      course_id: 'COURSE-EXT-02',
      course_name: 'Lai xe van tai',
      training_date: '2026-03-21',
      training_month: '2026-03',
      delivery_type: 'Đào tạo bên ngoài',
      training_format: 'Cử nhân sự đi học',
      training_unit: 'Don vi B',
      location: 'Online',
      attendance_status: 'Có mặt',
      plan_scope: 'Ngoài kế hoạch',
      program_type: 'Đào tạo tuân thủ',
      cost_per_pax: 500000,
      estimated_cost: 5000000,
      duration_hours: 4,
      score: 76,
      satisfaction: 4,
      applied_on_job: 'Có',
      registered_count: 10,
      actual_count: 10,
    },
    {
      row_status: 'ACTIVE',
      emp_id: 'E003',
      full_name: 'Tran Van C',
      department: 'QA',
      division: 'Chat luong',
      job_title: 'QA',
      session_id: 'SES-003',
      session_code: 'IN-202603-01',
      course_id: 'COURSE-IN-01',
      course_name: 'Hoi nhap',
      training_date: '2026-03-12',
      training_month: '2026-03',
      delivery_type: 'Đào tạo nội bộ',
      training_format: 'Offline',
      training_unit: 'Noi bo',
      location: 'Hiep Phuoc',
      attendance_status: 'Có mặt',
      plan_scope: 'Trong kế hoạch',
      program_type: 'Đào tạo hệ thống',
      cost_per_pax: 0,
      estimated_cost: 0,
      duration_hours: 2,
      score: 90,
      satisfaction: 5,
      applied_on_job: 'Có',
      registered_count: 15,
      actual_count: 10,
    },
    {
      row_status: 'ACTIVE',
      emp_id: 'E004',
      full_name: 'Pham Thi D',
      department: 'QA',
      division: 'Chat luong',
      job_title: 'QA',
      session_id: 'SES-003',
      session_code: 'IN-202603-01',
      course_id: 'COURSE-IN-01',
      course_name: 'Hoi nhap',
      training_date: '2026-03-12',
      training_month: '2026-03',
      delivery_type: 'Đào tạo nội bộ',
      training_format: 'Offline',
      training_unit: 'Noi bo',
      location: 'Hiep Phuoc',
      attendance_status: 'Vắng có phép',
      plan_scope: 'Trong kế hoạch',
      program_type: 'Đào tạo hệ thống',
      cost_per_pax: 0,
      estimated_cost: 0,
      duration_hours: 2,
      score: '',
      satisfaction: '',
      applied_on_job: 'Không',
      registered_count: 15,
      actual_count: 10,
    },
  ];
}

function createBaseContextWithAnalytics() {
  const context = createContext();
  loadGasFiles(context, [
    'Constants.gs',
    'Utils.gs',
    'ErrorHandler.gs',
    'TrainingSessionService.gs',
    'AnalyticsService.gs',
    'AnalyticsPresentationService.gs',
    'AnalyticsCustomReportService.gs',
    'ExternalAssignmentReportService.gs',
    'AppController.gs',
    'MenuService.gs',
  ]);
  context.getConfigMap_ = function () {
    return { CURRENT_FISCAL_YEAR: '2026' };
  };
  context.nowIsoString_ = function () { return '2026-04-01T10:00:00'; };
  context.ensureSheet_ = function (sheetKey) { return ensureSheetStub(context, sheetKey); };
  context.requireRole_ = function () {};
  context.runAction_ = function (actionName, executor) { return executor(); };
  context.showAlert_ = function (title, message) {
    context.__alerts.push({ title, message });
  };
  context.getDepartmentHeadcountMap_ = function () {
    return { Kho: 10, QA: 5 };
  };
  return context;
}

function testDepartmentAndCourseAndTrendReportsAreCustomReports() {
  const context = createBaseContextWithAnalytics();
  const APP_CONFIG = vm.runInContext('APP_CONFIG', context);

  assert.ok(
    Array.from(APP_CONFIG.CUSTOM_MANAGED_REPORT_SHEETS || []).includes('ANALYTICS_DEPARTMENT'),
    'Department analytics should be treated as a custom managed report'
  );
  assert.ok(
    Array.from(APP_CONFIG.CUSTOM_MANAGED_REPORT_SHEETS || []).includes('ANALYTICS_DEPARTMENT_COURSE'),
    'Department-course analytics should be treated as a custom managed report'
  );
  assert.ok(
    Array.from(APP_CONFIG.CUSTOM_MANAGED_REPORT_SHEETS || []).includes('ANALYTICS_TREND'),
    'Trend analytics should be treated as a custom managed report'
  );
  assert.deepStrictEqual(
    Array.from(context.getSheetDisplayHeaders_('ANALYTICS_DEPARTMENT')),
    [],
    'Custom report sheets should not expose flat display headers to QA/schema checks'
  );
}

function testBuildDepartmentAnalyticsReportDataAddsKpisRankingsAndDerivedColumns() {
  const context = createBaseContextWithAnalytics();
  const records = buildSampleTrainingRecords();

  assert.strictEqual(typeof context.buildDepartmentAnalyticsRows_, 'function', 'Department row builder should exist');
  assert.strictEqual(typeof context.buildDepartmentAnalyticsReportData_, 'function', 'Department report data builder should exist');

  const rows = context.buildDepartmentAnalyticsRows_(records, { departmentHeadcount: context.getDepartmentHeadcountMap_() });
  const report = context.buildDepartmentAnalyticsReportData_(rows, { fiscalYear: '2026', lastRefreshed: '2026-04-01T10:00:00' });

  assert.strictEqual(report.kpis.department_count, 2, 'Department KPI should count departments with training');
  assert.strictEqual(report.kpis.total_employees, 15, 'Department KPI should sum department headcount');
  assert.strictEqual(report.kpis.trained_employees, 4, 'Department KPI should sum unique trained employees across departments');
  assert.strictEqual(report.rankings.topCoverage.length > 0, true, 'Department report should expose top coverage block');
  assert.strictEqual(report.rankings.watchlist.length > 0, true, 'Department report should expose watchlist block');
  assert.ok(
    Object.prototype.hasOwnProperty.call(report.detailRows[0], 'hours_per_employee'),
    'Department detail rows should add hours_per_employee'
  );
  assert.ok(
    Object.prototype.hasOwnProperty.call(report.detailRows[0], 'cost_per_employee'),
    'Department detail rows should add cost_per_employee'
  );
  assert.ok(
    Object.prototype.hasOwnProperty.call(report.detailRows[0], 'coverage_rank'),
    'Department detail rows should add coverage_rank'
  );
}

function testBuildDepartmentCourseAnalyticsReportDataAddsExtraAnalysis() {
  const context = createBaseContextWithAnalytics();
  const records = buildSampleTrainingRecords();

  assert.strictEqual(typeof context.buildDepartmentCourseAnalyticsRows_, 'function', 'Department-course row builder should exist');
  assert.strictEqual(typeof context.buildDepartmentCourseAnalyticsReportData_, 'function', 'Department-course report data builder should exist');

  const rows = context.buildDepartmentCourseAnalyticsRows_(records, { departmentHeadcount: context.getDepartmentHeadcountMap_() });
  const report = context.buildDepartmentCourseAnalyticsReportData_(rows, { fiscalYear: '2026', lastRefreshed: '2026-04-01T10:00:00' });

  assert.strictEqual(report.kpis.department_course_count, rows.length, 'Department-course KPI should count unique department-course pairs');
  assert.ok(report.rankings.topParticipants.length > 0, 'Department-course report should expose top participants block');
  assert.ok(report.rankings.topCost.length > 0, 'Department-course report should expose top cost block');
  assert.ok(
    Object.prototype.hasOwnProperty.call(report.detailRows[0], 'participation_rate'),
    'Department-course detail rows should add participation_rate'
  );
  assert.ok(
    Object.prototype.hasOwnProperty.call(report.detailRows[0], 'hours_per_employee'),
    'Department-course detail rows should add hours_per_employee'
  );
  assert.ok(
    Object.prototype.hasOwnProperty.call(report.detailRows[0], 'participant_rank'),
    'Department-course detail rows should add participant_rank'
  );
}

function testBuildTrendAnalyticsReportDataAddsComparisonsAndHighlights() {
  const context = createBaseContextWithAnalytics();
  const records = buildSampleTrainingRecords();

  assert.strictEqual(typeof context.buildTrendAnalyticsReportData_, 'function', 'Trend report data builder should exist');

  const rows = context.buildTrendAnalyticsRows_(records);
  const report = context.buildTrendAnalyticsReportData_(rows, { fiscalYear: '2026', lastRefreshed: '2026-04-01T10:00:00' });

  assert.deepStrictEqual(
    Array.from(report.detailRows.map(function (row) { return row.month_key; })),
    ['2026-01', '2026-03'],
    'Trend report should keep YYYY-MM month keys'
  );
  assert.strictEqual(report.kpis.month_count, 2, 'Trend KPI should count populated months');
  assert.ok(report.comparison.latest_month, 'Trend report should expose latest month comparison');
  assert.ok(report.highlights.highest_participants, 'Trend report should expose highest participants highlight');
  assert.ok(
    Object.prototype.hasOwnProperty.call(report.detailRows[1], 'participant_delta_vs_previous'),
    'Trend detail rows should include participant delta vs previous month'
  );
  assert.ok(
    Object.prototype.hasOwnProperty.call(report.detailRows[1], 'cost_delta_vs_previous'),
    'Trend detail rows should include cost delta vs previous month'
  );
}

function testBuildSessionReconciliationReportDataComputesStatusesAndSuggestions() {
  const context = createBaseContextWithAnalytics();

  assert.strictEqual(typeof context.buildSessionReconciliationReportData_, 'function', 'Session reconciliation report data builder should exist');

  const trainingSessions = [
    {
      row_status: 'ACTIVE',
      session_id: 'SES-A',
      session_code: 'LOP-A',
      course_id: 'COURSE-A',
      course_name: 'Khoa A',
      training_date: '2026-03-20',
      training_month: '2026-03',
      training_unit: 'Don vi A',
      delivery_type: 'Đào tạo bên ngoài',
      training_format: 'Cử nhân sự đi học',
      registered_count: 3,
      actual_count: 2,
      estimated_cost: 3000000,
      total_hours: 8,
    },
    {
      row_status: 'ACTIVE',
      session_id: 'SES-B',
      session_code: 'LOP-B',
      course_id: 'COURSE-B',
      course_name: 'Khoa B',
      training_date: '2026-03-21',
      training_month: '2026-03',
      training_unit: 'Don vi B',
      delivery_type: 'Đào tạo nội bộ',
      training_format: 'Offline',
      registered_count: '',
      actual_count: '',
      estimated_cost: '',
      total_hours: 2,
    },
    {
      row_status: 'ACTIVE',
      session_id: 'SES-C',
      session_code: 'LOP-C',
      course_id: 'COURSE-C',
      course_name: 'Khoa C',
      training_date: '2026-03-22',
      training_month: '2026-03',
      training_unit: 'Don vi C',
      delivery_type: 'Đào tạo bên ngoài',
      training_format: 'Cử nhân sự đi học',
      registered_count: 1,
      actual_count: 1,
      estimated_cost: 1000000,
      total_hours: 4,
    },
    {
      row_status: 'ACTIVE',
      session_id: 'SES-D',
      session_code: 'LOP-D',
      course_id: 'COURSE-D',
      course_name: 'Khoa D',
      training_date: '2026-03-23',
      training_month: '2026-03',
      training_unit: 'Don vi D',
      delivery_type: 'Đào tạo bên ngoài',
      training_format: 'Cử nhân sự đi học',
      registered_count: 1,
      actual_count: 1,
      estimated_cost: 2000000,
      total_hours: 6,
    },
  ];
  const rawRows = [
    { row_status: 'ACTIVE', session_id: 'SES-A', session_code: 'LOP-A', emp_id: 'E001', attendance_status: 'Có mặt' },
    { row_status: 'ACTIVE', session_id: 'SES-A', session_code: 'LOP-A', emp_id: 'E002', attendance_status: 'Vắng có phép' },
    { row_status: 'ACTIVE', session_id: 'SES-B', session_code: 'LOP-B', email: 'guest@example.com', attendance_status: 'Có mặt' },
    { row_status: 'ACTIVE', session_id: 'SES-D', session_code: 'LOP-D', emp_id: 'E004', attendance_status: 'Có mặt' },
  ];
  const factRows = [
    { row_status: 'ACTIVE', session_id: 'SES-A', session_code: 'LOP-A', emp_id: 'E001', attendance_status: 'Có mặt' },
    { row_status: 'ACTIVE', session_id: 'SES-C', session_code: 'LOP-C', emp_id: 'E003', attendance_status: 'Có mặt' },
    { row_status: 'ACTIVE', session_id: 'SES-D', session_code: 'LOP-D', emp_id: 'E004', attendance_status: 'Có mặt' },
  ];

  const report = context.buildSessionReconciliationReportData_({
    fiscalYear: '2026',
    lastRefreshed: '2026-04-01T10:00:00',
    trainingSessions: trainingSessions,
    rawParticipants: rawRows,
    trainingRecords: factRows,
  });

  assert.strictEqual(report.kpis.total_sessions, 4, 'Reconciliation KPI should count classes in scope');
  assert.strictEqual(report.kpis.sessions_with_raw, 3, 'Reconciliation KPI should count classes that already have raw rows');
  assert.strictEqual(report.kpis.sessions_without_raw, 1, 'Reconciliation KPI should count classes without raw rows');
  assert.strictEqual(report.kpis.matched_sessions, 1, 'Reconciliation KPI should count fully matched classes');
  assert.strictEqual(report.kpis.mismatched_sessions, 3, 'Reconciliation KPI should count mismatched classes');

  const statusByCode = report.detailRows.reduce(function (accumulator, row) {
    accumulator[row.session_code] = row;
    return accumulator;
  }, {});

  assert.strictEqual(statusByCode['LOP-A'].reconciliation_status, 'Lệch số liệu', 'Session A should be flagged as mismatched');
  assert.strictEqual(statusByCode['LOP-A'].suggested_action, 'Kiểm tra số lượng HV đăng ký ở Lớp đào tạo', 'Session A should suggest checking manual registered headcount first');
  assert.strictEqual(statusByCode['LOP-B'].reconciliation_status, 'Đã có raw, chưa đồng bộ', 'Session B should be flagged as raw but not synced');
  assert.strictEqual(statusByCode['LOP-C'].reconciliation_status, 'Chưa có data raw', 'Session C should be flagged as missing raw data');
  assert.strictEqual(statusByCode['LOP-D'].reconciliation_status, 'Khớp', 'Session D should be flagged as matched');
}

function testRefreshSessionReconciliationActionAndMenuWriteOnlyTargetSheet() {
  const context = createBaseContextWithAnalytics();

  const trainingSessions = [
    {
      row_status: 'ACTIVE',
      session_id: 'SES-A',
      session_code: 'LOP-A',
      course_id: 'COURSE-A',
      course_name: 'Khoa A',
      training_date: '2026-03-20',
      training_month: '2026-03',
      training_unit: 'Don vi A',
      training_format: 'Offline',
      delivery_type: 'Đào tạo nội bộ',
      registered_count: 2,
      actual_count: 2,
      estimated_cost: 500000,
      total_hours: 2,
    },
  ];
  context.getTrainingSessions_ = function () { return trainingSessions; };
  context.getTrainingRawParticipantRows_ = function () { return []; };
  context.getActiveTrainingRecords_ = function () { return []; };

  assert.strictEqual(typeof context.refreshSessionReconciliationReport, 'function', 'AppController should expose a session reconciliation refresh action');
  const result = context.refreshSessionReconciliationReport();

  const reconciliationSheet = ensureSheetStub(context, 'REPORT_SESSION_RECONCILIATION');
  assert.ok(reconciliationSheet.rangeCalls.length > 0, 'Refreshing reconciliation should write to the custom report sheet');
  assert.strictEqual(ensureSheetStub(context, 'ANALYTICS_TREND').rangeCalls.length, 0, 'Refreshing reconciliation should not rewrite trend analytics');
  assert.strictEqual(result.data.total_sessions, 1, 'Refresh result should return reconciliation KPI summary');

  context.buildMenu_();
  assert.ok(
    context.__menuItems.includes('Làm mới đối chiếu lớp đào tạo'),
    'Menu should expose a dedicated session reconciliation refresh action'
  );
}

function testRefreshAnalyticsCoreSharesLoadedDataAcrossReports() {
  const context = createBaseContextWithAnalytics();
  const records = buildSampleTrainingRecords();
  const trainingSessions = [
    {
      row_status: 'ACTIVE',
      session_id: 'SES-001',
      session_code: 'EXT-202601-01',
      course_id: 'COURSE-EXT-01',
      course_name: 'Xuat nhap khau',
      training_date: '2026-01-25',
      training_month: '2026-01',
      training_unit: 'Don vi A',
      training_format: 'Cu nhan su di hoc',
      delivery_type: 'Dao tao ben ngoai',
      registered_count: 1,
      actual_count: 1,
      estimated_cost: 1000000,
      total_hours: 8,
      attendance_rate: 100,
    },
    {
      row_status: 'ACTIVE',
      session_id: 'SES-002',
      session_code: 'EXT-202603-01',
      course_id: 'COURSE-EXT-02',
      course_name: 'Lai xe van tai',
      training_date: '2026-03-21',
      training_month: '2026-03',
      training_unit: 'Don vi B',
      training_format: 'Cu nhan su di hoc',
      delivery_type: 'Dao tao ben ngoai',
      registered_count: 2,
      actual_count: 2,
      estimated_cost: 5000000,
      total_hours: 4,
      attendance_rate: 100,
    },
  ];
  const rawParticipants = records.map(function (row) {
    return {
      row_status: 'ACTIVE',
      session_id: row.session_id,
      session_code: row.session_code,
      course_id: row.course_id,
      course_name: row.course_name,
      training_date: row.training_date,
      training_month: row.training_month,
      emp_id: row.emp_id,
      full_name: row.full_name,
      email: row.emp_id + '@example.com',
      attendance_status: row.attendance_status,
    };
  });
  const counters = {
    configReads: 0,
    trainingSessionsReads: 0,
    rawReads: 0,
    recordReads: 0,
  };

  context.getConfigMap_ = function () {
    counters.configReads += 1;
    return { CURRENT_FISCAL_YEAR: '2026' };
  };
  context.getTrainingSessions_ = function () {
    counters.trainingSessionsReads += 1;
    return trainingSessions;
  };
  context.getTrainingRawParticipantRows_ = function () {
    counters.rawReads += 1;
    return rawParticipants;
  };
  context.getActiveTrainingRecords_ = function () {
    counters.recordReads += 1;
    return records;
  };
  context.writeDashboardRows_ = function () {};
  context.writeObjectsReplacingBody_ = function () {};
  context.readObjectsFromSheet_ = function () { return []; };
  context.getStagingSheetKeys_ = function () { return []; };
  context.getQueueBacklogCount_ = function () { return 0; };
  context.buildManagedSyncDashboardRows_ = function () { return []; };
  context.setConfigValue_ = function () {};

  context.refreshAnalyticsCore_({
    fiscalYear: '2026',
    skipRawSync: true,
    rawSyncSummary: {
      finalRecords: records,
      failed_rows: 0,
      updated_records: 1,
      inserted_records: 0,
      removed_records: 0,
      mode: 'quick',
    },
  });

  assert.strictEqual(counters.configReads, 0, 'Analytics refresh should not re-read config when fiscal year is already provided');
  assert.strictEqual(counters.trainingSessionsReads, 1, 'Analytics refresh should load training sessions once and share them across reports');
  assert.strictEqual(counters.rawReads, 1, 'Analytics refresh should load raw participants once and share them across reports');
  assert.strictEqual(counters.recordReads, 0, 'Analytics refresh should reuse provided final records instead of re-reading fact records');
}

function testSessionReconciliationResolvesRowsOncePerDatasetRow() {
  const context = createBaseContextWithAnalytics();
  const trainingSessions = [
    {
      row_status: 'ACTIVE',
      session_id: 'SES-A',
      session_code: 'LOP-A',
      course_id: 'COURSE-A',
      course_name: 'Khoa A',
      training_date: '2026-03-20',
      training_month: '2026-03',
      training_unit: 'Don vi A',
      training_format: 'Offline',
      delivery_type: 'Dao tao noi bo',
      registered_count: 2,
      actual_count: 2,
      estimated_cost: 500000,
      total_hours: 2,
    },
  ];
  const rawRows = [
    { row_status: 'ACTIVE', session_id: 'SES-A', emp_id: 'E001', full_name: 'A', email: 'a@example.com', attendance_status: 'Co mat' },
    { row_status: 'ACTIVE', session_code: 'LOP-A', emp_id: 'E002', full_name: 'B', email: 'b@example.com', attendance_status: 'Co mat' },
  ];
  const factRows = [
    { row_status: 'ACTIVE', session_id: 'SES-A', emp_id: 'E001', full_name: 'A', email: 'a@example.com', attendance_status: 'Co mat' },
    { row_status: 'ACTIVE', session_code: 'LOP-A', emp_id: 'E002', full_name: 'B', email: 'b@example.com', attendance_status: 'Co mat' },
  ];
  const originalResolve = context.resolveTrainingSession_;
  let resolveCount = 0;
  context.resolveTrainingSession_ = function () {
    resolveCount += 1;
    return originalResolve.apply(this, arguments);
  };

  context.buildSessionReconciliationReportData_({
    fiscalYear: '2026',
    trainingSessions: trainingSessions,
    rawParticipants: rawRows,
    trainingRecords: factRows,
  });

  assert.strictEqual(resolveCount, rawRows.length + factRows.length, 'Reconciliation should resolve each raw/fact row once before grouping');
}

function testPrepareAnalyticsReportSheetResetsOnlyUsedRange() {
  const context = createBaseContextWithAnalytics();
  const sheet = ensureSheetStub(context, 'ANALYTICS_TREND');
  sheet.lastRow = 24;
  sheet.lastColumn = 11;

  context.prepareAnalyticsReportSheet_('ANALYTICS_TREND', {
    frozenRows: 2,
    expectedWidth: 17,
  });

  assert.strictEqual(sheet.clearedContents, undefined, 'Optimized report writer should not clear the whole sheet');
  assert.strictEqual(sheet.clearedFormats, undefined, 'Optimized report writer should not clear formats for the whole sheet');
  assert.ok(sheet.rangeCalls.length > 0, 'Preparing report sheet should still touch the used range');
  assert.strictEqual(sheet.rangeCalls[0].numRows, 24, 'Preparing report sheet should only reset rows that were previously used');
  assert.strictEqual(sheet.rangeCalls[0].numColumns, 17, 'Preparing report sheet should only reset the used width needed for the report');
}

function run() {
  testDepartmentAndCourseAndTrendReportsAreCustomReports();
  testBuildDepartmentAnalyticsReportDataAddsKpisRankingsAndDerivedColumns();
  testBuildDepartmentCourseAnalyticsReportDataAddsExtraAnalysis();
  testBuildTrendAnalyticsReportDataAddsComparisonsAndHighlights();
  testBuildSessionReconciliationReportDataComputesStatusesAndSuggestions();
  testRefreshSessionReconciliationActionAndMenuWriteOnlyTargetSheet();
  testRefreshAnalyticsCoreSharesLoadedDataAcrossReports();
  testSessionReconciliationResolvesRowsOncePerDatasetRow();
  testPrepareAnalyticsReportSheetResetsOnlyUsedRange();
  console.log('PASS analytics-custom-report-tdd');
}

run();

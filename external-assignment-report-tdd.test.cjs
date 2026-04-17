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
    tabColor: '',
    frozenRows: 0,
    columnWidths: [],
    rangeCalls: [],
    maxRows: 200,
    maxColumns: 26,
    filterRange: null,
    getName() { return name; },
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
  const sheetStub = createSheetStub('Cử đi học bên ngoài');

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
    __sheetStub: sheetStub,
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

function buildSampleRecords() {
  return [
    {
      row_status: 'ACTIVE',
      emp_id: 'E001',
      full_name: 'Nguyen Van A',
      department: 'Kho',
      division: 'Supply',
      job_title: 'Staff',
      session_id: 'SES-001',
      session_code: 'XKN26_01-202601-01',
      course_id: 'XKN26_01',
      course_name: 'Huong dan XKN',
      training_date: '2026-01-25',
      training_month: '2026-01',
      delivery_type: 'Đào tạo bên ngoài',
      training_format: 'Cử nhân sự đi học',
      training_unit: 'Lien doan thuong mai',
      location: 'TP. HCM',
      attendance_status: 'Có mặt',
      plan_scope: 'Trong kế hoạch',
      program_type: 'Đào tạo tuân thủ',
      estimated_cost: 1000,
      avg_cost_per_pax: '',
      cost_per_pax: 400,
    },
    {
      row_status: 'ACTIVE',
      emp_id: 'E002',
      full_name: 'Tran Thi B',
      department: 'QA',
      division: 'Quality',
      job_title: 'Senior',
      session_id: 'SES-001',
      session_code: 'XKN26_01-202601-01',
      course_id: 'XKN26_01',
      course_name: 'Huong dan XKN',
      training_date: '2026-01-25',
      training_month: '2026-01',
      delivery_type: 'Đào tạo bên ngoài',
      training_format: 'Cử nhân sự đi học',
      training_unit: 'Lien doan thuong mai',
      location: 'TP. HCM',
      attendance_status: 'Vắng có phép',
      plan_scope: 'Trong kế hoạch',
      program_type: 'Đào tạo tuân thủ',
      estimated_cost: 1000,
      avg_cost_per_pax: '',
      cost_per_pax: 400,
    },
    {
      row_status: 'ACTIVE',
      emp_id: 'E003',
      full_name: 'Le Van C',
      department: 'Kho',
      division: 'Supply',
      job_title: 'Lead',
      session_id: 'SES-002',
      session_code: 'XKN26_02-202603-01',
      course_id: 'XKN26_02',
      course_name: 'Thu tuc hai quan',
      training_date: '2026-03-21',
      training_month: new Date('2026-03-01T00:00:00+07:00'),
      delivery_type: 'Đào tạo bên ngoài',
      training_format: 'Cử nhân sự đi học',
      training_unit: 'Truong boi duong',
      location: 'Online',
      attendance_status: 'Có mặt',
      plan_scope: 'Ngoài kế hoạch',
      program_type: 'Đào tạo tuân thủ',
      estimated_cost: '',
      avg_cost_per_pax: 350,
      cost_per_pax: '',
    },
    {
      row_status: 'ACTIVE',
      emp_id: 'E004',
      full_name: 'Pham Thi D',
      department: 'Logistics',
      division: 'Supply',
      job_title: 'Coordinator',
      session_id: 'SES-003',
      session_code: 'GIAONHAN26_01-202604-01',
      course_id: 'GIAONHAN26_01',
      course_name: 'Van tai',
      training_date: '2026-04-21',
      training_month: '2026-04',
      delivery_type: 'Đào tạo bên ngoài',
      training_format: 'Cử nhân sự đi học',
      training_unit: 'An toan lao dong',
      location: 'Online',
      attendance_status: 'Có mặt',
      plan_scope: 'Ngoài kế hoạch',
      program_type: 'Đào tạo tuân thủ',
      estimated_cost: '',
      avg_cost_per_pax: '',
      cost_per_pax: 500,
    },
    {
      row_status: 'ACTIVE',
      emp_id: 'E005',
      full_name: 'Vu Thi E',
      department: 'QA',
      division: 'Quality',
      job_title: 'Staff',
      session_id: 'SES-004',
      session_code: 'HC26_01-202605-01',
      course_id: 'HC26_01',
      course_name: 'PCCC',
      training_date: '2026-05-28',
      training_month: '2026-05',
      delivery_type: 'Đào tạo bên ngoài',
      training_format: 'Cử nhân sự đi học',
      training_unit: 'P. Hanh Chinh',
      location: 'Online',
      attendance_status: 'Có mặt',
      plan_scope: 'Trong kế hoạch',
      program_type: 'Đào tạo quy trình',
      estimated_cost: '',
      avg_cost_per_pax: '',
      cost_per_pax: '',
    },
    {
      row_status: 'ACTIVE',
      emp_id: 'E006',
      full_name: 'Ngo Thi F',
      department: 'QA',
      division: 'Quality',
      job_title: 'Staff',
      session_id: 'SES-005',
      session_code: 'HOINHAP26-202601-01',
      course_id: 'HOINHAP26',
      course_name: 'Hoi nhap',
      training_date: '2026-01-13',
      training_month: '2026-01',
      delivery_type: 'Đào tạo nội bộ',
      training_format: 'Offline',
      training_unit: 'Noi bo',
      location: 'Hiep Phuoc',
      attendance_status: 'Có mặt',
      plan_scope: 'Trong kế hoạch',
      program_type: 'Đào tạo hệ thống',
      estimated_cost: 200,
      avg_cost_per_pax: '',
      cost_per_pax: 100,
    },
    {
      row_status: 'ACTIVE',
      emp_id: 'E007',
      full_name: 'Hoang Thi G',
      department: 'QA',
      division: 'Quality',
      job_title: 'Staff',
      session_id: 'SES-006',
      session_code: 'XKN25_01-202512-01',
      course_id: 'XKN25_01',
      course_name: 'Cu nam',
      training_date: '2025-12-18',
      training_month: '2025-12',
      delivery_type: 'Đào tạo bên ngoài',
      training_format: 'Cử nhân sự đi học',
      training_unit: 'External',
      location: 'TP. HCM',
      attendance_status: 'Có mặt',
      plan_scope: 'Trong kế hoạch',
      program_type: 'Đào tạo tuân thủ',
      estimated_cost: 999,
      avg_cost_per_pax: '',
      cost_per_pax: '',
    },
  ];
}

function testReportBuilderFiltersCurrentYearAndComputesCostFallback() {
  const context = createContext();
  loadGasFiles(context, [
    'Constants.gs',
    'Utils.gs',
    'TrainingSessionService.gs',
    'ExternalAssignmentReportService.gs',
  ]);

  context.nowIsoString_ = function () { return '2026-03-31T14:45:00'; };
  context.getConfigMap_ = function () { return { CURRENT_FISCAL_YEAR: '2026' }; };

  const report = context.buildExternalAssignmentReportData_(buildSampleRecords(), {});

  assert.strictEqual(report.fiscalYear, '2026', 'Report should respect CURRENT_FISCAL_YEAR');
  assert.strictEqual(report.kpis.session_count, 4, 'Report should count only matching current-year external assignment sessions');
  assert.strictEqual(report.kpis.participant_count, 5, 'Report should count all matching participant rows');
  assert.strictEqual(report.kpis.unique_employee_count, 5, 'Report should count unique matching employees');
  assert.strictEqual(report.kpis.total_cost, 1850, 'Report should sum session totals using the configured fallback chain');

  const sessionRows = Array.from(report.sessionRows);
  assert.deepStrictEqual(
    sessionRows.map(function (row) { return row.cost_basis; }),
    ['ESTIMATED_COST', 'AVG_COST_PER_PAX', 'COST_PER_PAX', 'MISSING'],
    'Session summary should expose the configured cost fallback basis'
  );
  assert.deepStrictEqual(
    sessionRows.map(function (row) { return row.session_total_cost; }),
    [1000, 350, 500, ''],
    'Session summary should compute total cost from estimated, avg per pax, cost per pax, or blank'
  );

  const detailRows = Array.from(report.detailRows);
  assert.strictEqual(detailRows.length, 5, 'Detail rows should include all matching participants in the current year');
  assert.deepStrictEqual(
    detailRows.slice(0, 2).map(function (row) { return row.participant_cost; }),
    [500, 500],
    'Estimated session cost should be allocated evenly across participants in the session'
  );
  assert.ok(
    detailRows.every(function (row) { return row.year === '2026'; }),
    'Detail rows should carry the normalized report year'
  );
  assert.ok(
    detailRows.every(function (row) { return row.delivery_type === 'Đào tạo bên ngoài' && row.training_format === 'Cử nhân sự đi học'; }),
    'Detail rows should be hard-filtered to external assignment training records'
  );

  const departmentRows = Array.from(report.departmentRows);
  const khoRow = departmentRows.find(function (row) { return row.department === 'Kho'; });
  assert.ok(khoRow, 'Department summary should include participating departments');
  assert.strictEqual(khoRow.participant_count, 2, 'Department summary should aggregate participant counts');
  assert.strictEqual(khoRow.total_allocated_cost, 850, 'Department summary should aggregate participant-level allocated cost');
}

function testManualRefreshOnlyWritesDedicatedSheetAndMenuIsExposed() {
  const context = createContext();
  loadGasFiles(context, [
    'Constants.gs',
    'Utils.gs',
    'ErrorHandler.gs',
    'TrainingSessionService.gs',
    'ExternalAssignmentReportService.gs',
    'AnalyticsService.gs',
    'AnalyticsPresentationService.gs',
    'AnalyticsCustomReportService.gs',
    'AppController.gs',
    'MenuService.gs',
  ]);

  context.getConfigMap_ = function () { return { CURRENT_FISCAL_YEAR: '2026' }; };
  context.getActiveTrainingRecords_ = function () { return buildSampleRecords(); };
  context.ensureSheet_ = function (sheetKey) {
    if (sheetKey !== 'REPORT_EXTERNAL_ASSIGNMENT') {
      throw new Error('Manual report refresh should only ensure the dedicated report sheet.');
    }
    return context.__sheetStub;
  };
  context.runAction_ = function (actionName, executor) { return executor(); };
  context.requireRole_ = function () {};
  context.showAlert_ = function (title, message) {
    context.__alerts.push({ title, message });
  };
  context.writeObjectsReplacingBody_ = function (sheetKey) {
    context.__writes.push(sheetKey);
  };
  context.nowIsoString_ = function () { return '2026-03-31T14:45:00'; };

  assert.strictEqual(typeof context.refreshExternalAssignmentReport, 'function', 'AppController should expose a public refresh action for the report');

  const result = context.refreshExternalAssignmentReport();

  assert.strictEqual(result.data.participant_count, 5, 'Manual refresh should return the matching participant count');
  assert.strictEqual(result.data.session_count, 4, 'Manual refresh should return the matching session count');
  assert.deepStrictEqual(
    Array.from(context.__writes),
    [],
    'Manual refresh should not touch table-based analytics writers'
  );
  assert.ok(
    context.__sheetStub.rangeCalls.some(function (call) {
      return Array.isArray(call.values) && call.values.some(function (row) {
        return Array.isArray(row) && row.indexOf('Mã lớp') !== -1 && row.indexOf('Chi phí/người') !== -1;
      });
    }),
    'Custom writer should render the detail header block onto the dedicated sheet'
  );
  assert.ok(
    context.__alerts.some(function (entry) { return /cử đi học bên ngoài/i.test(entry.title); }),
    'Manual refresh should show a dedicated success alert'
  );

  context.buildMenu_();
  assert.ok(
    context.__menuItems.includes('Làm mới báo cáo cử đi học bên ngoài'),
    'Menu should expose a dedicated report refresh action'
  );
}

function testAnalyticsRefreshAlsoRebuildsExternalAssignmentReport() {
  const context = createContext();
  loadGasFiles(context, [
    'Constants.gs',
    'Utils.gs',
    'ErrorHandler.gs',
    'TrainingSessionService.gs',
    'ExternalAssignmentReportService.gs',
    'AnalyticsService.gs',
    'AnalyticsPresentationService.gs',
    'AnalyticsCustomReportService.gs',
  ]);

  let reportCall = null;
  context.readObjectsFromSheet_ = function (sheetKey) {
    if (sheetKey === 'QUEUE_JOBS' || sheetKey === 'QA_RESULTS' || sheetKey === 'CONFIG_SYSTEM') {
      return [];
    }
    if (sheetKey === 'MASTER_EMPLOYEES') {
      return [];
    }
    return [];
  };
  context.syncTrainingRecordsFromRawData_ = function () {
    return {
      failed_rows: 0,
      finalRecords: buildSampleRecords(),
    };
  };
  context.writeObjectsReplacingBody_ = function () {};
  context.writeDashboardRows_ = function () {};
  context.batchUpsertObjects_ = function () {};
  context.setConfigValue_ = function () {};
  context.getQueueBacklogCount_ = function () { return 0; };
  context.getStagingSheetKeys_ = function () { return []; };
  context.getEmployeeMasterRecords_ = function () { return []; };
  context.buildManagedSyncDashboardRows_ = function () { return []; };
  context.refreshDepartmentAnalyticsReportCore_ = function () { return { data: {} }; };
  context.refreshDepartmentCourseAnalyticsReportCore_ = function () { return { data: {} }; };
  context.refreshTrendAnalyticsReportCore_ = function () { return { data: {} }; };
  context.refreshSessionReconciliationReportCore_ = function () { return { data: {} }; };
  context.refreshExternalAssignmentReportCore_ = function (options) {
    reportCall = options;
    return { data: { participant_count: 5, session_count: 4 } };
  };

  context.refreshAnalyticsCore_({ mode: 'quick' });

  assert.ok(reportCall, 'Analytics refresh should also rebuild the external assignment report');
  assert.strictEqual(reportCall.records.length, buildSampleRecords().length, 'Analytics refresh should pass the active fact records to the report refresh');
}

function testReportBuilderFallsBackToTrainingSessionCostsAndFormatsDates() {
  const context = createContext();
  loadGasFiles(context, [
    'Constants.gs',
    'Utils.gs',
    'TrainingSessionService.gs',
    'ExternalAssignmentReportService.gs',
  ]);

  const reportRecords = [
    {
      row_status: 'ACTIVE',
      emp_id: 'E100',
      full_name: 'Nguyen Thi A',
      department: 'P. Xuat nhap khau',
      division: 'Khoi Kinh doanh',
      job_title: 'Truong bo phan',
      session_id: 'SES-XKN01',
      session_code: 'XKN26_01-202601-01',
      course_id: 'XKN26_01',
      course_name: 'Huong dan XKN',
      training_date: new Date('2026-01-25T00:00:00+07:00'),
      training_month: '2026-01',
      delivery_type: 'Đào tạo bên ngoài',
      training_format: 'Cử nhân sự đi học',
      training_unit: 'Lien doan thuong mai',
      location: 'TP. HCM',
      attendance_status: 'Có mặt',
      plan_scope: 'Trong kế hoạch',
      program_type: 'Đào tạo tuân thủ',
      estimated_cost: '',
      avg_cost_per_pax: '',
      cost_per_pax: '',
    },
    {
      row_status: 'ACTIVE',
      emp_id: 'E200',
      full_name: 'Le Van B',
      department: 'P. Kho vận',
      division: 'Khoi Van hanh',
      job_title: 'Nhan vien',
      session_id: 'SES-GN01',
      session_code: 'GIAONHAN26_01-202603-01',
      course_id: 'GIAONHAN26_01',
      course_name: 'Tap huan lai xe',
      training_date: new Date('2026-03-21T00:00:00+07:00'),
      training_month: new Date('2026-03-01T00:00:00+07:00'),
      delivery_type: 'Đào tạo bên ngoài',
      training_format: 'Cử nhân sự đi học',
      training_unit: 'An toan lao dong',
      location: 'Online',
      attendance_status: 'Có mặt',
      plan_scope: 'Ngoài kế hoạch',
      program_type: 'Đào tạo tuân thủ',
      estimated_cost: '',
      avg_cost_per_pax: '',
      cost_per_pax: '',
    },
  ];

  context.getConfigMap_ = function () { return { CURRENT_FISCAL_YEAR: '2026' }; };
  context.getTrainingSessions_ = function () {
    return [
      {
        session_id: 'SES-XKN01',
        session_code: 'XKN26_01-202601-01',
        estimated_cost: 1000000,
        avg_cost_per_pax: '',
        cost_per_pax: '',
      },
      {
        session_id: 'SES-GN01',
        session_code: 'GIAONHAN26_01-202603-01',
        estimated_cost: 5000000,
        avg_cost_per_pax: '',
        cost_per_pax: '',
      },
    ];
  };

  const report = context.buildExternalAssignmentReportData_(reportRecords, {});
  const sessionRows = Array.from(report.sessionRows);
  const detailRows = Array.from(report.detailRows);

  assert.deepStrictEqual(
    sessionRows.map(function (row) { return row.session_total_cost; }),
    [1000000, 5000000],
    'Report should fall back to the current Training Sessions cost when fact rows are missing session cost'
  );
  assert.deepStrictEqual(
    detailRows.map(function (row) { return row.training_date; }),
    ['25/01/2026', '21/03/2026'],
    'Report should render training dates in dd/MM/yyyy format'
  );
}

function testWriterResetsOnlyUsedRangeBeforeRefresh() {
  const context = createContext();
  loadGasFiles(context, [
    'Constants.gs',
    'Utils.gs',
    'TrainingSessionService.gs',
    'AnalyticsPresentationService.gs',
    'ExternalAssignmentReportService.gs',
  ]);

  context.ensureSheet_ = function () {
    return context.__sheetStub;
  };
  context.__sheetStub.lastRow = 18;
  context.__sheetStub.lastColumn = 12;

  const reportData = {
    fiscalYear: '2026',
    refreshedAt: '2026-03-31T18:00:00',
    filters: {
      delivery_type: 'Đào tạo bên ngoài',
      training_format: 'Cử nhân sự đi học',
    },
    kpis: {
      session_count: 1,
      participant_count: 1,
      unique_employee_count: 1,
      total_cost: 1000000,
    },
    sessionRows: [{
      month_key: '2026-01',
      training_date: '25/01/2026',
      session_code: 'XKN26_01-202601-01',
      course_id: 'XKN26_01',
      course_name: 'Huong dan XKN',
      training_unit: 'Lien doan thuong mai',
      location: 'TP. HCM',
      participant_count: 1,
      unique_employee_count: 1,
      departments_participating: 'P. Xuat nhap khau',
      session_total_cost: 1000000,
      avg_cost_per_pax: 1000000,
      cost_basis: 'ESTIMATED_COST',
      cost_basis_label: 'Chi phí dự kiến lớp',
    }],
    departmentRows: [],
    detailRows: [],
  };

  context.writeExternalAssignmentReportSheet_(reportData);

  const breakApartCall = context.__sheetStub.rangeCalls.find(function (call) {
    return call.row === 1 && call.column === 1 && call.numRows === 18 && call.numColumns === 22 && call.breakApart;
  });

  assert.ok(
    breakApartCall,
    'Writer should only reset the used report range before rewriting so refresh can rerun safely'
  );
  assert.strictEqual(context.__sheetStub.clearedContents, undefined, 'Writer should not clear the whole sheet contents');
  assert.strictEqual(context.__sheetStub.clearedFormats, undefined, 'Writer should not clear the whole sheet formats');
}

function testWriterCreatesFilterOnDetailBlockAndRendersHorizontalKpis() {
  const context = createContext();
  loadGasFiles(context, [
    'Constants.gs',
    'Utils.gs',
    'TrainingSessionService.gs',
    'ExternalAssignmentReportService.gs',
  ]);

  context.ensureSheet_ = function () {
    return context.__sheetStub;
  };

  const reportData = {
    fiscalYear: '2026',
    refreshedAt: '2026-03-31T18:00:00',
    filters: {
      delivery_type: 'Đào tạo bên ngoài',
      training_format: 'Cử nhân sự đi học',
    },
    kpis: {
      session_count: 3,
      participant_count: 12,
      unique_employee_count: 11,
      total_cost: 7800000,
    },
    sessionRows: [],
    departmentRows: [],
    detailRows: [{
      year: '2026',
      month_key: '2026-03',
      training_date: '21/03/2026',
      session_code: 'GIAONHAN26_01-202603-01',
      course_id: 'GIAONHAN26_01',
      course_name: 'Tap huan lai xe',
      training_unit: 'An toan lao dong',
      location: 'Online',
      emp_id: '50007088',
      full_name: 'Le Kim Phat',
      department: 'P. Kho van',
      division: 'Khoi Van hanh',
      job_title: 'Nhan vien',
      attendance_status: 'Có mặt',
      plan_scope: 'Ngoài kế hoạch',
      program_type: 'Đào tạo tuân thủ',
      delivery_type: 'Đào tạo bên ngoài',
      training_format: 'Cử nhân sự đi học',
      participant_cost: 500000,
      session_total_cost: 5000000,
      cost_basis_label: 'Chi phí dự kiến lớp',
    }],
  };

  context.writeExternalAssignmentReportSheet_(reportData);

  assert.ok(
    context.__sheetStub.filterRange,
    'Writer should create a ready-to-use filter on the detail table'
  );
  assert.ok(
    context.__sheetStub.filterRange.row > 1,
    'Detail filter should be scoped to the detail block, not the entire report canvas'
  );
  const kpiRows = context.__sheetStub.rangeCalls.filter(function (call) {
    return Array.isArray(call.values) && call.values.some(function (row) {
      return Array.isArray(row) && row.indexOf('Số lớp bên ngoài') !== -1;
    });
  });
  assert.ok(
    kpiRows.length > 0,
    'Writer should render horizontal KPI labels for the polished report header'
  );
}

function testCustomExternalAssignmentReportIsManagedButNotRowRestored() {
  const context = createContext();
  loadGasFiles(context, ['Constants.gs']);
  const appConfig = vm.runInContext('APP_CONFIG', context);

  assert.ok(
    Array.from(appConfig.MANAGED_OUTPUT_SHEETS || []).includes('REPORT_EXTERNAL_ASSIGNMENT'),
    'External assignment report should remain a managed output sheet'
  );
  assert.ok(
    !Array.from(appConfig.RESTORABLE_SHEETS || []).includes('REPORT_EXTERNAL_ASSIGNMENT'),
    'Custom multi-block report should rebuild from analytics refresh instead of row-based restore'
  );
}

try {
  testReportBuilderFiltersCurrentYearAndComputesCostFallback();
  testReportBuilderFallsBackToTrainingSessionCostsAndFormatsDates();
  testManualRefreshOnlyWritesDedicatedSheetAndMenuIsExposed();
  testAnalyticsRefreshAlsoRebuildsExternalAssignmentReport();
  testWriterResetsOnlyUsedRangeBeforeRefresh();
  testWriterCreatesFilterOnDetailBlockAndRendersHorizontalKpis();
  testCustomExternalAssignmentReportIsManagedButNotRowRestored();
  console.log('PASS external-assignment-report-tdd');
} catch (error) {
  console.error(error.stack || error.message || String(error));
  process.exitCode = 1;
}

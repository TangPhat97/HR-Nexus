const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

function createContext() {
  const auditEntries = [];
  const errorEntries = [];

  const FakeDate = class extends Date {
    constructor(value) {
      super(value || '2026-04-01T10:00:00.000Z');
    }
    static now() {
      return FakeDate.__now;
    }
  };
  FakeDate.__now = 1000;

  const context = {
    console,
    require,
    Buffer,
    Math,
    JSON,
    Date: FakeDate,
    Logger: { log() {} },
    logAudit_(action, status, details) {
      auditEntries.push({ action, status, details });
    },
    logError_(errorCode, action, error, details) {
      errorEntries.push({ errorCode, action, error, details });
    },
    requireRole_() {},
    withActionLock_(key, executor) {
      return executor();
    },
    nowIsoString_() {
      return '2026-04-01T10:00:00';
    },
    SpreadsheetApp: {
      getUi() {
        return {
          ButtonSet: { OK: 'OK' },
          alert() {},
        };
      },
    },
    __auditEntries: auditEntries,
    __errorEntries: errorEntries,
  };
  context.global = context;
  context.globalThis = context;
  return vm.createContext(context);
}

function loadErrorHandler(context) {
  const filePath = path.join(process.cwd(), 'gas', 'ErrorHandler.gs');
  vm.runInContext(fs.readFileSync(filePath, 'utf8'), context, { filename: filePath });
}

function testRunActionLogsDurationForSuccess() {
  const context = createContext();
  loadErrorHandler(context);

  context.Date.__now = 1000;
  const result = context.runAction_('syncTrainingData', function () {
    context.Date.__now = 2450;
    return context.createSuccessResult_('ok', { total_records: 123 });
  }, {
    successMessage: 'Đã đồng bộ dữ liệu đào tạo',
    targetSheet: 'Dữ liệu đào tạo',
  });

  assert.strictEqual(result.ok, true);
  assert.strictEqual(context.__auditEntries.length, 1, 'Successful action should create one audit entry');
  assert.strictEqual(context.__auditEntries[0].status, 'SUCCESS');
  assert.strictEqual(context.__auditEntries[0].details.duration_ms, 1450, 'Audit details should include live execution duration');
  assert.strictEqual(context.__auditEntries[0].details.started_at, '2026-04-01T10:00:00');
  assert.strictEqual(context.__auditEntries[0].details.finished_at, '2026-04-01T10:00:00');
  assert.ok(
    context.__auditEntries[0].details.message.indexOf('1450 ms') !== -1,
    'Audit message should show duration directly in Nhật ký tác động'
  );
}

function testRunActionLogsDurationForFailure() {
  const context = createContext();
  loadErrorHandler(context);

  context.Date.__now = 2000;
  let threw = false;
  try {
    context.runAction_('refreshAnalyticsFullRebuild', function () {
      context.Date.__now = 3125;
      throw new Error('Boom');
    }, {
      targetSheet: 'Phân tích xu hướng',
    });
  } catch (error) {
    threw = true;
    assert.strictEqual(error.message, 'Boom');
  }

  assert.strictEqual(threw, true, 'Action should still throw on failure');
  assert.strictEqual(context.__auditEntries.length, 1, 'Failed action should still create an audit entry');
  assert.strictEqual(context.__auditEntries[0].status, 'FAILED');
  assert.strictEqual(context.__auditEntries[0].details.duration_ms, 1125, 'Failed audit entry should include execution duration');
  assert.ok(
    context.__auditEntries[0].details.message.indexOf('1125 ms') !== -1,
    'Failed audit message should still surface duration'
  );
  assert.strictEqual(context.__errorEntries.length, 1, 'Failure should still be sent to error log');
}

function run() {
  testRunActionLogsDurationForSuccess();
  testRunActionLogsDurationForFailure();
  console.log('PASS action-telemetry-tdd');
}

run();

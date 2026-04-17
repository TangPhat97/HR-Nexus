const assert = require('assert');
const fs = require('fs');
const path = require('path');

const GAS_DIR = path.join(process.cwd(), 'gas');
const FILE_EXTENSIONS = new Set(['.gs', '.html', '.json']);
const SUSPECT_PATTERNS = [
  /Kh\?i t\?o h\? th\?ng/,
  /M\? Trang b\?t d\?u/,
  /Hu\?ng d\?n nh\?p li\?u/,
  /Ki\?m tra d\? li\?u/,
  /Làm m\?i data raw/,
  /kh\?ng kh\?p schema/i,
  /d\? li\?u/,
];

function listGasFiles(directory) {
  return fs.readdirSync(directory, { withFileTypes: true }).flatMap((entry) => {
    const fullPath = path.join(directory, entry.name);
    if (entry.isDirectory()) {
      return listGasFiles(fullPath);
    }
    if (!FILE_EXTENSIONS.has(path.extname(entry.name))) {
      return [];
    }
    return [fullPath];
  });
}

function runTests() {
  const gasFiles = listGasFiles(GAS_DIR);
  assert.ok(gasFiles.length > 0, 'Expected GAS source files to exist');

  gasFiles.forEach((filePath) => {
    assert.doesNotThrow(
      () => fs.readFileSync(filePath, 'utf8'),
      'Source file must be UTF-8 decodable: ' + filePath
    );
  });

  const suspiciousHits = [];
  gasFiles.forEach((filePath) => {
    const text = fs.readFileSync(filePath, 'utf8');
    SUSPECT_PATTERNS.forEach((pattern) => {
      if (pattern.test(text)) {
        suspiciousHits.push(filePath + ' :: ' + pattern);
      }
    });
  });

  assert.deepStrictEqual(
    suspiciousHits,
    [],
    'Source should not contain mojibake-style Vietnamese UI strings'
  );
}

runTests();

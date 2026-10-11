'use strict';
// Shared formulas: Excel stores copied formulas once and renumbers the groups on save, so cells must carry their
// own expanded formula - otherwise a renumbering looks like an edit and blocks merges of unrelated rows.
const test = require('node:test');
const assert = require('node:assert/strict');
const { shiftFormula } = require('../src/engine/ingest');

test('relative references move with the cell; absolute ones stay put', () => {
  assert.equal(shiftFormula('CEILING(D6/4,1)', 3, 0), 'CEILING(D9/4,1)');
  assert.equal(shiftFormula('E6*D6+$F$6+G$6*$H6', 2, 1), 'F8*E8+$F$6+H$6*$H8');
  assert.equal(shiftFormula('SUM(A1:B2)', 0, 26), 'SUM(AA1:AB2)');
});

test('function names, quoted text and quoted sheet names are not treated as references', () => {
  assert.equal(shiftFormula('LOG10(A1)', 1, 0), 'LOG10(A2)');
  assert.equal(shiftFormula('IF(A1="B2",B2,0)', 1, 0), 'IF(A2="B2",B3,0)');
  assert.equal(shiftFormula("'Q1 2024'!C3+Data!C3", 1, 0), "'Q1 2024'!C4+Data!C4");
});

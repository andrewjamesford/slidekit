import test from 'node:test';
import assert from 'node:assert/strict';
import { cssFontFamily } from '../src/css-font-family.js';
test('numeric multiword families are quoted while CSS fallback stacks retain semantics', () => {
  assert.equal(cssFontFamily('Source Sans 3'), '"Source Sans 3"');
  assert.equal(cssFontFamily('Roboto'), '"Roboto"');
  assert.equal(cssFontFamily('Source Sans 3, Arial, sans-serif'), '"Source Sans 3", "Arial", sans-serif');
  assert.equal(cssFontFamily('Noto Sans SC'), '"Noto Sans SC"');
});

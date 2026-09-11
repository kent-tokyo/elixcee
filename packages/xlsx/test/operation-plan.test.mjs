import assert from 'node:assert/strict';
import runtime from '../src/runtime.cjs';
import {
  executeOperationPlan,
  executeOperationPlanOnEditor,
  createOperationPluginRegistry,
  validateOperationPlugin,
  validateOperationPlan,
} from '../src/runtime.mjs';

assert.equal(runtime.executeOperationPlan, executeOperationPlan);
assert.equal(runtime.executeOperationPlanOnEditor, executeOperationPlanOnEditor);
assert.equal(runtime.createOperationPluginRegistry, createOperationPluginRegistry);
assert.equal(runtime.validateOperationPlugin, validateOperationPlugin);

const plan = { operations: [{ kind: 'setNumber', sheet: 'Sheet1', row: 2, col: 27, value: 42 }] };
const capabilities = ['cell.write.number', 'cell.write.string', 'cell.write.boolean'];
const validated = validateOperationPlan(plan, { capabilities });
assert.equal(validated[0].ref, 'AA2');

const workbook = { SheetNames: ['Sheet1'], Sheets: { Sheet1: { AA2: { t: 'n', v: 1 } } } };
const dryRun = executeOperationPlan(workbook, plan, { capabilities });
assert.equal(dryRun.dryRun, true);
assert.equal(workbook.Sheets.Sheet1.AA2.v, 1);
assert.equal(dryRun.changes[0].previous.v, 1);

const applied = executeOperationPlan(workbook, plan, { capabilities, apply: true });
assert.equal(applied.applied, true);
assert.equal(workbook.Sheets.Sheet1.AA2.v, 42);

const typedWorkbook = { Sheets: { Sheet1: {} } };
const typedPlan = {
  operations: [
    { kind: 'setString', sheet: 'Sheet1', row: 1, col: 1, value: 'ready' },
    { kind: 'setBoolean', sheet: 'Sheet1', row: 1, col: 2, value: true },
  ],
};
const typed = executeOperationPlan(typedWorkbook, typedPlan, { capabilities, apply: true });
assert.deepEqual(typedWorkbook.Sheets.Sheet1.A1, { t: 's', v: 'ready' });
assert.deepEqual(typedWorkbook.Sheets.Sheet1.B1, { t: 'b', v: true });
assert.deepEqual(typed.changes.map(({ kind, next }) => [kind, next.t]), [['setString', 's'], ['setBoolean', 'b']]);

assert.throws(
  () => validateOperationPlan(plan, { capabilities: [] }),
  (error) => error.code === 'ELIXCEE_CAPABILITY_DENIED'
);
assert.throws(
  () => validateOperationPlan({ operations: [{ ...plan.operations[0], row: 0 }] }, { capabilities }),
  (error) => error.code === 'ELIXCEE_INVALID_OPERATION_PLAN'
);
assert.throws(
  () => executeOperationPlan(workbook, { operations: [{ ...plan.operations[0], sheet: 'Missing' }] }, { capabilities }),
  (error) => error.code === 'ELIXCEE_SHEET_NOT_FOUND'
);

const editor = {
  calls: [],
  beginTransaction() { this.calls.push('begin'); },
  commitTransaction() { this.calls.push('commit'); return true; },
  abortTransaction() { this.calls.push('abort'); },
  setNumber(...args) { this.calls.push(['number', ...args]); },
  setString(...args) { this.calls.push(['string', ...args]); },
  setBoolean(...args) { this.calls.push(['boolean', ...args]); },
};
const editorPlan = executeOperationPlanOnEditor(editor, typedPlan, { capabilities });
assert.equal(editorPlan.dryRun, true);
assert.deepEqual(editor.calls, []);
executeOperationPlanOnEditor(editor, typedPlan, { capabilities, apply: true });
assert.deepEqual(editor.calls, [
  'begin',
  ['string', 'Sheet1', 1, 1, 'ready'],
  ['boolean', 'Sheet1', 1, 2, true],
  'commit',
]);

const failingEditor = {
  calls: [],
  beginTransaction() { this.calls.push('begin'); },
  commitTransaction() { this.calls.push('commit'); return true; },
  abortTransaction() { this.calls.push('abort'); },
  setNumber() { throw new Error('setter failed'); },
  setString() {},
  setBoolean() {},
};
assert.throws(
  () => executeOperationPlanOnEditor(failingEditor, plan, { capabilities, apply: true }),
  /setter failed/
);
assert.deepEqual(failingEditor.calls, ['begin', 'abort']);

const plugin = {
  name: 'write-ready-cell',
  capabilities: ['cell.write.string'],
  operations: [{ kind: 'setString', sheet: 'Sheet1', row: 1, col: 1, value: 'ready' }],
  limits: { maxOperations: 4, maxPlanBytes: 4096 },
};
const validatedPlugin = validateOperationPlugin(plugin);
assert.equal(validatedPlugin.name, 'write-ready-cell');
assert.equal(Object.isFrozen(validatedPlugin), true);
assert.throws(
  () => validateOperationPlugin({ ...plugin, run() {} }),
  (error) => error.code === 'ELIXCEE_PLUGIN_CODE_DENIED'
);
assert.throws(
  () => validateOperationPlugin({ ...plugin, module: './plugin.mjs' }),
  (error) => error.code === 'ELIXCEE_PLUGIN_SCHEMA_DENIED'
);
const registry = createOperationPluginRegistry({ maxPlugins: 1 });
registry.register(plugin);
assert.deepEqual(registry.list().map(({ name }) => name), ['write-ready-cell']);
const pluginWorkbook = { SheetNames: ['Sheet1'], Sheets: { Sheet1: {} } };
const pluginDryRun = registry.execute('write-ready-cell', pluginWorkbook, { capabilities: ['cell.write.string'] });
assert.equal(pluginDryRun.dryRun, true);
assert.equal(pluginWorkbook.Sheets.Sheet1.A1, undefined);
registry.execute('write-ready-cell', pluginWorkbook, { capabilities: ['cell.write.string'], apply: true });
assert.deepEqual(pluginWorkbook.Sheets.Sheet1.A1, { t: 's', v: 'ready' });
assert.throws(
  () => registry.execute('write-ready-cell', pluginWorkbook, { capabilities: [] }),
  (error) => error.code === 'ELIXCEE_CAPABILITY_DENIED'
);
assert.throws(
  () => registry.register(plugin),
  (error) => error.code === 'ELIXCEE_PLUGIN_DUPLICATE'
);
assert.throws(
  () => createOperationPluginRegistry({ maxOperations: 1 }).register({
    ...plugin,
    name: 'over-budget',
    limits: { maxOperations: 2 },
  }),
  (error) => error.code === 'ELIXCEE_PLUGIN_BUDGET'
);

console.log('operation plan: ok');

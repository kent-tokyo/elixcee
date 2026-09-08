'use strict';

// A deliberately data-only operation boundary for AI and automation callers.
// It never evaluates caller-supplied code or performs I/O. Plans are validated
// before any workbook mutation, and dry-run is the default.

const DEFAULT_LIMITS = Object.freeze({
  maxOperations: 1000,
  maxPlanBytes: 1024 * 1024,
  maxPlugins: 64,
});

function operationError(message, code = 'ELIXCEE_INVALID_OPERATION_PLAN') {
  const error = new Error(message);
  error.code = code;
  return error;
}

function byteLength(value) {
  return new TextEncoder().encode(value).length;
}

function cellRef(row, col) {
  let n = col;
  let letters = '';
  while (n > 0) {
    const remainder = (n - 1) % 26;
    letters = String.fromCharCode(65 + remainder) + letters;
    n = Math.floor((n - 1) / 26);
  }
  return `${letters}${row}`;
}

function validateOperationPlan(plan, options = {}) {
  if (!plan || typeof plan !== 'object' || Array.isArray(plan)) {
    throw operationError('operation plan must be an object');
  }
  if (!Array.isArray(plan.operations)) {
    throw operationError('operation plan must contain an operations array');
  }
  const limits = { ...DEFAULT_LIMITS, ...options };
  const serialized = JSON.stringify(plan);
  if (byteLength(serialized) > limits.maxPlanBytes) {
    throw operationError(`operation plan exceeds maxPlanBytes (${limits.maxPlanBytes})`, 'ELIXCEE_OPERATION_BUDGET');
  }
  if (plan.operations.length > limits.maxOperations) {
    throw operationError(`operation plan exceeds maxOperations (${limits.maxOperations})`, 'ELIXCEE_OPERATION_BUDGET');
  }
  const capabilities = new Set(options.capabilities || []);
  return plan.operations.map((operation, index) => {
    const allowedKinds = new Set(['setNumber', 'setString', 'setBoolean']);
    if (!operation || !allowedKinds.has(operation.kind)) {
      throw operationError(`operation ${index} is not an allowed typed cell operation`, 'ELIXCEE_CAPABILITY_DENIED');
    }
    const capability = `cell.write.${operation.kind.slice(3).toLowerCase()}`;
    if (!capabilities.has(capability)) {
      throw operationError(`capability ${capability} is required`, 'ELIXCEE_CAPABILITY_DENIED');
    }
    if (typeof operation.sheet !== 'string' || operation.sheet.trim() === '' || operation.sheet.length > 255) {
      throw operationError(`operation ${index} has an invalid sheet`);
    }
    if (!Number.isInteger(operation.row) || operation.row < 1 || operation.row > 1048576) {
      throw operationError(`operation ${index} has an invalid 1-based row`);
    }
    if (!Number.isInteger(operation.col) || operation.col < 1 || operation.col > 16384) {
      throw operationError(`operation ${index} has an invalid 1-based column`);
    }
    if (operation.kind === 'setNumber' && (typeof operation.value !== 'number' || !Number.isFinite(operation.value))) {
      throw operationError(`operation ${index} value must be a finite number`);
    }
    if (operation.kind === 'setString' && typeof operation.value !== 'string') {
      throw operationError(`operation ${index} value must be a string`);
    }
    if (operation.kind === 'setBoolean' && typeof operation.value !== 'boolean') {
      throw operationError(`operation ${index} value must be a boolean`);
    }
    return Object.freeze({
      kind: operation.kind,
      sheet: operation.sheet,
      row: operation.row,
      col: operation.col,
      ref: cellRef(operation.row, operation.col),
      value: operation.value,
    });
  });
}

function executeOperationPlan(workbook, plan, options = {}) {
  if (!workbook || typeof workbook !== 'object' || !workbook.Sheets) {
    throw operationError('workbook must be a WorkBook object');
  }
  const operations = validateOperationPlan(plan, options);
  const changes = operations.map((operation) => {
    const sheet = workbook.Sheets[operation.sheet];
    if (!sheet) throw operationError(`sheet '${operation.sheet}' was not found`, 'ELIXCEE_SHEET_NOT_FOUND');
    const previous = sheet[operation.ref];
    const type = operation.kind === 'setNumber' ? 'n' : operation.kind === 'setString' ? 's' : 'b';
    return { ...operation, previous: previous === undefined ? null : { ...previous }, next: { t: type, v: operation.value } };
  });
  if (options.apply === true) {
    for (const change of changes) workbook.Sheets[change.sheet][change.ref] = change.next;
  }
  return Object.freeze({ dryRun: options.apply !== true, applied: options.apply === true, changes });
}

function executeOperationPlanOnEditor(editor, plan, options = {}) {
  if (!editor || typeof editor !== 'object'
      || typeof editor.beginTransaction !== 'function'
      || typeof editor.commitTransaction !== 'function'
      || typeof editor.abortTransaction !== 'function') {
    throw operationError('editor must be a WorkbookEditor-like object');
  }
  const operations = validateOperationPlan(plan, options);
  const result = Object.freeze({
    dryRun: options.apply !== true,
    applied: options.apply === true,
    changes: operations.map((operation) => Object.freeze({ ...operation })),
  });
  if (options.apply !== true) return result;

  editor.beginTransaction();
  try {
    for (const operation of operations) {
      if (operation.kind === 'setNumber') editor.setNumber(operation.sheet, operation.row, operation.col, operation.value);
      else if (operation.kind === 'setString') editor.setString(operation.sheet, operation.row, operation.col, operation.value);
      else editor.setBoolean(operation.sheet, operation.row, operation.col, operation.value);
    }
    if (!editor.commitTransaction()) {
      throw operationError('editor transaction was not committed', 'ELIXCEE_TRANSACTION_FAILED');
    }
  } catch (error) {
    try { editor.abortTransaction(); } catch { /* preserve the original error */ }
    throw error;
  }
  return result;
}

// A plugin is intentionally a named, immutable operation plan. It is not a
// JavaScript callback, module, or script: registration cannot add arbitrary
// code or I/O to this boundary. Capabilities are still supplied at execution
// time, so registration is not an implicit grant.
function validateOperationPlugin(plugin, options = {}) {
  if (!plugin || typeof plugin !== 'object' || Array.isArray(plugin)) {
    throw operationError('operation plugin must be an object');
  }
  if (typeof plugin.name !== 'string' || plugin.name.trim() === '' || plugin.name.length > 128) {
    throw operationError('operation plugin name must be a non-empty string of at most 128 characters');
  }
  if (Object.prototype.hasOwnProperty.call(plugin, 'run')
      || Object.prototype.hasOwnProperty.call(plugin, 'execute')) {
    throw operationError('operation plugins cannot contain executable callbacks', 'ELIXCEE_PLUGIN_CODE_DENIED');
  }
  const allowedKeys = new Set(['name', 'capabilities', 'operations', 'limits']);
  const unexpectedKey = Object.keys(plugin).find((key) => !allowedKeys.has(key));
  if (unexpectedKey !== undefined) {
    throw operationError(`operation plugin field '${unexpectedKey}' is not allowed`, 'ELIXCEE_PLUGIN_SCHEMA_DENIED');
  }
  if (!Array.isArray(plugin.capabilities) || plugin.capabilities.some((value) => typeof value !== 'string')) {
    throw operationError('operation plugin capabilities must be an array of strings');
  }
  const capabilities = [...new Set(plugin.capabilities)];
  const ceiling = { ...DEFAULT_LIMITS, ...(options.limits || {}) };
  const requested = plugin.limits || {};
  for (const key of ['maxOperations', 'maxPlanBytes']) {
    if (requested[key] !== undefined && requested[key] > ceiling[key]) {
      throw operationError(`operation plugin ${key} exceeds the registry limit (${ceiling[key]})`, 'ELIXCEE_PLUGIN_BUDGET');
    }
  }
  const limits = { ...ceiling, ...requested };
  if (!Number.isInteger(limits.maxOperations) || limits.maxOperations < 0
      || !Number.isInteger(limits.maxPlanBytes) || limits.maxPlanBytes < 0) {
    throw operationError('operation plugin limits must be non-negative integers', 'ELIXCEE_PLUGIN_BUDGET');
  }
  const operations = validateOperationPlan(plugin, { ...limits, capabilities });
  return Object.freeze({
    name: plugin.name,
    capabilities: Object.freeze(capabilities),
    limits: Object.freeze({
      maxOperations: limits.maxOperations,
      maxPlanBytes: limits.maxPlanBytes,
    }),
    plan: Object.freeze({ operations: Object.freeze(operations.map((operation) => Object.freeze({
      kind: operation.kind,
      sheet: operation.sheet,
      row: operation.row,
      col: operation.col,
      value: operation.value,
    }))) }),
  });
}

function createOperationPluginRegistry(options = {}) {
  const limits = { ...DEFAULT_LIMITS, ...options };
  if (!Number.isInteger(limits.maxPlugins) || limits.maxPlugins < 0) {
    throw operationError('maxPlugins must be a non-negative integer', 'ELIXCEE_PLUGIN_BUDGET');
  }
  const plugins = new Map();
  return Object.freeze({
    register(plugin) {
      if (plugins.size >= limits.maxPlugins && !plugins.has(plugin?.name)) {
        throw operationError(`operation plugin registry exceeds maxPlugins (${limits.maxPlugins})`, 'ELIXCEE_PLUGIN_BUDGET');
      }
      const validated = validateOperationPlugin(plugin, { limits });
      if (plugins.has(validated.name)) {
        throw operationError(`operation plugin '${validated.name}' is already registered`, 'ELIXCEE_PLUGIN_DUPLICATE');
      }
      plugins.set(validated.name, validated);
      return validated;
    },
    list() {
      return Object.freeze([...plugins.values()].map((plugin) => Object.freeze({
        name: plugin.name,
        capabilities: plugin.capabilities,
        limits: plugin.limits,
      })));
    },
    execute(name, target, options = {}) {
      if (typeof name !== 'string' || !plugins.has(name)) {
        throw operationError(`operation plugin '${name}' was not found`, 'ELIXCEE_PLUGIN_NOT_FOUND');
      }
      const plugin = plugins.get(name);
      const granted = new Set(options.capabilities || []);
      for (const capability of plugin.capabilities) {
        if (!granted.has(capability)) {
          throw operationError(`capability ${capability} is required by plugin '${name}'`, 'ELIXCEE_CAPABILITY_DENIED');
        }
      }
      const executionOptions = {
        ...options,
        capabilities: [...granted],
        maxOperations: Math.min(plugin.limits.maxOperations, options.maxOperations ?? plugin.limits.maxOperations),
        maxPlanBytes: Math.min(plugin.limits.maxPlanBytes, options.maxPlanBytes ?? plugin.limits.maxPlanBytes),
      };
      if (target && typeof target.beginTransaction === 'function') {
        return executeOperationPlanOnEditor(target, plugin.plan, executionOptions);
      }
      return executeOperationPlan(target, plugin.plan, executionOptions);
    },
  });
}

module.exports = {
  DEFAULT_LIMITS,
  validateOperationPlan,
  executeOperationPlan,
  executeOperationPlanOnEditor,
  validateOperationPlugin,
  createOperationPluginRegistry,
};

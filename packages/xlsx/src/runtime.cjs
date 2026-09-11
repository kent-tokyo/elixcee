'use strict';

// Opt-in elixcee runtime API. Kept on a separate subpath so the root package
// remains an exact-compatible SheetJS-shaped surface.
const bridge = require('./internal/wasm/elixcee_wasm.node.cjs');
const {
  validateOperationPlan,
  executeOperationPlan,
  executeOperationPlanOnEditor,
  validateOperationPlugin,
  createOperationPluginRegistry,
} = require('./internal/operation-plan.cjs');

module.exports = {
  WorkbookEditor: bridge.WorkbookEditor,
  calculateWorkbook: bridge.calculateWorkbook,
  diagnoseWorkbook: bridge.diagnoseWorkbook,
  validateOperationPlan,
  executeOperationPlan,
  executeOperationPlanOnEditor,
  validateOperationPlugin,
  createOperationPluginRegistry,
};

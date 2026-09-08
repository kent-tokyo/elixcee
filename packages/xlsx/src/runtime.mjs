// Node ESM entry for the opt-in stateful Rust/WASM runtime.
export { WorkbookEditor, calculateWorkbook, diagnoseWorkbook } from './internal/wasm/elixcee_wasm.node.cjs';
export {
  validateOperationPlan,
  executeOperationPlan,
  executeOperationPlanOnEditor,
  validateOperationPlugin,
  createOperationPluginRegistry,
} from './internal/operation-plan.mjs';

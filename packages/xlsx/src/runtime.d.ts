export declare class WorkbookEditor {
  constructor(bytes: Uint8Array);
  setNumber(sheet: string, row: number, col: number, value: number): void;
  setString(sheet: string, row: number, col: number, value: string): void;
  setBoolean(sheet: string, row: number, col: number, value: boolean): void;
  recalculate(): string;
  snapshot(): string;
  undo(): boolean;
  redo(): boolean;
  beginTransaction(): void;
  commitTransaction(): boolean;
  abortTransaction(): boolean;
  canUndo(): boolean;
  canRedo(): boolean;
}

export declare function calculateWorkbook(bytes: Uint8Array): string;
export declare function diagnoseWorkbook(bytes: Uint8Array): string;

export interface OperationPlan {
  operations: Array<{
    kind: 'setNumber' | 'setString' | 'setBoolean';
    sheet: string;
    row: number;
    col: number;
    value: number | string | boolean;
  }>;
}

export interface OperationPlanOptions {
  capabilities?: string[];
  maxOperations?: number;
  maxPlanBytes?: number;
  apply?: boolean;
}

export declare function validateOperationPlan(
  plan: OperationPlan,
  options?: OperationPlanOptions
): Array<{
  kind: 'setNumber' | 'setString' | 'setBoolean';
  sheet: string;
  row: number;
  col: number;
  ref: string;
  value: number | string | boolean;
}>;

export declare function executeOperationPlan(
  workbook: Record<string, unknown>,
  plan: OperationPlan,
  options?: OperationPlanOptions
): { dryRun: boolean; applied: boolean; changes: Array<Record<string, unknown>> };

export declare function executeOperationPlanOnEditor(
  editor: WorkbookEditor,
  plan: OperationPlan,
  options?: OperationPlanOptions
): { dryRun: boolean; applied: boolean; changes: Array<Record<string, unknown>> };

export interface OperationPlugin {
  name: string;
  capabilities: string[];
  operations: OperationPlan['operations'];
  limits?: { maxOperations?: number; maxPlanBytes?: number };
}

export declare function validateOperationPlugin(
  plugin: OperationPlugin,
  options?: { limits?: { maxOperations?: number; maxPlanBytes?: number } }
): Readonly<{
  name: string;
  capabilities: readonly string[];
  limits: Readonly<{ maxOperations: number; maxPlanBytes: number }>;
  plan: Readonly<{ operations: ReadonlyArray<OperationPlan['operations'][number]> }>;
}>;

export interface OperationPluginRegistry {
  register(plugin: OperationPlugin): unknown;
  list(): ReadonlyArray<{ name: string; capabilities: readonly string[]; limits: Readonly<{ maxOperations: number; maxPlanBytes: number }> }>;
  execute(
    name: string,
    target: Record<string, unknown>,
    options?: OperationPlanOptions
  ): { dryRun: boolean; applied: boolean; changes: Array<Record<string, unknown>> };
}

export declare function createOperationPluginRegistry(
  options?: { maxPlugins?: number; maxOperations?: number; maxPlanBytes?: number }
): OperationPluginRegistry;

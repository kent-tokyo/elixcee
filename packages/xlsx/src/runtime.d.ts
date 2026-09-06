export declare class WorkbookEditor {
  constructor(bytes: Uint8Array);
  setNumber(sheet: string, row: number, col: number, value: number): void;
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

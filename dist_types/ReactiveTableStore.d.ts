/// <reference types="office-js" />
import { ColumnDef, RowData, UseTableParams, CellType, RowNode } from "./ReactiveTableStore.interface";
declare class ReactiveTableStore {
    private tableName;
    private sheetName;
    private columnDefs;
    private _loading;
    private _rows;
    private tablePositionColumn;
    private tablePositionRow;
    constructor(tableName: string, sheetName: string, columnDefs: ColumnDef[]);
    init(): Promise<void>;
    /**
     * Getters
     */
    get headers(): string[];
    get numberOfColumns(): number;
    get loading(): boolean;
    get rows(): {
        readonly [Symbol.iterator]: () => IterableIterator<RowNode>;
        readonly next: (...args: [] | [undefined]) => IteratorResult<RowNode, any>;
        readonly return?: ((value?: any) => IteratorResult<RowNode, any>) | undefined;
        readonly throw?: ((e?: any) => IteratorResult<RowNode, any>) | undefined;
    };
    /**
     * Initialization Functions
     */
    private constructTable;
    private loadRowsFromExcel;
    private registerEventHandlers;
    /**
     * Internal Helper Functions
     */
    runWithTable(func: (table: Excel.Table, context: Excel.RequestContext) => Promise<void>): Promise<void>;
    /**
     * Mutation Functions
     */
    addRow(row: RowData, index?: number): Promise<void>;
    deleteRow(index: number): Promise<void>;
    updateRow(index: number, data: RowData): Promise<void>;
    updateRowValue(rowIndex: number, columnIndex: number, value: CellType): void;
    get state(): RowNode[];
}
export declare const useTableStore: ({ tableName, sheetName, columnDefs }: UseTableParams) => ReactiveTableStore;
export {};

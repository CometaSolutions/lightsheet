// TODO We should reconsider the role/scope of this type
export type LightsheetOptions = {
  sheetName: string;
  data?: any[];
  onCellChange?: (colIndex: number, rowIndex: number, value: any) => void;
  onCellClick?: (colIndex: number, rowIndex: number) => void;
  onReady?: () => void;
  defaultRowCount?: number;
  defaultColCount?: number;
  isReadOnly?: boolean;
  toolbarOptions?: ToolbarOptions;
};

export type ToolbarOptions = {
  showToolbar?: boolean;
  items?: string[];
  element?: HTMLElement;
};

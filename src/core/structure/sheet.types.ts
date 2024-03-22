import { ColumnKey, RowKey } from "./key/keyTypes.ts";

import { CellState } from "./cell/cellState.ts";

export type PositionInfo = {
  columnKey?: ColumnKey;
  rowKey?: RowKey;
};

export type CellInfo = {
  position: PositionInfo;
  value?: string;
  state?: CellState;
};

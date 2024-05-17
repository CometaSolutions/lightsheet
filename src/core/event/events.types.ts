import { Coordinate } from "../../utils/common.types.ts";
import { CellState } from "../structure/cell/cellState.ts";

export type UISetCellPayload = {
  position: Coordinate;
  rawValue: string;
};

export type CoreSetCellPayload = {
  position: Coordinate;
  rawValue: string;
  formattedValue: string;
  state: CellState;

  clearCell: boolean;
  clearRow: boolean;
};

export type SetColumnLabelPayload = {
  columnIndex: number;
  label: string;
};

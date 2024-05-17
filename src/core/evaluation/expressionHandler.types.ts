import { SheetKey } from "../structure/key/keyTypes.ts";
import { Coordinate } from "../../utils/common.types.ts";
import { CellState } from "../structure/cell/cellState.ts";

export type CellSheetPosition = {
  sheetKey: SheetKey;
  position: Coordinate;
};

export type EvaluationResult = {
  value?: string;
  result: CellState;
  references: CellSheetPosition[];
};

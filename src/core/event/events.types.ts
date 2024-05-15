import { Coordinate } from "../../utils/common.types.ts";

export type UISetCellPayload = {
  position: Coordinate;
  rawValue: string;
};

export type CoreSetCellPayload = {
  position: Coordinate;
  rawValue: string;
  formattedValue: string;

  clearCell: boolean;
  clearRow: boolean;
};

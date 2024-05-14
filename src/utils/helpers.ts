import CellGroup from "../core/structure/group/cellGroup";
import Column from "../core/structure/group/column";
import { ColumnKey, RowKey } from "../core/structure/key/keyTypes";
import { PositionInfo } from "../core/structure/sheet.types";

export default class LightsheetHelper {
  static generateColumnLabel = (rowIndex: number) => {
    let label = "";
    const alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
    while (rowIndex > 0) {
      rowIndex--; // Adjust index to start from 0
      label = alphabet[rowIndex % 26] + label;
      rowIndex = Math.floor(rowIndex / 26);
    }
    return label || "A"; // Return "A" if index is 0
  };

  static getKeyPairFor = (cellGroup: CellGroup<ColumnKey | RowKey>, opposingKey: ColumnKey | RowKey): PositionInfo => {
    return {
      columnKey: CellGroup instanceof Column ? cellGroup.key as ColumnKey : opposingKey as ColumnKey,
      rowKey: CellGroup instanceof Column ? opposingKey as RowKey : cellGroup.key as RowKey,
    };
  }
}

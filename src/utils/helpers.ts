import CellGroup from "../core/structure/group/cellGroup";
import Column from "../core/structure/group/column";
import { ColumnKey, RowKey } from "../core/structure/key/keyTypes";
import { PositionInfo } from "../core/structure/sheet.types";

export default class LightsheetHelper {
  static generateColumnLabel = (columnIndex: number) => {
    let label = "";
    const alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
    while (columnIndex > 0) {
      label = alphabet[columnIndex % 26] + label;
      columnIndex = Math.floor(columnIndex / 26);
    }
    return label || "A"; // Return "A" if index is 0
  };

  static resolveColumnIndex(column: string): number {
    let index = 0;
    for (let i = 0; i < column.length; i++) {
      if (column[i] < "A" || column[i] > "Z") return -1;
      let baseIndex = column.charCodeAt(i) - "A".charCodeAt(0);
      if (i != column.length - 1) baseIndex += 1;
      // Character index gives us the exponent: for example, 'AAB' is 26^2 * 1 + 26^1 * 1 + 26^0 * 1 + 1
      const exp = column.length - i - 1;
      index += baseIndex * Math.pow(26, exp);
    }

    return index;
  }

  static getKeyPairFor = (
    cellGroup: CellGroup<ColumnKey | RowKey>,
    opposingKey: ColumnKey | RowKey,
  ): PositionInfo => {
    return {
      columnKey:
        cellGroup instanceof Column
          ? (cellGroup.key as ColumnKey)
          : (opposingKey as ColumnKey),
      rowKey:
        cellGroup instanceof Column
          ? (opposingKey as RowKey)
          : (cellGroup.key as RowKey),
    };
  };
}

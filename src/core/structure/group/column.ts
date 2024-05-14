import CellGroup from "./cellGroup.ts";
import {
  CellKey,
  ColumnKey,
  generateColumnKey,
  RowKey,
} from "../key/keyTypes.ts";
import CellStyle from "../cellStyle.ts";
import LightsheetHelper from "../../../utils/helpers.ts";

export default class Column extends CellGroup<ColumnKey> {
  cellIndex: Map<RowKey, CellKey>;
  cellFormatting: Map<RowKey, CellStyle>;

  constructor(
    width: number,
    position: number,
    name: string | undefined = undefined,
  ) {
    super(width, position, name);
    this.key = generateColumnKey();
    this.cellIndex = new Map<RowKey, CellKey>();
    this.cellFormatting = new Map<RowKey, CellStyle>();
  }

  getWidth(): number {
    return this.size_;
  }

  generateName(): string {
    return LightsheetHelper.generateColumnLabel(this.position + 1);
  }
}

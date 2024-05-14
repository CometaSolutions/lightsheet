import CellGroup from "./cellGroup.ts";
import { CellKey, ColumnKey, generateRowKey, RowKey } from "../key/keyTypes.ts";
import CellStyle from "../cellStyle.ts";

export default class Row extends CellGroup<RowKey> {
  cellIndex: Map<ColumnKey, CellKey>;
  cellFormatting: Map<ColumnKey, CellStyle>;

  constructor(width: number, position: number, name: string | undefined = undefined) {
    super(width, position, name);
    this.key = generateRowKey();
    this.cellIndex = new Map<ColumnKey, CellKey>();
    this.cellFormatting = new Map<ColumnKey, CellStyle>();
  }

  getHeight(): number {
    return this.size_;
  }

  generateName(): string {
    return `${this.position + 1}`
  }
}

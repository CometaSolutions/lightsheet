import { CellKey, ColumnKey, RowKey } from "../key/keyTypes.ts";
import CellStyle from "../cellStyle.ts";

export default abstract class CellGroup<TKey extends ColumnKey | RowKey> {
  key!: TKey;
  size_: number;
  name: string;
  position: number;

  defaultStyle: CellStyle | null;
  cellIndex: Map<ColumnKey | RowKey, CellKey>;
  cellFormatting: Map<ColumnKey | RowKey, CellStyle>;

  constructor(size: number, position: number, name: string | undefined) {
    this.size_ = size;
    this.position = position;
    this.name = name || this.generateName();
    this.defaultStyle = null;
    this.cellIndex = new Map<ColumnKey | RowKey, CellKey>();
    this.cellFormatting = new Map<ColumnKey | RowKey, CellStyle>();
  }

  abstract generateName() : string;
}

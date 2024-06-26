import {
  CellKey,
  ColumnKey,
  generateSheetKey,
  RowKey,
  SheetKey,
} from "./key/keyTypes.ts";
import Cell from "./cell/cell.ts";
import Column from "./group/column.ts";
import Row from "./group/row.ts";
import { CellInfo, PositionInfo, ShiftDirection } from "./sheet.types.ts";
import ExpressionHandler from "../evaluation/expressionHandler.ts";
import CellStyle from "./cellStyle.ts";
import CellGroup from "./group/cellGroup.ts";
import Events from "../event/events.ts";
import LightsheetEvent from "../event/event.ts";
import {
  CoreSetCellPayload,
  SetColumnLabelPayload,
  UISetCellPayload,
} from "../event/events.types.ts";
import EventType from "../event/eventType.ts";
import { CellState } from "./cell/cellState.ts";
import { EvaluationResult } from "../evaluation/expressionHandler.types.ts";
import SheetHolder from "./sheetHolder.ts";
import { CellReference } from "./cell/types.cell.ts";
import { Coordinate } from "../../utils/common.types.ts";
import LightsheetHelper from "../../utils/helpers.ts";

export default class Sheet {
  readonly key: SheetKey;
  readonly name: string;
  private readonly sheetHolder: SheetHolder;

  private readonly cellData: Map<CellKey, Cell>;

  private readonly rows: Map<RowKey, Row>;
  private readonly columns: Map<ColumnKey, Column>;
  private readonly rowPositions: Map<number, RowKey>;
  private readonly columnPositions: Map<number, ColumnKey>;
  private readonly columnLabels: Map<string, ColumnKey>;

  private defaultStyle: any;
  private defaultWidth: number;
  private defaultHeight: number;

  private events: Events;

  constructor(name: string, events: Events | null = null) {
    this.key = generateSheetKey();
    this.name = name;
    this.cellData = new Map<CellKey, Cell>();

    this.sheetHolder = SheetHolder.getInstance();
    this.sheetHolder.addSheet(this);

    this.defaultStyle = new CellStyle(); // TODO This should be configurable.
    this.rows = new Map<RowKey, Row>();
    this.columns = new Map<ColumnKey, Column>();

    this.rowPositions = new Map<number, RowKey>();
    this.columnPositions = new Map<number, ColumnKey>();
    this.columnLabels = new Map<string, ColumnKey>();

    this.defaultWidth = 40;
    this.defaultHeight = 20;

    this.events = events ?? new Events();
    this.registerEvents();
  }

  getRowIndex(rowKey: RowKey): number | undefined {
    return this.rows.get(rowKey)?.position;
  }

  getColumnIndex(colKey: ColumnKey): number | undefined {
    return this.columns.get(colKey)?.position;
  }

  setCellAt(colPos: number, rowPos: number, value: string): CellInfo {
    const position = this.initializePosition(colPos, rowPos);
    return this.setCell(position.columnKey!, position.rowKey!, value)!;
  }

  setCell(colKey: ColumnKey, rowKey: RowKey, formula: string): CellInfo | null {
    if (!this.columns.has(colKey) || !this.rows.has(rowKey)) {
      return null;
    }

    let cell = this.getCell(colKey, rowKey);

    if (!cell) {
      if (formula == "") return null;
      cell = this.createCell(colKey, rowKey, formula);
    }

    cell!.rawValue = formula;

    const colIndex = this.getColumnIndex(colKey)!;
    const rowIndex = this.getRowIndex(rowKey)!;
    const deleted = this.deleteCellIfUnused(colKey, rowKey);
    if (deleted) {
      cell = null;
    } else {
      this.resolveCell(cell!, colKey, rowKey);
    }

    this.emitSetCellEvent(colIndex, rowIndex, cell);

    return {
      rawValue: cell ? cell.rawValue : undefined,
      resolvedValue: cell ? cell.resolvedValue : undefined,
      formattedValue: cell ? cell.formattedValue : undefined,
      state: cell ? cell.state : undefined,
      position: {
        rowKey: this.rows.has(rowKey) ? rowKey : undefined,
        columnKey: this.columns.has(colKey) ? colKey : undefined,
      },
    };
  }

  public moveCell(
    from: Coordinate,
    to: Coordinate,
    moveStyling: boolean = true,
  ) {
    const fromPosition = this.getCellInfoAt(from.column, from.row)?.position;
    let toPosition = this.getCellInfoAt(to.column, to.row)?.position;

    if (!fromPosition) return false;

    if (toPosition) {
      // Cell already exists - delete it first. This may invalidate the position as well.
      this.deleteCell(toPosition.columnKey!, toPosition.rowKey!);
    }

    toPosition = this.initializePosition(to.column, to.row);

    const fromCol = this.columns.get(fromPosition.columnKey!)!;
    const fromRow = this.rows.get(fromPosition.rowKey!)!;

    const toCol = this.columns.get(toPosition.columnKey!)!;
    const toRow = this.rows.get(toPosition.rowKey!)!;

    const cellKey = fromCol.cellIndex.get(fromRow.key)!;
    const style = fromCol.cellFormatting.get(fromRow.key);

    fromCol.cellIndex.delete(fromRow.key);
    toCol.cellIndex.set(toRow.key, cellKey);
    fromRow.cellIndex.delete(fromCol.key);
    toRow.cellIndex.set(toCol.key, cellKey);

    if (style && moveStyling) {
      toCol.cellFormatting.set(toRow.key, style);
      toRow.cellFormatting.set(toCol.key, style);

      fromCol.cellFormatting.delete(fromRow.key);
      fromRow.cellFormatting.delete(fromCol.key);
    }

    this.updateReferenceSymbols(this.cellData.get(cellKey)!, {
      from: from,
      to: to,
    });
    return true;
  }

  public getCellInfoAt(colPos: number, rowPos: number): CellInfo | null {
    const colKey = this.columnPositions.get(colPos);
    const rowKey = this.rowPositions.get(rowPos);
    if (!colKey || !rowKey) return null;

    const cell = this.getCell(colKey, rowKey)!;
    return cell
      ? {
          rawValue: cell.rawValue,
          resolvedValue: cell.resolvedValue,
          formattedValue: cell.formattedValue,
          state: cell.state,
          position: {
            columnKey: colKey,
            rowKey: rowKey,
          },
        }
      : null;
  }

  moveColumn(from: number, to: number): boolean {
    return this.moveCellGroup(from, to, this.columns, this.columnPositions);
  }

  moveRow(from: number, to: number): boolean {
    return this.moveCellGroup(from, to, this.rows, this.rowPositions);
  }

  public setColumnLabel(columnIndex: number, label: string | null): boolean {
    const mathPattern = /[\d+-/*^.,<>!]/; // Try to catch symbols that would make expression parsing fail.
    if (label == "") label = null;

    if (label && mathPattern.test(label)) {
      throw new Error(`Column label "${label}" contains invalid symbols!`);
    }

    if (label && LightsheetHelper.resolveColumnIndex(label) != -1) {
      throw new Error(
        `Column label "${label}" collides with default column names!`,
      );
    }

    let targetKey = this.columnPositions.get(columnIndex);
    if (!targetKey) {
      if (!label) return false; // Target column doesn't exist and label is being cleared.
      targetKey = this.initializeColumn(columnIndex);
    }
    const column = this.columns.get(targetKey)!;

    if (
      label &&
      this.columnLabels.has(label) &&
      this.columnLabels.get(label) != column.key
    ) {
      throw new Error(
        `Column label "${label}" is already assigned to another column!`,
      );
    }

    // Delete old entry.
    if (column.label) this.columnLabels.delete(column.label);

    const defaultLabel = LightsheetHelper.generateColumnLabel(columnIndex);
    this.updateColumnLabelReferences(
      column,
      column.label ?? defaultLabel,
      label ?? defaultLabel,
    );

    if (label == "" || label == null) {
      label = defaultLabel;
      column.label = undefined;
      this.deleteColumnIfUnused(column); // Clearing label may lead to the column being unused.
    } else if (label != defaultLabel) {
      column.label = label;
      this.columnLabels.set(label, column.key);
    }

    this.emitSetColumnLabelEvent(columnIndex, label);
    return true;
  }

  public getColumnLabel(columnIndex: number): string {
    const targetKey = this.columnPositions.get(columnIndex);
    const columnName = targetKey ? this.columns.get(targetKey)!.label : null;
    return columnName ?? LightsheetHelper.generateColumnLabel(columnIndex);
  }

  public getColumnByLabel(label: string): Column | null {
    let targetKey = this.columnLabels.get(label);
    if (!targetKey) {
      const position = LightsheetHelper.resolveColumnIndex(label);
      targetKey = this.columnPositions.get(position);
      if (!targetKey) return null;
    }

    return this.columns.get(targetKey) ?? null;
  }

  insertColumn(position: number): boolean {
    this.shiftCellGroups(
      position,
      ShiftDirection.forward,
      this.columns,
      this.columnPositions,
    );
    return true;
  }

  insertRow(position: number): boolean {
    this.shiftCellGroups(
      position,
      ShiftDirection.forward,
      this.rows,
      this.rowPositions,
    );
    return true;
  }

  deleteColumn(position: number): boolean {
    this.setColumnLabel(position, null);
    return this.deleteCellGroup(position, this.columns, this.columnPositions);
  }

  deleteRow(position: number): boolean {
    return this.deleteCellGroup(position, this.rows, this.rowPositions);
  }

  private deleteCellGroup(
    position: number,
    target: Map<ColumnKey | RowKey, CellGroup<ColumnKey | RowKey>>,
    targetPositions: Map<number, ColumnKey | RowKey>,
  ): boolean {
    const groupKey = targetPositions.get(position);
    const lastPosition = Math.max(...targetPositions.keys());

    if (groupKey !== undefined) {
      const group = target.get(groupKey)!;

      // Delete all cells in this group. This will also lead to the group being deleted as it'll be empty.
      for (const [oppositeKey] of group.cellIndex) {
        const keyPair = LightsheetHelper.getKeyPairFor(group, oppositeKey);
        this.deleteCell(keyPair.columnKey!, keyPair.rowKey!);
      }
    }

    if (position !== lastPosition) {
      this.shiftCellGroups(
        lastPosition,
        ShiftDirection.backward,
        target,
        targetPositions,
      );
    }
    return true;
  }

  private moveCellGroup(
    from: number,
    to: number,
    target: Map<ColumnKey | RowKey, CellGroup<ColumnKey | RowKey>>,
    targetPositions: Map<number, ColumnKey | RowKey>,
  ): boolean {
    if (from === to) return false;

    const groupKey = targetPositions.get(from);

    if (groupKey !== undefined) {
      const group = target.get(groupKey);
      if (group === undefined) {
        throw new Error(
          `CellGroup not found for key ${groupKey}, inconsistent state`,
        );
      }
      group.position = to;
      this.updateGroupPositionalRefs(group, from, to);
    }

    // We have to update all the positions to keep it consistent.
    targetPositions.delete(from);
    const direction =
      to > from ? ShiftDirection.backward : ShiftDirection.forward;
    this.shiftCellGroups(to, direction, target, targetPositions);

    if (groupKey === undefined) {
      targetPositions.delete(to);
    } else {
      targetPositions.set(to, groupKey);
    }

    return true;
  }

  /**
   * Shift a CellGroup (column or row) forward or backward till an empty position is found.
   */
  private shiftCellGroups(
    start: number,
    shiftDirection: ShiftDirection,
    target: Map<ColumnKey | RowKey, CellGroup<ColumnKey | RowKey>>,
    targetPositions: Map<number, ColumnKey | RowKey>,
  ): boolean {
    let previousValue: ColumnKey | RowKey | undefined = undefined;
    let currentPos = start;
    let tempCurrent;
    do {
      tempCurrent = targetPositions.get(currentPos);

      if (previousValue === undefined) {
        // First iteration; just clear the position we're shifting from.
        targetPositions.delete(currentPos);
      } else {
        targetPositions.set(currentPos, previousValue);
        const group = target.get(previousValue)!;
        this.updateGroupPositionalRefs(group, group.position, currentPos);
        group.position = currentPos;
      }

      previousValue = tempCurrent;

      if (shiftDirection === ShiftDirection.forward) {
        currentPos++;
      } else {
        if (currentPos === 0) {
          break;
        }
        currentPos--;
      }
    } while (tempCurrent !== undefined && previousValue !== undefined);

    return true;
  }

  private updateGroupPositionalRefs(
    group: CellGroup<ColumnKey | RowKey>,
    from: number,
    to: number,
  ) {
    if (group instanceof Column && group.label) {
      return; // References to named columns are unaffected.
    }

    for (const [oppositeKey, cellKey] of group!.cellIndex) {
      const cell = this.cellData.get(cellKey)!;
      if (!cell.referencesIn) continue; // Only cells with incoming references are affected.

      // Convert the information on how the group is shifted into coordinates.
      // from and to can refer to either a column or row position depending on the group type.
      const oppositeGroupPos =
        group instanceof Column
          ? this.rows.get(oppositeKey as RowKey)!.position
          : this.columns.get(oppositeKey as ColumnKey)!.position;

      const fromCoord: Coordinate = {
        column: group instanceof Column ? from : oppositeGroupPos!,
        row: group instanceof Row ? from : oppositeGroupPos!,
      };

      const toCoord: Coordinate = {
        column: group instanceof Column ? to : oppositeGroupPos!,
        row: group instanceof Row ? to : oppositeGroupPos!,
      };

      this.updateReferenceSymbols(cell, { from: fromCoord, to: toCoord });
    }
  }

  private updateColumnLabelReferences(
    column: Column,
    fromLabel: string,
    toLabel: string,
  ) {
    for (const [rowKey, cellKey] of column!.cellIndex) {
      const cell = this.cellData.get(cellKey)!;
      if (!cell.referencesIn) continue; // Only cells with incoming references are affected.

      const fromRow = this.getRowIndex(rowKey)!;
      const coordinate: Coordinate = { column: column.position, row: fromRow };
      this.updateReferenceSymbols(
        cell,
        { from: coordinate, to: coordinate }, // Position does not change.
        { from: fromLabel, to: toLabel },
      );
    }
  }

  private updateReferenceSymbols(
    cell: Cell,
    position: { from: Coordinate; to: Coordinate },
    label?: { from: string; to: string },
  ) {
    // Update reference symbols for all cell formulas that refer to the cell being moved.
    for (const [refCellKey, refInfo] of cell.referencesIn) {
      const refSheet = this.sheetHolder.getSheet(refInfo.sheetKey)!;
      const refCell = refSheet.cellData.get(refCellKey)!;

      const expr = new ExpressionHandler(refSheet, refInfo, refCell.rawValue);

      // Update either column label reference or positional reference depending on parameters.
      const newValue = label
        ? expr.updateReferenceSymbols(
            label.from,
            label.to,
            position.from.row,
            position.to.row,
          )
        : expr.updatePositionalReferences(position.from, position.to, this);

      // The formula may not change if the cell is being referenced indirectly through a range.
      if (refCell.rawValue === newValue) continue;
      refCell.rawValue = newValue;

      // Emit event for the rawValue change.
      refSheet.emitSetCellEvent(
        refSheet.getColumnIndex(refInfo.column)!,
        refSheet.getRowIndex(refInfo.row)!,
        refCell,
      );
    }
  }

  private deleteCell(colKey: ColumnKey, rowKey: RowKey): boolean {
    const col = this.columns.get(colKey);
    const row = this.rows.get(rowKey);
    if (!col || !row) return false;

    if (!col.cellIndex.has(row.key)) return false;
    const cellKey = col.cellIndex.get(row.key)!;
    const cell = this.cellData.get(cellKey)!;

    // Clear the cell's formula to clean up any references it may have.
    cell.rawValue = "";
    this.resolveCellFormula(cell, colKey, rowKey);

    // If other cells are referring to this cell, remove the reference and invalidate them.
    cell.referencesIn.forEach((cellRef, refCellKey) => {
      const refSheet = this.sheetHolder.getSheet(cellRef.sheetKey)!;
      const referredCell = refSheet.cellData.get(refCellKey)!;
      referredCell.referencesOut.delete(cellKey);
      referredCell.setState(CellState.INVALID_REFERENCE);
    });

    // Delete cell data and all references to it in its column and row.
    this.cellData.delete(cellKey);

    col.cellIndex.delete(row.key);
    col.cellFormatting.delete(row.key);
    row.cellIndex.delete(col.key);
    row.cellFormatting.delete(col.key);

    // Clean up empty columns and rows if needed.
    this.deleteColumnIfUnused(col);
    this.deleteRowIfUnused(row);

    return true;
  }

  getCellStyle(colKey?: ColumnKey, rowKey?: RowKey): CellStyle {
    const col = colKey ? this.columns.get(colKey) : null;
    const row = rowKey ? this.rows.get(rowKey) : null;
    if (!col && !row) return this.defaultStyle;

    const existingStyle = col && row ? col.cellFormatting.get(row.key) : null;
    const cellStyle = new CellStyle().clone(existingStyle);

    // Apply style properties with priority: cell style > column style > row style > default style.
    cellStyle
      .applyStylesOf(col ? col.defaultStyle : null)
      .applyStylesOf(row ? row.defaultStyle : null)
      .applyStylesOf(this.defaultStyle);

    return cellStyle;
  }

  setCellStyle(
    colKey: ColumnKey,
    rowKey: RowKey,
    style: CellStyle | null,
  ): boolean {
    const col = this.columns.get(colKey);
    const row = this.rows.get(rowKey);
    if (!col || !row) return false;

    if (style == null) {
      return this.clearCellStyle(colKey, rowKey);
    }

    // TODO Style could be non-null but empty; should we allow this?
    style = new CellStyle().clone(style);

    col.cellFormatting.set(row.key, style);
    row.cellFormatting.set(col.key, style);

    if (style.formatter) {
      this.applyCellFormatter(this.getCell(colKey, rowKey)!, colKey, rowKey);
    }

    return true;
  }

  setColumnStyle(colKey: ColumnKey, style: CellStyle | null): boolean {
    const col = this.columns.get(colKey);
    if (!col) return false;

    this.setCellGroupStyle(col, style);
    return true;
  }

  setRowStyle(rowKey: RowKey, style: CellStyle | null): boolean {
    const row = this.rows.get(rowKey);
    if (!row) return false;

    this.setCellGroupStyle(row, style);
    return true;
  }

  private setCellGroupStyle(
    group: CellGroup<ColumnKey | RowKey>,
    style: CellStyle | null,
  ) {
    style = style ? new CellStyle().clone(style) : null;
    const formatterChanged = style?.formatter != group.defaultStyle?.formatter;
    group.defaultStyle = style;

    // Iterate through formatted cells in this group and clear any styling properties set by the new style.
    for (const [opposingKey, cellStyle] of group.cellFormatting) {
      const shouldClear = cellStyle.clearStylingSetBy(style);
      if (!shouldClear) continue;

      // The cell's style will have no properties after applying this group's new style; clear it.
      const keyPair = LightsheetHelper.getKeyPairFor(group, opposingKey);
      this.clearCellStyle(keyPair.columnKey!, keyPair.rowKey!);
    }

    if (!formatterChanged) return;

    // Apply new formatter to all cells in this group.
    for (const [opposingKey] of group.cellIndex) {
      const cell = this.cellData.get(group.cellIndex.get(opposingKey)!)!;
      const keyPair = LightsheetHelper.getKeyPairFor(group, opposingKey);
      this.applyCellFormatter(cell, keyPair.columnKey!, keyPair.rowKey!);
    }
  }

  private clearCellStyle(colKey: ColumnKey, rowKey: RowKey): boolean {
    const col = this.columns.get(colKey);
    const row = this.rows.get(rowKey);
    if (!col || !row) return false;

    const style = col.cellFormatting.get(row.key);
    if (style?.formatter) {
      this.applyCellFormatter(this.getCell(colKey, rowKey)!, colKey, rowKey);
    }

    col.cellFormatting.delete(row.key);
    row.cellFormatting.delete(col.key);

    // Clearing a cell's style may leave it completely empty - delete if needed.
    this.deleteCellIfUnused(colKey, rowKey);

    return true;
  }

  // <Row index, <Column index, value>>
  exportData(): Map<number, Map<number, string>> {
    const data = new Map<number, Map<number, string>>();

    for (const [rowPos, rowKey] of this.rowPositions) {
      const rowData = new Map<number, string>();
      const row = this.rows.get(rowKey)!;

      // Use row's cell index to get keys for each cell and their corresponding columns.
      for (const [colKey, cellKey] of row.cellIndex) {
        const cell = this.cellData.get(cellKey)!;
        const column = this.columns.get(colKey)!;
        rowData.set(column.position, cell.formattedValue);
      }

      data.set(rowPos, new Map([...rowData.entries()].sort()));
    }
    return new Map([...data.entries()].sort());
  }

  private createCell(colKey: ColumnKey, rowKey: RowKey, value: string): Cell {
    const col = this.columns.get(colKey);
    const row = this.rows.get(rowKey);
    if (!col || !row) {
      throw new Error(
        `Failed to create cell at col: ${col} row: ${row}: Column or Row not found.`,
      );
    }

    if (col.cellIndex.has(row.key)) {
      throw new Error(
        `Failed to create cell at col: ${col} row: ${row}: Cell already exists.`,
      );
    }

    const cell = new Cell();
    cell.rawValue = value;
    this.cellData.set(cell.key, cell);
    this.resolveCell(cell, colKey, rowKey);

    col.cellIndex.set(row.key, cell.key);
    row.cellIndex.set(col.key, cell.key);

    return cell;
  }

  private getCell(colKey: ColumnKey, rowKey: RowKey): Cell | null {
    const col = this.columns.get(colKey);
    const row = this.rows.get(rowKey);

    if (!col || !row) return null;

    if (!col.cellIndex.has(row.key)) return null;
    const cellKey = col.cellIndex.get(row.key)!;
    return this.cellData.get(cellKey)!;
  }

  private resolveCell(cell: Cell, colKey: ColumnKey, rowKey: RowKey): boolean {
    const valueChanged = this.resolveCellFormula(cell, colKey, rowKey);
    if (valueChanged && cell.state == CellState.OK) {
      this.applyCellFormatter(cell, colKey, rowKey);
    }

    return valueChanged;
  }

  private resolveCellFormula(
    cell: Cell,
    colKey: ColumnKey,
    rowKey: RowKey,
  ): boolean {
    const expressionHandler = new ExpressionHandler(
      this,
      { sheetKey: this.key, column: colKey, row: rowKey },
      cell.rawValue,
    );

    const evalResult = expressionHandler.evaluate();
    const prevState = cell.state;
    const prevValue = cell.resolvedValue;

    this.processEvaluationReferences(cell, colKey, rowKey, evalResult);
    // After references are updated, check for circular references.
    if (evalResult?.references && this.hasCircularReference(cell)) {
      // TODO Should the whole reference chain be invalidated?
      cell.setState(CellState.CIRCULAR_REFERENCE);
      return prevValue != "";
    }

    const newState = evalResult?.result ?? CellState.INVALID_EXPRESSION;
    cell.resolvedValue = evalResult?.value ?? "";
    cell.setState(newState);

    // If the value of the cell hasn't changed (no change in state or value),
    // there's no need to update cells that reference this cell.
    if (
      prevState == newState &&
      prevValue == (evalResult ? evalResult.value : "")
    )
      return false;

    // Update cells that reference this cell.
    for (const [refKey, refInfo] of cell.referencesIn) {
      const referringSheet = this.sheetHolder.getSheet(refInfo.sheetKey)!;
      const referringCell = referringSheet.cellData.get(refKey)!;
      const refUpdated = referringSheet.resolveCell(
        referringCell,
        refInfo.column,
        refInfo.row,
      );

      // Emit event if the referred cell's value has changed (for the referring sheet's events).
      if (refUpdated) {
        referringSheet.emitSetCellEvent(
          referringSheet.getColumnIndex(refInfo.column)!,
          referringSheet.getRowIndex(refInfo.row)!,
          referringCell,
        );
      }
    }
    return true;
  }

  private applyCellFormatter(
    cell: Cell,
    colKey: ColumnKey,
    rowKey: RowKey,
  ): boolean {
    const style = this.getCellStyle(colKey, rowKey);
    let formattedValue: string | null = cell.resolvedValue;
    if (style?.formatter) {
      formattedValue = style.formatter.format(formattedValue);
      if (formattedValue == null) {
        cell.setState(CellState.INVALID_FORMAT);
        return false;
      }
    }

    cell.formattedValue = formattedValue;
    return true;
  }

  /**
   * Delete a cell if it's empty, has no formatting and is not referenced by any other cell.
   */
  private deleteCellIfUnused(colKey: ColumnKey, rowKey: RowKey): boolean {
    const cell = this.getCell(colKey, rowKey)!;
    if (cell.rawValue != "") return false;

    // Check if this cell is referenced by anything.
    if (cell.referencesIn.size > 0) return false;

    // Check if cell-specific formatting is set.
    const cellCol = this.columns.get(colKey)!;
    if (cellCol.cellFormatting.has(rowKey)) return false;

    return this.deleteCell(colKey, rowKey);
  }

  private deleteColumnIfUnused(column: Column): boolean {
    if (!column.isUnused()) return false;

    this.columns.delete(column.key);
    this.columnPositions.delete(column.position);
    return true;
  }

  private deleteRowIfUnused(row: Row): boolean {
    if (!row.isUnused()) return false;

    this.rows.delete(row.key);
    this.rowPositions.delete(row.position);
    return true;
  }

  /**
   * Update reference collections of cells for all cells whose values are affected.
   */
  private processEvaluationReferences(
    cell: Cell,
    columnKey: ColumnKey,
    rowKey: RowKey,
    evalResult: EvaluationResult | null,
  ) {
    // Update referencesOut of this cell and referencesIn of newly referenced cells.
    const oldOut = new Map<CellKey, CellReference>(cell.referencesOut);
    cell.referencesOut.clear();

    evalResult?.references.forEach((ref) => {
      const refSheet = this.sheetHolder.getSheet(ref.sheetKey)!;

      // Initialize the referred cell if it doesn't exist yet.
      const position = refSheet.initializePosition(
        ref.position.column,
        ref.position.row,
      );
      if (!refSheet.getCellInfoAt(ref.position.column, ref.position.row)) {
        refSheet.createCell(position.columnKey!, position.rowKey!, "");
      }

      const referredCell = refSheet.getCell(
        position.columnKey!,
        position.rowKey!,
      )!;

      // Add referred cells to this cell's referencesOut.
      cell.referencesOut.set(referredCell.key, {
        sheetKey: refSheet.key,
        column: position.columnKey!,
        row: position.rowKey!,
      });

      // Add this cell to the referred cell's referencesIn.
      referredCell.referencesIn.set(cell.key, {
        sheetKey: this.key,
        column: columnKey,
        row: rowKey,
      });
    });

    // Resolve (oldOut - newOut) to get references that were removed from the formula.
    const removedReferences = new Map<CellKey, CellReference>(
      [...oldOut].filter(([cellKey]) => !cell.referencesOut.has(cellKey)),
    );

    // Clean up the referencesIn sets of cells that are no longer referenced by the formula.
    for (const [cellKey, refInfo] of removedReferences) {
      const referredSheet = this.sheetHolder.getSheet(refInfo.sheetKey)!;
      const referredCell = referredSheet.cellData.get(cellKey)!;
      referredCell.referencesIn.delete(cell.key);

      // This may result in the cell being unused - delete if necessary.
      referredSheet.deleteCellIfUnused(refInfo.column, refInfo.row);
    }
  }

  private hasCircularReference(cell: Cell): boolean {
    const refStack: [CellKey, SheetKey][] = [[cell.key, this.key]];
    let initial = true;

    // Depth-first search.
    while (refStack.length > 0) {
      const current = refStack.pop()!;
      const currCellKey = current[0];
      if (currCellKey === cell.key && !initial) return true;

      const currSheet = this.sheetHolder.getSheet(current[1])!;
      const currentCell = currSheet.cellData.get(currCellKey)!;
      if (currentCell.state == CellState.CIRCULAR_REFERENCE && !initial)
        return true;

      currentCell.referencesOut.forEach((cellRef, refCellKey) => {
        refStack.push([refCellKey, cellRef.sheetKey]);
      });

      initial = false;
    }
    return false;
  }

  private initializePosition(colPos: number, rowPos: number): PositionInfo {
    let rowKey;

    // Create row and column if they don't exist yet.
    if (!this.rowPositions.has(rowPos)) {
      const row = new Row(this.defaultHeight, rowPos);
      this.rows.set(row.key, row);
      this.rowPositions.set(rowPos, row.key);

      rowKey = row.key;
    } else {
      rowKey = this.rowPositions.get(rowPos)!;
    }

    const colKey = this.initializeColumn(colPos);

    return { rowKey: rowKey, columnKey: colKey };
  }

  private initializeColumn(position: number): ColumnKey {
    if (!this.columnPositions.has(position)) {
      const col = new Column(this.defaultWidth, position);
      this.columns.set(col.key, col);
      this.columnPositions.set(position, col.key);

      return col.key;
    } else {
      return this.columnPositions.get(position)!;
    }
  }

  private emitSetCellEvent(colPos: number, rowPos: number, cell: Cell | null) {
    const payload: CoreSetCellPayload = {
      position: {
        column: colPos,
        row: rowPos,
      },
      rawValue: cell ? cell.rawValue : "",
      formattedValue: cell ? cell.formattedValue : "",
      state: cell ? cell.state : CellState.OK,
      clearCell: cell == null,
      clearRow: this.rowPositions.get(rowPos) == null,
    };

    this.events.emit(new LightsheetEvent(EventType.CORE_SET_CELL, payload));
  }

  private emitSetColumnLabelEvent(columnIndex: number, label: string) {
    const payload: SetColumnLabelPayload = {
      columnIndex: columnIndex,
      label: label,
    };

    this.events.emit(
      new LightsheetEvent(EventType.CORE_SET_COLUMN_LABEL, payload),
    );
  }

  private registerEvents() {
    this.events.on(EventType.UI_SET_CELL, (event) =>
      this.handleUISetCell(event),
    );

    this.events.on(EventType.UI_SET_COLUMN_LABEL, (event) =>
      this.handleSetColumnLabel(event),
    );
  }

  private handleUISetCell(event: LightsheetEvent) {
    const payload = event.payload as UISetCellPayload;
    this.setCellAt(
      payload.position.column,
      payload.position.row,
      payload.rawValue,
    );
  }

  private handleSetColumnLabel(event: LightsheetEvent) {
    const payload = event.payload as SetColumnLabelPayload;

    this.setColumnLabel(payload.columnIndex, payload.label);
  }
}

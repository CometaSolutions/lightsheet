import LightsheetEvent from "../core/event/event.ts";
import {
  CoreSetCellPayload,
  SetColumnLabelPayload,
  UISetCellPayload,
} from "../core/event/events.types.ts";
import EventType from "../core/event/eventType.ts";
import { LightsheetOptions, ToolbarOptions } from "../main.types";
import LightsheetHelper from "../utils/helpers.ts";
import { ToolbarItems } from "../utils/constants.ts";
import { Coordinate } from "../utils/common.types.ts";
import Events from "../core/event/events.ts";
import { CellState } from "../core/structure/cell/cellState.ts";

export default class UI {
  tableEl!: Element;
  toolbarDom: HTMLElement | undefined;
  formulaBarDom!: HTMLElement | null;
  formulaInput!: HTMLInputElement;
  selectedCellDisplay!: HTMLElement;
  tableHeadDom!: Element;
  tableBodyDom!: Element;
  selectedCell: HTMLElement | null = null;
  rangeSelectionEnd: HTMLElement | null = null;
  selectedRowNumberCell: HTMLElement | null = null;
  selectedHeaderCell: HTMLElement | null = null;
  toolbarOptions: ToolbarOptions;
  isReadOnly: boolean;
  sheetName: string;
  tableContainerDom: Element;
  private events: Events;

  constructor(
    lightSheetContainerDom: Element,
    lightSheetOptions: LightsheetOptions,
    events: Events | null = null,
  ) {
    this.events = events ?? new Events();
    this.registerEvents();
    this.initializeKeyEvents();

    this.toolbarOptions = {
      showToolbar: false,
      element: undefined,
      items: ToolbarItems,
      ...lightSheetOptions.toolbarOptions,
    };
    this.isReadOnly = lightSheetOptions.isReadOnly || false;
    this.sheetName = lightSheetOptions.sheetName;

    this.tableContainerDom = lightSheetContainerDom;
    lightSheetContainerDom.classList.add("lightsheet_table_container");

    //table
    this.createTableDom();

    //toolbar
    this.createToolbar();

    //formula bar
    this.createFormulaBar();
  }

  private createTableDom() {
    const tableDom = document.createElement("table");
    tableDom.classList.add("lightsheet_table");
    tableDom.setAttribute("cellpadding", "0");
    tableDom.setAttribute("cellspacing", "0");
    tableDom.setAttribute("unselectable", "yes");
    this.tableContainerDom.appendChild(tableDom);
    this.tableEl = tableDom;

    //thead
    this.tableHeadDom = document.createElement("thead");
    tableDom.appendChild(this.tableHeadDom);
    const tableHeadRowDom = document.createElement("tr");
    this.tableHeadDom.appendChild(tableHeadRowDom);
    const rowNumberCell = document.createElement("th");
    rowNumberCell.classList.add(
      "lightsheet_table_row_number",
      "lightsheet_table_td",
    );
    rowNumberCell.textContent = " ";
    tableHeadRowDom.appendChild(rowNumberCell);

    //tbody
    this.tableBodyDom = document.createElement("tbody");
    tableDom.appendChild(this.tableBodyDom);
  }

  private createToolbar() {
    if (
      !this.toolbarOptions.showToolbar ||
      this.toolbarOptions.items?.length == 0
    )
      return;
    //Element
    this.toolbarDom = document.createElement("div");
    this.toolbarDom.classList.add("lightsheet_table_toolbar");

    if (this.toolbarOptions.element != null) {
      this.toolbarOptions.element.appendChild(this.toolbarDom);
    } else {
      this.tableEl.parentNode!.insertBefore(this.toolbarDom, this.tableEl);
    }

    for (let i = 0; i < this.toolbarOptions.items!.length; i++) {
      const toolbarItem = document.createElement("i");
      toolbarItem.classList.add("lightSheet_toolbar_item");
      toolbarItem.classList.add("material-symbols-outlined");
      toolbarItem.textContent = this.toolbarOptions.items![i];
      this.toolbarDom.appendChild(toolbarItem);
    }
  }

  private initializeKeyEvents() {
    document.addEventListener("keydown", (e: KeyboardEvent) => {
      const targetEl = e.target as HTMLElement;
      const inputSelected = targetEl.classList.contains(
        "lightsheet_table_cell_input",
      );

      if (
        !this.selectedCell &&
        !this.selectedHeaderCell &&
        !this.selectedRowNumberCell
      )
        return; // Ignore keyboard events if nothing is selected.

      const selectedInput = this.getSelectedCellInput();

      if (e.key === "Escape") {
        // Clear selection on esc.
        if (e.target == selectedInput) {
          selectedInput!.blur();
          return;
        } else if (e.target == this.formulaInput) {
          this.formulaInput.blur();
          return;
        }
        this.removeGroupSelection();
        this.removeCellRangeSelection();
        this.removeCellSelection();
      } else if (e.key.startsWith("Arrow") && !inputSelected) {
        e.preventDefault();
        const currPos = this.getCellCoordinate(this.selectedCell)!;
        const moveMap: Map<string, number[]> = new Map([
          ["ArrowLeft", [-1, 0]],
          ["ArrowRight", [1, 0]],
          ["ArrowUp", [0, -1]],
          ["ArrowDown", [0, 1]],
        ]);
        const shift = moveMap.get(e.key)!;
        const newCell = document.getElementById(
          this.getIndexedCellId(
            shift[0] + currPos.column,
            shift[1] + currPos.row,
          ),
        );
        if (!newCell) return;
        this.setSelectedCell(newCell);
      } else if (e.key !== "Enter" && e.target != this.formulaInput) {
        // Default to appending text to selected cell. Note that in this case Excel also clears the cell.
        selectedInput?.focus();
      }
    });
  }

  removeToolbar() {
    if (this.toolbarDom) this.toolbarDom.remove();
  }

  showToolbar(isShown: boolean) {
    this.toolbarOptions.showToolbar = isShown;
    if (isShown) {
      this.createToolbar();
    } else {
      this.removeToolbar();
    }
  }

  createFormulaBar() {
    if (this.isReadOnly || this.formulaBarDom instanceof HTMLElement) {
      return;
    }
    this.formulaBarDom = document.createElement("div");
    this.formulaBarDom.classList.add("lightsheet_table_formula_bar");
    this.tableContainerDom.insertBefore(this.formulaBarDom, this.tableEl);
    //selected cell display element
    this.selectedCellDisplay = document.createElement("div");
    this.selectedCellDisplay.classList.add("lightsheet_selected_cell_display");
    this.formulaBarDom.appendChild(this.selectedCellDisplay);

    //"fx" label element
    const fxLabel = document.createElement("div");
    fxLabel.textContent = "fx";
    fxLabel.classList.add("lightsheet_fx_label");
    this.formulaBarDom.appendChild(fxLabel);

    //formula input
    this.formulaInput = document.createElement("input");
    this.formulaInput.classList.add("lightsheet_formula_input");
    this.formulaBarDom.appendChild(this.formulaInput);
    this.setFormulaBar();
  }

  setFormulaBar() {
    this.formulaInput.addEventListener("input", () => {
      const newValue = this.formulaInput.value;
      const selectedCellInput = this.getSelectedCellInput();
      if (selectedCellInput) {
        selectedCellInput.value = newValue;
      }
    });
    this.formulaInput.addEventListener("keyup", (event) => {
      const newValue = this.formulaInput.value;
      if (!this.selectedCell) return;

      if (event.key === "Enter") {
        const selected = this.getSelectedCellCoordinate();
        if (selected) {
          this.onUICellValueChange(newValue, selected.column, selected.row);
        }
        this.formulaInput.blur();
        this.formulaInput.value = "";
        this.removeCellSelection();
      }
    });

    this.formulaInput.addEventListener("mousedown", (e) => {
      if (!this.selectedCell) e.preventDefault();
    });

    this.formulaInput.onblur = () => {
      const newValue = this.formulaInput.value;
      const selected = this.getSelectedCellCoordinate();
      if (selected) {
        this.onUICellValueChange(newValue, selected.column, selected.row);
      }
    };
  }

  removeFormulaBar() {
    if (this.formulaBarDom) {
      this.formulaBarDom.remove();
      this.formulaBarDom = null;
    }
  }

  addColumn() {
    const headerCellDom = document.createElement("th");
    headerCellDom.classList.add(
      "lightsheet_table_header",
      "lightsheet_table_td",
    );

    const columnIndex = this.getColumnCount();
    const newHeaderValue = LightsheetHelper.generateColumnLabel(columnIndex);

    headerCellDom.textContent = newHeaderValue;
    headerCellDom.onclick = (e: MouseEvent) =>
      this.onClickHeaderCell(e, columnIndex);

    const rowCount = this.getRowCount();
    for (let rowIndex = 0; rowIndex < rowCount; rowIndex++) {
      const rowDom = this.tableBodyDom.children[rowIndex];
      this.addCell(rowDom, columnIndex, rowIndex);
    }

    this.tableHeadDom.children[0].appendChild(headerCellDom);
  }

  private onClickHeaderCell(e: MouseEvent, columnIndex: number) {
    const selectedColumn = e.target as HTMLElement;
    if (!selectedColumn) return;
    const prevSelection = this.selectedHeaderCell;
    this.removeGroupSelection();
    this.removeCellRangeSelection();

    if (prevSelection !== selectedColumn) {
      selectedColumn.classList.add(
        "lightsheet_table_selected_row_number_header_cell",
      );
      this.selectedHeaderCell = selectedColumn;
      Array.from(this.tableBodyDom.children).forEach((childElement) => {
        childElement.children[columnIndex + 1].classList.add(
          "lightsheet_table_selected_row_column",
        );
      });
    }
  }

  addRow(): HTMLElement {
    const rowCount = this.getRowCount();
    const rowDom = this.createRowElement(rowCount);
    this.tableBodyDom.appendChild(rowDom);

    const columnCount = this.getColumnCount();
    for (let columnIndex = 0; columnIndex < columnCount; columnIndex++) {
      this.addCell(rowDom, columnIndex, rowCount);
    }
    return rowDom;
  }

  private createRowElement(labelCount: number): HTMLElement {
    const rowDom = document.createElement("tr");
    rowDom.id = this.getIndexedRowId(labelCount);
    const rowNumberCell = document.createElement("td");
    rowNumberCell.innerHTML = `${labelCount + 1}`; // Row numbers start from 1
    rowNumberCell.classList.add(
      "lightsheet_table_row_number",
      "lightsheet_table_row_cell",
      "lightsheet_table_td",
    );
    rowDom.appendChild(rowNumberCell);
    rowNumberCell.onclick = (e: MouseEvent) => {
      const selectedRow = e.target as HTMLElement;
      if (!selectedRow) return;
      const prevSelection = this.selectedRowNumberCell;
      this.removeGroupSelection();
      this.removeCellRangeSelection();

      if (prevSelection !== selectedRow) {
        selectedRow.classList.add(
          "lightsheet_table_selected_row_number_header_cell",
        );
        this.selectedRowNumberCell = selectedRow;

        const parentElement = selectedRow.parentElement;
        if (parentElement) {
          for (let i = 1; i < parentElement.children.length; i++) {
            const childElement = parentElement.children[i];
            if (childElement !== selectedRow) {
              childElement.classList.add(
                "lightsheet_table_selected_row_column",
              );
            }
          }
        }
      }
    };
    return rowDom;
  }

  getRow(rowKey: string): HTMLElement | null {
    return document.getElementById(rowKey);
  }

  addCell(rowDom: Element, colIndex: number, rowIndex: number): HTMLElement {
    const cellDom = document.createElement("td");
    cellDom.classList.add(
      "lightsheet_table_cell",
      "lightsheet_table_row_cell",
      "lightsheet_table_td",
    );
    rowDom.appendChild(cellDom);
    cellDom.id = this.getIndexedCellId(colIndex, rowIndex);

    const inputDom = document.createElement("input");
    inputDom.classList.add("lightsheet_table_cell_input");
    inputDom.readOnly = this.isReadOnly;
    cellDom.appendChild(inputDom);

    inputDom.addEventListener("input", (e: Event) => {
      const newValue = (e.target as HTMLInputElement).value;
      if (this.formulaBarDom) this.formulaInput.value = newValue;
    });

    inputDom.onchange = (e: Event) =>
      this.onUICellValueChange(
        (e.target as HTMLInputElement).value,
        colIndex,
        rowIndex,
      );

    const selectCell = (e: Event) => {
      if (this.selectedCell != cellDom || this.rangeSelectionEnd)
        e.preventDefault();

      this.removeGroupSelection();
      this.removeCellRangeSelection();

      if (this.selectedCell == cellDom) {
        return;
      }

      this.setSelectedCell(cellDom);
    };

    cellDom.addEventListener("mousedown", selectCell);
    inputDom.onfocus = selectCell;

    inputDom.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        this.removeCellSelection();
      }
    });

    inputDom.onmouseover = (e: MouseEvent) => {
      if (e.buttons === 1) {
        this.handleMouseOver(e, colIndex, rowIndex);
      }
    };

    return cellDom;
  }

  setReadOnly(readonly: boolean) {
    const inputElements = this.tableBodyDom.querySelectorAll("input");
    inputElements.forEach((input) => {
      (input as HTMLInputElement).readOnly = readonly;
    });
    this.isReadOnly = readonly;
    if (readonly) {
      this.removeFormulaBar();
    } else {
      this.createFormulaBar();
    }
  }

  onUICellValueChange(rawValue: string, colIndex: number, rowIndex: number) {
    const payload: UISetCellPayload = {
      position: { column: colIndex, row: rowIndex },
      rawValue,
    };
    this.events.emit(new LightsheetEvent(EventType.UI_SET_CELL, payload));
  }

  private registerEvents() {
    this.events.on(EventType.CORE_SET_CELL, (event) => {
      this.onCoreSetCell(event);
    });

    this.events.on(EventType.CORE_SET_COLUMN_LABEL, (event) => {
      this.onSetColumnLabel(event);
    });
  }

  private onCoreSetCell(event: LightsheetEvent) {
    const payload = event.payload as CoreSetCellPayload;
    // Create new columns if the column index is greater than the current column count.
    const newColumns = payload.position.column - this.getColumnCount() + 1;
    for (let i = 0; i < newColumns; i++) {
      this.addColumn();
    }

    const newRows = payload.position.row - this.getRowCount() + 1;
    for (let i = 0; i < newRows; i++) {
      this.addRow();
    }

    const cellDom = document.getElementById(
      this.getIndexedCellId(payload.position.column, payload.position.row),
    )!;

    // Update input element with values from the core.
    const inputEl = cellDom.firstChild! as HTMLInputElement;
    inputEl.setAttribute("rawValue", payload.rawValue);
    inputEl.setAttribute("resolvedValue", payload.formattedValue);

    if (payload.state == CellState.OK) {
      inputEl.removeAttribute("state");
    } else {
      inputEl.setAttribute("state", Number(payload.state).toString());
    }

    this.updateCellDisplay(inputEl);
    if (inputEl == this.getSelectedCellInput() && this.formulaBarDom) {
      this.formulaInput.value = payload.rawValue;
    }
  }

  private onSetColumnLabel(event: LightsheetEvent) {
    const payload = event.payload as SetColumnLabelPayload;
    const colCount = this.getColumnCount();
    for (let i = 0; i < payload.columnIndex - colCount + 1; i++) {
      this.addColumn();
    }

    const labelElement =
      this.tableHeadDom.children[0].children[payload.columnIndex + 1];
    labelElement.textContent = payload.label;
  }

  getColumnCount() {
    return this.tableHeadDom.children[0].children.length - 1;
  }

  getRowCount() {
    return this.tableBodyDom.children.length;
  }

  removeGroupSelection() {
    this.removeColumnSelection();
    this.removeRowSelection();
  }

  removeRowSelection() {
    if (!this.selectedRowNumberCell) return;

    this.selectedRowNumberCell.classList.remove(
      "lightsheet_table_selected_row_number_header_cell",
    );
    const parentElement = this.selectedRowNumberCell.parentElement;
    if (parentElement) {
      for (let i = 1; i < parentElement.children.length; i++) {
        const childElement = parentElement.children[i];
        if (childElement !== this.selectedRowNumberCell) {
          childElement.classList.remove("lightsheet_table_selected_row_column");
        }
      }
    }

    this.selectedRowNumberCell = null;
  }

  removeColumnSelection() {
    if (!this.selectedHeaderCell) return;

    this.selectedHeaderCell.classList.remove(
      "lightsheet_table_selected_row_number_header_cell",
    );
    for (const childElement of this.tableBodyDom.children) {
      Array.from(childElement.children).forEach((element: Element) => {
        element.classList.remove("lightsheet_table_selected_row_column");
      });
    }

    this.selectedHeaderCell = null;
  }

  removeCellRangeSelection() {
    if (!this.rangeSelectionEnd || !this.selectedCell) return;

    this.updateSelection(
      this.getCellCoordinate(this.rangeSelectionEnd)!,
      this.getSelectedCellCoordinate()!,
    );
    this.rangeSelectionEnd = null;
  }

  private updateSelection(previousEnd: Coordinate, newEnd: Coordinate) {
    const selRoot = this.getSelectedCellCoordinate()!;

    // If the selection root is within the rectangle formed by the update points,
    // clear the previous selection first.
    const topLeft = {
      column: Math.min(previousEnd.column, newEnd!.column),
      row: Math.min(previousEnd.row, newEnd!.row),
    };
    const botRight = {
      column: Math.max(previousEnd.column, newEnd!.column),
      row: Math.max(previousEnd.row, newEnd!.row),
    };
    if (
      (selRoot.column > topLeft.column && selRoot.column < botRight.column) ||
      (selRoot.row > topLeft.row && selRoot.row < botRight.row)
    ) {
      this.updateSelection(previousEnd, selRoot);
      previousEnd = selRoot;
    }

    const selectionUpdate = this.getSelectionDiff(previousEnd, newEnd);

    // TODO updating each cell class can cause performance issues in larger sheets.
    for (const added of selectionUpdate.added) {
      const cell = document.getElementById(
        this.getIndexedCellId(added.column, added.row),
      );
      cell?.classList.add("lightsheet_table_selected_cell_range");
    }

    for (const removed of selectionUpdate.removed) {
      const cell = document.getElementById(
        this.getIndexedCellId(removed.column, removed.row),
      );
      cell?.classList.remove("lightsheet_table_selected_cell_range");
    }
  }

  private getSelectionDiff(
    previousEnd: Coordinate,
    newEnd: Coordinate,
  ): { added: Array<Coordinate>; removed: Array<Coordinate> } {
    const selRoot = this.getSelectedCellCoordinate()!;

    const added: Array<Coordinate> = [];
    const removed: Array<Coordinate> = [];

    // Negative when shrinking, positive when growing.
    const colDiff =
      Math.abs(selRoot.column - newEnd.column) -
      Math.abs(selRoot.column - previousEnd.column);

    const diffMinColumn = Math.min(previousEnd.column, newEnd.column);
    const newSelectionMinRow = Math.min(selRoot.row, newEnd.row);
    const newSelectionMaxRow = Math.max(selRoot.row, newEnd.row);

    // Add cells from horizontal expansion.
    for (let c = diffMinColumn; c < diffMinColumn + colDiff; c++) {
      for (let r = newSelectionMinRow; r < newSelectionMaxRow + 1; r++) {
        const colIndex = newEnd.column - previousEnd.column > 0 ? c + 1 : c;
        added.push({ column: colIndex, row: r });
      }
    }

    const rowDiff =
      Math.abs(selRoot.row - newEnd.row) -
      Math.abs(selRoot.row - previousEnd.row);

    const diffMinRow = Math.min(previousEnd.row, newEnd.row);
    const newSelectionMinCol = Math.min(selRoot.column, newEnd.column);
    const newSelectionMaxCol = Math.max(selRoot.column, newEnd.column);

    // Add cells from vertical expansion.
    for (let r = diffMinRow; r < diffMinRow + rowDiff; r++) {
      for (let c = newSelectionMinCol; c < newSelectionMaxCol + 1; c++) {
        const rowIndex = newEnd.row - previousEnd.row > 0 ? r + 1 : r;
        added.push({ column: c, row: rowIndex });
      }
    }

    const diffMaxCol = Math.max(previousEnd.column, newEnd.column);
    const minAffectedRow = Math.min(previousEnd.row, newEnd.row, selRoot.row);
    const maxAffectedRow = Math.max(previousEnd.row, newEnd.row, selRoot.row);

    // Remove cells from horizontal reduction.
    for (let c = diffMaxCol; c > diffMaxCol + colDiff; c--) {
      for (let r = minAffectedRow; r < maxAffectedRow + 1; r++) {
        const colIndex = newEnd.column - previousEnd.column > 0 ? c - 1 : c;
        removed.push({ column: colIndex, row: r });
      }
    }

    const minAffectedCol = Math.min(
      previousEnd.column,
      newEnd.column,
      selRoot.column,
    );
    const maxAffectedCol = Math.max(
      previousEnd.column,
      newEnd.column,
      selRoot.column,
    );
    const diffMaxRow = Math.max(previousEnd.row, newEnd.row);

    // Remove cells from vertical reduction.
    for (let r = diffMaxRow; r > diffMaxRow + rowDiff; r--) {
      for (let c = minAffectedCol; c < maxAffectedCol + 1; c++) {
        const rowIndex = newEnd.row - previousEnd.row > 0 ? r - 1 : r;
        removed.push({ column: c, row: rowIndex });
      }
    }

    // Results will contain duplicate points where horizontal and vertical movement coincide,
    // but this does not affect current usage.
    return { added: added, removed: removed };
  }

  handleMouseOver(e: MouseEvent, colIndex: number, rowIndex: number) {
    if (e.button !== 0 || !this.selectedCell) return;

    const oldEnd =
      this.getCellCoordinate(this.rangeSelectionEnd) ??
      this.getSelectedCellCoordinate()!;

    const newEnd = { column: colIndex, row: rowIndex };
    this.updateSelection(oldEnd, newEnd);

    if (!this.rangeSelectionEnd) {
      this.updateCellDisplay(this.getSelectedCellInput()!); // Selected cell will hide formula when range selection is started.
    }

    this.rangeSelectionEnd = document.getElementById(
      this.getIndexedCellId(colIndex, rowIndex),
    );
  }

  private getSelectedCellCoordinate(): Coordinate | null {
    const selection = this.selectedCell;
    if (!selection) return null;

    return this.getCellCoordinate(selection);
  }

  private getCellCoordinate(
    cellElement: HTMLElement | null,
  ): Coordinate | null {
    if (!cellElement) return null;

    const cellCoordinate = cellElement.id.split("_").slice(-2);
    return {
      column: Number(cellCoordinate[0]),
      row: Number(cellCoordinate[1]),
    };
  }

  private updateCellDisplay(cellInput: HTMLInputElement) {
    // Display formula for selected cell.
    if (this.getSelectedCellInput() === cellInput && !this.rangeSelectionEnd) {
      cellInput.value = cellInput.getAttribute("rawvalue") ?? "";
      return;
    }

    // Display error state for unselected cell.
    const cellState = cellInput.getAttribute("state");
    if (cellState) {
      const errorInfo = UI.getErrorForState(Number(cellState));
      cellInput.value = `#${errorInfo.code}`;
      cellInput.title = errorInfo.description;
      return;
    }

    // Display resolved value for unselected cell without error state.
    cellInput.removeAttribute("title");
    cellInput.value = cellInput.getAttribute("resolvedvalue") ?? "";
  }

  private setSelectedCell(cell: HTMLElement) {
    this.removeCellSelection();
    this.selectedCell = cell;
    cell.classList.add("lightsheet_table_selected_cell");

    const selInput = this.getSelectedCellInput()!;
    this.updateCellDisplay(selInput);
    if (this.formulaBarDom) {
      this.formulaInput.value = selInput.getAttribute("rawValue")!;
      const selectedPos = this.getSelectedCellCoordinate()!;
      this.selectedCellDisplay.textContent = `${this.getColumnHeader(selectedPos.column)!.textContent}${selectedPos.row + 1}`;
      this.formulaInput.blur();
    }
  }

  private removeCellSelection() {
    const selection = this.getSelectedCellInput();
    if (!selection) return;
    selection.blur();
    selection.parentElement!.classList.remove("lightsheet_table_selected_cell");
    this.selectedCell = null;
    this.updateCellDisplay(selection);
    if (this.formulaBarDom) this.formulaInput.value = "";
  }

  private getSelectedCellInput(): HTMLInputElement | null {
    return this.selectedCell?.firstChild as HTMLInputElement;
  }

  private getIndexedRowId(rowIndex: number) {
    return `${this.sheetName}_row_${rowIndex}`;
  }

  private getIndexedCellId(colIndex: number, rowIndex: number) {
    return `${this.sheetName}_${colIndex}_${rowIndex}`;
  }

  private getColumnHeader(index: number): Element | null {
    return this.tableHeadDom.children[0].children[index + 1];
  }

  // TODO Consider locale?
  private static getErrorForState(state: CellState): {
    code: string;
    description: string;
  } {
    switch (state) {
      case CellState.INVALID_EXPRESSION:
        return {
          code: "SYNTAX",
          description: "The expression is syntactically invalid.",
        };
      case CellState.INVALID_SYMBOL:
        return {
          code: "SYMBOL",
          description: "The expression contains an invalid symbol.",
        };
      case CellState.INVALID_OPERANDS:
        return {
          code: "OPERANDS",
          description: "The expression contains invalid operands.",
        };
      case CellState.INVALID_REFERENCE:
        return {
          code: "REFERENCE",
          description:
            "The formula contains a reference that couldn't be resolved.",
        };
      case CellState.CIRCULAR_REFERENCE:
        return {
          code: "CIRC.REF",
          description: "The formula creates a circular reference.",
        };
      case CellState.INVALID_FORMAT:
        return {
          code: "FORMAT",
          description: "The formatter is incompatible with the cell contents.",
        };
      case CellState.UNKNOWN_ERROR:
        return {
          code: "ERROR",
          description:
            "Evaluating the formula failed. See the console for details.",
        };
      default:
        return {
          code: "",
          description: "",
        };
    }
  }
}

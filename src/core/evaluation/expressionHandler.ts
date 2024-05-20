import {
  create,
  parseDependencies,
  addDependencies,
  subtractDependencies,
  multiplyDependencies,
  divideDependencies,
  SymbolNodeDependencies,
  FunctionNodeDependencies,
  sumDependencies,
  MathNode,
} from "mathjs/number";

import Sheet from "../structure/sheet.ts";
import {
  CellSheetPosition,
  EvaluationResult,
} from "./expressionHandler.types.ts";

import { Coordinate } from "../../utils/common.types.ts";
import { CellReference } from "../structure/cell/types.cell.ts";
import SheetHolder from "../structure/sheetHolder.ts";
import { CellState } from "../structure/cell/cellState.ts";
import LightsheetHelper from "../../utils/helpers.ts";

const math = create({
  parseDependencies,
  addDependencies,
  subtractDependencies,
  multiplyDependencies,
  divideDependencies,
  SymbolNodeDependencies,
  FunctionNodeDependencies,
  sumDependencies,
});

export default class ExpressionHandler {
  private static functions: Map<
    string,
    (cellRef: CellReference, ...args: any[]) => string
  > = new Map();

  private sheet: Sheet;

  private cellRefHolder: Array<CellSheetPosition>;
  private rawValue: string;
  private targetCellRef: CellReference;
  private evaluationResult: CellState;

  constructor(targetSheet: Sheet, targetCell: CellReference, rawValue: string) {
    this.sheet = targetSheet;
    this.targetCellRef = targetCell;
    this.evaluationResult = CellState.OK;

    this.rawValue = rawValue;
    this.cellRefHolder = [];

    math.FunctionNode.onUndefinedFunction = (name: string) =>
      this.resolveFunction(name);

    //Configure mathjs to allow colons in symbols.
    const isAlphaOriginal = math.parse.isAlpha;
    math.parse.isAlpha = (c: string, cPrev: string, cNext: string): boolean => {
      return isAlphaOriginal(c, cPrev, cNext) || c == ":" || c == "!";
    };

    math.SymbolNode.onUndefinedSymbol = (name: string) =>
      this.resolveSymbol(name);
  }

  evaluate(): EvaluationResult | null {
    if (!this.rawValue.startsWith("="))
      return { value: this.rawValue, references: [], result: CellState.OK };

    const expression = this.rawValue.substring(1);
    let parsed;
    try {
      parsed = math.parse(expression);
      parsed.transform((node) =>
        this.injectFunctionParameter(node, this.targetCellRef),
      );
    } catch (e) {
      return null; // Expression syntax is invalid.
    }

    try {
      const value = parsed!.evaluate().toString();
      return {
        value: value,
        references: this.cellRefHolder,
        result: CellState.OK,
      };
    } catch (e) {
      if (this.evaluationResult == CellState.OK) {
        // Assign a default result (and print exception) if the exception was thrown by an unknown reason.
        // This can for example be caused by a function crash.
        this.evaluationResult = (e as Error)
          .toString()
          .includes("Cannot convert")
          ? CellState.INVALID_OPERANDS
          : CellState.UNKNOWN_ERROR;
      }
      return { references: this.cellRefHolder, result: this.evaluationResult };
    }
  }

  updatePositionalReferences(
    from: Coordinate,
    to: Coordinate,
    targetSheet: Sheet,
  ): string {
    if (!this.rawValue.startsWith("=")) return this.rawValue;

    const fromLabel = targetSheet.getColumnLabel(from.column)!;
    const toLabel = targetSheet.getColumnLabel(to.column)!;
    return this.updateReferenceSymbols(fromLabel, toLabel, from.row, to.row);
  }

  updateReferenceSymbols(
    fromLabel: string,
    toLabel: string,
    fromRow: number,
    toRow: number,
  ): string {
    if (!this.rawValue.startsWith("=")) return this.rawValue;

    const expression = this.rawValue.substring(1);
    const parseResult = math.parse(expression);
    const fromSymbol = fromLabel + (fromRow + 1);
    const toSymbol = toLabel + (toRow + 1);

    // Update each symbol in the expression.
    const transform = parseResult.transform((node) =>
      this.updateReferenceSymbol(node, fromSymbol, toSymbol),
    );
    return `=${transform.toString()}`;
  }

  private injectFunctionParameter(node: MathNode, cellPos: CellReference) {
    if (node instanceof math.FunctionNode) {
      const fName = (node as math.FunctionNode).fn.name;
      // If the node is a custom function, inject the cell position as the first argument.
      // This allows custom functions to access the cell invoking the function.
      if (ExpressionHandler.functions.has(fName)) {
        // Copy CellReference into mathjs node structure.
        const record: Record<string, any> = {};
        for (const key in cellPos) {
          record[key] = new math.ConstantNode(
            cellPos[key as keyof CellReference],
          );
        }
        node.args.unshift(new math.ObjectNode(record));
      }
    }
    return node;
  }

  private updateReferenceSymbol(
    node: MathNode,
    targetSymbol: string,
    newSymbol: string,
  ) {
    if (!(node instanceof math.SymbolNode)) {
      return node;
    }
    const symbolNode = node as math.SymbolNode;

    let prefix = "";
    let symbol = symbolNode.name;
    if (symbol.includes("!")) {
      const parts = symbol.split("!");
      prefix = parts[0] + "!";
      symbol = parts[1];
    }

    if (symbol.includes(":")) {
      const parts = symbol.split(":");
      if (parts.length != 2) {
        return symbolNode;
      }
      const newStart = parts[0] === targetSymbol ? newSymbol : parts[0];
      const newEnd = parts[1] === targetSymbol ? newSymbol : parts[1];
      symbolNode.name = prefix + newStart + ":" + newEnd;
      return symbolNode;
    }

    if (symbol !== targetSymbol) {
      return symbolNode;
    }

    symbolNode.name = prefix + newSymbol;
    return symbolNode;
  }

  private resolveFunction(
    name: string,
  ): (cellRef: CellReference, ...args: any[]) => string {
    const fun = ExpressionHandler.functions.get(name.toLowerCase());
    if (!fun) {
      console.log(
        `Undefined function "${name}" in expression "${this.rawValue}"`,
      );
      return () => "";
    }

    return fun;
  }

  static registerFunction(
    name: string,
    fun: (cellRef: CellReference, ...args: any[]) => string,
  ) {
    ExpressionHandler.functions.set(name.toLowerCase(), fun);
  }

  private resolveSymbol(symbol: string): any {
    let targetSheet = this.sheet;
    if (symbol.includes("!")) {
      const parts = symbol.split("!").filter((s) => s !== "");
      if (parts.length != 2) {
        this.evaluationResult = CellState.INVALID_REFERENCE;
        throw new Error("Invalid sheet reference: " + symbol);
      }

      const sheetName = parts[0];
      const refSheet = SheetHolder.getInstance().getSheetByName(sheetName);
      if (!refSheet) {
        this.evaluationResult = CellState.INVALID_REFERENCE;
        throw new Error("Invalid sheet reference: " + symbol);
      }
      targetSheet = refSheet;
      symbol = parts[1];
    }

    if (symbol.includes(":")) {
      const rangeParts = symbol.split(":").filter((s) => s !== "");
      if (rangeParts.length != 2) {
        this.evaluationResult = CellState.INVALID_EXPRESSION;
        throw new Error("Invalid range: " + symbol);
      }

      const start = this.parseSymbolToPosition(rangeParts[0], targetSheet);
      const end = this.parseSymbolToPosition(rangeParts[1], targetSheet);

      const values: string[] = [];
      for (let i = start.rowIndex; i <= end.rowIndex; i++) {
        for (let j = start.colIndex; j <= end.colIndex; j++) {
          const cellInfo = targetSheet.getCellInfoAt(j, i);
          this.cellRefHolder.push({
            sheetKey: targetSheet.key,
            position: { column: j, row: i },
          });
          values.push(cellInfo?.resolvedValue ?? "");
        }
      }
      return values;
    }

    const { colIndex, rowIndex } = this.parseSymbolToPosition(
      symbol,
      targetSheet,
    );

    const cellInfo = targetSheet.getCellInfoAt(colIndex, rowIndex);

    this.cellRefHolder.push({
      sheetKey: targetSheet.key,
      position: {
        column: colIndex,
        row: rowIndex,
      },
    });
    return cellInfo?.resolvedValue ?? "";
  }

  private parseSymbolToPosition(
    symbol: string,
    targetSheet: Sheet,
  ): {
    colIndex: number;
    rowIndex: number;
  } {
    const letterGroups = symbol.split(/[0-9]+/).filter((s) => s !== "");
    if (letterGroups.length != 1) {
      this.evaluationResult = CellState.INVALID_SYMBOL;
      throw new Error("Invalid symbol: " + symbol);
    }

    const columnStr = letterGroups[0];
    const colIndex =
      targetSheet.getColumnByLabel(columnStr)?.position ??
      LightsheetHelper.resolveColumnIndex(columnStr);

    if (colIndex == -1) {
      this.evaluationResult = CellState.INVALID_SYMBOL;
      throw new Error("Invalid symbol: " + symbol);
    }

    const rowStr = symbol.substring(columnStr.length);
    const rowIndex = parseInt(rowStr) - 1;

    if (isNaN(rowIndex)) {
      this.evaluationResult = CellState.INVALID_SYMBOL;
      throw new Error("Invalid symbol: " + symbol);
    }

    return { colIndex, rowIndex };
  }
}

import Sheet from "../../../src/core/structure/sheet";

describe("Column label tests", () => {
  let sheet: Sheet;

  beforeEach(() => {
    window.sheetHolder?.clear();
    sheet = new Sheet("Sheet");
    sheet.setCellAt(0, 0, "A1");
    sheet.setCellAt(0, 1, "A2");
    sheet.setCellAt(1, 0, "B1");
    sheet.setCellAt(1, 1, "B2");
  });

  it("should set and clear column label", () => {
    sheet.setColumnLabel(0, "First");
    const col = sheet["columns"].get(sheet["columnPositions"].get(0)!)!;
    expect(sheet.getColumnLabel(0)).toBe("First");
    expect(sheet.getColumnByLabel("First")).toBe(col);

    sheet.setColumnLabel(0, null);
    expect(sheet.getColumnLabel(0)).toBe("A");
    expect(sheet.getColumnByLabel("First")).toBe(null);
  });

  it("should resolve a reference to a named column", () => {
    sheet.setColumnLabel(2, "Third");
    sheet.setCellAt(2, 0, "test");
    const info = sheet.setCellAt(0, 0, "=Third1");
    expect(info.resolvedValue).toBe("test");
  });

  it("should resolve a reference to a labeled column in another sheet", () => {
    const secSheet = new Sheet("Sheet2");
    secSheet.setCellAt(0, 0, "test");
    secSheet.setColumnLabel(0, "First");
    sheet.setCellAt(0, 0, "=Sheet2!First1");
    sheet.setCellAt(0, 1, "=Sheet2!A1");

    const nameRef = sheet.getCellInfoAt(0, 0);
    const posRef = sheet.getCellInfoAt(0, 1);

    expect(nameRef!.resolvedValue).toBe("test");
    expect(posRef!.resolvedValue).toBe("test");
  });
});

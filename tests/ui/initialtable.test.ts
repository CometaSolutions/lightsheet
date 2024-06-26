import Lightsheet from "../../src/main";

describe("Lightsheet", () => {
  let lightSheet: Lightsheet;
  let targetElementMock: HTMLElement;
  const onReady = jest.fn();

  beforeEach(() => {
    // Mocking UI and target element
    targetElementMock = document.createElement("div");

    // Creating instance of Lightsheet with mocked dependencies
    window.sheetHolder?.clear();
    lightSheet = new Lightsheet(
      {
        sheetName: "Sheet",
        onReady: onReady,
      },
      targetElementMock,
    );

    document.body.appendChild(targetElementMock);
    expect(onReady).toHaveBeenCalledWith();
  });

  afterEach(() => {
    document.body.removeChild(targetElementMock);
  });

  test("initializeData method creates header with the correct number of children in UI", () => {
    // Get the number of children of tableHeadDom
    const tbodyElement = targetElementMock.querySelector("tbody");
    // Get the number of children of tableBodyDom
    const rowCount = tbodyElement?.children.length;
    const colCount = tbodyElement?.children[0].children.length;

    expect(lightSheet.options.defaultColCount).toBeDefined();
    expect(lightSheet.options.defaultRowCount).toBeDefined();
    // Expect that the number of children matches the defaultRowCount since no data is provided
    expect(colCount).toBe(lightSheet.options.defaultColCount! + 1);
    expect(rowCount).toBe(lightSheet.options.defaultRowCount);
  });
});

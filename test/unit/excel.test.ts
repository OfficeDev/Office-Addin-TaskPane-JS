import * as assert from "assert";
import "mocha";
import { OfficeMockObject } from "office-addin-mock";

/* global describe, global, it, require */

const excelMockContext = {
  workbook: {
    range: {
      address: "G4",
      format: {
        fill: {},
      },
    },
    getSelectedRange: function () {
      return this.range;
    },
  },
};

const ExcelMockData = {
  context: excelMockContext,
  run: async function (callback: (context: typeof excelMockContext) => Promise<void> | void) {
    await callback(this.context);
  },
};

const OfficeMockData = {
  onReady: async function () {},
};

describe("Excel", function () {
  it("Run", async function () {
    const excelMock: OfficeMockObject = new OfficeMockObject(ExcelMockData); // Mocking the host specific namespace
    global.Excel = excelMock as any;
    global.Office = new OfficeMockObject(OfficeMockData) as any; // Mocking the common office-js namespace

    const { run } = require("../../src/taskpane/excel");
    await run();

    assert.strictEqual(excelMock.context.workbook.range.format.fill.color, "yellow");
  });
});

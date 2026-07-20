import * as assert from "assert";
import "mocha";
import { OfficeMockObject } from "office-addin-mock";

/* global describe, global, it, require, Word */

type MockParagraph = {
  font: {
    color?: string;
  };
  insertLocation?: string;
  text: string;
};

const wordMockContext = {
  document: {
    body: {
      paragraph: {
        font: {},
        text: "",
      } as MockParagraph,
      insertParagraph: function (paragraphText: string, insertLocation: string): MockParagraph {
        this.paragraph.text = paragraphText;
        this.paragraph.insertLocation = insertLocation;
        return this.paragraph;
      },
    },
  },
};

const WordMockData = {
  context: wordMockContext,
  InsertLocation: {
    end: "End",
  },
  run: async function (callback: (context: typeof wordMockContext) => Promise<void> | void) {
    await callback(this.context);
  },
};

const OfficeMockData = {
  onReady: async function () {},
};

describe("Word", function () {
  it("Run", async function () {
    const wordMock: OfficeMockObject = new OfficeMockObject(WordMockData); // Mocking the host specific namespace
    global.Word = wordMock as any;
    global.Office = new OfficeMockObject(OfficeMockData) as any; // Mocking the common office-js namespace

    const { run } = require("../../src/taskpane/word");
    await run();

    assert.strictEqual(wordMock.context.document.body.paragraph.font.color, "blue");
  });
});

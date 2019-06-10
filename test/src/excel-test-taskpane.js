import { pingTestServer, sendTestResults } from "office-addin-test-helpers";
import { run } from "../../src/taskpane/excel";
import * as testHelpers from "./test-helpers";
const port = 4201;
let testValues = [];


Office.onReady(async (info) => {
    if (info.host === Office.HostType.Excel) {
        const testServerResponse = await pingTestServer(port);
        if (testServerResponse["status"] == 200) {
            await runTest();
        }
    }
});

export async function runTest() {
    return new Promise(async (resolve, reject) => {
        try {
            // Execute taskpane code
            await run();

            // Get output of executed taskpane code
            await Excel.run(async context => {
                const range = context.workbook.getSelectedRange();
                const cellFill = range.format.fill;
                cellFill.load('color');
                await context.sync();

                testHelpers.addTestResult(testValues, "fill-color", cellFill.color, "#FFFF00");
                await sendTestResults(testValues, port);
                testValues.pop();
                await testHelpers.closeWorkbook();
                resolve();
            });
        } catch {
            reject();
        }
    });
}
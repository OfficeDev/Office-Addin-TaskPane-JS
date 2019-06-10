import { pingTestServer, sendTestResults } from "office-addin-test-helpers";
import { run } from "../../src/taskpane/word";
import * as testHelpers from "./test-helpers";
const port = 4201;
let testValues = [];

Office.onReady(async(info) => {
    if (info.host === Office.HostType.Word) {
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
            Word.run(async (context) => {
                var firstParagraph = context.document.body.paragraphs.getFirst();
                firstParagraph.load("text");
                await context.sync();

                testHelpers.addTestResult(testValues, "output-message", firstParagraph.text, "Hello World");
                await sendTestResults(testValues, port);
                testValues.pop();
                resolve();
            });
        } catch {
            reject();
        }
    });
}
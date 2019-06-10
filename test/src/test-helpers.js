import * as childProcess from "child_process";
import * as cps from "current-processes";

export async function closeDesktopApplication(application) {
    return new Promise(async function (resolve, reject) {
        let processName = "";
        switch (application.toLowerCase()) {
            case "excel":
                processName = "Excel";
                break;
            case "powerpoint":
                processName = (process.platform === "win32") ? "Powerpnt" : "Powerpoint";
                break;
            case "onenote":
                processName = "Onenote";
                break;
            case "outlook":
                processName = "Outlook";
                break;
            case "project":
                processName = "Project";
                break;
            case "word":
                processName = (process.platform === "win32") ? "Winword" : "Word";
                break;
            default:
                reject(`${application} is not a valid Office desktop application.`);
        }

        try {
            let appClosed = false;
            if (process.platform == "win32") {
                const cmdLine = `tskill ${processName}`;
                appClosed = await executeCommandLine(cmdLine);
            } else {
                const pid = await getProcessId(processName);
                if (pid != undefined) {
                    process.kill(pid);
                    appClosed = true;
                } else {
                    resolve(false);
                }
            }
            resolve(appClosed);
        } catch (err) {
            reject(`Unable to kill ${application} process. ${err}`);
        }
    });
}

export async function closeWorkbook() {
    return new Promise(async (resolve, reject) => {
        try {
            await Excel.run(async context => {
                // @ts-ignore
                context.workbook.close(Excel.CloseBehavior.skipSave);
                resolve();
            });
        } catch {
            reject();
        }
    });
}

export function addTestResult(testValues, resultName, resultValue, expectedValue) {
    var data = {};
    data["expectedValue"] = expectedValue;
    data["resultName"] = resultName;
    data["resultValue"] = resultValue;
    testValues.push(data);
}

async function getProcessId(processName) {
    return new Promise(async function (resolve, reject) {
        cps.get(function (err, processes) {
            try {
                const processArray = processes.filter(function (p) {
                    return (p.name.indexOf(processName) > 0);
                });
                resolve(processArray.length > 0 ? processArray[0].pid : undefined);
            }
            catch {
                reject(err);
            }
        });
    });
}

async function executeCommandLine(cmdLine) {
    return new Promise((resolve, reject) => {
        childProcess.exec(cmdLine, (error) => {
            if (error) {
                reject(false);
            } else {
                resolve(true);
            }
        });
    });
}
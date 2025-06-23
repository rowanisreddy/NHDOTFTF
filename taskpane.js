Office.onReady(() => {
    document.getElementById("submit-form").onclick = submitForm;
    document.getElementById("sync-online").onclick = syncToServer;
    loadOfflineQueue();
});

let offlineQueue = [];

function loadOfflineQueue() {
    let saved = Office.context.document.settings.get("offlineQueue");
    offlineQueue = saved ? JSON.parse(saved) : [];
}

function saveOfflineQueue() {
    Office.context.document.settings.set("offlineQueue", JSON.stringify(offlineQueue));
    Office.context.document.settings.saveAsync();
}

async function submitForm() {
    try {
        await Excel.run(async (context) => {
            let inspectionSheet = context.workbook.worksheets.getItem("Inspection");
            // Load form inputs in the order of InspectionData table columns (B to AW)
            const ranges = [
                "B2", "B5", "B3", "B14", "B15", "B16", "B17", "B18", "B19", "B20", "B21",
                "C14", "C15", "C16", "C17", "C18", "C19", "C20", "C21",
                "B23", "B24", "B25", "B26", "B27", "B28", "B29", "B30", "B31", "B32",
                "C23", "C24", "C25", "C26", "C27", "C28", "C29", "C30", "C31", "C32",
                "B33", "B35", "B36", "B37", "B38", "B39", "B40", "B41", "B42"
            ];
            let rangeObjects = ranges.map(range => inspectionSheet.getRange(range).load("values"));
            await context.sync();

            // Extract values from ranges
            let values = rangeObjects.map(range => range.values[0][0]);

            // Normalize boolean values for photo fields (indices 40 to 47 correspond to B35:B42)
            for (let i = 40; i <= 47; i++) {
                values[i] = values[i] === "TRUE" || values[i] === true;
            }

            // Handle offline submission
            if (!navigator.onLine) {
                offlineQueue.push(values);
                saveOfflineQueue();
                document.getElementById("status").innerText = "Offline: Form queued.";
                await clearForm(context, inspectionSheet);
                return;
            }

            // Online submission to InspectionData table
            let dataSheet = context.workbook.worksheets.getItem("InspectionData");
            let table = dataSheet.tables.getItem("InspectionDataTable");
            let dataRange = table.getDataBodyRange().load("values");
            await context.sync();

            // Generate new LogID
            let lastLogId = dataRange.values.length ? dataRange.values[dataRange.values.length - 1][0] : "INS000";
            let newLogId = "INS" + (parseInt(lastLogId.replace("INS", "")) + 1).toString().padStart(3, "0");

            // Create new row with 49 columns (LogID + 48 values)
            let newRow = [newLogId, ...values];
            table.rows.add(null, [newRow]);
            await context.sync();

            // Clear form inputs except B2, B3, B5
            await clearForm(context, inspectionSheet);
            document.getElementById("status").innerText = `Submitted! LogID: ${newLogId}`;
        });
    } catch (error) {
        document.getElementById("status").innerText = `Error: ${error.message}`;
    }
}

async function syncToServer() {
    try {
        if (!navigator.onLine) {
            document.getElementById("status").innerText = "Offline: Cannot sync.";
            return;
        }
        if (offlineQueue.length === 0) {
            document.getElementById("status").innerText = "No queued submissions.";
            return;
        }
        await Excel.run(async (context) => {
            let dataSheet = context.workbook.worksheets.getItem("InspectionData");
            let table = dataSheet.tables.getItem("InspectionDataTable");
            let dataRange = table.getDataBodyRange().load("values");
            await context.sync();

            // Sync each queued submission
            for (let submission of offlineQueue) {
                let lastLogId = dataRange.values.length ? dataRange.values[dataRange.values.length - 1][0] : "INS000";
                let newLogId = "INS" + (parseInt(lastLogId.replace("INS", "")) + 1).toString().padStart(3, "0");
                let newRow = [newLogId, ...submission];
                table.rows.add(null, [newRow]);
            }
            await context.sync();

            // Placeholder for SharePoint sync (to be updated with IT-provided endpoint)
            // for (let submission of offlineQueue) {
            //     await fetch("https://your-sharepoint-site/_api/lists/getbytitle('InspectionData')/items", {
            //         method: "POST",
            //         headers: { "Content-Type": "application/json", "Authorization": "Bearer your-access-token" },
            //         body: JSON.stringify({
            //             Title: newLogId,
            //             Date: submission[0],
            //             StationID: submission[1],
            //             TechnicianName: submission[2],
            //             // Map other fields as needed
            //         })
            //     });
            // }

            // Clear queue after syncing
            offlineQueue = [];
            saveOfflineQueue();
            document.getElementById("status").innerText = "Synced to Excel!";
        });
    } catch (error) {
        document.getElementById("status").innerText = `Sync Error: ${error.message}`;
    }
}

async function clearForm(context, inspectionSheet) {
    // Clear input ranges except B2, B3, B5
    let clearRanges = ["B14:B21", "B23:B33", "B35:B42", "C14:C21", "C23:C32"];
    for (let range of clearRanges) {
        let [start, end] = range.split(":");
        let rowStart = parseInt(start.match(/\d+/)[0]);
        let rowEnd = end ? parseInt(end.match(/\d+/)[0]) : rowStart;
        let numRows = rowEnd - rowStart + 1;
        let values = Array(numRows).fill([null]);
        inspectionSheet.getRange(range).values = values;
    }
    await context.sync();
}
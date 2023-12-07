
(function () {
    "use strict";

    // Declaration of global variables for later use
    let messageBanner;
    let dialog
    let accessToken;  // used to store user's access token
    let runningEnvir
    let tableListeners = { "sensei_lessonslearned": null }
    let myTables = { "sensei_lessonslearned": null }

    // Constants for client ID, redirect URL, and resource domain for authentication
    const clientId = "be63874f-f40e-433a-9f35-46afa1aef385"
    const redirectUrl = "https://seamus77x.github.io/index.html"
    const resourceDomain = "https://gsis-pmo-australia-sensei-dev.crm6.dynamics.com/"

    // Initialization function that runs each time a new page is loaded.
    Office.initialize = function (reason) {
        $(function () {
            try {

                //Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
                //Office.context.document.settings.saveAsync();

                Office.addin.setStartupBehavior(Office.StartupBehavior.load);
                Office.addin.showAsTaskpane();
                //Office.addin.hide();
                //Office.addin.setStartupBehavior(Office.StartupBehavior.none);

                switch (Office.context.platform) {
                    case Office.PlatformType.PC:
                        runningEnvir = Office.PlatformType.PC
                        console.log('I am running in Desktop Excel on Windows');
                        break;
                    case Office.PlatformType.Mac:
                        runningEnvir = Office.PlatformType.Mac
                        console.log('I am running in Desktop Excel on Mac');
                        break;
                    case Office.PlatformType.OfficeOnline:
                        runningEnvir = Office.PlatformType.OfficeOnline
                        console.log('I am running in Web Excel');
                        break;
                    case Office.PlatformType.iOS:
                        runningEnvir = Office.PlatformType.iOS
                        console.log('I am running in Excel on iOS');
                        break;
                    case Office.PlatformType.Android:
                        runningEnvir = Office.PlatformType.Android
                        console.log('I am running in Excel on Android');
                        break;
                    // You can add more cases here as needed
                    default:
                        runningEnvir = PlatformNotFound
                        console.log('Platform not identified');
                        break;
                }

                // Notification mechanism initialization and hiding it initially
                let element = document.querySelector('.MessageBanner');
                messageBanner = new components.MessageBanner(element);
                messageBanner.hideBanner();

                // Fallback logic for versions of Excel older than 2016
                if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
                    throw new Error("Sorry, this add-in only works with newer versions of Excel.")
                }

                // add external js
                //$('#myScriptX').attr('src', 'Test.js')
                //$.getScript('Test.js', function () {
                //    externalFun()
                //})

                // UI text setting for buttons and descriptions
                $('#button1-text').text("Download");
                $("#button1").attr("title", "Load Data to Excel")
                $('#button1').on("click", loadSampleData);

                $('#button2-text').text("Button 2");

                // Authentication and access token retrieval logic
                if (typeof accessToken === 'undefined') {
                    // Constructing authentication URL
                    let authUrl = "https://login.microsoftonline.com/common/oauth2/authorize" +
                        "?client_id=" + clientId +
                        "&response_type=token" +
                        "&redirect_uri=" + redirectUrl +
                        "&response_mode=fragment" +
                        "&resource=" + resourceDomain;

                    // Displaying authentication dialog
                    Office.context.ui.displayDialogAsync(authUrl, { height: 30, width: 30, requireHTTPS: true },
                        function (result) {
                            if (result.status === Office.AsyncResultStatus.Failed) {
                                // If the dialog fails to open, throw an error
                                throw new Error("Failed to open dialog: " + result.error.message);
                            }
                            dialog = result.value;
                            dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
                        }
                    );
                }
            } catch (error) {
                errorHandler(error.message)
            }
        });
    }

    Office.actions.associate("buttonFunction", function (event) {
        console.log('Hey, you just pressed a ribbon button.')
        console.log(myTables)

        event.completed();
    })

    // Process message (access token) received from the dialog
    function processMessage(arg) {
        try {
            // Check if the message is present
            if (!arg.message) {
                throw new Error("No message received from the dialog.");
            }

            // Parse the JSON message received from the dialog
            const response = JSON.parse(arg.message);

            // Check the status of the response
            if (response.Status === "Success") {
                // store the token in memory for later use
                accessToken = response.AccessToken
                console.log("Access Token Received")
            } else if (response.Status === "Error") {
                // Handle the error scenario
                errorHandler(response.Message || "An error occurred.");
            } else {
                // Handle unexpected status
                errorHandler("Unexpected response status.");
            }

        } catch (error) {
            // Handle any errors that occur during processing
            errorHandler(error.message);
        } finally {
            // Close the dialog, regardless of whether an error occurred
            if (dialog) {
                dialog.close();
            }
        }
    }

    // Function to load sample data
    async function loadSampleData() {
        const tableName = 'sensei_lessonslearned'
        const excludedColsNames = ['@odata.etag', 'sensei_lessonlearnedid']
        const odataCondition = '?$select=sensei_lessonlearnedid,sensei_name,sensei_lessonlearned,sensei_observation,sensei_actiontaken&$top=5'

        await loadData(`${resourceDomain}api/data/v9.1/${tableName}${odataCondition}`
            , tableName, 1, 'Sheet1', 'A1', excludedColsNames)

        registerTableChangeEvent(tableName)
    }
    //sc_integrationrecentgranulartransactions
    //sensei_financialtransaction
    //sensei_financialtransactions?$select=sc_kbrkey,sc_vendorname,sensei_value,sc_docdate,sensei_financialtransactionid&$top=50000

    // Function to retrieve data from Dynamics 365
    async function loadData(resourceUrl, tableName, Col_To_Paste_In_Table = 1, defaultSheet = 'Sheet1', defaultTpLeftRng = 'A1', excludedColsNames = ['@odata.etag']) {
        try {
            let DataArr = await Read_D365(resourceUrl);

            // act as the corresponding table in memory, which records the change in Excel table
            myTables[tableName] = JSON.parse(JSON.stringify(DataArr))

            // delete unwanted cols from the array which is going to be pasted into Excel
            let colIndices = excludedColsNames.map(colName => DataArr[0].indexOf(colName)).filter(index => index !== -1);
            // Sort the indices in descending order to avoid index shifting issues during removal
            colIndices.sort((a, b) => b - a);
            // Remove the columns with the found indices
            DataArr.map(row => {
                colIndices.forEach(colIndex => row.splice(colIndex, 1));
            });
            // report an error and interupt if failed to read data from Dataverse
            if (!DataArr || DataArr.length === 0) {
                throw new Error("No data retrieved or data array is empty");
            }
            // paste data into Excel worksheet 
            await Excel.run(async (ctx) => {
                const ThisWorkbook = ctx.workbook;
                const Worksheets = ThisWorkbook.worksheets;
                ctx.application.calculationMode = Excel.CalculationMode.manual;
                Worksheets.load("items/tables/items/name");

                await ctx.sync();

                let tableFound = false;
                let table;
                let oldRangeAddress;
                let oldFirstRow_formula
                let sheet

                if (tableName !== 'not using a table') {

                    // Attempt to find the existing table.
                    for (sheet of Worksheets.items) {
                        const tables = sheet.tables;

                        // Check if the table exists in the current sheet
                        table = tables.items.find(t => t.name === tableName);

                        // if the table found, delete the existing data
                        if (table) {
                            tableFound = true;
                            // Clear the data body range.
                            const dataBodyRange = table.getDataBodyRange();
                            dataBodyRange.load("address");
                            let firstRow = dataBodyRange.getRow(0);
                            firstRow.load('formulas');

                            dataBodyRange.clear();
                            await ctx.sync();
                            // Load the address of the range for new data insertion.
                            oldRangeAddress = dataBodyRange.address.split('!')[1];
                            oldFirstRow_formula = firstRow.formulas;
                            break;
                        }
                    }

                    if (tableFound) {
                        // delete header row of DataArr
                        DataArr.shift()

                        // add LHS and RHS formula cols to expand dataArr
                        let excelTableRightColNo = columnNameToNumber(oldRangeAddress.split(":")[1].replace(/\d+$/, ''))
                        let ppTableRightColNo = columnNameToNumber(oldRangeAddress.split(":")[0].replace(/\d+$/, '')) + Col_To_Paste_In_Table - 1 + DataArr[0].length - 1
                        DataArr.forEach(row => {
                            if (Col_To_Paste_In_Table > 1) {
                                let tempRowFormula = oldFirstRow_formula
                                row.unshift(...tempRowFormula[0].slice(0, Col_To_Paste_In_Table - 1))
                            }

                            if (excelTableRightColNo > ppTableRightColNo) {
                                let tempRowFormula = oldFirstRow_formula
                                row.push(...tempRowFormula[0].slice(ppTableRightColNo - excelTableRightColNo))
                            }
                        })

                        let newRangeAdress = oldRangeAddress.replace(/\d+$/, parseInt(oldRangeAddress.match(/\d+/)[0], 10) + DataArr.length - 1)
                        let range = sheet.getRange(newRangeAdress);

                        if (runningEnvir !== Office.PlatformType.OfficeOnline) {
                            range.values = DataArr;
                        } else {
                            pasteChunksToExcel(splitArrayIntoSmallPieces(DataArr), newRangeAdress, sheet, ctx)
                        }

                        // include header row when resize
                        let newRangeAdressWithHeader = newRangeAdress.replace(/\d+/, oldRangeAddress.match(/\d+/)[0] - 1)
                        let WholeTableRange = sheet.getRange(newRangeAdressWithHeader)
                        table.resize(WholeTableRange)

                        range.format.autofitColumns();
                        range.format.autofitRows();
                    } else {
                        // Situation 2: If the table doesn't exist, create a new one.
                        let tgtSheet = Worksheets.getItem(defaultSheet);
                        let endCellCol = columnNumberToName(columnNameToNumber(defaultTpLeftRng.replace(/\d+$/, "")) - 1 + DataArr[0].length)
                        let endCellRow = parseInt(defaultTpLeftRng.match(/\d+$/)[0], 10) + DataArr.length - 1
                        let rangeAddress = defaultTpLeftRng + ":" + endCellCol + endCellRow;
                        let range = tgtSheet.getRange(rangeAddress);

                        if (runningEnvir !== Office.PlatformType.OfficeOnline) {
                            range.values = DataArr;
                        } else {
                            pasteChunksToExcel(splitArrayIntoSmallPieces(DataArr), rangeAddress, tgtSheet, ctx)
                        }

                        let newTable = tgtSheet.tables.add(rangeAddress, true /* hasHeaders */);
                        newTable.name = tableName;

                        newTable.getRange().format.autofitColumns();
                        newTable.getRange().format.autofitRows();
                    }

                } else {
                    // Situation 3: paste the data in sheet directly, no table format
                    let tgtSheet = Worksheets.getItem(defaultSheet);
                    let endCellCol = columnNumberToName(columnNameToNumber(defaultTpLeftRng.replace(/\d+$/, "")) - 1 + DataArr[0].length)
                    let endCellRow = parseInt(defaultTpLeftRng.match(/\d+$/)[0], 10) + DataArr.length - 1
                    let rangeAddress = defaultTpLeftRng + ":" + endCellCol + endCellRow;
                    let range = tgtSheet.getRange(rangeAddress);

                    if (runningEnvir !== Office.PlatformType.OfficeOnline) {
                        range.values = DataArr;
                    } else {
                        pasteChunksToExcel(splitArrayIntoSmallPieces(DataArr), rangeAddress, tgtSheet, ctx)
                    }

                    range.format.autofitColumns();
                    range.format.autofitRows();
                }

                await ctx.sync();
            })  // end of pasting data
        } catch (error) {
            errorHandler(error.message)
        } finally {
            await Excel.run(async (ctx) => {
                ctx.application.calculationMode = Excel.CalculationMode.automatic;
                await ctx.sync()
            })
        }
    }

    function splitArrayIntoSmallPieces(data, maxChunkSizeInMB = 3.3) {

        const jsonString = JSON.stringify(data);
        const sizeInBytes = new TextEncoder().encode(jsonString).length;
        const sizeInMB = sizeInBytes / (1024 * 1024);

        console.log(`Total size: ${sizeInMB} MB`);

        if (sizeInMB <= maxChunkSizeInMB) {
            return [data]; // No need to chunk
        }

        let chunks = [];
        const totalRows = data.length;
        const rowsPerChunk = Math.ceil(totalRows * maxChunkSizeInMB / sizeInMB);

        for (let i = 0; i < totalRows; i += rowsPerChunk) {
            const chunk = data.slice(i, i + rowsPerChunk);
            chunks.push(chunk);
        }

        return chunks;
    }

    async function pasteChunksToExcel(chunks, rangeAddressToPaste, sheet, ctx) {
        const startCol = rangeAddressToPaste.match(/[A-Za-z]+/)[0];
        let startRow = parseInt(rangeAddressToPaste.match(/\d+/)[0], 10);

        const numberOfCols = chunks[0][0].length;
        const endCol = columnNumberToName(columnNameToNumber(startCol) + numberOfCols - 1);

        for (const chunk of chunks) {
            const chunkRowCount = chunk.length;
            const endRow = startRow + chunkRowCount - 1; // Calculate end row for the current chunk
            const rangeAddress = `${startCol}${startRow}:${endCol}${endRow}`;
            const range = sheet.getRange(rangeAddress);
            range.values = chunk;
            await ctx.sync();

            startRow += chunkRowCount; // Update startRow for the next chunk
        }
    }


    async function updateData() {
        Update_D365('sensei_lessonslearned', '0f0db491-3421-ee11-9966-000d3a798402', { 'sc_additionalcommentsnotes': 'Update Test' })
        //Create_D365('sensei_lessonslearned', { 'sensei_name': 'Add Test', 'sc_additionalcommentsnotes': 'ADD test from Web Add-In' })
        //Delete_D365('sensei_lessonslearned','f38edda5-8d8d-ee11-be35-6045bd3db52a')
    }


    // Function to create data in Dynamics 365
    async function Create_D365(entityLogicalName, addedData) {
        const url = `${resourceDomain}api/data/v9.1/${entityLogicalName}`;

        try {
            const response = await fetch(url, {
                method: 'POST',
                headers: {
                    'OData-MaxVersion': '4.0',
                    'OData-Version': '4.0',
                    'Accept': 'application/json',
                    'Content-Type': 'application/json; charset=utf-8',
                    'Authorization': `Bearer ${accessToken}`,
                    'Prefer': 'return=representation'
                },
                body: JSON.stringify(addedData)
            });

            if (!response.ok) {
                const errorData = await response.json();
                throw new Error(`Server responded with status ${response.status}: ${errorData.error?.message}`);
            }

            const responseData = await response.json();
            console.log("Record added successfully. New record ID:");
            //console.log(JSON.stringify(responseData))
            return responseData
        } catch (error) {
            if (error.name === 'TypeError') {
                // Handle network errors (e.g., no internet connection)
                errorHandler("Network error: " + error.message);
            } else {
                // Handle other types of errors (e.g., server responded with error code)
                errorHandler("Error encountered when adding new records in Dataverse:" + error.message);
            }
        }
    }
    // Function to read data in Dynamics 365
    async function Read_D365(url) {
        let totalRecords = 0;
        let finalArr = [];
        let startTime = new Date().getTime();

        try {
            do {
                let response = await fetch(url, {
                    method: 'GET',
                    headers: {
                        'OData-MaxVersion': '4.0',
                        'OData-Version': '4.0',
                        'Accept': 'application/json',
                        'Content-Type': 'application/json; charset=utf-8',
                        'Authorization': `Bearer ${accessToken}`,
                    }
                });

                if (!response.ok) {
                    const errorData = await response.json();
                    throw new Error(`Server responded with status ${response.status}: ${errorData.error?.message}`);
                }

                let jsonObj = await response.json();
                let headers = [];
                let tempArr_5k = [];

                if (jsonObj["value"] && jsonObj["value"].length > 0) {
                    for (let fieldName in jsonObj["value"][0]) {
                        if (typeof jsonObj["value"][0][fieldName] === "object" && jsonObj["value"][0][fieldName] != null) {
                            for (let relatedField in jsonObj["value"][0][fieldName]) {
                                let expandedFieldName = `${fieldName} / ${relatedField}`;
                                headers.push(expandedFieldName);
                            }
                        } else {
                            headers.push(fieldName);
                        }
                    }

                    tempArr_5k = [headers];

                    jsonObj["value"].forEach((row) => {
                        let itemWithRelatedFields = {};

                        for (let cell in row) {
                            if (typeof row[cell] === "object" && row[cell] !== null) {
                                for (let field in row[cell]) {
                                    let relatedFieldName = `${cell} / ${field}`;
                                    itemWithRelatedFields[relatedFieldName] = row[cell][field];
                                }
                            } else {
                                itemWithRelatedFields[cell] = row[cell];
                            }
                        }

                        let tempValRow = headers.map((header) => {
                            return itemWithRelatedFields[header] || null;
                        });

                        tempArr_5k.push(tempValRow);
                    });

                    if (totalRecords >= 1) {

                        let tempArr = [];
                        let headerRow = tempArr_5k[0];

                        for (let row of tempArr_5k) {
                            let tempValRow = [];
                            for (let fieldName of finalArr[0]) {
                                let trueColNo = headerRow.indexOf(fieldName);
                                tempValRow.push(row[trueColNo] || null);
                            }
                            tempArr.push(tempValRow);
                        }

                        tempArr.splice(0, 1);
                        finalArr = finalArr.concat(tempArr);
                    } else {
                        finalArr = finalArr.concat(tempArr_5k);
                    }
                }

                if (jsonObj["@odata.nextLink"]) {
                    url = jsonObj["@odata.nextLink"];
                } else {
                    url = null; // No more pages to retrieve
                }

                totalRecords += 1;
                console.log('HTTP Status Code: ' + response.status + ' - Page: ' + totalRecords);

            } while (url != null);

            // Update Excel with the collected data
            if (finalArr.length > 0) {
                let finishTime = new Date().getTime();
                console.log(`${(finishTime - startTime) / 1000} s used to download ${finalArr.length} records with ${finalArr[0].length} cols.`);

                return finalArr
            } else {
                throw new Error("No data downloaded");
            }


        } catch (error) {
            if (error.name === 'TypeError') {
                // Handle network errors (e.g., no internet connection)
                errorHandler("Network error: " + error.message);
            } else {
                // Handle other types of errors (e.g., server responded with error code)
                errorHandler("Error encountered when retrieving records from Dataverse:" + error.message);
            }
        }
    }
    // Function to update data in Dynamics 365
    async function Update_D365(entityLogicalName, recordId, updatedData) {
        const url = `${resourceDomain}api/data/v9.1/${entityLogicalName}(${recordId})`;

        try {
            const response = await fetch(url, {
                method: 'PATCH',
                headers: {
                    'OData-MaxVersion': '4.0',
                    'OData-Version': '4.0',
                    'Accept': 'application/json',
                    'Content-Type': 'application/json; charset=utf-8',
                    'Authorization': `Bearer ${accessToken}`,
                    //'Prefer': 'return=representation'
                },
                body: JSON.stringify(updatedData)
            });

            if (!response.ok) {
                // If the server responded with a non-OK status, handle the error
                const errorData = await response.json();
                throw new Error(`Server responded with status ${response.status}: ${errorData.error?.message}`);
            }

            console.log(`Record updated successfully. Updated record ID: [${recordId}]`);
        } catch (error) {
            if (error.name === 'TypeError') {
                // Handle network errors (e.g., no internet connection)
                errorHandler("Network error: " + error.message);
            } else {
                // Handle other types of errors (e.g., server responded with error code)
                errorHandler("Error encountered when updating records in Dataverse" + error.message);
            }
        }
    }
    // Function to delete data in Dynamics 365
    async function Delete_D365(entityLogicalName, recordId) {
        const url = `${resourceDomain}api/data/v9.1/${entityLogicalName}(${recordId})`;

        try {
            const response = await fetch(url, {
                method: 'DELETE',
                headers: {
                    'OData-MaxVersion': '4.0',
                    'OData-Version': '4.0',
                    'Authorization': `Bearer ${accessToken}`,
                    'Content-Type': 'application/json; charset=utf-8'
                }
            });

            if (!response.ok) {
                const errorData = await response.json();
                throw new Error(`Server responded with status ${response.status}: ${errorData.error?.message}`);
            }

            console.log(`Record with ID [${recordId}] deleted successfully.`);
        } catch (error) {
            if (error.name === 'TypeError') {
                // Handle network errors (e.g., no internet connection)
                errorHandler("Network error: " + error.message);
            } else {
                // Handle other types of errors (e.g., server responded with error code)
                errorHandler("Error encountered when deleting new records in Dataverse:" + error.message);
            }
        }
    }

    // Progress bar update function
    function updateProgressBar(progress) {
        let elem = document.getElementById("myProgressBar");
        elem.style.width = progress + '%';
        //elem.innerHTML = progress + '%';
    }
    //// Example: Update the progress bar every second
    //let progress = 0;
    //let interval = setInterval(function () {
    //    progress += 10; // Increment progress
    //    updateProgressBar(progress);

    //    if (progress >= 100) clearInterval(interval); // Clear interval at 100%
    //}, 1000);


    // Utility function to convert column number to name
    function columnNumberToName(columnNumber) {
        let columnName = "";
        while (columnNumber > 0) {
            let remainder = (columnNumber - 1) % 26;
            columnName = String.fromCharCode(65 + remainder) + columnName;
            columnNumber = Math.floor((columnNumber - 1) / 26);
        }
        return columnName;
    }
    // Utility function to convert column name to number
    function columnNameToNumber(columnName) {
        let columnNumber = 0;
        for (let i = 0; i < columnName.length; i++) {
            columnNumber *= 26;
            columnNumber += columnName.charCodeAt(i) - 64;
        }
        return columnNumber;
    }

    // Helper function for treating errors
    function errorHandler(error) {
        // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
        showNotification("Error", error);
        console.error("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }

    async function registerTableChangeEvent(tableName) {


        //Excel.run(function (context) {
        //    var sheet = context.workbook.worksheets.getActiveWorksheet();

        //    var table = sheet.tables.getItem("sensei_lessonslearned");

        //    var headerRange = table.getHeaderRowRange();
        //    headerRange.load("values, cellCount, id");

        //    return context.sync()
        //        .then(function () {
        //            for (var i = 0; i < headerRange.values[0].length; i++) {
        //                console.log("Header cell value: " + headerRange.values[0][i] + headerRange.clientId);
        //            }
        //        });
        //}).catch(function (error) {
        //    console.error("Error: " + error);
        //    if (error instanceof OfficeExtension.Error) {
        //        console.error("Debug info: " + JSON.stringify(error.debugInfo));
        //    }
        //});

        try {
            if (tableListeners[tableName] === true) {
                return
            }

            let ThisWorkbook;
            let Worksheets;

            Excel.run(function (ctx) {

                ThisWorkbook = ctx.workbook;
                Worksheets = ThisWorkbook.worksheets;
                Worksheets.load("items/tables/items/name");
                return ctx.sync().then(() => {
                    for (let sheet of Worksheets.items) {
                        const tables = sheet.tables;
                        // Check if the 'Test' table exists in the current sheet
                        let table = tables.items.find(t => t.name === tableName);

                        if (table) {
                            // if the table found, then listen to the change in the table
                            table.onChanged.add(handleTableChange);
                            tableListeners[tableName] = true
                            console.log(`I am tracking the changes in ${tableName}`)
                            break;
                        }
                    }

                    if (!tableListeners[tableName]) {
                        // if the table not found, then raise an error
                        throw new Error(`[${tableName}] table is not found in Excel`);
                    }
                }).then(ctx.sync);

            })
        } catch (error) {
            // Error handling for issues within the Excel.run block
            errorHandler("Error in registerTableChangeEvent: " + error.message);
        }
    }


    // hanle table change.    tip: get after value from Excel if multiple range changes
    function handleTableChange(eventArgs) {
        try {
            Excel.run(function (ctx) {
                // get the Range changed and the table changed
                let range = eventArgs.getRange(ctx)
                range.load("values, address, rowIndex, columnIndex, cellCount")
                let table = ctx.workbook.tables.getItem(eventArgs.tableId);
                table.load("name")
                let tableRange = table.getRange()
                tableRange.load("rowIndex, columnIndex")

                return ctx.sync().then(function () {

                    let tableStartRow = tableRange.rowIndex;
                    let tableStartCol = tableRange.columnIndex;

                    switch (eventArgs.changeType) {
                        case 'RangeEdited':
                            let rangeRowRelative = range.rowIndex - tableStartRow;
                            let rangeColRelative = range.columnIndex - tableStartCol;

                            console.log(`Table [${rangeRowRelative + 1}, ${rangeColRelative + 1}] is updated`);
                            console.log(`Range [${eventArgs.address}] in table [${table.name}] was just updated.`);

                            if (range.cellCount === 1) {
                                let jsonPayLoad = {}
                                jsonPayLoad[myTables[table.name][0][rangeColRelative + 2]] = range.values[0][0]
                                Update_D365(table.name, myTables[table.name][rangeRowRelative][1], jsonPayLoad)
                            }

                            break;
                        case "RowInserted":
                            console.log(`Row [${eventArgs.address}] was just inserted.`)
                            break;
                        case "RowDeleted":
                            console.log(`Row [${eventArgs.address}] was just deleted.`)
                            break;
                        case "ColumnInserted":
                            console.log(`Column [${eventArgs.address}] was just inserted.`)
                            break;
                        case "ColumnDeleted":
                            console.log(`Column [${eventArgs.address}] was just deleted.`)
                            break;
                        case "CellInserted":
                            console.log(`Cell [${eventArgs.address}] was just inserted.`)
                            break;
                        case "CellDeleted":
                            console.log(`Cell [${eventArgs.address}] was just deleted.`)
                            break;
                        default:
                            console.log(`Unknown action.`)
                            break;
                    }
                })
            })
        } catch (error) {
            errorHandler(error.message)
        }


    }







})();


(function () {
    "use strict";

    let cellToHighlight;
    let messageBanner;
    let dialog
    let accessToken

    const clientId = "be63874f-f40e-433a-9f35-46afa1aef385"
    const redirectUrl = "https://seamus77x.github.io/index.html"
    const resourceDomain = "https://gsis-pmo-australia-sensei-prod.crm6.dynamics.com/"
    //const clientSecret = "eaJ8Q~EGmAyxPpL8JUQ95-EP8cdkZjBBTHyAXcHh"
    //const tokenEndpoint = "https://login.microsoftonline.com/common/oauth2/token"


    //console.log(getData(ResourceUrl))

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(function () {
            // Initialize the notification mechanism and hide it
            let element = document.querySelector('.MessageBanner');
            messageBanner = new components.MessageBanner(element);
            messageBanner.hideBanner();

            // If not using Excel 2016, use fallback logic.
            if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
                $("#template-description").text("This sample will display the value of the cells that you have selected in the spreadsheet.");
                $('#button-text').text("Display!");
                $('#button-desc').text("Display the selection");

                $('#highlight-button').on("click", displaySelectedCells);
                return;
            }

            $("#template-description").text("This sample highlights the highest value from the cells you have selected in the spreadsheet.");
            $("#highlight-button").attr("title", "highlight the cell which contains the largest number in the selected range")
            $('#button-text').text("Highlight!");
            $('#button-desc').text("Highlights the largest number")
            $("#load-data-button").attr("title", "load some random numbers to the worksheet")
            $('#load-data-button-text').text("Load Data");
            $('#load-data-button-text-desc').text("Load Data to Excel");

            // Add a click event handler for the highlight button.
            $('#load-data-button').on("click", loadSampleData);
            $('#highlight-button').on("click", hightlightHighestValue);

            try {
                if (typeof accessToken === 'undefined') {
                    // If Access Token is not got yet, retrieve it.
                    let authUrl = "https://login.microsoftonline.com/common/oauth2/authorize" +
                        "?client_id=" + clientId +
                        "&response_type=token" +
                        "&redirect_uri=" + redirectUrl +
                        "&response_mode=fragment" +
                        "&resource=" + resourceDomain;

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
            }catch (error) {
                errorHandler(error.message)
            }


        });
    }





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

                // Proceed to exchange the auth code for an access token
                //exchangeAuthCodeForAccessToken(response.AuthCode);
                console.log("Access Token Received")
                accessToken = response.AccessToken

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
            // You may also choose to show a notification to the user, log the error, or take other actions
        } finally {
            // Close the dialog, regardless of whether an error occurred
            if (dialog) {
                dialog.close();
            }
        }
    }

    function loadSampleData() {
        loadData(`${resourceDomain}api/data/v9.1/sc_integrationrecentgranulartransactions?$top=4000`, 'Sheet1', 'A1', 'CJIT')
    }

    function displaySelectedCells() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    showNotification('Error', result.error.message);
                }
            });
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

    async function Retrieve_D365(url) {
        let totalRecords = 0;
        let finalArr = [];
        //let url = `https://gsis-pmo-australia-sensei-dev.api.crm6.dynamics.com/api/data/v9.1/sc_integrationrecentgranulartransactions?$top=500000`;

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
                0
                if (!response.ok) {
                    throw new Error('HTTP error, status = ' + response.status);
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
            errorHandler(error.message);
            // Handle the error appropriately
        }
    }


    function hightlightHighestValue() {
        loadData(`${resourceDomain}api/data/v9.1/sc_integrationrecentgranulartransactions?$top=100000`, 'Sheet1', 'F10', 'CJIT')
    }

    async function loadData(resourceUrl, defaultSheet, defaultTpLeftRng, tableName) {
        const CJI3_DataArr = await Retrieve_D365(resourceUrl);

        if (!CJI3_DataArr || CJI3_DataArr.length === 0) {
            throw new Error("No data retrieved or data array is empty");
        }

        await Excel.run(async (ctx) => {
            const workbook = ctx.workbook;
            const sheets = workbook.worksheets;
            sheets.load("items/tables/items/name");
            await ctx.sync();

            let tableFound = false;
            let table;
            let oldRangeAddress;
            let sheet

            // Attempt to find the existing table.
            for (sheet of sheets.items) {
                const tables = sheet.tables;

                // Check if the table exists in the current sheet
                table = tables.items.find(t => t.name === tableName);

                if (table) {
                    tableFound = true;
                    // Clear the data body range.
                    const dataBodyRange = table.getDataBodyRange();
                    dataBodyRange.load("address");
                    dataBodyRange.clear();
                    await ctx.sync();

                    // Load the address of the range for new data insertion.
                    oldRangeAddress = dataBodyRange.address.split('!')[1];
                    break;
                }
            }

            if (tableFound) {
                // Insert new data into the cleared data body range.
                const startCell = oldRangeAddress.replace(/\d+/, parseInt(oldRangeAddress.match(/\d+/)[0], 10) - 1).split(":")[0]
                const endCell = oldRangeAddress.replace(/\d+$/, parseInt(oldRangeAddress.match(/\d+/)[0], 10) + CJI3_DataArr.length - 2).split(":")[1]
                const range = sheet.getRange(`${startCell}:${endCell}`);
                range.values = CJI3_DataArr;
                table.resize(range)

                range.format.autofitColumns();
                range.format.autofitRows();
            } else {
                // If the table doesn't exist, create a new one.
                let tgtSheet = sheets.getItem(defaultSheet);
                let endCellCol = columnNumberToName( columnNameToNumber(defaultTpLeftRng.replace(/\d+$/, "")) - 1 + CJI3_DataArr[0].length )
                let endCellRow = parseInt(defaultTpLeftRng.match(/\d+$/)[0], 10) + CJI3_DataArr.length - 1
                const rangeAddress = defaultTpLeftRng + ":" + endCellCol + endCellRow;
                const range = tgtSheet.getRange(rangeAddress);
                range.values = CJI3_DataArr;
                const newTable = tgtSheet.tables.add(rangeAddress, true /* hasHeaders */);
                newTable.name = tableName;

                newTable.getRange().format.autofitColumns();
                newTable.getRange().format.autofitRows();
            }

            await ctx.sync();
        });
        // Catch block removed for brevity.
    }



    function updateProgressBar(progress) {
        let elem = document.getElementById("myProgressBar");
        elem.style.width = progress + '%';
        //elem.innerHTML = progress + '%';
    }


    // Example: Update the progress bar every second
    //let progress = 0;
    //let interval = setInterval(function () {
    //    progress += 10; // Increment progress
    //    updateProgressBar(progress);

    //    if (progress >= 100) clearInterval(interval); // Clear interval at 100%
    //}, 1000);

    function columnNumberToName(columnNumber) {
        let columnName = "";
        while (columnNumber > 0) {
            let remainder = (columnNumber - 1) % 26;
            columnName = String.fromCharCode(65 + remainder) + columnName;
            columnNumber = Math.floor((columnNumber - 1) / 26);
        }
        return columnName;
    }

    function columnNameToNumber(columnName) {
        let columnNumber = 0;
        for (let i = 0; i < columnName.length; i++) {
            columnNumber *= 26;
            columnNumber += columnName.charCodeAt(i) - 64;
        }
        return columnNumber;
    }

 


})();


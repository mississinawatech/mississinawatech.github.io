const labelTemplate = document.getElementById("label0").cloneNode(true); // Clone of initial blank label form for future reference
var lastLabelId = 0; // The number of the last label id
var messageBoxLocked = false;

var clientId = "622742585541-ra14g5hujq25gi5kbjj96a470ebjcae5.apps.googleusercontent.com";
var clientSecret = "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";

var spreadsheetId = "1j8J-HuPIwkLZCAPeSBgRK5macqOwLM81oY_0lqccmSU";
var touchDictionaryId = "1XU9F5bJbSnZMkl2RzZZDT1QxmU-WhGOFj6yj6aVHWuU";
var templateSheetId = "982309210";
var currentSheetId = "";
var currentSheetName = "";
var devices = [];
var matches = [];
var dictionary = [];

function insertBefore(referenceNode, newNode) {

    referenceNode.parentNode.insertBefore(newNode, referenceNode);

}

// Creates a new blank label on the page
function newLabel() {
    if (lastLabelId < 9) {
        const newLabelButton = document.getElementById("new-label");

        var cloneOfLabelTemplate = labelTemplate.cloneNode(true);
        lastLabelId++;
        cloneOfLabelTemplate.id = "label" + lastLabelId;

        insertBefore(newLabelButton, cloneOfLabelTemplate);
        
        if (lastLabelId == 9) {
            newLabelButton.remove();
        }

    }
}

// Pulls the access token from the url on a redirect and resets the url
function onLoad() {

    var url = window.location;

    accessTokenInURL = new URLSearchParams(url.hash).get('access_token');

    if (accessTokenInURL != null) {
        localStorage.setItem("accessToken", accessTokenInURL);
        window.location = window.location.pathname;
    }

    const statusIndicator = document.getElementById("status");
    statusIndicator.innerHTML = "Console Info Loading...";

    checkTokenExpiration();

    setInterval(function() {
        checkTokenExpiration();
    }, 60 * 1000); // 60 * 1000 milsec

    getModelDictionary();
    getChromebookList();

}

// Displays a message for a short period of time
function displayMessage(message) {

    const messageBox = document.getElementById("message-box");

    if (!messageBoxLocked) {
        messageBox.classList.add("message-box-displayed");
        messageBox.innerHTML = message;

        setTimeout(closeMessage, 5 * 1000); // 5 Seconds
    }

}

function closeMessage() {

    if (!messageBoxLocked) {
        const messageBox = document.getElementById("message-box");

        messageBox.classList.remove("message-box-displayed");
    }

}

// Authenticates with Google
function oauthSignIn() {

    // Google's OAuth 2.0 endpoint for requesting an access token
    var oauth2Endpoint = 'https://accounts.google.com/o/oauth2/v2/auth?access_type=offline';
  
    // Create <form> element to submit parameters to OAuth 2.0 endpoint.
    var form = document.createElement('form');
    form.setAttribute('method', 'GET'); // Send as a GET request.
    form.setAttribute('action', oauth2Endpoint);

    // Parameters to pass to OAuth 2.0 endpoint.
    var params = {'client_id': clientId,
                  'redirect_uri': window.location.protocol + '//' + window.location.host + window.location.pathname,
                  'response_type': 'token',
                  'scope': 'https://www.googleapis.com/auth/admin.directory.device.chromeos.readonly https://www.googleapis.com/auth/spreadsheets',
                  'include_granted_scopes': 'true',
                  'state': 'pass-through value'};

    console.log("Sending authentication request to Google:");

    // Add form parameters as hidden input values.
    for (var p in params) {
        var input = document.createElement('input');
        input.setAttribute('type', 'hidden');
        input.setAttribute('name', p);
        input.setAttribute('value', params[p]);
        form.appendChild(input);
        console.log(p + " - " + params[p]);
    }
  
    // Add form to page and submit it to open the OAuth 2.0 endpoint.
    document.body.appendChild(form);
    form.submit();

}

// Populates some of the data of a label using a device's id
function populateById(label) {

    resetDropdown(label);

    var currentLabelId = label.id;
    var currentLabelNumber = currentLabelId.replace("label",'');

    var id = label.children[0].children[3].children[1].value;

    // console.log("Searching for Chromebook by this ID: " + id);

    matches[currentLabelNumber] = []; // Resets match array for label

    for (var i = 0; i < devices.length; i++) {

        if (devices[i][1] == id) {
            matches[currentLabelNumber].push(devices[i]); // Adds device as a match for the asset-id
        }

    }

    if (matches[currentLabelNumber].length > 1) {

        // Shows the dropdown
        const dropdown = label.children[0].children[3].children[2];
        dropdown.style.display = "inline-block"; // Displays the dropdown

        const select = dropdown.getElementsByTagName("select")[0];
        // Clears all options from the select
        while (select.firstChild) {
            select.removeChild(select.lastChild);
        }

        // Adds options to dropdown
        for (var i = 0; i < matches[currentLabelNumber].length; i++) {
        
            var option = document.createElement("option");

            option.innerHTML = matches[currentLabelNumber][i][0];
            option.value = matches[currentLabelNumber][i][2];

            select.appendChild(option);   
        
        }

        // Sets the dropdown value to the last chromebook in matches
        select.value = select.firstChild.value;

        // Chooses the first Chromebook in matches (most likely to be the correct one)
        label.children[0].children[1].children[1].value = matches[currentLabelNumber][0][0];
        label.children[0].children[3].children[1].value = matches[currentLabelNumber][0][1];
        label.children[0].children[4].children[1].value = matches[currentLabelNumber][0][2];

        label.children[0].children[6].children[1].value = isTouchscreen(matches[currentLabelNumber][0][0]);

    } else if (matches[currentLabelNumber].length > 0) { // Checks if there are any matches

        label.children[0].children[1].children[1].value = matches[currentLabelNumber][0][0];
        label.children[0].children[3].children[1].value = matches[currentLabelNumber][0][1];
        label.children[0].children[4].children[1].value = matches[currentLabelNumber][0][2];

        label.children[0].children[6].children[1].value = isTouchscreen(matches[currentLabelNumber][0][0]);
    
    } else {
        // console.log("No Matching Chromebooks");
    }
}

// Populates some of the data of a label using a device's serial number
function populateBySerialNumber(label) {

    resetDropdown(label);

    var SN = label.children[0].children[4].children[1].value;

    var currentLabelId = label.id;
    var currentLabelNumber = currentLabelId.replace("label",'');

    // console.log("Searching for Chromebook by this Serial Number: " + SN);

    matches[currentLabelNumber] = []; // Resets match array

    for (var i = 0; i < devices.length; i++) {

        if (devices[i][2] == SN) {
            matches[currentLabelNumber].push(devices[i]); // Adds device as a match for the serial number
        }

    }

    if (matches[currentLabelNumber].length > 1) {
        console.log("There is two Chromebooks with the same Serial Number!!!");
    }
    if (matches[currentLabelNumber].length > 0) { // Checks if there are any matches

        label.children[0].children[1].children[1].value = matches[currentLabelNumber][0][0];
        label.children[0].children[3].children[1].value = matches[currentLabelNumber][0][1];
        label.children[0].children[4].children[1].value = matches[currentLabelNumber][0][2];

        label.children[0].children[6].children[1].value = isTouchscreen(matches[currentLabelNumber][0][0]);
    
    } else {
        // console.log("No Matching Chromebooks");
    }
}

// Hides the asset id dropdown
function resetDropdown(label) {

    const dropdown = label.children[0].children[3].children[2];
    dropdown.style.display = "none"; // Hides the dropdown

}

function updateDropdownId(label, value) {

    var currentLabelId = label.id;
    var currentLabelNumber = currentLabelId.replace("label",'');

    console.log("Update Chromebook via Dropdown: " + value)

    var matchIndex; // Index of the match to use

    // Finds the match based on selection
    for (var i = 0; i < matches[currentLabelNumber].length; i++) {
        
        if (matches[currentLabelNumber][i][2] == value) {
            matchIndex = i;
        }

    }

    // Changes the inputs based on selection
    label.children[0].children[1].children[1].value = matches[currentLabelNumber][matchIndex][0];
    label.children[0].children[3].children[1].value = matches[currentLabelNumber][matchIndex][1];
    label.children[0].children[4].children[1].value = matches[currentLabelNumber][matchIndex][2];

    label.children[0].children[6].children[1].value = isTouchscreen(matches[currentLabelNumber][matchIndex][0]); 

}

// Returns Yes if a chromebook is found that is touchscreen in the dictionary
function isTouchscreen(model) {

    for (var i = 0; i < dictionary.length; i++) {
        
        if (dictionary[i][0] == model) {
            return dictionary[i][1];
        }

    }

    return "No";

}

// Retrieves a list of all Chromebooks
function getChromebookList() {

    console.log("Retrieving Chromebook List");

    // Authentication
    var myHeaders = new Headers();
    var authString = "Bearer " + localStorage.getItem("accessToken");
    myHeaders.append("Authorization", authString);

    // Settings for GET request
    var requestOptions = {
        method: 'GET',
        headers: myHeaders,
        redirect: 'follow'
    };

    var url = "https://admin.googleapis.com/admin/directory/v1/customer/C019q0wqs/devices/chromeos?OrderBy=LAST_SYNC&maxResults=1000";

    devices = []; // A global array to populate with JSON objects

    getChromebookPage(url, requestOptions) // Recursively gets paginated list of Chromebooks
    
}

// Retrieves a single page of information
function getChromebookPage(
    url,
    requestOptions,
    pageToken = ""
  ) {
    return fetch(`${url}${pageToken}`, requestOptions) // Append the page token to the base URL
      .then(response => response.json())
      .then(result => {

        var chromebooks = result.chromeosdevices; // Gets Chromebook list from the result

        // Runs through all chromebooks in one page and adds their model, asset-id, and serial number to an array
        for (var i = 0; i < chromebooks.length; i++) {

            devices.push([chromebooks[i].model,chromebooks[i].annotatedAssetId,chromebooks[i].serialNumber]);

        }

        if (result.nextPageToken) {
            return getChromebookPage(url, requestOptions, ("&pageToken=" + result.nextPageToken));
        }

        console.log("Chromebook list has been loaded (" + devices.length + " total)");
 
        const statusIndicator = document.getElementById("status");
        statusIndicator.innerHTML = "Console Info Loaded &#10003;";

    })
      .catch(error => console.log('error', error));
  }

// Exports data to the Google Sheet
function exportToSheets() {

    console.log("Exporting to Google Sheets");

    createSheet();

}

// Duplicates the template of the sheet in the Google Sheet
function createSheet() {

    // Authentication
    var myHeaders = new Headers();
    var authString = "Bearer " + localStorage.getItem("accessToken");
    myHeaders.append("Authorization", authString);


    const date = new Date();

    const month = String(date.getMonth() + 1);
    const day = String(date.getDate());
    const year = String(date.getFullYear());

    var dateString = month + "-" + day + "-" + year;

    // var dateString = String(Math.random());

    currentSheetName = dateString;

    // Operations to perform on the spreadsheet
    var requestBody = {
        requests: [
            {
                duplicateSheet: {
                    sourceSheetId: templateSheetId,
                    newSheetName: dateString,
                    insertSheetIndex: 1000
                }
            }
        ]
    }

    // Settings for POST request
    var requestOptions = {
        method: 'POST',
        headers: myHeaders,
        redirect: 'follow',
        includeSpreadSheetInResponse: false,
        body: JSON.stringify(requestBody)
    }

    var url = "https://sheets.googleapis.com/v4/spreadsheets/" + spreadsheetId + ":batchUpdate";

    // Duplicates the template sheet
    fetch(url, requestOptions)
    .then(response => response.json())
    .then(result => {

        if (result.hasOwnProperty("error")) { // Handles errors

            var code = result.error.code;

            if (code == 401) { // Unauthorized
                displayMessage("You are not signed in!");
            } else if (code >= 400) { // Error

                if (result.error.message.includes("already exists")) { // A sheet with the same name already exists
                    displayMessage("You have already created a sheet today! Rename or delete that sheet to create another!");
                } else {
                    displayMessage("An error occurred!");
                }

            }

        } else {

            displayMessage("The sheet was created successfully!");

            currentSheetId = result.replies[0].duplicateSheet.properties.sheetId;

            populateSheet(dateString);

        }

    })
    .catch(error => console.log(error));

}

// Populates the new Google Sheet
function populateSheet(dateString) {

    // Authentication
    var myHeaders = new Headers();
    var authString = "Bearer " + localStorage.getItem("accessToken");
    myHeaders.append("Authorization", authString);

    var ranges = [];

    ranges[0] = currentSheetName + '!B1:B8';
    ranges[1] = currentSheetName + '!D1:D8';
    ranges[2] = currentSheetName + '!B9:B16';
    ranges[3] = currentSheetName + '!D9:D16';
    ranges[4] = currentSheetName + '!B17:B24';
    ranges[5] = currentSheetName + '!D17:D24';
    ranges[6] = currentSheetName + '!B25:B32';
    ranges[7] = currentSheetName + '!D25:D32';
    ranges[8] = currentSheetName + '!B33:B40';
    ranges[9] = currentSheetName + '!D33:D40';
    
    // Changes date to a / format
    dateString = dateString.replace('-', '/');
    dateString = dateString.replace('-', '/');

    var requestData = [];
    var numberOfLabels = lastLabelId + 1;
    for (var i = 0; i < numberOfLabels; i++) {

        var jsonData = {};
        jsonData['majorDimension'] = 'ROWS';
        jsonData['range'] = ranges[i];

        var labelId = "label" + i;
        var selectedLabel = document.getElementById(labelId);
        var inputs = selectedLabel.getElementsByTagName('input');
        var textAreas = selectedLabel.getElementsByTagName('textarea');
        var selection = selectedLabel.getElementsByTagName('select');

        var valueArray = [
            [inputs[0].value],
            [inputs[1].value],
            [textAreas[0].value],
            [inputs[2].value],
            [inputs[3].value],
            [textAreas[1].value],
            [selection[selection.length - 1].options[selection[selection.length - 1].selectedIndex].text],
            [dateString]
        ];

        jsonData['values'] = valueArray;
        
        requestData.push(jsonData);
    }

    // Operations to perform on the spreadsheet
    var requestBody = {
        valueInputOption: 'RAW',
        responseDateTimeRenderOption: 'SERIAL_NUMBER',
        responseValueRenderOption: 'FORMATTED_VALUE',
        includeValuesInResponse: false,
        data: requestData
    }


    // Settings for POST request
    var requestOptions = {
        method: 'POST',
        headers: myHeaders,
        redirect: 'follow',
        includeSpreadSheetInResponse: false,
        body: JSON.stringify(requestBody)
    }

    var url = "https://sheets.googleapis.com/v4/spreadsheets/" + spreadsheetId + "/values:batchUpdate";

    // Duplicates the template sheet
    fetch(url, requestOptions)
    .then(response => response.json())
    .then(result => {

        console.log(result);

        if (result.hasOwnProperty("error")) { // Handles errors

            var code = result.error.code;

            if (code == 401) { // Unauthorized
                displayMessage("You are not signed in!");
            } else if (code >= 400) { // Error
                displayMessage("An error occurred!");
            }

        } else {

            displayMessage("The sheet was populated successfully!");

        }

    })
    .catch(error => console.log(error));

}

function getModelDictionary() {

    console.log("Retrieving Model Dictionary");

    // Authentication
    var myHeaders = new Headers();
    var authString = "Bearer " + localStorage.getItem("accessToken");
    myHeaders.append("Authorization", authString);

    // Settings for GET request
    var requestOptions = {
        method: 'GET',
        headers: myHeaders,
        redirect: 'follow'
    };

    var url = "https://sheets.googleapis.com/v4/spreadsheets/" + touchDictionaryId + "/values:batchGet?ranges=A2:B50";

    dictionary = []; // A global array to populate with JSON objects

    fetch(url, requestOptions) // Append the page token to the base URL
    .then(response => response.json())
    .then(result => {

        dictionary = result["valueRanges"][0]["values"];
        console.log("Model Dictionary has been loaded");

    })
    .catch(error => console.log('error', error));

}

function checkTokenExpiration() {

    // Authentication
    var myHeaders = new Headers();
    var authString = "Bearer " + localStorage.getItem("accessToken");
    myHeaders.append("Authorization", authString);

    // Settings for GET request
    var requestOptions = {
        method: 'GET',
        headers: myHeaders,
        redirect: 'follow'
    };

    var url="https://oauth2.googleapis.com/tokeninfo?access_token=" + localStorage.getItem("accessToken");

    fetch(url, requestOptions) // Append the page token to the base URL
    .then(response => response.json())
    .then(result => {

        const timer = document.getElementById("timer");

        if (result["error"] == "invalid_token") {

            timer.innerHTML = "Please Sign In!"

            const statusIndicator = document.getElementById("status");
            statusIndicator.innerHTML = "Console Info Can't be Loaded";

        } else {

            var minutes = Math.floor(result["expires_in"] / 60);

            timer.innerHTML = minutes + " Minutes Before Sign Out";

        }

    })
    .catch(error => console.log('error', error));

}

var fs = require('fs');
var express = require('express');
var request = require('request');
var cheerio = require('cheerio');
var google = require('googleapis');
var googleAuth = require('google-auth-library');

// If modifying these scopes, delete your previously saved credentials
// at ~/.credentials/script-nodejs-quickstart.json
var SCOPES = ['https://www.googleapis.com/auth/spreadsheets'];
var TOKEN_DIR = (process.env.HOME || process.env.HOMEPATH ||
    process.env.USERPROFILE) + '/.credentials/';
var TOKEN_PATH = TOKEN_DIR + 'script-nodejs-quickstart.json';

var GLOBAL_CREDENTIALS ="";
var COUNT = 0;

// Load client secrets from a local file.
fs.readFile('client_secret.json', function processClientSecrets(err, content) {
    if (err) {
        console.log('Error loading client secret file: ' + err);
        return;
    }
    // Authorize a client with the loaded credentials, then call the
    // Google Apps Script Execution API.
    GLOBAL_CREDENTIALS = content;
    authorize(JSON.parse(content), callAppsScript);
});

/**
 * Create an OAuth2 client with the given credentials, and then execute the
 * given callback function.
 *
 * @param {Object} credentials The authorization client credentials.
 * @param {function} callback The callback to call with the authorized client.
 */
function authorize(credentials, callback) {
    var clientSecret = credentials.installed.client_secret;
    var clientId = credentials.installed.client_id;
    var redirectUrl = credentials.installed.redirect_uris[0];
    var auth = new googleAuth();
    var oauth2Client = new auth.OAuth2(clientId, clientSecret, redirectUrl);

    // Check if we have previously stored a token.
    fs.readFile(TOKEN_PATH, function(err, token) {
        if (err) {
            getNewToken(oauth2Client, callback);
        } else {
            oauth2Client.credentials = JSON.parse(token);
            callback(oauth2Client);
        }
    });
}

/**
 * Get and store new token after prompting for user authorization, and then
 * execute the given callback with the authorized OAuth2 client.
 *
 * @param {google.auth.OAuth2} oauth2Client The OAuth2 client to get token for.
 * @param {getEventsCallback} callback The callback to call with the authorized
 *     client.
 */
function getNewToken(oauth2Client, callback) {
    var authUrl = oauth2Client.generateAuthUrl({
        access_type: 'offline',
        scope: SCOPES
    });
    console.log('Authorize this app by visiting this url: ', authUrl);
    var rl = readline.createInterface({
        input: process.stdin,
        output: process.stdout
    });
    rl.question('Enter the code from that page here: ', function(code) {
        rl.close();
        oauth2Client.getToken(code, function(err, token) {
            if (err) {
                console.log('Error while trying to retrieve access token', err);
                return;
            }
            oauth2Client.credentials = token;
            storeToken(token);
            callback(oauth2Client);
        });
    });
}

/**
 * Store token to disk be used in later program executions.
 *
 * @param {Object} token The token to store to disk.
 */
function storeToken(token) {
    try {
        fs.mkdirSync(TOKEN_DIR);
    } catch (err) {
        if (err.code != 'EEXIST') {
            throw err;
        }
    }
    fs.writeFile(TOKEN_PATH, JSON.stringify(token));
    console.log('Token stored to ' + TOKEN_PATH);
}

/**
 * Call an Apps Script function to list the folders in the user's root
 * Drive folder.
 *
 * @param {google.auth.OAuth2} auth An authorized OAuth2 client.
 */
function callAppsScript(auth) {
    var scriptId = 'MVRGsU5m0FkGFCr5aQwY3zgXvYlv25EAc';
    var script = google.script('v1');

    // Make the API request. The request object is included here as 'resource'.
    script.scripts.run({
        auth: auth,
        resource: {
            function: 'getSpreadSheetAsObject'
        },
        scriptId: scriptId
    }, function(err, resp) {
        if (err) {
            // The API encountered a problem before the script started executing.
            console.log('The API returned an error: ' + err);
            return;
        }
        if (resp.error) {
            // The API executed, but the script returned an error.

            // Extract the first (and only) set of error details. The values of this
            // object are the script's 'errorMessage' and 'errorType', and an array
            // of stack trace elements.
            var error = resp.error.details[0];
            console.log('Script error message: ' + error.errorMessage);
            console.log('Script error stacktrace:');

            if (error.scriptStackTraceElements) {
                // There may not be a stacktrace if the script didn't start executing.
                for (var i = 0; i < error.scriptStackTraceElements.length; i++) {
                    var trace = error.scriptStackTraceElements[i];
                    console.log('\t%s: %s', trace.function, trace.lineNumber);
                }
            }
        }
        else {
            // Write result to file
            var inputFile = "./InputFiles/Input.json";
            var allClassesObj = resp.response.result;

            var allClassesString = JSON.stringify(allClassesObj);
            //fs.writeFile(inputFile, allClassesString);

            parseClassesData(allClassesObj);
        }

    });
}

/**
 * Parses a Javascript object and kicks off the parsing process
 *
 * @param {Object} Enumerates all sheets (class day/names) and people in each class
 */
function parseClassesData(allClassesObj){
    for (var className in allClassesObj){
        if(allClassesObj.hasOwnProperty(className)){
            var classObj = allClassesObj[className];
            for(var personName in classObj) {
                if (classObj.hasOwnProperty(personName)) {
                    var personObj = classObj[personName];
                    var nameArray = personName.split(",");
                    if (nameArray.length != 2){
                        console.log("Malformed name: " + personName);
                        continue;
                    }
                    var firstName = nameArray[1].trim().toLowerCase();
                    var lastName = nameArray[0].trim().toLowerCase();
                    getStudentInformation(allClassesObj, personObj, firstName, lastName);
                }
            }
        }
    }
}

/**
 * Sends HTTP request to Vassar student directory and saves result to a local file
 *
 * @param {Object} Enumerates all sheets (class day/names) and people in each class
 * @param {Object} Contains person information: isStudent, isEmployee, sName, eName
 * @param {String} First name
 * @param {String} Last name
 */
function getStudentInformation(allClassesObj, personObj, firstName, lastName){
    // send http request
    request.post({
            url: 'https://aisapps.vassar.edu/directory/studirctry_rslts.php',
            form: { firstname: firstName, lastname: lastName}
        },
        // Callback function
        function (error, response, htmlData){
            if (!error && response.statusCode == 200) {
                // Write result to file
                var fileName = "./HtmlFiles/" + lastName + "_" + firstName + "_S" + ".html";
                fs.writeFile(fileName, htmlData);
                getEmployeeInformation(allClassesObj, personObj, firstName, lastName);
            }
        }
    );
}

/**
 * Sends HTTP request to Vassar Employee directory and saves result to a local file
 *
 * @param {Object} Enumerates all sheets (class day/names) and people in each class
 * @param {Object} Contains person information: isStudent, isEmployee, sName, eName
 * @param {String} First name
 * @param {String} Last name
 */
function getEmployeeInformation(allClassesObj, personObj, firstName, lastName){
    // send http request
    request.post({
            url: 'https://aisapps.vassar.edu/directory/empdir_rslts.php',
            form: { firstname: firstName, lastname: lastName}
        },
        // Callback function
        function (error, response, htmlData){
            if (!error && response.statusCode == 200) {
                var fileName = "./HtmlFiles/" + lastName + "_" + firstName + "_E" + ".html";
                fs.writeFile(fileName, htmlData);
            }
        }
    );
}

 
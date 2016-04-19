/*
 * Creates and returns a well-formed Javascript object based on class information in spreadsheet
 *
 */
function getSpreadSheetAsObject(){
    var parentSheet = SpreadsheetApp.openById('1CrS9aupqZ8z76tR5aGUsncLL7BwcYzeSbEN8Ok_SFzU');

    // object that we are going to return
    var outputObject = {};

    // Iterate through all sheets/classes in the app
    var sheetsArray = parentSheet.getSheets();
    for(var i=0; i<sheetsArray.length; i++){
        var sheet = sheetsArray[i];
        var sheetName = sheet.getName();

        if(!sheetName.match('Mon')){
            // Create new object corresponding to sheet
            var sheetObject = {};
            outputObject[sheetName] = sheetObject;
            Logger.log(sheetName);

            var lastRow = sheet.getLastRow();
            for(var row = 3; row <= lastRow - 4; row++){
                // Get name and age and whatnot
                var nameColumn = 1;
                var range = sheet.getRange(row, nameColumn);
                var name = range.getValue().trim();

                // Create a new person object. Add to parent (sheetObject)
                if(name !== "" ){
                    var personObject = {};
                    personObject.isStudent = false;
                    personObject.isEmployee = false;
                    personObject.sName = "";
                    personObject.eName = "";
                    sheetObject[name] = personObject;
                }
            }
        }
    }

    return outputObject;
}

/*
 * Updates the spreadsheet based on whether each person is Student, Staff, or Unassociated with Vassar College
 *
 * @param {String}: a JSON string containing class / person information
 */
function updateSpreadSheetData(inputData) {
    var STUDENT = 'Student';
    var STAFF = 'Staff';
    var NONVC = 'Non-VC';

    var allClassesObj = JSON.parse(inputData);

    var parentSheet = SpreadsheetApp.openById('1CrS9aupqZ8z76tR5aGUsncLL7BwcYzeSbEN8Ok_SFzU');

    var sheetsArray = parentSheet.getSheets();
    for(var i=0; i<sheetsArray.length; i++){
        var sheet = sheetsArray[i];
        var sheetName = sheet.getName();

        if(allClassesObj.hasOwnProperty(sheetName)){
            // Get the class object corresponding to this particular sheet (e.g. Tuesday Yoga 5pm)
            var classObj = allClassesObj[sheetName];

            // Iterate through list of names
            var lastRow = sheet.getLastRow();
            for(var row = 3; row <= lastRow - 4; row++){
                // Get name and type / person category
                var nameColumn = 1;
                var typeColumn = 2;

                var name = sheet.getRange(row, nameColumn).getValue().trim();
                var typeCell = sheet.getRange(row, typeColumn);
                var type = typeCell.getValue().trim();

                // See if we have a match in our input object. Only update the non-vc (most error prone) ones
                // If we do, update the spreadsheet with correct value
                if(name !== "" && name !== STUDENT && name !== STAFF && name !== NONVC && type === NONVC){
                    if(classObj.hasOwnProperty(name)) {
                        var personObj = classObj[name];
                        if(personObj.isEmployee) { type = STAFF; }
                        else if(personObj.isStudent) { type = STUDENT; }
                        else { type = NONVC; }

                        typeCell.setValue(type);
                    }
                }
            }
        }
    }
    return "Success!";
}

/*
 * Parses the spreadsheet and adds a new page with the names of all Non-VC class participants
 *
 */
function getAllNonVC(){
    var STUDENT = 'Student';
    var STAFF = 'Staff';
    var NONVC = 'Non-VC';

    var nonVassarArray = [];

    // Iterate through all sheets/classes in the app
    var parentSheet = SpreadsheetApp.openById('1CrS9aupqZ8z76tR5aGUsncLL7BwcYzeSbEN8Ok_SFzU');
    var sheetsArray = parentSheet.getSheets();
    for(var i=0; i<sheetsArray.length; i++){
        var sheet = sheetsArray[i];
        var sheetName = sheet.getName();


        // Iterate through list of names
        var lastRow = sheet.getLastRow();
        for(var row = 3; row <= lastRow - 4; row++){
            // Get name and type
            var nameColumn = 1;
            var typeColumn = 2;
            var name = sheet.getRange(row, nameColumn).getValue().trim();
            var type = sheet.getRange(row, typeColumn).getValue().trim();

            if(type === NONVC) {
                Logger.log(name);
                nonVassarArray.push(name);
            }
        }
    }

    // Create new sheet and add information
    var nonVassarSheet = parentSheet.insertSheet("Non-VC: May need to verify");
    nonVassarSheet.insertColumnBefore(1);
    nonVassarSheet.appendRow(["Names"]);
    for(var index=0; index<nonVassarArray.length; index++){
        nonVassarSheet.appendRow([nonVassarArray[index]]);
    }

}

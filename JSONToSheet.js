/* to work with this script

1.Open your Google Sheet.

2. Go to Extensions → Apps Script.

3. Replace the existing code below: 
4. Save the script.


5. Add a Button to Your Sheet:

a. Go back to your Google Sheet.
   b. Insert a drawing or a button image (Insert → Drawing → New).
   c. After you create the drawing or button, click on it and then click the three dots in the top right corner.
   d. Select "Assign Script" and enter JSONToSheet.

6. Prepare Your Sheet:
   a. Make sure your JSON data is in a specific cell (e.g., A1).

7.Run the Script:

   a. Click the button you created. This will run the JSONToSheet function.
   b.The function will then read the JSON from the specified cell and write the data to the sheet.

*/


function JSONToSheet() {
    var sheet = SpreadsheetApp.getActiveSheet();
    var jsonString = sheet.getRange("A1").getValue();

    try {
        var data = JSON.parse(jsonString);
    } catch (e) {
        sheet.getRange("A2").setValue("Error parsing JSON");
        return;
    }

    if (typeof data !== 'object' || data === null || Array.isArray(data)) {
        sheet.getRange("A2").setValue("JSON is not an object");
        return;
    }

    var keys = Object.keys(data);
    var values = Object.values(data);

    // Clear previous data
    sheet.getRange("A2:B" + sheet.getLastRow()).clearContent();

    // Write keys in column A and values in column B
    for (var i = 0; i < keys.length; i++) {
        sheet.getRange(i + 2, 1).setValue(keys[i]); // Keys in column A
        sheet.getRange(i + 2, 2).setValue(values[i]); // Values in column B
    }
}




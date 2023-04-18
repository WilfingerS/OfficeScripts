// AUTHOR: SETH WILFINGER
// CREATED DATE: 4.13.2023
// LAST UPDATED: 4.18.2023

function main(workbook: ExcelScript.Workbook) {
    const totalSheets = workbook.getWorksheets(); // All the Sheets in the current excel sheet, from OfficeScripts API
    let data = {["Total Club"]:{BCM:0,TOTAL:0}}; // Initializes the Dictionary

    for(let i = 1; i < totalSheets.length; i ++){	// Nested For Loops to Loop through Sheets That are Imported
        let range = totalSheets[i].getUsedRange(); //get the cells that are used in the sheet, from OfficeScripts API
        let c1 = range.find("Agreement Type", { completeMatch: true, matchCase: true }).getColumnIndex(); //gets the index for the column that contains all Agreement Types for the worksheet, from OfficeScripts API
        let c2 = range.find("Membership Type", { completeMatch: true, matchCase: true }).getColumnIndex(); // same as above but for Membership type from OfficeScripts API
        let c3 = range.find("Sales Person", { completeMatch: true, matchCase: true }).getColumnIndex(); // same as above but for sales person from OfficeScripts API

        for (let r = 1; r<=range.getRowCount();r++){ // For Each Sheet Loops through the rows looking for valid information
            let at = range.getCell(r,c1).getValue().toString(); // agreeement type (New Agreement or Conversion)
            let mt = range.getCell(r, c2).getValue().toString(); // membership type (10NR, 15NR, etc..)
            let sp = range.getCell(r, c3).getValue().toString(); // sales person

            if (at.startsWith("N")){ // Only Stores Information about New Agreement 
				if (mt.includes("SP") || mt.includes("SS")) { continue; } // skips over 10SS, 10SP
                if (sp == "-") { sp = "WEB"; } // "-" is web joins/ online joins
                if (!data[sp]) { data[sp] = {BCM:0,TOTAL:0}; } // if there is no Dictionary Array then it creates one
                if (mt.startsWith("Black")) { data[sp]["BCM"] += 1; data["Total Club"]["BCM"] +=1; } // Checks if Membership is Black Card Plus or Black Card Member and tally's 
                data["Total Club"]["TOTAL"] += 1; // If no conditions meet then it must be a classic membership and tallys up the total for the club
                data[sp]["TOTAL"] += 1; // and tally's up the total for the sales person.
            }
        }
    }

    let rowC = 0; //initializes row counter
    for (let name in data){ // nested dictionary for loop 
        let cell = totalSheets[0].getCell(rowC,0); // Gets the First Cell of the Row
        cell.setValue(name); //Sets the Value to the key:String for the dictionary (Typically the name of the SalesPerson)
        let columnC = 1; // initializes column count
        for (let info in data[name]){ // another for loop through the values of data[key]
            cell = totalSheets[0].getCell(rowC,columnC); // gets the cell to the right of the last cell used
            cell.setValue(`${info}: ${data[name][info]}`); // Sets the value
            columnC++; // increments column counter
        }
        rowC++;// increments to next row
    }
    //console.log(data);
}

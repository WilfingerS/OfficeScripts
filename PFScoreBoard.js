// AUTHOR: SETH WILFINGER
// DATE: 4.13.2023
function main(workbook: ExcelScript.Workbook) {
    const totalSheets = workbook.getWorksheets();
    let data = {["Total Club"]:{BCM:0,TOTAL:0}};

    for(let i = 1; i < totalSheets.length; i ++){
        let range = totalSheets[i].getUsedRange();
        let c1 = range.find("Agreement Type", { completeMatch: true, matchCase: true }).getColumnIndex();
        let c2 = range.find("Membership Type", { completeMatch: true, matchCase: true }).getColumnIndex();
        let c3 = range.find("Sales Person", { completeMatch: true, matchCase: true }).getColumnIndex();

        for (let r = 1; r<=range.getRowCount();r++){
            let at = range.getCell(r,c1).getValue().toString();
            let mt = range.getCell(r, c2).getValue().toString();
            let sp = range.getCell(r, c3).getValue().toString();

            if (at.startsWith("N")){
                if (sp == "-") { sp = "WEB"; }
                if (!data[sp]) { data[sp] = {BCM:0,TOTAL:0}; }
                if (mt.startsWith("Black")) { data[sp]["BCM"] += 1; data["Total Club"]["BCM"] +=1; }
                data["Total Club"]["TOTAL"] += 1;
                data[sp]["TOTAL"] += 1;
            }
        }
    }

    let rowC = 0;
    for (let name in data){
        let cell = totalSheets[0].getCell(rowC,0);
        cell.setValue(name);
        let columnC = 1;
        for (let info in data[name]){
            cell = totalSheets[0].getCell(rowC,columnC);
            cell.setValue(`${info}: ${data[name][info]}`);
            columnC++;
        }
        rowC++;
    }
    //console.log(data);
}

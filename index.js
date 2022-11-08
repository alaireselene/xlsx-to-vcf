var vCardsJS = require('vcards-js');
var XLSX = require("xlsx");
const fs = require('fs');
var datastring = "";

function main() {
    // Handling data. Put your data in the source_sample.xlsx file, then rename it to source.xlsx.
    var workbook = XLSX.readFile("source.xlsx");
    var sheet = workbook.Sheets['data_table'];

    // Generating vCards.
    for (var r = 1; r !== 72; r++) {
        var vCard = vCardsJS();

        // Basic information.
        vCard.firstName = sheet['B' + r.toString()].v;
        vCard.gender = sheet['D' + r.toString()].v;
        vCard.organization = "Tổ Hỗ trợ triển khai Hệ thống iCTSV";

        // Contact information.
        vCard.workPhone = sheet['F' + r.toString()].v;
        vCard.cellPhone = sheet['E' + r.toString()].v;
        vCard.homeAddress = sheet['D' + r.toString()].v;
        vCard.url = sheet['G' + r.toString()].v;
        vCard.workEmail = sheet['A' + r.toString()].v;

        datastring += (vCard.getFormattedString() + "\n");
    }

    //Write file
    fs.writeFile('output.vcf', datastring, function (err) {
        if (err) throw err;
        console.log('Saved!');
    });
    
}
main();
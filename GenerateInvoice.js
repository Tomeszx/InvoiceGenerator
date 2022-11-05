function get_partner(){
  const data_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DATA").getRange("3:3").getValues()[0];
  const folder_PDF = DriveApp.getFolderById(data_sheet[0]).getId();
  const manual_invoice = SpreadsheetApp.openById(data_sheet[3]).getSheetByName("CN invoice - Example");
  const data_partners = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Invoices");
  const rows = data_partners.getDataRange().getValues();
  const start_time_script = new Date().getTime()

  rows.forEach(function(row, index){

    var end_time_script = (new Date().getTime() - start_time_script) / 1000;
    if (Number(end_time_script / 60) > 27){return};

    var start_time = new Date().getTime()
    var partner = row[0]
    if (partner == "" || index < 3 || row[3] != ""){return}

    //SpreadsheetApp.flush(); if (data_partners.getRange(index + 1, 6).getValue() == true){return}
    //else{data_partners.getRange(index + 1, 6).setValue(true)}; SpreadsheetApp.flush();

    manual_invoice.getRange("A3").setValue(partner);
    SpreadsheetApp.flush();

    var arr = generate_spread(data_sheet, manual_invoice, partner);
    var sheet_id = arr[0];
    var invoice_id = arr[1];

    var id_pdf = convertPDF(sheet_id);

    var url_pdf = moveFile(id_pdf, folder_PDF);

    console.log(partner, "-> Duration:",(new Date().getTime() - start_time) / 1000), "sec";
    data_partners.getRange(index + 1, 3, 1, 3).setValues([[invoice_id, "https://docs.google.com/spreadsheets/d/" + sheet_id, url_pdf]]);

    console.log(partner, "-> Duration:",(new Date().getTime() - start_time) / 1000), "sec";
  })
}

function generate_spread(data_sheet, invoice_sheet, partner) {
  const folder_sheet = DriveApp.getFolderById(data_sheet[1]);
  const template = DriveApp.getFileById(data_sheet[2]);
  const invoice_id = invoice_sheet.getRange("H4").getValue()

  var report_title = `${invoice_id} - ${partner} - CN invoice`
  try{var template_copy = template.makeCopy(report_title, folder_sheet);}
  catch{var template_copy = template.makeCopy(report_title, folder_sheet);}
  var new_template = SpreadsheetApp.openById(template_copy.getId());

  var data_dict = {}
  const data = SpreadsheetApp.openById(data_sheet[3]).getSheets();
  for(var i in data){ 
    var sheet = data[i]; 
    if (sheet.getName().includes("invoice")){ data_dict[sheet.getName()] = sheet.getRange("A:H").getValues(); break}
    else{ data_dict[sheet.getName()] = sheet.getDataRange().getValues();}
  }
  const keys = Object.keys(data_dict);
  
  for (var i in keys){
    var key = keys[i];
    var sheet_data = data_dict[key];

    // EXXPORT DATE TO 1st PAGE
    if (key.includes("invoice")){
      for (var x in sheet_data){
        var row = sheet_data[x];
        if (row.indexOf("AMOUNT") > -1){ var start_row = Number(x)}
        if (row.indexOf("Payment section") > -1){var row_2nd_page = Number(x); break;}
        new_template.getSheets()[0].getRange(Number(x) + 1, 1, 1, row.length).setValues([row])
      }

      // DELETE EMPTY ROWS
      var s = 0
      for (var x in sheet_data){
        var row = sheet_data[x];
        if (x > start_row){
          if (row[1] == ""){break}

          if (row[7] == 0){ 
            new_template.getSheets()[0].deleteRow(Number(x) + 1 - s)
            s += 1
          }
        }
      }
      // EXXPORT 2nd PAGE
      for (var x in sheet_data){
        var row = sheet_data[x];
        if (row[0] == "Orders"){var start_row = Number(x); break;}

        if(x > row_2nd_page){ 
          new_template.getSheets()[1].getRange(Number(x) - row_2nd_page + 1, 1, 1, row.length).setValues([row])
        }
      }

      // EXXPORT 3rd - x PAGE
      var delete_sheet = 0
      for (sh in new_template.getSheets()){
        var sh_index = sh - delete_sheet
        var sheet = new_template.getSheets()[sh_index]
        if (sh < 2){continue}
        
        var start_row = invoice_sheet.createTextFinder(sheet.getName()).matchEntireCell(true).matchCase(true).findAll()[1].getRow() - 1;
        try{var end_row = invoice_sheet.createTextFinder(new_template.getSheets()[Number(sh_index) + 1].getName()).matchEntireCell(true).matchCase(true).findAll()[1].getRow() - 1;}
        catch{end_row = invoice_sheet.getLastRow()}

        //EXPORT 3rd PAGE
        var data_source = []
        for (var x = start_row; end_row > x; x++){
          var row = sheet_data[x];
          if(row.indexOf("#N/A") > -1 || row.indexOf("#REF!") > -1){break}

          if(row[7] != ""){
            data_source.push(row)

            if (row[7] != "" && row[0] == ""){break}
          }
        }
        if(data_source.length > 1){
          sheet.deleteRows(data_source.length + 1, sheet.getLastRow() - data_source.length - 1); //delete empty rows
          sheet.getRange(2, 1, data_source.length, data_source[0].length).setValues(data_source); // import data to sheet
        }
        else {
          new_template.deleteSheet(sheet)
          delete_sheet += 1
        }
      }
    }
  }  
  SpreadsheetApp.flush();
  return [new_template.getId(), invoice_id]
}

function convertPDF(sheet_id) {
  const spread = DriveApp.getFileById(sheet_id)

  const docblob = spread.getAs('application/pdf');
  /* Add the PDF extension */
  docblob.setName(spread.getName() + ".pdf");
  const file = DriveApp.createFile(docblob);

  return file.getId();
}


function moveFile(fileId, toFolderId) {
   const file = DriveApp.getFileById(fileId);
   const source_folder = DriveApp.getFileById(fileId).getParents().next();
   const folder = DriveApp.getFolderById(toFolderId)

   folder.addFile(file);
   source_folder.removeFile(file);

  return file.getUrl();
}

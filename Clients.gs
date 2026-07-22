/**
 * ======================================================
 * Clients.gs
 * Client Management Module
 * ======================================================
 */

function generateClientId_() {
  const sh = SpreadsheetApp.getActive().getSheetByName("Clients");

  if (!sh || sh.getLastRow() < 2) {
    return "CL000001";
  }

  const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().flat();

  let max = 0;

  ids.forEach(id => {
    if (!id) return;

    const num = parseInt(String(id).replace("CL", ""), 10);

    if (!isNaN(num) && num > max) {
      max = num;
    }
  });

  return "CL" + Utilities.formatString("%06d", max + 1);
}

function addClient(clientName, officeName) {

  const sh = SpreadsheetApp.getActive().getSheetByName("Clients");

  const clientId = generateClientId_();

  sh.appendRow([
    clientId,
    clientName,
    officeName,
    "",
    "Active"
  ]);

  return clientId;

}

function getClientByOffice(officeName) {

  const sh = SpreadsheetApp.getActive().getSheetByName("Clients");

  if (sh.getLastRow() < 2) return null;

  const data = sh.getRange(2,1,sh.getLastRow()-1,5).getValues();

  for (let i = 0; i < data.length; i++) {

    if (String(data[i][2]).trim() === String(officeName).trim()) {

      return {
        row:i+2,
        clientId:data[i][0],
        clientName:data[i][1],
        office:data[i][2],
        docId:data[i][3],
        status:data[i][4]
      };

    }

  }

  return null;

}

function getClientById(clientId){

  const sh = SpreadsheetApp.getActive().getSheetByName("Clients");

  if(sh.getLastRow()<2) return null;

  const data = sh.getRange(2,1,sh.getLastRow()-1,5).getValues();

  for(let i=0;i<data.length;i++){

    if(data[i][0]==clientId){

      return{
        row:i+2,
        clientId:data[i][0],
        clientName:data[i][1],
        office:data[i][2],
        docId:data[i][3],
        status:data[i][4]
      };

    }

  }

  return null;

}

function updateClientDocId(clientId, docId){

  const client = getClientById(clientId);

  if(!client) return false;

  SpreadsheetApp.getActive()
    .getSheetByName("Clients")
    .getRange(client.row,4)
    .setValue(docId);

  return true;

}

function getActiveClients(){

  const sh = SpreadsheetApp.getActive().getSheetByName("Clients");

  if(sh.getLastRow()<2) return [];

  const data = sh.getRange(2,1,sh.getLastRow()-1,5).getValues();

  return data.filter(r => String(r[4]).toLowerCase()=="active");

}

function deactivateClient(clientId){

  const client = getClientById(clientId);

  if(!client) return false;

  SpreadsheetApp.getActive()
    .getSheetByName("Clients")
    .getRange(client.row,5)
    .setValue("Inactive");

  return true;

}

function activateClient(clientId){

  const client = getClientById(clientId);

  if(!client) return false;

  SpreadsheetApp.getActive()
    .getSheetByName("Clients")
    .getRange(client.row,5)
    .setValue("Active");

  return true;

}

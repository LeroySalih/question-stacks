import { ConsoleTranscriptLogger } from 'botbuilder';
import xlsx from 'node-xlsx';

const __dirname = ""

const parseSheet = (ds) => {

  const db = []
  const pupils = {}

  const YEAR = 1
  const NAME = 2
  const EMAIL = 4
  const UPN = 11
  const PARENT_EMAIL = "14"

  let currentName = ""
  let currentEmail = ""
  let currentYear = ""
  let currentUPN = ""
  let currentParentEmail = ""

  // Build the Populated Array
  Object.values(ds).forEach((row, i) => {


    let pupilYear = row[YEAR]
    let pupilName = row[NAME];
    let pupilEmail = row[EMAIL];
    let pupilUPN = row[UPN];
    let parentEmail = row[PARENT_EMAIL];

    if (pupilName != currentName && pupilName != undefined) {
      currentYear = pupilYear
      currentName = pupilName 
      currentEmail = pupilEmail
      currentUPN = pupilUPN   
      currentParentEmail =  parentEmail
    }

    if (parentEmail != currentParentEmail && parentEmail != undefined){
      currentParentEmail = parentEmail
    }

    db.push ({name: currentName, email: currentEmail, upn: currentUPN, year: currentYear, parentEmail: currentParentEmail});

    
    
  });

  // Convert Array to dictionary
  db.forEach(pupil => {
    if (pupils[pupil.upn] == undefined){
      pupils[pupil.upn] = pupil 
      pupils[pupil.upn]['emails'] = []
    }

    pupils[pupil.upn]['emails'].push (pupil.parentEmail);
  })

  return pupils;
}

const contactDB = (fileName) => {

  const workSheetsFromFile = xlsx.parse(fileName);

  const pupilDataSheet = workSheetsFromFile[0]

  return  parseSheet(pupilDataSheet.data.slice(1, pupilDataSheet.data.length));
}

export {contactDB}

/*
const result = createDatabase(`./pupil-db-2021.xlsx`)
console.log(result);
*/

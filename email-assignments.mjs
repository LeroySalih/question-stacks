import {moveFile} from 'move-file';
import fs from 'fs'
import xlsx from 'node-xlsx';
import {contactDB} from './contact-db.mjs';
import nodeoutlook from 'nodejs-nodemailer-outlook'
import { argv } from 'process';
import readline from 'readline';

const cleanUp = () => {
  fs.readdirSync('./assignments-in').forEach(file => moveFile(`./assignments-in/${file}`, `./assignments-out/${file}`))
}

// email, task name, score, feedback
const readFile = (fileName) => {

  const TASK_NAME = 3
  const EMAIL = 2
  const workSheetsFromFile = xlsx.parse(fileName);
  const workSheet = workSheetsFromFile[0].data;
  const pupilsOnly = workSheet.slice(1, workSheet.length);

  const taskFeedback = {}

  pupilsOnly.forEach(row => {
    taskFeedback[row[EMAIL]] = [ workSheet[0][3], ...row]
  })
  
  return taskFeedback

}

const consolidateTasks = (tasks, assignment) => {

  Object.keys(assignment).forEach(pupil => {
      
    if (tasks[pupil] == undefined) {
      tasks[pupil] = [];
    }

    tasks[pupil].push(assignment[pupil]) ;

  });

  return tasks;

}

const getParentEmails = (cdb, email) => {
  const pupilRecords = Object.values(cdb)
          .filter(row => row.email == email)
          .map(row => row.emails)
          .join()
          .replaceAll(',', '; ')
  return pupilRecords;
}

const sleep = (ms) => {
  return new Promise(res => setTimeout(res, ms))
}

async function emailPupils (pupils, cdb, details) {

  for (const key of Object.keys(pupils)){

    const pupil = pupils[key]
    const parentsEmail = getParentEmails(cdb, pupil[0][3]);

    if (parentsEmail == ''){
      console.log(`${pupil[0][1]} - no parent email found (${pupil[0][3]})`);
    } else {
      console.log(`Sending email for ${pupil[0][1]} to ${parentsEmail}`);

      var result = await sendEmail(pupil, parentsEmail, details, parentsEmail);
    }

  }
  

}

const displayPupils = (pupils, cdb) => {
  const TITLE = 0
  const FIRST_NAME = 1
  const FAMILY_NAME = 2
  const MARKS = 4
  const FEEDBACK = 6
  Object.keys(pupils).forEach(key => {
    const parentsEmail = getParentEmails(cdb, key)
    console.log( pupils[key][0][FIRST_NAME], pupils[key][0][FAMILY_NAME], key, `[${parentsEmail}]`)
    
    pupils[key].forEach(task => console.log(task[TITLE], `\t${task[MARKS]}%`, task[FEEDBACK]));
    console.log('\n')
  })

  
}

const emailBody = (pupil, details, parentsEmail) => {
  const TITLE = 0
  const FIRST_NAME = 1
  const FAMILY_NAME = 2
  const MARKS = 4
  const FEEDBACK = 6

  const tasks = pupil.map(task => `<tr><td>${task[TITLE]}&nbsp;</td><td>${task[MARKS]}%&nbsp;</td><td>${task[FEEDBACK]}</td></tr>`)
  
  const parentsEmailText = !details.live ? `<p>${parentsEmail}</p>` : '<p></p>'

  return `<p>Dear Parents of ${pupil[0][FIRST_NAME]} ${pupil[0][FAMILY_NAME]}</p>
  ${parentsEmailText}
  <p>Please find below the summary of work recently completed by ${pupil[0][FIRST_NAME]} in his maths class.</p>

  <table>
  ${tasks}
  </table>
  
  <p>If you have any questions or concerns, please do not hesitate to contact me on leroysalih@bisak.org.</p>
  <p>If you would like to stop receiving these updates, please notify me by email.</p>
  <p>Yours Sincerely,</p>
  <p>Mr Leroy Salih</p>
  `
}

const sendEmail = (pupil, parentsEmail, details) => {
  return new Promise((res, rej) => {

    nodeoutlook.sendEmail({
      auth: {
          user: details.signIn,
          pass: details.password
      },
      from: details.signIn,
      to: details.live ? parentsEmail : 'sleroy@bisak.org',
      // to: details.live ? 'sleroy@bisak.org' : 'sleroy@bisak.org',
      subject: ` Yr10 Maths: Feedback for ${pupil[0][1]} ${pupil[0][2]}`,
      html: emailBody(pupil,details, parentsEmail),
      text: 'This is text version!',
      replyTo: 'sleroy@bisak.org',
      onError: (e) => { rej(e) },
      onSuccess: (i) => {res(`Success ${pupil[0][1]} ${pupil[0][2]}`)}
    },
    
    );
  })
  

}


const getSignIn = () => {

  return new Promise ((res, rej) => {

    var rl = readline.createInterface({
      input: process.stdin,
      output: process.stdout
    });
    
    rl.stdoutMuted = false;
    
    rl.question('Sign In: ', function(signIn) {
      rl.close();
      res(signIn);
    });
    
    rl._writeToOutput = function _writeToOutput(stringToWrite) {
      if (rl.stdoutMuted)
        rl.output.write("*");
      else
        rl.output.write(stringToWrite);
    };

  })

  

}

const getClassName = () => {

  return new Promise ((res, rej) => {

    var rl = readline.createInterface({
      input: process.stdin,
      output: process.stdout
    });
    
    rl.stdoutMuted = false;
    
    rl.question('Class name: ', function(signIn) {
      rl.close();
      res(signIn);
    });
    
    rl._writeToOutput = function _writeToOutput(stringToWrite) {
      if (rl.stdoutMuted)
        rl.output.write("*");
      else
        rl.output.write(stringToWrite);
    };

  })

  

}

const getIsLive = () => {

  return new Promise ((res, rej) => {

    var rl = readline.createInterface({
      input: process.stdin,
      output: process.stdout
    });
    
    rl.stdoutMuted = false;
    
    rl.question('Live (y/n): ', function(signIn) {
      rl.close();
      res(signIn);
    });
    
    rl._writeToOutput = function _writeToOutput(stringToWrite) {
      if (rl.stdoutMuted)
        rl.output.write("*");
      else
        rl.output.write(stringToWrite);
    };

  })
}

const getPassword = () => {

  return new Promise ((res, rej) => {

    var rl = readline.createInterface({
      input: process.stdin,
      output: process.stdout
    });
    
    rl.stdoutMuted = true;
    
    rl.question('Password: ', function(password) {
      rl.close();
      res(password);
    });
    
    rl._writeToOutput = function _writeToOutput(stringToWrite) {
      if (rl.stdoutMuted)
        rl.output.write("*");
      else
        rl.output.write(stringToWrite);
    };

  })

  

}



const getDetails = async () => {
  const signIn = await getSignIn();
  const password = await getPassword();
  const className = await getClassName();
  const isLive = await getIsLive()
  return {signIn, password, className, live: isLive == 'y'}
}

const main = async () => {


  const details = await getDetails();

  const tasks = {}

  const cdb = contactDB(`./pupil-db-2021.xlsx`);

  fs.readdirSync('./assignments-in')
    .sort((a, b) => a > b ? 1 : -1)
    .forEach(file => {
      console.log(`Reading file ${file}`)
      const result = readFile(`./assignments-in/${file}`);
    
    consolidateTasks (tasks, result);

    
    moveFile(`./assignments-in/${file}`, `./assignments-out/${file}`);
  }
  );

  // displayPupils (tasks, cdb);
  emailPupils (tasks, cdb, details);
}

main()




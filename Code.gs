const DB_ID = "1dBO8ThI7FEKb24D9sPVWokfXLuWUx5aCQvisrT9wBvI";
const USERS_SHEET = "NOA-PROP";
const ADMIN_EMAIL = "mcddatamanagement.cog@megaworld-lifestyle.com";

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('NOA Monitoring App')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function expandUserProperties(userProperties, allDbProperties) {
  let expandedList = [];
  
  const dbPropsMap = allDbProperties.map(p => ({
    original: p,
    clean: p.toString().trim().toLowerCase()
  }));

  userProperties.forEach(userProp => {
    const cleanUserProp = userProp.toString().trim();
    expandedList.push(cleanUserProp); 

    const match = cleanUserProp.match(/^[A-Z0-9]{2,5}\s+-\s+(.*)$/i);
    
    if (match) {
      const rootName = match[1].trim().toLowerCase(); 
      
      dbPropsMap.forEach(dbItem => {
        if (dbItem.clean === rootName) {
          expandedList.push(dbItem.original);
        }
        else if (dbItem.clean === cleanUserProp.toLowerCase()) {
          expandedList.push(dbItem.original);
        }
      });
    }
  });

  return [...new Set(expandedList)];
}

function getPropertyOptions() {

  const data = getData();
  const rawProperties = [...new Set(data.map(d => d.property ? d.property.toString().trim() : "").filter(Boolean))];
  const codedProps = [];
  const plainProps = [];
  
  const codePattern = /^[A-Z0-9]{2,5}\s+-\s+/i; 

  rawProperties.forEach(p => {
    if (codePattern.test(p)) {
      codedProps.push(p);
    } else {
      plainProps.push(p);
    }
  });

  const codedRoots = new Set();
  
  codedProps.forEach(p => {
    const parts = p.split(/\s+-\s+/); 
    if (parts.length > 1) {
      const rootName = parts.slice(1).join(" - ").trim().toLowerCase();
      codedRoots.add(rootName);
    }
  });
  
  const finalOptions = [...codedProps];
  
  plainProps.forEach(plain => {
    const cleanPlain = plain.trim().toLowerCase();
    
    if (!codedRoots.has(cleanPlain)) {
      finalOptions.push(plain);
    }
  });

  return finalOptions.sort();
}

function getUsersSheet() {
  const ss = SpreadsheetApp.openById(DB_ID);
  let sheet = ss.getSheetByName(USERS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(USERS_SHEET);
    sheet.appendRow([
      "Timestamp",       // A
      "Name",            // B
      "Email",           // C
      "Password",        // D
      "Properties",      // E
      "SecQuestion",     // F
      "SecAnswer",       // G
      "SessionToken",    // H
      "TokenExpiry",     // I
      "Status of User",  // J
      "EmailSent",       // K
      "Status of Employee", // L
      "AlertSeen"        // M 
    ]);
  }
  return sheet;
}

function hashString(str) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, str);
  return raw.map(function(e) {return ("0" + (e < 0 ? e + 256 : e).toString(16)).slice(-2)}).join("");
}

function sanitize(str) {
  if (typeof str !== 'string') return str;
  return str.replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function registerUser(form) {
  const sheet = getUsersSheet();
  const data = sheet.getDataRange().getDisplayValues();
  const email = sanitize(form.email).toLowerCase();
  
  const exists = data.some(r => r[2].toString().toLowerCase() === email);
  if (exists) return { success: false, message: "Email already registered." };

  const passHash = hashString(form.password);
  const secAnsHash = hashString(form.secAnswer.toLowerCase().trim());
  const timestamp = new Date();
  
  if (!form.properties || form.properties.length === 0) {
     return { success: false, message: "Please select at least one property." };
  }

  form.properties.forEach(prop => {
    sheet.appendRow([
      timestamp,
      sanitize(form.name),
      email,
      passHash,
      prop,             // Col E: Single Property
      form.secQuestion,
      secAnsHash,
      "", "",           // Token slots empty
      "Pending"         // Col J: Default Status
    ]);
  });
  
  try {
    MailApp.sendEmail({
      to: ADMIN_EMAIL,
      subject: "NOA App: New User Registration - " + form.name,
      htmlBody: `
        <h3>New Account Registration</h3>
        <p><strong>Name:</strong> ${sanitize(form.name)}</p>
        <p><strong>Email:</strong> ${email}</p>
        <p><strong>Properties Requested:</strong> ${form.properties.join(", ")}</p>
        <p>Please check the spreadsheet <b>"${USERS_SHEET}"</b> to approve or reject specific properties.</p>
        <p>Link: <a href="https://docs.google.com/spreadsheets/d/${DB_ID}/edit">Open Spreadsheet</a></p>
      `
    });
  } catch(e) {
    console.log("Email error: " + e.toString());
  }
  
  // Updated Message for User
  return { success: true, message: "Registration successful! We will notify you via email once your account has been approved by the Admin." };
}

function loginUser(credentials) {
  const sheet = getUsersSheet();
  const data = sheet.getDataRange().getDisplayValues();
  const email = credentials.email.toLowerCase();
  const passHash = hashString(credentials.password);

  let userRows = []; 
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][2].toString().toLowerCase() === email && data[i][3] === passHash) {
      const employeeStatus = data[i][11] ? data[i][11].toString().trim().toUpperCase() : "";
      if (employeeStatus === "RESIGNED") {
        return { success: false, message: "ACCESS DENIED: This account is deactivated. The User is marked as 'RESIGNED'" };
      }
      userRows.push({ index: i + 1, data: data[i] });
    }
  }

  if (userRows.length === 0) return { success: false, message: "Invalid email or password." };

  const approvedRows = userRows.filter(item => {
    const status = item.data[9].toString().trim().toUpperCase();
    return status.includes("APPROVE"); 
  });

  if (approvedRows.length === 0) {
    return { success: false, message: "Your account is still PENDING approval or has been rejected." };
  }

  const assignedProps = approvedRows.map(item => item.data[4]); 
  
  const allDbData = getData(); 
  const allDbProps = [...new Set(allDbData.map(d => d.property).filter(Boolean))];

  const finalAccessList = expandUserProperties(assignedProps, allDbProps);

  const token = Utilities.getUuid();
  const expiry = new Date().getTime() + (3600 * 1000 * 24); 
  
  userRows.forEach(item => {
    sheet.getRange(item.index, 8).setValue(token); 
    sheet.getRange(item.index, 9).setValue(expiry); 
  });

  const newAlerts = checkUserAlerts(email);

  return { 
    success: true, 
    token: token,
    alerts: newAlerts, // Pass alerts to frontend
    user: {
      name: approvedRows[0].data[1], 
      email: approvedRows[0].data[2],
      properties: finalAccessList
    }
  };
}

function verifySession(token) {
  if(!token) return { valid: false };
  
  const sheet = getUsersSheet();
  const data = sheet.getDataRange().getDisplayValues();
  const now = new Date().getTime();
  
  let validUserEmail = null;
  let userData = null;

  for (let i = 1; i < data.length; i++) {
    if (data[i][7] === token) {
      const expiry = parseInt(data[i][8]);
      const employeeStatus = data[i][11] ? data[i][11].toString().trim().toUpperCase() : "";
      
      if (employeeStatus === "RESIGNED") return { valid: false };

      if (now < expiry) {
        validUserEmail = data[i][2]; 
        userData = data[i];
        break;
      }
    }
  }

  if (!validUserEmail) return { valid: false };

  const assignedProps = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === validUserEmail) {
       const status = data[i][9].toString().trim().toUpperCase();
       const empStatus = data[i][11] ? data[i][11].toString().trim().toUpperCase() : "";
       
       if (status.includes("APPROVE") && empStatus !== "RESIGNED") {
         assignedProps.push(data[i][4]);
       }
    }
  }

  if (assignedProps.length === 0) return { valid: false };

  const allDbData = getData();
  const allDbProps = [...new Set(allDbData.map(d => d.property).filter(Boolean))];

  const finalAccessList = expandUserProperties(assignedProps, allDbProps);

  const newAlerts = checkUserAlerts(validUserEmail);

  return { 
    valid: true, 
    alerts: newAlerts, // Pass alerts to frontend
    user: {
      name: userData[1],
      email: userData[2],
      properties: finalAccessList
    }
  };
}

function recoverPassword(email, secQuestion, secAnswer, newPassword) {
  const sheet = getUsersSheet();
  const data = sheet.getDataRange().getDisplayValues();
  const targetEmail = email.toLowerCase();
  
  let userRowIndices = [];
  let storedQuestion = "";
  let storedAnsHash = "";

  for (let i = 1; i < data.length; i++) {
    if (data[i][2].toString().toLowerCase() === targetEmail) {
      userRowIndices.push(i + 1);
      // Capture security details from the first occurrence
      if (!storedQuestion) {
        storedQuestion = data[i][5];
        storedAnsHash = data[i][6];
      }
    }
  }

  if (userRowIndices.length === 0) return { success: false, message: "Email not found." };
  
  if (!secAnswer && !newPassword) {
    return { success: true, step: "question", question: storedQuestion };
  }

  const inputAnsHash = hashString(secAnswer.toLowerCase().trim());
  if (inputAnsHash !== storedAnsHash) {
    return { success: false, message: "Incorrect security answer." };
  }

  const newPassHash = hashString(newPassword);
  
  userRowIndices.forEach(idx => {
    sheet.getRange(idx, 4).setValue(newPassHash); // Col D
  });

  return { success: true, step: "done", message: "Password reset successful." };
}

function getData() {
  const ss = SpreadsheetApp.openById("1m7bOgXL4UJHUd0euaYMguAakhhuElRPxBqZ6R_GTRj4");
  const ws = ss.getSheetByName("Form Responses 1");
  const lastRow = ws.getLastRow();
  
  const ssAction = SpreadsheetApp.openById("1Aa2O2XVinhL7x2zaBlCgpAcv7IcLAcqeUgGVfRNidzw");
  const wsAction = ssAction.getSheetByName("LE UNIFICADO");
  
  let validRefsInUnificado = new Set();
  const lastRowAction = wsAction.getLastRow();
  
  if (lastRowAction > 1) {
    const actionData = wsAction.getRange(2, 1, lastRowAction - 1).getDisplayValues(); 
    actionData.forEach(r => {
      if(r[0]) validRefsInUnificado.add(r[0].trim());
    });
  }
  
  if (lastRow < 6) return [];
  
  const dataRange = ws.getRange(6, 1, lastRow - 5, 31);
  const values = dataRange.getDisplayValues(); 
  
  const mappedData = values.map((row) => {
    const timestamp = row[0].trim();
    const refNo = row[23].trim(); 
    
    if (timestamp === "" && refNo === "") {
      return null; 
    }

    const colY_LinkOrStatus = row[24]; 
    const colAE_ActionStatus = row[30]; 

    let status = "Pending";

    const colYString = colY_LinkOrStatus.toString().trim();
    const colYLower = colYString.toLowerCase();

    if (colYLower.includes("disapproved")) {
      status = "Disapproved";
    } else if (colYString !== "") {
      
      const isRefValid = validRefsInUnificado.has(refNo);
      const isActionDoneInAE = colAE_ActionStatus && colAE_ActionStatus.toString().trim() !== "";

      if (isRefValid && isActionDoneInAE) {
        status = "Validated with Action Done";
      } else {
        status = "Validated with Pending Action";
      }

    } else {
      status = "Pending";
    }

    return {
      timestamp: row[0],
      releasedBy: row[1],
      property: row[2],
      payor: row[3],
      msp: row[4],
      serviceType: row[5],
      startDate: row[7],
      endDate: row[8],
      kindOfNoa: row[9],
      sector: row[12],
      refNo: row[23],
      pdfLink: row[24],
      status: status
    };
  })
  .filter(item => item !== null);

  return mappedData;
}

function processEmailQueue() {
  const ss = SpreadsheetApp.openById(DB_ID);
  const sheet = ss.getSheetByName(USERS_SHEET);
  const dataRange = sheet.getDataRange();
  const data = dataRange.getDisplayValues();
  
  let emailQueue = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const name = row[1];       // Col B
    const email = row[2];      // Col C
    const property = row[4];   // Col E
    const status = row[9].toString().toUpperCase();     // Col J
    const emailSent = row[10]; // Col K

    if ((status.includes("APPROVE") || status.includes("REJECT")) && emailSent !== "DONE" && email.includes("@")) {
      
      if (!emailQueue[email]) {
        emailQueue[email] = {
          name: name,
          approvedProps: [],
          rejectedProps: [],
          rowIndices: []
        };
      }
      
      if (status.includes("APPROVE")) {
        emailQueue[email].approvedProps.push(property);
      } else {
        emailQueue[email].rejectedProps.push(property);
      }
      
      emailQueue[email].rowIndices.push(i + 1);
    }
  }

  if (Object.keys(emailQueue).length === 0) return;
  
  for (const email in emailQueue) {
    const userData = emailQueue[email];
    
    let messageBody = `<p>Hi <strong>${userData.name}</strong>,</p><p>Here is the status update on your requested properties:</p>`;
    
    if (userData.approvedProps.length > 0) {
      messageBody += `<h3 style="color: green;">✅ APPROVED ACCESS:</h3><ul>`;
      userData.approvedProps.forEach(p => messageBody += `<li>${p}</li>`);
      messageBody += `</ul>`;
    }

    if (userData.rejectedProps.length > 0) {
      messageBody += `<h3 style="color: red;">❌ DISAPPROVED / REJECTED:</h3><ul>`;
      userData.rejectedProps.forEach(p => messageBody += `<li>${p}</li>`);
      messageBody += `</ul><p>Please contact admin for more details regarding disapproved items.</p>`;
    }

    messageBody += `<br><hr><p style="font-size: 12px; color: #666;">Admin Team - NOA Monitoring</p>`;

    try {
      MailApp.sendEmail({
        to: email,
        subject: "Updates on your NOA Account Access",
        htmlBody: `<div style="font-family: Arial, sans-serif; padding: 20px; border: 1px solid #ddd;">${messageBody}</div>`
      });
      
      userData.rowIndices.forEach(rowIndex => {
        sheet.getRange(rowIndex, 11).setValue("DONE"); 
      });
      
      console.log(`Emailed update to ${email}`);
      
    } catch (e) {
      console.error(`Failed to email ${email}`);
    }
  }
}

function requestAdditionalProperties(token, newProperties) {
  const session = verifySession(token);
  if (!session.valid) return { success: false, message: "Session expired. Please login again." };
  
  const userEmail = session.user.email;
  const userName = session.user.name;
  
  const sheet = getUsersSheet();
  const data = sheet.getDataRange().getDisplayValues();
  const timestamp = new Date();
  
  let userStaticData = null;
  
  const existingPropStatus = {};
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][2].toString().toLowerCase() === userEmail.toLowerCase()) {
      if (!userStaticData) {
        userStaticData = {
          passHash: data[i][3],    // Col D
          secQuestion: data[i][5], // Col F
          secAnsHash: data[i][6]   // Col G
        };
      }
      
      const propName = data[i][4];
      const status = data[i][9].toString().toUpperCase(); // Col J (Status)
      existingPropStatus[propName] = status; 
    }
  }
  
  if (!userStaticData) return { success: false, message: "User record not found." };
  
  let addedCount = 0;
  let requestedPropsList = [];
  let ignoredProps = []; 

  newProperties.forEach(prop => {
    const currentStatus = existingPropStatus[prop];
    
    if (!currentStatus || currentStatus.includes("REJECT") || currentStatus.includes("DISAPPROVED")) {
      
      sheet.appendRow([
        timestamp,
        userName,
        userEmail,
        userStaticData.passHash, 
        prop,                    
        userStaticData.secQuestion,
        userStaticData.secAnsHash,
        token,                   
        "",                      
        "Pending",             
        "",                     
        "",                     
        ""                     
      ]);
      requestedPropsList.push(prop);
      addedCount++;
      
    } else {
      ignoredProps.push(prop);
    }
  });
  
  if (addedCount === 0) {
    if (ignoredProps.length > 0) {
      return { success: false, message: "Request skipped. You already have pending or approved access for these properties." };
    }
    return { success: false, message: "No new properties selected." };
  }
  
  try {
    MailApp.sendEmail({
      to: ADMIN_EMAIL,
      subject: "NOA App: Additional Property Request - " + userName,
      htmlBody: `
        <h3>Existing User Requesting New Access</h3>
        <p><strong>Name:</strong> ${userName}</p>
        <p><strong>Email:</strong> ${userEmail}</p>
        <p><strong>New Properties Requested:</strong> ${requestedPropsList.join(", ")}</p>
        <hr>
        <p>Please check the spreadsheet <b>"${USERS_SHEET}"</b>. Set status to <b>APPROVE</b> or <b>REJECT</b>.</p>
        <p><a href="https://docs.google.com/spreadsheets/d/${DB_ID}/edit">Open Spreadsheet</a></p>
      `
    });
  } catch(e) {
    console.log("Admin email error: " + e.toString());
  }
  
  return { success: true, message: "Request submitted successfully! Admin will review your request." };
}

function checkUserAlerts(email) {
  const sheet = getUsersSheet();
  const data = sheet.getDataRange().getDisplayValues();
  const targetEmail = email.toLowerCase();
  
  let alerts = {
    approved: [],
    rejected: []
  };
  
  let rowsToUpdate = [];

  for (let i = 1; i < data.length; i++) {
    const rowEmail = data[i][2].toString().toLowerCase();
    
    if (rowEmail === targetEmail) {
      const property = data[i][4];      // Col E
      const status = data[i][9].toString().toUpperCase(); // Col J
      const alertSeen = data[i][12] ? data[i][12].toString().toUpperCase() : ""; // Col M (AlertSeen)

      if ((status.includes("APPROVE") || status.includes("REJECT")) && alertSeen !== "SEEN") {
        
        if (status.includes("APPROVE")) {
          alerts.approved.push(property);
        } else {
          alerts.rejected.push(property);
        }
        
        rowsToUpdate.push(i + 1);
      }
    }
  }

  if (rowsToUpdate.length > 0) {
    rowsToUpdate.forEach(rowIndex => {
      sheet.getRange(rowIndex, 13).setValue("SEEN");
    });
  }

  return alerts;
}

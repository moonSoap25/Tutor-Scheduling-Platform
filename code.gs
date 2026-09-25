const SHEET_ID = '10ifKBszw6Ahj9HEmvwUHoYnER9DqNGwErSklYxJRyEQ';
const ADMIN_EMAIL = 'hana.sapunxhiu@gmail.com';

function doGet(e) {
  const page = e.parameter.page;

  if (page === 'tutor') {
    return HtmlService.createHtmlOutputFromFile('tutor-form')
      .setTitle('Tutor Registration');
  }
  if (page === 'student') {
    return HtmlService.createHtmlOutputFromFile('student-form')
      .setTitle('Student Sign-Up');
  }
  if (page === 'admin') {
    return HtmlService.createHtmlOutputFromFile('admin')
      .setTitle('Admin Panel');
  }
  if (page === 'tutor-profile') {
    return HtmlService.createHtmlOutputFromFile('tutor-profile')
      .setTitle('Tutor Profile');
  }

  return HtmlService.createHtmlOutput('<p>Page not found.</p>');
}

function getSpreadsheet() {
  return SpreadsheetApp.openById(SHEET_ID);
}

// ─── TUTOR FUNCTIONS ───────────────────────────────────────

function saveTutor(data) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Tutors');
  const existing = sheet.getDataRange().getValues();

  // Check for duplicate email
  for (let i = 1; i < existing.length; i++) {
    if (existing[i][2].toLowerCase() === data.email.toLowerCase()) {
      return { 
        success: false, 
        message: 'This email is already registered. Use the profile page to update your information.'
      };
    }
  }

  sheet.appendRow([
    data.name,
    data.grade,
    data.email,
    data.subjects.join(', '),
    data.times.join(', '),
    new Date().toLocaleDateString()
  ]);

  return { success: true };
}

function getTutorByEmail(email) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Tutors');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][2].toLowerCase() === email.toLowerCase()) {
      return {
        rowIndex: i + 1,
        name: data[i][0],
        grade: data[i][1],
        email: data[i][2],
        subjects: data[i][3] ? data[i][3].split(', ') : [],
        times: data[i][4] ? data[i][4].split(', ') : []
      };
    }
  }
  return null;
}

function updateTutor(data) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Tutors');

  sheet.getRange(data.rowIndex, 1, 1, 6).setValues([[
    data.name,
    data.grade,
    data.email,
    data.subjects.join(', '),
    data.times.join(', '),
    new Date().toLocaleDateString()
  ]]);

  try {
    GmailApp.sendEmail(
      ADMIN_EMAIL,
      'Tutor profile updated — ' + data.name,
      `A tutor has updated their profile.\n\nName: ${data.name}\nGrade: ${data.grade}\nEmail: ${data.email}\nSubjects: ${data.subjects.join(', ')}\nAvailable timings: ${data.times.join(', ')}\n\nThis change is now live in the Tutors sheet.`
    );
  } catch(e) {
    Logger.log('Email failed: ' + e.message);
  }

  return { success: true };
}

// ─── STUDENT FORM FUNCTIONS ────────────────────────────────

function getAvailableSubjects() {
  const ss = getSpreadsheet();
  const tutorSheet = ss.getSheetByName('Tutors');
  const matchSheet = ss.getSheetByName('Matches');
  const tutorData = tutorSheet.getDataRange().getValues();
  const matchData = matchSheet.getDataRange().getValues();

  const booked = {};
  matchData.slice(1).forEach(row => {
    const tutorEmail = row[5];
    const time = row[7];
    if (tutorEmail && time) {
      if (!booked[tutorEmail]) booked[tutorEmail] = [];
      booked[tutorEmail].push(time);
    }
  });

  const availableSubjects = new Set();

  tutorData.slice(1).forEach(row => {
    const tutorEmail = row[2];
    const subjects = row[3] ? row[3].split(', ') : [];
    const times = row[4] ? row[4].split(', ') : [];
    const bookedTimes = booked[tutorEmail] || [];

    // This tutor still has at least one free slot
    const hasFreeSlot = times.some(t => !bookedTimes.includes(t));

    if (hasFreeSlot) {
      subjects.forEach(s => availableSubjects.add(s));
    }
  });

  return [...availableSubjects].sort();
}

function getTimesForSubject(subject) {
  const ss = getSpreadsheet();
  const tutorSheet = ss.getSheetByName('Tutors');
  const matchSheet = ss.getSheetByName('Matches');
  const tutorData = tutorSheet.getDataRange().getValues();
  const matchData = matchSheet.getDataRange().getValues();

  // Build booked slots per tutor
  const booked = {};
  matchData.slice(1).forEach(row => {
    const tutorEmail = row[5];
    const time = row[7];
    if (tutorEmail && time) {
      if (!booked[tutorEmail]) booked[tutorEmail] = [];
      booked[tutorEmail].push(time);
    }
  });

  // A timeslot is available if at least one tutor
  // who teaches this subject still has it free
  const availableTimes = new Set();

  tutorData.slice(1).forEach(row => {
    const tutorEmail = row[2];
    const subjects = row[3] ? row[3].split(', ') : [];
    const times = row[4] ? row[4].split(', ') : [];
    const bookedTimes = booked[tutorEmail] || [];

    if (!subjects.includes(subject)) return;

    times.forEach(t => {
      if (!bookedTimes.includes(t)) {
        availableTimes.add(t);
      }
    });
  });

  return [...availableTimes].sort();
}

function submitStudentSession(studentData, subject, time) {
  const ss = getSpreadsheet();
  const tutorSheet = ss.getSheetByName('Tutors');
  const matchSheet = ss.getSheetByName('Matches');
  const studentSheet = ss.getSheetByName('Students');
  const tutorData = tutorSheet.getDataRange().getValues();
  const matchData = matchSheet.getDataRange().getValues();

  // Build booked slots per tutor
  const booked = {};
  matchData.slice(1).forEach(row => {
    const tutorEmail = row[5];
    const time = row[7];
    if (tutorEmail && time) {
      if (!booked[tutorEmail]) booked[tutorEmail] = [];
      booked[tutorEmail].push(time);
    }
  });

  // Find best available tutor
  let bestTutor = null;
  let bestLoad = Infinity;

  tutorData.slice(1).forEach(row => {
    const tutorName = row[0];
    const tutorGrade = parseInt(row[1]);
    const tutorEmail = row[2];
    const subjects = row[3] ? row[3].split(', ') : [];
    const times = row[4] ? row[4].split(', ') : [];
    const bookedTimes = booked[tutorEmail] || [];

    if (!subjects.includes(subject)) return;
    if (tutorGrade < studentData.grade) return;
    if (!times.includes(time)) return;
    if (bookedTimes.includes(time)) return;

    const load = bookedTimes.length;
    if (load < bestLoad) {
      bestLoad = load;
      bestTutor = { name: tutorName, grade: tutorGrade, email: tutorEmail };
    }
  });

  if (!bestTutor) {
    return { success: false, message: 'Sorry, that slot was just taken. Please pick another time.' };
  }

  // Save student to Students sheet
  studentSheet.appendRow([
    studentData.name,
    studentData.grade,
    studentData.email,
    studentData.parentEmail,
    JSON.stringify([{ subject, times: [time] }]),
    new Date().toLocaleDateString()
  ]);

  // Save match to Matches sheet
  const weekOf = new Date().toLocaleDateString();
  matchSheet.appendRow([
    weekOf,
    studentData.name,
    studentData.email,
    studentData.parentEmail,
    bestTutor.name,
    bestTutor.email,
    subject,
    time
  ]);

  return {
    success: true,
    tutorName: bestTutor.name,
    subject,
    time
  };
}

// ─── ADMIN FUNCTIONS ───────────────────────────────────────

function getAdminStats() {
  const ss = getSpreadsheet();
  const studentSheet = ss.getSheetByName('Students');
  const matchSheet = ss.getSheetByName('Matches');

  const studentCount = Math.max(0, studentSheet.getLastRow() - 1);
  const matchCount = Math.max(0, matchSheet.getLastRow() - 1);

  const studentData = studentSheet.getDataRange().getValues();
  const totalRequests = studentData.slice(1).reduce((acc, row) => {
    try { return acc + JSON.parse(row[4]).length; }
    catch(e) { return acc; }
  }, 0);

  return {
    studentCount,
    matchCount,
    unmatchedCount: Math.max(0, totalRequests - matchCount)
  };
}

function getCurrentMatches() {
  const ss = getSpreadsheet();
  const matchSheet = ss.getSheetByName('Matches');
  const data = matchSheet.getDataRange().getValues();

  if (data.length <= 1) return { matched: [] };

  const matched = data.slice(1).map(row => ({
    studentName: row[1],
    studentEmail: row[2],
    parentEmail: row[3],
    tutorName: row[4],
    tutorEmail: row[5],
    subject: row[6],
    bestTime: row[7]
  }));

  return { matched };
}

function sendAllMatchEmails() {
  const ss = getSpreadsheet();
  const matchSheet = ss.getSheetByName('Matches');
  const data = matchSheet.getDataRange().getValues();

  if (data.length <= 1) return { success: false, message: 'No matches found.' };

  data.slice(1).forEach(row => {
    const studentName = row[1];
    const studentEmail = row[2];
    const parentEmail = row[3];
    const tutorName = row[4];
    const tutorEmail = row[5];
    const subject = row[6];
    const time = row[7];

    const emailSubject = `Tutoring Session Confirmed — ${subject}`;

    const studentBody = `Hi ${studentName},\n\nYou've been matched with a tutor for the upcoming week!\n\nSubject: ${subject}\nTutor: ${tutorName}\nTime: ${time}\n\nPlease make sure to show up on time. If you need to cancel, contact your school coordinator as soon as possible.\n\nGood luck!`;
    const tutorBody = `Hi ${tutorName},\n\nYou have a new tutoring session scheduled for the upcoming week!\n\nSubject: ${subject}\nStudent: ${studentName}\nTime: ${time}\n\nPlease make sure to show up on time. If you need to cancel, contact your school coordinator as soon as possible.\n\nThank you for tutoring!`;
    const parentBody = `Hi,\n\nYour child ${studentName} has been matched with a tutor for the upcoming week!\n\nSubject: ${subject}\nTutor: ${tutorName}\nTime: ${time}\n\nIf you have any questions, please contact your school coordinator.\n\nThank you!`;

    GmailApp.sendEmail(studentEmail, emailSubject, studentBody);
    GmailApp.sendEmail(tutorEmail, emailSubject, tutorBody);
    if (parentEmail) GmailApp.sendEmail(parentEmail, emailSubject, parentBody);
  });

  return { success: true, count: data.length - 1 };
}

// ─── WEEKLY RESET ──────────────────────────────────────────

function clearStudentsWeekly() {
  const ss = getSpreadsheet();
  const studentSheet = ss.getSheetByName('Students');
  const matchSheet = ss.getSheetByName('Matches');

  let archiveSheet = ss.getSheetByName('Archive');
  if (!archiveSheet) {
    archiveSheet = ss.insertSheet('Archive');
    archiveSheet.appendRow([
      'Week Of', 'Student Name', 'Student Email', 'Parent Email',
      'Tutor Name', 'Tutor Email', 'Subject', 'Time', 'Archived On'
    ]);
  }

  const matchData = matchSheet.getDataRange().getValues();
  const weekOf = new Date().toLocaleDateString();
  matchData.slice(1).forEach(row => {
    if (row[0]) archiveSheet.appendRow([...row, weekOf]);
  });

  const studentLastRow = studentSheet.getLastRow();
  if (studentLastRow > 1) studentSheet.deleteRows(2, studentLastRow - 1);

  const matchLastRow = matchSheet.getLastRow();
  if (matchLastRow > 1) matchSheet.deleteRows(2, matchLastRow - 1);
}

// 
function getTutorSessions(email) {
  try {
    const ss = getSpreadsheet();
    const matchSheet = ss.getSheetByName('Matches');
    if (!matchSheet) return []; // Return empty array if tab is missing

    const data = matchSheet.getDataRange().getValues();
    if (data.length <= 1) return []; // Return empty array if only headers exist

    const sessions = [];
    const thisWeek = new Date().toLocaleDateString();

    data.slice(1).forEach((row, i) => {
      const tutorEmail = row[5] ? String(row[5]).trim().toLowerCase() : '';
      if (tutorEmail === email.trim().toLowerCase()) {
        let sessionDate = '';
        if (row[0] instanceof Date) {
          sessionDate = row[0].toLocaleDateString();
        } else if (row[0] !== undefined && row[0] !== null) {
          sessionDate = String(row[0]);
        }

        sessions.push({
          rowIndex: i + 2,
          weekOf: sessionDate,
          studentName: row[1] || '',
          subject: row[6] || '',
          time: row[7] || '',
          status: row[8] || 'Pending',
          isThisWeek: sessionDate === thisWeek
        });
      }
    });

    return sessions;
  } catch (err) {
    Logger.log('Error in getTutorSessions: ' + err.toString());
    throw new Error(err.message);
  }
}

function updateSessionStatus(rowIndex, status) {
  const ss = getSpreadsheet();
  const matchSheet = ss.getSheetByName('Matches');
  matchSheet.getRange(rowIndex, 9).setValue(status);
  return { success: true };
}

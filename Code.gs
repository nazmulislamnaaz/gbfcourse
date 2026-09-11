/**
 * =====================================================================
 *  COURSE CONTENT WEB APP — "ঘরে বসে ফ্রিলান্সিং"
 *  Instructor: Nazmul Islam Naaz
 *  Paid batch — Registration (with photo) → Admin Approval → Content Access
 * =====================================================================
 */

// ---------------------- CONFIG ----------------------
const SS = SpreadsheetApp.getActiveSpreadsheet();
const STUDENTS_SHEET      = 'Students';
const MODULES_SHEET       = 'Modules';
const CONTENTS_SHEET      = 'Contents';
const REGISTRATIONS_SHEET = 'Registrations';
const PHOTO_FOLDER_NAME   = 'Course Student Photos';
const SESSION_DURATION    = 6 * 60 * 60; // সেশন মেয়াদ: ৬ ঘণ্টা


// ---------------------- WEB APP ENTRY ----------------------
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ঘরে বসে ফ্রিলান্সিং - কোর্স')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


/** Facebook/YouTube Live অন থাকলে সেই তথ্য রিটার্ন করে — এডমিন Settings শিট থেকে আপডেট করেন */
function getLiveStatus() {
  const sheet = SS.getSheetByName('Settings');
  if (!sheet) return { success: true, facebook: null, youtube: null };

  const data = sheet.getDataRange().getValues();
  const result = { success: true, facebook: null, youtube: null };

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const platform = (row[0] || '').toString().trim().toLowerCase();
    const url = (row[1] || '').toString().trim();
    const isLive = row[2] === true || (row[2] || '').toString().trim().toLowerCase() === 'true';

    if (platform === 'facebook') result.facebook = { url: url, isLive: isLive };
    if (platform === 'youtube') result.youtube = { url: url, isLive: isLive };
  }

  return result;
}

/** হোমপেজের Hero ভিডিও Settings শিটের 'HomepageVideo' row থেকে নিয়ন্ত্রিত হয় */
function getHomepageVideoConfig_() {
  const sheet = SS.getSheetByName('Settings');
  if (!sheet) return { active: false, videoId: '' };

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const platform = (data[i][0] || '').toString().trim().toLowerCase();
    if (platform === 'homepagevideo') {
      const url = (data[i][1] || '').toString().trim();
      const isActive = data[i][2] === true || (data[i][2] || '').toString().trim().toLowerCase() === 'true';
      const videoId = extractYouTubeId_(url) || url; // পুরো লিংক অথবা শুধু ভিডিও ID — দুটোই কাজ করবে
      return { active: isActive && !!videoId, videoId: videoId };
    }
  }
  return { active: false, videoId: '' };
}

/** Dashboard-এর উপরে দেখানো এডমিন ঘোষণা — Announcement শিট থেকে পড়ে */
function getAnnouncement() {
  const sheet = SS.getSheetByName('Announcement');
  if (!sheet) return { success: true, active: false };

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { success: true, active: false };

  const row = data[1]; // প্রথম ডেটা row-ই একমাত্র/সক্রিয় ঘোষণা হিসেবে ব্যবহৃত হয়
  const message = (row[0] || '').toString().trim();
  const linkUrl = (row[1] || '').toString().trim();
  const linkLabel = (row[2] || '').toString().trim();
  const isActive = row[3] === true || (row[3] || '').toString().trim().toLowerCase() === 'true';

  if (!isActive || !message) return { success: true, active: false };

  return {
    success: true,
    active: true,
    message: message,
    linkUrl: linkUrl,
    linkLabel: linkLabel || 'বিস্তারিত দেখুন'
  };
}


// ---------------------- SYLLABUS STRUCTURE (CACHED — সবচেয়ে বেশি পড়া হয়, তাই cache করা) ----------------------
const SYLLABUS_CACHE_KEY = 'syllabus_structure_v1';
const SYLLABUS_CACHE_TTL = 300; // ৫ মিনিট — এডমিন Modules/Contents শিট বদলালে সর্বোচ্চ ৫ মিনিটে reflect করবে

/** Modules+Contents একসাথে বানিয়ে (embedType/embedUrl সহ) cache করে রাখে — বারবার শিট না পড়ে দ্রুত রেসপন্স দেয় */
function getSyllabusStructure_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(SYLLABUS_CACHE_KEY);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) { /* কর্প্ট cache হলে নিচে থেকে রিবিল্ড হবে */ }
  }

  const modules = getModulesList_();
  const modMap  = {};
  modules.forEach(function (m) { modMap[m.id] = m; });

  getContentsList_().forEach(function (c) {
    const embed = buildEmbedInfo_(c.type, c.url);
    c.embedType = embed.embedType;
    c.embedUrl  = embed.embedUrl;
    if (modMap[c.moduleId]) modMap[c.moduleId].contents.push(c);
  });

  try {
    cache.put(SYLLABUS_CACHE_KEY, JSON.stringify(modules), SYLLABUS_CACHE_TTL);
  } catch (e) { /* সাইজ লিমিটে সমস্যা হলে cache স্কিপ, তবু ডেটা রিটার্ন হবে */ }

  return modules;
}

/** এডমিন Modules/Contents শিটে ম্যানুয়ালি বদলানোর পর সাথে সাথে reflect করাতে চাইলে এই ফাংশন রান করলে cache সাথে সাথে খালি হয়ে যাবে */
function clearSyllabusCache() {
  CacheService.getScriptCache().remove(SYLLABUS_CACHE_KEY);
}


// ---------------------- PUBLIC HOMEPAGE DATA ----------------------
/** লগইন ছাড়াই হোমপেজে সিলেবাস দেখানোর জন্য — URL এক্সপোজ করে না, শুধু টাইটেল */
function getPublicSyllabus() {
  const modules = getSyllabusStructure_(); // cached — URL সহ পূর্ণ স্ট্রাকচার, নিচে স্ট্রিপ করা হচ্ছে
  const stripped = modules.map(function (m) {
    return {
      id: m.id,
      title: m.title,
      order: m.order,
      description: m.description,
      contents: m.contents.map(function (c) { return { title: c.title, type: c.type }; })
    };
  });
  return { success: true, modules: stripped };
}

/** হোমপেজের Contact ফর্ম থেকে মেসেজ আসলে Messages শিটে জমা রাখে */
function submitContactMessage(data) {
  const name    = (data.name || '').toString().trim();
  const contact = (data.contact || '').toString().trim();
  const message = (data.message || '').toString().trim();

  if (!name || !contact || !message) {
    return { success: false, message: 'নাম, যোগাযোগের তথ্য ও মেসেজ — সবগুলো দিন।' };
  }

  const sheet = SS.getSheetByName('Messages');
  if (!sheet) return { success: false, message: 'সিস্টেম এরর: Messages শিট পাওয়া যায়নি।' };

  sheet.appendRow([new Date(), name, contact, message]);
  return { success: true, message: 'আপনার মেসেজ পাঠানো হয়েছে। শীঘ্রই যোগাযোগ করা হবে। ধন্যবাদ!' };
}


/** পারফরম্যান্স: হোমপেজ লোডে ২টা আলাদা রাউন্ড-ট্রিপের বদলে ১টায় সিলেবাস + ফিডব্যাক একসাথে পাঠায় */
function getHomepageBootstrap() {
  const syllabus = getPublicSyllabus();
  const feedbacks = getApprovedFeedbacks();
  const video = getHomepageVideoConfig_();
  return {
    success: true,
    modules: syllabus.modules || [],
    feedbacks: feedbacks.feedbacks || [],
    video: video
  };
}


// ---------------------- REGISTRATION ----------------------
/**
 * data = { name, phone, email, paymentNumber, photoBase64, photoMime }
 * সরাসরি অ্যাক্সেস দেয় না — Registrations শিটে Pending রাখে।
 * এডমিন পেমেন্ট যাচাই করে Students শিটে ম্যানুয়ালি এড করলে অ্যাক্সেস চালু হয়।
 */
function submitRegistration(data) {
  const name          = (data.name || '').toString().trim();
  const phone         = (data.phone || '').toString().trim();
  const email         = (data.email || '').toString().trim().toLowerCase();
  const paymentNumber = (data.paymentNumber || '').toString().trim();
  const password      = (data.password || '').toString().trim();
  const photoBase64   = data.photoBase64 || '';

  if (!name || !phone) {
    return { success: false, message: 'নাম ও ফোন নাম্বার আবশ্যক।' };
  }
  if (!paymentNumber) {
    return { success: false, message: 'যে নাম্বার থেকে পেমেন্ট করেছেন সেটি দিন।' };
  }
  if (password.length < 4) {
    return { success: false, message: 'পাসওয়ার্ড কমপক্ষে ৪ ক্যারেক্টার হতে হবে।' };
  }
  if (!photoBase64) {
    return { success: false, message: 'প্রোফাইল ছবি আবশ্যক — স্টুডেন্ট আইডির জন্য প্রয়োজন।' };
  }

  if (findStudentByIdentifier_(phone.toLowerCase()) || (email && findStudentByIdentifier_(email))) {
    return { success: false, message: 'এই ফোন/ইমেইল দিয়ে ইতিমধ্যে একটি অ্যাকাউন্ট আছে। লগইন ট্যাব থেকে লগইন করুন।' };
  }

  const regSheet = SS.getSheetByName(REGISTRATIONS_SHEET);
  if (!regSheet) return { success: false, message: 'সিস্টেম এরর: Registrations শিট পাওয়া যায়নি।' };

  const regData = regSheet.getDataRange().getValues();
  for (let i = 1; i < regData.length; i++) {
    const row = regData[i];
    const rPhone  = (row[2] || '').toString().trim();
    const rEmail  = (row[3] || '').toString().trim().toLowerCase();
    const rStatus = (row[7] || '').toString().trim().toLowerCase();
    if (rStatus === 'pending' && (rPhone === phone || (email && rEmail === email))) {
      return { success: false, message: 'আপনার একটি আবেদন ইতিমধ্যে পর্যালোচনাধীন আছে। অনুগ্রহ করে অপেক্ষা করুন।' };
    }
  }

  let photoUrl = '';
  try {
    photoUrl = savePhotoToDrive_(photoBase64, data.photoMime || 'image/jpeg', phone);
  } catch (err) {
    return { success: false, message: 'ছবি আপলোডে সমস্যা হয়েছে, আবার চেষ্টা করুন।' };
  }

  // Header: Timestamp|Name|Phone|Email|PaymentNumber|Password|PhotoURL|Status|AdminNote
  regSheet.appendRow([new Date(), name, phone, email, paymentNumber, password, photoUrl, 'Pending', '']);

  return {
    success: true,
    message: 'আপনার আবেদন সফলভাবে জমা হয়েছে। পেমেন্ট যাচাই শেষে এডমিন অ্যাক্সেস চালু করবেন — আপনার সেট করা পাসওয়ার্ড দিয়েই তখন লগইন করতে পারবেন।'
  };
}

/** Base64 ছবি Drive-এ সেভ করে পাবলিক থাম্বনেইল URL রিটার্ন করে */
function savePhotoToDrive_(base64Data, mimeType, phone) {
  const folder = getOrCreatePhotoFolder_();
  const bytes = Utilities.base64Decode(base64Data);
  const blob = Utilities.newBlob(bytes, mimeType, 'student_' + phone + '_' + new Date().getTime() + '.jpg');
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w300';
}

function getOrCreatePhotoFolder_() {
  const folders = DriveApp.getFoldersByName(PHOTO_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(PHOTO_FOLDER_NAME);
}

/** স্টুডেন্ট নিজের আবেদনের status যাচাই করতে পারে */
function checkRegistrationStatus(identifier) {
  identifier = (identifier || '').toString().trim().toLowerCase();
  if (!identifier) return { success: false, message: 'ফোন অথবা ইমেইল দিন।' };

  const student = findStudentByIdentifier_(identifier);
  if (student) {
    if ((student.status || '').toLowerCase() === 'inactive') {
      return { success: true, stage: 'inactive', message: 'আপনার অ্যাকাউন্ট নিষ্ক্রিয় করা আছে। এডমিনের সাথে যোগাযোগ করুন।' };
    }
    return { success: true, stage: 'approved', message: 'আপনার অ্যাকাউন্ট সক্রিয়! লগইন ট্যাব থেকে লগইন করুন।' };
  }

  const regSheet = SS.getSheetByName(REGISTRATIONS_SHEET);
  if (!regSheet) return { success: false, message: 'কোনো আবেদন পাওয়া যায়নি।' };

  const regData = regSheet.getDataRange().getValues();
  let latest = null;
  for (let i = 1; i < regData.length; i++) {
    const row = regData[i];
    const rPhone = (row[2] || '').toString().trim().toLowerCase();
    const rEmail = (row[3] || '').toString().trim().toLowerCase();
    if (rPhone === identifier || rEmail === identifier) latest = row;
  }

  if (!latest) return { success: false, message: 'এই তথ্য দিয়ে কোনো আবেদন খুঁজে পাওয়া যায়নি।' };

  const status = (latest[7] || '').toString().trim().toLowerCase();
  if (status === 'rejected') {
    return { success: true, stage: 'rejected', message: 'দুঃখিত, আপনার আবেদনটি গ্রহণ করা হয়নি। ' + (latest[8] || 'বিস্তারিত জানতে এডমিনের সাথে যোগাযোগ করুন।') };
  }
  return { success: true, stage: 'pending', message: 'আপনার আবেদন পর্যালোচনাধীন আছে। পেমেন্ট যাচাই হলেই অ্যাক্সেস চালু হবে।' };
}


// ---------------------- AUTHENTICATION ----------------------
function authenticateUser(identifier, password, deviceLabel) {
  identifier = (identifier || '').toString().trim().toLowerCase();
  password   = (password || '').toString().trim();

  if (!identifier || !password) {
    return { success: false, message: 'ফোন/ইমেইল ও পাসওয়ার্ড দিন।' };
  }

  const row = findStudentRow_(identifier);
  if (!row) {
    return { success: false, message: 'অ্যাকাউন্ট পাওয়া যায়নি। রেজিস্ট্রেশন করেছেন? "স্ট্যাটাস চেক" ট্যাবে দেখুন।' };
  }

  if (row.values.password !== password) {
    return { success: false, message: 'পাসওয়ার্ড সঠিক নয়।' };
  }
  if ((row.values.status || '').toLowerCase() === 'inactive') {
    return { success: false, message: 'আপনার অ্যাকাউন্টটি নিষ্ক্রিয় করা হয়েছে। এডমিনের সাথে যোগাযোগ করুন।' };
  }

  // ডিভাইস তথ্য ও শেষ লগইন সময় আপডেট
  const sheet = SS.getSheetByName(STUDENTS_SHEET);
  sheet.getRange(row.rowIndex, 8).setValue(deviceLabel || 'Unknown device'); // LastDevice
  sheet.getRange(row.rowIndex, 9).setValue(new Date());                      // LastLogin

  const token = Utilities.getUuid();
  CacheService.getScriptCache().put(
    'session_' + token,
    JSON.stringify({ name: row.values.name, phone: row.values.phone, batch: row.values.batch }),
    SESSION_DURATION
  );

  return {
    success: true,
    token: token,
    name: row.values.name,
    batch: row.values.batch,
    photoUrl: row.values.photoUrl,
    deviceLabel: deviceLabel || ''
  };
}

/** Students শিটে identifier (phone/email) দিয়ে স্টুডেন্ট রো খোঁজে, rowIndex সহ */
function findStudentRow_(identifier) {
  const sheet = SS.getSheetByName(STUDENTS_SHEET);
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  // Header: Name|Phone|Email|Password|Batch|Status|PhotoURL|LastDevice|LastLogin|StudentCode

  const headers = data[0] || [];
  let codeColIdx = headers.indexOf('StudentCode'); // 0-based; -1 হলে কলাম এখনো তৈরি হয়নি

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    const phone = (row[1] || '').toString().trim().toLowerCase();
    const email = (row[2] || '').toString().trim().toLowerCase();
    if (identifier === phone || identifier === email) {
      return {
        rowIndex: i + 1,
        values: {
          name: row[0].toString().trim(),
          phone: (row[1] || '').toString().trim(),
          email: (row[2] || '').toString().trim(),
          password: (row[3] || '').toString().trim(),
          batch: (row[4] || '').toString().trim(),
          status: (row[5] || '').toString().trim(),
          photoUrl: (row[6] || '').toString().trim(),
          code: codeColIdx >= 0 ? (row[codeColIdx] || '').toString().trim() : ''
        }
      };
    }
  }
  return null;
}

function findStudentByIdentifier_(identifier) {
  const row = findStudentRow_(identifier);
  return row ? row.values : null;
}

function validateSession(token) {
  if (!token) return null;
  const raw = CacheService.getScriptCache().get('session_' + token);
  return raw ? JSON.parse(raw) : null;
}

function logoutUser(token) {
  if (token) CacheService.getScriptCache().remove('session_' + token);
  return true;
}


// ---------------------- PROGRESS TRACKING ----------------------
/** একটা কনটেন্ট আইটেম দেখা হয়েছে হিসেবে মার্ক করে (সার্ভার-সাইড, তাই সব ডিভাইসে সিঙ্ক থাকে) */
function markViewed(token, contentId) {
  const session = validateSession(token);
  if (!session || !contentId) return { success: false };

  const sheet = SS.getSheetByName('Progress');
  if (!sheet) return { success: false };

  const phone = session.phone;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if ((data[i][0] || '').toString().trim() === phone && (data[i][1] || '').toString().trim() === contentId) {
      return { success: true }; // আগে থেকেই মার্ক করা আছে
    }
  }

  sheet.appendRow([phone, contentId, new Date()]);
  return { success: true };
}

/** এই স্টুডেন্টের দেখা সব কনটেন্ট আইডির লিস্ট রিটার্ন করে */
function getViewedContentIds_(phone) {
  const sheet = SS.getSheetByName('Progress');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const ids = [];
  for (let i = 1; i < data.length; i++) {
    if ((data[i][0] || '').toString().trim() === phone) {
      ids.push((data[i][1] || '').toString().trim());
    }
  }
  return ids;
}


// ---------------------- FEEDBACK / TESTIMONIALS ----------------------
/** লগইনকৃত স্টুডেন্ট ফিডব্যাক সাবমিট করে — সরাসরি পাবলিক হয় না, এডমিন Approve করলে হোমপেজে দেখাবে */
function submitFeedback(token, rating, message) {
  const session = validateSession(token);
  if (!session) return { success: false, message: 'সেশনের মেয়াদ শেষ। আবার লগইন করুন।' };

  message = (message || '').toString().trim();
  rating = Number(rating) || 0;

  if (!message) return { success: false, message: 'মতামত লিখুন।' };
  if (rating < 1 || rating > 5) return { success: false, message: 'রেটিং সিলেক্ট করুন।' };

  const row = findStudentRow_(session.phone.toLowerCase());
  const name = row ? row.values.name : session.name;
  const photoUrl = row ? row.values.photoUrl : '';

  const sheet = SS.getSheetByName('Feedback');
  if (!sheet) return { success: false, message: 'সিস্টেম এরর: Feedback শিট পাওয়া যায়নি।' };

  sheet.appendRow([new Date(), session.phone, name, photoUrl, rating, message, 'Pending']);
  return { success: true, message: 'আপনার মতামতের জন্য ধন্যবাদ! এডমিন Approve করার পর এটি সবাই দেখতে পাবে।' };
}

/** শুধুমাত্র Approved ফিডব্যাক পাবলিকলি রিটার্ন করে (হোমপেজে দেখানোর জন্য, লগইন লাগে না) */
function getApprovedFeedbacks() {
  const sheet = SS.getSheetByName('Feedback');
  if (!sheet) return { success: true, feedbacks: [] };

  const data = sheet.getDataRange().getValues();
  const feedbacks = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const status = (row[6] || '').toString().trim().toLowerCase();
    if (status === 'approved') {
      feedbacks.push({
        name: row[2] || 'শিক্ষার্থী',
        photoUrl: row[3] || '',
        rating: Number(row[4]) || 5,
        message: row[5] || ''
      });
    }
  }
  feedbacks.reverse(); // সাম্প্রতিক আগে
  return { success: true, feedbacks: feedbacks };
}


/** স্টুডেন্টের জন্য ইউনিক ৬-সংখ্যার কোড — আগে থেকে থাকলে সেটাই, না থাকলে নতুন জেনারেট করে সেভ করে */
/** Students শিটে 'StudentCode' হেডার কলাম আছে কিনা যাচাই করে, না থাকলে তৈরি করে দেয় (পুরনো শিটেও কাজ করবে) */
function ensureStudentCodeColumn_(sheet) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  let colIndex = headers.indexOf('StudentCode') + 1; // 1-based; না পেলে 0 হবে

  if (colIndex === 0) {
    colIndex = lastCol + 1;
    sheet.getRange(1, colIndex).setValue('StudentCode')
      .setFontWeight('bold').setBackground('#0B3D3A').setFontColor('#FFFFFF');
  }
  return colIndex;
}

function getOrCreateStudentCode_(sheet, rowIndex, currentCode) {
  currentCode = (currentCode || '').toString().trim();
  const codeCol = ensureStudentCodeColumn_(sheet);

  if (!currentCode) {
    // rowIndex-এর বর্তমান মান আবার সরাসরি চেক করা হচ্ছে, যাতে column সদ্য তৈরি হয়ে থাকলেও ঠিক মান পাওয়া যায়
    currentCode = sheet.getRange(rowIndex, codeCol).getValue().toString().trim();
  }
  if (currentCode) return currentCode;

  const allData = sheet.getDataRange().getValues();
  const existing = {};
  for (let i = 1; i < allData.length; i++) {
    const c = (allData[i][codeCol - 1] || '').toString().trim();
    if (c) existing[c] = true;
  }

  let code;
  let attempts = 0;
  do {
    code = Math.floor(100000 + Math.random() * 900000).toString();
    attempts++;
  } while (existing[code] && attempts < 50);

  sheet.getRange(rowIndex, codeCol).setValue(code); // StudentCode কলামে সেভ — Google Sheet-এ সাথে সাথে আপডেট হয়
  SpreadsheetApp.flush(); // নিশ্চিত করে যে লেখাটা তৎক্ষণাৎ শিটে সেভ হয়ে যায়
  return code;
}


// ---------------------- SELF-SERVICE PROFILE UPDATE (নাম + ছবি) ----------------------
/**
 * স্টুডেন্ট নিজের নাম ও প্রোফাইল ছবি পরিবর্তন করতে পারে।
 * ফোন/ইমেইল বদলানো যায় না — এটাই লগইন identifier, নিরাপত্তার জন্য অপরিবর্তনীয়।
 */
function updateProfile(token, newName, photoBase64, photoMime) {
  const session = validateSession(token);
  if (!session) return { success: false, message: 'সেশনের মেয়াদ শেষ। আবার লগইন করুন।' };

  newName = (newName || '').toString().trim();
  if (!newName) return { success: false, message: 'নাম দিন।' };
  if (newName.length > 60) return { success: false, message: 'নাম খুব বড়, ছোট করে দিন।' };

  const row = findStudentRow_(session.phone.toLowerCase());
  if (!row) return { success: false, message: 'অ্যাকাউন্ট পাওয়া যায়নি।' };

  const sheet = SS.getSheetByName(STUDENTS_SHEET);
  sheet.getRange(row.rowIndex, 1).setValue(newName); // Name কলাম

  let photoUrl = row.values.photoUrl;
  if (photoBase64) {
    try {
      photoUrl = savePhotoToDrive_(photoBase64, photoMime || 'image/jpeg', session.phone);
      sheet.getRange(row.rowIndex, 7).setValue(photoUrl); // PhotoURL কলাম
    } catch (err) {
      return { success: false, message: 'ছবি আপলোডে সমস্যা হয়েছে, আবার চেষ্টা করুন।' };
    }
  }

  // সেশন cache-এও নতুন নাম আপডেট করা হচ্ছে যাতে সাথে সাথে সব জায়গায় reflect করে
  CacheService.getScriptCache().put(
    'session_' + token,
    JSON.stringify({ name: newName, phone: session.phone, batch: session.batch }),
    SESSION_DURATION
  );

  return { success: true, message: 'প্রোফাইল সফলভাবে আপডেট হয়েছে।', name: newName, photoUrl: photoUrl };
}


// ---------------------- SELF-SERVICE PASSWORD CHANGE ----------------------
function changePassword(token, oldPassword, newPassword) {
  const session = validateSession(token);
  if (!session) return { success: false, message: 'সেশনের মেয়াদ শেষ। আবার লগইন করুন।' };

  newPassword = (newPassword || '').toString().trim();
  oldPassword = (oldPassword || '').toString().trim();

  if (newPassword.length < 4) {
    return { success: false, message: 'নতুন পাসওয়ার্ড কমপক্ষে ৪ ক্যারেক্টার হতে হবে।' };
  }

  const row = findStudentRow_(session.phone.toLowerCase());
  if (!row) return { success: false, message: 'অ্যাকাউন্ট পাওয়া যায়নি।' };

  if (row.values.password !== oldPassword) {
    return { success: false, message: 'বর্তমান পাসওয়ার্ড সঠিক নয়।' };
  }

  const sheet = SS.getSheetByName(STUDENTS_SHEET);
  sheet.getRange(row.rowIndex, 4).setValue(newPassword); // Password কলাম
  return { success: true, message: 'পাসওয়ার্ড সফলভাবে পরিবর্তন হয়েছে।' };
}


// ---------------------- COURSE CONTENT ----------------------
/** পারফরম্যান্স: লগইনের পর ৩টা আলাদা রাউন্ড-ট্রিপ (course+live+announcement) এর বদলে ১টায় সব পাঠায় */
function getDashboardBootstrap(token) {
  const course = getCourseContent(token);
  if (!course.success) return course;

  const live = getLiveStatus();
  const announcement = getAnnouncement();

  course.live = { facebook: live.facebook, youtube: live.youtube };
  course.announcement = {
    active: announcement.active,
    message: announcement.message || '',
    linkUrl: announcement.linkUrl || '',
    linkLabel: announcement.linkLabel || ''
  };
  return course;
}

function getCourseContent(token) {
  const session = validateSession(token);
  if (!session) return { success: false, message: 'সেশনের মেয়াদ শেষ হয়ে গেছে। আবার লগইন করুন।' };

  const modules = getSyllabusStructure_(); // cached — বারবার শিট পড়তে হয় না

  const studentRow = findStudentRow_(session.phone.toLowerCase());
  let studentCode = '';
  if (studentRow) {
    const stuSheet = SS.getSheetByName(STUDENTS_SHEET);
    studentCode = getOrCreateStudentCode_(stuSheet, studentRow.rowIndex, studentRow.values.code);
  }

  return {
    success: true,
    studentName: session.name,
    batch: session.batch,
    phone: studentRow ? studentRow.values.phone : session.phone,
    email: studentRow ? studentRow.values.email : '',
    photoUrl: studentRow ? studentRow.values.photoUrl : '',
    studentCode: studentCode,
    viewedIds: getViewedContentIds_(session.phone),
    modules: modules
  };
}

function getModulesList_() {
  const sheet = SS.getSheetByName(MODULES_SHEET);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const modules = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    modules.push({
      id: row[0].toString().trim(),
      title: row[1] || '',
      order: Number(row[2]) || 0,
      description: row[3] || '',
      contents: []
    });
  }
  modules.sort(function (a, b) { return a.order - b.order; });
  return modules;
}

function getContentsList_() {
  const sheet = SS.getSheetByName(CONTENTS_SHEET);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    rows.push({
      id: row[0].toString().trim(),
      moduleId: (row[1] || '').toString().trim(),
      title: row[2] || '',
      type: row[3] || '',
      url: row[4] || '',
      order: Number(row[5]) || 0,
      description: row[6] || ''
    });
  }
  rows.sort(function (a, b) { return a.order - b.order; });
  return rows;
}


// ---------------------- EMBED HELPERS ----------------------
function buildEmbedInfo_(type, url) {
  type = (type || '').toString().trim().toLowerCase();
  url  = (url || '').toString().trim();

  if (type.indexOf('youtube') !== -1) {
    const id = extractYouTubeId_(url);
    return { embedType: 'youtube', embedUrl: id ? 'https://www.youtube.com/embed/' + id : url };
  }
  if (type.indexOf('drive') !== -1) {
    const id = extractDriveId_(url);
    return { embedType: 'drive', embedUrl: id ? 'https://drive.google.com/file/d/' + id + '/preview' : url };
  }
  return { embedType: 'link', embedUrl: url };
}

function extractYouTubeId_(url) {
  const patterns = [
    /youtu\.be\/([^?&]+)/,
    /youtube\.com\/watch\?v=([^&]+)/,
    /youtube\.com\/embed\/([^?&]+)/,
    /youtube\.com\/shorts\/([^?&]+)/
  ];
  for (let i = 0; i < patterns.length; i++) {
    const m = url.match(patterns[i]);
    if (m) return m[1];
  }
  return null;
}

function extractDriveId_(url) {
  const patterns = [/\/d\/([a-zA-Z0-9_-]+)/, /id=([a-zA-Z0-9_-]+)/];
  for (let i = 0; i < patterns.length; i++) {
    const m = url.match(patterns[i]);
    if (m) return m[1];
  }
  return null;
}

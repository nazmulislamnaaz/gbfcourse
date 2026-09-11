/**
 * =====================================================================
 *  SETUP SCRIPT — একবারই রান করবেন
 *  Students, Registrations, Modules, Contents শিট তৈরি করে এবং
 *  পুরো সিলেবাস (১০ মডিউল, ৬০ ভিডিও, ১০টি Assessment, ৫টি Cheat Sheet)
 *  অটো-ফিল করে দেয়।
 *
 *  চালানোর নিয়ম: Apps Script এডিটরে ফাংশন ড্রপডাউন থেকে "setupSyllabus"
 *  সিলেক্ট করে Run চাপুন (প্রথমবার permission চাইবে, Allow করুন)।
 *
 *  ⚠️ রি-রান করলে Modules/Contents শিট মুছে আবার নতুন করে লেখা হবে।
 * =====================================================================
 */

function setupSyllabus() {
  createStudentsSheet_();
  createRegistrationsSheet_();
  createModulesAndContents_();
  createMessagesSheet_();
  createSettingsSheet_();
  createAnnouncementSheet_();
  createProgressSheet_();
  createFeedbackSheet_();
  SpreadsheetApp.getUi().alert('✅ সেটআপ সম্পন্ন! এখন Modules/Contents শিটে গিয়ে প্রতিটি ভিডিও, অ্যাসেসমেন্ট ও চিট শীটের URL কলামে লিংক বসান।');
}

function createProgressSheet_() {
  let sheet = SS.getSheetByName('Progress');
  if (!sheet) {
    sheet = SS.insertSheet('Progress');
    sheet.appendRow(['Phone', 'ContentID', 'ViewedAt']);
    sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#0B3D3A').setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, 3);
  }
}

function createFeedbackSheet_() {
  let sheet = SS.getSheetByName('Feedback');
  if (!sheet) {
    sheet = SS.insertSheet('Feedback');
    sheet.appendRow(['Timestamp', 'Phone', 'Name', 'PhotoURL', 'Rating', 'Message', 'Status']);
    sheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#0B3D3A').setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, 7);
  }
}

function createSettingsSheet_() {
  let sheet = SS.getSheetByName('Settings');
  if (!sheet) {
    sheet = SS.insertSheet('Settings');
    sheet.appendRow(['Platform', 'LiveURL', 'IsLive', 'UpdatedAt']);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#0B3D3A').setFontColor('#FFFFFF');
    sheet.appendRow(['Facebook', '', false, '']);
    sheet.appendRow(['YouTube', '', false, '']);
    sheet.appendRow(['HomepageVideo', '', false, '']);
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, 4);
  }
}

function createAnnouncementSheet_() {
  let sheet = SS.getSheetByName('Announcement');
  if (!sheet) {
    sheet = SS.insertSheet('Announcement');
    sheet.appendRow(['Message', 'LinkURL', 'LinkLabel', 'IsActive', 'UpdatedAt']);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#0B3D3A').setFontColor('#FFFFFF');
    sheet.appendRow(['', '', '', false, '']);
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, 5);
    sheet.setColumnWidth(1, 380);
  }
}

function createMessagesSheet_() {
  let sheet = SS.getSheetByName('Messages');
  if (!sheet) {
    sheet = SS.insertSheet('Messages');
    sheet.appendRow(['Timestamp', 'Name', 'Contact', 'Message']);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#0B3D3A').setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, 4);
  }
}

function createStudentsSheet_() {
  let sheet = SS.getSheetByName(STUDENTS_SHEET);
  if (!sheet) {
    sheet = SS.insertSheet(STUDENTS_SHEET);
    sheet.appendRow(['Name', 'Phone', 'Email', 'Password', 'Batch', 'Status', 'PhotoURL', 'LastDevice', 'LastLogin', 'StudentCode']);
    sheet.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground('#0B3D3A').setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, 10);
  }
}

function createRegistrationsSheet_() {
  let sheet = SS.getSheetByName(REGISTRATIONS_SHEET);
  if (!sheet) {
    sheet = SS.insertSheet(REGISTRATIONS_SHEET);
    sheet.appendRow(['Timestamp', 'Name', 'Phone', 'Email', 'PaymentNumber', 'Password', 'PhotoURL', 'Status', 'AdminNote']);
    sheet.getRange(1, 1, 1, 9).setFontWeight('bold').setBackground('#0B3D3A').setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, 9);
  }
}

function createModulesAndContents_() {
  let modSheet = SS.getSheetByName(MODULES_SHEET);
  if (modSheet) SS.deleteSheet(modSheet);
  modSheet = SS.insertSheet(MODULES_SHEET);
  modSheet.appendRow(['ModuleID', 'ModuleTitle', 'ModuleOrder', 'ModuleDescription']);
  modSheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#0B3D3A').setFontColor('#FFFFFF');
  modSheet.setFrozenRows(1);

  let conSheet = SS.getSheetByName(CONTENTS_SHEET);
  if (conSheet) SS.deleteSheet(conSheet);
  conSheet = SS.insertSheet(CONTENTS_SHEET);
  conSheet.appendRow(['ContentID', 'ModuleID', 'Title', 'Type', 'URL', 'Order', 'Description']);
  conSheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#0B3D3A').setFontColor('#FFFFFF');
  conSheet.setFrozenRows(1);

  const syllabus = getSyllabusData_();
  const modRows = [];
  const conRows = [];
  let contentCounter = 1;

  syllabus.forEach(function (mod) {
    modRows.push([mod.id, mod.title, mod.order, mod.description || '']);
    mod.items.forEach(function (item, cIdx) {
      conRows.push([
        'C' + contentCounter,
        mod.id,
        item.title,
        item.type,
        '', // URL admin পরে বসাবেন
        cIdx + 1,
        item.description || ''
      ]);
      contentCounter++;
    });
  });

  modSheet.getRange(2, 1, modRows.length, 4).setValues(modRows);
  conSheet.getRange(2, 1, conRows.length, 7).setValues(conRows);
  modSheet.autoResizeColumns(1, 4);
  conSheet.autoResizeColumns(1, 7);
}

/** পুরো সিলেবাস — রিসোর্স মডিউল (৫ চিট শীট) + ১০ মডিউল (৬০ ভিডিও + ১০টি Assessment) */
function getSyllabusData_() {
  return [
    {
      id: 'R0',
      order: 0,
      title: '📎 কোর্স রিসোর্স ও চিট শীট',
      description: 'পুরো কোর্সের ৫টি ডাউনলোডযোগ্য চিট শীট',
      items: [
        { title: 'চিট শীট ১', type: 'Drive', description: 'লিংক বসান' },
        { title: 'চিট শীট ২', type: 'Drive', description: 'লিংক বসান' },
        { title: 'চিট শীট ৩', type: 'Drive', description: 'লিংক বসান' },
        { title: 'চিট শীট ৪', type: 'Drive', description: 'লিংক বসান' },
        { title: 'চিট শীট ৫', type: 'Drive', description: 'লিংক বসান' }
      ]
    },
    {
      id: 'M1', order: 1, title: 'মডিউল ১: ফ্রিল্যান্সিং পরিচিতি', description: '৫টি ভিডিও',
      items: [
        { title: 'ভিডিও-১: ফ্রিল্যান্সিং কী?', type: 'YouTube' },
        { title: 'ভিডিও-২: ফ্রিল্যান্সিং সম্পর্কে প্রচলিত ভুল ধারণা ও বাস্তবতা', type: 'YouTube' },
        { title: 'ভিডিও-৩: চাকরি, ব্যবসা ও ফ্রিল্যান্সিং-এর মধ্যে পার্থক্য', type: 'YouTube' },
        { title: 'ভিডিও-৪: একজন সফল ফ্রিল্যান্সারের Career Roadmap', type: 'YouTube' },
        { title: 'ভিডিও-৫: Beginner হিসেবে সঠিকভাবে Freelancing শুরু করার পরিকল্পনা', type: 'YouTube' },
        { title: 'Module 01 Assessment Test', type: 'Link' }
      ]
    },
    {
      id: 'M2', order: 2, title: 'মডিউল ২: আন্তর্জাতিক Freelancing Marketplace', description: '৫টি ভিডিও',
      items: [
        { title: 'ভিডিও-৬: Freelancing Marketplace কীভাবে কাজ করে', type: 'YouTube' },
        { title: 'ভিডিও-৭: Fiverr সম্পর্কে সম্পূর্ণ ধারণা', type: 'YouTube' },
        { title: 'ভিডিও-৮: Upwork সম্পর্কে সম্পূর্ণ ধারণা', type: 'YouTube' },
        { title: 'ভিডিও-৯: Freelancer.com ও অন্যান্য জনপ্রিয় Marketplace পরিচিতি', type: 'YouTube' },
        { title: 'ভিডিও-১০: Buyer, Seller ও Marketplace-এর কার্যপ্রণালী', type: 'YouTube' },
        { title: 'Module 02 Assessment Test', type: 'Link' }
      ]
    },
    {
      id: 'M3', order: 3, title: 'মডিউল ৩: High-Demand Freelancing Skills', description: '১০টি ভিডিও',
      items: [
        { title: 'ভিডিও-১১: বর্তমান সময়ের সবচেয়ে চাহিদাসম্পন্ন Freelancing Skills', type: 'YouTube' },
        { title: 'ভিডিও-১২: Canva দিয়ে Professional Design-এর বেসিক', type: 'YouTube' },
        { title: 'ভিডিও-১৩: Content Writing-এর মৌলিক ধারণা', type: 'YouTube' },
        { title: 'ভিডিও-১৪: Data Entry ও Web Research', type: 'YouTube' },
        { title: 'ভিডিও-১৫: Virtual Assistant হিসেবে কাজ করার কৌশল', type: 'YouTube' },
        { title: 'ভিডিও-১৬: Professional Email Template Design', type: 'YouTube' },
        { title: 'ভিডিও-১৭: Instagram Marketing-এর বেসিক', type: 'YouTube' },
        { title: 'ভিডিও-১৮: YouTube Channel Management ও Optimization', type: 'YouTube' },
        { title: 'ভিডিও-১৯: AI Tools ব্যবহার করে দ্রুত ও স্মার্টভাবে কাজ করা', type: 'YouTube' },
        { title: 'ভিডিও-২০: নিজের জন্য সঠিক Skill নির্বাচন করার কৌশল', type: 'YouTube' },
        { title: 'Module 03 Assessment Test', type: 'Link' }
      ]
    },
    {
      id: 'M4', order: 4, title: 'মডিউল ৪: Professional CV, LinkedIn ও Portfolio', description: '৬টি ভিডিও',
      items: [
        { title: 'ভিডিও-২১: আন্তর্জাতিক মানের Professional CV তৈরি', type: 'YouTube' },
        { title: 'ভিডিও-২২: ATS Friendly CV তৈরির নিয়ম', type: 'YouTube' },
        { title: 'ভিডিও-২৩: Professional LinkedIn Account তৈরি', type: 'YouTube' },
        { title: 'ভিডিও-২৪: LinkedIn Profile Optimization', type: 'YouTube' },
        { title: 'ভিডিও-২৫: Smart Portfolio তৈরির কৌশল', type: 'YouTube' },
        { title: 'ভিডিও-২৬: Portfolio-কে Professionalভাবে উপস্থাপন করার নিয়ম', type: 'YouTube' },
        { title: 'Module 04 Assessment Test', type: 'Link' }
      ]
    },
    {
      id: 'M5', order: 5, title: 'মডিউল ৫: WordPress Portfolio Website', description: '৩টি ভিডিও',
      items: [
        { title: 'ভিডিও-২৭: WordPress-এর পরিচিতি, Domain ও Hosting সম্পর্কে প্রাথমিক ধারণা', type: 'YouTube' },
        { title: 'ভিডিও-২৮: WordPress ব্যবহার করে Professional Portfolio Website তৈরি', type: 'YouTube' },
        { title: 'ভিডিও-২৯: Portfolio, Project, CV ও Contact Form যুক্ত করে Website Publish', type: 'YouTube' },
        { title: 'Module 05 Assessment Test', type: 'Link' }
      ]
    },
    {
      id: 'M6', order: 6, title: 'মডিউল ৬: Fiverr Complete Guide', description: '৯টি ভিডিও',
      items: [
        { title: 'ভিডিও-৩০: Fiverr Account Create', type: 'YouTube' },
        { title: 'ভিডিও-৩১: Professional Profile Setup ও Optimization', type: 'YouTube' },
        { title: 'ভিডিও-৩২: লাভজনক Gig Research করার কৌশল', type: 'YouTube' },
        { title: 'ভিডিও-৩৩: SEO Friendly Gig Title ও Description লেখা', type: 'YouTube' },
        { title: 'ভিডিও-৩৪: Professional Gig Image Design', type: 'YouTube' },
        { title: 'ভিডিও-৩৫: Pricing, FAQ ও Requirements সেটআপ', type: 'YouTube' },
        { title: 'ভিডিও-৩৬: Gig Publish ও Optimization', type: 'YouTube' },
        { title: 'ভিডিও-৩৭: Gig Ranking ও Order পাওয়ার কৌশল', type: 'YouTube' },
        { title: 'ভিডিও-৩৮: Fiverr Success Tips ও Common Mistakes', type: 'YouTube' },
        { title: 'Module 06 Assessment Test', type: 'Link' }
      ]
    },
    {
      id: 'M7', order: 7, title: 'মডিউল ৭: Upwork Complete Guide', description: '৭টি ভিডিও',
      items: [
        { title: 'ভিডিও-৩৯: Upwork Account Create', type: 'YouTube' },
        { title: 'ভিডিও-৪০: Professional Profile Setup ও Optimization', type: 'YouTube' },
        { title: 'ভিডিও-৪১: Specialized Profile তৈরি', type: 'YouTube' },
        { title: 'ভিডিও-৪২: সঠিক Job Search করার কৌশল', type: 'YouTube' },
        { title: 'ভিডিও-৪৩: Winning Proposal Writing', type: 'YouTube' },
        { title: 'ভিডিও-৪৪: Client Interview ও Professional Communication', type: 'YouTube' },
        { title: 'ভিডিও-৪৫: Upwork-এ প্রথম কাজ পাওয়ার কার্যকর কৌশল', type: 'YouTube' },
        { title: 'Module 07 Assessment Test', type: 'Link' }
      ]
    },
    {
      id: 'M8', order: 8, title: 'মডিউল ৮: Client Communication & Project Management', description: '৫টি ভিডিও',
      items: [
        { title: 'ভিডিও-৪৬: Professional Client Communication', type: 'YouTube' },
        { title: 'ভিডিও-৪৭: Project Discussion ও Requirement Analysis', type: 'YouTube' },
        { title: 'ভিডিও-৪৮: Revision Management ও Client Satisfaction', type: 'YouTube' },
        { title: 'ভিডিও-৪৯: Professional Project Delivery', type: 'YouTube' },
        { title: 'ভিডিও-৫০: ৫-স্টার Review অর্জনের কার্যকর কৌশল', type: 'YouTube' },
        { title: 'Module 08 Assessment Test', type: 'Link' }
      ]
    },
    {
      id: 'M9', order: 9, title: 'মডিউল ৯: Payment Method, Career Growth & Success', description: '৫টি ভিডিও',
      items: [
        { title: 'ভিডিও-৫১: Payoneer Account Setup ও Verification', type: 'YouTube' },
        { title: 'ভিডিও-৫২: Wise ও অন্যান্য Payment Method সম্পর্কে ধারণা', type: 'YouTube' },
        { title: 'ভিডিও-৫৩: Marketplace থেকে নিরাপদে Payment Withdraw করার নিয়ম', type: 'YouTube' },
        { title: 'ভিডিও-৫৪: Personal Branding ও Long-term Freelancing Career তৈরি', type: 'YouTube' },
        { title: 'ভিডিও-৫৫: মাসিক আয়ের পরিকল্পনা, Time Management ও Success Roadmap', type: 'YouTube' },
        { title: 'Module 09 Assessment Test', type: 'Link' }
      ]
    },
    {
      id: 'M10', order: 10, title: 'বোনাস মডিউল', description: '৫টি ভিডিও',
      items: [
        { title: 'ভিডিও-৫৬: Beginner Freelancer Checklist', type: 'YouTube' },
        { title: 'ভিডিও-৫৭: Free Learning Resources ও প্রয়োজনীয় Websites', type: 'YouTube' },
        { title: 'ভিডিও-৫৮: Productivity Tools ও AI Tools-এর কার্যকর ব্যবহার', type: 'YouTube' },
        { title: 'ভিডিও-৫৯: Beginner-দের সবচেয়ে সাধারণ ভুল এবং সেগুলো এড়ানোর উপায়', type: 'YouTube' },
        { title: 'ভিডিও-৬০: Final Career Guideline, Future Learning Roadmap এবং Next Steps', type: 'YouTube' },
        { title: 'Module 10 Assessment Test', type: 'Link' }
      ]
    }
  ];
}

// --- تنظیمات و ساختار دیتابیس ---
const CONFIG = {
  SHEETS: {
    PEOPLE: "افراد",
    ATT: "حضور و غیاب",
    SETTINGS: "تنظیمات",
    NOTES: "یادداشت_جلسات",
    PLANS: "طرح_درس_کامل",
    ESSENTIALS: "ملزومات",
    MI_CONFIG: "تنظیمات_هوش",
    MI_LOGS: "لاگ_هوش",
    ARCHIVE_LOGS: "سوابق_فعالیت"
  },
  DEFAULT_IMG: "https://cdn-icons-png.flaticon.com/512/847/847969.png"
};

const SHEET_STRUCTURE = {
  "افراد": ["نام", "غیبت", "تاخیر", "امتیاز", "وضعیت", "عکس", "امتیاز_دستی", "بیوگرافی", "تلفن_متربی", "تلفن_والدین", "تولد", "مدرسه", "پزشکی", "یادداشت_والدین"],
  "حضور و غیاب": ["نام"],
  "تنظیمات": ["کلید", "مقدار", "توضیحات"],
  "یادداشت_جلسات": ["تاریخ", "یادداشت"],
  "طرح_درس_کامل": ["تاریخ", "عنوان", "اولویت", "ماژول‌ها", "سین", "ID", "وضعیت"],
  "ملزومات": ["ID_طرح", "تاریخ", "عنوان", "اولویت", "نام_وسیله", "ID_یکتا", "وضعیت", "نوع"],
  "تنظیمات_هوش": ["نوع_هوش", "توضیحات_علمی", "رفتارهای_مثبت (+)", "رفتارهای_منفی (-)"],
  "لاگ_هوش": ["تاریخ", "نام_متربی", "نوع_هوش", "رفتار_مشاهده_شده", "امتیاز"],
  "سوابق_فعالیت": ["تاریخ", "نوع", "عنوان", "جزئیات_JSON"]
};

// --- توابع اصلی وب‌اپ ---

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('مدیریت هوشمند کلاس')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getAppData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  checkSheets(ss);
  updateCalculations(ss);

  const shP = ss.getSheetByName(CONFIG.SHEETS.PEOPLE);
  // اگر شیت افراد نبود، بازگشت ساختار خالی
  if (!shP) return { students: [], trend: {labels:[], data:[]}, notes: [], plans: [], settings: {}, moduleHistory: [] };

  const pData = shP.getDataRange().getValues();
  let students = [];
  if (pData.length > 1) {
    for (let i = 1; i < pData.length; i++) {
      if (pData[i][0]) {
        students.push({
          name: String(pData[i][0]),
          score: Number(pData[i][3]) || 0,
          manualScore: Number(pData[i][6]) || 0,
          image: fixUrl(pData[i][5]),
          fire: (Number(pData[i][1]) === 0 && Number(pData[i][2]) === 0)
        });
      }
    }
  }
  // مرتب‌سازی دانش‌آموزان بر اساس امتیاز (بیشترین بالا)
  students.sort((a, b) => b.score - a.score);

  const plans = getPlans(ss);

  return {
    students: students,
    trend: getTrendData(ss),
    topStudent: students.length > 0 ? students[0].name : "---",
    notes: getNotesList(ss),
    plans: plans,
    settings: getSystemSettings(ss),
    moduleHistory: getCombinedHistory(ss)
  };
}

function getFullStudentProfile(studentName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shP = ss.getSheetByName(CONFIG.SHEETS.PEOPLE);
  if (!shP) return { success: false, msg: "دیتابیس یافت نشد" };

  const pData = shP.getDataRange().getValues();
  let info = {};
  let found = false;

  for(let i=1; i<pData.length; i++) {
    if(String(pData[i][0]) === String(studentName)) {
      info = {
        name: pData[i][0],
        score: pData[i][3],
        image: pData[i][5] || "",
        displayImage: fixUrl(pData[i][5]),
        bio: pData[i][7] ? String(pData[i][7]) : "",
        phone: pData[i][8] ? String(pData[i][8]) : "",
        parentPhone: pData[i][9] ? String(pData[i][9]) : "",
        dob: pData[i][10] ? String(pData[i][10]) : "",
        school: pData[i][11] ? String(pData[i][11]) : "",
        medical: pData[i][12] ? String(pData[i][12]) : "",
        parentNote: pData[i][13] ? String(pData[i][13]) : ""
      };
      found = true;
      break;
    }
  }

  if (!found) return { success: false, msg: "دانش‌آموز یافت نشد" };

  return {
    success: true,
    info: info,
    mi: getStudentMIProfile(studentName),
    stats: getStudentDetails(studentName),
    fullHistory: getComprehensiveStudentHistory(ss, studentName)
  };
}

function getComprehensiveStudentHistory(ss, studentName) {
  let combined = [];

  // 1. MI Logs
  const miProfile = getStudentMIProfile(studentName);
  miProfile.history.forEach(h => {
    let statusText = h.score === 1 ? "✅ مثبت" : (h.score === -1 ? "❌ منفی" : "⚪ خنثی");
    combined.push({
      date: h.date,
      type: 'mi',
      title: h.type,
      desc: h.behavior + " (" + statusText + ")",
      score: h.score,
      icon: '🧠'
    });
  });

  // 2. Attendance
  const details = getStudentDetails(studentName);
  details.history.forEach(h => {
    let icon = '📅';
    if(h.status.includes('حاضر')) icon = '✅';
    else if(h.status.includes('غیبت')) icon = '❌';
    else if(h.status.includes('تاخیر')) icon = '⏰';

    combined.push({
      date: h.date,
      type: 'att',
      title: 'حضور و غیاب',
      desc: h.status,
      icon: icon
    });
  });

  // 3. General Activities (if present on that day)
  const allActivities = getCombinedHistory(ss);
  const attDates = details.history.filter(h => h.status.includes('حاضر') || h.status.includes('تاخیر')).map(h => normalizeDateStr(h.date));

  allActivities.forEach(act => {
    if(attDates.includes(normalizeDateStr(act.date))) {
      combined.push({
        date: act.date,
        type: 'activity',
        title: act.title,
        desc: act.type === 'grouping' ? 'شرکت در گروه‌بندی' : (act.details || 'شرکت در فعالیت کلاسی'),
        icon: '🎯'
      });
    }
  });

  // Sort by date descending
  return combined.sort((a, b) => b.date.localeCompare(a.date));
}

// ==========================================
//          سیستم تلگرام (بدون تغییر منطق)
// ==========================================

function doPost(e) {
  const output = ContentService.createTextOutput("ok");
  if (!e || !e.postData) return output;

  try {
    const update = JSON.parse(e.postData.contents);
    if (update.callback_query) {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const settings = getSystemSettings(ss);
      const token = settings.BOT_TOKEN;
      const cb = update.callback_query;

      if(token) {
        try {
          UrlFetchApp.fetch(`https://api.telegram.org/bot${token}/answerCallbackQuery`, {
            method: 'post', contentType: 'application/json',
            payload: JSON.stringify({ callback_query_id: cb.id, text: "بررسی شد." })
          });
        } catch(err) { Logger.log("AnswerCallback Error: " + err); }
      }

      if (cb.data.startsWith("DONE_")) {
        const itemId = cb.data.split("_")[1];
        const lock = LockService.getScriptLock();
        if (lock.tryLock(5000)) {
          try {
            const shE = ss.getSheetByName(CONFIG.SHEETS.ESSENTIALS);
            const vals = shE.getDataRange().getValues();
            let foundRow = -1;
            for (let i = 1; i < vals.length; i++) {
              if (String(vals[i][5]).trim() === String(itemId).trim()) { foundRow = i + 1; break; }
            }
            if (foundRow !== -1) {
              shE.getRange(foundRow, 7).setValue(1);
              let oldKeyboard = cb.message.reply_markup.inline_keyboard;
              let newKeyboard = [];
              for (let i = 0; i < oldKeyboard.length; i++) {
                let newRow = [];
                for (let j = 0; j < oldKeyboard[i].length; j++) {
                  if (oldKeyboard[i][j].callback_data !== cb.data) newRow.push(oldKeyboard[i][j]);
                }
                if (newRow.length > 0) newKeyboard.push(newRow);
              }
              if (newKeyboard.length > 0) {
                editMessageReplyMarkup(token, cb.message.chat.id, cb.message.message_id, {inline_keyboard: newKeyboard});
              } else {
                editMessageText(token, cb.message.chat.id, cb.message.message_id, "🎉 همه موارد خریداری شد.");
              }
            }
          } catch (innerE) { Logger.log("Lock Logic Error: " + innerE); } finally { lock.releaseLock(); }
        }
      }
    }
  } catch (err) { Logger.log("doPost Error: " + err.message); }
  return output;
}

function setupTelegramWebhook() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = getSystemSettings(ss);
  const token = settings.BOT_TOKEN;
  const url = ScriptApp.getService().getUrl();
  if(!token) return "خطا: توکن نیست.";
  try { return "نتیجه: " + UrlFetchApp.fetch(`https://api.telegram.org/bot${token}/setWebhook?url=${url}`).getContentText(); }
  catch(e) { return "خطا: " + e.message; }
}

function editMessageText(token, chatId, messageId, text) { try { UrlFetchApp.fetch(`https://api.telegram.org/bot${token}/editMessageText`, { method: 'post', contentType: 'application/json', payload: JSON.stringify({ chat_id: chatId, message_id: messageId, text: text }) }); } catch(e){} }
function editMessageReplyMarkup(token, chatId, messageId, replyMarkup) { try { UrlFetchApp.fetch(`https://api.telegram.org/bot${token}/editMessageReplyMarkup`, { method: 'post', contentType: 'application/json', payload: JSON.stringify({ chat_id: chatId, message_id: messageId, reply_markup: replyMarkup }) }); } catch(e){} }
function sendTelegramMsgWithBtn(token, chatId, text, markup) { try { UrlFetchApp.fetch(`https://api.telegram.org/bot${token}/sendMessage`, { method: 'post', contentType: 'application/json', payload: JSON.stringify({ chat_id: chatId, text: text, reply_markup: markup, parse_mode: 'Markdown' }) }); } catch(e){} }

// ==========================================

function checkDailyReminders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const todayStr = getTodayStr();
  const shE = ss.getSheetByName(CONFIG.SHEETS.ESSENTIALS);
  const shP = ss.getSheetByName(CONFIG.SHEETS.PLANS);
  if(!shE || !shP) return;
  const essentials = shE.getDataRange().getValues();
  const plansData = shP.getDataRange().getDisplayValues();
  let shoppingList = {};
  for(let i=1; i<essentials.length; i++) {
    let [pid, , , , item, uid, status, type] = essentials[i];
    if(type === 'buy' && status != 1 && String(status).toLowerCase() !== "true") {
      if(!shoppingList[pid]) shoppingList[pid] = [];
      shoppingList[pid].push({name: item, id: uid});
    }
  }
  const settings = getSystemSettings(ss);
  const token = settings.BOT_TOKEN;
  const chats = String(settings.ADMIN_CHAT_IDS || "").split(",").map(id => id.trim()).filter(id => id !== "");
  if(!token || chats.length === 0) return;
  plansData.forEach((p, index) => {
    if(index === 0) return;
    let pid = p[5], pDate = normalizeDateStr(p[0]), pPrio = p[2], planTitle = p[1];
    let isToday = (pDate === todayStr);
    if (pDate !== "" && pDate < todayStr && !isToday) return;
    if(shoppingList[pid] && shoppingList[pid].length > 0) {
      if(isToday || pPrio === 'high') {
        let itemsMsg = "", keyboard = [], currentRow = [];
        let list = shoppingList[pid];
        for(let j=0; j<list.length; j++) {
            itemsMsg += `\n${j+1}. ${list[j].name}`;
            currentRow.push({text: `✅ ${j+1}`, callback_data: `DONE_${list[j].id}`});
            if(currentRow.length === 4) { keyboard.push(currentRow); currentRow = []; }
        }
        if(currentRow.length > 0) keyboard.push(currentRow);
        let msg = `${isToday ? "🚨 *فوری*" : "🛒 *خرید*"} (${pDate})\n\n📌 *${planTitle}*${itemsMsg}`;
        chats.forEach(id => sendTelegramMsgWithBtn(token, id, msg, {inline_keyboard: keyboard}));
      }
    }
  });
}

// --- سایر توابع کمکی ---

function getCombinedHistory(ss) {
  let history = [];
  const shP = ss.getSheetByName(CONFIG.SHEETS.PLANS);
  if (shP) {
    const pData = shP.getDataRange().getDisplayValues();
    if (pData.length > 1) {
      for (let i = 1; i < pData.length; i++) {
        let date = normalizeDateStr(pData[i][0]);
        let planTitle = pData[i][1];
        let planId = pData[i][5];
        let modules = [];
        try { modules = JSON.parse(pData[i][3] || "[]"); } catch (e) { continue; }
        modules.forEach(m => {
          if(m.name && m.name.trim() !== "") history.push({date: date, source: 'plan', planTitle: planTitle, planId: planId, type: m.type, title: m.name, details: m.desc || ""});
        });
      }
    }
  }
  const shA = ss.getSheetByName(CONFIG.SHEETS.ARCHIVE_LOGS);
  if (shA) {
    const aData = shA.getDataRange().getDisplayValues();
    if (aData.length > 1) {
      for (let i = 1; i < aData.length; i++) {
        history.push({date: normalizeDateStr(aData[i][0]), source: 'manual', planTitle: 'فعالیت ثبت شده دستی', type: aData[i][1], title: aData[i][2], details: aData[i][3]});
      }
    }
  }
  return history.sort((a,b) => b.date.localeCompare(a.date));
}

function saveManualLog(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CONFIG.SHEETS.ARCHIVE_LOGS);
  if(sh) sh.appendRow([data.date, data.type, data.title, data.details]);
  return {success: true, msg: "✅ ثبت شد."};
}

function checkSheets(ss) {
  for (let sheetName in SHEET_STRUCTURE) {
    let sh = ss.getSheetByName(sheetName);
    if (!sh) {
      sh = ss.insertSheet(sheetName);
      sh.appendRow(SHEET_STRUCTURE[sheetName]);
      if(sheetName === CONFIG.SHEETS.SETTINGS) {
        sh.appendRow(["BOT_TOKEN", "", "توکن ربات تلگرام"]);
        sh.appendRow(["ADMIN_CHAT_IDS", "", "آیدی مدیران تلگرام"]);
      }
      if(sheetName === CONFIG.SHEETS.MI_CONFIG) fillMIData(sh);
    } else {
        if(sheetName === CONFIG.SHEETS.SETTINGS) {
            const data = sh.getDataRange().getValues();
            const keys = data.map(r => r[0]);
            if(!keys.includes("BOT_TOKEN")) sh.appendRow(["BOT_TOKEN", "", "توکن ربات تلگرام"]);
            if(!keys.includes("ADMIN_CHAT_IDS")) sh.appendRow(["ADMIN_CHAT_IDS", "", "آیدی مدیران تلگرام"]);
        }
        // اگر شیت هوش وجود داشت اما دیتا قدیمی بود، آپدیت کن
        if(sheetName === CONFIG.SHEETS.MI_CONFIG) {
          const firstVal = sh.getRange(2, 1).getValue();
          if(firstVal === "زبانی-کلامی 🗣️") fillMIData(sh);
        }
    }
  }
}

function fillMIData(sh) {
  const data = [
    ["هوش کلامی 🗣️", "افراد با هوش کلامی بالا، مجریان و سخنوران توانمندی هستند.", "پرگو و وراجی,استفاده از واژگان غنی,شنوندگان فعال,تکلم زودهنگام در کودکی,تعامل کلامی با بزرگسالان,صحبت با جزئیات,صدای جذاب و رسا,تمایل به آموزش در کلاس", ""],
    ["هوش منطقی-ریاضی 🔢", "افراد با هوش منطقی-ریاضی اغلب مدرسین خود را به چالش می‌کشند.", "پرسش‌های مکرر و غیرمعمول,توانایی در دسته‌بندی,یادگیری سریع رنگ‌ها و اسامی,صحبت فراتر از سن,توانایی در فهم ریاضیات,علاقه‌مندی به بازی‌های فکری,علاقه به یادگیری,تمرکز بالا,نگاه انتقادی", ""],
    ["هوش بین فردی 🤝", "توانایی درک و تعامل مؤثر با دیگران.", "درک احساسات اطرافیان,برقراری ارتباط سریع از کودکی,واکنش‌های هیجانی مناسب,موفقیت در بیان,مسالمت‌جو در میان همسالان,سازش‌پذیری و معاشرت,دوستان زیاد و رهبری گروه,درک شرایط افراد", ""],
    ["هوش درون فردی 🧘", "توانایی شناخت خود و کنترل هیجانات درونی.", "آرامش و وقار,ترجیح تنهایی,شهود قوی,اعتماد به نفس بالا,کنترل احساسات,اصلاح رفتار,لذت از فعالیت‌های فردی,خلاقیت و تخیل قوی,پرسش‌های وجودی,رفتار و پوشش فراتر از سن", ""],
    ["هوش تصویری/فضایی 🎨", "توانایی تجسم اجسام و خلق آثار بصری.", "مهارت در نقاشی,علاقه به مونتاژ و دمونتاژ,توانایی در پازل,استعداد در ساخت لگو,موفقیت در فعالیت‌های دستی,عملکرد خوب در املا,حافظه تصویری قوی,علاقه به اسباب‌بازی‌های هندسی,توانایی در حل روبیک", ""],
    ["هوش جنبشی-حرکتی ⚽", "توانایی استفاده از مهارت‌های بدنی و حرکتی.", "تحرک زیاد,حرکات تکراری,ترجیح بازی‌های حرکتی,توانمند در ورزش,داوطلب کمک,بی‌علاقگی به نشستن طولانی,واکنش‌های سریع,تشخیص بیش‌فعالی,چالش تعادلی", ""],
    ["هوش طبیعت‌گرا 🌿", "علاقه به شناخت و حفظ محیط زیست و موجودات زنده.", "علاقه به گیاهان و حیوانات,علاقه‌مند به مباحث حیات وحش,علاقه به حضور در طبیعت,علاقه به جمع‌آوری نمونه حشرات,حساسیت به حفظ طبیعت", ""],
    ["هوش موسیقیایی 🎵", "حساسیت به ریتم، آهنگ و صداها.", "تشخیص ریتم,زمزمه,آواز خواندن,درک تن صدا", "بی‌توجهی به صداها"]
  ];
  sh.getRange(2, 1, data.length, 4).setValues(data);
}

function updateCalculations(ss) {
  const shP = ss.getSheetByName(CONFIG.SHEETS.PEOPLE);
  const shA = ss.getSheetByName(CONFIG.SHEETS.ATT);
  if (!shP || !shA) return;
  const pD = shP.getDataRange().getValues();
  const aD = shA.getDataRange().getValues();
  let m = {};
  for(let i=1; i<pD.length; i++) { if(pD[i][0]) m[pD[i][0]] = {r: i+1, ab: 0, la: 0, sy: 0, ma: Number(pD[i][6]) || 0}; }
  if(aD.length > 0) { for(let c=1; c<aD[0].length; c++) { for(let r=1; r<aD.length; r++) { let n = aD[r][0]; let v = String(aD[r][c]); if(m[n]) { if(v.includes("غیبت")) { m[n].ab++; m[n].sy -= 1; } else if(v.includes("تاخیر")) { m[n].la++; } else if(v.includes("حاضر")) { m[n].sy += 1; } else if(v.includes("موجه")) { m[n].sy += 0.5; } } } } }
  for(let n in m) { shP.getRange(m[n].r, 2, 1, 3).setValues([[m[n].ab, m[n].la, m[n].sy + m[n].ma]]); }
}

function getSystemSettings(ss) {
  const sh = ss.getSheetByName(CONFIG.SHEETS.SETTINGS);
  if(!sh) return {};
  const d = sh.getDataRange().getValues();
  let s = {};
  for(let i=1; i<d.length; i++) s[d[i][0]] = d[i][1];
  return s;
}

function saveSystemSettings(f) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CONFIG.SHEETS.SETTINGS);
  const d = sh.getDataRange().getValues();
  for(let i=1; i<d.length; i++) {
    if(f[d[i][0]] !== undefined) sh.getRange(i+1, 2).setValue(f[d[i][0]]);
  }
  return {success: true, msg: "✅ تنظیمات ذخیره شد."};
}

function normalizeDateStr(d) {
  if(!d) return "";
  let s = String(d).replace(/[۰-۹]/g, c => '0123456789'['۰۱۲۳۴۵۶۷۸۹'.indexOf(c)]);
  let p = s.split('/');
  if(p.length !== 3) return s;
  return `${p[0]}/${p[1].padStart(2, '0')}/${p[2].padStart(2, '0')}`;
}

function getTodayStr() { return normalizeDateStr(new Date().toLocaleDateString('fa-IR')); }

function fixUrl(u) {
  if(!u) return CONFIG.DEFAULT_IMG;
  if(u.includes("/d/")) return "https://lh3.googleusercontent.com/d/" + u.split("/d/")[1].split("/")[0] + "=s400";
  return u;
}

function getNotesList(ss){
    const sh = ss.getSheetByName(CONFIG.SHEETS.NOTES);
    if (!sh) return [];
    const d=sh.getDataRange().getDisplayValues();
    let n=[];
    if(d.length>1)for(let i=1;i<d.length;i++)n.push({date:d[i][0],text:d[i][1]});
    return n.reverse();
}

// --- اصلاح شده برای مرتب‌سازی تاریخ (نزولی) ---
function getPlans(ss) {
    const sh = ss.getSheetByName(CONFIG.SHEETS.PLANS);
    if (!sh) return [];
    const d = sh.getDataRange().getDisplayValues();
    let plans = [];
    if(d.length > 1) {
        for(let i=1; i<d.length; i++) {
            plans.push({
                date: normalizeDateStr(d[i][0]), // استاندارد سازی تاریخ
                title: d[i][1],
                priority: d[i][2],
                modules: d[i][3],
                sin: d[i][4],
                id: d[i][5],
                status: d[i][6]
            });
        }
    }

    // مرتب‌سازی: جدیدترین تاریخ (بزرگترین رشته در فرمت YYYY/MM/DD) اول
    // قدیمی‌ترین تاریخ پایین لیست می‌رود
    plans.sort((a, b) => b.date.localeCompare(a.date));

    return plans;
}

// --- توابع ذخیره سازی اصلاح شده ---

function savePlan(id, date, title, priority, sin, modules) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const shP = ss.getSheetByName(CONFIG.SHEETS.PLANS);
    const shE = ss.getSheetByName(CONFIG.SHEETS.ESSENTIALS);

    let currentId = id;
    if (!currentId || currentId === "undefined" || currentId === "") {
        currentId = String(Date.now());
    }

    let rowIndex = -1;
    const data = shP.getDataRange().getValues();

    for(let i=1; i<data.length; i++) {
        if(String(data[i][5]).trim() === String(currentId).trim()) {
            rowIndex = i + 1;
            break;
        }
    }

    if(rowIndex === -1) {
        shP.appendRow([date, title, priority, modules, sin, currentId, 'Active']);
    } else {
        shP.getRange(rowIndex, 1, 1, 6).setValues([[date, title, priority, modules, sin, currentId]]);
    }

    // حذف ملزومات قدیمی
    const essData = shE.getDataRange().getValues();
    for(let i=essData.length-1; i>=1; i--) {
        if(String(essData[i][0]).trim() === String(currentId).trim()) {
            shE.deleteRow(i+1);
        }
    }

    // افزودن ملزومات جدید
    try {
        const mods = JSON.parse(modules);
        let newRows = [];
        mods.forEach(m => {
            if(m.items && m.items.length > 0) {
                m.items.forEach(it => {
                   newRows.push([currentId, date, title, priority, it.name, Math.floor(Math.random()*100000), 0, it.type]);
                });
            }
        });
        if(newRows.length > 0) shE.getRange(shE.getLastRow()+1, 1, newRows.length, 8).setValues(newRows);
    } catch(e) {}

    return {success: true, msg: "✅ طرح درس با موفقیت ذخیره شد."};
}

function submitAttendance(data) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const shA = ss.getSheetByName(CONFIG.SHEETS.ATT);

    const headers = shA.getRange(1, 1, 1, shA.getLastColumn() || 1).getValues()[0];
    let colIndex = -1;
    const today = normalizeDateStr(data.date);

    for(let i=1; i<headers.length; i++) {
        let hDate = headers[i] instanceof Date ? headers[i].toLocaleDateString('fa-IR') : headers[i];
        if(normalizeDateStr(hDate) === today) { colIndex = i + 1; break; }
    }

    if(colIndex === -1) {
        colIndex = (shA.getLastColumn() || 1) + 1;
        shA.getRange(1, colIndex).setValue(data.date);
    }

    const rows = shA.getDataRange().getValues();
    let nameRowMap = {};
    for(let i=1; i<rows.length; i++) nameRowMap[rows[i][0]] = i + 1;

    data.records.forEach(rec => {
        let r = nameRowMap[rec.name];
        if(!r) {
            shA.appendRow([rec.name]);
            r = shA.getLastRow();
            nameRowMap[rec.name] = r;
        }

        let statusText = "";
        if(rec.status === 'Present') statusText = "حاضر";
        else if(rec.status === 'Absent') statusText = "غیبت";
        else if(rec.status === 'Late') statusText = `تاخیر (${rec.min} دقیقه)`;
        else if(rec.status === 'Excused') statusText = "موجه";

        shA.getRange(r, colIndex).setValue(statusText);
    });

    updateCalculations(ss);
    return {success: true, msg: "✅ حضور و غیاب ثبت شد."};
}

function saveNote(d,t){SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.NOTES).appendRow([d,t]);return{success:true,msg:"یادداشت ذخیره شد"};}
function addStudent(n){SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.PEOPLE).appendRow([n,0,0,0,"ثبت نام","",0]);return{success:true};}
function delStudent(n){const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.PEOPLE);const d=sh.getDataRange().getValues();for(let i=0;i<d.length;i++)if(d[i][0]==n){sh.deleteRow(i+1);return{success:true};}}
function addManualScore(n,p){const ss=SpreadsheetApp.getActiveSpreadsheet();const sh=ss.getSheetByName(CONFIG.SHEETS.PEOPLE);const d=sh.getDataRange().getValues();for(let i=1;i<d.length;i++)if(d[i][0]==n){sh.getRange(i+1,7).setValue((Number(d[i][6])||0)+Number(p));updateCalculations(ss);return{success:true,msg:"✅ امتیاز ثبت شد"}}; return {success:false};}
function submitMILog(sn,it,bh,sc){SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.MI_LOGS).appendRow([new Date(),sn,it,bh,sc]);return{success:true,msg:"✅ ثبت شد"};}
function deletePlan(id){const ss=SpreadsheetApp.getActiveSpreadsheet();const sh=ss.getSheetByName(CONFIG.SHEETS.PLANS);const d=sh.getDataRange().getValues();for(let i=1;i<d.length;i++)if(String(d[i][5]).trim()==String(id).trim()){sh.deleteRow(i+1);const shE=ss.getSheetByName(CONFIG.SHEETS.ESSENTIALS);const edat=shE.getDataRange().getValues();for(let j=edat.length-1;j>=1;j--)if(String(edat[j][0]).trim()==String(id).trim())shE.deleteRow(j+1);return{success:true,msg:"🗑 طرح درس حذف شد"}}return{success:false}}
function updateStudentProfile(data){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const shP=ss.getSheetByName(CONFIG.SHEETS.PEOPLE);
  const rows=shP.getDataRange().getValues();
  for(let i=1; i<rows.length; i++){
    if(rows[i][0] == data.originalName){
      shP.getRange(i+1,6).setValue(data.image);
      shP.getRange(i+1,8).setValue(data.bio || "");
      shP.getRange(i+1,9).setValue(data.phone || "[]");
      shP.getRange(i+1,11).setValue(data.dob || "");
      shP.getRange(i+1,12).setValue(data.school || "");
      shP.getRange(i+1,13).setValue(data.medical || "");
      shP.getRange(i+1,14).setValue(data.parentNote || "");
      return{success:true,msg:"✅ بروزرسانی شد"};
    }
  }
  return {success:false};
}
function updateLastStatus(sh, n, s) { const d = sh.getDataRange().getValues(); for(let i=0; i<d.length; i++) if(d[i][0] == n) { sh.getRange(i+1, 5).setValue(s); break; } }
function getTrendData(ss) {
  const shA = ss.getSheetByName(CONFIG.SHEETS.ATT);
  if(!shA) return {labels:[], data:[], details: []};
  const data = shA.getDataRange().getValues();
  let labels = [], trend = [], details = [];
  if(data.length > 0) {
    const headers = data[0];
    for(let c=Math.max(1, headers.length-10); c<headers.length; c++){
      labels.push(headers[c] instanceof Date ? headers[c].toLocaleDateString('fa-IR') : String(headers[c]));
      let p=0, a=0, l=0, e=0, t=0;
      for(let r=1; r<data.length; r++) {
        if(data[r][c]) {
          t++;
          let val = String(data[r][c]);
          if(val.includes("حاضر")) p++;
          else if(val.includes("غیبت")) a++;
          else if(val.includes("تاخیر")) l++;
          else if(val.includes("موجه")) e++;
        }
      }
      trend.push(t>0 ? Math.round((p/t)*100) : 0);
      details.push({p, a, l, e, total: t});
    }
  }
  return {labels, data: trend, details};
}
function getStudentDetails(n){ const ss=SpreadsheetApp.getActiveSpreadsheet(); const shP=ss.getSheetByName(CONFIG.SHEETS.PEOPLE); const shA=ss.getSheetByName(CONFIG.SHEETS.ATT); if(!shA) return {history:[],stats:{p:0,a:0,l:0,e:0},scores:{system:0,manual:0,total:0},growth:{labels:[],data:[]}}; const pD=shP.getDataRange().getValues(); let m=0; const rP=pD.find(r=>r[0]==n); if(rP) m=Number(rP[6])||0; const d=shA.getDataRange().getValues(); let h=[],s={p:0,a:0,l:0,e:0},gl=[],gd=[],sys=0,ri=-1; for(let i=1;i<d.length;i++)if(d[i][0]==n){ri=i;break} if(ri>-1){ const he=d[0]; for(let c=1;c<he.length;c++){ let v=String(d[ri][c]),da=(he[c]instanceof Date)?he[c].toLocaleDateString('fa-IR'):String(he[c]); if(v&&v!=""){ if(v.includes("حاضر")){sys++;s.p++}else if(v.includes("غیبت")){sys--;s.a++}else if(v.includes("تاخیر")){s.l++} gl.push(da); gd.push(sys+m) } } for(let c=he.length-1;c>=1;c--){ let v=String(d[ri][c]),da=(he[c]instanceof Date)?he[c].toLocaleDateString('fa-IR'):String(he[c]); if(v!="") h.push({date:da,status:v}) } } return {history:h,stats:s,scores:{system:sys,manual:m,total:sys+m},growth:{labels:gl,data:gd}}; }
function getStudentMIProfile(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shLogs = ss.getSheetByName(CONFIG.SHEETS.MI_LOGS);
  const shConf = ss.getSheetByName(CONFIG.SHEETS.MI_CONFIG);
  if(!shLogs || !shConf) return {config:[], chart:{labels:[],data:[]}, history:[], traitStates:{}};

  const cData = shConf.getDataRange().getValues();
  let confArr = [];
  for(let i=1; i<cData.length; i++) {
    confArr.push({
      type: cData[i][0],
      desc: cData[i][1],
      pos: cData[i][2] ? String(cData[i][2]).split(',') : [],
      neg: cData[i][3] ? String(cData[i][3]).split(',') : []
    });
  }

  const lData = shLogs.getDataRange().getValues();
  let traitStates = {};
  let history = [];

  for(let i=1; i<lData.length; i++) {
    if(String(lData[i][1]) === String(name)) {
      let type = lData[i][2];
      let behavior = lData[i][3];
      let score = Number(lData[i][4]);

      // ذخیره آخرین وضعیت برای هر رفتار
      traitStates[type + "_" + behavior] = score;

      history.push({
        date: new Date(lData[i][0]).toLocaleDateString('fa-IR'),
        type: type,
        behavior: behavior,
        score: score
      });
    }
  }

  let scores = {};
  confArr.forEach(c => {
    scores[c.type] = 0;
    // جمع امتیازات بر اساس آخرین وضعیت هر ویژگی تعریف شده
    c.pos.forEach(p => {
      scores[c.type] += (traitStates[c.type + "_" + p] || 0);
    });
    c.neg.forEach(n => {
      scores[c.type] += (traitStates[c.type + "_" + n] || 0);
    });
  });

  let chartData = []; let labels = [];
  confArr.forEach(c => {
    labels.push(c.type);
    chartData.push(Math.max(0, scores[c.type]));
  });

  return {
    config: confArr,
    chart: { labels, data: chartData },
    history: history.reverse().slice(0, 20),
    traitStates: traitStates
  };
}

function getClassMIData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shLogs = ss.getSheetByName(CONFIG.SHEETS.MI_LOGS);
  const shConf = ss.getSheetByName(CONFIG.SHEETS.MI_CONFIG);
  const shP = ss.getSheetByName(CONFIG.SHEETS.PEOPLE);
  if(!shLogs || !shConf || !shP) return {labels: [], averages: []};

  const students = shP.getDataRange().getValues().slice(1).map(r => r[0]).filter(n => n);
  const studentCount = Math.max(1, students.length);

  const cData = shConf.getDataRange().getValues();
  let labels = [];
  let confArr = [];
  for(let i=1; i<cData.length; i++) {
    labels.push(cData[i][0]);
    confArr.push({
      type: cData[i][0],
      traits: (cData[i][2] ? String(cData[i][2]).split(',') : []).concat(cData[i][3] ? String(cData[i][3]).split(',') : [])
    });
  }

  const lData = shLogs.getDataRange().getValues();
  let studentTraitStates = {}; // { "Student|Type|Trait": score }

  for(let i=1; i<lData.length; i++) {
    let sName = lData[i][1];
    let type = lData[i][2];
    let trait = lData[i][3];
    let score = Number(lData[i][4]);
    studentTraitStates[sName + "|" + type + "|" + trait] = score;
  }

  let totalScores = {};
  labels.forEach(l => totalScores[l] = 0);

  students.forEach(sName => {
    confArr.forEach(c => {
      let studentTypeScore = 0;
      c.traits.forEach(t => {
        studentTypeScore += (studentTraitStates[sName + "|" + c.type + "|" + t] || 0);
      });
      totalScores[c.type] += Math.max(0, studentTypeScore);
    });
  });

  let averages = labels.map(l => Number((totalScores[l] / studentCount).toFixed(2)));
  return {labels, averages};
}
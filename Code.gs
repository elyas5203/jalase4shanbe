// --- تنظیمات و ساختار دیتابیس ---
const FB_CONFIG = {
  URL: "[YOUR_FIREBASE_URL]", // e.g. https://your-db-name.firebaseio.com/
  SECRET: "[YOUR_FIREBASE_SECRET]"
};

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
  "افراد": ["نام", "غیبت", "تاخیر", "امتیاز", "وضعیت", "عکس", "امتیاز_دستی", "بیوگرافی", "تلفن_متربی", "تلفن_والدین", "تولد", "مدرسه", "پزشکی", "یادداشت_والدین", "مجموع_دقایق_تاخیر"],
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
  return getComprehensiveAppData();
}

/**
 * این تابع تمام اطلاعات سیستم را در یک ساختار JSON واحد جمع‌آوری می‌کند
 * برای سینک با فایربیس و استفاده در کلاینت بدون تاخیر
 */
function getComprehensiveAppData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  checkSheets(ss);
  updateCalculations(ss);

  const students = {};
  const shP = ss.getSheetByName(CONFIG.SHEETS.PEOPLE);
  if (!shP) return { students: {} };

  const pData = shP.getDataRange().getValues();

  // دریافت اطلاعات عمومی یکبار برای کل کلاس
  const notes = getNotesList(ss);
  const plans = getPlans(ss);
  const settings = getSystemSettings(ss);
  const moduleHistory = getCombinedHistory(ss);
  const trend = getTrendData(ss);
  const classMIData = getClassMIData();

  // دریافت لاگ‌های هوش یکبار برای پردازش
  const shLogs = ss.getSheetByName(CONFIG.SHEETS.MI_LOGS);
  const shConf = ss.getSheetByName(CONFIG.SHEETS.MI_CONFIG);
  const miLogs = shLogs ? shLogs.getDataRange().getValues() : [];
  const miConfRaw = shConf ? shConf.getDataRange().getValues() : [];
  const miConf = [];
  for(let i=1; i<miConfRaw.length; i++) {
    miConf.push({
      type: miConfRaw[i][0],
      desc: miConfRaw[i][1],
      pos: miConfRaw[i][2] ? String(miConfRaw[i][2]).split(',') : [],
      neg: miConfRaw[i][3] ? String(miConfRaw[i][3]).split(',') : []
    });
  }

  // دریافت اطلاعات حضور و غیاب
  const shA = ss.getSheetByName(CONFIG.SHEETS.ATT);
  const attData = shA ? shA.getDataRange().getValues() : [];

  for (let i = 1; i < pData.length; i++) {
    const name = String(pData[i][0]);
    if (!name) continue;

    // ۱. اطلاعات پایه
    const info = {
      name: name,
      absent: Number(pData[i][1]) || 0,
      late: Number(pData[i][2]) || 0,
      score: Number(pData[i][3]) || 0,
      image: pData[i][5] || "",
      displayImage: fixUrl(pData[i][5]),
      manualScore: Number(pData[i][6]) || 0,
      bio: pData[i][7] ? String(pData[i][7]) : "",
      phone: pData[i][8] ? String(pData[i][8]) : "[]",
      dob: pData[i][10] ? String(pData[i][10]) : "",
      school: pData[i][11] ? String(pData[i][11]) : "",
      medical: pData[i][12] ? String(pData[i][12]) : "",
      parentNote: pData[i][13] ? String(pData[i][13]) : "",
      totalDelayMins: Number(pData[i][14]) || 0,
      fire: (Number(pData[i][1]) === 0)
    };

    // ۲. پروفایل هوش برای این فرد
    const traitStates = {};
    const miHistory = [];
    if (miLogs.length > 0) {
      for(let j=1; j<miLogs.length; j++) {
        if(String(miLogs[j][1]) === name) {
          let type = miLogs[j][2];
          let behavior = miLogs[j][3];
          let score = Number(miLogs[j][4]);
          traitStates[type + "_" + behavior] = score;
          miHistory.push({ date: new Date(miLogs[j][0]).toLocaleDateString('fa-IR'), type, behavior, score });
        }
      }
    }

    const miScores = {};
    miConf.forEach(c => {
      miScores[c.type] = 0;
      c.pos.forEach(p => miScores[c.type] += (traitStates[c.type + "_" + p] || 0));
      c.neg.forEach(n => miScores[c.type] += (traitStates[c.type + "_" + n] || 0));
    });

    const studentMI = {
      chart: { labels: miConf.map(c => c.type), data: miConf.map(c => Math.max(0, miScores[c.type])) },
      history: miHistory.reverse().slice(0, 20),
      traitStates: traitStates
    };

    // ۳. آمار و تاریخچه حضور
    let ri = -1;
    for (let k = 1; k < attData.length; k++) if (String(attData[k][0]) === name) { ri = k; break; }

    const h = [], s = { p: 0, a: 0, l: 0, e: 0, lm: 0 }, gl = [], gd = [];
    let sys = 0;
    if (ri > -1) {
      const he = attData[0];
      for (let c = 1; c < he.length; c++) {
        let v = String(attData[ri][c]), da = (he[c] instanceof Date) ? he[c].toLocaleDateString('fa-IR') : String(he[c]);
        if (v && v != "") {
          if (v.includes("حاضر")) { sys++; s.p++ }
          else if (v.includes("غیبت")) { sys--; s.a++ }
          else if (v.includes("تاخیر")) {
            sys++; s.l++;
            let minsMatch = v.match(/\(([^)]+)\)/);
            if (minsMatch) {
              let minsStr = minsMatch[1].replace(/[^0-9۰-۹]/g, '');
              s.lm += parseInt(toEnglishDigits(minsStr)) || 0;
            }
          }
          else if (v.includes("موجه")) { s.e++ }
          gl.push(normalizeDateStr(da));
          gd.push(sys + info.manualScore);
        }
      }
      for (let c = he.length - 1; c >= 1; c--) {
        let v = String(attData[ri][c]), da = (he[c] instanceof Date) ? he[c].toLocaleDateString('fa-IR') : String(he[c]);
        if (v != "") h.push({ date: normalizeDateStr(da), status: v });
      }
    }

    // ۴. تاریخچه جامع
    const fullHistory = [];
    miHistory.forEach(mh => {
      let statusText = mh.score === 1 ? "✅ مثبت" : (mh.score === -1 ? "❌ منفی" : "⚪ خنثی");
      fullHistory.push({ date: normalizeDateStr(mh.date), type: 'mi', title: mh.type, desc: mh.behavior + " (" + statusText + ")", score: mh.score, icon: '🧠' });
    });
    h.forEach(att => fullHistory.push({ date: att.date, type: 'att', title: 'حضور و غیاب', desc: att.status, icon: att.status.includes('حاضر')?'✅':(att.status.includes('غیبت')?'❌':'⏰') }));

    const attDates = h.map(x => x.date);
    moduleHistory.forEach(act => {
      if(attDates.includes(act.date)) fullHistory.push({ date: act.date, type: 'activity', title: act.title, desc: act.type === 'grouping' ? 'شرکت در گروه‌بندی' : (act.details || 'شرکت در فعالیت کلاسی'), icon: '🎯' });
    });
    fullHistory.sort((a, b) => b.date.localeCompare(a.date));

    students[name] = {
      info: info,
      mi: studentMI,
      stats: { history: h, stats: s, scores: { system: sys, manual: info.manualScore, total: sys + info.manualScore }, growth: { labels: gl, data: gd } },
      fullHistory: fullHistory
    };
  }

  const sortedStudentList = Object.values(students).sort((a, b) => b.info.score - a.info.score);
  const maxScore = sortedStudentList.length > 0 ? sortedStudentList[0].info.score : -1;
  const topStudents = sortedStudentList.filter(s => s.info.score === maxScore && maxScore > 0).map(s => s.info.name).join(" - ");

  return {
    students: students,
    trend: trend,
    topStudent: topStudents || "---",
    notes: notes,
    plans: plans,
    settings: settings,
    moduleHistory: moduleHistory,
    classMI: classMIData,
    miConfig: miConf
  };
}

/**
 * همگام‌سازی اطلاعات شیت با فایربیس (REST API)
 */
function syncSheetToFirebase() {
  if (FB_CONFIG.URL.includes("[YOUR_")) return {success: false, msg: "تنظیمات فایربیس انجام نشده است"};

  const data = getComprehensiveAppData();
  const url = `${FB_CONFIG.URL}/classDB.json?auth=${FB_CONFIG.SECRET}`;
  const options = {
    method: 'put',
    contentType: 'application/json',
    payload: JSON.stringify(data)
  };

  try {
    UrlFetchApp.fetch(url, options);
    return { success: true, msg: "Firebase Synced 🔄" };
  } catch (e) {
    Logger.log("Sync Error: " + e.message);
    return { success: false, msg: e.message };
  }
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
      date: normalizeDateStr(h.date),
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
      date: normalizeDateStr(h.date),
      type: 'att',
      title: 'حضور و غیاب',
      desc: h.status,
      icon: icon
    });
  });

  // 3. General Activities (if present on that day)
  const allActivities = getCombinedHistory(ss);
  const attDates = details.history.map(h => normalizeDateStr(h.date));

  allActivities.forEach(act => {
    let normDate = normalizeDateStr(act.date);
    if(attDates.includes(normDate)) {
      combined.push({
        date: normDate,
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
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settings = getSystemSettings(ss);
    const token = settings.BOT_TOKEN;

    if (!token) return output;

    // Handle Message Commands
    if (update.message && update.message.text) {
      let txt = update.message.text;
      if (txt === "/start" || txt === "/status") {
        sendSimpleMsg(token, update.message.chat.id, "<b>🚀 سامانه مدیریت هوشمند کلاس</b>\n<b>────────────────</b>\n✅ وضعیت: <code>فعال و متصل</code>\n⏱ زمان: <code>" + new Date().toLocaleString('fa-IR') + "</code>\n<b>────────────────</b>\n✨ آماده دریافت گزارش‌ها.");
        return output;
      }
    }

    // Handle Callback Queries (Buttons)
    if (update.callback_query) {
      const cb = update.callback_query;

      // Feedback to User
      let feedbackText = "انجام شد.";
      if (cb.data.startsWith("DONE_")) feedbackText = "✅ مورد خریداری شد.";
      if (cb.data === "TEST_BTN") feedbackText = "🚀 دکمه با موفقیت کار می‌کند!";

      try {
        UrlFetchApp.fetch(`https://api.telegram.org/bot${token}/answerCallbackQuery`, {
          method: 'post', contentType: 'application/json',
          payload: JSON.stringify({
            callback_query_id: cb.id,
            text: feedbackText,
            show_alert: cb.data === "TEST_BTN"
          })
        });
      } catch(err) { Logger.log("AnswerCallback Error: " + err); }

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
                editMessageText(token, cb.message.chat.id, cb.message.message_id, "<b>✅ خرید تکمیل شد</b>\n<b>────────────────</b>\n🎉 تمامی موارد مورد نیاز برای این طرح درس خریداری و ثبت شدند.\n<b>────────────────</b>");
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

function editMessageText(token, chatId, messageId, text) { try { UrlFetchApp.fetch(`https://api.telegram.org/bot${token}/editMessageText`, { method: 'post', contentType: 'application/json', payload: JSON.stringify({ chat_id: chatId, message_id: messageId, text: text, parse_mode: 'HTML' }) }); } catch(e){} }
function editMessageReplyMarkup(token, chatId, messageId, replyMarkup) { try { UrlFetchApp.fetch(`https://api.telegram.org/bot${token}/editMessageReplyMarkup`, { method: 'post', contentType: 'application/json', payload: JSON.stringify({ chat_id: chatId, message_id: messageId, reply_markup: replyMarkup }) }); } catch(e){} }
function sendTelegramMsgWithBtn(token, chatId, text, markup) { try { UrlFetchApp.fetch(`https://api.telegram.org/bot${token}/sendMessage`, { method: 'post', contentType: 'application/json', payload: JSON.stringify({ chat_id: chatId, text: text, reply_markup: markup, parse_mode: 'HTML', disable_web_page_preview: true }) }); } catch(e){} }
function sendSimpleMsg(token, chatId, text) { try { UrlFetchApp.fetch(`https://api.telegram.org/bot${token}/sendMessage`, { method: 'post', contentType: 'application/json', payload: JSON.stringify({ chat_id: chatId, text: text, parse_mode: 'HTML' }) }); } catch(e){} }

function notifyAdmins(msg, markup = null, photoUrl = null) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = getSystemSettings(ss);
  const token = settings.BOT_TOKEN;
  const chats = String(settings.ADMIN_CHAT_IDS || "").split(",").map(id => id.trim()).filter(id => id !== "");
  if(!token || chats.length === 0) return;

  const webAppUrl = ScriptApp.getService().getUrl();
  let finalMarkup = markup;

  if (!finalMarkup && webAppUrl) {
    finalMarkup = { inline_keyboard: [[{ text: "🌐 ورود به سامانه مدیریت", url: webAppUrl }]] };
  } else if (finalMarkup && webAppUrl) {
    let hasDash = false;
    finalMarkup.inline_keyboard.forEach(row => row.forEach(btn => { if(btn.url === webAppUrl) hasDash = true; }));
    if (!hasDash) finalMarkup.inline_keyboard.push([{ text: "🌐 ورود به سامانه مدیریت", url: webAppUrl }]);
  }

  chats.forEach(chatId => {
    try {
      let method = photoUrl ? "sendPhoto" : "sendMessage";
      let payload = {
        chat_id: chatId,
        parse_mode: 'HTML'
      };
      if (photoUrl) {
        payload.photo = photoUrl;
        payload.caption = msg;
      } else {
        payload.text = msg;
        payload.disable_web_page_preview = true;
      }
      if (finalMarkup) payload.reply_markup = finalMarkup;

      UrlFetchApp.fetch(`https://api.telegram.org/bot${token}/${method}`, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload)
      });
    } catch(e) {
      Logger.log("NotifyAdmins Error (" + chatId + "): " + e.message);
    }
  });
}

/**
 * ایجاد یک پیام زیبا و ساختاریافته برای تلگرام
 */
function sendBeautifulNotification(title, icon, sections, markup = null, photoUrl = null) {
  let msg = `<b>${icon} ${title}</b>\n`;
  msg += `<b>────────────────</b>\n`;

  sections.forEach(sec => {
    if (sec.label) {
      msg += `<b>${sec.label}:</b> <code>${sec.value}</code>\n`;
    } else if (sec.pre) {
      msg += `<b>${sec.title}:</b>\n<pre>${sec.pre}</pre>\n`;
    } else if (sec.italic) {
      msg += `<i>${sec.italic}</i>\n`;
    } else if (sec.raw) {
      msg += sec.raw + "\n";
    }
  });

  msg += `<b>────────────────</b>`;

  notifyAdmins(msg, markup, photoUrl);
}

function testTelegramConnection() {
  const markup = {
    inline_keyboard: [[{ text: "🔘 تست دکمه شیشه‌ای", callback_data: "TEST_BTN" }]]
  };

  sendBeautifulNotification("تست موفقیت‌آمیز اتصال", "🚀", [
    { italic: "سیستم مدیریت هوشمند کلاس با موفقیت به این بات متصل شد." },
    { label: "نسخه سیستم", value: "4.0.0" },
    { label: "وضعیت سرور", value: "Online 🟢" },
    { raw: "\n<i>برای اطمینان از کارکرد دکمه‌ها، روی دکمه زیر کلیک کنید:</i>" }
  ], markup);

  return {success: true, msg: "✅ پیام تست ارسال شد. لطفاً تلگرام خود را چک کنید."};
}

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
      if(isToday) {
        let itemsMsg = "", keyboard = [], currentRow = [];
        let list = shoppingList[pid];
        for(let j=0; j<list.length; j++) {
            itemsMsg += `\n${j+1}. ${escapeHtml(list[j].name)}`;
            currentRow.push({text: `✅ ${j+1}`, callback_data: `DONE_${list[j].id}`});
            if(currentRow.length === 4) { keyboard.push(currentRow); currentRow = []; }
        }
        if(currentRow.length > 0) keyboard.push(currentRow);
        let msg = `<b>${isToday ? "🚨 یادآوری فوری" : "🛒 لیست خرید"}</b>\n`;
        msg += `<b>────────────────</b>\n`;
        msg += `🗓 <b>تاریخ اجرا:</b> <code>${escapeHtml(pDate)}</code>\n`;
        msg += `📌 <b>طرح درس:</b> ${escapeHtml(planTitle)}\n`;
        msg += `<b>────────────────</b>\n`;
        msg += `🛍 <b>موارد مورد نیاز:</b>\n${itemsMsg}\n\n`;
        msg += `<i>برای تایید خرید، روی شماره مربوطه کلیک کنید:</i>`;

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
  if (sh) sh.appendRow([data.date, data.type, data.title, data.details]);

  let typeName = data.type === 'grouping' ? 'گروه‌بندی' : 'فعالیت آزاد';
  let typeIcon = data.type === 'grouping' ? '👥' : '📝';

  let sections = [
    { label: "عنوان", value: escapeHtml(data.title) },
    { label: "تاریخ", value: escapeHtml(data.date) }
  ];

  if (data.type === 'grouping') {
    try {
      let d = JSON.parse(data.details);
      if (d.isMultiGroup) {
        let grps = d.groups.map(g => `▫️ <b>${escapeHtml(g.name)}:</b>\n${escapeHtml(g.members.join('، '))}`).join('\n\n');
        sections.push({ raw: `<b>🔍 اعضای گروه‌ها:</b>\n${grps}` });
      } else {
        sections.push({ title: "جزئیات", pre: escapeHtml(data.details) });
      }
    } catch (e) { sections.push({ title: "جزئیات", pre: escapeHtml(data.details) }); }
  } else {
    sections.push({ italic: escapeHtml(data.details) });
  }

  sendBeautifulNotification(typeName, typeIcon, sections);

  syncSheetToFirebase();
  return { success: true, msg: "✅ ثبت شد." };
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
        // Ensure all columns exist
        let lastCol = sh.getLastColumn();
        if (lastCol > 0) {
          let headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
          let expectedHeaders = SHEET_STRUCTURE[sheetName];
          expectedHeaders.forEach(h => {
            if (!headers.includes(h)) {
              sh.getRange(1, lastCol + 1).setValue(h);
              lastCol++;
            }
          });
        }

        if(sheetName === CONFIG.SHEETS.SETTINGS) {
            const data = sh.getDataRange().getValues();
            const keys = data.map(r => r[0]);
            if(!keys.includes("BOT_TOKEN")) sh.appendRow(["BOT_TOKEN", "", "توکن ربات تلگرام"]);
            if(!keys.includes("ADMIN_CHAT_IDS")) sh.appendRow(["ADMIN_CHAT_IDS", "", "آیدی مدیران تلگرام"]);
        }
        if(sheetName === CONFIG.SHEETS.MI_CONFIG) {
          const firstVal = sh.getRange(2, 1).getValue();
          if(firstVal === "زبانی-کلامی 🗣️" || firstVal === "هوش کلامی 🗣️") {
             const rowValue = sh.getRange(2, 3).getValue();
             if(!rowValue.includes("(")) fillMIData(sh);
          }
        }
    }
  }
}

function fillMIData(sh) {
  const data = [
    ["هوش کلامی 🗣️", "افراد با هوش کلامی بالا، مجریان و سخنوران توانمندی هستند.", "پرگو و وراجی (تمایل زیاد به صحبت کردن),استفاده از واژگان غنی (دایره لغات وسیع و حاضرجوابی),شنوندگان فعال (دقت به صحبت‌های دیگران),تکلم زودهنگام در کودکی,تعامل کلامی با بزرگسالان در کودکی,صحبت با جزئیات (توضیحات و روایت‌های مفصل),صدای جذاب و رسا (لحن دلنشین و گیرا),تمایل به آموزش در کلاس", ""],
    ["هوش منطقی-ریاضی 🔢", "افراد با هوش منطقی-ریاضی اغلب مدرسین خود را به چالش می‌کشند.", "پرسش‌های مکرر و غیرمعمول (سوالات چالش‌برانگیز),توانایی در دسته‌بندی اطلاعات,یادگیری سریع رنگ‌ها و اسامی,صحبت فراتر از سن (گفتار پخته),توانایی در فهم ریاضیات,علاقه‌مندی به بازی‌های فکری و حل مسئله,علاقه زیاد به یادگیری,تمرکز بالا و تحرک کمتر,نگاه انتقادی به مسائل", ""],
    ["هوش بین فردی 🤝", "توانایی درک و تعامل مؤثر با دیگران.", "درک سریع احساسات اطرافیان,برقراری ارتباط سریع از کودکی,واکنش‌های هیجانی سنجیده و مناسب,موفقیت در بیان عقاید و احساسات,مسالمت‌جو و صلح‌جو در میان همسالان,سازش‌پذیری و معاشرت بالا,دوستان زیاد و توانایی رهبری گروه,درک بالای شرایط و احوال دیگران", ""],
    ["هوش درون فردی 🧘", "توانایی شناخت خود و کنترل هیجانات درونی.", "آرامش و وقار (متانت خاص),ترجیح تنهایی برای حل مسائل,شهود و حس ششم قوی,اعتماد به نفس بالا,توانایی بالا در کنترل احساسات (خشم),اصلاح رفتار پس از اشتباه,لذت از فعالیت‌های فردی (مطالعه و نوشتن),خلاقیت و تخیل بسیار قوی,طرح پرسش‌های وجودی و عمیق,رفتار و پوشش فراتر از سن", ""],
    ["هوش تصویری/فضایی 🎨", "توانایی تجسم اجسام و خلق آثار بصری.", "مهارت بالا در نقاشی,علاقه به مونتاژ و دمونتاژ اسباب‌بازی,توانایی عالی در حل پازل,استعداد در ساخت سازه‌های لگو,موفقیت در فعالیت‌های دستی (مجسمه‌سازی),عملکرد خوب در درس املا,حافظه تصویری قوی,علاقه به اسباب‌بازی‌های هندسی,توانایی در حل مکعب روبیک", ""],
    ["هوش جنبشی-حرکتی ⚽", "توانایی استفاده از مهارت‌های بدنی و حرکتی.", "تحرک زیاد نسبت به همسالان,حرکات تکراری دست و پا هنگام نشستن,ترجیح بازی‌های حرکتی به بازی‌های فکری,مستعد و توانمند در ورزش,داوطلب شدن برای کمک به دیگران,بی‌علاقگی به نشستن طولانی‌مدت,واکنش‌های سریع و ناگهانی,تشخیص بیش‌فعالی (احتمالی),به چالش کشیدن تعادل بدن", ""],
    ["هوش طبیعت‌گرا 🌿", "علاقه به شناخت و حفظ محیط زیست و موجودات زنده.", "علاقه زیاد به گیاهان و حیوانات (نگهداری در منزل),علاقه‌مند به مباحث حیات وحش و طبیعت,علاقه به حضور در طبیعت و فضاهای سرسبز,علاقه به جمع‌آوری نمونه‌های طبیعت (حشرات),حساسیت به حفظ طبیعت و نظافت محیط", ""],
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
  for(let i=1; i<pD.length; i++) {
    if(pD[i][0]) m[pD[i][0]] = {r: i+1, ab: 0, la: 0, lm: 0, sy: 0, ma: Number(pD[i][6]) || 0};
  }
  if(aD.length > 0) {
    for(let c=1; c<aD[0].length; c++) {
      for(let r=1; r<aD.length; r++) {
        let n = aD[r][0];
        let v = String(aD[r][c]);
        if(m[n]) {
          if(v.includes("غیبت")) { m[n].ab++; m[n].sy -= 1; }
          else if(v.includes("تاخیر")) {
            m[n].la++;
            m[n].sy += 1;
            let minsMatch = v.match(/\(([^)]+)\)/);
            if (minsMatch) {
              let minsStr = minsMatch[1].replace(/[^0-9۰-۹]/g, '');
              m[n].lm += parseInt(toEnglishDigits(minsStr)) || 0;
            }
          }
          else if(v.includes("حاضر")) { m[n].sy += 1; }
          else if(v.includes("موجه")) { m[n].sy += 0.5; }
        }
      }
    }
  }
  for(let n in m) {
    shP.getRange(m[n].r, 2, 1, 3).setValues([[m[n].ab, m[n].la, m[n].sy + m[n].ma]]);
    // Always check for the new column and set total minutes
    let headers = shP.getRange(1, 1, 1, shP.getLastColumn()).getValues()[0];
    let lmIdx = headers.indexOf("مجموع_دقایق_تاخیر");
    if (lmIdx !== -1) {
      shP.getRange(m[n].r, lmIdx + 1).setValue(m[n].lm);
    }
  }
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
  syncSheetToFirebase();
  return {success: true, msg: "✅ تنظیمات ذخیره شد."};
}

function normalizeDateStr(d) {
  if(!d) return "";
  let s = toEnglishDigits(String(d));
  let p = s.split('/');
  if(p.length !== 3) return s;
  return `${p[0]}/${p[1].padStart(2, '0')}/${p[2].padStart(2, '0')}`;
}

function toEnglishDigits(str) {
  if (typeof str !== 'string') str = String(str);
  return str.replace(/[۰-۹]/g, c => '0123456789'['۰۱۲۳۴۵۶۷۸۹'.indexOf(c)]);
}

function getTodayStr() { return normalizeDateStr(new Date().toLocaleDateString('fa-IR')); }

function generateAttendanceChartUrl(trendData) {
  if (!trendData || !trendData.labels || trendData.labels.length === 0) return null;

  const chartConfig = {
    type: 'line',
    data: {
      labels: trendData.labels,
      datasets: [{
        label: 'حضور %',
        data: trendData.data,
        borderColor: '#4f46e5',
        backgroundColor: 'rgba(79, 70, 229, 0.1)',
        fill: true,
        pointRadius: 4,
        lineTension: 0.4
      }]
    },
    options: {
      title: { display: true, text: 'روند حضور و غیاب کلاس' },
      scales: {
        yAxes: [{ ticks: { beginAtZero: true, max: 100 } }]
      }
    }
  };

  return "https://quickchart.io/chart?c=" + encodeURIComponent(JSON.stringify(chartConfig)) + "&w=500&h=300&bkg=white";
}

function escapeHtml(text) {
  if (!text) return "";
  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

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

function savePlan(id, date, title, sin, modules) {
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

    let priority = "low"; // Removed from UI, defaulting
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

    // Telegram Notification for Shopping List
    let buyList = [];
    try {
        const mods = JSON.parse(modules);
        mods.forEach(m => {
            if(m.items) m.items.forEach(it => { if(it.type === 'buy') buyList.push(it.name); });
        });
    } catch(e) {}

    let sections = [
      { label: "عنوان", value: escapeHtml(title) },
      { label: "تاریخ اجرا", value: escapeHtml(date) }
    ];

    if (sin && sin.trim() !== "") {
      sections.push({ title: "⏰ سین برنامه (زمان‌بندی)", pre: escapeHtml(sin) });
    }

    try {
      const mods = JSON.parse(modules);
      if (mods && mods.length > 0) {
        let modText = mods.map((m, i) => {
          let t = `<b>${i+1}. ${escapeHtml(m.name)}</b>`;
          if (m.desc) t += `\n<i>${escapeHtml(m.desc)}</i>`;
          return t;
        }).join('\n\n');
        sections.push({ raw: `<b>🧩 ماژول‌های برنامه:</b>\n${modText}` });
      }
    } catch(e) {}

    if(buyList.length > 0) {
        let listText = buyList.map((item, idx) => `${idx + 1}. ${escapeHtml(item)}`).join('\n');
        sections.push({ title: "🛒 لیست خرید لوازم", pre: listText });
    }

    sendBeautifulNotification("طرح درس جدید ثبت شد", "📚", sections);

    syncSheetToFirebase();
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

    let p = [], a = [], l = [], e = [];

    data.records.forEach(rec => {
        let r = nameRowMap[rec.name];
        if(!r) {
            shA.appendRow([rec.name]);
            r = shA.getLastRow();
            nameRowMap[rec.name] = r;
        }

        let statusText = "";
        if(rec.status === 'Present') {
          statusText = "حاضر";
          p.push(rec.name);
        }
        else if(rec.status === 'Absent') {
          statusText = "غیبت";
          a.push(rec.name);
        }
        else if(rec.status === 'Late') {
          let mins = toEnglishDigits(String(rec.min || "0")).replace(/[^0-9]/g, '');
          statusText = `تاخیر (${mins} دقیقه)`;
          l.push(`${rec.name} (${mins}د)`);
        }
        else if(rec.status === 'Excused') {
          statusText = "موجه";
          e.push(rec.name);
        }

        shA.getRange(r, colIndex).setValue(statusText);
    });

    // Update before calculating chart
    updateCalculations(ss);

    // Telegram Notification
    const trendData = getTrendData(ss);
    const photoUrl = generateAttendanceChartUrl(trendData);
    const currentRate = trendData.data.length > 0 ? trendData.data[trendData.data.length - 1] : 0;

    let sections = [
      { label: "تاریخ", value: escapeHtml(today) },
      { label: "درصد حضور این جلسه", value: currentRate + "%" }
    ];

    if(p.length) sections.push({ title: "✅ حضور به‌موقع", pre: escapeHtml(p.join('، ')) });
    if(l.length) sections.push({ title: "⏰ حضور با تاخیر", pre: escapeHtml(l.join('، ')) });
    if(a.length) sections.push({ title: "❌ غایبین", pre: escapeHtml(a.join('، ')) });
    if(e.length) sections.push({ title: "🏳️ غایبین موجه", pre: escapeHtml(e.join('، ')) });

    sendBeautifulNotification("گزارش حضور و غیاب", "📊", sections, null, photoUrl);
    syncSheetToFirebase();
    return {success: true, msg: "✅ حضور و غیاب ثبت شد."};
}
function saveNote(d, t) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.NOTES).appendRow([d, t]);

  sendBeautifulNotification("یادداشت جدید جلسه", "📒", [
    { label: "تاریخ", value: escapeHtml(d) },
    { italic: escapeHtml(t) }
  ]);

  syncSheetToFirebase();
  return { success: true, msg: "یادداشت ذخیره شد" };
}
function addStudent(n){SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.PEOPLE).appendRow([n,0,0,0,"ثبت نام","",0]); syncSheetToFirebase(); return{success:true};}
function delStudent(n){const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.PEOPLE);const d=sh.getDataRange().getValues();for(let i=0;i<d.length;i++)if(d[i][0]==n){sh.deleteRow(i+1); syncSheetToFirebase(); return{success:true};}}
function addManualScore(n, p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CONFIG.SHEETS.PEOPLE);
  const d = sh.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) {
    if (d[i][0] == n) {
      sh.getRange(i + 1, 7).setValue((Number(d[i][6]) || 0) + Number(p));
      updateCalculations(ss);

      sendBeautifulNotification("ثبت امتیاز دستی", "⭐", [
        { label: "متربی", value: escapeHtml(n) },
        { label: "تغییر امتیاز", value: (p > 0 ? '+' : '') + p },
        { label: "مجموع امتیازات", value: (Number(d[i][3]) || 0) + Number(p) }
      ]);
      syncSheetToFirebase();
      return { success: true, msg: "✅ امتیاز ثبت شد" };
    }
  }
  return { success: false };
}
function submitMILog(sn, it, bh, sc) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.MI_LOGS).appendRow([new Date(), sn, it, bh, sc]);

  let statusText = sc === 1 ? "✅ مثبت" : (sc === -1 ? "❌ منفی" : "⚪ خنثی");

  sendBeautifulNotification("ثبت هوش چندگانه", "🧠", [
    { label: "متربی", value: escapeHtml(sn) },
    { label: "نوع هوش", value: escapeHtml(it) },
    { label: "وضعیت", value: statusText },
    { italic: escapeHtml(bh) }
  ]);

  syncSheetToFirebase();
  return { success: true, msg: "✅ ثبت شد" };
}
function deletePlan(id){const ss=SpreadsheetApp.getActiveSpreadsheet();const sh=ss.getSheetByName(CONFIG.SHEETS.PLANS);const d=sh.getDataRange().getValues();for(let i=1;i<d.length;i++)if(String(d[i][5]).trim()==String(id).trim()){sh.deleteRow(i+1);const shE=ss.getSheetByName(CONFIG.SHEETS.ESSENTIALS);const edat=shE.getDataRange().getValues();for(let j=edat.length-1;j>=1;j--)if(String(edat[j][0]).trim()==String(id).trim())shE.deleteRow(j+1); syncSheetToFirebase(); return{success:true,msg:"🗑 طرح درس حذف شد"}}return{success:false}}
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
      syncSheetToFirebase();
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
function getStudentDetails(n) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shP = ss.getSheetByName(CONFIG.SHEETS.PEOPLE);
  const shA = ss.getSheetByName(CONFIG.SHEETS.ATT);
  if (!shA) return { history: [], stats: { p: 0, a: 0, l: 0, e: 0, lm: 0 }, scores: { system: 0, manual: 0, total: 0 }, growth: { labels: [], data: [] } };

  const pD = shP.getDataRange().getValues();
  let m = 0;
  const rP = pD.find(r => r[0] == n);
  if (rP) m = Number(rP[6]) || 0;

  const d = shA.getDataRange().getValues();
  let h = [], s = { p: 0, a: 0, l: 0, e: 0, lm: 0 }, gl = [], gd = [], sys = 0, ri = -1;

  for (let i = 1; i < d.length; i++) if (d[i][0] == n) { ri = i; break }

  if (ri > -1) {
    const he = d[0];
    for (let c = 1; c < he.length; c++) {
      let v = String(d[ri][c]), da = (he[c] instanceof Date) ? he[c].toLocaleDateString('fa-IR') : String(he[c]);
      if (v && v != "") {
        if (v.includes("حاضر")) { sys++; s.p++ }
        else if (v.includes("غیبت")) { sys--; s.a++ }
        else if (v.includes("تاخیر")) {
          sys++;
          s.l++;
          let minsMatch = v.match(/\(([^)]+)\)/);
          if (minsMatch) {
            let minsStr = minsMatch[1].replace(/[^0-9۰-۹]/g, '');
            s.lm += parseInt(toEnglishDigits(minsStr)) || 0;
          }
        }
        else if (v.includes("موجه")) { s.e++ }

        gl.push(normalizeDateStr(da));
        gd.push(sys + m)
      }
    }
    for (let c = he.length - 1; c >= 1; c--) {
      let v = String(d[ri][c]), da = (he[c] instanceof Date) ? he[c].toLocaleDateString('fa-IR') : String(he[c]);
      if (v != "") h.push({ date: normalizeDateStr(da), status: v })
    }
  }
  return { history: h, stats: s, scores: { system: sys, manual: m, total: sys + m }, growth: { labels: gl, data: gd } };
}
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
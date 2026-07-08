/**
 * @OnlyCurrentDoc
 */

var TELEGRAM_BOT_TOKEN = "8844519360:AAGsebksu24KjnxWvKovDOx1AxsLgXU_uO4";
var TELEGRAM_CHAT_ID = "-1001874067843";
var TELEGRAM_TOPIC_ID = "2596";
var TELEGRAM_TOPIC_URL = "https://t.me/c/1874067843/2596";

// ✅ استضافة الأيقونة على GitHub Pages (حل نهائي لمشكلة عدم ظهور الأيقونة
// عند "إضافة إلى الشاشة الرئيسية"، لأن نطاق script.google.com يُعامل أحياناً
// كخدمة Google عامة ويُظهر أيقونة افتراضية بدلاً من أيقونة التطبيق)
//
// بعد رفع ملفات الأيقونات (icon-72.png ... icon-512.png) و manifest.json
// إلى مستودع GitHub وتفعيل GitHub Pages، استبدل "YOUR_GITHUB_USERNAME" و
// "YOUR_REPO_NAME" أدناه بالقيم الحقيقية فقط - لا حاجة لأي تعديل آخر.
var GITHUB_ICON_BASE = "https://YOUR_GITHUB_USERNAME.github.io/YOUR_REPO_NAME";
var APP_ICON_URL = GITHUB_ICON_BASE + "/icon-512.png";

function doGet(e) {
  var action = "";
  var params = (e && e.parameter) ? e.parameter : {};

  if (params.action) {
    action = params.action.toLowerCase().trim();
  }

  if (action && action !== "") {
    try {
      return handleAction(action, params);
    } catch (err) {
      return jsonResponse(false, "خطأ في التنفيذ: " + err.toString());
    }
  }

  return HtmlService.createHtmlOutput(getHtmlPage())
    .setTitle("تسجيل الحضور")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ✅ يتحقق من صيغة الوقت المرسل من المستخدم (HH:mm)
// إذا كان غير صالح أو غير موجود، يتم استخدام الوقت الحالي تلقائياً
function resolveChosenTime(paramTime, currentTime) {
  if (!paramTime) return currentTime;
  var t = paramTime.toString().trim();
  var match = t.match(/^([01]?\d|2[0-3]):([0-5]\d)$/);
  if (!match) return currentTime;
  var hh = match[1].length === 1 ? "0" + match[1] : match[1];
  return hh + ":" + match[2];
}

// ✅ يسترجع بيانات اليوم المخزنة (للحفاظ على كل معلومات رسالة Telegram)
// وإن لم توجد بيانات محفوظة (مثلاً بيانات قديمة قبل التحديث) يعيد بناءها من الشيت كحل احتياطي
function loadDayData(props, dataKey, name, date, sheet, rowNumber) {
  var raw = props.getProperty(dataKey);
  if (raw) {
    try {
      return JSON.parse(raw);
    } catch (e) {
      // تجاهل الخطأ والمتابعة للبناء الاحتياطي
    }
  }
  var inTime = sheet.getRange(rowNumber, 4).getDisplayValue();
  return {
    name: name,
    date: date,
    inTime: inTime || null,
    breakOutTime: null,
    breakInTime: null,
    breakMinutes: null,
    outTime: null
  };
}

function handleAction(action, params) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("💼 الحضور والانصراف");

  var cellA1Value = sheet.getRange("A1").getValue();
  var tz = ss.getSpreadsheetTimeZone();
  var todayStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
  var now = new Date();
  var currentTime = Utilities.formatDate(now, tz, "HH:mm");

  // ✅ الوقت الذي اختاره المستخدم من مربع اختيار الوقت (إن وُجد)، وإلا الوقت الحالي تلقائياً
  var chosenTime = resolveChosenTime(params && params.time, currentTime);

  var dataRange = sheet.getRange("C12:C43");
  var values = dataRange.getValues();
  var displayValues = dataRange.getDisplayValues();

  for (var i = 0; i < values.length; i++) {
    var cellValue = values[i][0];
    var cellText = displayValues[i][0];
    var matchFound = false;

    if (cellValue instanceof Date) {
      var formattedCellDate = Utilities.formatDate(cellValue, tz, "yyyy-MM-dd");
      if (formattedCellDate === todayStr) matchFound = true;
    } else if (cellText) {
      var cleanCellText = cellText.replace(/[^\d-]/g, "").trim();
      var todayReversed = todayStr.split("-").reverse().join("-");
      if (cleanCellText === todayStr || cleanCellText === todayReversed) matchFound = true;
    }

    if (matchFound) {
      var rowNumber = i + 12;
      var props = PropertiesService.getScriptProperties();
      var msgKey = "msg_" + todayStr + "_" + rowNumber;
      var dataKey = "data_" + todayStr + "_" + rowNumber;

      if (action === "in") {
        sheet.getRange(rowNumber, 4).setValue(chosenTime);

        var dayData = {
          name: cellA1Value,
          date: todayStr,
          inTime: chosenTime,
          breakOutTime: null,
          breakInTime: null,
          breakMinutes: null,
          outTime: null
        };
        props.setProperty(dataKey, JSON.stringify(dayData));

        var message = buildMessage(dayData);
        var msgId = sendTelegram(message, null);
        if (msgId) props.setProperty(msgKey, msgId.toString());
        return jsonResponse(true, "تم تسجيل الحضور!", chosenTime);
      }

      if (action === "break_out") {
        var cellF = sheet.getRange(rowNumber, 6);
        cellF.setNumberFormat("@");
        cellF.setValue("OUT_" + currentTime);

        var dayData = loadDayData(props, dataKey, cellA1Value, todayStr, sheet, rowNumber);
        dayData.breakOutTime = currentTime;
        props.setProperty(dataKey, JSON.stringify(dayData));

        var msgId = props.getProperty(msgKey);
        var message = buildMessage(dayData);
        editTelegram(msgId, message);

        return jsonResponse(true, "تم تسجيل الخروج من المكتب!", currentTime);
      }

      if (action === "break_in") {
        var cellF = sheet.getRange(rowNumber, 6);
        var fValue = cellF.getValue();
        var fText = fValue.toString();
        var breakOutMatch = fText.match(/OUT_(\d{1,2}:\d{2})/);

        if (breakOutMatch) {
          var breakOutTime = breakOutMatch[1];
          var breakOutDate = parseTimeString(breakOutTime);
          var breakInDate = parseTimeString(currentTime);

          if (breakOutDate && breakInDate) {
            var diffMs = breakInDate - breakOutDate;
            var diffMinutes = Math.round(diffMs / 60000);
            if (diffMinutes < 0) diffMinutes += 24 * 60;

            var existingMinutes = 0;
            var currentFVal = sheet.getRange(rowNumber, 6).getValue();
            if (typeof currentFVal === 'number' && currentFVal > 0) {
              existingMinutes = currentFVal;
            }
            var totalMinutes = existingMinutes + diffMinutes;

            cellF.setNumberFormat("0");
            cellF.setValue(totalMinutes);

            var dayData = loadDayData(props, dataKey, cellA1Value, todayStr, sheet, rowNumber);
            dayData.breakOutTime = breakOutTime;
            dayData.breakInTime = currentTime;
            dayData.breakMinutes = totalMinutes;
            props.setProperty(dataKey, JSON.stringify(dayData));

            var msgId = props.getProperty(msgKey);
            var message = buildMessage(dayData);
            editTelegram(msgId, message);

            return jsonResponse(true, "عدتَ للمكتب! غبت " + totalMinutes + " دقيقة", currentTime, totalMinutes);
          }
        }
        return jsonResponse(false, "خطأ في حساب الوقت");
      }

      if (action === "out") {
        sheet.getRange(rowNumber, 5).setValue(chosenTime);

        // ✅ لا يتم حذف أي بيانات سابقة (الحضور / الخروج المؤقت / العودة)
        // يتم فقط إضافة وقت الانصراف إلى نفس البيانات المحفوظة
        var dayData = loadDayData(props, dataKey, cellA1Value, todayStr, sheet, rowNumber);
        dayData.outTime = chosenTime;
        props.setProperty(dataKey, JSON.stringify(dayData));

        var msgId = props.getProperty(msgKey);
        var message = buildMessage(dayData);
        editTelegram(msgId, message);

        return jsonResponse(true, "تم تسجيل الانصراف!", chosenTime);
      }

      if (action === "clear_memory") {
        var allProps = props.getProperties();
        var count = 0;
        for (var key in allProps) {
          if (key.indexOf("msg_") === 0 || key.indexOf("data_") === 0) {
            props.deleteProperty(key);
            count++;
          }
        }
        return jsonResponse(true, "تم تفريغ الذاكرة! (" + count + " رسالة محذوفة)");
      }
    }
  }

  return jsonResponse(false, "لم يتم العثور على تاريخ اليوم في الجدول");
}

function convertDriveUrl(url) {
  if (!url) return "";
  var match1 = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (match1) return "https://lh3.googleusercontent.com/d/" + match1[1];
  var match2 = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (match2) return "https://lh3.googleusercontent.com/d/" + match2[1];
  if (url.includes("lh3.googleusercontent.com")) return url;
  if (url.startsWith("http")) return url;
  return "";
}

function getHtmlPage() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("💼 الحضور والانصراف");
  var employeeName = sheet.getRange("A1").getValue();
  var driveUrl = sheet.getRange("A2").getValue();
  var imageUrl = convertDriveUrl(driveUrl);
  var spreadsheetUrl = ss.getUrl();
  var scriptUrl = ScriptApp.getService().getUrl();

  // ✅ يدعم أي صيغة رابط (مشاركة Drive عادية أو رابط lh3 مباشر)
  var iconUrl = convertDriveUrl(APP_ICON_URL) || APP_ICON_URL;

  return `
<!DOCTYPE html>
<html dir="rtl" lang="ar">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">

  <!-- ✅ Favicon -->
  <link rel="icon" type="image/png" href="${iconUrl}">
  <link rel="shortcut icon" type="image/png" href="${iconUrl}">

  <!-- ✅ Apple Touch Icon - أيقونة الشاشة الرئيسية على iPhone -->
  <link rel="apple-touch-icon" href="${GITHUB_ICON_BASE}/icon-192.png">
  <link rel="apple-touch-icon" sizes="152x152" href="${GITHUB_ICON_BASE}/icon-152.png">
  <link rel="apple-touch-icon" sizes="180x180" href="${GITHUB_ICON_BASE}/icon-192.png">
  <link rel="apple-touch-icon" sizes="167x167" href="${GITHUB_ICON_BASE}/icon-192.png">
  <link rel="apple-touch-icon-precomposed" href="${GITHUB_ICON_BASE}/icon-192.png">

  <!-- ✅ Apple Splash Screen - شاشة البدء عند فتح التطبيق على iPhone/iPad -->
  <link rel="apple-touch-startup-image" href="${iconUrl}">
  <link rel="apple-touch-startup-image" media="(device-width: 320px) and (device-height: 568px)" href="${iconUrl}">
  <link rel="apple-touch-startup-image" media="(device-width: 375px) and (device-height: 667px)" href="${iconUrl}">
  <link rel="apple-touch-startup-image" media="(device-width: 414px) and (device-height: 896px)" href="${iconUrl}">
  <link rel="apple-touch-startup-image" media="(device-width: 390px) and (device-height: 844px)" href="${iconUrl}">
  <link rel="apple-touch-startup-image" media="(device-width: 428px) and (device-height: 926px)" href="${iconUrl}">

  <!-- ✅ PWA Manifest -->
  <link rel="manifest" href="data:application/json;base64,${Utilities.base64Encode(JSON.stringify({
    name: "تسجيل الحضور",
    short_name: "حضور",
    description: "تطبيق تسجيل الحضور والانصراف",
    start_url: scriptUrl,
    display: "standalone",
    background_color: "#0a0a0f",
    theme_color: "#ff2a6d",
    icons: [
      { src: GITHUB_ICON_BASE + "/icon-72.png",  sizes: "72x72",   type: "image/png", purpose: "any" },
      { src: GITHUB_ICON_BASE + "/icon-96.png",  sizes: "96x96",   type: "image/png", purpose: "any" },
      { src: GITHUB_ICON_BASE + "/icon-128.png", sizes: "128x128", type: "image/png", purpose: "any" },
      { src: GITHUB_ICON_BASE + "/icon-144.png", sizes: "144x144", type: "image/png", purpose: "any" },
      { src: GITHUB_ICON_BASE + "/icon-152.png", sizes: "152x152", type: "image/png", purpose: "any" },
      { src: GITHUB_ICON_BASE + "/icon-192.png", sizes: "192x192", type: "image/png", purpose: "any" },
      { src: GITHUB_ICON_BASE + "/icon-384.png", sizes: "384x384", type: "image/png", purpose: "any" },
      { src: GITHUB_ICON_BASE + "/icon-512.png", sizes: "512x512", type: "image/png", purpose: "any" },
      { src: GITHUB_ICON_BASE + "/icon-192.png", sizes: "192x192", type: "image/png", purpose: "maskable" },
      { src: GITHUB_ICON_BASE + "/icon-512.png", sizes: "512x512", type: "image/png", purpose: "maskable" }
    ]
  }))}">

  <!-- ✅ iOS Meta - لإنشاء اختصار على الشاشة الرئيسية -->
  <meta name="apple-mobile-web-app-capable" content="yes">
  <meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
  <meta name="apple-mobile-web-app-title" content="حضور">

  <!-- ✅ Theme -->
  <meta name="theme-color" content="#0a0a0f">
  <meta name="msapplication-TileColor" content="#0a0a0f">
  <meta name="msapplication-TileImage" content="${iconUrl}">
  <meta name="msapplication-square150x150logo" content="${iconUrl}">

  <!-- ✅ Android Icon (يُستخدم بواسطة Chrome عبر ملف الـ Manifest أعلاه) -->
  <link rel="icon" type="image/png" sizes="192x192" href="${iconUrl}">
  <link rel="icon" type="image/png" sizes="512x512" href="${iconUrl}">

  <title>تسجيل الحضور - ${employeeName}</title>
  <style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;900&display=swap');

    * { 
      margin: 0; 
      padding: 0; 
      box-sizing: border-box; 
      -webkit-tap-highlight-color: transparent;
      touch-action: manipulation;
    }

    :root {
      --gta-pink: #ff2a6d;
      --gta-cyan: #05d9e8;
      --gta-orange: #ff8c42;
      --gta-purple: #d300c5;
      --gta-yellow: #f9c80e;
      --gta-dark: #0a0a0f;
      --gta-card: rgba(255,255,255,0.04);
      --gta-border: rgba(255,255,255,0.08);
    }

    body {
      font-family: 'Cairo', sans-serif;
      background: var(--gta-dark);
      min-height: 100vh;
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: flex-start;
      padding: 16px;
      overflow-x: hidden;
    }

    .bg-orb {
      position: fixed;
      border-radius: 50%;
      filter: blur(80px);
      opacity: 0.15;
      pointer-events: none;
      z-index: 0;
    }
    .orb-1 {
      width: 300px; height: 300px;
      background: var(--gta-pink);
      top: -100px; right: -100px;
      animation: float 8s ease-in-out infinite;
    }
    .orb-2 {
      width: 250px; height: 250px;
      background: var(--gta-cyan);
      bottom: -80px; left: -80px;
      animation: float 10s ease-in-out infinite reverse;
    }
    .orb-3 {
      width: 200px; height: 200px;
      background: var(--gta-purple);
      top: 50%; left: 50%;
      animation: float 12s ease-in-out infinite 2s;
    }

    @keyframes float {
      0%, 100% { transform: translate(0, 0) scale(1); }
      33% { transform: translate(30px, -30px) scale(1.1); }
      66% { transform: translate(-20px, 20px) scale(0.9); }
    }

    .container {
      width: 100%;
      max-width: 420px;
      position: relative;
      z-index: 1;
    }

    .profile {
      text-align: center;
      margin-bottom: 16px;
      position: relative;
    }

    .profile-img-wrap {
      width: 90px;
      height: 90px;
      margin: 0 auto 10px;
      position: relative;
    }

    .profile-img-wrap::before {
      content: '';
      position: absolute;
      inset: -3px;
      border-radius: 50%;
      background: linear-gradient(135deg, var(--gta-pink), var(--gta-cyan), var(--gta-orange));
      animation: rotate 3s linear infinite;
    }

    @keyframes rotate {
      to { transform: rotate(360deg); }
    }

    .profile-img {
      width: 100%;
      height: 100%;
      border-radius: 50%;
      object-fit: cover;
      position: relative;
      z-index: 1;
      border: 3px solid var(--gta-dark);
      background: #1a1a2e;
    }

    .profile-img-placeholder {
      width: 100%;
      height: 100%;
      border-radius: 50%;
      background: linear-gradient(135deg, #1a1a2e, #2d2d44);
      display: flex;
      align-items: center;
      justify-content: center;
      font-size: 40px;
      position: relative;
      z-index: 1;
      border: 3px solid var(--gta-dark);
    }

    .profile-name {
      color: #fff;
      font-size: 22px;
      font-weight: 900;
      text-shadow: 0 0 20px rgba(255,42,109,0.3);
      letter-spacing: 1px;
    }

    .profile-status {
      color: rgba(255,255,255,0.35);
      font-size: 11px;
      font-weight: 600;
      margin-top: 4px;
      min-height: 16px;
      transition: all 0.3s ease;
    }

    .profile-status.active {
      color: var(--gta-cyan);
      text-shadow: 0 0 10px rgba(5,217,232,0.4);
      animation: pulse-text 1.5s ease-in-out infinite;
    }

    @keyframes pulse-text {
      0%, 100% { opacity: 0.6; }
      50% { opacity: 1; }
    }

    .time-card {
      background: var(--gta-card);
      border: 1px solid var(--gta-border);
      border-radius: 16px;
      padding: 16px;
      text-align: center;
      margin-bottom: 10px;
      backdrop-filter: blur(10px);
      position: relative;
      overflow: hidden;
    }

    .time-card::after {
      content: '';
      position: absolute;
      top: 0; left: 0; right: 0;
      height: 2px;
      background: linear-gradient(90deg, var(--gta-pink), var(--gta-cyan), var(--gta-orange));
    }

    .time-display {
      color: var(--gta-cyan);
      font-size: 40px;
      font-weight: 900;
      font-family: 'Cairo', sans-serif;
      letter-spacing: 4px;
      line-height: 1;
      text-shadow: 0 0 30px rgba(5,217,232,0.3);
    }

    .date-display {
      color: rgba(255,255,255,0.4);
      font-size: 12px;
      margin-top: 6px;
      font-weight: 600;
    }

    .buttons-grid {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 8px;
    }

    .btn {
      border: none;
      border-radius: 14px;
      padding: 14px 8px;
      font-size: 13px;
      font-weight: 700;
      font-family: 'Cairo', sans-serif;
      cursor: pointer;
      transition: all 0.15s ease;
      display: flex;
      flex-direction: column;
      align-items: center;
      gap: 4px;
      color: #fff;
      position: relative;
      overflow: hidden;
      background: var(--gta-card);
      border: 1px solid var(--gta-border);
      backdrop-filter: blur(10px);
    }

    .btn::before {
      content: '';
      position: absolute;
      inset: 0;
      opacity: 0;
      transition: opacity 0.2s;
    }

    .btn:active {
      transform: scale(0.96);
    }

    .btn:active::before {
      opacity: 0.15;
    }

    .btn .icon { 
      font-size: 22px; 
      filter: drop-shadow(0 2px 4px rgba(0,0,0,0.3));
    }

    .btn-in { 
      grid-column: 1 / -1; 
    }
    .btn-in::before {
      background: linear-gradient(135deg, #00b894, #00cec9);
    }
    .btn-in .icon { color: #00cec9; }

    .btn-break-out::before {
      background: linear-gradient(135deg, #fdcb6e, #e17055);
    }
    .btn-break-out .icon { color: #fdcb6e; }

    .btn-break-in::before {
      background: linear-gradient(135deg, #74b9ff, #a29bfe);
    }
    .btn-break-in .icon { color: #74b9ff; }

    .btn-out { 
      grid-column: 1 / -1; 
    }
    .btn-out::before {
      background: linear-gradient(135deg, #ff7675, #d63031);
    }
    .btn-out .icon { color: #ff7675; }

    .btn-clear {
      grid-column: 1 / -1;
      padding: 10px;
      font-size: 11px;
      margin-top: 4px;
      opacity: 0.6;
    }
    .btn-clear::before {
      background: linear-gradient(135deg, #636e72, #b2bec3);
    }
    .btn-clear .icon { color: #b2bec3; font-size: 16px; }

    .btn:disabled {
      opacity: 0.3 !important;
      cursor: not-allowed;
      transform: none !important;
    }

    .links-section {
      margin-top: 16px;
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 8px;
    }

    .link-btn {
      display: flex;
      align-items: center;
      justify-content: center;
      gap: 8px;
      padding: 12px 10px;
      border-radius: 12px;
      text-decoration: none;
      font-size: 12px;
      font-weight: 700;
      font-family: 'Cairo', sans-serif;
      color: #fff;
      transition: all 0.2s ease;
      border: 1px solid var(--gta-border);
      background: var(--gta-card);
      backdrop-filter: blur(10px);
    }

    .link-btn:active {
      transform: scale(0.96);
    }

    .link-btn .logo {
      width: 20px;
      height: 20px;
      object-fit: contain;
      filter: drop-shadow(0 1px 2px rgba(0,0,0,0.3));
    }

    .link-btn.sheets {
      background: linear-gradient(135deg, rgba(66,133,244,0.15), rgba(52,168,83,0.15));
      border-color: rgba(66,133,244,0.2);
    }
    .link-btn.sheets:hover {
      background: linear-gradient(135deg, rgba(66,133,244,0.25), rgba(52,168,83,0.25));
      box-shadow: 0 4px 20px rgba(66,133,244,0.15);
    }

    .link-btn.telegram {
      background: linear-gradient(135deg, rgba(0,136,204,0.15), rgba(0,178,255,0.15));
      border-color: rgba(0,136,204,0.2);
    }
    .link-btn.telegram:hover {
      background: linear-gradient(135deg, rgba(0,136,204,0.25), rgba(0,178,255,0.25));
      box-shadow: 0 4px 20px rgba(0,136,204,0.15);
    }

    .status-bar {
      margin-top: 12px;
      padding: 12px;
      border-radius: 12px;
      text-align: center;
      font-size: 13px;
      font-weight: 700;
      display: none;
      animation: popIn 0.3s cubic-bezier(0.175, 0.885, 0.32, 1.275);
      backdrop-filter: blur(10px);
    }

    .status-bar.success {
      display: block;
      background: rgba(0,206,201,0.1);
      color: #00cec9;
      border: 1px solid rgba(0,206,201,0.2);
    }

    .status-bar.error {
      display: block;
      background: rgba(255,118,117,0.1);
      color: #ff7675;
      border: 1px solid rgba(255,118,117,0.2);
    }

    @keyframes popIn {
      from { opacity: 0; transform: scale(0.9) translateY(10px); }
      to { opacity: 1; transform: scale(1) translateY(0); }
    }

    .btn-in { box-shadow: 0 4px 20px rgba(0,206,201,0.15); }
    .btn-break-out { box-shadow: 0 4px 20px rgba(253,203,110,0.1); }
    .btn-break-in { box-shadow: 0 4px 20px rgba(116,185,255,0.1); }
    .btn-out { box-shadow: 0 4px 20px rgba(255,118,117,0.15); }

    .btn-in:hover { box-shadow: 0 4px 30px rgba(0,206,201,0.25); }
    .btn-break-out:hover { box-shadow: 0 4px 30px rgba(253,203,110,0.2); }
    .btn-break-in:hover { box-shadow: 0 4px 30px rgba(116,185,255,0.2); }
    .btn-out:hover { box-shadow: 0 4px 30px rgba(255,118,117,0.25); }

    @media (prefers-reduced-motion: reduce) {
      *, *::before, *::after {
        animation-duration: 0.01ms !important;
        animation-iteration-count: 1 !important;
        transition-duration: 0.01ms !important;
      }
    }

    /* ✅ نافذة اختيار الوقت (Time Picker Modal) */
    .modal-overlay {
      position: fixed;
      inset: 0;
      background: rgba(0,0,0,0.6);
      backdrop-filter: blur(6px);
      display: flex;
      align-items: center;
      justify-content: center;
      z-index: 100;
      opacity: 0;
      pointer-events: none;
      transition: opacity 0.2s ease;
      padding: 16px;
    }

    .modal-overlay.show {
      opacity: 1;
      pointer-events: all;
    }

    .modal-card {
      width: 100%;
      max-width: 340px;
      background: #14141f;
      border: 1px solid var(--gta-border);
      border-radius: 18px;
      padding: 22px;
      text-align: center;
      position: relative;
      overflow: hidden;
      animation: popIn 0.3s cubic-bezier(0.175, 0.885, 0.32, 1.275);
    }

    .modal-card::before {
      content: '';
      position: absolute;
      top: 0; left: 0; right: 0;
      height: 2px;
      background: linear-gradient(90deg, var(--gta-pink), var(--gta-cyan), var(--gta-orange));
    }

    .modal-title {
      color: #fff;
      font-size: 17px;
      font-weight: 900;
      margin-bottom: 4px;
    }

    .modal-subtitle {
      color: rgba(255,255,255,0.4);
      font-size: 12px;
      font-weight: 600;
      margin-bottom: 18px;
    }

    .modal-time-input {
      width: 100%;
      background: var(--gta-card);
      border: 1px solid var(--gta-border);
      border-radius: 14px;
      padding: 14px;
      color: var(--gta-cyan);
      font-family: 'Cairo', sans-serif;
      font-size: 26px;
      font-weight: 900;
      text-align: center;
      letter-spacing: 2px;
      margin-bottom: 18px;
      outline: none;
      color-scheme: dark;
    }

    .modal-time-input::-webkit-calendar-picker-indicator {
      filter: invert(1);
      cursor: pointer;
    }

    .modal-actions {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 8px;
    }

    .modal-btn {
      border: none;
      border-radius: 12px;
      padding: 13px 8px;
      font-size: 14px;
      font-weight: 700;
      font-family: 'Cairo', sans-serif;
      cursor: pointer;
      transition: all 0.15s ease;
      color: #fff;
    }

    .modal-btn:active {
      transform: scale(0.96);
    }

    .modal-cancel {
      background: var(--gta-card);
      border: 1px solid var(--gta-border);
    }

    .modal-confirm {
      background: linear-gradient(135deg, #00b894, #00cec9);
    }
  </style>
</head>
<body>
  <div class="bg-orb orb-1"></div>
  <div class="bg-orb orb-2"></div>
  <div class="bg-orb orb-3"></div>

  <div class="container">
    <div class="profile">
      <div class="profile-img-wrap">
        ${imageUrl ? 
          `<img src="${imageUrl}" class="profile-img" alt="${employeeName}" onerror="this.style.display='none';this.parentElement.innerHTML='<div class=\'profile-img-placeholder\'>👤</div>'">` : 
          `<div class="profile-img-placeholder">👤</div>`
        }
      </div>
      <div class="profile-name">${employeeName}</div>
      <div class="profile-status" id="profileStatus">جاهز للتسجيل</div>
    </div>

    <div class="time-card">
      <div class="time-display" id="clock">--:--</div>
      <div class="date-display" id="date">--</div>
    </div>

    <div class="buttons-grid">
      <button class="btn btn-in" onclick="openTimeModal('in', this)">
        <span class="icon">📥</span>
        <span>تسجيل حضور</span>
      </button>

      <button class="btn btn-break-out" onclick="sendAction('break_out', this)">
        <span class="icon">🚪</span>
        <span>خروج</span>
      </button>

      <button class="btn btn-break-in" onclick="sendAction('break_in', this)">
        <span class="icon">✅</span>
        <span>عودة</span>
      </button>

      <button class="btn btn-out" onclick="openTimeModal('out', this)">
        <span class="icon">📤</span>
        <span>تسجيل انصراف</span>
      </button>

      <button class="btn btn-clear" onclick="sendAction('clear_memory', this)">
        <span class="icon">🗑️</span>
        <span>تفريغ الذاكرة</span>
      </button>
    </div>

    <div class="links-section">
      <a href="${spreadsheetUrl}" target="_blank" class="link-btn sheets">
        <svg class="logo" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path d="M14.5 2H6C4.9 2 4 2.9 4 4V20C4 21.1 4.9 22 6 22H18C19.1 22 20 21.1 20 20V7.5L14.5 2Z" fill="#34A853"/>
          <path d="M14.5 2V7.5H20" stroke="#fff" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/>
          <path d="M16 13H8M16 17H8M10 9H8" stroke="#fff" stroke-width="1.5" stroke-linecap="round"/>
        </svg>
        <span>جدول البيانات</span>
      </a>

      <a href="${TELEGRAM_TOPIC_URL}" target="_blank" class="link-btn telegram">
        <svg class="logo" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path d="M12 22C17.5228 22 22 17.5228 22 12C22 6.47715 17.5228 2 12 2C6.47715 2 2 6.47715 2 12C2 17.5228 6.47715 22 12 22Z" fill="#0088CC"/>
          <path d="M8.5 14.5L9.5 17.5L17 9.5L7 13L8.5 14.5Z" fill="#fff"/>
          <path d="M9.5 17.5L10.5 14.5L17 9.5" stroke="#fff" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/>
        </svg>
        <span>موضوع التيليجرام</span>
      </a>
    </div>

    <div class="status-bar" id="status"></div>
  </div>

  <!-- ✅ نافذة اختيار/تعديل وقت الحضور أو الانصراف -->
  <div class="modal-overlay" id="timeModalOverlay">
    <div class="modal-card">
      <div class="modal-title" id="modalTitle">تحديد الوقت</div>
      <div class="modal-subtitle">يمكنك تعديل الوقت أو تركه كما هو</div>
      <input type="time" id="modalTimeInput" class="modal-time-input">
      <div class="modal-actions">
        <button class="modal-btn modal-cancel" onclick="closeTimeModal()">إلغاء</button>
        <button class="modal-btn modal-confirm" onclick="confirmTimeModal()">تأكيد</button>
      </div>
    </div>
  </div>

  <script>
    var SCRIPT_URL = "${scriptUrl}";

    // ✅ أسماء الأشهر بالعربية (جوان، جويلية...)
    var arabicMonths = [
      "جانفي", "فيفري", "مارس", "أفريل", "ماي", "جوان",
      "جويلية", "أوت", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"
    ];

    // ✅ أسماء الأيام بالعربية
    var arabicDays = [
      "الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"
    ];

    function updateClock() {
      const now = new Date();
      const h = String(now.getHours()).padStart(2,'0');
      const m = String(now.getMinutes()).padStart(2,'0');
      // ✅ أرقام إنجليزية 12345
      document.getElementById('clock').textContent = h + ':' + m;

      // ✅ تاريخ بأشهر عربية (جوان، جويلية...)
      const dayName = arabicDays[now.getDay()];
      const day = now.getDate();
      const monthName = arabicMonths[now.getMonth()];
      const year = now.getFullYear();
      document.getElementById('date').textContent = dayName + '، ' + day + ' ' + monthName + ' ' + year;
    }
    updateClock();
    setInterval(updateClock, 1000);

    // ✅ إلغاء ميزة الضغط المزدوج - منع الضغط المتكرر
    var isProcessing = false;

    // ✅ متغيرات نافذة اختيار الوقت
    var pendingAction = null;
    var pendingBtn = null;

    function pad2(n) { return String(n).padStart(2, '0'); }

    // ✅ فتح نافذة اختيار الوقت مع عرض الوقت الحالي تلقائياً
    function openTimeModal(action, btn) {
      pendingAction = action;
      pendingBtn = btn;

      var now = new Date();
      var hh = pad2(now.getHours());
      var mm = pad2(now.getMinutes());
      document.getElementById('modalTimeInput').value = hh + ':' + mm;

      document.getElementById('modalTitle').textContent =
        action === 'in' ? '⏰ تحديد وقت الحضور' : '⏰ تحديد وقت الانصراف';

      document.getElementById('timeModalOverlay').classList.add('show');
    }

    function closeTimeModal() {
      document.getElementById('timeModalOverlay').classList.remove('show');
      pendingAction = null;
      pendingBtn = null;
    }

    // ✅ عند التأكيد: يتم استخدام الوقت المختار (المعدَّل أو الحالي إذا لم يُغيَّر)
    function confirmTimeModal() {
      var chosenTime = document.getElementById('modalTimeInput').value;
      var action = pendingAction;
      var btn = pendingBtn;

      document.getElementById('timeModalOverlay').classList.remove('show');

      if (!chosenTime) {
        var now = new Date();
        chosenTime = pad2(now.getHours()) + ':' + pad2(now.getMinutes());
      }

      sendAction(action, btn, chosenTime);
    }

    async function sendAction(action, btn, customTime) {
      // ✅ منع الضغط المتكرر
      if (isProcessing) return;
      isProcessing = true;

      const status = document.getElementById('status');
      const profileStatus = document.getElementById('profileStatus');
      const buttons = document.querySelectorAll('.btn');

      buttons.forEach(b => b.disabled = true);
      profileStatus.textContent = 'جاري التنفيذ...';
      profileStatus.classList.add('active');
      status.style.display = 'none';

      try {
        var requestUrl = SCRIPT_URL + '?action=' + action;
        if (customTime) requestUrl += '&time=' + encodeURIComponent(customTime);

        const response = await fetch(requestUrl);
        const contentType = response.headers.get('content-type');

        if (!contentType || !contentType.includes('application/json')) {
          const text = await response.text();
          throw new Error('الرد ليس JSON');
        }

        const data = await response.json();

        if (data.success) {
          status.className = 'status-bar success';
          var msg = data.message;
          if (data.time) msg += ' | ' + data.time;
          if (data.minutes !== undefined) msg += ' | ' + data.minutes + ' دقيقة';
          status.textContent = '✅ ' + msg;
          profileStatus.textContent = 'تم التسجيل';
          profileStatus.classList.remove('active');
          setTimeout(() => { profileStatus.textContent = 'جاهز للتسجيل'; }, 2000);
        } else {
          status.className = 'status-bar error';
          status.textContent = '❌ ' + data.message;
          profileStatus.textContent = 'خطأ في التسجيل';
          profileStatus.classList.remove('active');
        }
      } catch (err) {
        status.className = 'status-bar error';
        status.textContent = '❌ خطأ في الاتصال';
        profileStatus.textContent = 'خطأ في الاتصال';
        profileStatus.classList.remove('active');
        console.error(err);
      }

      setTimeout(() => {
        buttons.forEach(b => b.disabled = false);
        isProcessing = false;
      }, 1500);
    }
  </script>
</body>
</html>
  `;
}

function jsonResponse(success, message, time, minutes) {
  var obj = { success: success, message: message };
  if (time) obj.time = time;
  if (minutes !== undefined) obj.minutes = minutes;
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function sendTelegram(text, msgId) {
  var url = "https://api.telegram.org/bot" + TELEGRAM_BOT_TOKEN + "/sendMessage";

  var payload = {
    chat_id: TELEGRAM_CHAT_ID,
    message_thread_id: parseInt(TELEGRAM_TOPIC_ID),
    text: text
  };

  var options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var result = JSON.parse(response.getContentText());

    if (result.ok) return result.result.message_id;

    payload = { chat_id: TELEGRAM_CHAT_ID, text: text };
    options.payload = JSON.stringify(payload);
    response = UrlFetchApp.fetch(url, options);
    result = JSON.parse(response.getContentText());

    if (result.ok) return result.result.message_id;
    return null;

  } catch (err) {
    return null;
  }
}

function editTelegram(messageId, text) {
  if (!messageId) return false;

  var url = "https://api.telegram.org/bot" + TELEGRAM_BOT_TOKEN + "/editMessageText";
  var payload = { chat_id: TELEGRAM_CHAT_ID, message_id: parseInt(messageId), text: text };
  var options = { method: "post", contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var result = JSON.parse(response.getContentText());
    return result.ok;
  } catch (err) {
    return false;
  }
}

function buildMessage(data) {
  var text = "";
  text += "📊 " + data.name + "\n";
  text += "📅 التاريخ: " + data.date + "\n\n";
  text += "📥 الحضور: " + (data.inTime || "--:--") + "\n";
  if (data.breakOutTime) text += "🚪 خروج مؤقت: " + data.breakOutTime + "\n";
  if (data.breakInTime) text += "✅ عودة: " + data.breakInTime + "\n";
  if (data.breakMinutes !== null && data.breakMinutes !== undefined) {
    text += "⏱ مدة الغياب: " + data.breakMinutes + " دقيقة\n";
  }
  if (data.outTime) text += "📤 الانصراف: " + data.outTime + "\n";
  text += "\n✅ تم التوثيق";
  return text;
}

function parseTimeString(timeStr) {
  if (!timeStr) return null;
  var timeText = timeStr.toString().trim();
  var parts = timeText.split(":");
  if (parts.length >= 2) {
    var hours = parseInt(parts[0], 10);
    var minutes = parseInt(parts[1], 10);
    if (!isNaN(hours) && !isNaN(minutes)) {
      var date = new Date();
      date.setHours(hours, minutes, 0, 0);
      return date;
    }
  }
  return null;
}

/**
 * الشامل — إرسال إشعارات العروض تلقائياً من الشيت عبر OneSignal
 *
 * ماذا يفعل؟
 *  1) عند نزول عرض (وقت starts_at) يرسل إشعار "عرض جديد".
 *  2) قبل انتهاء العرض بـ 6 ساعات يرسل إشعار تذكير.
 *  3) يسجّل وقت الإرسال في عمودي notified_start و notified_6h حتى لا يتكرر.
 *
 * مهم: مفتاح REST لا يُكتب هنا. يُحفظ في Script Properties (انظر الشرح المرفق).
 */

const CFG = {
  APP_ID: "d8176c2f-571d-4fee-bf4b-5c738e304bd4",
  SITE_URL: "https://alshamilsyr.github.io",
  TZ: "Africa/Cairo",
  SEGMENT: "Total Subscriptions",   // الشريحة الافتراضية لكل المشتركين في OneSignal
  REMIND_HOURS: 6
};

function getKey_() {
  const k = PropertiesService.getScriptProperties().getProperty("ONESIGNAL_REST_KEY");
  if (!k) throw new Error("لم يتم حفظ ONESIGNAL_REST_KEY في Script Properties");
  return k.trim();
}

function sendPush_(title, body) {
  const payload = {
    app_id: CFG.APP_ID,
    target_channel: "push",
    included_segments: [CFG.SEGMENT],
    headings: { en: title, ar: title },
    contents: { en: body, ar: body },
    url: CFG.SITE_URL
  };
  const res = UrlFetchApp.fetch("https://api.onesignal.com/notifications?c=push", {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: "Key " + getKey_() },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  if (code < 200 || code >= 300) throw new Error("OneSignal " + code + ": " + res.getContentText());
  return JSON.parse(res.getContentText());
}

function pad_(n) { n = String(n); return n.length < 2 ? "0" + n : n; }

function toDate_(v) {
  if (v instanceof Date) return v;
  const s = String(v || "").trim();
  if (!s) return null;
  const m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2}))?/);
  if (!m) return null;
  const str = m[1] + "-" + pad_(m[2]) + "-" + pad_(m[3]) + " " + pad_(m[4] || "0") + ":" + (m[5] || "00");
  return Utilities.parseDate(str, CFG.TZ, "yyyy-MM-dd HH:mm");
}

function stamp_() {
  return Utilities.formatDate(new Date(), CFG.TZ, "yyyy-MM-dd HH:mm");
}

/** الدالة الرئيسية: تُشغَّل تلقائياً كل 15 دقيقة */
function checkOffers() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return;
  try {
    const ss = SpreadsheetApp.getActive();
    const offersSheet = ss.getSheetByName("Offers");
    const prodSheet = ss.getSheetByName("Products");
    if (!offersSheet || !prodSheet) throw new Error("تبويب Offers أو Products غير موجود");

    // أسماء المنتجات
    const pv = prodSheet.getDataRange().getValues();
    const ph = pv[0].map(h => String(h).trim().toLowerCase());
    const iId = ph.indexOf("id"), iName = ph.indexOf("name");
    const names = {};
    for (let r = 1; r < pv.length; r++) names[String(pv[r][iId]).trim()] = pv[r][iName];

    // العروض
    const data = offersSheet.getDataRange().getValues();
    const h = data[0].map(x => String(x).trim().toLowerCase());
    const c = n => h.indexOf(n);
    ["product_id", "discount_percent", "starts_at", "ends_at", "notified_start", "notified_6h"].forEach(n => {
      if (c(n) < 0) throw new Error("العمود " + n + " غير موجود في تبويب Offers");
    });

    const now = new Date();
    const remindMs = CFG.REMIND_HOURS * 3600 * 1000;

    for (let r = 1; r < data.length; r++) {
      const row = data[r];
      const pid = String(row[c("product_id")]).trim();
      if (!pid) continue;
      const pct = Number(row[c("discount_percent")]) || 0;
      const start = toDate_(row[c("starts_at")]);
      const end = toDate_(row[c("ends_at")]);
      if (!end || pct <= 0 || end <= now) continue;          // العرض منتهي أو غير صالح

      const name = names[pid] || "منتج مميز";
      const endTxt = Utilities.formatDate(end, CFG.TZ, "d/M HH:mm");
      const sNote = String(row[c("notified_start")]).trim();
      const rNote = String(row[c("notified_6h")]).trim();
      let sentStartNow = false;

      // 1) إشعار نزول العرض
      if (!sNote && (!start || start <= now)) {
        sendPush_("🎁 عرض جديد من الشامل", name + " بخصم " + pct + "% — العرض ينتهي " + endTxt);
        offersSheet.getRange(r + 1, c("notified_start") + 1).setValue(stamp_());
        sentStartNow = true;
      }

      // 2) إشعار قبل الانتهاء بـ 6 ساعات
      if (!rNote && (end - now) <= remindMs) {
        const shortOffer = start && (end - start) <= remindMs;   // عرض مدته أقصر من 6 ساعات
        if (sentStartNow || shortOffer) {
          offersSheet.getRange(r + 1, c("notified_6h") + 1).setValue("تخطي");
        } else {
          sendPush_("⏰ باقي " + CFG.REMIND_HOURS + " ساعات على انتهاء العرض", name + " بخصم " + pct + "% — احجزه قبل ما يخلص");
          offersSheet.getRange(r + 1, c("notified_6h") + 1).setValue(stamp_());
        }
      }
    }
  } finally {
    lock.releaseLock();
  }
}

/** شغّلها مرة واحدة لإنشاء التشغيل التلقائي كل 15 دقيقة */
function setupTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "checkOffers") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("checkOffers").timeBased().everyMinutes(15).create();
}

/** شغّلها لتجربة إرسال إشعار للمشتركين */
function testPush() {
  const r = sendPush_("🔔 تجربة من الشامل", "لو وصلتك الرسالة دي يبقى الإشعارات شغالة تمام");
  Logger.log(JSON.stringify(r));
}

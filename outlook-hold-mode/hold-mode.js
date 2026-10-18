/*
 * Hold Mode for Outlook: shared core + event handlers.
 *
 * This one file is loaded in three places:
 *   - classic Outlook on Windows, by its JavaScript-only event runtime (no DOM, no imports)
 *   - commands.html, the event runtime for new Outlook / Outlook on the web
 *   - taskpane.html, for the parser, formatting and saved-state helpers
 * so it is written in plain ES5 with no dependencies.
 */
var HoldCore = (function () {
  "use strict";

  var KEY_UNTIL = "holdUntil";      // ms since epoch, or absent when hold mode is off
  var KEY_MORNING = "morningTime";  // "HH:MM", the time used for "tomorrow", "monday", etc.
  var DEFAULT_MORNING = "08:00";

  var DAY_SHORT = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
  var MONTH_SHORT = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

  var WEEKDAYS = {
    sun: 0, sunday: 0,
    mon: 1, monday: 1,
    tue: 2, tues: 2, tuesday: 2,
    wed: 3, weds: 3, wednesday: 3,
    thu: 4, thur: 4, thurs: 4, thursday: 4,
    fri: 5, friday: 5,
    sat: 6, saturday: 6
  };
  var MONTHS = {
    jan: 0, january: 0, feb: 1, february: 1, mar: 2, march: 2, apr: 3, april: 3,
    may: 4, jun: 5, june: 5, jul: 6, july: 6, aug: 7, august: 7,
    sep: 8, sept: 8, september: 8, oct: 9, october: 9, nov: 10, november: 10, dec: 11, december: 11
  };
  var UNIT_MINUTES = {
    m: 1, min: 1, mins: 1, minute: 1, minutes: 1,
    h: 60, hr: 60, hrs: 60, hour: 60, hours: 60,
    d: 1440, day: 1440, days: 1440,
    w: 10080, wk: 10080, wks: 10080, week: 10080, weeks: 10080
  };

  // ---------- saved state (Outlook roaming settings, stored in the mailbox) ----------

  var settingsOverride = null;  // stand-in store when the pane is opened outside Outlook

  function settings() {
    return settingsOverride || Office.context.roamingSettings;
  }

  function useSettings(store) {
    settingsOverride = store;
  }

  /** The hold time as a Date if hold mode is on and still in the future, otherwise null. */
  function getActiveHold() {
    var ms = settings().get(KEY_UNTIL);
    if (typeof ms !== "number" || ms <= Date.now()) return null;
    return new Date(ms);
  }

  /** Raw saved value (may be in the past), so the task pane can tidy up an expired hold. */
  function getSavedHold() {
    var ms = settings().get(KEY_UNTIL);
    return typeof ms === "number" ? new Date(ms) : null;
  }

  function setHold(date, callback) {
    if (date) settings().set(KEY_UNTIL, date.getTime());
    else settings().remove(KEY_UNTIL);
    settings().saveAsync(function (r) {
      if (callback) callback(r.status === Office.AsyncResultStatus.Succeeded ? null : r.error);
    });
  }

  function getMorning() {
    var v = settings().get(KEY_MORNING);
    return parseClock(v) || parseClock(DEFAULT_MORNING);
  }

  function setMorning(hhmm, callback) {
    settings().set(KEY_MORNING, hhmm);
    settings().saveAsync(function (r) {
      if (callback) callback(r.status === Office.AsyncResultStatus.Succeeded ? null : r.error);
    });
  }

  function parseClock(v) {
    var m = /^(\d{1,2}):(\d{2})$/.exec(v || "");
    if (!m) return null;
    var h = +m[1], mi = +m[2];
    if (h > 23 || mi > 59) return null;
    return { h: h, m: mi };
  }

  // ---------- date helpers ----------

  function startOfDay(d) {
    return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }

  function addDays(d, n) {
    return new Date(d.getFullYear(), d.getMonth(), d.getDate() + n, d.getHours(), d.getMinutes());
  }

  function at(day, clock) {
    return new Date(day.getFullYear(), day.getMonth(), day.getDate(), clock.h, clock.m);
  }

  /** Next date (after today) falling on the given weekday. */
  function nextWeekday(now, wd) {
    var diff = (wd - now.getDay() + 7) % 7;
    return addDays(startOfDay(now), diff === 0 ? 7 : diff);
  }

  /** The next morning: today if the morning time hasn't passed yet, otherwise tomorrow. */
  function nextMorning(now, morning) {
    var today = at(now, morning);
    return today > now ? today : at(addDays(startOfDay(now), 1), morning);
  }

  /** The next weekday (Mon-Fri) morning strictly after today. */
  function nextBusinessMorning(now, morning) {
    var d = addDays(startOfDay(now), 1);
    while (d.getDay() === 0 || d.getDay() === 6) d = addDays(d, 1);
    return at(d, morning);
  }

  // ---------- quick options ----------

  /** Preset choices shown in the picker, computed for "now". Duplicate times are dropped. */
  function presets(now, morning) {
    var tomorrow = at(addDays(startOfDay(now), 1), morning);
    var list = [
      { label: "In 1 hour", date: new Date(now.getTime() + 60 * 60000) },
      { label: "In 3 hours", date: new Date(now.getTime() + 3 * 60 * 60000) }
    ];
    var thisMorning = at(now, morning);
    if (thisMorning > now) list.push({ label: "This morning", date: thisMorning });  // e.g. writing at 1 AM
    list.push({ label: "Tomorrow morning", date: tomorrow });
    var biz = nextBusinessMorning(now, morning);
    if (biz.getTime() !== tomorrow.getTime()) list.push({ label: "Next weekday morning", date: biz });
    var monday = at(nextWeekday(now, 1), morning);
    if (monday.getTime() !== biz.getTime() && monday.getTime() !== tomorrow.getTime()) {
      list.push({ label: "Monday morning", date: monday });
    }
    return list;
  }

  // ---------- natural-language parser ----------

  /**
   * Parses things like "in 2h", "in 1h 30m", "tomorrow 9am", "fri 2pm", "monday morning",
   * "oct 24 8:30", "24 oct", "2026-10-24 08:00", "noon", "next week".
   * Returns { date } on success, { error } when the text is understood but unusable,
   * or null when it can't be read.
   */
  function parseWhen(input, now, morning) {
    var s = String(input || "").toLowerCase()
      .replace(/[,.](?!\d)/g, " ")
      .replace(/(\d)(st|nd|rd|th)\b/g, "$1")
      .replace(/@/g, " ")
      .replace(/\b(at|on|this)\b/g, " ")
      .replace(/(\d)\s+(am|pm|a|p)\b/g, "$1$2")
      .replace(/\s+/g, " ")
      .trim();
    if (!s) return null;

    var rel = parseRelative(s, now);
    if (rel) return rel;

    var tokens = s.split(" ");
    var day = null;    // Date at midnight
    var clock = null;  // { h, m }
    var i = 0, t, n, iso;

    function setDay(d) { if (day) return false; day = d; return true; }
    function setClock(c) { if (clock) return false; clock = c; return true; }

    while (i < tokens.length) {
      t = tokens[i];
      n = tokens[i + 1];

      if (t === "today") { if (!setDay(startOfDay(now))) return null; i++; continue; }
      if (t === "tonight") {
        if (!setDay(startOfDay(now))) return null;
        if (!clock) clock = { h: 20, m: 0, soft: true };
        i++; continue;
      }
      if (t === "tomorrow" || t === "tmrw" || t === "tmr" || t === "tom") {
        if (!setDay(addDays(startOfDay(now), 1))) return null; i++; continue;
      }
      if (t === "next" && n === "week") {
        if (!setDay(nextWeekday(now, 1))) return null; i += 2; continue;
      }
      if (t === "next" && WEEKDAYS.hasOwnProperty(n)) {
        if (!setDay(nextWeekday(now, WEEKDAYS[n]))) return null; i += 2; continue;
      }
      if (WEEKDAYS.hasOwnProperty(t)) {
        if (!setDay(nextWeekday(now, WEEKDAYS[t]))) return null; i++; continue;
      }
      // "oct 24", "oct 24 2027"
      if (MONTHS.hasOwnProperty(t) && /^\d{1,2}$/.test(n || "")) {
        var y1 = /^\d{4}$/.test(tokens[i + 2] || "") ? +tokens[i + 2] : null;
        var d1 = monthDay(now, MONTHS[t], +n, y1);
        if (!d1 || !setDay(d1)) return null;
        i += y1 ? 3 : 2; continue;
      }
      // "24 oct", "24 oct 2027"
      if (/^\d{1,2}$/.test(t) && MONTHS.hasOwnProperty(n)) {
        var y2 = /^\d{4}$/.test(tokens[i + 2] || "") ? +tokens[i + 2] : null;
        var d2 = monthDay(now, MONTHS[n], +t, y2);
        if (!d2 || !setDay(d2)) return null;
        i += y2 ? 3 : 2; continue;
      }
      // "2026-10-24"
      if ((iso = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(t))) {
        var d3 = new Date(+iso[1], +iso[2] - 1, +iso[3]);
        if (d3.getMonth() !== +iso[2] - 1 || !setDay(d3)) return null;
        i++; continue;
      }
      // named times of day
      if (t === "morning") { if (!setClock(morning)) return null; i++; continue; }
      if (t === "noon" || t === "midday") { if (!setClock({ h: 12, m: 0 })) return null; i++; continue; }
      if (t === "afternoon") { if (!setClock({ h: 13, m: 0 })) return null; i++; continue; }
      if (t === "eod") { if (!setClock({ h: 17, m: 0 })) return null; i++; continue; }
      if (t === "evening") { if (!setClock({ h: 18, m: 0 })) return null; i++; continue; }
      if (t === "night") { if (!setClock({ h: 20, m: 0 })) return null; i++; continue; }
      if (t === "midnight") { if (!setClock({ h: 0, m: 0 })) return null; i++; continue; }

      var c = parseTimeToken(t);
      if (c) {
        if (clock && clock.soft) {
          // "tonight 9" overrides the 8 PM default and means 9 PM
          if (!c.explicit && c.h < 12) c.h += 12;
          clock = null;
        }
        if (!setClock(c)) return null;
        i++; continue;
      }
      return null;  // unknown word
    }

    var date;
    if (day && clock) {
      date = at(day, clock);
    } else if (day) {
      date = at(day, morning);
    } else if (clock && tokens.length === 1 && tokens[0] === "morning") {
      date = nextMorning(now, morning);
    } else if (clock) {
      date = at(now, clock);
      if (date <= now) date = at(addDays(startOfDay(now), 1), clock);
    } else {
      return null;
    }
    if (date <= now) return { error: "That time has already passed." };
    return { date: date };
  }

  function parseRelative(s, now) {
    if (s.indexOf("in ") !== 0) return null;
    var rest = s.slice(3)
      .replace(/\bhalf an? hour\b/g, "30m")
      .replace(/\ban? /g, "1 ")
      .replace(/\s+/g, "");
    var re = /(\d+(?:\.\d+)?)([a-z]+)/g;
    var total = 0, consumed = 0, m;
    while ((m = re.exec(rest))) {
      if (m.index !== consumed || !UNIT_MINUTES.hasOwnProperty(m[2])) return null;
      total += parseFloat(m[1]) * UNIT_MINUTES[m[2]];
      consumed = re.lastIndex;
    }
    if (consumed !== rest.length || total <= 0) return null;
    return { date: new Date(now.getTime() + Math.round(total) * 60000) };
  }

  /** "8", "8am", "8:30", "8:30pm", "20:00", "8p". Bare 1-6 means PM (office hours). */
  function parseTimeToken(t) {
    var m = /^(\d{1,2})(?::(\d{2}))?(am|pm|a|p)?$/.exec(t);
    if (!m) return null;
    var h = +m[1], mi = m[2] ? +m[2] : 0, ap = m[3];
    if (mi > 59) return null;
    if (ap) {
      if (h < 1 || h > 12) return null;
      if (ap.charAt(0) === "p" && h !== 12) h += 12;
      if (ap.charAt(0) === "a" && h === 12) h = 0;
    } else {
      if (h > 23) return null;
      if (h >= 1 && h <= 6) h += 12;
    }
    return { h: h, m: mi, explicit: !!ap };
  }

  /** Month/day with no year means the next time that date comes around. */
  function monthDay(now, month, dayOfMonth, year) {
    var y = year || now.getFullYear();
    var d = new Date(y, month, dayOfMonth);
    if (d.getMonth() !== month) return null;
    if (!year && d < startOfDay(now)) d = new Date(y + 1, month, dayOfMonth);
    return d;
  }

  // ---------- formatting (manual: the classic event runtime may not have Intl) ----------

  function clockText(d) {
    var h = d.getHours(), m = d.getMinutes();
    var ap = h < 12 ? "AM" : "PM";
    h = h % 12 || 12;
    return h + ":" + (m < 10 ? "0" : "") + m + " " + ap;
  }

  /** "Mon, Oct 20, 8:00 AM" (adds the year when it isn't this year). */
  function formatFull(d, now) {
    now = now || new Date();
    var s = DAY_SHORT[d.getDay()] + ", " + MONTH_SHORT[d.getMonth()] + " " + d.getDate();
    if (d.getFullYear() !== now.getFullYear()) s += ", " + d.getFullYear();
    return s + ", " + clockText(d);
  }

  /** "Today, 3:42 PM" / "Tomorrow, 8:00 AM" / "Mon, 8:00 AM" / "Mon, Oct 27, 8:00 AM". */
  function formatSmart(d, now) {
    now = now || new Date();
    var days = Math.round((startOfDay(d) - startOfDay(now)) / 86400000);
    if (days === 0) return "Today, " + clockText(d);
    if (days === 1) return "Tomorrow, " + clockText(d);
    if (days > 1 && days < 7) return DAY_SHORT[d.getDay()] + ", " + clockText(d);
    return formatFull(d, now);
  }

  /** "in 2 days, 14 hours" / "in 3 hours, 5 min" / "in 12 min" / "in less than a minute". */
  function formatRelative(d, now) {
    now = now || new Date();
    var mins = Math.round((d - now) / 60000);
    if (mins < 1) return "in less than a minute";
    var days = Math.floor(mins / 1440), hours = Math.floor((mins % 1440) / 60), m = mins % 60;
    function plural(n, w) { return n + " " + w + (n === 1 ? "" : "s"); }
    if (days) return "in " + plural(days, "day") + (hours ? ", " + plural(hours, "hour") : "");
    if (hours) return "in " + plural(hours, "hour") + (m ? ", " + m + " min" : "");
    return "in " + m + " min";
  }

  function bannerText(d) {
    return "Hold mode is on. This email will be delivered " + formatFull(d) + ".";
  }

  return {
    getActiveHold: getActiveHold,
    getSavedHold: getSavedHold,
    setHold: setHold,
    useSettings: useSettings,
    getMorning: getMorning,
    setMorning: setMorning,
    parseClock: parseClock,
    presets: presets,
    parseWhen: parseWhen,
    formatFull: formatFull,
    formatSmart: formatSmart,
    formatRelative: formatRelative,
    clockText: clockText,
    bannerText: bannerText
  };
})();

// ---------- event handlers ----------

var BANNER_KEY = "holdModeBanner";

function showHoldBanner(item, holdUntil, done) {
  if (!holdUntil) {
    item.notificationMessages.removeAsync(BANNER_KEY, function () { done(); });
    return;
  }
  item.notificationMessages.replaceAsync(BANNER_KEY, {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: HoldCore.bannerText(holdUntil),
    icon: "Icon.16x16",
    persistent: true
  }, function () { done(); });
}

/** New email, reply or forward: show the banner when hold mode is on. */
function onNewMessageComposeHandler(event) {
  var holdUntil = HoldCore.getActiveHold();
  if (!holdUntil) { event.completed(); return; }
  showHoldBanner(Office.context.mailbox.item, holdUntil, function () { event.completed(); });
}

/** Send: when hold mode is on, schedule delivery for the hold time (held on the server). */
function onMessageSendHandler(event) {
  var holdUntil = HoldCore.getActiveHold();
  if (!holdUntil) { event.completed({ allowEvent: true }); return; }

  var item = Office.context.mailbox.item;
  item.delayDeliveryTime.getAsync(function (got) {
    var existing = got.status === Office.AsyncResultStatus.Succeeded && got.value ? new Date(got.value) : null;
    if (existing && existing >= holdUntil) {
      // Already scheduled for the hold time or later (e.g. by hand); leave it alone.
      event.completed({ allowEvent: true });
      return;
    }
    item.delayDeliveryTime.setAsync(holdUntil, function (set) {
      if (set.status === Office.AsyncResultStatus.Failed) {
        event.completed({
          allowEvent: false,
          errorMessage: "Hold mode couldn't schedule this email for " + HoldCore.formatFull(holdUntil) +
            " (" + set.error.message + "). Choose Don't Send and try again, or Send Anyway to send it now."
        });
        return;
      }
      event.completed({ allowEvent: true });
    });
  });
}

if (typeof Office !== "undefined" && Office.actions && Office.actions.associate) {
  Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
}

if (typeof module !== "undefined" && module.exports) module.exports = HoldCore;

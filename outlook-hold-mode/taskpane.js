/* Hold Mode task pane: status card, Beeper-style quick picks, typed times, calendar picker. */
(function () {
  "use strict";

  var $ = function (id) { return document.getElementById(id); };
  var parsedDate = null;  // result of the typed time, if valid

  Office.onReady(function (info) {
    if (!info.host) installBrowserShim();  // opened outside Outlook, e.g. for a quick look in a browser
    init();
  });

  function init() {
    $("morning").value = clockValue(HoldCore.getMorning());

    $("turnOff").addEventListener("click", function () { save(null); });
    $("whenText").addEventListener("input", updatePreview);
    $("customForm").addEventListener("submit", function (e) {
      e.preventDefault();
      if (parsedDate) save(parsedDate);
    });
    $("calendarForm").addEventListener("submit", function (e) {
      e.preventDefault();
      var v = $("whenPicker").value;
      if (!v) return showMessage("Pick a date and time first.");
      var d = fromLocalInput(v);
      if (d <= new Date()) return showMessage("That time has already passed.");
      save(d);
    });
    $("morning").addEventListener("change", function () {
      var c = HoldCore.parseClock($("morning").value);
      if (!c) return;
      HoldCore.setMorning(clockValue(c), function (err) {
        if (err) showMessage("Couldn't save the morning time: " + err.message);
        render();
      });
    });

    // Tidy up a hold that has already expired, then keep the countdown and presets fresh.
    var saved = HoldCore.getSavedHold();
    if (saved && !HoldCore.getActiveHold()) HoldCore.setHold(null);
    render();
    setInterval(function () {
      if (HoldCore.getSavedHold() && !HoldCore.getActiveHold()) {
        HoldCore.setHold(null, function () { render(); updateBanner(); });
      } else {
        render();
      }
    }, 30000);
  }

  function save(date) {
    showMessage("");
    HoldCore.setHold(date, function (err) {
      if (err) return showMessage("Couldn't save: " + err.message);
      $("whenText").value = "";
      updatePreview();
      render();
      updateBanner();
    });
  }

  function render() {
    var now = new Date();
    var hold = HoldCore.getActiveHold();
    var status = $("status");

    status.classList.toggle("on", !!hold);
    status.classList.toggle("off", !hold);
    $("turnOff").hidden = !hold;
    if (hold) {
      $("statusTitle").textContent = "Hold mode is on";
      $("statusDetail").innerHTML = "";
      $("statusDetail").append(
        "Mail you send is delivered ",
        bold(HoldCore.formatSmart(hold, now)),
        " (" + HoldCore.formatRelative(hold, now) + ")."
      );
      $("pickerTitle").textContent = "Change hold time";
    } else {
      $("statusTitle").textContent = "Hold mode is off";
      $("statusDetail").textContent = "Mail sends normally.";
      $("pickerTitle").textContent = "Hold outgoing mail until…";
    }

    var list = $("presets");
    list.innerHTML = "";
    HoldCore.presets(now, HoldCore.getMorning()).forEach(function (p) {
      var btn = document.createElement("button");
      btn.type = "button";
      btn.className = "preset";
      var label = document.createElement("span");
      label.textContent = p.label;
      var time = document.createElement("span");
      time.className = "preset-time";
      time.textContent = HoldCore.formatSmart(p.date, now);
      btn.append(label, time);
      btn.addEventListener("click", function () { save(p.date); });
      var li = document.createElement("li");
      li.append(btn);
      list.append(li);
    });

    if (!$("whenPicker").value) {
      var tomorrow = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1);
      var morning = HoldCore.getMorning();
      tomorrow.setHours(morning.h, morning.m);
      $("whenPicker").value = toLocalInput(tomorrow);
    }
    updatePreview();
  }

  function updatePreview() {
    var text = $("whenText").value;
    var preview = $("preview");
    var apply = $("applyText");
    parsedDate = null;
    preview.className = "preview";

    if (!text.trim()) {
      preview.innerHTML = "&nbsp;";
      apply.textContent = "Hold until…";
      apply.disabled = true;
      return;
    }
    var r = HoldCore.parseWhen(text, new Date(), HoldCore.getMorning());
    if (r && r.date) {
      parsedDate = r.date;
      preview.className = "preview ok";
      preview.textContent = "→ " + HoldCore.formatFull(r.date) + " (" + HoldCore.formatRelative(r.date) + ")";
      apply.textContent = "Hold until " + HoldCore.formatSmart(r.date);
      apply.disabled = false;
    } else {
      preview.className = "preview bad";
      preview.textContent = r && r.error ? r.error : "Not sure what time that is. Try “tomorrow 9am” or “oct 24 8:30”.";
      apply.textContent = "Hold until…";
      apply.disabled = true;
    }
  }

  /** If the pane is open on an email being written, refresh that email's banner right away. */
  function updateBanner() {
    var item = Office.context && Office.context.mailbox && Office.context.mailbox.item;
    if (!item || !item.delayDeliveryTime || typeof showHoldBanner !== "function") return;
    showHoldBanner(item, HoldCore.getActiveHold(), function () {});
  }

  function showMessage(text) { $("message").textContent = text; }

  function bold(text) {
    var b = document.createElement("b");
    b.textContent = text;
    return b;
  }

  function clockValue(c) {
    return (c.h < 10 ? "0" : "") + c.h + ":" + (c.m < 10 ? "0" : "") + c.m;
  }

  function toLocalInput(d) {
    var p = function (n) { return (n < 10 ? "0" : "") + n; };
    return d.getFullYear() + "-" + p(d.getMonth() + 1) + "-" + p(d.getDate()) + "T" + p(d.getHours()) + ":" + p(d.getMinutes());
  }

  function fromLocalInput(v) {
    var m = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})/.exec(v);
    return new Date(+m[1], +m[2] - 1, +m[3], +m[4], +m[5]);
  }

  /** Lets the pane run in a normal browser tab (settings kept in localStorage) for previewing. */
  function installBrowserShim() {
    var store = {};
    try { store = JSON.parse(localStorage.getItem("holdModeShim") || "{}"); } catch (e) { /* ignore */ }
    var persist = function () { try { localStorage.setItem("holdModeShim", JSON.stringify(store)); } catch (e) { /* ignore */ } };
    HoldCore.useSettings({
      get: function (k) { return store[k]; },
      set: function (k, v) { store[k] = v; },
      remove: function (k) { delete store[k]; },
      saveAsync: function (cb) { persist(); if (cb) cb({ status: Office.AsyncResultStatus.Succeeded }); }
    });
  }
})();

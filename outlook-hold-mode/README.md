# Hold Mode for Outlook

Turn on "hold mode" and every email you send is scheduled for a time you choose, until you turn it off
or that time passes. Delivery is scheduled on the Exchange server (`item.delayDeliveryTime`), so
Outlook doesn't need to be running when the mail goes out.

Works in classic Outlook for Windows, new Outlook for Windows, and Outlook on the web (Mailbox requirement set 1.13).

## Install

1. Download `manifest.xml`: open
   https://github.com/ktgre514-hash/office-extensions/blob/main/outlook-hold-mode/manifest.xml
   and click **Download raw file**.
2. In classic Outlook: **Home → Get Add-ins → My add-ins → Custom Addins → Add a custom add-in → Add from File**,
   and pick the downloaded file. (Some accounts also offer **Add from URL**; paste
   `https://ktgre514-hash.github.io/office-extensions/outlook-hold-mode/manifest.xml` there.)
3. Restart classic Outlook if the button doesn't appear.

The manifest only points at the files on GitHub Pages, so code changes reach Outlook without reinstalling.
Reinstall only when `manifest.xml` itself changes (bump its `<Version>` when it does).

## Use

- **Hold Mode** button (Home ribbon when reading mail, Message ribbon when writing; may be under **Apps** or **…**)
  opens the panel. Pin it (pushpin icon) to keep the status visible while you move between emails.
- Pick a quick option, type a time (`tomorrow 9am`, `fri 2pm`, `in 90m`, `oct 24 8:30`, `monday morning`),
  or use the calendar.
- While hold mode is on, every new email, reply and forward shows a banner with the delivery time.
- Hold mode turns itself off when the time passes.

**Classic Outlook:** held emails skip the Outbox and can't be cancelled after Send; they appear in Sent Items.
In new Outlook / on the web they wait in Drafts and can still be edited.

## Files

| File | Purpose |
| --- | --- |
| `manifest.xml` | Add-in manifest (ribbon buttons, OnNewMessageCompose / OnMessageSend events) |
| `hold-mode.js` | Saved state, time parser, formatting, and the event handlers. Plain ES5: classic Outlook runs it in a JavaScript-only runtime. |
| `commands.html` | Event runtime page for new Outlook / web |
| `taskpane.*` | Status + picker panel |
| `tools/make-icons.js` | Regenerates `assets/icon-*.png` |
| `test/parse.test.js` | `node --test outlook-hold-mode/test/parse.test.js` |

To preview the panel in a browser, serve the repo root (e.g. `npx http-server -c-1`) and open
`/outlook-hold-mode/taskpane.html`; outside Outlook it stores settings in localStorage.

// Run with: node --test outlook-hold-mode/test
const test = require("node:test");
const assert = require("node:assert");
const Hold = require("../hold-mode.js");

// Friday, Oct 16 2026, 3:00 PM local time
const NOW = new Date(2026, 9, 16, 15, 0);
const MORNING = { h: 8, m: 0 };
const d = (mo, day, h, mi = 0, y = 2026) => new Date(y, mo - 1, day, h, mi).getTime();

function when(text) {
  const r = Hold.parseWhen(text, NOW, MORNING);
  return r && r.date ? r.date.getTime() : r;
}

const cases = [
  ["in 2h", NOW.getTime() + 120 * 60000],
  ["in 90 minutes", NOW.getTime() + 90 * 60000],
  ["in 1h 30m", NOW.getTime() + 90 * 60000],
  ["in an hour", NOW.getTime() + 60 * 60000],
  ["in half an hour", NOW.getTime() + 30 * 60000],
  ["in 2 days", NOW.getTime() + 2 * 1440 * 60000],
  ["tomorrow", d(10, 17, 8)],
  ["tomorrow 9am", d(10, 17, 9)],
  ["tmrw at 9:30 pm", d(10, 17, 21, 30)],
  ["tomorrow morning", d(10, 17, 8)],
  ["mon", d(10, 19, 8)],
  ["Monday 9", d(10, 19, 9)],
  ["next monday 10am", d(10, 19, 10)],
  ["fri 2", d(10, 23, 14)],          // today is Friday, so "fri" is next week; bare 2 = 2 PM
  ["friday 2:30", d(10, 23, 14, 30)],
  ["oct 24 8:30", d(10, 24, 8, 30)],
  ["24 oct", d(10, 24, 8)],
  ["October 24th at 1pm", d(10, 24, 13)],
  ["jan 5", d(1, 5, 8, 0, 2027)],    // already past this year -> next year
  ["2026-11-02 07:45", d(11, 2, 7, 45)],
  ["5pm", d(10, 16, 17)],            // later today
  ["2pm", d(10, 17, 14)],            // already passed today -> tomorrow
  ["noon", d(10, 17, 12)],
  ["tonight", d(10, 16, 20)],
  ["tonight 9", d(10, 16, 21)],
  ["morning", d(10, 17, 8)],
  ["next week", d(10, 19, 8)],
  ["monday morning", d(10, 19, 8)],
  ["monday evening", d(10, 19, 18)],
  ["sat 8 am", d(10, 17, 8)]
];

for (const [text, expected] of cases) {
  test(`parses "${text}"`, () => assert.strictEqual(when(text), expected));
}

test("rejects gibberish and conflicts", () => {
  for (const t of ["", "banana", "in 2 bananas", "tomorrow monday", "9am 10am", "feb 30", "13pm", "8:75"]) {
    assert.strictEqual(when(t), null, t);
  }
});

test("reports past times", () => {
  assert.deepStrictEqual(Hold.parseWhen("today 9am", NOW, MORNING), { error: "That time has already passed." });
  assert.deepStrictEqual(Hold.parseWhen("2026-10-01", NOW, MORNING), { error: "That time has already passed." });
});

test("presets on a Friday afternoon", () => {
  const p = Hold.presets(NOW, MORNING).map((x) => [x.label, x.date.getTime()]);
  assert.deepStrictEqual(p, [
    ["In 1 hour", d(10, 16, 16)],
    ["In 3 hours", d(10, 16, 18)],
    ["Tomorrow morning", d(10, 17, 8)],
    ["Next weekday morning", d(10, 19, 8)]
  ]);
});

test("presets on a Tuesday include Monday", () => {
  const tue = new Date(2026, 9, 13, 10, 0);
  const labels = Hold.presets(tue, MORNING).map((x) => x.label);
  assert.deepStrictEqual(labels, ["In 1 hour", "In 3 hours", "Tomorrow morning", "Monday morning"]);
});

test("presets after midnight include this morning", () => {
  const satEarly = new Date(2026, 9, 17, 1, 0);
  const p = Hold.presets(satEarly, MORNING).map((x) => [x.label, x.date.getTime()]);
  assert.deepStrictEqual(p.slice(2), [
    ["This morning", d(10, 17, 8)],
    ["Tomorrow morning", d(10, 18, 8)],
    ["Next weekday morning", d(10, 19, 8)]
  ]);
});

test("formatting", () => {
  assert.strictEqual(Hold.formatFull(new Date(d(10, 19, 8)), NOW), "Mon, Oct 19, 8:00 AM");
  assert.strictEqual(Hold.formatFull(new Date(d(1, 5, 13, 5, 2027)), NOW), "Tue, Jan 5, 2027, 1:05 PM");
  assert.strictEqual(Hold.formatSmart(new Date(d(10, 16, 17)), NOW), "Today, 5:00 PM");
  assert.strictEqual(Hold.formatSmart(new Date(d(10, 17, 8)), NOW), "Tomorrow, 8:00 AM");
  assert.strictEqual(Hold.formatSmart(new Date(d(10, 19, 8)), NOW), "Mon, 8:00 AM");
  assert.strictEqual(Hold.formatSmart(new Date(d(10, 26, 8)), NOW), "Mon, Oct 26, 8:00 AM");
  assert.strictEqual(Hold.formatRelative(new Date(d(10, 19, 8)), NOW), "in 2 days, 17 hours");
  assert.strictEqual(Hold.formatRelative(new Date(NOW.getTime() + 65 * 60000), NOW), "in 1 hour, 5 min");
  assert.strictEqual(Hold.formatRelative(new Date(NOW.getTime() + 12 * 60000), NOW), "in 12 min");
});

/**
 * 空き時間・会議室照合ツール（公開用サンプル）
 * Copyright (c) 2026. MIT License.
 * 管理者が初回に setupSampleSpreadsheet() を実行してください。
 * 本番利用前に README の共有権限・権限審査の注意事項を確認してください。
 */
const APP = Object.freeze({
  CONFIG_KEY: 'MEMBER_SPREADSHEET_ID',
  TZ: 'Asia/Tokyo',
  START_HOUR: 10,
  END_HOUR: 19,
  MIN_MINUTES: 60,
  SLOT_STEP_MINUTES: 30,
  MAX_DAYS: 31,
  MAX_CALENDARS: 40,
  HOLIDAY_CALENDAR_IDS: [
    'ja.japanese.official#holiday@group.v.calendar.google.com',
    'ja.japanese#holiday@group.v.calendar.google.com'
  ]
});

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('空き時間・会議室検索')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** 実行者本人のGoogle Driveに設定シートを作成。最初の一回のみ手動実行。 */
function setupSampleSpreadsheet() {
  const ss = SpreadsheetApp.create('空き時間検索_公開版_メンバー設定');
  const sheet = ss.getSheets()[0];
  sheet.setName('メンバー');
  sheet.getRange(1, 1, 5, 3).setValues([
    ['区分', '表示名', 'カレンダーID（通常はメールアドレス）'],
    ['営業', '山田（サンプル・要変更）', 'yamada@example.com'],
    ['制作', '佐藤（サンプル・要変更）', 'sato@example.com'],
    ['その他', '追加メンバー（サンプル・要変更）', 'member@example.com'],
    ['会議室', '会議室A（サンプル・要変更）', 'room@example.com']
  ]);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 3);
  PropertiesService.getScriptProperties().setProperty(APP.CONFIG_KEY, ss.getId());
  Logger.log('設定シート: ' + ss.getUrl());
  Logger.log('保存した設定ID: ' + ss.getId());
  return ss.getUrl();
}

/** 既存の設定シートを利用する場合、IDを渡してスクリプトプロパティに保存する。 */
function setConfigSpreadsheetId(id) {
  if (!id || !/^[a-zA-Z0-9_-]{20,}$/.test(String(id))) {
    throw new Error('正しいスプレッドシートIDを入力してください。');
  }
  SpreadsheetApp.openById(String(id)).getSheetByName('メンバー') || (() => { throw new Error('「メンバー」シートがありません。'); })();
  PropertiesService.getScriptProperties().setProperty(APP.CONFIG_KEY, String(id));
}

function getConfig_() {
  const id = PropertiesService.getScriptProperties().getProperty(APP.CONFIG_KEY);
  if (!id) throw new Error('初期設定が未完了です。管理者が setupSampleSpreadsheet() を実行してください。');
  const sh = SpreadsheetApp.openById(id).getSheetByName('メンバー');
  if (!sh) throw new Error('「メンバー」シートがありません。');
  const rows = sh.getDataRange().getDisplayValues().slice(1);
  const groups = {};
  const rooms = [];
  const seen = new Set();
  for (const row of rows) {
    const group = String(row[0] || '').trim();
    const name = String(row[1] || '').trim();
    const id = String(row[2] || '').trim();
    if (!group || !name || !id || !id.includes('@')) continue;
    const normalized = id.toLowerCase();
    if (seen.has(normalized)) continue;
    seen.add(normalized);
    const entry = {name, id};
    if (group === '会議室') rooms.push(entry);
    else (groups[group] = groups[group] || []).push(entry);
  }
  return {groups, rooms};
}

function getBootstrapData() {
  return getConfig_();
}

function parseDate_(s) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) throw new Error('日付の形式が不正です。');
  const [y, m, d] = s.split('-').map(Number);
  const date = new Date(y, m - 1, d);
  if (date.getFullYear() !== y || date.getMonth() !== m - 1 || date.getDate() !== d) {
    throw new Error('存在しない日付です。');
  }
  return date;
}
function localDay_(date) {
  return Utilities.formatDate(date, APP.TZ, 'yyyy-MM-dd');
}
function formatTime_(ms) {
  return Utilities.formatDate(new Date(ms), APP.TZ, 'HH:mm');
}
function roundUp_(ms) {
  const step = APP.SLOT_STEP_MINUTES * 60000;
  return Math.ceil(ms / step) * step;
}
function roundDown_(ms) {
  const step = APP.SLOT_STEP_MINUTES * 60000;
  return Math.floor(ms / step) * step;
}

function getDateRange_(mode, start, end) {
  const today = parseDate_(localDay_(new Date()));
  let first, last;
  if (mode === 'this_week') {
    first = new Date(today);
    first.setDate(first.getDate() + 1); // 当日は対象外
    last = new Date(today);
    const deltaFriday = (5 - today.getDay() + 7) % 7;
    last.setDate(last.getDate() + deltaFriday);
  } else if (mode === 'next_week') {
    const mondayOffset = (8 - today.getDay()) % 7 || 7;
    first = new Date(today);
    first.setDate(first.getDate() + mondayOffset);
    last = new Date(first);
    last.setDate(last.getDate() + 4);
  } else if (mode === 'custom') {
    first = parseDate_(start);
    last = parseDate_(end);
    if (first.getTime() < today.getTime()) throw new Error('指定期間の開始日は今日以降にしてください。');
  } else {
    throw new Error('期間指定が正しくありません。');
  }
  if (first > last) return {first, last, empty: true};
  const days = Math.round((last.getTime() - first.getTime()) / 86400000) + 1;
  if (days > APP.MAX_DAYS || days <= 0) throw new Error('検索可能期間は最大31日です。');
  return {first, last, empty: false};
}

/** APIが予定情報にアクセスできない場合は「空き」ではなくエラーにする。 */
function getBusy_(ids, min, max) {
  if (ids.length === 0) return {};
  if (ids.length > APP.MAX_CALENDARS) throw new Error('照会できるカレンダーは最大40件です。');
  const response = Calendar.Freebusy.query({
    timeMin: min.toISOString(),
    timeMax: max.toISOString(),
    timeZone: APP.TZ,
    items: ids.map(id => ({id}))
  });
  const calendars = response.calendars || {};
  const out = {};
  for (const id of ids) {
    const cal = calendars[id];
    if (!cal || (cal.errors && cal.errors.length)) {
      throw new Error('予定情報を取得できないカレンダーがあります。共有権限とIDを確認してください: ' + id);
    }
    out[id] = (cal.busy || []).map(b => ({start: Date.parse(b.start), end: Date.parse(b.end)}));
  }
  return out;
}

/**
 * 利用者が購読する日本の祝日カレンダーを自動判定する。
 * 旧ID・新IDの双方に対応。見つからない場合は祝日を無視せず停止する。
 */
function findHolidayCalendarId_() {
  const ids = [];
  let pageToken;
  do {
    const options = {maxResults: 250};
    if (pageToken) options.pageToken = pageToken;
    const page = Calendar.CalendarList.list(options);
    for (const item of (page.items || [])) {
      const id = String(item.id || '');
      const name = String(item.summary || '');
      // 祝日IDを優先。Googleアカウントの言語による表示名差異を避ける。
      if (APP.HOLIDAY_CALENDAR_IDS.includes(id) ||
          (/japanese.*holiday/i.test(id) && /group\.v\.calendar\.google\.com$/.test(id)) ||
          (/^(日本の祝日|Holidays in Japan)$/i.test(name) && /holiday/i.test(id))) {
        ids.push(id);
      }
    }
    pageToken = page.nextPageToken;
  } while (pageToken);

  for (const knownId of APP.HOLIDAY_CALENDAR_IDS) {
    if (ids.includes(knownId)) return knownId;
  }
  return ids[0] || null;
}

function holidayDates_(start, end) {
  const id = findHolidayCalendarId_();
  if (!id) {
    throw new Error('日本の祝日カレンダーを確認できません。Googleカレンダーの「関心のあるカレンダー」から「日本の祝日」を追加し、再実行してください。');
  }
  const dates = new Set();
  const until = new Date(end);
  until.setDate(until.getDate() + 1);
  let pageToken;
  do {
    const options = {
      timeMin: start.toISOString(),
      timeMax: until.toISOString(),
      singleEvents: true,
      maxResults: 250,
      timeZone: APP.TZ
    };
    if (pageToken) options.pageToken = pageToken;
    const page = Calendar.Events.list(id, options);
    for (const event of (page.items || [])) {
      if (event.status === 'cancelled') continue;
      // 祝日カレンダーの終日イベントは開始日を yyyy-MM-dd で返す。
      const day = event.start && (event.start.date ||
        (event.start.dateTime ? localDay_(new Date(event.start.dateTime)) : null));
      if (day) dates.add(day);
    }
    pageToken = page.nextPageToken;
  } while (pageToken);
  return dates;
}

/** @returns {{dates:Array,rooms:Array}} 日時はISO 8601で返し、表示処理と分離 */
function findFreeSlots(payload) {
  payload = payload || {};
  const config = getConfig_();
  const selected = (payload.ids || []).map(String).map(x => x.trim()).filter(Boolean);
  const ids = Array.from(new Set(['primary', ...selected]));
  if (selected.some(id => !id.includes('@'))) throw new Error('参加者のカレンダーIDを確認してください。');
  if (ids.length > APP.MAX_CALENDARS) throw new Error('参加者のカレンダー数が上限を超えています。');
  const {first, last, empty} = getDateRange_(payload.mode, payload.start, payload.end);
  if (empty) return {dates: [], rooms: config.rooms, message: '今週の対象となる平日がありません。'};
  const apiStart = new Date(first);
  const apiEnd = new Date(last);
  apiEnd.setDate(apiEnd.getDate() + 1);
  const busy = getBusy_(ids, apiStart, apiEnd);
  const holidays = holidayDates_(first, last);
  const dates = [];
  for (let date = new Date(first); date <= last; date.setDate(date.getDate() + 1)) {
    if (date.getDay() === 0 || date.getDay() === 6 || holidays.has(localDay_(date))) continue;
    const y = date.getFullYear(), m = date.getMonth(), d = date.getDate();
    const dayStart = new Date(y, m, d, APP.START_HOUR).getTime();
    const dayEnd = new Date(y, m, d, APP.END_HOUR).getTime();
    const occupied = ids.flatMap(id => busy[id] || [])
      .filter(x => x.end > dayStart && x.start < dayEnd)
      .map(x => ({start: Math.max(x.start, dayStart), end: Math.min(x.end, dayEnd)}))
      .sort((a,b) => a.start - b.start);
    const merged = [];
    for (const x of occupied) {
      const lastBusy = merged[merged.length - 1];
      if (lastBusy && x.start <= lastBusy.end) lastBusy.end = Math.max(lastBusy.end, x.end);
      else merged.push({...x});
    }
    let current = dayStart;
    const slots = [];
    for (const b of [...merged, {start: dayEnd, end: dayEnd}]) {
      const from = roundUp_(current), to = roundDown_(b.start);
      if (to - from >= APP.MIN_MINUTES * 60000) {
        slots.push({start: new Date(from).toISOString(), end: new Date(to).toISOString()});
      }
      current = Math.max(current, b.end);
    }
    if (slots.length) dates.push({date: localDay_(date), slots});
  }
  return {dates, rooms: config.rooms, message: dates.length ? '' : '条件に合う空き時間はありませんでした。'};
}

/** 候補枠を改ざんして照合する必要がないように期間と参加者を再計算。 */
function findSlotsWithRooms(payload) {
  const result = findFreeSlots(payload);
  if (!result.dates.length) return result;
  if (!result.rooms.length) throw new Error('「メンバー」シートに会議室が登録されていません。');
  const roomIds = result.rooms.map(r => r.id);
  const allSlots = result.dates.flatMap(d => d.slots);
  const first = new Date(Math.min(...allSlots.map(s => Date.parse(s.start))));
  const last = new Date(Math.max(...allSlots.map(s => Date.parse(s.end))));
  const busyRooms = getBusy_(roomIds, first, last);
  for (const day of result.dates) {
    for (const slot of day.slots) {
      const start = Date.parse(slot.start), end = Date.parse(slot.end);
      slot.freeRooms = result.rooms.filter(r => !(busyRooms[r.id] || [])
        .some(b => b.start < end && b.end > start)).map(r => r.name);
    }
  }
  return result;
}

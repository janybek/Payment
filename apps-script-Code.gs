/**
 * Сайт оплат — API для сайта (Google Apps Script).
 * Вставить в: Таблица → Расширения → Apps Script → Code.gs
 */

const SHEET_STUDENTS = 'Студенты';
const SHEET_LOG = 'Журнал';
const TZ = 'Asia/Bishkek';
const TOKEN_HOURS = 12;
const STATUSES = ['Оплатил', 'Не оплатил'];
const ENR = 'Номер студ. билета';
const PAID_DATE = 'Дата оплаты';
// Семестры, скрытые на публичной странице (админка видит всё)
const HIDDEN_SEMESTERS = ['1'];
const EDIT_FIELDS = [ENR, 'ФИО', 'ФИО (лат.)', 'Группа', 'Курс', 'Семестр', 'Статус оплаты',
  'ИНН / паспорт', 'Email', 'Дата рождения', 'Пол', 'Проверить', 'Примечание'];

/* ================= ВХОД В API ================= */

function doGet(e) { return respond_(handle_((e && e.parameter) || {})); }

function doPost(e) {
  let p = {};
  try { p = JSON.parse(e.postData.contents); } catch (x) {}
  return respond_(handle_(p));
}

function respond_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function handle_(p) {
  try {
    switch (p.action) {
      case 'list':      return { ok: true, students: listPublic_(), deadline: deadline_() };
      case 'login':     return login_(p.login, p.password);
      case 'adminList': return adminList_(auth_(p.token));
      case 'setStatus': return setStatus_(editor_(p.token), p.id, p.status);
      case 'setStatusMany': return setStatusMany_(editor_(p.token), p.ids || [], p.status);
      case 'save':      return save_(editor_(p.token), p.student || {});
      case 'delete':    return delete_(editor_(p.token), p.id);
      case 'deleteMany': return deleteMany_(editor_(p.token), p.group, p.semester);
      default:          return { ok: false, error: 'Неизвестное действие' };
    }
  } catch (err) {
    return { ok: false, error: String(err && err.message || err), auth: !!(err && err.auth) };
  }
}

/* ================= ЧТЕНИЕ ================= */

// Что видит пользователь «только просмотр» (учителя): без ИНН, email, даты рождения и т.п.
const VIEWER_FIELDS = ['ID', ENR, 'ФИО', 'ФИО (лат.)', 'Группа', 'Курс', 'Семестр', 'Статус оплаты', PAID_DATE];

function adminList_(user) {
  const rows = readAll_().rows;
  if (user.role !== 'viewer') return { ok: true, role: user.role, deadline: deadline_(), students: rows };
  return { ok: true, role: 'viewer', deadline: deadline_(), students: rows.map(s => {
    const o = {}; VIEWER_FIELDS.forEach(f => o[f] = s[f] || ''); return o;
  }) };
}

/** Проверка, что пользователь может менять данные (не «только просмотр»). */
function editor_(token) {
  const u = auth_(token);
  if (u.role === 'viewer') throw new Error('Только просмотр — изменения недоступны');
  return u;
}

function sheet_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('Нет листа «' + name + '»');
  return sh;
}

function readAll_() {
  const sh = sheet_(SHEET_STUDENTS);
  const values = sh.getDataRange().getDisplayValues();
  const headers = values[0].map(h => String(h).trim());
  const col = {};
  headers.forEach((h, i) => { if (h) col[h] = i; });
  ['ID', 'ФИО', 'Группа', 'Статус оплаты'].forEach(h => {
    if (!(h in col)) throw new Error('В листе «Студенты» нет колонки «' + h + '»');
  });
  const rows = [];
  for (let r = 1; r < values.length; r++) {
    const v = values[r];
    if (!v[col['ID']]) continue;
    const o = { _row: r + 1 };
    headers.forEach((h, i) => { if (h) o[h] = v[i]; });
    rows.push(o);
  }
  return { sh, headers, col, rows };
}

/** Публичный список — БЕЗ ФИО: только номер студ. билета, группа, курс, семестр, статус. */
function listPublic_() {
  return readAll_().rows
    .filter(s => s[ENR] && HIDDEN_SEMESTERS.indexOf(String(s['Семестр'])) < 0)
    .map(s => ({ e: String(s[ENR]), g: s['Группа'], c: s['Курс'], s: s['Семестр'],
                 p: s['Статус оплаты'] === 'Оплатил' ? 1 : 0 }));
}

/* ================= ИЗМЕНЕНИЯ ================= */

function now_() { return Utilities.formatDate(new Date(), TZ, 'dd.MM.yyyy HH:mm'); }

function log_(user, action, id, fio, before, after) {
  sheet_(SHEET_LOG).appendRow([now_(), user.n, action, id, fio, before, after]);
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function setStatus_(user, id, status) {
  if (STATUSES.indexOf(status) < 0) throw new Error('Неверный статус');
  return withLock_(() => {
    const db = readAll_();
    const s = db.rows.find(x => x['ID'] === id);
    if (!s) throw new Error('Студент ' + id + ' не найден');
    const before = s['Статус оплаты'];
    if (before === status) return { ok: true, student: s };
    db.sh.getRange(s._row, db.col['Статус оплаты'] + 1).setValue(status);
    const pd = ensureCol_(db, PAID_DATE);
    const date = status === 'Оплатил' ? today_() : '';
    db.sh.getRange(s._row, pd + 1).setNumberFormat('@').setValue(date);
    stamp_(db, s._row, user);
    log_(user, 'Статус', id, s['ФИО'], before, status);
    s['Статус оплаты'] = status; s[PAID_DATE] = date;
    return { ok: true, student: s, date: date };
  });
}

/** Массовая отметка оплаты: одна запись в таблицу для всех выбранных. */
function setStatusMany_(user, ids, status) {
  if (STATUSES.indexOf(status) < 0) throw new Error('Неверный статус');
  if (!ids.length) throw new Error('Никто не выбран');
  if (ids.length > 2000) throw new Error('Слишком много');
  return withLock_(() => {
    const db = readAll_();
    const pd = ensureCol_(db, PAID_DATE);
    const last = db.sh.getLastRow(), n = last - 1, w = db.headers.length;
    const range = db.sh.getRange(2, 1, n, w);
    const vals = range.getValues();
    const want = {}; ids.forEach(i => want[i] = 1);
    const cS = db.col['Статус оплаты'], cI = db.col['ID'];
    const cM = db.col['Изменено'], cK = db.col['Кем изменено'];
    const date = status === 'Оплатил' ? today_() : '', stamp = now_(), logRows = [];
    vals.forEach(row => {
      if (!want[row[cI]] || row[cS] === status) return;
      logRows.push([stamp, user.n, 'Статус (массово)', row[cI], row[db.col['ФИО']], row[cS], status]);
      row[cS] = status; row[pd] = date;
      if (cM != null) row[cM] = stamp;
      if (cK != null) row[cK] = user.n;
    });
    if (logRows.length) {
      db.sh.getRange(2, pd + 1, n, 1).setNumberFormat('@');
      range.setValues(vals);
      const lg = sheet_(SHEET_LOG);
      lg.getRange(lg.getLastRow() + 1, 1, logRows.length, 7).setValues(logRows);
    }
    return { ok: true, count: logRows.length, date: date };
  });
}

function today_() { return Utilities.formatDate(new Date(), TZ, 'dd.MM.yyyy'); }

/** Добавляет колонку в конец листа, если её нет. Возвращает индекс (с 0). */
function ensureCol_(db, name) {
  if (name in db.col) return db.col[name];
  const c = db.headers.length;
  db.sh.getRange(1, c + 1).setValue(name).setFontWeight('bold');
  db.headers.push(name); db.col[name] = c;
  return c;
}

/** Срок оплаты (dd.MM.yyyy) — задаётся в таблице: меню «Сайт оплат → Срок оплаты». */
function deadline_() { return props_().getProperty('DEADLINE') || ''; }

function stamp_(db, row, user) {
  if ('Изменено' in db.col) db.sh.getRange(row, db.col['Изменено'] + 1).setValue(now_());
  if ('Кем изменено' in db.col) db.sh.getRange(row, db.col['Кем изменено'] + 1).setValue(user.n);
}

function clean_(st) {
  const o = {};
  EDIT_FIELDS.forEach(f => { if (f in st) o[f] = String(st[f] == null ? '' : st[f]).trim(); });
  if ('ФИО' in o && !o['ФИО']) throw new Error('Укажите ФИО');
  if ('Группа' in o && !o['Группа']) throw new Error('Укажите группу');
  if ('Статус оплаты' in o && STATUSES.indexOf(o['Статус оплаты']) < 0) throw new Error('Неверный статус');
  if (o[ENR] && !/^\d{4,10}$/.test(o[ENR])) throw new Error('Номер студ. билета — только цифры');
  ['Курс', 'Семестр'].forEach(f => {
    if (o[f]) { if (!/^\d{1,2}$/.test(o[f])) throw new Error('«' + f + '» — только число'); o[f] = Number(o[f]); }
  });
  return o;
}

function save_(user, st) {
  const data = clean_(st);
  return withLock_(() => {
    const db = readAll_();
    if (data[ENR]) {
      const twin = db.rows.find(x => String(x[ENR]) === data[ENR] && x['ID'] !== st['ID']);
      if (twin) throw new Error('Номер ' + data[ENR] + ' уже есть у студента ' + twin['ID'] + ' (' + twin['Группа'] + ')');
    }
    if (st['ID']) {
      // --- редактирование ---
      const s = db.rows.find(x => x['ID'] === st['ID']);
      if (!s) throw new Error('Студент ' + st['ID'] + ' не найден');
      const changes = [];
      Object.keys(data).forEach(f => {
        if (!(f in db.col)) return;
        if (String(s[f]) === String(data[f])) return;
        db.sh.getRange(s._row, db.col[f] + 1).setValue(data[f]);
        changes.push([f, s[f], data[f]]);
      });
      const sc = changes.find(c => c[0] === 'Статус оплаты');
      if (sc) {
        const pd = ensureCol_(db, PAID_DATE);
        db.sh.getRange(s._row, pd + 1).setNumberFormat('@').setValue(sc[2] === 'Оплатил' ? today_() : '');
      }
      if (changes.length) {
        stamp_(db, s._row, user);
        log_(user, 'Редактирование', s['ID'], data['ФИО'] || s['ФИО'],
          changes.map(c => c[0] + ': ' + c[1]).join('; '),
          changes.map(c => c[0] + ': ' + c[2]).join('; '));
      }
      return { ok: true, id: s['ID'], changed: changes.length };
    }
    // --- новый студент ---
    if (!data['ФИО'] || !data['Группа']) throw new Error('Укажите ФИО и группу');
    if (!data['Статус оплаты']) data['Статус оплаты'] = 'Не оплатил';
    let max = 0;
    db.rows.forEach(x => { const m = /^S(\d+)$/.exec(x['ID']); if (m) max = Math.max(max, +m[1]); });
    const bin = SpreadsheetApp.getActive().getSheetByName('Удалённые');   // не выдавать ID удалённых повторно
    if (bin && bin.getLastRow() > 1) bin.getRange(2, 3, bin.getLastRow() - 1, 1).getValues()
      .forEach(r => { const m = /^S(\d+)$/.exec(r[0]); if (m) max = Math.max(max, +m[1]); });
    const id = 'S' + String(max + 1).padStart(4, '0');
    const row = db.headers.map(h => {
      if (h === 'ID') return id;
      if (h === 'Изменено') return now_();
      if (h === 'Кем изменено') return user.n;
      if (h === PAID_DATE) return data['Статус оплаты'] === 'Оплатил' ? today_() : '';
      return h in data ? data[h] : '';
    });
    db.sh.appendRow(row);
    log_(user, 'Добавлен', id, data['ФИО'], '', data['Группа'] + ', ' + data['Статус оплаты']);
    return { ok: true, id: id };
  });
}

/** Удаление: строка переносится на лист «Удалённые» (можно вернуть вручную). */
function delete_(user, id) {
  if (!id) throw new Error('Не указан студент');
  return withLock_(() => {
    const db = readAll_();
    const s = db.rows.find(x => x['ID'] === id);
    if (!s) throw new Error('Студент ' + id + ' не найден');
    const bin = binSheet_(db.headers);
    const rowVals = db.sh.getRange(s._row, 1, 1, db.headers.length).getValues()[0];
    bin.appendRow([now_(), user.n].concat(rowVals));
    db.sh.deleteRow(s._row);
    log_(user, 'Удалён', id, s['ФИО'], s['Группа'] + ', ' + s['Статус оплаты'], 'перенесён на лист «Удалённые»');
    return { ok: true, id: id };
  });
}

function binSheet_(headers) {
  const ss = SpreadsheetApp.getActive();
  let bin = ss.getSheetByName('Удалённые');
  if (!bin) {
    bin = ss.insertSheet('Удалённые');
    bin.appendRow(['Удалено', 'Кем удалено'].concat(headers));
    bin.getRange(1, 1, 1, headers.length + 2).setFontWeight('bold');
    bin.setFrozenRows(1);
  }
  return bin;
}

/** Массовое удаление: вся группа и/или весь семестр. Строки уходят на лист «Удалённые». */
function deleteMany_(user, group, semester) {
  group = String(group || '').trim(); semester = String(semester || '').trim();
  if (!group && !semester) throw new Error('Не выбрана группа или семестр');
  return withLock_(() => {
    const db = readAll_();
    const width = db.headers.length;
    const last = db.sh.getLastRow();
    if (last < 2) throw new Error('Список пуст');
    const vals = db.sh.getRange(2, 1, last - 1, width).getValues();
    const disp = db.sh.getRange(2, 1, last - 1, width).getDisplayValues();
    const cG = db.col['Группа'], cS = db.col['Семестр'], cId = db.col['ID'];
    const hit = r => disp[r][cId] && (!group || disp[r][cG] === group) && (!semester || String(disp[r][cS]) === semester);
    const keep = [], gone = [];
    vals.forEach((row, r) => (hit(r) ? gone : keep).push(row));
    if (!gone.length) throw new Error('Никого не найдено для удаления');
    const stamp = now_();
    const bin = binSheet_(db.headers);
    bin.getRange(bin.getLastRow() + 1, 1, gone.length, width + 2)
       .setValues(gone.map(row => [stamp, user.n].concat(row)));
    // перезаписываем лист без удалённых строк
    if (keep.length) db.sh.getRange(2, 1, keep.length, width).setValues(keep);
    db.sh.getRange(2 + keep.length, 1, gone.length, width).clearContent();
    const what = [group ? 'группа ' + group : '', semester ? 'семестр ' + semester : ''].filter(String).join(', ');
    log_(user, 'Удалено массово', '', what, gone.length + ' студ.', 'перенесены на лист «Удалённые»');
    return { ok: true, count: gone.length };
  });
}

/* ================= ВХОД / ПАРОЛИ ================= */

function props_() { return PropertiesService.getScriptProperties(); }

function users_() {
  try { return JSON.parse(props_().getProperty('USERS') || '{}'); } catch (e) { return {}; }
}

function secret_() {
  let s = props_().getProperty('SECRET');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); props_().setProperty('SECRET', s); }
  return s;
}

function hash_(salt, pass) {
  let h = salt + '|' + pass;
  for (let i = 0; i < 500; i++) {
    h = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h + salt));
  }
  return h;
}

function login_(login, password) {
  login = String(login || '').trim().toLowerCase();
  const cache = CacheService.getScriptCache();
  const failKey = 'fail_' + login;
  const fails = +(cache.get(failKey) || 0);
  if (fails >= 5) return { ok: false, error: 'Слишком много попыток. Подождите 10 минут.' };
  const u = users_()[login];
  if (!u || hash_(u.salt, String(password || '')) !== u.hash) {
    cache.put(failKey, String(fails + 1), 600);
    Utilities.sleep(800);
    return { ok: false, error: 'Неверный логин или пароль' };
  }
  cache.remove(failKey);
  const payload = Utilities.base64EncodeWebSafe(Utilities.newBlob(JSON.stringify(
    { u: login, n: u.name, v: u.ver || 1, e: Date.now() + TOKEN_HOURS * 3600e3 })).getBytes());
  return { ok: true, token: payload + '.' + sign_(payload), name: u.name, role: u.role || 'editor' };
}

function sign_(payload) {
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(payload, secret_()));
}

function auth_(token) {
  const fail = () => { const e = new Error('Сессия истекла, войдите заново'); e.auth = true; throw e; };
  if (!token || token.indexOf('.') < 0) fail();
  const [payload, sig] = token.split('.');
  if (sign_(payload) !== sig) fail();
  let d;
  try { d = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(payload)).getDataAsString()); } catch (e) { fail(); }
  if (!d || d.e < Date.now()) fail();
  const u = users_()[d.u];
  if (!u || (u.ver || 1) !== d.v) fail(); // пользователь удалён или сменил пароль
  return { login: d.u, n: u.name, role: u.role || 'editor' };
}

/* ================= МЕНЮ В ТАБЛИЦЕ (только для владельца) ================= */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Сайт оплат')
    .addItem('Добавить пользователя / сменить пароль', 'menuSetUser')
    .addItem('Добавить учителя (только просмотр)', 'menuSetViewer')
    .addItem('Срок оплаты (баннер на сайте)', 'menuDeadline')
    .addItem('Список пользователей', 'menuListUsers')
    .addItem('Удалить пользователя', 'menuDeleteUser')
    .addToUi();
}

function ask_(title, text) {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt(title, text, ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return null;
  return r.getResponseText().trim();
}

function menuSetViewer() { menuSetUser('viewer'); }

function menuSetUser(role) {
  role = role === 'viewer' ? 'viewer' : 'editor';
  const ui = SpreadsheetApp.getUi();
  const login = (ask_('Логин', 'Латиницей, например: dekan') || '').toLowerCase();
  if (!login) return;
  if (!/^[a-z0-9._-]{3,30}$/.test(login)) return ui.alert('Логин: 3–30 символов, латиница/цифры');
  const users = users_();
  const old = users[login];
  const name = ask_('Имя', 'Как показывать в журнале, например: Декан' + (old ? '\nСейчас: ' + old.name : ''));
  if (name === null) return;
  const pass = ask_('Пароль', 'Минимум 8 символов');
  if (!pass) return;
  if (pass.length < 8) return ui.alert('Пароль слишком короткий — минимум 8 символов');
  const salt = Utilities.getUuid();
  users[login] = { name: name || (old && old.name) || login, salt: salt, hash: hash_(salt, pass),
                   ver: old ? (old.ver || 1) + 1 : 1, role: role };
  props_().setProperty('USERS', JSON.stringify(users));
  secret_();
  ui.alert('Готово: ' + login + ' (' + users[login].name + ') — ' + (role === 'viewer' ? 'только просмотр' : 'может изменять') + (old ? '\nСтарые входы этого пользователя сброшены.' : ''));
}

function menuDeadline() {
  const ui = SpreadsheetApp.getUi();
  const cur = deadline_();
  const v = ask_('Срок оплаты', 'Дата в формате дд.мм.гггг, например 01.12.2026.\nПусто — убрать баннер.' + (cur ? '\nСейчас: ' + cur : ''));
  if (v === null) return;
  if (v && !/^\d{2}\.\d{2}\.\d{4}$/.test(v)) return ui.alert('Формат: дд.мм.гггг');
  if (v) props_().setProperty('DEADLINE', v); else props_().deleteProperty('DEADLINE');
  ui.alert(v ? 'Срок оплаты: ' + v : 'Баннер убран');
}

function menuListUsers() {
  const users = users_();
  const list = Object.keys(users).map(l => '• ' + l + ' — ' + users[l].name + (users[l].role === 'viewer' ? ' (просмотр)' : ' (изменение)')).join('\n');
  SpreadsheetApp.getUi().alert(list || 'Пользователей пока нет');
}

function menuDeleteUser() {
  const ui = SpreadsheetApp.getUi();
  const login = (ask_('Удалить пользователя', 'Введите логин') || '').toLowerCase();
  if (!login) return;
  const users = users_();
  if (!users[login]) return ui.alert('Нет такого пользователя');
  delete users[login];
  props_().setProperty('USERS', JSON.stringify(users));
  ui.alert('Удалён: ' + login + '. Его вход на сайт сразу перестаёт работать.');
}

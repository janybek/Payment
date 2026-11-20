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
// Медстраховка — видна только деканату (не учителям и не на публичной странице)
const INS_UNTIL = 'Страховка до';
const INS_PAID = 'Страховка оплата';
const LEADER_COL = 'Староста';
const LEADER_LOGIN = 'Логин старосты';
// Семестры, скрытые на публичной странице (админка видит всё)
const HIDDEN_SEMESTERS = ['1'];
const EDIT_FIELDS = [INS_UNTIL, INS_PAID, ENR, 'ФИО', 'ФИО (лат.)', 'Группа', 'Курс', 'Семестр', 'Статус оплаты',
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
      case 'setStatus': return p.kind === 'ins'
        ? setStatus_(insEditor_(p.token, [p.id]), p.id, p.status, 'ins')
        : setStatus_(editor_(p.token), p.id, p.status, p.kind);
      case 'setStatusMany': return p.kind === 'ins'
        ? setStatusMany_(insEditor_(p.token, p.ids || []), p.ids || [], p.status, 'ins')
        : setStatusMany_(editor_(p.token), p.ids || [], p.status, p.kind);
      case 'save':      return saveAny_(auth_(p.token), p.student || {});
      case 'setLeader': return setLeader_(editor_(p.token), p.id, !!p.on, p.email, !!p.reset);
      case 'history':   return history_(auth_(p.token), p.id);
      case 'repeatList':   return repeatList_(auth_(p.token));
      case 'repeatSave':   return repeatSave_(editor_(p.token), p.rec || {});
      case 'repeatDelete': return repeatDelete_(editor_(p.token), p.rid);
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
const VIEWER_FIELDS = ['ID', ENR, 'ФИО', 'ФИО (лат.)', 'Группа', 'Курс', 'Семестр', 'Статус оплаты', PAID_DATE, LEADER_COL];

// Что видит староста: только свою группу, без оплаты контракта
const LEADER_FIELDS = ['ID', ENR, 'ФИО', 'ФИО (лат.)', 'Группа', 'Курс', 'Семестр', 'Email', 'Пол', 'Дата рождения', INS_UNTIL, INS_PAID, LEADER_COL];
// Что староста может менять (личные данные + страховка: дата и оплата)
const LEADER_EDIT = ['ФИО', 'ФИО (лат.)', 'Email', ENR, 'Пол', 'Дата рождения', INS_UNTIL, INS_PAID];

function adminList_(user) {
  const rows = readAll_().rows;
  if (user.role === 'leader') return { ok: true, role: 'leader', groups: leaderGroups_(user, rows), students: rows
    .filter(s => leaderGroups_(user, rows).indexOf(s['Группа']) >= 0)
    .map(s => { const o = {}; LEADER_FIELDS.forEach(f => o[f] = s[f] || ''); return o; }) };
  if (user.role !== 'viewer') return { ok: true, role: user.role, deadline: deadline_(), students: rows };
  return { ok: true, role: 'viewer', deadline: deadline_(), students: rows.map(s => {
    const o = {}; VIEWER_FIELDS.forEach(f => o[f] = s[f] || ''); return o;
  }) };
}

/** Проверка, что пользователь может менять данные (не «только просмотр»). */
function editor_(token) {
  const u = auth_(token);
  if (u.role !== 'editor') throw new Error('Только просмотр — изменения недоступны');
  return u;
}

/** Группа старосты = текущая группа его студенческой записи (перевели — права переехали). */
function leaderGroups_(user, rows) {
  const me = (rows || readAll_().rows).find(s => s['ID'] === user.sid);
  return me && String(me[LEADER_COL]) === 'Да' ? [me['Группа']] : [];
}

/** Оплату страховки меняют деканат и староста (только в своей группе). Оплату контракта — только деканат. */
function insEditor_(token, ids) {
  const u = auth_(token);
  if (u.role === 'editor') return u;
  if (u.role !== 'leader') throw new Error('Только просмотр — изменения недоступны');
  const rows = readAll_().rows, my = leaderGroups_(u, rows), want = {};
  ids.forEach(i => want[i] = 1);
  if (!ids.length || rows.some(s => want[s['ID']] && my.indexOf(s['Группа']) < 0))
    throw new Error('Можно менять только студентов своей группы');
  return u;
}

/** Сохранение карточки: деканат — всё; староста — только личные данные своей группы. */
function saveAny_(user, st) {
  if (user.role === 'editor') return save_(user, st);
  if (user.role !== 'leader') throw new Error('Только просмотр — изменения недоступны');
  if (!st['ID']) throw new Error('Староста не может добавлять студентов');
  const rows = readAll_().rows, s = rows.find(x => x['ID'] === st['ID']);
  if (!s || leaderGroups_(user, rows).indexOf(s['Группа']) < 0) throw new Error('Можно менять только студентов своей группы');
  const locked = Object.keys(st).filter(f => f !== 'ID' && LEADER_EDIT.indexOf(f) < 0 && EDIT_FIELDS.indexOf(f) >= 0 &&
                                            String(st[f] == null ? '' : st[f]).trim() !== String(s[f] == null ? '' : s[f]).trim());
  if (locked.length) throw new Error('Староста не может менять: ' + locked.join(', '));
  const data = { ID: st['ID'] };
  LEADER_EDIT.forEach(f => { if (f in st) data[f] = st[f]; });
  return save_(user, data);
}

/** История изменений одного студента из «Журнала». */
function history_(user, id) {
  if (user.role === 'leader') {
    const rows = readAll_().rows, s = rows.find(x => x['ID'] === id);
    if (!s || leaderGroups_(user, rows).indexOf(s['Группа']) < 0) throw new Error('Нет доступа');
  } else if (user.role !== 'editor') throw new Error('Нет доступа');
  const lg = sheet_(SHEET_LOG), n = lg.getLastRow() - 1;
  if (n < 1) return { ok: true, items: [] };
  const v = lg.getRange(2, 1, n, 7).getDisplayValues();
  const items = v.filter(r => r[3] === id).slice(-30).reverse()
    .map(r => ({ when: r[0], who: r[1], what: r[2], before: r[5], after: r[6] }));
  return { ok: true, items: items };
}

/** Назначить / снять старосту, выдать новый пароль. Логин — Gmail. */
function setLeader_(user, id, on, email, reset) {
  email = String(email || '').trim().toLowerCase();
  return withLock_(() => {
    const db = readAll_();
    const s = db.rows.find(x => x['ID'] === id);
    if (!s) throw new Error('Студент не найден');
    const cL = ensureCol_(db, LEADER_COL), cE = ensureCol_(db, LEADER_LOGIN);
    const L = leaders_();
    const mine = Object.keys(L).filter(k => L[k].sid === id);
    let pass = '';
    if (!on) {
      mine.forEach(k => delete L[k]);
      db.sh.getRange(s._row, cL + 1).setValue('');
      db.sh.getRange(s._row, cE + 1).setValue('');
      log_(user, 'Староста снят', id, s['ФИО'], s[LEADER_LOGIN] || '', '');
    } else {
      if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('Укажите почту (Gmail) старосты');
      if (L[email] && L[email].sid !== id) throw new Error('Эта почта уже логин другого старосты');
      const prev = mine.length ? L[mine[0]] : null;
      mine.filter(k => k !== email).forEach(k => delete L[k]);
      if (!prev || reset || mine[0] !== email) {
        pass = newPassword_();
        const salt = Utilities.getUuid();
        L[email] = { sid: id, salt: salt, hash: hash_(salt, pass), ver: prev ? (prev.ver || 1) + 1 : 1 };
      } else L[email] = prev;
      db.sh.getRange(s._row, cL + 1).setValue('Да');
      db.sh.getRange(s._row, cE + 1).setValue(email);
      log_(user, pass ? (prev ? 'Староста: новый пароль' : 'Назначен старостой') : 'Староста', id, s['ФИО'],
           s[LEADER_LOGIN] || '', email + ' (' + s['Группа'] + ')');
    }
    props_().setProperty('LEADERS', JSON.stringify(L));
    secret_();
    stamp_(db, s._row, user);
    return { ok: true, password: pass, email: on ? email : '' };
  });
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

function setStatus_(user, id, status, kind) {
  if (STATUSES.indexOf(status) < 0) throw new Error('Неверный статус');
  if (kind === 'ins') return setInsPaid_(user, [id], status, false);
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
/** Оплата страховки: одному или нескольким студентам. */
function setInsPaid_(user, ids, status, many) {
  if (!ids.length) throw new Error('Никто не выбран');
  return withLock_(() => {
    const db = readAll_();
    const c = ensureCol_(db, INS_PAID);
    const want = {}; ids.forEach(i => want[i] = 1);
    const stamp = now_(), logRows = [];
    db.rows.forEach(s => {
      if (!want[s['ID']] || (s[INS_PAID] || '') === status) return;
      db.sh.getRange(s._row, c + 1).setValue(status);
      stamp_(db, s._row, user);
      logRows.push([stamp, user.n, many ? 'Страховка (массово)' : 'Страховка', s['ID'], s['ФИО'], s[INS_PAID] || '', status]);
    });
    if (logRows.length) {
      const lg = sheet_(SHEET_LOG);
      lg.getRange(lg.getLastRow() + 1, 1, logRows.length, 7).setValues(logRows);
    }
    return { ok: true, count: logRows.length };
  });
}

function setStatusMany_(user, ids, status, kind) {
  if (STATUSES.indexOf(status) < 0) throw new Error('Неверный статус');
  if (!ids.length) throw new Error('Никто не выбран');
  if (ids.length > 2000) throw new Error('Слишком много');
  if (kind === 'ins') return setInsPaid_(user, ids, status, true);
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

function validDate_(v) {
  const m = /^(\d{2})\.(\d{2})\.(\d{4})$/.exec(v || '');
  if (!m) return false;
  const d = new Date(+m[3], +m[2] - 1, +m[1]);
  return d.getFullYear() === +m[3] && d.getMonth() === +m[2] - 1 && d.getDate() === +m[1];
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
  if (o[INS_UNTIL] && o[INS_UNTIL].toLowerCase() !== 'нет' && !validDate_(o[INS_UNTIL]))
    throw new Error('«Страховка до» — дата дд.мм.гггг или «нет»');
  if (o[INS_PAID] && STATUSES.indexOf(o[INS_PAID]) < 0) throw new Error('Неверный статус страховки');
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
        if (!(f in db.col)) {
          if ((f === INS_UNTIL || f === INS_PAID) && data[f]) ensureCol_(db, f); else return;
        }
        if (String(s[f]) === String(data[f])) return;
        db.sh.getRange(s._row, db.col[f] + 1).setValue(f === INS_UNTIL && data[f] ? "'" + data[f] : data[f]);
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
      if (h === INS_UNTIL && data[h]) return "'" + data[h];
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
  const isLeader = login.indexOf('@') > 0;              // старосты входят по Gmail
  const key = isLeader ? 'mail:' + login : login;
  const u = isLeader ? leaders_()[login] : users_()[login];
  if (!u || hash_(u.salt, String(password || '').trim()) !== u.hash) {
    cache.put(failKey, String(fails + 1), 600);
    Utilities.sleep(800);
    return { ok: false, error: 'Неверный логин или пароль' };
  }
  cache.remove(failKey);
  const role = isLeader ? 'leader' : (u.role || 'editor');
  const name = isLeader ? leaderName_(u) : u.name;
  const payload = Utilities.base64EncodeWebSafe(Utilities.newBlob(JSON.stringify(
    { u: key, n: name, v: u.ver || 1, e: Date.now() + TOKEN_HOURS * 3600e3 })).getBytes());
  return { ok: true, token: payload + '.' + sign_(payload), name: name, role: role };
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
  if (String(d.u).indexOf('mail:') === 0) {          // староста
    const L = leaders_()[d.u.slice(5)];
    if (!L || (L.ver || 1) !== d.v) fail();
    return { login: d.u, n: d.n, role: 'leader', sid: L.sid };
  }
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

/* ================= ПОВТОРНЫЙ КУРС ================= */
// Лист «Повторный курс»: видят деканат и учителя, меняет только деканат. Студенты и старосты не видят.
const SHEET_REPEAT = 'Повторный курс';
const REPEAT_HEAD = ['RID', 'ID студента', 'ФИО (по приказу)', 'Группа', 'Семестр', 'Предмет', 'За семестр',
                     'Форма контроля', 'Преподаватель', 'Часы', 'Подгруппа', 'Приказ', 'Примечание'];

function repeatSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_REPEAT);
  if (!sh) {
    sh = ss.insertSheet(SHEET_REPEAT);
    sh.getRange(1, 1, 1, REPEAT_HEAD.length).setValues([REPEAT_HEAD]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function readRepeat_() {
  const sh = repeatSheet_();
  const v = sh.getDataRange().getDisplayValues();
  const h = v[0].map(String), rows = [];
  for (let i = 1; i < v.length; i++) {
    if (!v[i][0]) continue;
    const o = { _row: i + 1 };
    h.forEach((k, j) => { if (k) o[k] = v[i][j]; });
    rows.push(o);
  }
  return { sh: sh, head: h, rows: rows };
}

function repeatList_(user) {
  if (user.role !== 'editor' && user.role !== 'viewer') throw new Error('Нет доступа');
  return { ok: true, rows: readRepeat_().rows.map(r => { delete r._row; return r; }) };
}

function repeatSave_(user, rec) {
  const o = {};
  REPEAT_HEAD.slice(1).forEach(k => { if (k in rec) o[k] = String(rec[k] == null ? '' : rec[k]).trim(); });
  ['ФИО (по приказу)', 'Группа', 'Предмет'].forEach(k => { if (!o[k]) throw new Error('Заполните поле «' + k + '»'); });
  ['Семестр', 'За семестр', 'Часы'].forEach(k => { if (o[k] && !/^\d{1,3}$/.test(o[k])) throw new Error('«' + k + '» — только число'); });
  return withLock_(() => {
    const db = readRepeat_();
    const col = {}; db.head.forEach((k, j) => col[k] = j);
    if (rec.RID) {
      const r = db.rows.find(x => x.RID === rec.RID);
      if (!r) throw new Error('Запись не найдена');
      const ch = [];
      Object.keys(o).forEach(k => {
        if (!(k in col) || String(r[k] || '') === o[k]) return;
        db.sh.getRange(r._row, col[k] + 1).setValue(o[k]); ch.push([k, r[k] || '', o[k]]);
      });
      if (ch.length) log_(user, 'Повторный курс: изменено', rec.RID, o['ФИО (по приказу)'],
                          ch.map(c => c[0] + ': ' + c[1]).join('; '), ch.map(c => c[0] + ': ' + c[2]).join('; '));
      return { ok: true, rid: rec.RID };
    }
    let max = 0;
    db.rows.forEach(x => { const m = /^R(\d+)$/.exec(x.RID); if (m) max = Math.max(max, +m[1]); });
    const rid = 'R' + String(max + 1).padStart(4, '0');
    db.sh.appendRow(db.head.map(k => k === 'RID' ? rid : (o[k] || '')));
    log_(user, 'Повторный курс: добавлено', rid, o['ФИО (по приказу)'], '', o['Предмет'] + ' · ' + o['Группа']);
    return { ok: true, rid: rid };
  });
}

function repeatDelete_(user, rid) {
  return withLock_(() => {
    const db = readRepeat_();
    const r = db.rows.find(x => x.RID === rid);
    if (!r) throw new Error('Запись не найдена');
    db.sh.deleteRow(r._row);
    log_(user, 'Повторный курс: удалено', rid, r['ФИО (по приказу)'], r['Предмет'] + ' · ' + r['Группа'] + ' · ' + r['Преподаватель'], '');
    return { ok: true };
  });
}

/* ================= СТАРОСТЫ ================= */
// Хранятся в свойствах скрипта: { gmail: { sid: ID студента, salt, hash, ver } }.
// Назначаются на сайте: карточка студента → «★ Староста группы».
function leaders_() {
  try { return JSON.parse(props_().getProperty('LEADERS') || '{}'); } catch (e) { return {}; }
}
function leaderName_(L) {
  const s = readAll_().rows.find(x => x['ID'] === L.sid);
  return s ? 'Староста ' + s['Группа'] + ' (' + s['ФИО'] + ')' : 'Староста';
}
function newPassword_() {
  const abc = 'abcdefghjkmnpqrstuvwxyz23456789';     // без похожих символов (l/1, o/0)
  let s = '';
  for (let i = 0; i < 6; i++) s += abc.charAt(Math.floor(Math.random() * abc.length));
  return s;
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

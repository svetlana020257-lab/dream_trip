/**
 * ГНУ · Путь героя — сервер игры (Google Apps Script)
 *
 * Хранит состояние игры в этой таблице, проверяет PIN куратора,
 * принимает отметки заданий и анкеты героев.
 *
 * Настройки (Проект → Настройки проекта → Свойства скрипта):
 *   CURATOR_PIN      — PIN для входа куратора (обязательно)
 *   SITE_URL         — адрес сайта на GitHub Pages (для листа «Ссылки»)
 *
 * После правки кода: Развернуть → Управление развёртываниями → карандаш →
 * Версия: «Новая версия» → Развернуть. Ссылка /exec останется прежней.
 */

const STATE_SHEET = 'state';
// Служебные листы state, claims и profiles не трогай руками. Смотреть удобно листы «Герои», «Анкеты», «Журнал», «Сводка», «Ссылки».
const CHUNK = 40000;

function doGet(e) {
  const p = (e && e.parameter) || {};
  const action = p.action || 'state';
  if (action === 'state') {
    const state = loadState_();
    const profiles = loadProfiles_();
    const out = { ok: true, state: publicState_(state, profiles) };
    const hero = heroByCode_(state, p.code);
    if (hero) {
      out.me = hero.id;
      const c = loadClaims_()[hero.id];
      out.claims = c ? { [hero.id]: c } : {};
      out.profile = profiles[hero.id] || null;
    }
    return json_(out);
  }
  return json_({ ok: false, error: 'Неизвестное действие' });
}

/* участникам не отдаём чужие личные коды, служебную копию и чужие свитки.
   Из анкет других героев видно только имя, класс, стихию и снаряжение. */
function publicState_(state, profiles) {
  if (!state) return null;
  const s = JSON.parse(JSON.stringify(state));
  delete s.prev;
  (s.heroes || []).forEach(h => {
    delete h.code;
    const p = profiles && profiles[h.id];
    if (p) h.look = { name: p.name || '', cls: p.cls || '', v: p.v == 2 ? 2 : 1, triple: p.triple || '', gear: p.gear || null, onb: !!p.onb };
  });
  return s;
}
function heroByCode_(state, code) {
  if (!state || !code) return null;
  return (state.heroes || []).find(h => h.code && h.code === String(code)) || null;
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); }
  catch (err) { return json_({ ok: false, error: 'Неверный запрос' }); }

  try {
    if (body.action === 'check') return json_({ ok: checkPin_(body.pin) });

    if (body.action === 'state') {
      if (!checkPin_(body.pin)) return json_({ ok: false, error: 'Неверный PIN куратора' });
      return json_({ ok: true, full: true, state: loadState_(), claims: loadClaims_(), profiles: loadProfiles_() });
    }

    if (body.action === 'claim') {
      const state = loadState_();
      const hero = heroByCode_(state, body.code);
      if (!hero) return json_({ ok: false, error: 'Ссылка героя не найдена. Попроси у куратора новую.' });
      const c = body.claim || {};
      const clean = {
        week: state.week,
        task: !!c.task, refl: !!c.refl, sphere: !!c.sphere, trap: !!c.trap,
        min: (Array.isArray(c.min) ? c.min : []).slice(0, 7).map(x => (x ? 1 : 0)),
        at: new Date().toISOString(),
      };
      while (clean.min.length < 7) clean.min.push(0);
      const lock = LockService.getScriptLock();
      lock.waitLock(20000);
      try { saveClaim_(hero.id, clean); } finally { lock.releaseLock(); }
      return json_({ ok: true });
    }

    if (body.action === 'profile') {
      const state = loadState_();
      const hero = heroByCode_(state, body.code);
      if (!hero) return json_({ ok: false, error: 'Ссылка героя не найдена. Попроси у куратора новую.' });
      const p = body.profile || {};
      const str = (v, n) => String(v == null ? '' : v).slice(0, n);
      const g = p.gear || {};
      const clean = {
        name: str(p.name, 40).trim(),
        cls: CLASS_NAMES[p.cls] ? p.cls : '',
        triple: TRIPLE_NAMES[p.triple] ? p.triple : '',
        v: p.v == 2 ? 2 : 1,
        gear: { glow: str(g.glow, 20), aura: str(g.aura, 20), amulet: str(g.amulet, 20), plate: str(g.plate, 20) },
        scroll: { goal: str((p.scroll || {}).goal, 400), train: str((p.scroll || {}).train, 400), boost: str((p.scroll || {}).boost, 400) },
        onb: !!p.onb,
        at: new Date().toISOString(),
      };
      const lock = LockService.getScriptLock();
      lock.waitLock(20000);
      try { saveRow_('profiles', hero.id, clean); mirrorProfiles_(state); } finally { lock.releaseLock(); }
      return json_({ ok: true });
    }

    if (body.action === 'save') {
      if (!checkPin_(body.pin)) return json_({ ok: false, error: 'Неверный PIN куратора' });
      const lock = LockService.getScriptLock();
      lock.waitLock(20000);
      try {
        saveState_(body.state);
        mirrorSheets_(body.state);
      } finally { lock.releaseLock(); }
      return json_({ ok: true });
    }

    return json_({ ok: false, error: 'Неизвестное действие' });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

/* ---------------- хранение ---------------- */

function sheet_(name) {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function loadState_() {
  const sh = sheet_(STATE_SHEET);
  const last = sh.getLastRow();
  if (!last) return null;
  const text = sh.getRange(1, 1, last, 1).getValues().map(r => r[0]).join('');
  if (!text) return null;
  try { return JSON.parse(text); } catch (e) { return null; }
}

function saveState_(state) {
  const text = JSON.stringify(state || {});
  const rows = [];
  for (let i = 0; i < text.length; i += CHUNK) rows.push([text.slice(i, i + CHUNK)]);
  const sh = sheet_(STATE_SHEET);
  sh.clearContents();
  sh.getRange(1, 1, rows.length, 1).setNumberFormat('@').setValues(rows);
}

/* заявки участников: лист claims, строка на героя */
function loadClaims_() {
  const sh = sheet_('claims');
  const last = sh.getLastRow();
  const out = {};
  if (!last) return out;
  sh.getRange(1, 1, last, 2).getValues().forEach(r => {
    if (!r[0]) return;
    try { out[r[0]] = JSON.parse(r[1]); } catch (e) {}
  });
  return out;
}
function saveClaim_(heroId, claim) { saveRow_('claims', heroId, claim); }

/* анкеты героев: лист profiles, строка на героя */
function loadProfiles_() {
  const sh = sheet_('profiles');
  const last = sh.getLastRow();
  const out = {};
  if (!last) return out;
  sh.getRange(1, 1, last, 2).getValues().forEach(r => {
    if (!r[0]) return;
    try { out[r[0]] = JSON.parse(r[1]); } catch (e) {}
  });
  return out;
}

/* строка «id → JSON» в служебном листе */
function saveRow_(name, heroId, obj) {
  const sh = sheet_(name);
  const last = sh.getLastRow();
  const ids = last ? sh.getRange(1, 1, last, 1).getValues().map(r => r[0]) : [];
  const i = ids.indexOf(heroId);
  const row = i >= 0 ? i + 1 : last + 1;
  sh.getRange(row, 1, 1, 2).setNumberFormat('@').setValues([[heroId, JSON.stringify(obj)]]);
}

function checkPin_(pin) {
  const real = PropertiesService.getScriptProperties().getProperty('CURATOR_PIN');
  return !!real && String(pin || '') === String(real);
}

/* ---------------- красивые листы для просмотра ---------------- */

const CLASS_NAMES = { voin: 'Воин', hran: 'Хранитель', torg: 'Торговец', cel: 'Целитель', pal: 'Паладин', mag: 'Маг' };
const TRIPLE_NAMES = { fire: 'Огонь', water: 'Вода', earth: 'Земля', air: 'Воздух' };

function world_(n) { return n <= 0 ? 'Старт' : n <= 40 ? 'Тьма' : n <= 80 ? 'Рассвет' : 'Золото'; }

function mirrorSheets_(s) {
  if (!s || !s.heroes) return;
  // Герои
  const heroes = s.heroes.slice().sort((a, b) => b.pos - a.pos).map(h => {
    const st = [];
    if (h.captured) st.push('в плену (копит ' + (h.held || 0) + ')');
    if (h.atCeiling) st.push('у потолка ' + h.atCeiling);
    if (h.trap) st.push('в ловушке');
    if (h.pos >= 120) st.push('на вершине');
    return [h.name, CLASS_NAMES[h.cls] || '', TRIPLE_NAMES[h.triple] || '', h.pos, world_(h.pos),
            (h.broken || []).join(', '), st.join('; ') || '—', h.last == null ? '' : h.last];
  });
  writeTable_('Герои', ['Герой', 'Класс', 'Тройка', 'Клетка', 'Мир', 'Пробиты потолки', 'Статус', 'Ходы за последнюю неделю'], heroes);

  // Журнал ходов
  const names = {};
  s.heroes.forEach(h => names[h.id] = h.name);
  const log = (s.log || []).slice().reverse().map(l => l.dark
    ? [l.week, 'Тьма', l.steps, '', l.pos, l.caught && l.caught.length ? 'в плен: ' + l.caught.join(', ') : 'никто не пойман']
    : [l.week, names[l.hero] || '?', l.delta, l.from, l.to, (l.parts || []).map(p => p[0] + ' ' + (p[1] > 0 ? '+' : '') + p[1]).join(', ')]);
  writeTable_('Журнал', ['Неделя', 'Кто', 'Ходы', 'Было', 'Стало', 'Из чего сложилось'], log);

  // Сводка
  const sum = [
    ['Неделя', s.week],
    ['Клетка Тьмы', s.dark ? s.dark.pos : ''],
    ['Добавка к ходу Тьмы', s.dark ? s.dark.bonus : ''],
    ['Фонарей у Наставника', s.lanterns],
    ['Яра отгоняла Тьму', s.yaraUsed ? 'да' : 'нет'],
    ['Открытые законы пути', (s.laws || []).join(' | ') || '—'],
    ['Обновлено', new Date()],
  ];
  Object.keys(TRIPLE_NAMES).forEach(k => {
    const t = (s.triples || {})[k] || {};
    sum.push(['Сила тройки ' + TRIPLE_NAMES[k], t.used ? 'применена на неделе ' + t.week : 'не тронута']);
  });
  writeTable_('Сводка', ['Показатель', 'Значение'], sum);

  mirrorProfiles_(s);

  // Личные ссылки участников (SITE_URL — адрес сайта на GitHub Pages)
  const site = PropertiesService.getScriptProperties().getProperty('SITE_URL') || '';
  writeTable_('Ссылки', ['Герой', 'Личная ссылка'],
    s.heroes.map(h => [h.name, site ? site.replace(/\/?$/, '/') + '?h=' + h.code : 'впиши SITE_URL в свойства скрипта, код: ' + h.code]));
}

/* лист «Анкеты»: что герои выбрали при первом входе */
function mirrorProfiles_(s) {
  if (!s || !s.heroes) return;
  const pr = loadProfiles_();
  const rows = s.heroes.map(h => {
    const p = pr[h.id];
    if (!p) return [h.name, '—', '', '', '', '', '', 'анкета не заполнена'];
    const sc = p.scroll || {};
    const ok = h.clsOk && h.tripleOk && h.cls === p.cls && h.triple === p.triple;
    return [p.name || h.name, CLASS_NAMES[p.cls] || '', TRIPLE_NAMES[p.triple] || '', sc.goal || '', sc.train || '', sc.boost || '',
            p.at ? new Date(p.at) : '', ok ? 'подтверждено' : 'ждёт подтверждения'];
  });
  writeTable_('Анкеты', ['Герой', 'Класс', 'Стихия', 'Цель на год', 'Натренировать', 'Усилить', 'Заполнено', 'Статус'], rows);
}

function writeTable_(name, header, rows) {
  const sh = sheet_(name);
  sh.clear();
  const all = [header].concat(rows);
  sh.getRange(1, 1, all.length, header.length).setValues(all);
  sh.getRange(1, 1, 1, header.length)
    .setFontWeight('bold').setBackground('#132150').setFontColor('#e9b84c').setVerticalAlignment('middle');
  sh.setRowHeight(1, 32);
  sh.setFrozenRows(1);
  if (rows.length) {
    sh.getRange(2, 1, rows.length, header.length).setVerticalAlignment('top').setWrap(true);
    const bg = rows.map((_, r) => header.map(() => (r % 2 ? '#f6f1e3' : '#ffffff')));
    sh.getRange(2, 1, rows.length, header.length).setBackgrounds(bg);
  }
  sh.autoResizeColumns(1, header.length);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

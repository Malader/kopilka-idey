/**
 * «Копилка идей»: приём заявок с формы в Google Таблицу,
 * письмо ответственным и рассылка в телеграм подписчикам бота.
 *
 * Куда вставлять: Google Таблица -> Расширения -> Apps Script -> заменить
 * содержимое файла Code.gs на этот текст -> Сохранить -> Развернуть.
 * Подробная инструкция в README.md.
 */

/* =====================  НАСТРОЙКИ  ===================== */

var CONFIG = {
  // Название листа с заявками. Если листа нет, он создастся сам.
  SHEET_NAME: 'Заявки',

  // Название листа со списком подписчиков телеграм-бота.
  SUBS_SHEET_NAME: 'Подписчики',

  // Кому приходит письмо о новой заявке. Можно несколько через запятую.
  NOTIFY_EMAILS: 'v.malyshev@g.nsu.ru, v.malysheva@greenway.group',

  // Отправлять автору идеи письмо «спасибо, заявка принята».
  SEND_CONFIRMATION: true,

  // Как подписывать письма.
  PROJECT_NAME: 'Копилка идей',

  // Минимальное время заполнения формы в секундах.
  // Всё, что отправлено быстрее, считается ботом. 0 выключает проверку.
  MIN_SECONDS: 4,

  // Как часто бот проверяет команды подписчиков, в минутах.
  TELEGRAM_POLL_MINUTES: 1,

  // Сколько последних заявок показывает команда /last.
  TELEGRAM_LAST_COUNT: 5
};

/* Токен бота и служебные значения лежат в свойствах скрипта
   (Настройки проекта -> Свойства скрипта), а не в коде, потому что
   репозиторий публичный:

     TELEGRAM_TOKEN   строка от @BotFather вида 1234567890:AAE...
     TELEGRAM_PAROL   необязательно: кодовое слово для подписки.
                      Если заполнено, подписаться можно только
                      командой «/start слово».
     TELEGRAM_OFFSET  служебное, скрипт ведёт сам, руками не трогать.
*/

/* ============  СООТВЕТСТВИЕ ПОЛЕЙ И КОЛОНОК  ============ */

var FIELDS = [
  ['fullName',      'ФИО'],
  ['email',         'Почта'],
  ['title',         'Название идеи'],
  ['idea',          'Суть идеи'],
  ['problem',       'Какую проблему решает'],
  ['topic',         'Тема'],
  ['benefit',       'Что даст или улучшит'],
  ['metrics',       'Эффект в цифрах'],
  ['resources',     'Нужные ресурсы'],
  ['deadline',      'Срок реализации'],
  ['lead',          'Возможный руководитель'],
  ['participation', 'Готовность участвовать'],
  ['materials',     'Ссылка на материалы'],
  ['consent',       'Согласие на обработку ПД']
];

var REQUIRED = ['fullName', 'email', 'title', 'idea', 'problem', 'topic',
  'benefit', 'resources', 'deadline', 'lead', 'participation'];

/* =====================  ТОЧКИ ВХОДА  ===================== */

function doGet() {
  return ContentService
    .createTextOutput('Копилка идей: сервис приёма заявок работает.')
    .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  try {
    var data = parsePayload_(e);

    // Скрытое поле-ловушка: люди его не видят, боты заполняют.
    if (data.website) return json_({ ok: true, skipped: 'honeypot' });

    // Слишком быстрое заполнение тоже похоже на бота.
    var elapsed = Number(data.elapsed || 0);
    if (CONFIG.MIN_SECONDS > 0 && elapsed > 0 && elapsed < CONFIG.MIN_SECONDS) {
      return json_({ ok: true, skipped: 'too_fast' });
    }

    var missing = missingFields_(data);
    if (missing.length) {
      return json_({ ok: false, error: 'Не заполнены обязательные поля: ' + missing.join(', ') });
    }

    var saved = saveRow_(data);

    // Повтор той же отправки (запасной путь сработал поверх основного):
    // строка уже есть, второй раз писать и слать уведомления не нужно.
    // Уведомления обёрнуты в safely_: заявка уже в таблице, и сбой письма
    // или телеграма не должен возвращать человеку ошибку.
    if (!saved.duplicate) {
      safely_(function () { notifyTelegram_(data, saved.number); });
      safely_(function () { notifyTeam_(data, saved.number); });
      if (CONFIG.SEND_CONFIRMATION) safely_(function () { notifyAuthor_(data, saved.number); });
    }

    return json_({ ok: true, number: saved.number, duplicate: saved.duplicate });
  } catch (err) {
    logError_(err, e);
    return json_({ ok: false, error: 'Внутренняя ошибка сервиса. Попробуйте ещё раз позже.' });
  }
}

/* =====================  РАБОТА С ТАБЛИЦЕЙ  ===================== */

function sheet_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sh) sh = ss.insertSheet(CONFIG.SHEET_NAME);

  if (sh.getLastRow() === 0) {
    var header = ['№', 'Дата и время'];
    FIELDS.forEach(function (f) { header.push(f[1]); });
    header.push('Статус', 'Комментарий');

    sh.appendRow(header);
    sh.getRange(1, 1, 1, header.length)
      .setFontWeight('bold')
      .setBackground('#edf5ff')
      .setVerticalAlignment('middle');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 50);
    sh.setColumnWidth(2, 140);
    for (var c = 3; c <= header.length; c++) sh.setColumnWidth(c, 220);
  }
  return sh;
}

function saveRow_(data) {
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    var cache = CacheService.getScriptCache();
    var key = clean_(data.submissionId) ? 'sub_' + clean_(data.submissionId) : '';

    // ту же отправку второй раз не записываем, возвращаем номер первой
    if (key) {
      var seen = cache.get(key);
      if (seen) return { number: Number(seen), duplicate: true };
    }

    var sh = sheet_();
    var number = sh.getLastRow(); // строка заголовка занимает первую, поэтому это и есть номер заявки

    var row = [number, new Date()];
    FIELDS.forEach(function (f) { row.push(clean_(data[f[0]])); });
    row.push('Новая', '');

    sh.appendRow(row);
    var last = sh.getLastRow();
    sh.getRange(last, 1, 1, row.length).setVerticalAlignment('top').setWrap(true);
    sh.getRange(last, 2).setNumberFormat('dd.MM.yyyy HH:mm');

    if (key) cache.put(key, String(number), 1800);

    return { number: number, duplicate: false };
  } finally {
    lock.releaseLock();
  }
}

/* =====================  ПИСЬМА  ===================== */

function notifyTeam_(data, number) {
  var to = String(CONFIG.NOTIFY_EMAILS || '').trim();
  if (!to) return;

  var rows = FIELDS.map(function (f) {
    var v = clean_(data[f[0]]);
    if (!v) return '';
    return '<tr>' +
      '<td style="padding:8px 14px 8px 0;vertical-align:top;color:#7f8da6;font-size:13px;white-space:nowrap">' + esc_(f[1]) + '</td>' +
      '<td style="padding:8px 0;vertical-align:top;color:#13233f;font-size:14px">' + esc_(v).replace(/\n/g, '<br>') + '</td>' +
      '</tr>';
  }).join('');

  var html =
    '<div style="font-family:Arial,Helvetica,sans-serif;max-width:640px">' +
    '<h2 style="margin:0 0 4px;font-size:20px;color:#13233f">Новая идея №' + number + '</h2>' +
    '<p style="margin:0 0 18px;color:#7f8da6;font-size:13px">' + esc_(CONFIG.PROJECT_NAME) + ', ' +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm') + '</p>' +
    '<table style="border-collapse:collapse;width:100%">' + rows + '</table>' +
    '<p style="margin:22px 0 0"><a href="' + SpreadsheetApp.getActive().getUrl() +
    '" style="display:inline-block;padding:10px 18px;border-radius:6px;background:#1769e8;color:#fff;text-decoration:none;font-size:14px">Открыть таблицу заявок</a></p>' +
    '</div>';

  var options = {
    name: CONFIG.PROJECT_NAME,
    htmlBody: html
  };
  if (isEmail_(data.email)) options.replyTo = String(data.email).trim();

  MailApp.sendEmail(to, 'Новая идея №' + number + ': ' + clean_(data.title), stripTags_(html), options);
}

function notifyAuthor_(data, number) {
  if (!isEmail_(data.email)) return;

  var html =
    '<div style="font-family:Arial,Helvetica,sans-serif;max-width:560px">' +
    '<h2 style="margin:0 0 12px;font-size:20px;color:#13233f">Спасибо, идея принята</h2>' +
    '<p style="margin:0 0 14px;color:#13233f;font-size:15px;line-height:1.5">' +
    'Ваша заявка зарегистрирована под номером <b>' + number + '</b>.</p>' +
    '<p style="margin:0 0 6px;color:#7f8da6;font-size:13px">Название идеи</p>' +
    '<p style="margin:0 0 18px;color:#13233f;font-size:15px"><b>' + esc_(clean_(data.title)) + '</b></p>' +
    '<p style="margin:0;color:#7f8da6;font-size:13px;line-height:1.5">' +
    'Мы рассмотрим идею и свяжемся с вами по этому адресу. Отвечать на это письмо не нужно.</p>' +
    '</div>';

  MailApp.sendEmail(String(data.email).trim(),
    CONFIG.PROJECT_NAME + ': идея №' + number + ' принята',
    stripTags_(html),
    { name: CONFIG.PROJECT_NAME, htmlBody: html });
}

/* =====================  ТЕЛЕГРАМ: ОСНОВА  ===================== */

function tgProp_(name) {
  try {
    return String(PropertiesService.getScriptProperties().getProperty(name) || '').trim();
  } catch (err) {
    return '';
  }
}

// Настоящий токен выглядит как 1234567890:AAE... Пока в свойстве лежит
// заглушка, считаем что телеграм не настроен и никуда не ходим.
function tgToken_() {
  var t = tgProp_('TELEGRAM_TOKEN');
  return /^\d{5,}:[A-Za-z0-9_-]{20,}$/.test(t) ? t : '';
}

function tgCall_(method, payload) {
  var token = tgToken_();
  if (!token) return null;

  var res = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/' + method, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload || {}),
    muteHttpExceptions: true
  });

  try {
    return JSON.parse(res.getContentText());
  } catch (err) {
    return { ok: false, description: res.getContentText().slice(0, 300) };
  }
}

function tgSend_(chatId, text) {
  return tgCall_('sendMessage', {
    chat_id: String(chatId),
    text: text,
    parse_mode: 'HTML',
    disable_web_page_preview: true
  });
}

/* =====================  ТЕЛЕГРАМ: ПОДПИСЧИКИ  ===================== */

function subsSheet_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(CONFIG.SUBS_SHEET_NAME);
  if (!sh) sh = ss.insertSheet(CONFIG.SUBS_SHEET_NAME);

  if (sh.getLastRow() === 0) {
    sh.appendRow(['Chat ID', 'Имя', 'Ник', 'Подписан', 'Статус']);
    sh.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#edf5ff');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 130);
    sh.setColumnWidth(2, 220);
    sh.setColumnWidth(3, 160);
    sh.setColumnWidth(4, 140);
    sh.setColumnWidth(5, 170);
  }
  return sh;
}

function subsAll_() {
  var sh = subsSheet_();
  if (sh.getLastRow() < 2) return [];

  return sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues().map(function (r, i) {
    return { row: i + 2, id: String(r[0]).trim(), name: r[1], nick: r[2], status: String(r[4]).trim() };
  }).filter(function (s) { return s.id; });
}

function subsActive_() {
  return subsAll_().filter(function (s) { return s.status === 'активен'; });
}

function subsFind_(chatId) {
  var id = String(chatId);
  var found = subsAll_().filter(function (s) { return s.id === id; });
  return found.length ? found[0] : null;
}

function subsSetStatus_(row, status) {
  subsSheet_().getRange(row, 5).setValue(status);
}

function subsAdd_(chat) {
  var sh = subsSheet_();
  var name = ((chat.first_name || '') + ' ' + (chat.last_name || '')).trim() || chat.title || 'без имени';
  sh.appendRow([
    String(chat.id),
    name,
    chat.username ? '@' + chat.username : '',
    new Date(),
    'активен'
  ]);
  sh.getRange(sh.getLastRow(), 4).setNumberFormat('dd.MM.yyyy HH:mm');
}

/* =====================  ТЕЛЕГРАМ: РАССЫЛКА ЗАЯВОК  ===================== */

function notifyTelegram_(data, number) {
  if (!tgToken_()) return;

  var subs = subsActive_();
  if (!subs.length) return;

  var text = zayavkaText_(data, number);

  subs.forEach(function (s) {
    var res = tgSend_(s.id, text);
    // 403 это «бот заблокирован» или «чат удалён»: помечаем и больше не дёргаем
    if (res && res.ok === false && (res.error_code === 403 || res.error_code === 400)) {
      subsSetStatus_(s.row, 'недоступен');
    }
  });
}

function zayavkaText_(data, number) {
  var line = function (label, key, limit) {
    var v = clean_(data[key]);
    if (!v) return '';
    if (limit && v.length > limit) v = v.slice(0, limit) + '...';
    return '\n<b>' + esc_(label) + ':</b> ' + esc_(v);
  };

  var text =
    '<b>Новая идея №' + number + '</b>\n' +
    esc_(clean_(data.title)) + '\n' +
    line('Автор', 'fullName') +
    line('Почта', 'email') +
    line('Тема', 'topic') +
    line('Срок', 'deadline') +
    line('Участие', 'participation') +
    line('Руководитель', 'lead') +
    '\n' +
    line('Суть', 'idea', 600) +
    line('Проблема', 'problem', 600) +
    line('Что даст', 'benefit', 600) +
    line('Эффект в цифрах', 'metrics', 300) +
    line('Ресурсы', 'resources', 400) +
    line('Материалы', 'materials', 300) +
    '\n\n<a href="' + SpreadsheetApp.getActive().getUrl() + '">Открыть таблицу заявок</a>';

  if (text.length > 4000) text = text.slice(0, 3900) + '\n\n[сообщение обрезано, подробности в таблице]';
  return text;
}

/* =====================  ТЕЛЕГРАМ: КОМАНДЫ  ===================== */

/**
 * Запускается по расписанию (триггер ставит podklyuchitTelegram).
 * Забирает новые сообщения боту и отвечает на команды.
 */
function obrabotatKomandy() {
  if (!tgToken_()) return;

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return; // предыдущий запуск ещё работает

  try {
    var props = PropertiesService.getScriptProperties();
    var offset = Number(props.getProperty('TELEGRAM_OFFSET') || 0);

    var res = tgCall_('getUpdates', { offset: offset, timeout: 0, allowed_updates: ['message'] });
    if (!res || !res.ok) return;

    var list = res.result || [];
    if (!list.length) return;

    list.forEach(function (u) {
      if (u.update_id >= offset) offset = u.update_id + 1;
      if (u.message) safely_(function () { obrabotatSoobshchenie_(u.message); });
    });

    props.setProperty('TELEGRAM_OFFSET', String(offset));
  } finally {
    lock.releaseLock();
  }
}

function obrabotatSoobshchenie_(m) {
  var chat = m.chat;
  if (!chat || !chat.id) return;

  var text = String(m.text || '').trim();
  var cmd = text.split(/\s+/)[0].toLowerCase().replace(/@.*$/, '');
  var arg = text.slice(text.split(/\s+/)[0].length).trim();

  if (cmd === '/start') return komandaStart_(chat, arg);
  if (cmd === '/stop') return komandaStop_(chat);
  if (cmd === '/help' || cmd === '/помощь') return komandaHelp_(chat);
  if (cmd === '/last' || cmd === '/последние') return komandaLast_(chat);

  tgSend_(chat.id, 'Не понял команду. Напишите /help, чтобы увидеть список.');
}

function komandaStart_(chat, arg) {
  var parol = tgProp_('TELEGRAM_PAROL');
  if (parol && arg !== parol) {
    tgSend_(chat.id,
      'Подписка на этого бота закрыта кодовым словом.\n' +
      'Отправьте команду вместе с ним: <code>/start кодовое-слово</code>');
    return;
  }

  var found = subsFind_(chat.id);

  if (found && found.status === 'активен') {
    tgSend_(chat.id, 'Вы уже подписаны. Новые идеи приходят сюда сразу после отправки формы.');
    return;
  }

  if (found) {
    subsSetStatus_(found.row, 'активен');
  } else {
    subsAdd_(chat);
  }

  tgSend_(chat.id,
    '<b>Подписка оформлена</b>\n' +
    'Каждая новая идея из формы будет приходить сюда: кто предложил, ' +
    'по какой теме, в чём суть и что она даст.\n\n' +
    '/last показать последние заявки\n' +
    '/stop отписаться\n' +
    '/help что умеет бот');
}

function komandaStop_(chat) {
  var found = subsFind_(chat.id);

  if (!found || found.status !== 'активен') {
    tgSend_(chat.id, 'Вы и так не подписаны. Чтобы снова получать идеи, отправьте /start.');
    return;
  }

  subsSetStatus_(found.row, 'отписан');
  tgSend_(chat.id, 'Отписал. Новые идеи приходить не будут. Вернуться можно командой /start.');
}

function komandaHelp_(chat) {
  var sub = subsFind_(chat.id);
  var state = (sub && sub.status === 'активен') ? 'подписаны' : 'не подписаны';

  tgSend_(chat.id,
    '<b>' + esc_(CONFIG.PROJECT_NAME) + '</b>\n' +
    'Бот присылает каждую новую идею, которую сотрудники отправляют через форму.\n\n' +
    '/start подписаться на новые идеи\n' +
    '/stop отписаться\n' +
    '/last последние ' + CONFIG.TELEGRAM_LAST_COUNT + ' заявок\n' +
    '/help это сообщение\n\n' +
    'Сейчас вы <b>' + state + '</b>. Всего подписчиков: ' + subsActive_().length + '.');
}

function komandaLast_(chat) {
  var sh = sheet_();
  var total = sh.getLastRow() - 1;

  if (total < 1) {
    tgSend_(chat.id, 'Заявок пока нет. Как только придёт первая, она сразу прилетит сюда.');
    return;
  }

  var n = Math.min(CONFIG.TELEGRAM_LAST_COUNT, total);
  var values = sh.getRange(sh.getLastRow() - n + 1, 1, n, 2 + FIELDS.length).getValues().reverse();
  var tz = Session.getScriptTimeZone();

  var lines = values.map(function (r) {
    var when = (r[1] instanceof Date) ? Utilities.formatDate(r[1], tz, 'dd.MM HH:mm') : String(r[1]);
    return '<b>№' + r[0] + '</b> ' + esc_(String(r[4])) + '\n' +
      '<i>' + esc_(when) + ', ' + esc_(String(r[2])) + ', ' + esc_(String(r[7])) + '</i>';
  });

  tgSend_(chat.id,
    '<b>Последние заявки</b> (всего ' + total + ')\n\n' + lines.join('\n\n') +
    '\n\n<a href="' + SpreadsheetApp.getActive().getUrl() + '">Открыть таблицу</a>');
}

/* ===== Настройка телеграма, запускается вручную один раз =====

   1. В телеграме напишите @BotFather команду /newbot и придумайте боту имя.
      BotFather пришлёт строку вида 1234567890:AAE... это и есть токен.
   2. Настройки проекта -> Свойства скрипта -> впишите токен в TELEGRAM_TOKEN.
   3. Вернитесь в редактор, выберите функцию podklyuchitTelegram и нажмите
      «Выполнить». Скрипт заведёт лист подписчиков, поставит расписание
      и напишет ссылку на бота.
   4. Ссылку раздайте тем, кто должен получать идеи. Каждый жмёт «Запустить»
      или отправляет /start и попадает в рассылку.
============================================================== */

function podklyuchitTelegram() {
  if (!tgToken_()) {
    console.log('В свойстве TELEGRAM_TOKEN нет настоящего токена. ' +
      'Настройки проекта -> Свойства скрипта -> впишите строку от @BotFather ' +
      'вида 1234567890:AAE... вместо заглушки.');
    return;
  }

  var me = tgCall_('getMe', {});
  if (!me || !me.ok) {
    console.log('Телеграм не принял токен. Ответ: ' + JSON.stringify(me));
    return;
  }

  subsSheet_();
  postavitRaspisanie_();
  obrabotatKomandy(); // разбираем команды, которые уже успели прислать

  var bot = me.result || {};
  console.log(
    'Бот подключён: @' + bot.username + ' («' + (bot.first_name || '') + '»).\n' +
    'Ссылка для подписчиков: https://t.me/' + bot.username + '\n' +
    'Каждый, кто откроет её и нажмёт «Запустить», начнёт получать новые идеи.\n' +
    'Список подписчиков виден на листе «' + CONFIG.SUBS_SHEET_NAME + '».\n' +
    'Сейчас активных подписчиков: ' + subsActive_().length + '.');
}

function otklyuchitTelegram() {
  snyatRaspisanie_();
  console.log('Расписание снято, бот больше не отвечает на команды и не рассылает идеи. ' +
    'Письма продолжают приходить. Список подписчиков остался на листе «' +
    CONFIG.SUBS_SHEET_NAME + '».');
}

function postavitRaspisanie_() {
  snyatRaspisanie_();
  ScriptApp.newTrigger('obrabotatKomandy')
    .timeBased()
    .everyMinutes(CONFIG.TELEGRAM_POLL_MINUTES)
    .create();
}

function snyatRaspisanie_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'obrabotatKomandy') ScriptApp.deleteTrigger(t);
  });
}

/* =====================  ВСПОМОГАТЕЛЬНОЕ  ===================== */

function safely_(fn) {
  try {
    fn();
  } catch (err) {
    logError_(err);
  }
}

function parsePayload_(e) {
  if (e && e.postData && e.postData.contents) {
    var raw = String(e.postData.contents).trim();
    if (raw.charAt(0) === '{') {
      try { return JSON.parse(raw); } catch (ignore) {}
    }
  }
  return (e && e.parameter) ? e.parameter : {};
}

function missingFields_(data) {
  var labels = {};
  FIELDS.forEach(function (f) { labels[f[0]] = f[1]; });

  return REQUIRED.filter(function (key) {
    return !clean_(data[key]);
  }).map(function (key) {
    return labels[key] || key;
  });
}

function clean_(v) {
  if (v === null || v === undefined) return '';
  return String(v).trim().slice(0, 5000);
}

function isEmail_(v) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(String(v || '').trim());
}

function esc_(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function stripTags_(html) {
  return String(html).replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function logError_(err, e) {
  try {
    console.error(err && err.stack ? err.stack : String(err));
    if (e && e.postData) console.error('payload: ' + String(e.postData.contents).slice(0, 2000));
  } catch (ignore) {}
}

/* =====================  ПРОВЕРКА ВРУЧНУЮ  =====================
   Запустите эту функцию один раз из редактора Apps Script
   (выбрать testSubmit в списке функций и нажать «Выполнить»),
   чтобы выдать разрешения и убедиться, что строка и письмо уходят.
============================================================== */

function testSubmit() {
  var fake = {
    postData: {
      contents: JSON.stringify({
        fullName: 'Иванов Иван Иванович',
        email: 'test@example.ru',
        title: 'Тестовая идея',
        idea: 'Проверяем, что форма доезжает до таблицы.',
        problem: 'Пока непонятно, работает ли связка.',
        topic: 'Оптимизация процессов',
        benefit: 'Убедимся, что всё настроено верно.',
        metrics: 'экономия 1 час на проверку',
        resources: 'ничего не нужно',
        deadline: 'До 1 месяца',
        lead: 'Отдел развития',
        participation: 'Готов(а) консультировать',
        materials: '',
        consent: 'Да',
        elapsed: 30,
        submissionId: 'test-' + new Date().getTime()
      })
    }
  };
  var res = doPost(fake);
  console.log(res.getContent());
}

/**
 * «Копилка идей»: приём заявок с формы в Google Таблицу + письмо ответственным.
 *
 * Куда вставлять: Google Таблица -> Расширения -> Apps Script -> заменить
 * содержимое файла Code.gs на этот текст -> Сохранить -> Развернуть.
 * Подробная инструкция в README.md.
 */

/* =====================  НАСТРОЙКИ  ===================== */

var CONFIG = {
  // Название листа в таблице. Если листа нет, он создастся сам.
  SHEET_NAME: 'Заявки',

  // Кому приходит письмо о новой заявке. Можно несколько через запятую.
  NOTIFY_EMAILS: 'v.malyshev@g.nsu.ru',

  // Отправлять автору идеи письмо «спасибо, заявка принята».
  SEND_CONFIRMATION: true,

  // Как подписывать письма.
  PROJECT_NAME: 'Копилка идей',

  // Минимальное время заполнения формы в секундах.
  // Всё, что отправлено быстрее, считается ботом. 0 выключает проверку.
  MIN_SECONDS: 4

  // Телеграм настраивается не здесь, а в свойствах скрипта
  // (Настройки проекта -> Свойства скрипта): TELEGRAM_TOKEN и TELEGRAM_CHAT_ID.
  // Токен бота не место в коде, который лежит в публичном репозитории.
  // Как настроить: впишите токен в TELEGRAM_TOKEN, напишите боту любое
  // сообщение и запустите функцию podklyuchitTelegram (см. в конце файла).
};

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
    // Сами уведомления обёрнуты в safely_: заявка уже в таблице, и сбой
    // письма или телеграма не должен возвращать человеку ошибку.
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

/* =====================  ТЕЛЕГРАМ  ===================== */

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
  return JSON.parse(res.getContentText());
}

function notifyTelegram_(data, number) {
  if (!tgToken_() || !tgProp_('TELEGRAM_CHAT_ID')) return; // не настроен, молча пропускаем

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

  tgCall_('sendMessage', {
    chat_id: tgProp_('TELEGRAM_CHAT_ID'),
    text: text,
    parse_mode: 'HTML',
    disable_web_page_preview: true
  });
}

/* ===== Настройка телеграма, запускается вручную один раз =====

   1. В телеграме напишите @BotFather команду /newbot и придумайте боту имя.
      BotFather пришлёт строку вида 1234567890:AAE... это и есть токен.
   2. Настройки проекта -> Свойства скрипта -> впишите токен в TELEGRAM_TOKEN.
   3. Напишите своему боту любое сообщение (или добавьте его в группу
      и напишите там, тогда заявки будут падать в группу).
   4. Выберите в списке функций podklyuchitTelegram и нажмите «Выполнить».
      Скрипт сам найдёт чат, запомнит его и пришлёт туда проверочное сообщение.
============================================================== */

function podklyuchitTelegram() {
  if (!tgToken_()) {
    console.log('В свойстве TELEGRAM_TOKEN нет настоящего токена. ' +
      'Настройки проекта -> Свойства скрипта -> впишите строку от @BotFather ' +
      'вида 1234567890:AAE... вместо заглушки.');
    return;
  }

  var upd = tgCall_('getUpdates', {});
  if (!upd || !upd.ok) {
    console.log('Телеграм ответил ошибкой. Проверьте токен. Ответ: ' + JSON.stringify(upd));
    return;
  }

  var list = upd.result || [];
  if (!list.length) {
    console.log('Телеграм не видит ни одного сообщения боту. Напишите боту любое ' +
      'сообщение (в группе тоже подойдёт) и запустите функцию ещё раз.');
    return;
  }

  var last = list[list.length - 1];
  var src = last.message || last.channel_post || last.edited_message || last.my_chat_member || {};
  var chat = src.chat;
  if (!chat || !chat.id) {
    console.log('Не удалось определить чат. Напишите боту обычное текстовое сообщение и повторите.');
    return;
  }

  PropertiesService.getScriptProperties().setProperty('TELEGRAM_CHAT_ID', String(chat.id));

  var name = chat.title || chat.username || ((chat.first_name || '') + ' ' + (chat.last_name || '')).trim();
  var sent = tgCall_('sendMessage', {
    chat_id: String(chat.id),
    text: 'Копилка идей подключена. Новые заявки будут приходить сюда.'
  });

  if (sent && sent.ok) {
    console.log('Готово. Заявки будут приходить в чат «' + name + '» (id ' + chat.id + ').');
  } else {
    console.log('Чат запомнен (id ' + chat.id + '), но проверочное сообщение не ушло: ' + JSON.stringify(sent));
  }
}

function otklyuchitTelegram() {
  PropertiesService.getScriptProperties().deleteProperty('TELEGRAM_CHAT_ID');
  console.log('Отправка в телеграм выключена. Письма продолжают приходить.');
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

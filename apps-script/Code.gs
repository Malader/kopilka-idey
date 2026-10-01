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
  TELEGRAM_LAST_COUNT: 5,

  MAX_ATTACHMENTS: 5,
  MAX_ATTACHMENT_BYTES: 10 * 1024 * 1024
};

/* Токен бота и служебные значения лежат в свойствах скрипта
   (Настройки проекта -> Свойства скрипта), а не в коде, потому что
   репозиторий публичный:

     TELEGRAM_TOKEN   строка от @BotFather вида 1234567890:AAE...
     TELEGRAM_PAROL   необязательно: кодовое слово для подписки.
                      Если заполнено, подписаться можно только
                      командой «/start слово».
     TELEGRAM_OFFSET  служебное, скрипт ведёт сам, руками не трогать.
     ATTACHMENTS_FOLDER_ID  ID закрытой папки Google Drive для файлов.
     SECURITY_TELEGRAM_ALLOWED_CHAT_IDS  разрешённые chat ID через запятую.
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

function doGet(e) {
  if (e && e.parameter && Object.prototype.hasOwnProperty.call(e.parameter, 'receipt')) {
    var id = e.parameter.receipt;
    if (typeof id !== 'string' || (id.length !== 32 && id.length !== 36) ||
        !/^(?:[a-f0-9]{32}|[a-f0-9]{8}-[a-f0-9]{4}-4[a-f0-9]{3}-[89ab][a-f0-9]{3}-[a-f0-9]{12})$/i.test(id)) {
      return json_({ ok: false, error: 'Некорректный идентификатор отправки.' });
    }

    try {
      var seen = CacheService.getScriptCache().get('sub_' + id);
      if (!seen) return json_({ ok: false, pending: true });
      var receipt = JSON.parse(seen);
      if (typeof receipt === 'number') receipt = { number: receipt, attachmentCount: 0 };
      if (!receipt || !Number.isSafeInteger(receipt.number) || receipt.number <= 0 ||
          !Number.isSafeInteger(receipt.attachmentCount) || receipt.attachmentCount < 0 ||
          receipt.attachmentCount > CONFIG.MAX_ATTACHMENTS) {
        return json_({ ok: false, pending: true });
      }
      return json_({ ok: true, number: receipt.number, attachmentCount: receipt.attachmentCount });
    } catch (err) {
      return json_({ ok: false, pending: true });
    }
  }

  return ContentService
    .createTextOutput('Копилка идей: сервис приёма заявок работает.')
    .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  try {
    var data = parseSecurePayload_(e);
    // Квитанция уже сохранённой заявки доступна без нового списания лимитов.
    var existing = cachedReceipt_(data.submissionId);
    if (existing) return json_({ ok: true, number: existing.number,
      attachmentCount: existing.attachmentCount, duplicate: true });
    admitRequest_(data.clientId);
    validateSecurePayload_(data);
    if (data.website || data.elapsed < CONFIG.MIN_SECONDS) {
      throw securityError_('INVALID_PAYLOAD', 'Обновите страницу формы и заполните её ещё раз.');
    }
    var saved = saveRow_(data);
    if (!saved.duplicate) {
      safely_(function () { notifyTelegram_(data, saved.number); });
      safely_(function () { notifyTeam_(data, saved.number, saved.attachments); });
      if (CONFIG.SEND_CONFIRMATION) safely_(function () { notifyAuthor_(data, saved.number); });
    }
    return json_({ ok: true, number: saved.number, duplicate: saved.duplicate,
      attachmentCount: saved.attachmentCount });
  } catch (err) {
    logError_(err, e);
    var result = { ok: false, error: err && (err.securityCode || err.attachmentError) ? err.message :
      'Внутренняя ошибка сервиса. Данные остались в форме. Попробуйте позже.',
      code: err && err.securityCode || 'SERVICE_ERROR' };
    if (err && err.retryAfterSeconds) result.retryAfterSeconds = err.retryAfterSeconds;
    return json_(result);
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
  if (!lock.tryLock(200)) throw securityError_('RATE_LIMITED', 'Сервис занят. Попробуйте через несколько секунд.', 5);
  try {
    var cache = CacheService.getScriptCache();
    var key = 'sub_' + data.submissionId;
    var existing = cachedReceipt_(data.submissionId);
    if (existing) return { number: existing.number, duplicate: true, attachmentCount: existing.attachmentCount };
    var reservation = reserveDailyBudget_(attachmentBytes_(data.attachments));
    var attachments = [];
    var files = [];
    var sh, number, row;
    var writeAttempted = false;
    var storageUncertain = false;
    try {
      attachments = attachmentBlobs_(data.attachments);
      if (attachments.length) {
        var folder = attachmentsFolder_();
        attachments.forEach(function (blob) {
          storageUncertain = true;
          files.push(folder.createFile(blob));
          storageUncertain = false;
        });
      }
      sh = sheet_();
      number = sh.getLastRow();
      row = [number, new Date()];
      FIELDS.forEach(function (f) { row.push(sheetText_(data[f[0]])); });
      row.push('Новая', '');
      if (files.length) {
        var column = attachmentColumn_(sh);
        while (row.length < column) row.push('');
        row[column - 1] = files.map(function (file, i) {
          return attachments[i].getName() + ': ' + file.getUrl();
        }).join('\n');
      }
      writeAttempted = true;
      sh.appendRow(row);
      var last = sh.getLastRow();
      sh.getRange(last, 1, 1, row.length).setVerticalAlignment('top').setWrap(true);
      sh.getRange(last, 2).setNumberFormat('dd.MM.yyyy HH:mm');
      SpreadsheetApp.flush();
      cache.put(key, JSON.stringify({ number: number, attachmentCount: attachments.length }), 1800);
    } catch (err) {
      var clean = !writeAttempted && !storageUncertain;
      if (attachments.length) {
        try {
          if (writeAttempted && sh && row) {
            if (sh.getLastRow() === number) clean = !storageUncertain;
            else if (sh.getLastRow() === number + 1) {
              var savedRow = sh.getRange(number + 1, 1, 1, 2).getValues()[0];
              if (savedRow[0] === number && savedRow[1] instanceof Date && savedRow[1].getTime() === row[1].getTime()) {
                sh.deleteRow(number + 1);
                SpreadsheetApp.flush();
                clean = !storageUncertain;
              }
            }
          }
        } catch (cleanupError) { clean = false; logError_(cleanupError); }
        files.forEach(function (file) {
          try { file.setTrashed(true); } catch (cleanupError) { clean = false; logError_(cleanupError); }
        });
      }
      // Не возвращаем бюджет, если сохранение или очистка не подтверждены.
      if (clean) safely_(function () { refundDailyBudget_(reservation); });
      if (!attachments.length || err && (err.securityCode || err.attachmentError)) throw err;
      throw attachmentError_('Не удалось сохранить заявку с файлами. Попробуйте ещё раз позже.');
    }
    return { number: number, duplicate: false, attachments: attachments, attachmentCount: attachments.length };
  } finally { lock.releaseLock(); }
}

function attachmentError_(message) {
  var err = new Error(message);
  err.attachmentError = true;
  return err;
}

function attachmentBlobs_(items) {
  if (items === undefined) return [];
  if (!Array.isArray(items)) throw attachmentError_('Некорректный список прикреплённых файлов.');
  if (items.length > CONFIG.MAX_ATTACHMENTS) {
    throw attachmentError_('Можно прикрепить не более ' + CONFIG.MAX_ATTACHMENTS + ' файлов.');
  }

  var mimeTypes = {
    pdf: 'application/pdf',
    docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    txt: 'text/plain', csv: 'text/csv', jpg: 'image/jpeg', jpeg: 'image/jpeg',
    png: 'image/png', webp: 'image/webp', gif: 'image/gif'
  };
  var total = 0;
  return items.map(function (item) {
    if (!item || typeof item !== 'object' || Array.isArray(item)) {
      throw attachmentError_('Некорректные данные прикреплённого файла.');
    }
    if (typeof item.name !== 'string' || !item.name.trim()) {
      throw attachmentError_('Не указано имя прикреплённого файла.');
    }
    var name = item.name.replace(/[\u0000-\u001f\u007f]/g, '')
      .split(/[\\/]/).pop().replace(/[<>:"|?*]/g, '_').replace(/^[=+@-]+/, '').trim();
    var dot = name.lastIndexOf('.');
    var extension = name.slice(dot + 1).toLowerCase();
    if (dot < 1 || !Object.prototype.hasOwnProperty.call(mimeTypes, extension)) {
      throw attachmentError_('Недопустимый формат файла: ' + name.slice(0, 180) + '.');
    }
    name = name.slice(0, Math.min(dot, 180 - (name.length - dot))) + name.slice(dot);
    if (typeof item.size !== 'number' || !isFinite(item.size) ||
        item.size <= 0 || Math.floor(item.size) !== item.size) {
      throw attachmentError_('Некорректный размер файла: ' + name + '. Пустые файлы не принимаются.');
    }
    if (typeof item.base64 !== 'string' || !item.base64.length || item.base64.length % 4 !== 0 ||
        !/^[A-Za-z0-9+/]*={0,2}$/.test(item.base64)) {
      throw attachmentError_('Некорректные данные base64 файла: ' + name + '.');
    }
    var padding = item.base64.slice(-2) === '==' ? 2 : (item.base64.slice(-1) === '=' ? 1 : 0);
    var size = item.base64.length / 4 * 3 - padding;
    if (size !== item.size) throw attachmentError_('Размер файла не совпадает с данными: ' + name + '.');
    total += size;
    if (total > CONFIG.MAX_ATTACHMENT_BYTES) {
      throw attachmentError_('Общий размер файлов не должен превышать 10 МБ.');
    }
    var bytes;
    try {
      bytes = Utilities.base64Decode(item.base64);
    } catch (err) {
      throw attachmentError_('Некорректные данные base64 файла: ' + name + '.');
    }
    if (bytes.length !== item.size || Utilities.base64Encode(bytes) !== item.base64) {
      throw attachmentError_('Некорректные данные base64 файла: ' + name + '.');
    }
    return Utilities.newBlob(bytes, mimeTypes[extension], name);
  });
}

function attachmentsFolder_() {
  var id;
  try {
    id = String(PropertiesService.getScriptProperties().getProperty('ATTACHMENTS_FOLDER_ID') || '').trim();
  } catch (err) {
    throw attachmentError_('Не удалось прочитать настройки ATTACHMENTS_FOLDER_ID. Попробуйте ещё раз позже.');
  }
  if (!id) {
    throw attachmentError_('Приём файлов не настроен: укажите ATTACHMENTS_FOLDER_ID в свойствах скрипта.');
  }
  var folder;
  try {
    folder = DriveApp.getFolderById(id);
    if (folder.isTrashed()) throw attachmentError_('Папка для файлов находится в корзине.');
    if (folder.getSharingAccess() !== DriveApp.Access.PRIVATE) {
      throw attachmentError_('Папка для файлов должна быть закрыта: отключите общий доступ по ссылке и для домена.');
    }
  } catch (err) {
    if (err && err.attachmentError) throw err;
    throw attachmentError_('Нет доступа к папке для файлов. Проверьте ATTACHMENTS_FOLDER_ID и разрешения Google Drive.');
  }
  return folder;
}

function attachmentColumn_(sh) {
  var last = sh.getLastColumn();
  var header = sh.getRange(1, 1, 1, last).getValues()[0];
  var index = header.indexOf('Файлы');
  if (index >= 0) return index + 1;
  sh.getRange(1, last + 1).setValue('Файлы').setFontWeight('bold')
    .setBackground('#edf5ff').setVerticalAlignment('middle');
  sh.setColumnWidth(last + 1, 220);
  return last + 1;
}

/* =====================  ПИСЬМА  ===================== */

function notifyTeam_(data, number, attachments) {
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
  if (attachments && attachments.length) options.attachments = attachments;

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
  return subsAll_().filter(function (s) { return s.status === 'активен' && botRecipientAllowed_(s.id); });
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
    sheetText_(name),
    sheetText_(chat.username ? '@' + chat.username : ''),
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
  if (!botRecipientAllowed_(chat.id)) {
    tgSend_(chat.id, 'Доступ к заявкам закрыт. Обратитесь к организатору.');
    return;
  }
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
    'Сейчас вы <b>' + state + '</b>. Доступ к заявкам выдаёт организатор.');
}

function komandaLast_(chat) {
  if (!botRecipientAllowed_(chat.id)) {
    tgSend_(chat.id, 'Доступ к заявкам закрыт. Обратитесь к организатору.');
    return;
  }
  var subscriber = subsFind_(chat.id);
  if (!subscriber || subscriber.status !== 'активен') {
    tgSend_(chat.id, 'Для просмотра заявок оформите разрешённую подписку командой /start.');
    return;
  }
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
   4. В SECURITY_TELEGRAM_ALLOWED_CHAT_IDS внесите проверенные chat ID.
      Только разрешённые чаты смогут подписаться и просматривать заявки.
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

function logError_(err, e) { safeLogError_(err, e); }
/* Ручная отправка тестов запрещена: проверяйте через ?preview=1 и локальные тесты. */

var SECURITY = {
  POST_BYTES: 15 * 1024 * 1024,
  REQUESTS_PER_MINUTE: 20,
  CLIENT_REQUESTS_PER_HOUR: 5,
  CLIENT_INTERVAL_MS: 60000,
  SUBMISSIONS_PER_DAY: 100,
  FILE_BYTES_PER_DAY: 200 * 1024 * 1024
};

function securityError_(code, message, retryAfter) {
  var err = new Error(message);
  err.securityCode = code;
  if (retryAfter) err.retryAfterSeconds = retryAfter;
  return err;
}

function secureId_(value) {
  return typeof value === 'string' && (value.length === 32 || value.length === 36) &&
    /^(?:[a-f0-9]{32}|[a-f0-9]{8}-[a-f0-9]{4}-4[a-f0-9]{3}-[89ab][a-f0-9]{3}-[a-f0-9]{12})$/i.test(value);
}

function cachedReceipt_(id) {
  if (!secureId_(id)) return null;
  var raw = CacheService.getScriptCache().get('sub_' + id);
  if (!raw) return null;
  var receipt = JSON.parse(raw);
  if (typeof receipt === 'number') receipt = { number: receipt, attachmentCount: 0 };
  if (!receipt || !Number.isSafeInteger(receipt.number) || receipt.number < 1 ||
      !Number.isSafeInteger(receipt.attachmentCount) || receipt.attachmentCount < 0 ||
      receipt.attachmentCount > CONFIG.MAX_ATTACHMENTS) return null;
  return receipt;
}

function parseSecurePayload_(e) {
  if (!e || !e.postData || typeof e.postData.contents !== 'string') {
    throw securityError_('INVALID_PAYLOAD', 'Некорректный запрос формы. Обновите страницу.');
  }
  var raw = e.postData.contents;
  if (raw.length > SECURITY.POST_BYTES || Number(e.contentLength || e.postData.length || 0) > SECURITY.POST_BYTES) {
    throw securityError_('PAYLOAD_TOO_LARGE', 'Размер запроса превышает допустимый предел.');
  }
  if (Utilities.newBlob(raw).getBytes().length > SECURITY.POST_BYTES) {
    throw securityError_('PAYLOAD_TOO_LARGE', 'Размер запроса превышает допустимый предел.');
  }
  var data;
  try { data = JSON.parse(raw); } catch (err) {
    throw securityError_('INVALID_PAYLOAD', 'Некорректный запрос формы. Обновите страницу.');
  }
  if (!data || typeof data !== 'object' || Array.isArray(data)) {
    throw securityError_('INVALID_PAYLOAD', 'Некорректный запрос формы. Обновите страницу.');
  }
  return data;
}

function validateSecurePayload_(data) {
  var limits = { fullName: 200, email: 254, title: 150, idea: 4000, problem: 4000,
    topic: 80, benefit: 4000, metrics: 300, resources: 4000, deadline: 80,
    lead: 200, participation: 100, materials: 500, consent: 3 };
  Object.keys(limits).forEach(function (key) {
    var value = data[key];
    if (value === undefined && REQUIRED.indexOf(key) === -1 && key !== 'consent') return;
    if (typeof value !== 'string' || value.length > limits[key] || /[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/.test(value)) {
      throw securityError_('INVALID_FIELD', 'Некорректное поле формы: ' + key + '.');
    }
    if (REQUIRED.indexOf(key) !== -1 && !value.trim()) {
      throw securityError_('INVALID_FIELD', 'Заполните обязательное поле: ' + key + '.');
    }
  });
  if (!isEmail_(data.email) || data.consent !== 'Да' || !secureId_(data.submissionId) || !secureId_(data.clientId)) {
    throw securityError_('INVALID_FIELD', 'Проверьте почту и согласие. Если ошибка повторяется, обновите страницу.');
  }
  var choices = {
    topic: ['Оптимизация процессов', 'Сокращение затрат', 'Рост ТО', 'Идеи про продукты', 'Прочее'],
    deadline: ['До 1 месяца', '1–3 месяца', '3–6 месяцев', 'Более 6 месяцев', 'Пока сложно оценить'],
    participation: ['Да, готов(а) стать руководителем', 'Да, готов(а) участвовать в команде', 'Готов(а) консультировать', 'Нет, хочу только предложить идею']
  };
  Object.keys(choices).forEach(function (key) {
    if (choices[key].indexOf(data[key]) === -1) throw securityError_('INVALID_FIELD', 'Выберите допустимый вариант: ' + key + '.');
  });
  if (data.materials && !/^https?:\/\/[^\s]+$/i.test(data.materials)) {
    throw securityError_('INVALID_FIELD', 'Ссылка на материалы должна начинаться с https:// или http://.');
  }
  if (!Number.isSafeInteger(data.elapsed) || data.elapsed < 0 || data.elapsed > 86400 ||
      (data.website !== undefined && (typeof data.website !== 'string' || data.website.length > 200))) {
    throw securityError_('INVALID_FIELD', 'Обновите страницу формы и попробуйте снова.');
  }
  var allowedKeys = Object.keys(limits).concat(['submissionId', 'clientId', 'attachments', 'website', 'elapsed', 'page']);
  if (Object.keys(data).some(function (key) { return allowedKeys.indexOf(key) === -1; })) {
    throw securityError_('INVALID_PAYLOAD', 'Запрос содержит неизвестные поля. Обновите страницу формы.');
  }
  if (data.page !== undefined && (typeof data.page !== 'string' || data.page.length > 2000)) {
    throw securityError_('INVALID_FIELD', 'Некорректный адрес страницы формы.');
  }
  attachmentBytes_(data.attachments);
}

function attachmentBytes_(items) {
  if (items === undefined) return 0;
  if (!Array.isArray(items) || items.length > CONFIG.MAX_ATTACHMENTS) {
    throw securityError_('INVALID_FILES', 'Можно прикрепить не более пяти файлов.');
  }
  var total = 0;
  items.forEach(function (item) {
    if (!item || typeof item !== 'object' || !Number.isSafeInteger(item.size) || item.size < 1 ||
        typeof item.name !== 'string' || item.name.length > 255 || typeof item.base64 !== 'string' ||
        !item.base64.length || item.base64.length % 4 !== 0 || !/^[A-Za-z0-9+/]*={0,2}$/.test(item.base64)) {
      throw securityError_('INVALID_FILES', 'Некорректные данные прикреплённого файла.');
    }
    var padding = item.base64.slice(-2) === '==' ? 2 : item.base64.slice(-1) === '=' ? 1 : 0;
    if (item.base64.length / 4 * 3 - padding !== item.size ||
        !/\.(pdf|docx|xlsx|pptx|txt|csv|jpg|jpeg|png|webp|gif)$/i.test(item.name)) {
      throw securityError_('INVALID_FILES', 'Проверьте размер и формат прикреплённого файла.');
    }
    total += item.size;
    if (total > CONFIG.MAX_ATTACHMENT_BYTES) throw securityError_('INVALID_FILES', 'Общий размер файлов не должен превышать 10 МБ.');
  });
  return total;
}

function readSecurityProperty_(props, key, fallback) {
  var raw = props.getProperty(key);
  if (!raw) return fallback;
  try { return JSON.parse(raw); } catch (err) {
    throw securityError_('PROTECTION_UNAVAILABLE', 'Защита сервиса временно недоступна. Попробуйте позже.');
  }
}

function clientHash_(id) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, id, Utilities.Charset.UTF_8)
    .map(function (byte) { return ('0' + ((byte + 256) % 256).toString(16)).slice(-2); }).join('');
}

function admitRequest_(clientId) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(200)) throw securityError_('RATE_LIMITED', 'Сервис занят. Попробуйте через несколько секунд.', 5);
  try {
    var now = Date.now();
    var props = PropertiesService.getScriptProperties();
    var globalTimes = readSecurityProperty_(props, 'SECURITY_RATE_WINDOW', []);
    if (!Array.isArray(globalTimes) || globalTimes.some(function (t) { return !Number.isSafeInteger(t) || t > now; })) {
      throw securityError_('PROTECTION_UNAVAILABLE', 'Защита сервиса временно недоступна. Попробуйте позже.');
    }
    globalTimes = globalTimes.filter(function (t) { return t > now - 60000; });
    if (globalTimes.length >= SECURITY.REQUESTS_PER_MINUTE) {
      throw securityError_('RATE_LIMITED', 'Слишком много отправок. Попробуйте через минуту.', Math.max(1, Math.ceil((globalTimes[0] + 60000 - now) / 1000)));
    }
    // Кэш браузерных ограничений отделён от ScriptCache с квитанциями.
    globalTimes.push(now);
    props.setProperty('SECURITY_RATE_WINDOW', JSON.stringify(globalTimes));
    if (!secureId_(clientId)) return;
    var cache = CacheService.getUserCache();
    var key = 'security_client_' + clientHash_(clientId);
    var times = [];
    try { times = JSON.parse(cache.get(key) || '[]'); } catch (err) {}
    if (!Array.isArray(times)) times = [];
    times = times.filter(function (t) { return Number.isSafeInteger(t) && t <= now && t > now - 3600000; });
    var wait = times.length && times[times.length - 1] + SECURITY.CLIENT_INTERVAL_MS - now;
    if (wait > 0 || times.length >= SECURITY.CLIENT_REQUESTS_PER_HOUR) {
      var retry = wait > 0 ? wait : times[0] + 3600000 - now;
      throw securityError_('RATE_LIMITED', 'Слишком частые отправки из этого браузера. Попробуйте позже.', Math.max(1, Math.ceil(retry / 1000)));
    }
    times.push(now);
    cache.put(key, JSON.stringify(times), 3600);
  } catch (err) {
    if (err.securityCode) throw err;
    throw securityError_('PROTECTION_UNAVAILABLE', 'Защита сервиса временно недоступна. Попробуйте позже.');
  } finally { lock.releaseLock(); }
}

function dailyState_(props) {
  var day = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var state = readSecurityProperty_(props, 'SECURITY_DAILY_BUDGET', { day: day, submissions: 0, bytes: 0 });
  if (!state || typeof state.day !== 'string' || !Number.isSafeInteger(state.submissions) || state.submissions < 0 ||
      !Number.isSafeInteger(state.bytes) || state.bytes < 0) {
    throw securityError_('PROTECTION_UNAVAILABLE', 'Защита сервиса временно недоступна. Попробуйте позже.');
  }
  return state.day === day ? state : { day: day, submissions: 0, bytes: 0 };
}

function reserveDailyBudget_(bytes) {
  var props = PropertiesService.getScriptProperties();
  var state = dailyState_(props);
  if (state.submissions >= SECURITY.SUBMISSIONS_PER_DAY || state.bytes + bytes > SECURITY.FILE_BYTES_PER_DAY) {
    throw securityError_('DAILY_LIMIT', 'Достигнут дневной лимит сервиса. Попробуйте завтра или свяжитесь с организатором.');
  }
  state.submissions++;
  state.bytes += bytes;
  props.setProperty('SECURITY_DAILY_BUDGET', JSON.stringify(state));
  return { day: state.day, bytes: bytes };
}

function refundDailyBudget_(reservation) {
  if (!reservation) return;
  var props = PropertiesService.getScriptProperties();
  var state = dailyState_(props);
  if (state.day !== reservation.day) return;
  state.submissions = Math.max(0, state.submissions - 1);
  state.bytes = Math.max(0, state.bytes - reservation.bytes);
  props.setProperty('SECURITY_DAILY_BUDGET', JSON.stringify(state));
}

function sheetText_(value) {
  var text = clean_(value);
  return /^[=+@-]/.test(text) ? "'" + text : text;
}

function botRecipientAllowed_(id) {
  var allowed = tgProp_('SECURITY_TELEGRAM_ALLOWED_CHAT_IDS').split(',').map(function (value) { return value.trim(); }).filter(Boolean);
  return allowed.indexOf(String(id)) !== -1;
}

function safeLogError_(err, e) {
  try {
    console.error(JSON.stringify({ code: err && (err.securityCode || err.name) || 'ServiceError',
      time: new Date().toISOString(), requestBytes: e && Number(e.contentLength) || undefined }));
  } catch (ignore) {}
}

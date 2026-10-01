const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const test = require('node:test');

const source = fs.readFileSync(path.join(__dirname, '../apps-script/Code.gs'), 'utf8');
const header = ['№', 'Дата и время', 'ФИО', 'Почта', 'Название идеи', 'Суть идеи', 'Какую проблему решает', 'Тема', 'Что даст или улучшит', 'Эффект в цифрах', 'Нужные ресурсы', 'Срок реализации', 'Возможный руководитель', 'Готовность участвовать', 'Ссылка на материалы', 'Согласие на обработку ПД', 'Статус', 'Комментарий'];

function payload(extra = {}) {
  return { fullName: 'Иван Иванов', email: 'test@example.ru', title: 'Идея', idea: 'Суть', problem: 'Проблема', topic: 'Оптимизация процессов', benefit: 'Польза', resources: 'Нет', deadline: 'До 1 месяца', lead: 'Отдел', participation: 'Готов(а) консультировать', consent: 'Да', elapsed: 30, submissionId: '7fbf0eaf1d1e4a57a4059683e62eab76', clientId: '7fbf0eaf1d1e4a57a4059683e62eab76', ...extra };
}

function attachment(name = 'idea.pdf', contents = 'abc') {
  const bytes = Buffer.isBuffer(contents) ? contents : Buffer.from(contents);
  return { name, mimeType: 'untrusted/client', size: bytes.length, base64: bytes.toString('base64') };
}

function harness(options = {}) {
  const state = { rows: options.rows ? options.rows.map(row => [...row]) : [], files: [], mails: [], telegram: [], logs: [], locked: false, folderReads: 0, createAttempts: 0, lockRequests: 0, cacheWrites: 0, flushes: 0, captcha: [], propsWrites: [], sheetReads: 0, now: Date.parse('2026-10-01T05:00:00Z') };
  const cache = new Map(options.cacheEntries || []);
  const clientCache = new Map();
  const props = { ATTACHMENTS_FOLDER_ID: 'private-folder', SECURITY_TELEGRAM_ALLOWED_CHAT_IDS: '123', ...options.props };
  function range(row, col, count = 1, width = 1) {
    const result = {};
    for (const method of ['setFontWeight', 'setBackground', 'setVerticalAlignment', 'setWrap', 'setNumberFormat']) result[method] = () => {
      if (options.formatFails && row > 1) throw new Error('range formatting failure');
      return result;
    };
    result.getValues = () => Array.from({ length: count }, (_, i) => Array.from({ length: width }, (_, j) => state.rows[row - 1 + i]?.[col - 1 + j] ?? ''));
    result.setValue = value => { while (state.rows.length < row) state.rows.push([]); state.rows[row - 1][col - 1] = value; return result; };
    result.setValues = values => { values.forEach((valuesRow, i) => valuesRow.forEach((value, j) => range(row + i, col + j).setValue(value))); return result; };
    return result;
  }
  const sheet = {
    getLastRow: () => state.rows.length,
    getLastColumn: () => Math.max(0, ...state.rows.map(row => row.length)),
    getRange: range,
    appendRow(row) {
      if (options.appendFails && state.rows.length) throw new Error('sheet failure');
      state.rows.push([...row]);
      if (options.appendWritesThenFails && state.rows.length > 1) throw new Error('sheet failure after write');
    },
    deleteRow: row => state.rows.splice(row - 1, 1),
    setFrozenRows() {}, setColumnWidth() {}
  };
  const folder = {
    getSharingAccess: () => options.folderAccess || 'PRIVATE',
    isTrashed: () => Boolean(options.folderTrashed),
    createFile(blob) {
      assert.equal(state.locked, true, 'storage must run under existing submission lock');
      state.createAttempts++;
      if (state.createAttempts === options.failCreateAt) throw new Error('Drive unavailable');
      const file = { blob, trashed: false, getUrl: () => 'https://drive.google.com/file/d/file-' + (state.files.indexOf(file) + 1) + '/view', setTrashed(value) { if (options.trashFails) throw new Error('cleanup failure'); this.trashed = value; }, setSharing() { throw new Error('must never grant access'); } };
      state.files.push(file);
      return file;
    }
  };
  class ClockDate extends Date {
    constructor(...args) { super(...(args.length ? args : [state.now])); }
    static now() { return state.now; }
  }
  const context = vm.createContext({
    Date: ClockDate,
    console: { log() {}, error: (...args) => state.logs.push(args.join(' ')) },
    ContentService: { MimeType: { JSON: 'JSON', TEXT: 'TEXT' }, createTextOutput: text => ({ mimeType: null, setMimeType(value) { this.mimeType = value; return this; }, getContent: () => text }) },
    SpreadsheetApp: { getActive: () => ({ getSheetByName: name => { if (name === 'Заявки') state.sheetReads++; return name === 'Подписчики' && options.subscribers ? {
      getLastRow: () => options.subscribers.length,
      getRange: (row, col, count = 1, width = 1) => ({
        getValues: () => Array.from({ length: count }, (_, i) => Array.from({ length: width }, (_, j) => options.subscribers[row - 1 + i]?.[col - 1 + j] ?? '')),
        setValue: value => { options.subscribers[row - 1][col - 1] = value; }
      })
    } : (state.rows.length ? sheet : null); }, insertSheet: () => sheet, getUrl: () => 'https://docs.google.com/spreadsheets/d/local-test' }), flush() { state.flushes++; if (options.flushFails) throw new Error('sheet flush failure'); if (options.onFlush) options.onFlush({ state, cache }); } },
    LockService: { getScriptLock: () => { state.lockRequests++; return { waitLock: () => { state.locked = true; }, tryLock: () => { if (options.busyLock) return false; assert.equal(state.locked, false); state.locked = true; return true; }, releaseLock: () => { state.locked = false; } }; } },
    CacheService: { getUserCache: () => ({ get: key => clientCache.get(key), put: (key, value) => clientCache.set(key, value) }), getScriptCache: () => ({ get: key => { if (options.cacheReadFails) throw new Error('cache read failure'); return cache.get(key); }, put: (key, value) => { if (options.cacheFails) throw new Error('cache write failure'); state.cacheWrites++; cache.set(key, value); } }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: key => { if (options.propertyFails && key === 'ATTACHMENTS_FOLDER_ID') throw new Error('property service unavailable'); if (options.securityPropsFail && key.startsWith('SECURITY_')) throw new Error('secret details'); return props[key] || null; }, setProperty: (key, value) => { if (options.securityPropsFail) throw new Error('secret details'); props[key] = value; state.propsWrites.push(key); } }) },
    DriveApp: { Access: { PRIVATE: 'PRIVATE' }, getFolderById: id => { state.folderReads++; assert.equal(id, 'private-folder'); if (options.folderFails) throw new Error('missing folder'); return folder; } },
    Utilities: { formatDate: (date, tz, format) => format === 'yyyy-MM-dd' ? new Date(date.getTime() + 7 * 3600000).toISOString().slice(0, 10) : '01.10.2026 12:00', DigestAlgorithm: { SHA_256: 'sha256' }, Charset: { UTF_8: 'utf8' }, computeDigest: (algorithm, value) => Array.from(require('node:crypto').createHash('sha256').update(value).digest()), base64Decode: value => Array.from(Buffer.from(value, 'base64'), byte => byte > 127 ? byte - 256 : byte), base64Encode: bytes => Buffer.from(bytes).toString('base64'), newBlob: (bytes, mimeType, name) => ({ getBytes: () => typeof bytes === 'string' ? Array.from(Buffer.from(bytes)) : bytes, getContentType: () => mimeType, getName: () => name }) },
    Session: { getScriptTimeZone: () => 'Asia/Novosibirsk' },
    MailApp: { sendEmail: (to, subject, body, mailOptions) => { if (options.onMail) options.onMail({ state, context }); if (options.mailFails) throw new Error('mail quota'); state.mails.push({ to, subject, body, options: mailOptions }); } },
    UrlFetchApp: { fetch: (url, request) => { if (url.includes('smartcaptcha')) {
      state.captcha.push({ url, request });
      if (options.captchaFails) throw new Error('secret details');
      return { getResponseCode: () => options.captchaStatus || 200,
        getContentText: () => options.captchaRaw || JSON.stringify(options.captchaReply || { status: 'ok', host: 'kopilkaidei.ru' }) };
    }
    state.telegram.push({ url, request }); return { getContentText: () => '{"ok":true}' }; } }
  });
  vm.runInContext(source, context, { filename: 'Code.gs' });
  const submit = data => JSON.parse(context.doPost({ postData: { contents: JSON.stringify(data) } }).getContent());
  const lookup = id => JSON.parse(context.doGet({ parameter: { receipt: id } }).getContent());
  return { state, submit, lookup, context, props, cache, clientCache };
}


module.exports = { harness, payload, attachment, header };

const test = require('node:test');
const assert = require('node:assert/strict');
const { harness, payload, attachment, header } = require('./harness.cjs');

const receiptId = '7fbf0eaf1d1e4a57a4059683e62eab76';
const uuidReceiptId = '7fbf0eaf-1d1e-4a57-a405-9683e62eab76';

test('GET without receipt keeps the existing text health response', () => {
  const app = harness();
  for (const event of [undefined, {}, { parameter: {} }, { parameter: { other: 'value' } }]) {
    const response = app.context.doGet(event);
    assert.equal(response.getContent(), 'Копилка идей: сервис приёма заявок работает.');
    assert.equal(response.mimeType, 'TEXT');
  }
});

test('saved receipt can be read before synchronous notification finishes', () => {
  let observed;
  let first = true;
  const app = harness({ onMail({ state, context }) {
    if (!first) return;
    first = false;
    assert.equal(state.rows.length, 2);
    assert.equal(state.files.length, 1);
    assert.equal(state.mails.length, 0, 'the first mail has not completed');
    assert.equal(state.locked, false);
    observed = JSON.parse(context.doGet({ parameter: { receipt: receiptId } }).getContent());
  } });
  const result = app.submit(payload({ submissionId: receiptId, attachments: [attachment()] }));
  assert.equal(result.ok, true);
  assert.deepEqual(observed, { ok: true, number: 1, attachmentCount: 1 });
  assert.equal(app.state.mails.length, 2);
});

test('receipt lookup accepts strong hex and version 4 UUID identifiers', () => {
  for (const id of [receiptId, uuidReceiptId, receiptId.toUpperCase(), uuidReceiptId.toUpperCase()]) {
    const app = harness({ cacheEntries: [['sub_' + id, JSON.stringify({ number: 12, attachmentCount: 2 })]] });
    assert.deepEqual(app.lookup(id), { ok: true, number: 12, attachmentCount: 2 });
  }
});

test('receipt exposes only the acceptance number and exact saved attachment count', () => {
  const app = harness({ cacheEntries: [['sub_' + receiptId, JSON.stringify({ number: 12, attachmentCount: 5, fullName: 'private', email: 'private@example.ru', url: 'https://example.ru/private', token: 'secret' })]] });
  assert.deepEqual(app.lookup(receiptId), { ok: true, number: 12, attachmentCount: 5 });
});

test('legacy numeric cache receipt confirms zero attachments', () => {
  const app = harness({ cacheEntries: [['sub_' + receiptId, '42']] });
  assert.deepEqual(app.lookup(receiptId), { ok: true, number: 42, attachmentCount: 0 });
});

test('missing and expired cache entries stay pending rather than claiming submission failure', () => {
  const app = harness();
  assert.deepEqual(app.lookup(receiptId), { ok: false, pending: true });
  app.cache.set('sub_' + receiptId, JSON.stringify({ number: 4, attachmentCount: 1 }));
  app.cache.delete('sub_' + receiptId);
  assert.deepEqual(app.lookup(receiptId), { ok: false, pending: true });
});

test('weak and malformed receipt identifiers never expose cached receipts', () => {
  for (const id of ['', 'request-1', '1720000000000-random', receiptId + '0', receiptId.slice(1), 'z'.repeat(32), '7fbf0eaf-1d1e-1a57-a405-9683e62eab76', '7fbf0eaf-1d1e-4a57-7405-9683e62eab76', receiptId + '\n', ' ' + receiptId]) {
    const app = harness({ cacheEntries: [['sub_' + id, '{"number":4,"attachmentCount":1}']] });
    const result = app.lookup(id);
    assert.equal(result.ok, false);
    assert.equal(typeof result.error, 'string');
    assert.equal(result.number, undefined);
    assert.equal(result.attachmentCount, undefined);
  }
});

test('corrupt and invalid cached receipts remain unconfirmed', () => {
  const invalid = ['not json', 'null', '{}', '[]', '"42"', '0', '-1', '1.5', '9007199254740992'];
  for (const number of [0, -1, 1.5, '12', Number.MAX_SAFE_INTEGER + 1]) invalid.push(JSON.stringify({ number, attachmentCount: 1 }));
  for (const attachmentCount of [-1, 1.5, '1', 6, Number.MAX_SAFE_INTEGER + 1, null]) invalid.push(JSON.stringify({ number: 12, attachmentCount }));
  invalid.push(JSON.stringify({ number: 12 }));
  for (const value of invalid) {
    const app = harness({ cacheEntries: [['sub_' + receiptId, value]] });
    const result = app.lookup(receiptId);
    assert.equal(result.ok, false, value);
    assert.equal(result.number, undefined, value);
  }
});

test('cache read failure stays pending and does not invent a receipt', () => {
  const app = harness({ cacheReadFails: true });
  assert.deepEqual(app.lookup(receiptId), { ok: false, pending: true });
});

test('receipt lookup never writes to storage, takes a lock, or sends notifications', () => {
  const app = harness({ cacheEntries: [['sub_' + receiptId, '{"number":4,"attachmentCount":1}']] });
  const before = JSON.stringify(app.state);
  const originalEntries = Array.from(app.cache);
  app.context.SpreadsheetApp.getActive = () => { throw new Error('receipt lookup must not access Sheets'); };
  assert.deepEqual(app.lookup(receiptId), { ok: true, number: 4, attachmentCount: 1 });
  assert.equal(JSON.stringify(app.state), before);
  assert.deepEqual(Array.from(app.cache), originalEntries);
});

test('receipt is published only after pending Sheet writes are flushed', () => {
  let flushed = false;
  const app = harness({ onFlush({ state, cache }) {
    assert.equal(state.rows.length, 2);
    assert.equal(cache.has('sub_' + receiptId), false);
    flushed = true;
  }, onMail({ state }) {
    assert.equal(flushed, true);
    assert.equal(state.flushes, 1);
  } });
  assert.equal(app.submit(payload({ submissionId: receiptId, attachments: [attachment()] })).ok, true);
  assert.equal(flushed, true);
  assert.deepEqual(app.lookup(receiptId), { ok: true, number: 1, attachmentCount: 1 });
});

test('link-only submissions retain original column order and send ordinary notifications', () => {
  const app = harness({ props: { ATTACHMENTS_FOLDER_ID: '' } });
  assert.deepEqual(app.submit(payload({ materials: 'https://example.ru/materials' })), { ok: true, number: 1, duplicate: false, attachmentCount: 0 });
  assert.deepEqual(app.state.rows[0], header);
  assert.equal(app.state.rows[1][14], 'https://example.ru/materials');
  assert.deepEqual(app.state.rows[1].slice(16), ['Новая', '']);
  assert.equal(app.state.mails.length, 2);
  assert.equal(app.state.mails[0].options.attachments, undefined);
  assert.equal(app.state.folderReads, 0);
});

test('uploaded bytes are stored privately, linked after existing columns and attached to team email', () => {
  const app = harness({ rows: [header, [1, 'date', 'Старый автор']] });
  assert.deepEqual(app.submit(payload({ attachments: [attachment('deck.PPTX', Buffer.from([0, 128, 255]))] })), { ok: true, number: 2, duplicate: false, attachmentCount: 1 });
  assert.deepEqual(app.state.rows[0], [...header, 'Файлы']);
  assert.equal(app.state.rows[1].length, 3, 'old row values must not move');
  assert.deepEqual(app.state.rows[2].slice(16, 18), ['Новая', '']);
  assert.match(app.state.rows[2][18], /deck\.PPTX.*https:\/\/drive\.google\.com/s);
  assert.equal(app.state.files.length, 1);
  const blob = app.state.mails[0].options.attachments[0];
  assert.equal(blob.getName(), 'deck.PPTX');
  assert.equal(blob.getContentType(), 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
  assert.deepEqual(Buffer.from(blob.getBytes()), Buffer.from([0, 128, 255]));
  assert.equal(app.state.mails[1].options.attachments, undefined);
  assert.equal(app.state.files[0].trashed, false);
});

test('existing Files column is reused without shifting custom columns', () => {
  const app = harness({ rows: [[...header, 'Файлы', 'Внутренний код']] });
  assert.equal(app.submit(payload({ attachments: [attachment()] })).ok, true);
  assert.equal(app.state.rows[0].length, 20);
  assert.match(app.state.rows[1][18], /idea\.pdf/);
});

test('file names discard paths, controls and keep allowed extension while truncating', () => {
  const app = harness();
  assert.equal(app.submit(payload({ attachments: [attachment('C:\\fakepath\\bad\r\nname.pdf'), attachment('a'.repeat(240) + '.docx')] })).ok, true);
  assert.equal(app.state.files[0].blob.getName(), 'badname.pdf');
  const name = app.state.files[1].blob.getName();
  assert.ok(name.length <= 180);
  assert.ok(name.endsWith('.docx'));
});

test('repeat submissions do not store files or send notifications twice', () => {
  const app = harness();
  const data = payload({ attachments: [attachment()] });
  assert.equal(app.submit(data).ok, true);
  app.props.ATTACHMENTS_FOLDER_ID = '';
  assert.deepEqual(app.submit(data), { ok: true, number: 1, duplicate: true, attachmentCount: 1 });
  assert.equal(app.state.rows.length, 2);
  assert.equal(app.state.files.length, 1);
  assert.equal(app.state.mails.length, 2);
  assert.equal(app.state.locked, false);
});

for (const [name, uploads, message] of [
  ['non-array attachments', {}, /список|массив/i],
  ['six files', Array.from({ length: 6 }, () => attachment()), /5/],
  ['empty bytes', [attachment('empty.pdf', '')], /пуст|размер/i],
  ['declared size mismatch', [{ ...attachment(), size: 1 }], /размер/i],
  ['fractional size', [{ ...attachment(), size: 3.1 }], /размер/i],
  ['string size', [{ ...attachment(), size: '3' }], /размер/i],
  ['invalid base64 alphabet', [{ ...attachment(), base64: 'YWJ!' }], /данные|base64/i],
  ['noncanonical pad bits', [{ ...attachment('idea.pdf', 'a'), base64: 'YR==' }], /данные|base64/i],
  ['missing base64 padding', [{ ...attachment('idea.pdf', 'a'), base64: 'YQ' }], /данные|base64/i],
  ['line breaks in base64', [{ ...attachment(), base64: 'YW\nJj' }], /данные|base64/i],
  ['unsupported executable', [attachment('run.exe')], /формат|тип|расширение/i],
  ['unsupported html', [attachment('page.html')], /формат|тип|расширение/i],
  ['unsupported svg', [attachment('image.svg')], /формат|тип|расширение/i],
  ['missing name', [{ ...attachment(), name: '' }], /имя|название/i],
  ['null attachment', [null], /файл|данные/i]
]) {
  test('rejects ' + name + ' without writing rows, files or notifications', () => {
    const app = harness();
    const result = app.submit(payload({ attachments: uploads }));
    assert.equal(result.ok, false);
    assert.match(result.error, /файл|данные|размер|формат|список|5|пяти/i);
    assert.equal(app.state.rows.length, 0);
    assert.equal(app.state.files.length, 0);
    assert.equal(app.state.mails.length, 0);
    assert.equal(app.state.locked, false);
  });
}

test('decoded total above 10 MiB is rejected before storage', () => {
  const app = harness();
  const file = attachment('large.pdf', Buffer.alloc(5 * 1024 * 1024 + 1, 1));
  const result = app.submit(payload({ attachments: [file, file] }));
  assert.equal(result.ok, false);
  assert.match(result.error, /10.*М[Бб]|10.*MiB/);
  assert.equal(app.state.files.length, 0);
  assert.equal(app.state.rows.length, 0);
});

test('five files totaling exactly 10 MiB are accepted', () => {
  const app = harness();
  const file = attachment('archive.pdf', Buffer.alloc(2 * 1024 * 1024, 1));
  assert.equal(app.submit(payload({ attachments: [file, file, file, file, file] })).ok, true);
  assert.equal(app.state.files.length, 5);
});

for (const [name, options, error] of [
  ['missing configured folder', { props: { ATTACHMENTS_FOLDER_ID: '' } }, /ATTACHMENTS_FOLDER_ID/],
  ['unavailable folder', { folderFails: true }, /папк|доступ/i],
  ['public folder', { folderAccess: 'ANYONE_WITH_LINK' }, /доступ|закрыт/i],
  ['trashed folder', { folderTrashed: true }, /папк|корзин/i]
]) {
  test('rejects ' + name + ' explicitly before storing submission', () => {
    const app = harness(options);
    const result = app.submit(payload({ attachments: [attachment()] }));
    assert.equal(result.ok, false);
    assert.match(result.error, error);
    assert.equal(app.state.rows.length, 0);
    assert.equal(app.state.files.length, 0);
    assert.equal(app.state.mails.length, 0);
  });
}

test('Drive failure rolls back only files created by current submission', () => {
  const app = harness({ failCreateAt: 3 });
  assert.equal(app.submit(payload({ attachments: [attachment('previous.pdf')] })).ok, true);
  const result = app.submit(payload({ submissionId: '7fbf0eaf1d1e4a57a4059683e62eab77', clientId: '7fbf0eaf1d1e4a57a4059683e62eab77', attachments: [attachment('first.pdf'), attachment('second.pdf')] }));
  assert.equal(result.ok, false);
  assert.match(result.error, /сохран|загруз.*файл/i);
  assert.equal(app.state.rows.length, 2);
  assert.equal(app.state.files[0].trashed, false);
  assert.equal(app.state.files[1].trashed, true);
  assert.equal(app.state.mails.length, 2);
});

for (const options of [{ appendFails: true }, { appendWritesThenFails: true }]) {
  test('sheet append failure leaves no partial submission and trashes current files: ' + JSON.stringify(options), () => {
    const app = harness({ ...options, rows: [header] });
    const result = app.submit(payload({ attachments: [attachment()] }));
    assert.equal(result.ok, false);
    assert.match(result.error, /сохран|запис/i);
    assert.equal(app.state.rows.length, 1);
    assert.equal(app.state.files[0].trashed, true);
    assert.equal(app.state.mails.length, 0);
  });
}

test('notification errors preserve saved application and files', () => {
  const app = harness({ mailFails: true });
  assert.equal(app.submit(payload({ attachments: [attachment()] })).ok, true);
  assert.equal(app.state.rows.length, 2);
  assert.equal(app.state.files[0].trashed, false);
});

for (const options of [{ formatFails: true }, { flushFails: true }, { cacheFails: true }]) {
  test('failure before durable submission acknowledgment rolls back row and files: ' + JSON.stringify(options), () => {
    const app = harness({ ...options, rows: [header] });
    const result = app.submit(payload({ attachments: [attachment()] }));
    assert.equal(result.ok, false);
    assert.match(result.error, /сохран|запис/i);
    assert.equal(app.state.rows.length, 1);
    assert.equal(app.state.files[0].trashed, true);
  });
}

test('unavailable script properties yield explicit attachment setup failure', () => {
  const app = harness({ propertyFails: true });
  const result = app.submit(payload({ attachments: [attachment()] }));
  assert.equal(result.ok, false);
  assert.match(result.error, /настрой|ATTACHMENTS_FOLDER_ID/i);
  assert.equal(app.state.rows.length, 0);
});

test('attachments keep existing Telegram text and subscription delivery without duplicates', () => {
  const app = harness({
    props: { TELEGRAM_TOKEN: '1234567890:abcdefghijklmnopqrstuvwxyz' },
    subscribers: [['Chat ID', 'Имя', 'Ник', 'Подписан', 'Статус'], ['123', 'Иван', '@ivan', 'date', 'активен']]
  });
  const data = payload({ materials: 'https://example.ru/materials', attachments: [attachment()] });
  assert.equal(app.submit(data).ok, true);
  assert.equal(app.state.telegram.length, 1);
  const sent = JSON.parse(app.state.telegram[0].request.payload);
  assert.equal(sent.chat_id, '123');
  assert.match(sent.text, /Новая идея №1/);
  assert.match(sent.text, /https:\/\/example\.ru\/materials/);
  assert.equal(app.submit(data).duplicate, true);
  assert.equal(app.state.telegram.length, 1);
});

test('receipt reports exact accepted attachment count for multi-file submission and duplicate', () => {
  const app = harness();
  const data = payload({ attachments: [attachment('one.pdf'), attachment('two.pptx')] });
  assert.deepEqual(app.submit(data), { ok: true, number: 1, duplicate: false, attachmentCount: 2 });
  assert.deepEqual(app.submit(data), { ok: true, number: 1, duplicate: true, attachmentCount: 2 });
  assert.equal(app.state.files.length, 2);
  assert.equal(app.state.rows.length, 2);
  assert.equal(app.state.mails.length, 2);
});

test('legacy numeric cache receipt acknowledges zero files without new writes', () => {
  const app = harness({ cacheEntries: [['sub_7fbf0eaf1d1e4a57a4059683e62eab76', '42']] });
  assert.deepEqual(app.submit(payload()), { ok: true, number: 42, duplicate: true, attachmentCount: 0 });
  assert.equal(app.state.files.length, 0);
  assert.equal(app.state.rows.length, 0);
  assert.equal(app.state.mails.length, 0);
});

test('legacy numeric cache never claims newly submitted attachments were accepted', () => {
  const app = harness({ cacheEntries: [['sub_7fbf0eaf1d1e4a57a4059683e62eab76', '42']] });
  assert.deepEqual(app.submit(payload({ attachments: [attachment()] })), { ok: true, number: 42, duplicate: true, attachmentCount: 0 });
  assert.equal(app.state.files.length, 0);
  assert.equal(app.state.rows.length, 0);
  assert.equal(app.state.mails.length, 0);
});

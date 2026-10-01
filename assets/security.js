var IdeaSecurity = (function () {
  var memoryId = null;
  function strongId(id) {
    return typeof id === 'string' && (id.length === 32 || id.length === 36) &&
      /^(?:[a-f0-9]{32}|[a-f0-9]{8}-[a-f0-9]{4}-4[a-f0-9]{3}-[89ab][a-f0-9]{3}-[a-f0-9]{12})$/i.test(id);
  }
  function clientId() {
    if (memoryId) return memoryId;
    try { memoryId = sessionStorage.getItem('idea-client-id'); } catch (ignore) {}
    if (!strongId(memoryId)) memoryId = IdeaAttachments.newId();
    try { sessionStorage.setItem('idea-client-id', memoryId); } catch (ignore) {}
    return memoryId;
  }
  function checkTransport(preview) {
    if (!preview && location.protocol !== 'https:' && location.protocol !== 'http:') {
      return Promise.reject(new Error('Для отправки откройте сайт kopilkaidei.ru. Данные остались в форме.'));
    }
    return Promise.resolve();
  }
  function receipt(endpoint, id, count) {
    if (!strongId(id) || (location.protocol !== 'https:' && location.protocol !== 'http:')) return Promise.resolve(null);
    var controller = typeof AbortController !== 'undefined' ? new AbortController() : null;
    return new Promise(function (resolve) {
      var settled = false;
      var timer = setTimeout(function () { if (controller) controller.abort(); finish(null); }, 10000);
      function finish(value) {
        if (settled) return;
        settled = true;
        clearTimeout(timer);
        resolve(value);
      }
      var options = { method: 'GET', cache: 'no-store' };
      if (controller) options.signal = controller.signal;
      Promise.resolve().then(function () {
        return fetch(endpoint + (endpoint.indexOf('?') === -1 ? '?' : '&') + 'receipt=' + encodeURIComponent(id), options);
      }).then(function (response) { return response.ok ? response.json() : null; }).then(function (result) {
        finish(IdeaAttachments.isConfirmed(result) && result.attachmentCount === count ? result : null);
      }).catch(function () { finish(null); });
    });
  }
  return { clientId: clientId, checkTransport: checkTransport, receipt: receipt };
})();

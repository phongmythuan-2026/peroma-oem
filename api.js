// Thay thế google.script.run: gọi backend Google Apps Script (Web App) qua fetch.
// Dán URL Web App (kết thúc bằng /exec) vào đây sau khi Deploy.
window.GAS_API_URL = 'https://script.google.com/macros/s/AKfycbzQsZFtdxzwLt0iHp1PpaBjjMH1K7onXUtDDOtaX4Ly98aKeUuDD3lXt4FtJxhN1hfXYg/exec';

(function () {
  function announceDataSync(version) {
    if (version === undefined || version === null) return;
    const signal = { version: String(version), timestamp: Date.now() };
    try {
      localStorage.setItem('OEM_DATA_SYNC_SIGNAL', JSON.stringify(signal));
    } catch (error) {
      console.warn('Could not broadcast OEM data sync to other tabs:', error);
    }
    window.dispatchEvent(new CustomEvent('oem-data-sync', {
      detail: { version: signal.version, source: 'local' }
    }));
  }

  function createRunner(success, failure) {
    return new Proxy({}, {
      get: function (_, name) {
        if (name === 'withSuccessHandler') return function (fn) { return createRunner(fn, failure); };
        if (name === 'withFailureHandler') return function (fn) { return createRunner(success, fn); };
        if (name === 'withUserObject') return function () { return createRunner(success, failure); };
        if (name === 'then') return undefined;
        return function () {
          const args = Array.prototype.slice.call(arguments);
          // text/plain tránh preflight CORS của Apps Script
          fetch(window.GAS_API_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            body: JSON.stringify({ fn: String(name), args: args })
          })
            .then(function (r) { return r.json(); })
            .then(function (res) {
              if (res && res.__error) throw new Error(res.__error);
              if (res && res.syncVersion !== undefined) announceDataSync(res.syncVersion);
              if (success) success(res ? res.data : res);
            })
            .catch(function (err) {
              if (failure) failure(err);
              else console.error('GAS call failed:', name, err);
            });
        };
      }
    });
  }
  window.google = window.google || {};
  window.google.script = { run: createRunner(null, null) };
})();

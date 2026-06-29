// storage.js — localStorage auto-save with debounce
var AutoSaveStorage = (function () {
  'use strict';
  var NS = 'excel-web-editor';
  var DEBOUNCE = 500;

  function Storage(opts) {
    opts = opts || {};
    this.ns = opts.namespace || NS;
    this.debounce = opts.debounceMs || DEBOUNCE;
    this.onSave = opts.onSave || function () {};
    this.onError = opts.onError || function () {};
    this._timers = {};
    this._pending = {};
  }

  Storage.prototype._key = function (id) { return this.ns + '::' + id; };

  Storage.prototype.available = function () {
    try { var k = this._key('_p_'); localStorage.setItem(k, '1'); localStorage.removeItem(k); return true; }
    catch (_) { return false; }
  };

  Storage.prototype.save = function (id, val) {
    var self = this;
    this._pending[id] = val;
    if (this._timers[id]) clearTimeout(this._timers[id]);
    this._timers[id] = setTimeout(function () {
      delete self._timers[id];
      self._commit(id);
    }, this.debounce);
  };

  Storage.prototype._commit = function (id) {
    var val = this._pending[id];
    if (val === undefined) return;
    delete this._pending[id];
    try {
      localStorage.setItem(this._key(id), JSON.stringify(val));
      this.onSave(id, val);
    } catch (e) {
      this.onError(e, id);
    }
  };

  Storage.prototype.flushAll = function () {
    var self = this;
    Object.keys(this._pending).forEach(function (id) {
      if (self._timers[id]) clearTimeout(self._timers[id]);
      delete self._timers[id];
      self._commit(id);
    });
  };

  Storage.prototype.load = function (id) {
    var raw = localStorage.getItem(this._key(id));
    if (!raw) return null;
    try { return JSON.parse(raw); } catch (_) { return null; }
  };

  Storage.prototype.loadAll = function () {
    var result = {};
    var prefix = this.ns + '::';
    for (var i = 0; i < localStorage.length; i++) {
      var k = localStorage.key(i);
      if (k && k.indexOf(prefix) === 0) result[k.slice(prefix.length)] = this.load(k.slice(prefix.length));
    }
    return result;
  };

  Storage.prototype.remove = function (id) {
    if (this._timers[id]) clearTimeout(this._timers[id]);
    delete this._timers[id];
    delete this._pending[id];
    localStorage.removeItem(this._key(id));
  };

  Storage.prototype.listFiles = function () { return Object.keys(this.loadAll()); };

  return Storage;
})();

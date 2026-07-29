'use strict';

const CACHE_TTL_SECONDS = 21600; // max TTL GAS CacheService (6 jam); aman krn tiap mutasi invalidate cache eksplisit
const CACHE_KEY_PREFIX = 'cfg_';
const CACHE_KEY_ALL = 'cfg_all';
const CACHE_CHUNK_SIZE = 90000;

const CacheHelper = {
  getConfig: function(year) {
    const key = CACHE_KEY_PREFIX + String(year);
    try {
      const cache = CacheService.getScriptCache();
      const cached = this._getLargeCache(cache, key);
      if (cached) {
        const parsed = JSON.parse(cached);
        if (parsed && parsed.selectedYear) return parsed;
      }
    } catch (e) {
      console.warn('CacheHelper.getConfig cache read error: ' + e.message);
    }
    return this._fetchAndStore(year);
  },

  invalidate: function (year) {
    try {
      const cache = CacheService.getScriptCache();
      this._removeLargeCache(cache, CACHE_KEY_PREFIX + String(year));
    } catch (e) {
      console.warn('CacheHelper.invalidate error: ' + e.message);
    }
  },

  _registerYear_: function (year) {
    try {
      const cache = CacheService.getScriptCache();
      const existing = cache.get(CACHE_KEY_ALL);
      const years = existing ? JSON.parse(existing) : [];
      if (years.indexOf(year) === -1) {
        years.push(year);
        cache.put(CACHE_KEY_ALL, JSON.stringify(years), CACHE_TTL_SECONDS + 60);
      }
    } catch (e) {
      console.warn('CacheHelper._registerYear_ error: ' + e.message);
    }
  },

  _getCachedYearRegistry_: function () {
    try {
      const cache = CacheService.getScriptCache();
      const existing = cache.get(CACHE_KEY_ALL);
      return existing ? JSON.parse(existing) : [];
    } catch (e) {
      console.warn('CacheHelper._getCachedYearRegistry_ error: ' + e.message);
      return [];
    }
  },

  invalidateAll: function() {
    try {
      const cache = CacheService.getScriptCache();
      const registry = this._getCachedYearRegistry_();
      registry.forEach(y => this._removeLargeCache(cache, CACHE_KEY_PREFIX + String(y)));
      cache.remove(CACHE_KEY_ALL);
    } catch (e) {
      console.warn('CacheHelper.invalidateAll error: ' + e.message);
    }
  },

  _fetchAndStore: function(year) {
    const config = ConfigRepository.loadAll(year);
    try {
      const cache = CacheService.getScriptCache();
      const key = CACHE_KEY_PREFIX + String(year);
      const serialized = JSON.stringify(config);
      this._putLargeCache(cache, key, serialized, CACHE_TTL_SECONDS);
      this._registerYear_(year);
    } catch (e) {
      console.warn('CacheHelper._fetchAndStore cache write error: ' + e.message);
    }
    return config;
  },

  _putLargeCache: function(cache, key, str, ttl) {
    if (str.length < CACHE_CHUNK_SIZE) {
      cache.put(key, str, ttl);
      return;
    }
    const chunks = [];
    for (let i = 0; i < str.length; i += CACHE_CHUNK_SIZE) {
      chunks.push(str.substring(i, i + CACHE_CHUNK_SIZE));
    }
    const manifest = { chunks: chunks.length };
    const toPut = {};
    toPut[key] = JSON.stringify(manifest);
    for (let i = 0; i < chunks.length; i++) {
      toPut[key + '_c' + i] = chunks[i];
    }
    cache.putAll(toPut, ttl);
  },

  _getLargeCache: function(cache, key) {
    const head = cache.get(key);
    if (!head) return null;
    if (head.startsWith('{"chunks":')) {
      try {
        const manifest = JSON.parse(head);
        if (manifest.chunks) {
          const keys = [];
          for (let i = 0; i < manifest.chunks; i++) {
            keys.push(key + '_c' + i);
          }
          const chunks = cache.getAll(keys);
          let fullStr = '';
          for (let i = 0; i < manifest.chunks; i++) {
            const chunk = chunks[key + '_c' + i];
            if (!chunk) return null; 
            fullStr += chunk;
          }
          return fullStr;
        }
      } catch (e) {
        // Fallback to treat as normal string
      }
    }
    return head;
  },

  _removeLargeCache: function(cache, key) {
    const head = cache.get(key);
    if (head && head.startsWith('{"chunks":')) {
      try {
        const manifest = JSON.parse(head);
        if (manifest.chunks) {
          const keys = [key];
          for (let i = 0; i < manifest.chunks; i++) {
            keys.push(key + '_c' + i);
          }
          cache.removeAll(keys);
          return;
        }
      } catch (e) {}
    }
    cache.remove(key);
  }
};

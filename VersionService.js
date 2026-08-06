'use strict';

const VERSION_KEY = 'SYS_VERSION';
const VERSION_CACHE_TTL = 300;

function getVersion() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(VERSION_KEY);
  if (cached !== null) return { version: Number(cached) };
  const props = PropertiesService.getScriptProperties();
  const v = Number(props.getProperty(VERSION_KEY) || 0);
  cache.put(VERSION_KEY, String(v), VERSION_CACHE_TTL);
  return { version: v };
}

function bumpVersion() {
  const props = PropertiesService.getScriptProperties();
  const v = Number(props.getProperty(VERSION_KEY) || 0) + 1;
  props.setProperty(VERSION_KEY, String(v));
  const cache = CacheService.getScriptCache();
  cache.put(VERSION_KEY, String(v), VERSION_CACHE_TTL);
  CacheHelper.invalidateAll();
  return v;
}

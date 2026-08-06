'use strict';

// ============================================================
// Resumable Upload (>50 MB) via Drive API v3
// ============================================================

const RESUMABLE_SESSION_PROPERTY_PREFIX = 'upload_sess_';
const RESUMABLE_SESSION_TTL_MS = 30 * 60 * 1000;
const RESUMABLE_UPLOAD_MAX_CHUNK_BYTES = 10 * 1024 * 1024;

function assertResumableSession_(sessionUrl) {
  const value = String(sessionUrl || '');
  const prefix = 'https://www.googleapis.com/upload/drive/v3/files';
  if (value.indexOf(prefix) !== 0) throw new Error('Session upload tidak valid.');
  const suffix = value.slice(prefix.length);
  if (suffix && suffix.charAt(0) !== '?') throw new Error('Session upload tidak valid.');
  return value;
}

function cleanupExpiredResumableSessions_() {
  const props = PropertiesService.getScriptProperties();
  const keys = props.getKeys();
  const now = Date.now();
  for (let i = 0; i < keys.length; i++) {
    const key = keys[i];
    if (key.indexOf(RESUMABLE_SESSION_PROPERTY_PREFIX) !== 0) continue;
    try {
      const raw = props.getProperty(key);
      const record = raw ? JSON.parse(raw) : null;
      if (!record || !record.expiresAt || now >= Number(record.expiresAt)) props.deleteProperty(key);
    } catch (_) { props.deleteProperty(key); }
  }
}

function registerResumableSession_(sessionUrl, totalSize) {
  cleanupExpiredResumableSessions_();
  const actor = getRequestPortalUser_() || {};
  if (!actor.accountId) throw accessDeniedError_('UPLOAD_SESSION_ACTOR', 'Sesi portal tidak valid.');
  const token = Utilities.getUuid();
  const record = {
    sessionUrl: assertResumableSession_(sessionUrl),
    totalSize: Number(totalSize),
    accountId: actor.accountId,
    expiresAt: Date.now() + RESUMABLE_SESSION_TTL_MS
  };
  PropertiesService.getScriptProperties().setProperty(RESUMABLE_SESSION_PROPERTY_PREFIX + token, JSON.stringify(record));
  return token;
}

function loadResumableSession_(token, totalSize) {
  token = String(token || '');
  if (token.length < 20 || token.length > 100 || !/^[A-Za-z0-9_-]+$/.test(token)) {
    throw new Error('Session upload tidak valid.');
  }
  const props = PropertiesService.getScriptProperties();
  const key = RESUMABLE_SESSION_PROPERTY_PREFIX + token;
  const raw = props.getProperty(key);
  if (!raw) throw new Error('Session upload tidak ditemukan atau sudah berakhir.');
  let record;
  try { record = JSON.parse(raw); } catch (_) { props.deleteProperty(key); throw new Error('Session upload tidak valid.'); }
  if (!record.expiresAt || Date.now() >= Number(record.expiresAt)) {
    props.deleteProperty(key);
    throw new Error('Session upload sudah berakhir. Mulai ulang unggahan.');
  }
  const actor = getRequestPortalUser_() || {};
  if (!actor.accountId || actor.accountId !== record.accountId) {
    throw accessDeniedError_('UPLOAD_SESSION_OWNER', 'Session upload bukan milik akun ini.');
  }
  if (Number(totalSize) !== Number(record.totalSize)) throw new Error('Ukuran upload berubah dari sesi awal.');
  record.sessionUrl = assertResumableSession_(record.sessionUrl);
  return record;
}

function deleteResumableSession_(token) {
  try { PropertiesService.getScriptProperties().deleteProperty(RESUMABLE_SESSION_PROPERTY_PREFIX + String(token || '')); } catch (_) {}
}

function resumableUploadInit_(name, mimeType, totalSize, folderId) {
  const safeName = validateDriveItemName_(name, 'Nama file');
  const metadata = { name: safeName };
  if (folderId) metadata.parents = [cleanId_(folderId)];

  const token = ScriptApp.getOAuthToken();
  const headers = {
    Authorization: 'Bearer ' + token,
    'Content-Type': 'application/json; charset=UTF-8',
    'X-Upload-Content-Type': mimeType || 'application/octet-stream',
    'X-Upload-Content-Length': String(totalSize)
  };

  const response = UrlFetchApp.fetch(
    'https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable',
    {
      method: 'POST',
      headers: headers,
      payload: JSON.stringify(metadata),
      muteHttpExceptions: true
    }
  );

  const code = response.getResponseCode();
  if (code < 200 || code >= 300) {
    const body = response.getContentText();
    throw new Error('Gagal mulai upload besar: ' + (body ? body.slice(0, 200) : 'HTTP ' + code));
  }

  const location = response.getHeaders()['Location'];
  if (!location) throw new Error('Server tidak mengembalikan URL upload.');
  const sessionToken = registerResumableSession_(location, totalSize);
  return { sessionUrl: sessionToken, totalSize: totalSize };
}

function resumableUploadChunk_(sessionToken, chunkBase64, startByte, totalSize) {
  if (!Number.isInteger(startByte) || startByte < 0) throw new Error('Posisi chunk tidak valid.');
  const session = loadResumableSession_(sessionToken, totalSize);
  const chunkBytes = Utilities.base64Decode(chunkBase64);
  const chunkSize = chunkBytes.length;
  if (!chunkSize || chunkSize > RESUMABLE_UPLOAD_MAX_CHUNK_BYTES) throw new Error('Ukuran chunk tidak valid.');
  const endByte = startByte + chunkSize - 1;
  if (endByte >= totalSize) {
    if (endByte !== totalSize - 1) throw new Error('Chunk melewati ukuran file.');
  }

  const response = UrlFetchApp.fetch(session.sessionUrl, {
    method: 'PUT',
    headers: { 'Content-Range': 'bytes ' + startByte + '-' + endByte + '/' + totalSize },
    payload: chunkBytes,
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();
  if (code === 200 || code === 201) {
    deleteResumableSession_(sessionToken);
    const fileData = JSON.parse(response.getContentText());
    return {
      done: true,
      fileId: fileData.id,
      name: fileData.name,
      url: 'https://drive.google.com/file/d/' + fileData.id + '/view',
      mimeType: fileData.mimeType,
      size: Number(fileData.size || 0)
    };
  }

  if (code === 308) {
    const range = response.getHeaders()['Range'];
    let nextByte = startByte + chunkSize;
    if (range) {
      const m = range.match(/([0-9]+)$/);
      if (m) nextByte = parseInt(m[1], 10) + 1;
    }
    return { done: false, nextByte: nextByte };
  }

  if (code === 404 || code === 410) deleteResumableSession_(sessionToken);
  const body = response.getContentText();
  throw new Error('Upload gagal (HTTP ' + code + '): ' + (body ? body.slice(0, 200) : ''));
}

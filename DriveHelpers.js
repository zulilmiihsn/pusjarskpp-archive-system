'use strict';

function validateDriveItemName_(name, label) {
  const cleanName = String(name || '').trim();
  if (!cleanName) throw new Error((label || 'Nama') + ' wajib diisi.');
  if (cleanName.length > MAX_FILENAME_LENGTH) throw new Error((label || 'Nama') + ' maksimal ' + MAX_FILENAME_LENGTH + ' karakter.');
  if (/[\\/:*?"<>|]/.test(cleanName)) {
    throw new Error((label || 'Nama') + ' tidak boleh berisi karakter \\ / : * ? " < > |.');
  }
  return cleanName;
}

/**
 * Escape a value for safe inclusion in a Drive API v2 `q` query string.
 * Escapes backslashes and single quotes to prevent query injection.
 */
function escapeDriveQueryValue_(value) {
  return String(value || '').replace(/\\/g, '\\\\').replace(/'/g, "\\'");
}

function extractDriveTargetId_(value) {
  const text = String(value || '').trim();
  if (!text) return '';

  const patterns = [
    /\/d\/([-\w]{25,})/i,
    /\/folders\/([-\w]{25,})/i,
    /[?&]id=([-\w]{25,})/i
  ];
  for (let i = 0; i < patterns.length; i++) {
    const match = text.match(patterns[i]);
    if (match) return match[1];
  }

  const fallback = text.match(/[-\w]{25,}/);
  return fallback ? fallback[0] : '';
}

function resolveArchiveDocumentShortcutTarget_(url) {
  let targetId = extractDriveTargetId_(url);
  if (!targetId) throw new Error('Link Google Drive tidak valid.');

  for (let depth = 0; depth < 3; depth++) {
    const target = Drive.Files.get(targetId, {
      supportsAllDrives: true,
      fields: 'id,name,mimeType,shortcutDetails'
    });

    if (target.mimeType !== 'application/vnd.google-apps.shortcut') {
      return target;
    }

    if (!target.shortcutDetails || !target.shortcutDetails.targetId) {
      throw new Error('Gagal membuat link dokumen: Link mengarah ke shortcut Drive yang target aslinya tidak dapat dibaca.');
    }
    targetId = target.shortcutDetails.targetId;
  }

  throw new Error('Gagal membuat link dokumen: Link Drive berisi rangkaian shortcut terlalu dalam. Salin link file atau folder asli, bukan shortcut.');
}

function buildArchiveDocumentShortcutError_(error, target) {
  const message = String(error && error.message ? error.message : error || '');
  if (message.indexOf('Gagal membuat link dokumen:') === 0 || message.indexOf('Link Google Drive tidak valid.') === 0) {
    return new Error(message);
  }

  const targetLabel = target && target.name ? '"' + target.name + '"' : 'target link';
  const targetMime = target && target.mimeType ? ' (' + target.mimeType + ')' : '';
  if (/shortcutTargetInvalid|not an allowed shortcut target type/i.test(message)) {
    return new Error(
      'Gagal membuat link dokumen: Link mengarah ke ' + targetLabel + targetMime +
      ', tetapi tipe item ini tidak bisa dijadikan shortcut Drive. Salin link file atau folder asli di Google Drive.'
    );
  }
  if (/not found|File not found|insufficient|permission|access/i.test(message)) {
    return new Error('Gagal membuat link dokumen: Target link tidak dapat dibaca oleh akun ini. Pastikan file/folder asli masih ada dan akun Anda memiliki akses.');
  }
  return new Error('Gagal membuat link dokumen: Google Drive menolak pembuatan shortcut. Detail: ' + message);
}

function getUniqueFileName_(folder, requestedName) {
  const cleanName = String(requestedName || 'Dokumen Arsip.pdf').trim() || 'Dokumen Arsip.pdf';
  if (!fileNameExistsInFolder_(folder, cleanName)) return cleanName;

  const match = cleanName.match(/^(.*?)(\.[^.]+)?$/);
  const base = (match && match[1] ? match[1] : cleanName).trim();
  const ext = match && match[2] ? match[2] : '';
  let counter = 2;
  let candidate = base + ' (' + counter + ')' + ext;
  while (fileNameExistsInFolder_(folder, candidate)) {
    counter++;
    candidate = base + ' (' + counter + ')' + ext;
  }
  return candidate;
}

function fileNameExistsInFolder_(folder, name) {
  const folderId = folder.getId();
  let pageToken = null;
  do {
    let result;
    try {
      result = Drive.Files.list({
        q: "'" + folderId + "' in parents and name = '" + name.replace(/'/g, "\\'") + "' and trashed = false",
        fields: "nextPageToken, files(id)",
        pageSize: 10,
        pageToken: pageToken,
        supportsAllDrives: true,
        includeItemsFromAllDrives: true
      });
    } catch (e) {
      break;
    }
    if (result.files && result.files.length > 0) return true;
    pageToken = result.nextPageToken;
  } while (pageToken);
  return false;
}

function getFolderByPickerId_(folderId) {
  if (!folderId || String(folderId) === 'root') return DriveApp.getRootFolder();
  return DriveApp.getFolderById(cleanId_(folderId));
}

function normalizeDrivePageSize_(value, fallback, max) {
  const parsed = Number(value || fallback);
  if (!parsed || isNaN(parsed)) return fallback;
  return Math.max(10, Math.min(Math.floor(parsed), max || fallback));
}

function driveFileResourceToDto_(item) {
  item = item || {};
  const mimeType = item.mimeType || '';
  const id = item.id || '';
  const modified = item.modifiedTime ? new Date(item.modifiedTime).getTime() : 0;
  const created = item.createdTime ? new Date(item.createdTime).getTime() : 0;
  return {
    id: id,
    name: item.name || id,
    url: item.webViewLink || (mimeType === DRIVE_FOLDER_MIME_TYPE
      ? 'https://drive.google.com/drive/folders/' + id
      : 'https://drive.google.com/file/d/' + id + '/view'),
    downloadUrl: mimeType === DRIVE_FOLDER_MIME_TYPE ? '' : DriveService.getDownloadUrl(id, mimeType),
    mimeType: mimeType,
    size: Number(item.size || 0),
    created: created,
    updated: modified || created
  };
}

function getFirstParentFolder_(folder) {
  try {
    const parents = folder.getParents();
    if (!parents.hasNext()) return null;
    return DriveService.folderToDto(parents.next());
  } catch (error) {
    console.error('getFirstParentFolder_ failed: ' + error.message);
    return null;
  }
}

function chooseInboxSubfolder_(name) {
  if (/\.docx?$/i.test(name)) return 'DOCX';
  if (/\.pdf$/i.test(name)) return 'PDF';
  return 'Lainnya';
}

function dataUrlToBlob_(dataUrl, name, mimeType) {
  const parts = String(dataUrl).split(',');
  if (parts.length < 2 || !parts[1]) throw new Error('Format file upload tidak valid.');
  const header = parts[0];
  const contentType = mimeType || (header.match(/data:(.*?);base64/) || [])[1] || MimeType.BINARY;
  const bytes = Utilities.base64Decode(parts[1]);
  return Utilities.newBlob(bytes, contentType, name);
}


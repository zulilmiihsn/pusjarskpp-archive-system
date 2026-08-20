'use strict';

function wsGetOrCreateFolder_(parent, name, report) {
  const folders = parent.getFoldersByName(name);
  while (folders.hasNext()) {
    const folder = folders.next();
    if (!folder.isTrashed()) {
      wsPushReport_(report, 'found', 'Folder ditemukan: ' + name);
      return folder;
    }
  }
  const folder = withRetry_(() => parent.createFolder(name));
  // Invalidasi cache anak agar listing berikutnya melihat folder baru ini (C1).
  delete _wsFolderChildrenCache_[parent.getId()];
  wsPushReport_(report, 'created', 'Folder dibuat: ' + name);
  return folder;
}

function wsEnsureRestrictedSharing_(folder, report) {
  const folderName = folder && folder.getName ? folder.getName() : 'folder arsip';
  try {
    if (folder.getSharingAccess() === DriveApp.Access.PRIVATE) {
      wsPushReport_(report, 'found', 'Akses Restricted sudah aktif: ' + folderName);
      return folder;
    }
    folder.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
    wsPushReport_(report, 'updated', 'Link sharing dinonaktifkan (Restricted): ' + folderName);
    return folder;
  } catch (error) {
    throw new Error(
      'Gagal mengatur akses Restricted untuk "' + folderName +
      '". Periksa kebijakan berbagi Google Workspace. Detail: ' +
      (error && error.message ? error.message : String(error))
    );
  }
}

function wsFindOrCreateChildFolder_(parent, candidates, createName, report) {
  const found = wsFindChildFolder_(parent, candidates);
  if (found) {
    wsPushReport_(report, 'found', 'Folder ditemukan: ' + found.getName());
    return found;
  }
  const folder = withRetry_(function() { return parent.createFolder(createName); });
  if (typeof _wsFolderChildrenCache_ !== 'undefined') {
    delete _wsFolderChildrenCache_[parent.getId()];
  }
  wsPushReport_(report, 'created', 'Folder dibuat: ' + createName);
  return folder;
}

function wsFindChildFolder_(parent, candidates) {
  const items = wsListChildFolders_(parent);
  
  // exact match first (candidate order priority)
  for (let j = 0; j < candidates.length; j++) {
    for (let i = 0; i < items.length; i++) {
      if (items[i].getName() === candidates[j]) return items[i];
    }
  }
  // fuzzy match second (candidate order priority)
  for (let j = 0; j < candidates.length; j++) {
    const candNorm = wsNormalize_(candidates[j]);
    if (!candNorm) continue;
    for (let i = 0; i < items.length; i++) {
      const normalized = wsNormalize_(items[i].getName());
      if (normalized.indexOf(candNorm) >= 0) return items[i];
    }
  }
  
  return null;
}

const _wsFolderChildrenCache_ = {};

// Ambil SEMUA child (lintas halaman) untuk satu query Drive.Files.list. Memakai
// withRetry_ untuk error transient; pada kegagalan persisten ia MELEMPAR
// (fail-closed) — bukan mengembalikan list terpotong — supaya pemanggil tidak
// salah menyimpulkan "tidak ada" lalu membuat folder/berkas duplikat.
function listAllChildren_(q, fields) {
  const files = [];
  let pageToken = null;
  do {
    const result = withRetry_(function () {
      return Drive.Files.list({
        q: q,
        fields: fields || 'nextPageToken, files(id, name)',
        pageSize: 1000,
        pageToken: pageToken,
        supportsAllDrives: true,
        includeItemsFromAllDrives: true
      });
    });
    const items = result.files || [];
    for (let i = 0; i < items.length; i++) files.push(items[i]);
    pageToken = result.nextPageToken;
  } while (pageToken);
  return files;
}

function wsListChildFolders_(parent) {
  const parentId = parent.getId();
  if (_wsFolderChildrenCache_[parentId]) {
    return _wsFolderChildrenCache_[parentId];
  }

  // Bila listAllChildren_ melempar (error persisten), fungsi keluar sebelum baris
  // cache di bawah — jadi list terpotong tak pernah tersimpan (anti cache-poison).
  const sortedFolders = listAllChildren_(
    "'" + parentId + "' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false",
    'nextPageToken, files(id, name)'
  ).map(function (it) {
    return { folder: DriveApp.getFolderById(it.id), name: it.name };
  }).sort(function (a, b) {
    return a.name.localeCompare(b.name, 'id', { numeric: true });
  }).map(function (item) {
    return item.folder;
  });

  _wsFolderChildrenCache_[parentId] = sortedFolders;
  return sortedFolders;
}

function wsGetFileByNameAndMime_(folder, name, mimeType) {
  let q = "'" + folder.getId() + "' in parents and name = '" + name.replace(/'/g, "\\'") + "' and trashed = false";
  if (mimeType) {
    q += " and mimeType = '" + mimeType + "'";
  }
  const files = listAllChildren_(q, 'nextPageToken, files(id)');
  return files.length ? DriveApp.getFileById(files[0].id) : null;
}

function wsFindFirstArchiveSheet_(folder) {
  const files = listAllChildren_(
    "'" + folder.getId() + "' in parents and mimeType = '" + MimeType.GOOGLE_SHEETS + "' and trashed = false",
    'nextPageToken, files(id, name)'
  );
  for (let i = 0; i < files.length; i++) {
    const name = wsNormalize_(files[i].name);
    if (name.indexOf('daftar') >= 0 && name.indexOf('arsip') >= 0) {
      return DriveApp.getFileById(files[i].id);
    }
  }
  return null;
}

function wsFindFirstOfficeSpreadsheet_(folder) {
  const files = listAllChildren_(
    "'" + folder.getId() + "' in parents and (mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'     or mimeType = 'application/vnd.ms-excel' or name contains '.xls' or name contains '.xlsx') and trashed = false",
    'nextPageToken, files(id, name, mimeType)'
  );
  for (let i = 0; i < files.length; i++) {
    const name = files[i].name;
    const normalizedName = wsNormalize_(name);
    const mimeType = files[i].mimeType;
    if (normalizedName.indexOf('daftar') >= 0 && normalizedName.indexOf('arsip') >= 0 &&
      (/\.xlsx?$/i.test(name) || mimeType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || mimeType       === 'application/vnd.ms-excel')) {
      return DriveApp.getFileById(files[i].id);
    }
  }
  return null;
}

function wsMoveFileToFolder_(fileId, folder) {
  const file = DriveApp.getFileById(fileId);
  file.moveTo(folder);
}

function wsNormalize_(value) {
  return String(value || '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
}

function wsBuildPath_(rootName, parts) {
  return [rootName].concat(parts).join(' > ');
}

function wsPushReport_(report, status, label) {
  if (!report) return;
  report.push({
    status: status,
    label: label
  });
}


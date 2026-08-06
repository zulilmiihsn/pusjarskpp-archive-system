'use strict';

/**
 * Given an activity's target_folder_id (which lives inside 1. Persuratan > Tahun XXXX),
 * resolve the corresponding mirror folder inside 2. Dokumen > Tahun XXXX > same activity folder.
 * Returns the mirror activity folder, or null if the structure cannot be resolved.
 */
function resolveDokumenMirrorForActivity_(activityTargetFolderId) {
  try {
    // activityTargetFolderId points to e.g. "Folder 2 (Latsar CPNS)" inside "1. Persuratan > Tahun 2026"
    const activityFolder = DriveApp.getFolderById(activityTargetFolderId);
    const activityFolderName = activityFolder.getName();

    // Go up: activityFolder -> tahunFolder -> persuratanFolder -> naskahDinasFolder
    const tahunParents = activityFolder.getParents();
    if (!tahunParents.hasNext()) return null;
    const tahunFolder = tahunParents.next();
    const tahunFolderName = tahunFolder.getName();

    const persuratanParents = tahunFolder.getParents();
    if (!persuratanParents.hasNext()) return null;
    const persuratanFolder = persuratanParents.next();

    const naskahDinasParents = persuratanFolder.getParents();
    if (!naskahDinasParents.hasNext()) return null;
    const naskahDinasFolder = naskahDinasParents.next();

    // Navigate into 2. Dokumen > Tahun XXXX > same activity folder name
    const dokumenFolder = DriveService.getOrCreateChildFolder(naskahDinasFolder, '2. Dokumen');
    const dokumenTahunFolder = DriveService.getOrCreateChildFolder(dokumenFolder, tahunFolderName);
    return DriveService.getOrCreateChildFolder(dokumenTahunFolder, activityFolderName);
  } catch (e) {
    console.error('resolveDokumenMirrorForActivity_ failed: ' + e.message);
    return null;
  }
}

/**
 * Ensure a sub-activity folder also exists inside the 2. Dokumen mirror tree.
 * parentFolderName is for grouped activities (e.g. PKN/PKA/PKP inside Kepemimpinan).
 */
function ensureDokumenMirrorSubActivity_(activityTargetFolderId, subActivityName, parentFolderName) {
  try {
    const mirrorActivityFolder = resolveDokumenMirrorForActivity_(activityTargetFolderId);
    if (!mirrorActivityFolder) return null;

    if (parentFolderName) {
      const groupFolder = DriveService.getOrCreateChildFolder(mirrorActivityFolder, parentFolderName);
      return DriveService.getOrCreateChildFolder(groupFolder, subActivityName);
    }
    return DriveService.getOrCreateChildFolder(mirrorActivityFolder, subActivityName);
  } catch (e) {
    console.error('ensureDokumenMirrorSubActivity_ failed: ' + e.message);
    return null;
  }
}

/**
 * Rename a sub-activity folder inside the 2. Dokumen mirror tree.
 */
function renameDokumenMirrorSubActivity_(activityTargetFolderId, oldName, newName, parentFolderName) {
  try {
    const mirrorActivityFolder = resolveDokumenMirrorForActivity_(activityTargetFolderId);
    if (!mirrorActivityFolder) return;

    let searchParent = mirrorActivityFolder;
    if (parentFolderName) {
      const groupFolders = mirrorActivityFolder.getFoldersByName(parentFolderName);
      if (groupFolders.hasNext()) {
        searchParent = groupFolders.next();
      } else {
        return; // group folder doesn't exist, nothing to rename
      }
    }

    const folders = searchParent.getFoldersByName(oldName);
    if (folders.hasNext()) {
      folders.next().setName(newName);
    }
  } catch (e) {
    console.error('renameDokumenMirrorSubActivity_ failed: ' + e.message);
  }
}

/**
 * Resolve the "2. Dokumen" mirror folder for any folder under "1. Persuratan".
 * Read-only: returns null if mirror path doesn't exist.
 */
function resolveMirrorForFolder_(folderId) {
  try {
    const folder = DriveApp.getFolderById(folderId);
    const pathParts = [];
    let current = folder;
    let naskahDinasFolder = null;

    while (current) {
      const parents = current.getParents();
      if (!parents.hasNext()) break;
      const parent = parents.next();
      const pName = parent.getName().toLowerCase();

      if (pName.indexOf('persuratan') >= 0 || pName.indexOf('1. persuratan') >= 0) {
        const ndParents = parent.getParents();
        if (ndParents.hasNext()) {
          naskahDinasFolder = ndParents.next();
        }
        pathParts.unshift(current.getName());
        break;
      }

      pathParts.unshift(current.getName());
      current = parent;
    }

    if (!naskahDinasFolder || pathParts.length === 0) return null;

    const dokumenFolders = naskahDinasFolder.getFoldersByName('2. Dokumen');
    if (!dokumenFolders.hasNext()) return null;
    let mirror = dokumenFolders.next();

    for (let i = 0; i < pathParts.length; i++) {
      const subFolders = mirror.getFoldersByName(pathParts[i]);
      if (!subFolders.hasNext()) return null;
      mirror = subFolders.next();
    }
    return mirror;
  } catch (e) {
    console.error('resolveMirrorForFolder_ failed: ' + e.message);
    return null;
  }
}

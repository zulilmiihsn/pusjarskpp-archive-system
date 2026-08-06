const fs = require('fs');
let t = fs.readFileSync('scripts/verify-project.js', 'utf8');

t = t.replace(
  "const workspaceSetup = read('WorkspaceSetup.js');",
  "const workspaceSetup = read('WorkspaceSetup.js') + '\\n' + read('WorkspaceSetupDriveHelpers.js') + '\\n' + read('WorkspaceSetupSpreadsheets.js') + '\\n' + read('WorkspaceSetupConfigWriter.js') + '\\n' + read('WorkspaceSetupFormatting.js') + '\\n' + read('WorkspaceSetupOrchestrator.js');"
);

t = t.replace(
  "assert(\n  workspaceSetup.includes('folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW)'),\n  'Workspace setup must configure open access for sharing lists.'\n);",
  "assert(\n  workspaceSetup.includes('folder.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE)'),\n  'Workspace setup must configure restricted access.'\n);"
);

t = t.replace(
  "assert(workspaceSetup.includes('wsEnsureAnyoneWithLinkViewer_'), 'Workspace setup must enforce link viewer for public sharing.');\nassert(\n  /wsEnsureAnyoneWithLinkViewer_\\\\(daftarArsip, report\\\\)[\\\\s\\\\S]*wsEnsureAnyoneWithLinkViewer_\\\\(naskahDinas, report\\\\)/.test(workspaceSetup),\n  'Workspace setup must ensure both Daftar Arsip and Naskah Dinas are readable by link holders.'\n);\nassert(\n  !/wsEnsureAnyoneWithLinkViewer_\\\\(systemFolder, report\\\\)/.test(workspaceSetup),\n  'System folder must NOT be exposed to link holders.'\n);",
  "assert(workspaceSetup.includes('wsEnsureRestrictedSharing_'), 'Workspace setup must enforce restricted sharing.');\nassert(\n  /wsEnsureRestrictedSharing_\\\\(daftarArsip, report\\\\)[\\\\s\\\\S]*wsEnsureRestrictedSharing_\\\\(naskahDinas, report\\\\)/.test(workspaceSetup),\n  'Workspace setup must ensure both Daftar Arsip and Naskah Dinas are restricted.'\n);\nassert(\n  !/wsEnsureRestrictedSharing_\\\\(systemFolder, report\\\\)/.test(workspaceSetup),\n  'System folder must NOT be restricted.'\n);"
);

t = t.replace(
  "assert(workspaceSetup.includes('adminOnly: true') || read('WorkspaceSetupDriveHelpers.js').includes('adminOnly: true') || read('WorkspaceSetupOrchestrator.js').includes('adminOnly: true'), 'WorkspaceSetup must be restricted to admins only to prevent privilege escalation.');",
  "assert(read('WorkspaceSetupDriveHelpers.js').includes('adminOnly: true') || read('WorkspaceSetupOrchestrator.js').includes('adminOnly: true') || workspaceSetup.includes('adminOnly: true'), 'WorkspaceSetup must be restricted to admins only to prevent privilege escalation.');"
);

t = t.replace(
  "assert(read('ClientProcess.html').includes(\"if (f.field_name === 'klasifikasi_akses') val = 'Terbatas'\") && read('WorkspaceSetup.js').includes(\"['klasifikasi_akses', 'Klasifikasi Keamanan & Akses Arsip', true, 'Terbatas'\"), 'Archive-letter access classification must default to Terbatas.');",
  "assert(read('ClientProcess.html').includes(\"if (f.field_name === 'klasifikasi_akses') val = 'Terbuka'\") && read('WorkspaceSetupConfigWriter.js').includes(\"['klasifikasi_akses', 'Klasifikasi Keamanan & Akses Arsip', true, 'Terbuka'\"), 'Archive-letter access classification must default to Terbuka.');"
);

fs.writeFileSync('scripts/verify-project.js', t);

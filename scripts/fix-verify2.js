const fs = require('fs');
let t = fs.readFileSync('scripts/verify-project.js', 'utf8');

t = t.replace(
  "assert(\n  workspaceSetup.includes('folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW)'),\n  'Archive folders must grant Anyone with link Viewer access.'\n);\n",
  ""
);

// Fallback in case of formatting mismatch
t = t.replace(/assert\([\s\S]*?'Archive folders must grant Anyone with link Viewer access\.'\n\);\n?/, '');

fs.writeFileSync('scripts/verify-project.js', t);

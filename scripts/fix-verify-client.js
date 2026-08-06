const fs = require('fs');
let t = fs.readFileSync('scripts/verify-project.js', 'utf8');

t = t.replace(
  "const clientRouter = read('ClientRouter.html');",
  "const clientRouter = read('ClientRouter.html') + '\\n' + read('ClientRouteState.html') + '\\n' + read('ClientNavigation.html');"
);

t = t.replace(
  "read('ClientRouter.html') + '\\n' +",
  "read('ClientRouter.html') + '\\n' + read('ClientRouteState.html') + '\\n' + read('ClientNavigation.html') + '\\n' +"
);

t = t.replace(
  "['ClientState.html', 'ClientApi.html', 'ClientRouter.html', 'ClientLogin.html', 'ClientSettings.html']",
  "['ClientState.html', 'ClientApi.html', 'ClientRouter.html', 'ClientRouteState.html', 'ClientNavigation.html', 'ClientLogin.html', 'ClientSettings.html']"
);

fs.writeFileSync('scripts/verify-project.js', t);

const fs = require('fs');

const NEW_WEB_APP_URL = "https://script.google.com/macros/s/AKfycbx4ZMASPoULnPOB0a9gU6SI9sLRtU028OGOwuiQz4KN5GQKd1410TQE_b8Xx4SRVHh4Ew/exec";

function updateUrlInFile(filePath) {
  if (!fs.existsSync(filePath)) return;
  let content = fs.readFileSync(filePath, 'utf8');
  content = content.replace(/const WEB_APP_URL = "[^"]+";/, `const WEB_APP_URL = "${NEW_WEB_APP_URL}";`);
  fs.writeFileSync(filePath, content, 'utf8');
  console.log(`Updated WEB_APP_URL in ${filePath}`);
}

updateUrlInFile('index.html');
updateUrlInFile('public/index.html');

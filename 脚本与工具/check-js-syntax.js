const fs = require('fs');
const path = 'D:\\编程开发\\开源项目\\多维表分析软件\\MultiTableAddin\\src\\MultiTableAddin\\Html\\index.html';
const html = fs.readFileSync(path, 'utf8');
const m = html.match(/<script[^>]*>([\s\S]*?)<\/script>/);
if (!m) { console.log('no script'); process.exit(1); }
try {
  new Function(m[1]);
  console.log('JS syntax OK');
} catch (e) {
  console.error('JS SYNTAX ERROR:', e.message);
  process.exit(1);
}

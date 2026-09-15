import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const args = process.argv.slice(2);
if (args.length > 1 || args.length === 1 && args[0] !== '--check') throw new Error('Usage: node build.mjs [--check]');
const root = path.dirname(fileURLToPath(import.meta.url));
const sources = ['import-docx.js', 'import-pdf.js', 'interpretation.js', 'reader.js', 'app.js'];
const read = name => {
  const text = fs.readFileSync(path.join(root, 'src', name), 'utf8');
  if (text.includes('\r')) throw new Error(`${name}: source must use LF line endings`);
  return text;
};
const template = read('index.template.html');
const marker = '{{SCENARIO_SCRIPT}}';
if (template.split(marker).length !== 2) throw new Error('Expected one script insertion point');
const script = sources.map(name => `\n    // Source: src/${name}\n${read(name).trimEnd()}\n`).join('');
const result = template.replace(marker, () => script);
const target = path.join(root, 'index.html');
if (args[0] === '--check') {
  if (fs.readFileSync(target, 'utf8') !== result) {
    console.error('index.html differs from src/. Run node build.mjs.');
    process.exitCode = 1;
  } else console.log(`Generated HTML matches ${sources.length} source files.`);
} else {
  fs.writeFileSync(target, result, 'utf8');
  console.log('Built index.html. Deployment requires this HTML only.');
}

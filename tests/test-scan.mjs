import assert from 'node:assert/strict';
import fs from 'node:fs';
import { scanFile } from '../src/scan.js';

const cases = [
  { name: 'good.docx', score: 100, ids: [] },
  { name: 'bad.docx', score: 63, ids: ['docx.image.alt_missing', 'docx.table.header_missing', 'docx.link.ambiguous', 'docx.headings.missing', 'docx.title.missing', 'docx.text.tiny'] },
  { name: 'good.pptx', score: 100, ids: [] },
  { name: 'bad.pptx', score: 69, ids: ['pptx.title.missing', 'pptx.image.alt_poor', 'pptx.table.header_missing', 'pptx.hidden_slides', 'pptx.text.tiny'] },
];

for (const item of cases) {
  const data = fs.readFileSync(`../mvp/tests/fixtures/${item.name}`);
  const report = await scanFile(new File([data], item.name));
  assert.equal(report.score, item.score, `${item.name} score`);
  assert.deepEqual(report.issues.map((i) => i.id), item.ids, `${item.name} issue ids`);
  console.log(`ok ${item.name} score=${report.score} issues=${report.issues.length}`);
}

await assert.rejects(
  () => scanFile(new File([new Uint8Array([1, 2, 3])], 'broken.docx')),
  /ZIP container is damaged/
);
await assert.rejects(
  () => scanFile(new File([new Uint8Array([1])], 'file.pdf')),
  /Only \.docx and \.pptx/
);
console.log('ok error handling');

import { readdir, stat } from "node:fs/promises";
import { join } from "node:path";

const limits = {
  pptx: 800 * 1024,
  unpacked: 1024 * 1024,
  zip: 350 * 1024,
};

async function treeSize(path) {
  let bytes = 0;
  for (const entry of await readdir(path, { withFileTypes: true })) {
    const child = join(path, entry.name);
    if (entry.isDirectory()) bytes += await treeSize(child);
    else if (entry.isFile()) bytes += (await stat(child)).size;
  }
  return bytes;
}

const sizes = {
  pptx: (await stat("build/chrome/dist/pptx.js")).size,
  unpacked: await treeSize("build/chrome"),
  zip: (await stat("hover-file-preview-chrome.zip")).size,
};

console.log(`Chrome pptx.js: ${sizes.pptx} bytes`);
console.log(`Unpacked Chrome tree: ${sizes.unpacked} bytes`);
console.log(`Chrome package ZIP: ${sizes.zip} bytes`);

const failures = Object.entries(limits).filter(([name, limit]) => sizes[name] > limit);
if (failures.length) {
  for (const [name, limit] of failures) {
    console.error(`${name} exceeds its ${limit}-byte budget`);
  }
  process.exitCode = 1;
}

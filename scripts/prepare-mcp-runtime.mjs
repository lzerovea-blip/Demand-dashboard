import { createHash } from 'node:crypto';
import { cp, mkdir, mkdtemp, readFile, rm, writeFile, chmod } from 'node:fs/promises';
import { homedir, tmpdir } from 'node:os';
import { join } from 'node:path';
import { execFileSync } from 'node:child_process';
import JSZip from 'jszip';
const version = '22.21.1';
const target = process.argv[2];
if (!['win32-x64', 'darwin-arm64'].includes(target)) throw new Error('Unsupported MCP runtime target');
const name = `node-v${version}-${target === 'win32-x64' ? 'win-x64.zip' : 'darwin-arm64.tar.gz'}`;
const base = `https://nodejs.org/dist/v${version}/`;
async function download(url) { const response = await fetch(url, { signal: AbortSignal.timeout(60000) }); if (!response.ok) throw new Error(`Runtime download failed: ${response.status}`); return Buffer.from(await response.arrayBuffer()); }
const manifest = (await download(`${base}SHASUMS256.txt`)).toString();
const expected = manifest.split('\n').map(line => line.trim().split(/\s+/)).find(parts => parts[1] === name)?.[0];
if (!expected) throw new Error('Runtime checksum missing');
const cache = join(homedir(), '.cache', 'solution-requirements-mcp');
await mkdir(cache, { recursive: true });
const cached = join(cache, name);
const digest = value => createHash('sha256').update(value).digest('hex');
let archive = await readFile(cached).catch(() => null);
if (!archive || digest(archive) !== expected) archive = await download(`${base}${name}`);
if (digest(archive) !== expected) throw new Error('Runtime checksum mismatch');
await writeFile(cached, archive);
const output = join('build-resources', 'mcp-runtime', target);
await mkdir(output, { recursive: true });
const binary = join(output, target === 'win32-x64' ? 'node.exe' : 'node');
if (target === 'win32-x64') {
  const zip = await JSZip.loadAsync(archive);
  await writeFile(binary, await zip.file(`node-v${version}-win-x64/node.exe`).async('nodebuffer'));
  await writeFile(join(output, 'LICENSE'), await zip.file(`node-v${version}-win-x64/LICENSE`).async('nodebuffer'));
} else {
  const temporary = await mkdtemp(join(tmpdir(), 'mcp-runtime-'));
  try {
    execFileSync('tar', ['-xzf', cached, '-C', temporary]);
    await cp(join(temporary, `node-v${version}-darwin-arm64/bin/node`), binary);
    await cp(join(temporary, `node-v${version}-darwin-arm64/LICENSE`), join(output, 'LICENSE'));
    await chmod(binary, 0o755);
  } finally { await rm(temporary, { recursive: true, force: true }); }
}
await writeFile(join(output, 'runtime.json'), JSON.stringify({ version, target, url: `${base}${name}`, archiveSha256: expected, binarySha256: digest(await readFile(binary)) }, null, 2));
console.log(`Verified MCP runtime ${version}: ${target}`);

/**
 * 跨平台启动 Python 解析器：Windows 通常没有 python3 命令（exit 9009），依次尝试 py/python/python3。
 * 端口由 prep-dev.mjs 写入 data/.parser_dev_port（pnpm dev 会先跑 prep）；单独运行本脚本时回退为 8000。
 */
import { spawnSync } from 'node:child_process';
import fs from 'node:fs';
import { platform } from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.join(__dirname, '..');
const cwd = path.join(root, 'apps', 'parser');
const portFile = path.join(root, 'data', '.parser_dev_port');

if (!process.env.PARSER_PORT && fs.existsSync(portFile)) {
  process.env.PARSER_PORT = fs.readFileSync(portFile, 'utf8').trim() || '8000';
}

const attempts =
  platform() === 'win32'
    ? [
        ['py', ['-3', 'main.py']],
        ['python', ['main.py']],
        ['python3', ['main.py']],
      ]
    : [
        ['python3', ['main.py']],
        ['python', ['main.py']],
      ];

const win = platform() === 'win32';

for (const [cmd, args] of attempts) {
  const r = spawnSync(cmd, args, { cwd, stdio: 'inherit', shell: true, env: process.env });
  if (r.error) {
    if (r.error.code === 'ENOENT') continue;
    console.error(r.error.message);
    process.exit(1);
  }
  if ((win && r.status === 9009) || (!win && r.status === 127)) continue;
  process.exit(r.status ?? 1);
}

console.error(
  '未找到可用的 Python。Windows 请安装 Python 3 并勾选 “Add to PATH”，或使用：py -3 apps/parser/main.py'
);
process.exit(1);

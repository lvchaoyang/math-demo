/**
 * 在 concurrently 启动各服务前执行：检测解析器端口，并同步 apps/api/.env 中的 PARSER_URL。
 * 避免 8000 被占用（如 CLodop）时 API 仍指向 8000；也保证 API 首次启动即读到正确配置。
 */
import fs from 'node:fs';
import net from 'node:net';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.join(__dirname, '..');
const dataDir = path.join(root, 'data');
const apiEnvPath = path.join(root, 'apps', 'api', '.env');
const portFile = path.join(dataDir, '.parser_dev_port');

function canBindPort(port) {
  return new Promise((resolve) => {
    const server = net.createServer();
    server.unref();
    server.once('error', () => resolve(false));
    server.listen(port, '0.0.0.0', () => {
      server.close(() => resolve(true));
    });
  });
}

function mergeParserUrl(wantPort) {
  const wantUrl = `http://localhost:${wantPort}`;
  let body = '';
  if (fs.existsSync(apiEnvPath)) {
    body = fs.readFileSync(apiEnvPath, 'utf8');
  }
  if (/^\s*PARSER_URL\s*=/m.test(body)) {
    if (/PARSER_URL\s*=\s*http:\/\/localhost:\d+\b/.test(body)) {
      const next = body.replace(/PARSER_URL\s*=\s*http:\/\/localhost:\d+\b/g, `PARSER_URL=${wantUrl}`);
      if (next !== body) fs.writeFileSync(apiEnvPath, next, 'utf8');
    }
    return;
  }
  fs.appendFileSync(apiEnvPath, `${body && !body.endsWith('\n') ? '\n' : ''}PARSER_URL=${wantUrl}\n`, 'utf8');
}

const free8000 = await canBindPort(8000);
const chosen = free8000 ? 8000 : 8001;

fs.mkdirSync(dataDir, { recursive: true });
fs.writeFileSync(portFile, String(chosen), 'utf8');

if (chosen === 8001) {
  mergeParserUrl(8001);
  console.warn(
    '[math-demo] 端口 8000 已被占用（常见于 CLodop 等），解析服务将使用 8001；已同步 apps/api/.env 中的 PARSER_URL。\n'
  );
} else {
  if (fs.existsSync(apiEnvPath)) {
    const body = fs.readFileSync(apiEnvPath, 'utf8');
    if (/PARSER_URL\s*=\s*http:\/\/localhost:8001\b/.test(body)) {
      fs.writeFileSync(
        apiEnvPath,
        body.replace(/PARSER_URL\s*=\s*http:\/\/localhost:8001\b/g, 'PARSER_URL=http://localhost:8000'),
        'utf8'
      );
      console.warn('[math-demo] 已把 apps/api/.env 中的 PARSER_URL 改回 http://localhost:8000\n');
    }
  }
}

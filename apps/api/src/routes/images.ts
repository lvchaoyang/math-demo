import { Router } from 'express';
import path from 'path';
import fs from 'fs';
import { spawn } from 'child_process';
import { fileURLToPath } from 'url';
import { logger } from '../utils/logger.js';

const router = Router();
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const PROJECT_ROOT = path.resolve(__dirname, '../../../../');

const WMF_EXTENSIONS = ['.wmf', '.emf'];
const CACHE_DIR = path.join(PROJECT_ROOT, 'data', 'image_cache');

const WMF2SVG_CMD = process.env.WMF2SVG_PATH || 'wmf2svg';
const RSVG_CONVERT_CMD = process.env.RSVG_CONVERT_PATH || 'rsvg-convert';
const MAGICK_CMD = process.env.MAGICK_PATH || 'magick';
const INKSCAPE_CMD = process.env.INKSCAPE_PATH || 'inkscape';
const CONVERT_TIMEOUT_MS = Number(process.env.IMAGE_CONVERT_TIMEOUT_MS || 12000);

type CommandResult = {
  code: number | null;
  stderr: string;
  error?: string;
  timedOut: boolean;
};

function runCommandWithTimeout(cmd: string, args: string[], timeoutMs: number): Promise<CommandResult> {
  return new Promise((resolve) => {
    const child = spawn(cmd, args);
    let stderr = '';
    let settled = false;
    const timer = setTimeout(() => {
      if (settled) return;
      settled = true;
      child.kill('SIGKILL');
      resolve({ code: null, stderr, timedOut: true, error: `timeout after ${timeoutMs}ms` });
    }, timeoutMs);

    child.stderr.on('data', (data) => {
      stderr += data.toString();
    });

    child.on('close', (code) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      resolve({ code, stderr, timedOut: false });
    });

    child.on('error', (err) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      resolve({ code: null, stderr, timedOut: false, error: String(err) });
    });
  });
}

function ensureCacheDir() {
  if (!fs.existsSync(CACHE_DIR)) {
    fs.mkdirSync(CACHE_DIR, { recursive: true });
  }
}

function isWmfFile(filename: string): boolean {
  const ext = path.extname(filename).toLowerCase();
  return WMF_EXTENSIONS.includes(ext);
}

async function convertWmfToPng(wmfPath: string, pngPath: string): Promise<{ success: boolean; error?: string }> {
  ensureCacheDir();
  const tempSvgPath = pngPath.replace('.png', '.svg');
  const wmf2svg = await runCommandWithTimeout(WMF2SVG_CMD, ['-o', tempSvgPath, wmfPath], CONVERT_TIMEOUT_MS);
  if (wmf2svg.code !== 0 || !fs.existsSync(tempSvgPath)) {
    return { success: false, error: `wmf2svg failed: ${wmf2svg.error || wmf2svg.stderr}` };
  }
  const rsvg = await runCommandWithTimeout(RSVG_CONVERT_CMD, [tempSvgPath, '-o', pngPath], CONVERT_TIMEOUT_MS);
  if (fs.existsSync(tempSvgPath)) {
    fs.unlinkSync(tempSvgPath);
  }
  if (rsvg.code === 0 && fs.existsSync(pngPath)) {
    return { success: true };
  }
  return { success: false, error: `rsvg-convert failed: ${rsvg.error || rsvg.stderr}` };
}

async function convertWithImageMagick(wmfPath: string, pngPath: string): Promise<{ success: boolean; error?: string }> {
  const result = await runCommandWithTimeout(MAGICK_CMD, ['-density', '600', wmfPath, pngPath], CONVERT_TIMEOUT_MS);
  if (result.code === 0 && fs.existsSync(pngPath)) {
    return { success: true };
  }
  return { success: false, error: result.error || result.stderr };
}

async function convertWithInkscape(wmfPath: string, pngPath: string): Promise<{ success: boolean; error?: string }> {
  const result = await runCommandWithTimeout(INKSCAPE_CMD, [
    wmfPath,
    '--export-type=png',
    '--export-dpi=600',
    '--export-filename=' + pngPath
  ], CONVERT_TIMEOUT_MS);
  if (result.code === 0 && fs.existsSync(pngPath)) {
    return { success: true };
  }
  return { success: false, error: result.error || result.stderr };
}

type ConversionResult = {
  success: boolean;
  method: string;
  outputPath: string;
  error?: string;
  size?: number;
};

async function convertWmfBestEffort(wmfPath: string, finalPngPath: string): Promise<ConversionResult> {
  ensureCacheDir();
  const candidates = [
    { method: 'inkscape', out: finalPngPath.replace(/\.png$/, '.inkscape.png'), run: convertWithInkscape },
    { method: 'wmf2svg', out: finalPngPath.replace(/\.png$/, '.wmf2svg.png'), run: convertWmfToPng },
    { method: 'magick', out: finalPngPath.replace(/\.png$/, '.magick.png'), run: convertWithImageMagick }
  ];

  const ok: ConversionResult[] = [];
  const errors: string[] = [];

  for (const c of candidates) {
    try {
      const res = await c.run(wmfPath, c.out);
      if (res.success && fs.existsSync(c.out)) {
        const stat = fs.statSync(c.out);
        ok.push({ success: true, method: c.method, outputPath: c.out, size: stat.size });
      } else {
        errors.push(`${c.method}: ${res.error || 'failed'}`);
      }
    } catch (e) {
      errors.push(`${c.method}: ${String(e)}`);
    }
  }

  if (ok.length === 0) {
    return { success: false, method: 'none', outputPath: finalPngPath, error: errors.join(' | ') };
  }

  // 经验规则：文件体积更大的渲染通常细节保留更好
  ok.sort((a, b) => (b.size || 0) - (a.size || 0));
  const winner = ok[0];
  fs.copyFileSync(winner.outputPath, finalPngPath);

  for (const c of ok) {
    if (c.outputPath !== winner.outputPath && fs.existsSync(c.outputPath)) {
      fs.unlinkSync(c.outputPath);
    }
  }

  return { success: true, method: winner.method, outputPath: finalPngPath, size: winner.size };
}

router.get('/:fileId/:imageName(*)', async (req, res) => {
  try {
    const { fileId, imageName } = req.params;
    
    // 支持带 media/ 前缀的路径
    const imageDir = path.join(PROJECT_ROOT, 'data', 'images', fileId);
    let imagePath = path.join(imageDir, imageName);
    
    if (!fs.existsSync(imagePath)) {
      // 尝试添加 media 子目录
      if (!imageName.startsWith('media/')) {
        const prefixedPath = path.join(imageDir, 'media', imageName);
        if (fs.existsSync(prefixedPath)) {
          imagePath = prefixedPath;
        }
      } else {
        // 尝试去掉 media 前缀
        const prefixedPath = path.join(imageDir, imageName);
        if (fs.existsSync(prefixedPath)) {
          imagePath = prefixedPath;
        }
      }
    }
    
    if (!fs.existsSync(imagePath)) {
      // logger.error(`Image not found: ${imagePath}`);
      return res.status(404).json({
        success: false,
        message: '图片不存在'
      });
    }
    
    if (isWmfFile(imagePath)) {
      // 防止外部转换命令阻塞导致接口无响应
      res.setTimeout(CONVERT_TIMEOUT_MS * 4);
      ensureCacheDir();
      // 移除 media/ 前缀，避免文件路径问题
      const sanitizedImageName = imageName.replace(/^media\//, '');
      const cacheKey = `${fileId}_${sanitizedImageName}`.replace(/\.(wmf|emf)$/i, '.png');
      const cachedPngPath = path.join(CACHE_DIR, cacheKey);
      
      if (fs.existsSync(cachedPngPath)) {
        // logger.info(`Serving cached PNG: ${cachedPngPath}`);
        return res.sendFile(cachedPngPath);
      }
      
      const result = await convertWmfBestEffort(imagePath, cachedPngPath);
      
      if (result.success && fs.existsSync(cachedPngPath)) {
        logger.info(`WMF converted by ${result.method}: ${sanitizedImageName}`);
        // logger.info(`Serving converted PNG: ${cachedPngPath}`);
        return res.sendFile(cachedPngPath);
      }
      
      // logger.error(`WMF conversion failed for: ${imagePath}, error: ${result.error}`);
      return res.status(500).json({
        success: false,
        message: 'WMF 图片转换失败',
        error: result.error,
        originalFormat: 'wmf'
      });
    }
    
    res.sendFile(imagePath);
  } catch (error) {
    // logger.error(`Image serving error: ${error}`);
    res.status(500).json({
      success: false,
      message: '获取图片失败'
    });
  }
});

export { router as imagesRouter };

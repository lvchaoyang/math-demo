import { Router } from 'express';
import path from 'path';
import fs from 'fs';
import { spawn } from 'child_process';
import { fileURLToPath } from 'url';
import { logger } from '../utils/logger.js';
import os from 'os';

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
const SOFFICE_CMD = process.env.SOFFICE_PATH || 'soffice';
/** Windows：指向 tools/wmf-gdi-render 编译出的 WmfGdiRender.exe（GDI+ 栅格化 WMF/EMF）；未设置时尝试仓库默认输出路径 */
const CONVERT_TIMEOUT_MS = Number(process.env.IMAGE_CONVERT_TIMEOUT_MS || 12000);
const WMF_CACHE_VERSION = 'v14';

function getWmfGdiRenderExe(): string {
  const fromEnv = (process.env.WMF_GDI_RENDER_EXE || '').trim();
  if (fromEnv) return fromEnv;
  if (os.platform() !== 'win32') return '';
  const candidates = [
    path.join(PROJECT_ROOT, 'tools', 'wmf-gdi-render', 'bin', 'Release', 'net8.0-windows', 'WmfGdiRender.exe'),
    path.join(PROJECT_ROOT, 'tools', 'wmf-gdi-render', 'bin', 'Debug', 'net8.0-windows', 'WmfGdiRender.exe'),
  ];
  for (const p of candidates) {
    if (fs.existsSync(p)) return p;
  }
  return '';
}

function isGdiMetafileRendererAvailable(): boolean {
  const exe = getWmfGdiRenderExe();
  return os.platform() === 'win32' && exe.length > 0 && fs.existsSync(exe);
}

type CommandResult = {
  code: number | null;
  stdout: string;
  stderr: string;
  error?: string;
  timedOut: boolean;
};

function runCommandWithTimeout(cmd: string, args: string[], timeoutMs: number): Promise<CommandResult> {
  return new Promise((resolve) => {
    const child = spawn(cmd, args);
    let stdout = '';
    let stderr = '';
    let settled = false;
    const timer = setTimeout(() => {
      if (settled) return;
      settled = true;
      child.kill('SIGKILL');
      resolve({ code: null, stdout, stderr, timedOut: true, error: `timeout after ${timeoutMs}ms` });
    }, timeoutMs);

    child.stdout.on('data', (data) => {
      stdout += data.toString();
    });

    child.stderr.on('data', (data) => {
      stderr += data.toString();
    });

    child.on('close', (code) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      resolve({ code, stdout, stderr, timedOut: false });
    });

    child.on('error', (err) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      resolve({ code: null, stdout, stderr, timedOut: false, error: String(err) });
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

async function convertWmfToPng(
  wmfPath: string,
  pngPath: string,
  outlineText: boolean = false
): Promise<{ success: boolean; error?: string }> {
  ensureCacheDir();
  const tempSvgPath = pngPath.replace(/\.png$/i, '.svg');
  const wmf2svg = await runCommandWithTimeout(WMF2SVG_CMD, ['-o', tempSvgPath, wmfPath], CONVERT_TIMEOUT_MS);
  if (wmf2svg.code !== 0 || !fs.existsSync(tempSvgPath)) {
    return { success: false, error: `wmf2svg failed: ${wmf2svg.error || wmf2svg.stderr}` };
  }

  // 关键：SVG->PNG 优先用 Inkscape（rsvg-convert 常见字体度量导致叠字/挤压）
  const inkscape = await runCommandWithTimeout(
    INKSCAPE_CMD,
    [
      tempSvgPath,
      '--export-type=png',
      '--export-dpi=600',
      ...(outlineText ? ['--export-text-to-path'] : []),
      '--export-filename=' + pngPath
    ],
    CONVERT_TIMEOUT_MS
  );
  if (inkscape.code === 0 && fs.existsSync(pngPath)) {
    if (fs.existsSync(tempSvgPath)) fs.unlinkSync(tempSvgPath);
    return { success: true };
  }

  // 兼容旧版 inkscape：不支持 --export-text-to-path 时回退
  const inkscapeFallback = await runCommandWithTimeout(
    INKSCAPE_CMD,
    [
      tempSvgPath,
      '--export-type=png',
      '--export-dpi=600',
      '--export-filename=' + pngPath
    ],
    CONVERT_TIMEOUT_MS
  );
  if (inkscapeFallback.code === 0 && fs.existsSync(pngPath)) {
    if (fs.existsSync(tempSvgPath)) fs.unlinkSync(tempSvgPath);
    return { success: true };
  }

  // Inkscape 不可用/失败时，再回退 rsvg-convert
  const rsvg = await runCommandWithTimeout(RSVG_CONVERT_CMD, [tempSvgPath, '-o', pngPath], CONVERT_TIMEOUT_MS);
  if (rsvg.code === 0 && fs.existsSync(pngPath)) {
    if (fs.existsSync(tempSvgPath)) fs.unlinkSync(tempSvgPath);
    return { success: true };
  }
  if (fs.existsSync(tempSvgPath)) fs.unlinkSync(tempSvgPath);
  return {
    success: false,
    error: `svg->png failed: inkscape=${inkscape.error || inkscape.stderr} | inkscapeFallback=${inkscapeFallback.error || inkscapeFallback.stderr} | rsvg=${rsvg.error || rsvg.stderr}`
  };
}

async function convertWithImageMagick(wmfPath: string, pngPath: string): Promise<{ success: boolean; error?: string }> {
  const result = await runCommandWithTimeout(MAGICK_CMD, ['-density', '600', wmfPath, pngPath], CONVERT_TIMEOUT_MS);
  if (result.code === 0 && fs.existsSync(pngPath)) {
    return { success: true };
  }
  return { success: false, error: result.error || result.stderr };
}

function normalizeRasterDpi(value: unknown): number {
  const n = Number(value);
  if (!Number.isFinite(n) || n < 72) return 600;
  return Math.min(1200, Math.round(n));
}

async function convertWithGdiRender(
  wmfPath: string,
  pngPath: string,
  rasterDpi: number
): Promise<{ success: boolean; error?: string }> {
  if (!isGdiMetafileRendererAvailable()) {
    return {
      success: false,
      error:
        'GDI metafile renderer not available (need Windows, .NET 8 WmfGdiRender.exe under tools/wmf-gdi-render, or set WMF_GDI_RENDER_EXE)',
    };
  }
  const dpi = Math.min(1200, Math.max(72, Math.round(rasterDpi)));
  const exe = getWmfGdiRenderExe();
  const result = await runCommandWithTimeout(
    exe,
    [wmfPath, pngPath, '--dpi', String(dpi)],
    Math.max(CONVERT_TIMEOUT_MS * 3, 25000)
  );
  if (result.code === 0 && fs.existsSync(pngPath)) return { success: true };
  return { success: false, error: result.error || result.stderr || result.stdout };
}

async function convertWithSoffice(
  wmfPath: string,
  pngPath: string,
  rasterDpi: number = 600
): Promise<{ success: boolean; error?: string }> {
  // soffice 先转 PDF，再由 magick 转 PNG，通常比 SVG 渲染更不容易“符号乱码”
  ensureCacheDir();
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'wmf-soffice-'));
  try {
    const base = path.basename(wmfPath).replace(/\.(wmf|emf)$/i, '');
    const pdfPath = path.join(tmpDir, `${base}.pdf`);

    const toPdf = await runCommandWithTimeout(
      SOFFICE_CMD,
      ['--headless', '--nologo', '--nodefault', '--nolockcheck', '--convert-to', 'pdf', '--outdir', tmpDir, wmfPath],
      Math.max(CONVERT_TIMEOUT_MS * 4, 30000)
    );
    if (toPdf.code !== 0 || !fs.existsSync(pdfPath)) {
      return { success: false, error: `soffice->pdf failed: ${toPdf.error || toPdf.stderr || toPdf.stdout}` };
    }

    // 仅首页；白底去透明；高 DPI 有时能减轻“挤在一坨”的观感（根因仍是 LO 对 WMF 的排版）
    const pdfPage0 = `${pdfPath}[0]`;
    const toPng = await runCommandWithTimeout(
      MAGICK_CMD,
      [
        '-density',
        String(rasterDpi),
        pdfPage0,
        '-colorspace',
        'sRGB',
        '-background',
        'white',
        '-alpha',
        'remove',
        '-flatten',
        pngPath
      ],
      Math.max(CONVERT_TIMEOUT_MS * 4, 30000)
    );
    if (toPng.code === 0 && fs.existsSync(pngPath)) return { success: true };
    return { success: false, error: `pdf->png failed: ${toPng.error || toPng.stderr || toPng.stdout}` };
  } catch (e) {
    return { success: false, error: String(e) };
  } finally {
    try {
      fs.rmSync(tmpDir, { recursive: true, force: true });
    } catch {
      // ignore
    }
  }
}

async function convertWithInkscape(
  wmfPath: string,
  pngPath: string,
  outlineText: boolean = false
): Promise<{ success: boolean; error?: string }> {
  const args = [
    wmfPath,
    '--export-type=png',
    '--export-dpi=600',
    ...(outlineText ? ['--export-text-to-path'] : []),
    '--export-filename=' + pngPath
  ];
  let result = await runCommandWithTimeout(INKSCAPE_CMD, args, CONVERT_TIMEOUT_MS);
  if (result.code === 0 && fs.existsSync(pngPath)) return { success: true };

  // 兼容旧版 Inkscape：不支持 --export-text-to-path 时回退
  if (outlineText) {
    const fallbackArgs = [
      wmfPath,
      '--export-type=png',
      '--export-dpi=600',
      '--export-filename=' + pngPath
    ];
    result = await runCommandWithTimeout(INKSCAPE_CMD, fallbackArgs, CONVERT_TIMEOUT_MS);
    if (result.code === 0 && fs.existsSync(pngPath)) return { success: true };
  }

  return { success: false, error: result.error || result.stderr };
}

type ConversionResult = {
  success: boolean;
  method: string;
  outputPath: string;
  error?: string;
  size?: number;
  fillRatio?: number;
  normalized?: boolean;
};

function normalizeMethodParam(value: unknown): 'auto' | 'inkscape' | 'magick' | 'wmf2svg' | 'soffice' | 'gdi' {
  const v = String(value || 'auto').toLowerCase();
  if (v === 'inkscape' || v === 'magick' || v === 'wmf2svg' || v === 'soffice' || v === 'gdi') return v;
  return 'auto';
}

function normalizeFitParam(value: unknown): 'contain' | 'fill' {
  const v = String(value || 'contain').toLowerCase();
  return v === 'fill' ? 'fill' : 'contain';
}

function normalizeNormalizeParam(value: unknown): boolean {
  const v = String(value || '1').toLowerCase();
  // 支持 normalize=0/false/off 禁用
  if (v === '0' || v === 'false' || v === 'off' || v === 'no') return false;
  return true;
}

function normalizeOutlineParam(value: unknown): boolean {
  const v = String(value || '0').toLowerCase();
  // outline=1/true/on 启用 text-to-path
  if (v === '1' || v === 'true' || v === 'on' || v === 'yes') return true;
  return false;
}

async function getImageSize(imagePath: string): Promise<{ width: number; height: number } | null> {
  const result = await runCommandWithTimeout(
    MAGICK_CMD,
    ['identify', '-format', '%w %h', imagePath],
    CONVERT_TIMEOUT_MS
  );
  if (result.code !== 0) return null;
  const parts = result.stdout.trim().split(/\s+/);
  if (parts.length !== 2) return null;
  const width = Number(parts[0]);
  const height = Number(parts[1]);
  if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0) return null;
  return { width, height };
}

async function getTrimmedSize(imagePath: string): Promise<{ width: number; height: number } | null> {
  const result = await runCommandWithTimeout(
    MAGICK_CMD,
    ['convert', imagePath, '-trim', '+repage', '-format', '%w %h', 'info:'],
    CONVERT_TIMEOUT_MS
  );
  if (result.code !== 0) return null;
  const parts = result.stdout.trim().split(/\s+/);
  if (parts.length !== 2) return null;
  const width = Number(parts[0]);
  const height = Number(parts[1]);
  if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0) return null;
  return { width, height };
}

async function normalizeFormulaPng(
  imagePath: string,
  _fillPriority: boolean
): Promise<{ applied: boolean; fillRatio?: number }> {
  const original = await getImageSize(imagePath);
  const trimmed = await getTrimmedSize(imagePath);
  if (!original || !trimmed) return { applied: false };

  const fillRatio = trimmed.width / original.width;
  const heightRatio = trimmed.height / original.height;
  if (fillRatio >= 0.72 || heightRatio < 0.45) {
    return { applied: false, fillRatio };
  }

  const tempPath = imagePath.replace(/\.png$/i, '.normalized.png');
  const containArgs = [
    'convert',
    imagePath,
    '-trim',
    '+repage',
    '-filter',
    'Lanczos',
    '-resize',
    `${original.width}x${original.height}`,
    '-gravity',
    'center',
    '-extent',
    `${original.width}x${original.height}`,
    tempPath
  ];

  // 保真优先：仅使用 contain，避免 ! 拉伸造成字符重叠/变形
  const normalize = await runCommandWithTimeout(MAGICK_CMD, containArgs, CONVERT_TIMEOUT_MS);

  if (normalize.code === 0 && fs.existsSync(tempPath)) {
    fs.copyFileSync(tempPath, imagePath);
    fs.unlinkSync(tempPath);
    return { applied: true, fillRatio };
  }
  if (fs.existsSync(tempPath)) fs.unlinkSync(tempPath);
  return { applied: false, fillRatio };
}

async function convertWmfBestEffort(
  wmfPath: string,
  finalPngPath: string,
  fillPriority: boolean
  ,
  forcedMethod: 'auto' | 'inkscape' | 'magick' | 'wmf2svg' | 'soffice' | 'gdi' = 'auto',
  enableNormalize: boolean = true,
  outlineText: boolean = false,
  rasterDpi: number = 600
): Promise<ConversionResult> {
  ensureCacheDir();
  const candidates: Array<{
    method: string;
    out: string;
    run: (a: string, b: string) => Promise<{ success: boolean; error?: string }>;
  }> = [];

  if (isGdiMetafileRendererAvailable()) {
    candidates.push({
      method: 'gdi',
      out: finalPngPath.replace(/\.png$/, '.gdi.png'),
      run: (a: string, b: string) => convertWithGdiRender(a, b, rasterDpi)
    });
  }

  candidates.push(
    {
      method: 'inkscape',
      out: finalPngPath.replace(/\.png$/, '.inkscape.png'),
      run: (a: string, b: string) => convertWithInkscape(a, b, outlineText)
    },
    {
      method: 'wmf2svg',
      out: finalPngPath.replace(/\.png$/, '.wmf2svg.png'),
      run: (a: string, b: string) => convertWmfToPng(a, b, outlineText)
    },
    {
      method: 'soffice',
      out: finalPngPath.replace(/\.png$/, '.soffice.png'),
      run: (a: string, b: string) => convertWithSoffice(a, b, rasterDpi)
    },
    {
      method: 'magick',
      out: finalPngPath.replace(/\.png$/, '.magick.png'),
      run: (a: string, b: string) => convertWithImageMagick(a, b)
    }
  );

  const ok: ConversionResult[] = [];
  const errors: string[] = [];

  for (const c of candidates) {
    try {
      const res = await c.run(wmfPath, c.out);
      if (res.success && fs.existsSync(c.out)) {
        const stat = fs.statSync(c.out);
        const original = await getImageSize(c.out);
        const trimmed = await getTrimmedSize(c.out);
        let fillRatio: number | undefined = undefined;
        if (original && trimmed) fillRatio = trimmed.width / original.width;
        ok.push({ success: true, method: c.method, outputPath: c.out, size: stat.size, fillRatio });
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

  const autoPreference = isGdiMetafileRendererAvailable()
    ? ['gdi', 'wmf2svg', 'inkscape', 'magick', 'soffice']
    : ['wmf2svg', 'inkscape', 'magick', 'soffice'];

  let winner: ConversionResult | undefined;

  if (forcedMethod !== 'auto') {
    const group = ok.filter((x) => x.method === forcedMethod);
    if (group.length === 0) {
      return {
        success: false,
        method: 'none',
        outputPath: finalPngPath,
        error: errors.join(' | ') || `method=${forcedMethod} unavailable or failed`
      };
    }
    group.sort((a, b) => {
      const fillA = a.fillRatio ?? 0;
      const fillB = b.fillRatio ?? 0;
      if (Math.abs(fillB - fillA) > 0.03) return fillB - fillA;
      return (b.size || 0) - (a.size || 0);
    });
    winner = group[0];
  } else {
    for (const m of autoPreference) {
      const group = ok.filter((x) => x.method === m);
      if (group.length === 0) continue;
      group.sort((a, b) => {
        const fillA = a.fillRatio ?? 0;
        const fillB = b.fillRatio ?? 0;
        if (Math.abs(fillB - fillA) > 0.03) return fillB - fillA;
        return (b.size || 0) - (a.size || 0);
      });
      winner = group[0];
      break;
    }
    if (!winner) winner = ok[0];
  }

  fs.copyFileSync(winner.outputPath, finalPngPath);
  // GDI+ 栅格化已接近 Word，再用 ImageMagick trim/extent 易伤版式；与 README 中 method=gdi&normalize=0 建议一致
  if (enableNormalize && winner.method !== 'gdi') {
    const normalized = await normalizeFormulaPng(finalPngPath, fillPriority);
    if (normalized.applied) {
      winner.normalized = true;
      winner.fillRatio = normalized.fillRatio;
    }
  }

  for (const c of ok) {
    if (c.outputPath !== winner.outputPath && fs.existsSync(c.outputPath)) {
      fs.unlinkSync(c.outputPath);
    }
  }

  return {
    success: true,
    method: winner.method,
    outputPath: finalPngPath,
    size: winner.size,
    fillRatio: winner.fillRatio,
    normalized: winner.normalized
  };
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
        // 尝试 formula_assets 子目录（P3/P4 渲染产物）
        const formulaPath = path.join(imageDir, 'formula_assets', imageName);
        if (fs.existsSync(formulaPath)) {
          imagePath = formulaPath;
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
      const fit = normalizeFitParam(req.query.fit);
      const fillPriority = fit !== 'contain';
      const method = normalizeMethodParam(req.query.method);
      const enableNormalize = normalizeNormalizeParam(req.query.normalize);
      const outlineText = normalizeOutlineParam(req.query.outline);
      const rasterDpi = normalizeRasterDpi(req.query.dpi);
      // 移除 media/ 前缀，避免文件路径问题
      const sanitizedImageName = imageName.replace(/^media\//, '');
      const baseName = sanitizedImageName.replace(/\.(wmf|emf)$/i, '');
      const cachedFileName = `${fileId}_${baseName}.${WMF_CACHE_VERSION}.${fillPriority ? 'fill' : 'contain'}.${method}.${enableNormalize ? 'norm1' : 'norm0'}.${outlineText ? 'outline1' : 'outline0'}.r${rasterDpi}.png`;
      const cachedPngPath = path.join(CACHE_DIR, cachedFileName);
      
      if (fs.existsSync(cachedPngPath)) {
        // logger.info(`Serving cached PNG: ${cachedPngPath}`);
        res.type('png');
        return res.sendFile(cachedPngPath);
      }
      
      const result = await convertWmfBestEffort(
        imagePath,
        cachedPngPath,
        fillPriority,
        method,
        enableNormalize,
        outlineText,
        rasterDpi
      );
      
      if (result.success && fs.existsSync(cachedPngPath)) {
        logger.info(
          `WMF converted by ${result.method}${result.normalized ? ' + normalized' : ''}: ${sanitizedImageName} (fit=${fit}, method=${method}, normalize=${enableNormalize ? '1' : '0'}, dpi=${rasterDpi})`
        );
        // logger.info(`Serving converted PNG: ${cachedPngPath}`);
        res.type('png');
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

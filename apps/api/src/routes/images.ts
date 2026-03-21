import { Router } from 'express';
import path from 'path';
import fs from 'fs';
import { spawn } from 'child_process';
import { logger } from '../utils/logger.js';

const router = Router();

const WMF_EXTENSIONS = ['.wmf', '.emf'];
const CACHE_DIR = path.join(process.cwd(), '..', '..', 'data', 'image_cache');

const WMF2SVG_PATH = '/opt/homebrew/bin/wmf2svg';
const RSVG_CONVERT_PATH = '/opt/homebrew/bin/rsvg-convert';
const MAGICK_PATH = '/opt/homebrew/bin/magick';

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
  return new Promise((resolve) => {
    ensureCacheDir();
    
    const tempSvgPath = pngPath.replace('.png', '.svg');
    
    const wmf2svg = spawn(WMF2SVG_PATH, ['-o', tempSvgPath, wmfPath]);
    
    let wmf2svgError = '';
    wmf2svg.stderr.on('data', (data) => {
      wmf2svgError += data.toString();
    });
    
    wmf2svg.on('close', (code) => {
      if (code !== 0 || !fs.existsSync(tempSvgPath)) {
        // logger.error(`wmf2svg failed: ${wmf2svgError}`);
        resolve({ success: false, error: `wmf2svg failed: ${wmf2svgError}` });
        return;
      }
      
      const rsvg = spawn(RSVG_CONVERT_PATH, [tempSvgPath, '-o', pngPath]);
      
      let rsvgError = '';
      rsvg.stderr.on('data', (data) => {
        rsvgError += data.toString();
      });
      
      rsvg.on('close', (rsvgCode) => {
        if (fs.existsSync(tempSvgPath)) {
          fs.unlinkSync(tempSvgPath);
        }
        
        if (rsvgCode === 0 && fs.existsSync(pngPath)) {
          // logger.info(`WMF conversion successful: ${wmfPath} -> ${pngPath}`);
          resolve({ success: true });
        } else {
          // logger.error(`rsvg-convert failed: ${rsvgError}`);
          resolve({ success: false, error: `rsvg-convert failed: ${rsvgError}` });
        }
      });
      
      rsvg.on('error', (err) => {
        if (fs.existsSync(tempSvgPath)) {
          fs.unlinkSync(tempSvgPath);
        }
        // logger.error(`rsvg-convert process error: ${err}`);
        resolve({ success: false, error: `rsvg-convert process error: ${err}` });
      });
    });
    
    wmf2svg.on('error', (err) => {
      // logger.error(`wmf2svg process error: ${err}`);
      resolve({ success: false, error: `wmf2svg process error: ${err}` });
    });
  });
}

async function convertWithImageMagick(wmfPath: string, pngPath: string): Promise<{ success: boolean; error?: string }> {
  return new Promise((resolve) => {
    if (!fs.existsSync(MAGICK_PATH)) {
      resolve({ success: false, error: 'ImageMagick not found' });
      return;
    }
    
    const process = spawn(MAGICK_PATH, [wmfPath, pngPath]);
    
    let stderr = '';
    process.stderr.on('data', (data) => {
      stderr += data.toString();
    });
    
    process.on('close', (code) => {
      if (code === 0 && fs.existsSync(pngPath)) {
        // logger.info(`ImageMagick conversion successful: ${wmfPath} -> ${pngPath}`);
        resolve({ success: true });
      } else {
        // logger.error(`ImageMagick conversion failed: ${stderr}`);
        resolve({ success: false, error: stderr });
      }
    });
    
    process.on('error', (err) => {
      // logger.error(`ImageMagick process error: ${err}`);
      resolve({ success: false, error: `ImageMagick process error: ${err}` });
    });
  });
}

router.get('/:fileId/:imageName', async (req, res) => {
  try {
    const { fileId, imageName } = req.params;
    
    const imageDir = path.join(process.cwd(), '..', '..', 'data', 'images', fileId);
    let imagePath = path.join(imageDir, imageName);
    
    if (!fs.existsSync(imagePath)) {
      const prefixedPath = path.join(imageDir, `${fileId}_${imageName}`);
      if (fs.existsSync(prefixedPath)) {
        imagePath = prefixedPath;
      } else {
        // logger.error(`Image not found: ${imagePath}`);
        return res.status(404).json({
          success: false,
          message: '图片不存在'
        });
      }
    }
    
    if (isWmfFile(imagePath)) {
      ensureCacheDir();
      const cacheKey = `${fileId}_${imageName}`.replace(/\.(wmf|emf)$/i, '.png');
      const cachedPngPath = path.join(CACHE_DIR, cacheKey);
      
      if (fs.existsSync(cachedPngPath)) {
        // logger.info(`Serving cached PNG: ${cachedPngPath}`);
        return res.sendFile(cachedPngPath);
      }
      
      let result = { success: false, error: 'No conversion method available' };
      
      if (fs.existsSync(WMF2SVG_PATH) && fs.existsSync(RSVG_CONVERT_PATH)) {
        result = await convertWmfToPng(imagePath, cachedPngPath);
      }
      
      if (!result.success) {
        result = await convertWithImageMagick(imagePath, cachedPngPath);
      }
      
      if (result.success && fs.existsSync(cachedPngPath)) {
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

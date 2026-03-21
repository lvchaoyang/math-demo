import { Router } from 'express';
import axios from 'axios';

const router = Router();

router.get('/', async (req, res) => {
  try {
    // 检查 Parser 服务状态
    const PARSER_URL = process.env.PARSER_URL || 'http://localhost:8001';
    const parserHealth = await axios.get(`${PARSER_URL}/health`, { timeout: 5000 });
    
    res.json({
      status: 'healthy',
      services: {
        api: 'running',
        parser: parserHealth.data.status || 'unknown'
      },
      timestamp: new Date().toISOString()
    });
  } catch (error) {
    res.status(503).json({
      status: 'degraded',
      services: {
        api: 'running',
        parser: 'unreachable'
      },
      error: 'Parser service is not available',
      timestamp: new Date().toISOString()
    });
  }
});

export { router as healthRouter };

/**
 * MathJax 3 在 startup 完成后才挂好 typesetPromise；轮询直到可用，避免公式以原始 $...$ 显示。
 * （公式引擎已改为 main.ts 从 npm 包同步引入，不依赖 CDN。）
 */
export async function waitForMathJaxReady(timeoutMs = 20000): Promise<boolean> {
  const start = Date.now()
  while (Date.now() - start < timeoutMs) {
    const mj = (window as unknown as { MathJax?: { typesetPromise?: unknown } }).MathJax
    if (mj && typeof mj.typesetPromise === 'function') {
      return true
    }
    await new Promise((r) => setTimeout(r, 50))
  }
  return false
}

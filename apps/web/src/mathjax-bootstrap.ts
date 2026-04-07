/**
 * 必须在加载 MathJax 组件包之前执行（见 main.ts 中 import 顺序）。
 * 使用 node_modules 内 tex-mml-svg，避免 CDN 失败；SVG 输出避免 CHTML 字体未加载导致的符号错位。
 */
const w = window as unknown as { MathJax: Record<string, unknown> }
w.MathJax = {
  tex: {
    inlineMath: [
      ['$', '$'],
      ['\\(', '\\)'],
    ],
    displayMath: [
      ['$$', '$$'],
      ['\\[', '\\]'],
    ],
  },
  svg: {
    fontCache: 'global',
  },
  options: {
    skipHtmlTags: ['script', 'noscript', 'style', 'textarea', 'pre', 'code'],
  },
}

export {}

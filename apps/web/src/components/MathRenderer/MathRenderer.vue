<template>
  <div ref="mathContainer" class="math-renderer" v-html="renderedContent"></div>
</template>

<script setup lang="ts">
import { ref, watch, nextTick } from 'vue'
import { waitForMathJaxReady } from '../../utils/mathjaxReady'

interface Props {
  content: string
  displayMode?: boolean
}

const props = withDefaults(defineProps<Props>(), {
  displayMode: false
})

const mathContainer = ref<HTMLElement>()
const renderedContent = ref('')

// 声明 MathJax 全局变量
declare global {
  interface Window {
    MathJax: {
      typesetPromise: (elements?: HTMLElement[]) => Promise<void>
      tex: {
        inlineMath: string[][]
        displayMath: string[][]
      }
    }
  }
}

const renderMath = async () => {
  if (!props.content) {
    renderedContent.value = ''
    return
  }

  // 交给 MathJax 解析：后端已输出 $...$ / $$...$$ 及 math-inline/math-block 标记
  renderedContent.value = props.content
  
  // 等待 DOM 更新后渲染 MathJax
  await nextTick()
  
  // 如果内容只有公式图片，没有 MathJax 标记，避免重复 typeset 影响排版
  const shouldTypeset =
    props.content.includes('math-inline') ||
    props.content.includes('math-block') ||
    props.content.includes('data-latex') ||
    props.content.includes('$')

  if (!shouldTypeset || !mathContainer.value) {
    return
  }
  const mjReady = await waitForMathJaxReady()
  if (!mjReady) {
    console.warn('MathJax 加载超时，公式可能以源码显示')
    return
  }
  const mj = window.MathJax
  if (!mj?.typesetPromise) {
    return
  }
  try {
    await mj.typesetPromise([mathContainer.value])
  } catch (error) {
    console.error('MathJax 渲染失败:', error)
  }
}

watch(() => props.content, renderMath, { immediate: true })
</script>

<style scoped>
/* 行内公式与正文混排：不要用 middle 挤压 MathJax 3 CHTML 内部定位（根号横线、上下标会错位） */
.math-renderer {
  display: block;
  line-height: 1.65;
}

.math-renderer :deep(.math-inline) {
  display: inline;
  vertical-align: baseline;
  margin: 0 2px;
  line-height: normal;
}

.math-renderer :deep(.math-block) {
  display: block;
  text-align: center;
  margin: 12px 0;
  line-height: normal;
  overflow-x: auto;
}

/* MathJax 3：与中文混排时基线对齐；勿用 vertical-align: middle */
.math-renderer :deep(mjx-container),
.math-renderer :deep(.mjx-container) {
  vertical-align: baseline;
}

.math-renderer :deep(.question-image) {
  max-width: 100%;
  height: auto;
  vertical-align: middle;
  margin: 0 6px;
}

.math-renderer :deep(img) {
  max-width: 100%;
  height: auto;
  vertical-align: middle;
  margin: 0 6px;
  display: inline-block;
}

.math-renderer :deep(.formula-image) {
  /* 公式 WMF 图片不要强行 max-height：会导致挤压重叠 */
  max-width: 100%;
  height: 28px;
  vertical-align: middle;
  margin: 0 6px;
  display: inline-block;
}

.math-renderer :deep(.formula-image-block) {
  display: block;
  margin: 12px 6px;
  max-width: 100%;
  height: auto;
  object-fit: contain;
}

.math-renderer :deep(.option-image) {
  max-height: 18px;
  vertical-align: middle;
  margin: 0 6px;
  display: inline-block;
}
</style>

<template>
  <div ref="mathContainer" class="math-renderer" v-html="renderedContent"></div>
</template>

<script setup lang="ts">
import { ref, watch, onMounted, nextTick } from 'vue'

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
  if (!props.content) return
  
  let content = props.content
  
  // 处理 LaTeX 公式
  // 将 $...$ 替换为行内公式标记
  content = content.replace(/\$([^$]+)\$/g, (match, formula) => {
    return `\\(${formula}\\)`
  })
  
  // 将 $$...$$ 替换为块级公式标记
  content = content.replace(/\$\$([^$]+)\$\$/g, (match, formula) => {
    return `\\[${formula}\\]`
  })
  
  // 处理 HTML 中的公式标记
  content = content.replace(/<span class="math-inline">\$([^$]+)\$<\/span>/g, (match, formula) => {
    return `\\(${formula}\\)`
  })
  
  content = content.replace(/<div class="math-block">\$\$([^$]+)\$\$<\/div>/g, (match, formula) => {
    return `\\[${formula}\\]`
  })
  
  renderedContent.value = content
  
  // 等待 DOM 更新后渲染 MathJax
  await nextTick()
  
  if (window.MathJax && mathContainer.value) {
    try {
      await window.MathJax.typesetPromise([mathContainer.value])
    } catch (error) {
      console.error('MathJax 渲染失败:', error)
    }
  }
}

watch(() => props.content, renderMath, { immediate: true })

onMounted(() => {
  renderMath()
})
</script>

<style scoped>
.math-renderer {
  display: inline;
}

.math-renderer :deep(.MathJax) {
  display: inline-block;
  vertical-align: middle;
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

.math-renderer :deep(.formula-image),
.math-renderer :deep(.option-image) {
  max-height: 30px;
  vertical-align: middle;
  margin: 0 6px;
  display: inline-block;
}
</style>

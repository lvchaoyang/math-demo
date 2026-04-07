<template>
  <el-card class="question-list-card">
    <template #header>
      <div class="card-header">
        <span>题目列表 (共 {{ questions.length }} 道)</span>
        <el-checkbox
          :model-value="selectAll"
          @change="handleSelectAll"
        >
          全选
        </el-checkbox>
      </div>
    </template>
    
    <div class="question-list">
      <div
        v-for="question in questions"
        :key="question.id"
        class="question-item"
        :class="{ 'is-selected': isSelected(question), 'is-low-confidence': question.low_confidence }"
      >
        <div class="question-header">
          <el-checkbox
            :model-value="isSelected(question)"
            @change="(val: string | number | boolean) => handleSelect(question, val)"
          >
            <span class="question-number">第 {{ question.number }} 题</span>
            <el-tag size="small" class="question-type">{{ question.type_name }}</el-tag>
            <el-tag
              v-if="question.low_confidence"
              size="small"
              type="warning"
              class="confidence-tag"
              :title="(question.low_confidence_reasons || []).join('；')"
            >
              低置信度
            </el-tag>
          </el-checkbox>
        </div>
        
        <div class="question-content">
          <math-renderer :content="question.content_html" />
        </div>
        
        <!-- 选项（选择题） -->
        <div v-if="question.options && question.options.length > 0" class="question-options">
          <div
            v-for="option in question.options"
            :key="option.label"
            class="option-item"
          >
            <!-- <span class="option-label">{{ option.label }}.</span> -->
            <math-renderer :content="option.content_html" />
          </div>
        </div>
        
        <!-- 图片（包括公式图片） -->
        <!-- <div v-if="question.images && question.images.length > 0" class="question-images">
          <img
            v-for="img in question.images"
            :key="img"
            :src="getImageUrl(img)"
            class="question-image"
            :class="{ 'formula-image': isFormulaImage(img) }"
            alt="题目图片"
            @error="handleImageError"
          />
        </div> -->
      </div>
    </div>
  </el-card>
</template>

<script setup lang="ts">
import { computed } from 'vue'
import MathRenderer from '../MathRenderer/MathRenderer.vue'
import type { Question } from '../../types'

interface Props {
  questions: Question[]
  /** 当前试卷 ID，用于跨卷组卷 */
  fileId: string
  /** 已加入组卷的本卷题目 id */
  assemblyQuestionIds: string[]
}

const props = defineProps<Props>()

const emit = defineEmits<{
  (
    e: 'toggle-assembly',
    payload: { fileId: string; question: Question; selected: boolean }
  ): void
  (
    e: 'select-all-assembly',
    payload: { fileId: string; questions: Question[]; selected: boolean }
  ): void
}>()

const selectAll = computed(
  () =>
    props.questions.length > 0 &&
    props.questions.every((q) => props.assemblyQuestionIds.includes(q.id))
)

const isSelected = (question: Question) => {
  return props.assemblyQuestionIds.includes(question.id)
}

// 获取图片URL
const getImageUrl = (img: string) => {
  // 如果img已经是完整路径，直接返回
  if (img.startsWith('http') || img.startsWith('/')) {
    return img
  }
  // 从文件名中提取file_id（格式: {file_id}_{original_filename}）
  const parts = img.split('_')
  if (parts.length >= 2) {
    const fileId = parts[0]
    // 返回原始文件名，后端会处理查找
    return `/api/v1/images/${fileId}/${img}`
  }
  return `/uploads/images/${img}`
}

// 判断是否是公式图片
const isFormulaImage = (img: string) => {
  return img.endsWith('.wmf') || img.endsWith('.emf')
}

// 处理图片加载错误
const handleImageError = (e: Event) => {
  const img = e.target as HTMLImageElement
  console.error('图片加载失败:', img.src)
  // 可以在这里设置默认图片
}

function coerceChecked(raw: string | number | boolean): boolean {
  if (typeof raw === 'boolean') return raw
  return raw === 'true' || raw === 1
}

const handleSelect = (question: Question, raw: string | number | boolean) => {
  emit('toggle-assembly', {
    fileId: props.fileId,
    question,
    selected: coerceChecked(raw)
  })
}

const handleSelectAll = (val: boolean | string | number) => {
  const selected = val === true || val === 'true' || val === 1
  emit('select-all-assembly', {
    fileId: props.fileId,
    questions: props.questions,
    selected
  })
}

</script>

<style scoped>
.question-list-card {
  margin-bottom: 0;
  border: none;
  box-shadow: none !important;
  background: transparent;
}

.question-list-card :deep(.el-card__header) {
  padding: 0 0 12px 0;
  border-bottom: none;
  background: transparent;
}

.question-list-card :deep(.el-card__body) {
  padding: 0;
}

.card-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  font-size: 15px;
  font-weight: 600;
  color: var(--md-text, #1f2937);
}

.question-list {
  display: flex;
  flex-direction: column;
  gap: 12px;
}

.question-item {
  padding: 16px 18px;
  border: 1px solid var(--md-border, #e4e7ef);
  border-radius: var(--md-radius-sm, 8px);
  transition:
    border-color 0.2s ease,
    box-shadow 0.2s ease,
    background 0.2s ease;
  background: var(--md-surface, #fff);
  box-shadow: var(--md-shadow-sm);
}

.question-item:hover {
  box-shadow: var(--md-shadow-md);
  border-color: var(--md-border-strong, #dcdfe8);
}

.question-item.is-selected {
  border-color: var(--md-primary, #5b6ee8);
  background: var(--md-primary-soft, #eef0fc);
  box-shadow: 0 0 0 1px rgba(91, 110, 232, 0.2);
}

.question-item.is-low-confidence {
  border-color: #e8b86d;
  background: #fffbf0;
}

.question-header {
  margin-bottom: 10px;
}

.question-number {
  font-weight: 600;
  margin-right: 10px;
}

.question-type {
  margin-left: 10px;
}

.confidence-tag {
  margin-left: 8px;
}

.question-content {
  margin-left: 24px;
  /* 与 MathRenderer 内 line-height 协调；过大行高会放大行盒与公式基线差异 */
  line-height: 1.65;
  color: var(--md-text, #1f2937);
}

.question-options {
  margin-left: 24px;
  margin-top: 10px;
  display: grid;
  grid-template-columns: repeat(2, 1fr);
  gap: 10px;
}

.option-item {
  display: flex;
  align-items: center;
  gap: 5px;
  padding: 10px 12px;
  background: #f8f9fc;
  border: 1px solid var(--md-border, #e4e7ef);
  border-radius: var(--md-radius-sm, 8px);
  overflow: visible;
}

.option-item .math-renderer {
  flex: 1;
  min-width: 0;
}

.option-label {
  font-weight: 600;
  color: #606266;
  min-width: 20px;
}

.question-images {
  margin-left: 24px;
  margin-top: 15px;
  display: flex;
  flex-wrap: wrap;
  gap: 10px;
}

.question-image {
  max-width: 300px;
  max-height: 200px;
  border: 1px solid #ebeef5;
  border-radius: 4px;
}

.formula-image {
  max-height: 30px;
  vertical-align: middle;
  margin: 0 6px;
  display: inline-block;
}
</style>

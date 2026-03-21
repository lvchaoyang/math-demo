<template>
  <el-card class="question-list-card">
    <template #header>
      <div class="card-header">
        <span>题目列表 (共 {{ questions.length }} 道)</span>
        <el-checkbox
          v-model="selectAll"
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
        :class="{ 'is-selected': isSelected(question) }"
      >
        <div class="question-header">
          <el-checkbox
            :model-value="isSelected(question)"
            @change="(val: boolean) => handleSelect(question, val)"
          >
            <span class="question-number">第 {{ question.number }} 题</span>
            <el-tag size="small" class="question-type">{{ question.type_name }}</el-tag>
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
import { ref, watch } from 'vue'
import MathRenderer from '../MathRenderer/MathRenderer.vue'
import type { Question } from '../../types'

interface Props {
  questions: Question[]
}

const props = defineProps<Props>()

const emit = defineEmits<{
  (e: 'selection-change', selection: Question[]): void
}>()

const selectedQuestions = ref<Question[]>([])
const selectAll = ref(false)

const isSelected = (question: Question) => {
  return selectedQuestions.value.some(q => q.id === question.id)
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

const handleSelect = (question: Question, selected: boolean) => {
  if (selected) {
    if (!isSelected(question)) {
      selectedQuestions.value.push(question)
    }
  } else {
    selectedQuestions.value = selectedQuestions.value.filter(q => q.id !== question.id)
  }
  selectAll.value = selectedQuestions.value.length === props.questions.length
  emit('selection-change', selectedQuestions.value)
}

const handleSelectAll = (val: boolean) => {
  if (val) {
    selectedQuestions.value = [...props.questions]
  } else {
    selectedQuestions.value = []
  }
  emit('selection-change', selectedQuestions.value)
}

// 监听题目变化，清空选择
watch(() => props.questions, () => {
  selectedQuestions.value = []
  selectAll.value = false
  emit('selection-change', [])
}, { deep: true })
</script>

<style scoped>
.question-list-card {
  margin-bottom: 20px;
}

.card-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  font-size: 16px;
  font-weight: 600;
}

.question-list {
  display: flex;
  flex-direction: column;
  gap: 15px;
}

.question-item {
  padding: 15px;
  border: 1px solid #ebeef5;
  border-radius: 8px;
  transition: all 0.3s;
  background: #fff;
}

.question-item:hover {
  box-shadow: 0 2px 12px rgba(0, 0, 0, 0.1);
}

.question-item.is-selected {
  border-color: #409eff;
  background: #f0f9ff;
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

.question-content {
  margin-left: 24px;
  line-height: 1.8;
  color: #303133;
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
  align-items: flex-start;
  gap: 5px;
  padding: 8px;
  background: #f5f7fa;
  border-radius: 4px;
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

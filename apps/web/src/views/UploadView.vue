<template>
  <div class="upload-view">
    <!-- 上传区域 -->
    <el-card class="upload-card">
      <template #header>
        <div class="card-header">
          <span>上传试卷</span>
        </div>
      </template>
      
      <!-- 解析模式选择 -->
      <div class="mode-selector">
        <span class="mode-label">解析模式：</span>
        <el-radio-group v-model="parseMode" size="large">
          <el-radio-button label="questions">
            <el-icon><Document /></el-icon>
            题目拆分
          </el-radio-button>
          <el-radio-button label="html">
            <el-icon><Document /></el-icon>
            整体HTML
          </el-radio-button>
        </el-radio-group>
        <div class="mode-description">
          <template v-if="parseMode === 'questions'">
            <el-tag type="info" size="small">推荐</el-tag>
            <span>智能拆分题目，支持单独查看和导出</span>
          </template>
          <template v-else>
            <el-tag type="warning" size="small">快速</el-tag>
            <span>整体转换为HTML，公式图片内嵌，适合快速预览</span>
          </template>
        </div>
      </div>

      <el-upload
        class="upload-area"
        drag
        :action="uploadAction"
        accept=".docx"
        :data="{ mode: parseMode }"
        :on-success="handleUploadSuccess"
        :on-error="handleUploadError"
        :before-upload="beforeUpload"
        :show-file-list="false"
      >
        <el-icon class="upload-icon"><Upload /></el-icon>
        <div class="upload-text">
          <em>点击上传</em> 或拖拽文件到此处
        </div>
        <template #tip>
          <div class="upload-tip">
            仅支持 .docx 格式的 Word 文档，文件大小不超过 50MB
          </div>
        </template>
      </el-upload>
    </el-card>

    <!-- 解析进度 -->
    <el-card v-if="parsingStatus" class="progress-card">
      <template #header>
        <div class="card-header">
          <span>解析进度</span>
        </div>
      </template>
      <el-progress 
        :percentage="parseProgress" 
        :status="progressStatus"
        :stroke-width="20"
      />
      <p class="progress-message">{{ parseMessage }}</p>
    </el-card>
    
    <!-- HTML 预览 -->
    <el-card v-if="htmlContent" class="html-preview-card">
      <template #header>
        <div class="card-header">
          <span>文档预览</span>
          <el-button type="primary" size="small" @click="downloadHtml">
            <el-icon><Download /></el-icon>
            下载 HTML
          </el-button>
        </div>
      </template>
      <div ref="htmlPreviewRef" class="html-content" v-html="htmlContent"></div>
    </el-card>
    
    <!-- 题目列表 -->
    <question-list
      v-if="questions.length > 0"
      :questions="questions"
      @selection-change="handleSelectionChange"
    />
    
    <!-- 导出面板 -->
    <export-panel
      v-if="selectedQuestions.length > 0"
      :file-id="currentFileId"
      :selected-count="selectedQuestions.length"
      :selected-ids="selectedQuestionIds"
      @export-success="handleExportSuccess"
    />
  </div>
</template>

<script setup lang="ts">
import { ref, computed, nextTick } from 'vue'
import { ElMessage } from 'element-plus'
import { Upload, Document, Download } from '@element-plus/icons-vue'
import QuestionList from '../components/QuestionList/QuestionList.vue'
import ExportPanel from '../components/ExportPanel/ExportPanel.vue'
import type { Question } from '../types'

const parseMode = ref<'questions' | 'html'>('questions')
const parsingStatus = ref(false)
const parseProgress = ref(0)
const parseMessage = ref('')
const currentFileId = ref('')
const questions = ref<Question[]>([])
const selectedQuestions = ref<Question[]>([])
const htmlContent = ref('')
const htmlPreviewRef = ref<HTMLElement | null>(null)

const uploadAction = computed(() => '/api/v1/upload')

const progressStatus = computed(() => {
  if (parseProgress.value === 100) return 'success'
  return ''
})

const selectedQuestionIds = computed(() => {
  return selectedQuestions.value.map(q => q.id)
})

const beforeUpload = (file: File) => {
  const isDocx = file.name.endsWith('.docx')
  const isLt50M = file.size / 1024 / 1024 < 50

  if (!isDocx) {
    ElMessage.error('只支持 .docx 格式的文件!')
    return false
  }
  if (!isLt50M) {
    ElMessage.error('文件大小不能超过 50MB!')
    return false
  }
  
  parsingStatus.value = true
  parseProgress.value = 0
  parseMessage.value = '正在上传...'
  htmlContent.value = ''
  questions.value = []
  return true
}

const handleUploadSuccess = (response: any) => {
  if (response.success) {
    currentFileId.value = response.file_id
    parseMessage.value = '上传成功，开始解析...'
    startPollingProgress(response.file_id)
  } else {
    ElMessage.error(response.message || '上传失败')
    parsingStatus.value = false
  }
}

const handleUploadError = () => {
  ElMessage.error('上传失败')
  parsingStatus.value = false
}

const startPollingProgress = (fileId: string) => {
  const pollInterval = setInterval(async () => {
    try {
      const response = await fetch(`/api/v1/upload/progress/${fileId}`)
      const data = await response.json()
      
      parseProgress.value = data.progress
      parseMessage.value = data.message
      
      if (data.status === 'completed') {
        clearInterval(pollInterval)
        
        if (data.mode === 'html') {
          // HTML 模式
          htmlContent.value = data.html
          ElMessage.success('HTML 转换完成')

            // 让新插入的 HTML 也进行 MathJax typeset，避免公式重叠/错位
            await nextTick()
            if ((window as any).MathJax && htmlPreviewRef.value) {
              try {
                const shouldTypeset =
                  htmlContent.value.includes('math-inline') ||
                  htmlContent.value.includes('math-block') ||
                  htmlContent.value.includes('data-latex') ||
                  htmlContent.value.includes('$')

                if (shouldTypeset) {
                  await (window as any).MathJax.typesetPromise([htmlPreviewRef.value])
                }
              } catch (e) {
                console.error('HTML MathJax 渲染失败:', e)
              }
            }
        } else {
          // 题目拆分模式
          questions.value = data.questions
          ElMessage.success(`解析完成，共 ${data.questions.length} 道题目`)
        }
      } else if (data.status === 'error') {
        clearInterval(pollInterval)
        ElMessage.error(data.message)
      }
    } catch (error) {
      console.error('获取进度失败:', error)
    }
  }, 1000)
}

const handleSelectionChange = (selection: Question[]) => {
  selectedQuestions.value = selection
}

const handleExportSuccess = () => {
  ElMessage.success('导出成功')
}

const downloadHtml = () => {
  const blob = new Blob([htmlContent.value], { type: 'text/html' })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = `document_${currentFileId.value}.html`
  a.click()
  URL.revokeObjectURL(url)
  ElMessage.success('HTML 文件已下载')
}
</script>

<style scoped>
.upload-view {
  max-width: 1200px;
  margin: 0 auto;
}

.upload-card,
.progress-card,
.html-preview-card {
  margin-bottom: 20px;
}

.card-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  font-size: 16px;
  font-weight: 600;
}

.mode-selector {
  margin-bottom: 20px;
  padding: 15px;
  background-color: #f5f7fa;
  border-radius: 4px;
}

.mode-label {
  font-weight: 600;
  margin-right: 10px;
  color: #606266;
}

.mode-description {
  margin-top: 10px;
  font-size: 13px;
  color: #909399;
  display: flex;
  align-items: center;
  gap: 8px;
}

.upload-area {
  width: 100%;
}

.upload-icon {
  font-size: 60px;
  color: #c0c4cc;
  margin-bottom: 15px;
}

.upload-text {
  color: #606266;
  font-size: 14px;
}

.upload-text em {
  color: #409eff;
  font-style: normal;
}

.upload-tip {
  margin-top: 15px;
  color: #909399;
  font-size: 12px;
}

.progress-message {
  text-align: center;
  margin-top: 15px;
  color: #606266;
}

.html-content {
  padding: 20px;
  background-color: #fff;
  border: 1px solid #e4e7ed;
  border-radius: 4px;
  max-height: 800px;
  overflow-y: auto;
  line-height: 1.8;
}

.html-content :deep(img) {
  max-width: 100%;
  height: 20px;
  vertical-align: middle;
  margin: 10px 0;
  /* display: block; */
  /* margin: 10px auto; */
}

.html-content :deep(.formula-image) {
  display: inline-block;
  vertical-align: middle;
  margin: 0 5px;
}

.html-content :deep(.math-inline) {
  display: inline;
  margin: 0 2px;
}

.html-content :deep(.math-block) {
  display: block;
  text-align: center;
  margin: 15px 0;
}

.html-content :deep(p) {
  margin: 0.5em 0;
}

.html-content :deep(h1),
.html-content :deep(h2),
.html-content :deep(h3) {
  margin: 1em 0 0.5em;
  font-weight: 600;
}
</style>

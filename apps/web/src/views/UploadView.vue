<template>
  <div class="upload-view">
    <!-- 首屏：可选模式 + 大拖拽区 -->
    <el-card v-if="!shouldCompactUpload" class="upload-card">
      <template #header>
        <div class="card-header">
          <span>上传试卷</span>
        </div>
      </template>

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
            <span>支持多份试卷；勾选题目可跨卷组卷导出</span>
          </template>
          <template v-else>
            <el-tag type="warning" size="small">快速</el-tag>
            <span>整体转换为 HTML，公式图片内嵌，适合快速预览</span>
          </template>
        </div>
      </div>

      <el-upload
        class="upload-area"
        :class="{ 'is-parsing': parsingCounter > 0 }"
        drag
        :action="uploadAction"
        accept=".docx"
        :data="{ mode: parseMode }"
        :multiple="parseMode === 'questions'"
        :disabled="parseMode === 'html' && parsingCounter > 0"
        :on-success="handleUploadSuccess"
        :on-error="handleUploadError"
        :before-upload="beforeUpload"
        :show-file-list="false"
      >
        <el-icon class="upload-icon"><Upload /></el-icon>
        <div class="upload-text">
          <em>点击上传</em> 或拖拽文件到此处
          <span v-if="parseMode === 'questions'" class="multi-hint">（题目模式下可多选文件）</span>
        </div>
        <template #tip>
          <div class="upload-tip">仅支持 .docx，单文件不超过 50MB</div>
        </template>
      </el-upload>
    </el-card>

    <!-- 进入会话后：模式已锁定，单行「继续上传 + 当前模式 + 上传」 -->
    <el-card v-else class="upload-card upload-card--compact">
      <div class="upload-inline-row">
        <span class="upload-inline-label">继续上传</span>
        <el-tag type="info" effect="plain" size="small" class="upload-mode-tag">
          当前模式：{{ compactModeLabel }}
        </el-tag>
        <div class="upload-compact-upload-wrap">
          <el-upload
            class="upload-area upload-area--compact"
            :class="{ 'is-parsing': parsingCounter > 0 }"
            :action="uploadAction"
            accept=".docx"
            :data="{ mode: effectiveParseMode }"
            :multiple="effectiveParseMode === 'questions'"
            :disabled="effectiveParseMode === 'html' && parsingCounter > 0"
            :on-success="handleUploadSuccess"
            :on-error="handleUploadError"
            :before-upload="beforeUpload"
            :show-file-list="false"
          >
            <el-button type="primary">
              <el-icon><Upload /></el-icon>
              {{ effectiveParseMode === 'questions' ? '添加试卷' : '选择 Word' }}
            </el-button>
          </el-upload>
          <span class="upload-tip-compact">
            .docx，≤50MB<span v-if="effectiveParseMode === 'questions'">；可多选</span>
          </span>
        </div>
      </div>
    </el-card>

    <!-- 轻量状态条：不遮挡上传区，多卷时每卷卡片内另有骨架屏 -->
    <el-alert
      v-if="parsingCounter > 0"
      class="parse-status-alert"
      type="info"
      :closable="false"
      show-icon
    >
      <template #title>
        <span class="parse-status-title">
          正在处理 {{ parsingCounter }} 个任务
        </span>
      </template>
      <template #default>
        <p v-if="effectiveParseMode === 'questions'" class="parse-status-desc">
          题目模式下可继续上传其他试卷；下方各卷卡片会显示解析进度。
        </p>
        <p v-else class="parse-status-desc">正在转换为 HTML，请稍候。</p>
      </template>
    </el-alert>

    <!-- HTML 模式：解析中占位（避免整页空白） -->
    <el-card
      v-if="effectiveParseMode === 'html' && parsingCounter > 0 && !htmlContent"
      class="html-wait-card"
    >
      <template #header>
        <span>正在生成预览…</span>
      </template>
      <el-skeleton :rows="6" animated />
    </el-card>

    <!-- HTML 预览（单份） -->
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

    <!-- 多份试卷题目 -->
    <template v-if="effectiveParseMode === 'questions' && papers.length > 0">
      <el-card
        v-for="paper in papers"
        :key="paper.fileId"
        class="paper-card"
        :data-paper-id="paper.fileId"
      >
        <template #header>
          <div class="paper-header">
            <div class="paper-title">
              <el-icon><Document /></el-icon>
              <span class="paper-name">{{ paper.filename }}</span>
              <el-tag v-if="paper.parsing" type="warning" size="small">解析中</el-tag>
              <el-tag v-else type="success" size="small">{{ paper.questions.length }} 题</el-tag>
            </div>
            <el-button
              type="danger"
              text
              size="small"
              :disabled="paper.parsing"
              @click="removePaper(paper.fileId)"
            >
              移除此卷
            </el-button>
          </div>
        </template>

        <div v-if="paper.parsing" class="paper-skeleton">
          <p class="paper-skeleton-hint">正在拆分题目与渲染公式，大文件可能需稍长时间…</p>
          <el-skeleton :rows="5" animated />
        </div>
        <question-list
          v-else-if="paper.questions.length > 0"
          :file-id="paper.fileId"
          :questions="paper.questions"
          :assembly-question-ids="assemblyQuestionIdsFor(paper.fileId)"
          @toggle-assembly="onToggleAssembly"
          @select-all-assembly="onSelectAllAssembly"
        />
        <el-empty v-else description="本题卷无题目" :image-size="80" />
      </el-card>
    </template>

    <div v-if="assemblyItems.length > 0" class="assembly-sticky-wrap">
      <div class="assembly-toolbar">
        <span class="assembly-toolbar-text">
          已选 <strong>{{ assemblyItems.length }}</strong> 题，可跨卷组卷
        </span>
        <el-button type="danger" text size="small" @click="confirmClearAssembly">
          清空已选
        </el-button>
      </div>
      <export-panel
        :assembly-items="assemblyRows"
        @update-assembly="onUpdateAssemblyFromPreview"
      />
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref, computed, watch, nextTick } from 'vue'
import { ElMessage, ElMessageBox } from 'element-plus'
import { Upload, Document, Download } from '@element-plus/icons-vue'
import QuestionList from '../components/QuestionList/QuestionList.vue'
import ExportPanel from '../components/ExportPanel/ExportPanel.vue'
import type { Question } from '../types'
import { waitForMathJaxReady } from '../utils/mathjaxReady'

interface PaperEntry {
  fileId: string
  filename: string
  questions: Question[]
  parsing: boolean
}

interface AssemblyItem {
  fileId: string
  question: Question
  clientKey: string
}

const parseMode = ref<'questions' | 'html'>('questions')
const parsingCounter = ref(0)
const papers = ref<PaperEntry[]>([])
const assemblyItems = ref<AssemblyItem[]>([])
const htmlContent = ref('')
const htmlFileId = ref('')
const htmlPreviewRef = ref<HTMLElement | null>(null)

const pollIntervals = new Map<string, ReturnType<typeof setInterval>>()
const pollAttempts = new Map<string, number>()
/** 轮询间隔 1s，最多约 15 分钟，避免异常时无限轮询 */
const MAX_POLL_ATTEMPTS = 900

const uploadAction = computed(() => '/api/v1/upload')

/** 进入紧凑流程时锁定，避免与首屏 `parseMode` 双轨；上传与展示均以此为准 */
const lockedParseMode = ref<'questions' | 'html' | null>(null)

/** 已有解析任务、题目卷或 HTML 结果时，顶部改为紧凑条，少占纵向空间 */
const shouldCompactUpload = computed(() => {
  if (parsingCounter.value > 0) return true
  if (papers.value.length > 0) return true
  if (parseMode.value === 'html' && htmlContent.value.trim().length > 0) return true
  return false
})

watch(
  shouldCompactUpload,
  (compact) => {
    if (compact) {
      if (lockedParseMode.value === null) {
        lockedParseMode.value = parseMode.value
      }
    } else {
      lockedParseMode.value = null
    }
  },
  { flush: 'sync' }
)

const effectiveParseMode = computed(() => {
  if (shouldCompactUpload.value && lockedParseMode.value !== null) {
    return lockedParseMode.value
  }
  return parseMode.value
})

const compactModeLabel = computed(() =>
  effectiveParseMode.value === 'questions' ? '题目拆分' : '整体 HTML'
)

function newClientKey() {
  return crypto.randomUUID()
}

const assemblyRows = computed(() =>
  assemblyItems.value.map((item) => ({
    fileId: item.fileId,
    question: item.question,
    clientKey: item.clientKey,
    sourceLabel: papers.value.find((p) => p.fileId === item.fileId)?.filename ?? item.fileId.slice(0, 8)
  }))
)

function assemblyQuestionIdsFor(fileId: string) {
  return assemblyItems.value.filter((i) => i.fileId === fileId).map((i) => i.question.id)
}

function onUpdateAssemblyFromPreview(
  items: Array<{ fileId: string; question: Question; clientKey: string }>
) {
  assemblyItems.value = items
}

function stopPoll(fileId: string) {
  const id = pollIntervals.get(fileId)
  if (id !== undefined) {
    clearInterval(id)
    pollIntervals.delete(fileId)
  }
  pollAttempts.delete(fileId)
}

function failPollJob(
  fileId: string,
  mode: 'questions' | 'html',
  message: string
) {
  stopPoll(fileId)
  parsingCounter.value = Math.max(0, parsingCounter.value - 1)
  if (mode === 'questions') {
    papers.value = papers.value.filter((p) => p.fileId !== fileId)
  }
  ElMessage.error(message)
}

async function confirmClearAssembly() {
  try {
    await ElMessageBox.confirm('确定清空当前已选题目？', '清空组卷', {
      type: 'warning',
      confirmButtonText: '清空',
      cancelButtonText: '取消'
    })
    assemblyItems.value = []
    ElMessage.success('已清空')
  } catch {
    /* 用户取消 */
  }
}

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

  parsingCounter.value += 1
  if (effectiveParseMode.value === 'html') {
    htmlContent.value = ''
  }
  return true
}

const handleUploadSuccess = (response: any, uploadFile: { name: string }) => {
  if (!response?.success) {
    ElMessage.error(response?.message || '上传失败')
    parsingCounter.value = Math.max(0, parsingCounter.value - 1)
    return
  }

  const fileId = response.file_id as string
  const filename = uploadFile.name

  if (effectiveParseMode.value === 'html') {
    htmlFileId.value = fileId
    startPollingProgress(fileId, filename, 'html')
    return
  }

  papers.value.push({
    fileId,
    filename,
    questions: [],
    parsing: true
  })
  startPollingProgress(fileId, filename, 'questions')
}

const handleUploadError = () => {
  ElMessage.error('上传失败')
  parsingCounter.value = Math.max(0, parsingCounter.value - 1)
}

const startPollingProgress = (
  fileId: string,
  filename: string,
  mode: 'questions' | 'html'
) => {
  stopPoll(fileId)
  pollAttempts.set(fileId, 0)
  const interval = setInterval(async () => {
    try {
      const n = (pollAttempts.get(fileId) ?? 0) + 1
      pollAttempts.set(fileId, n)
      if (n > MAX_POLL_ATTEMPTS) {
        failPollJob(
          fileId,
          mode,
          '等待解析结果超时，请检查网络与解析服务后重试'
        )
        return
      }

      const res = await fetch(`/api/v1/upload/progress/${fileId}`)
      const data = await res.json()

      if (data.status === 'completed') {
        stopPoll(fileId)
        parsingCounter.value = Math.max(0, parsingCounter.value - 1)

        if (data.mode === 'html' || mode === 'html') {
          const raw = data.html
          htmlContent.value = typeof raw === 'string' ? raw : ''
          if (!htmlContent.value.trim()) {
            ElMessage.warning('HTML 内容为空，请确认已安装 Pandoc 或文档是否可读')
          } else {
            ElMessage.success('HTML 转换完成')
            await nextTick()
            await typesetHtmlIfNeeded()
          }
        } else {
          const paper = papers.value.find((p) => p.fileId === fileId)
          if (paper) {
            paper.parsing = false
            paper.questions = Array.isArray(data.questions) ? data.questions : []
          }
          const qn = paper?.questions.length ?? 0
          ElMessage.success(`「${filename}」解析完成，共 ${qn} 道题`)
          await nextTick()
          document
            .querySelector(`[data-paper-id="${fileId}"]`)
            ?.scrollIntoView({ behavior: 'smooth', block: 'nearest' })
        }
      } else if (data.status === 'error') {
        failPollJob(fileId, mode, data.message || '解析失败')
      }
    } catch (e) {
      console.error('获取进度失败:', e)
    }
  }, 1000)
  pollIntervals.set(fileId, interval)
}

async function typesetHtmlIfNeeded() {
  if (!htmlPreviewRef.value) return
  const ready = await waitForMathJaxReady()
  if (!ready) return
  const mj = (window as unknown as { MathJax?: { typesetPromise?: (n: Element[]) => Promise<void> } })
    .MathJax
  if (!mj?.typesetPromise) return
  const html = htmlContent.value
  try {
    const shouldTypeset =
      html.includes('math-inline') ||
      html.includes('math-block') ||
      html.includes('data-latex') ||
      html.includes('$')
    if (shouldTypeset) {
      await mj.typesetPromise([htmlPreviewRef.value])
    }
  } catch (e) {
    console.error('HTML MathJax 渲染失败:', e)
  }
}

function onToggleAssembly(payload: {
  fileId: string
  question: Question
  selected: boolean
}) {
  const { fileId, question, selected } = payload
  if (selected) {
    const exists = assemblyItems.value.some(
      (i) => i.fileId === fileId && i.question.id === question.id
    )
    if (!exists) {
      assemblyItems.value.push({
        fileId,
        question,
        clientKey: newClientKey()
      })
    }
  } else {
    assemblyItems.value = assemblyItems.value.filter(
      (i) => !(i.fileId === fileId && i.question.id === question.id)
    )
  }
}

function onSelectAllAssembly(payload: {
  fileId: string
  questions: Question[]
  selected: boolean
}) {
  const { fileId, questions, selected } = payload
  if (selected) {
    for (const q of questions) {
      const exists = assemblyItems.value.some(
        (i) => i.fileId === fileId && i.question.id === q.id
      )
      if (!exists) {
        assemblyItems.value.push({ fileId, question: q, clientKey: newClientKey() })
      }
    }
  } else {
    assemblyItems.value = assemblyItems.value.filter((i) => i.fileId !== fileId)
  }
}

function removePaper(fileId: string) {
  stopPoll(fileId)
  papers.value = papers.value.filter((p) => p.fileId !== fileId)
  assemblyItems.value = assemblyItems.value.filter((i) => i.fileId !== fileId)
}

const downloadHtml = () => {
  const blob = new Blob([htmlContent.value], { type: 'text/html' })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = `document_${htmlFileId.value}.html`
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

/* 组卷导出卡片由 ExportPanel 自带样式，此处排除避免叠层 */
.upload-view :deep(.el-card:not(.export-panel-card)) {
  border-radius: var(--md-radius, 12px);
  border: 1px solid var(--md-border, #e4e7ef);
  box-shadow: var(--md-shadow-sm, 0 1px 2px rgba(15, 23, 42, 0.05));
}

.upload-view :deep(.el-card__header) {
  padding: 14px 18px;
  border-bottom: 1px solid var(--md-border, #e4e7ef);
  background: linear-gradient(180deg, #fafbff 0%, #fff 100%);
}

.upload-card,
.html-preview-card,
.html-wait-card,
.paper-card {
  margin-bottom: 18px;
}

.upload-card--compact {
  border-left: 3px solid var(--md-primary, #5b6ee8);
}

.upload-card--compact :deep(.el-card__body) {
  padding: 12px 18px 14px;
}

.upload-inline-row {
  display: flex;
  flex-wrap: nowrap;
  align-items: center;
  gap: 10px 14px;
  overflow-x: auto;
  padding-bottom: 2px;
  -webkit-overflow-scrolling: touch;
}

.upload-inline-label {
  font-weight: 600;
  color: var(--md-text, #1f2937);
  font-size: 14px;
  flex-shrink: 0;
}

.upload-mode-tag {
  flex-shrink: 0;
  border-radius: 999px !important;
}

.upload-compact-upload-wrap {
  flex: 1 1 auto;
  min-width: 0;
  display: inline-flex;
  flex-direction: row;
  flex-wrap: nowrap;
  align-items: center;
  gap: 10px;
}

.upload-compact-upload-wrap :deep(.el-upload) {
  display: inline-block;
  width: auto;
}

.upload-tip-compact {
  flex-shrink: 0;
  font-size: 12px;
  color: var(--md-text-secondary, #64748b);
  line-height: 32px;
  white-space: nowrap;
}

.upload-compact-upload-wrap .upload-area--compact.is-parsing :deep(.el-button) {
  border-color: #c7d4fc;
  background: var(--md-primary-soft, #eef0fc);
}

.parse-status-alert {
  margin-bottom: 16px;
  border-radius: var(--md-radius-sm, 8px);
  border: 1px solid #d9e2ff;
}

.parse-status-title {
  font-weight: 600;
}

.parse-status-desc {
  margin: 4px 0 0;
  font-size: 13px;
  line-height: 1.55;
  color: var(--md-text-secondary, #64748b);
}

.upload-area :deep(.el-upload-dragger) {
  border-radius: var(--md-radius, 12px);
  border: 2px dashed var(--md-border-strong, #dcdfe8);
  background: #fafbfd;
  transition:
    border-color 0.2s ease,
    background 0.2s ease,
    box-shadow 0.2s ease;
}

.upload-area :deep(.el-upload-dragger:hover) {
  border-color: var(--md-primary, #5b6ee8);
  background: var(--md-primary-soft, #eef0fc);
}

.upload-area.is-parsing :deep(.el-upload-dragger) {
  border-color: #b8c5f5;
  background: #f5f7ff;
}

.assembly-sticky-wrap {
  position: sticky;
  bottom: 16px;
  z-index: 20;
  padding-top: 10px;
}

.assembly-toolbar {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  margin-bottom: 10px;
  padding: 10px 14px;
  background: linear-gradient(135deg, #f0f3ff 0%, #eef6ff 100%);
  border: 1px solid #dce4f7;
  border-radius: var(--md-radius-sm, 8px);
  box-shadow: var(--md-shadow-sm);
}

.assembly-toolbar-text {
  font-size: 13px;
  color: var(--md-text-secondary, #64748b);
}

.assembly-toolbar-text strong {
  color: var(--md-primary, #5b6ee8);
  font-weight: 700;
}

.paper-skeleton-hint {
  margin: 0 0 12px;
  font-size: 13px;
  color: var(--md-text-secondary, #64748b);
}

.card-header,
.paper-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  font-size: 15px;
  font-weight: 600;
  color: var(--md-text, #1f2937);
}

.paper-header {
  gap: 12px;
}

.paper-title {
  display: flex;
  align-items: center;
  gap: 10px;
  min-width: 0;
}

.paper-title .el-icon {
  color: var(--md-primary, #5b6ee8);
}

.paper-name {
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  max-width: min(420px, 55vw);
}

.paper-skeleton {
  padding: 8px 0;
}

.mode-selector {
  margin-bottom: 20px;
  padding: 16px 18px;
  background: var(--md-primary-soft, #eef0fc);
  border: 1px solid #e0e5f7;
  border-radius: var(--md-radius-sm, 8px);
}

.mode-label {
  font-weight: 600;
  margin-right: 10px;
  color: var(--md-text-secondary, #64748b);
  font-size: 13px;
}

.mode-description {
  margin-top: 12px;
  font-size: 13px;
  color: var(--md-text-secondary, #64748b);
  display: flex;
  align-items: center;
  gap: 8px;
  line-height: 1.5;
}

.upload-area {
  width: 100%;
}

.upload-icon {
  font-size: 56px;
  color: #b4b9c9;
  margin-bottom: 12px;
}

.upload-text {
  color: var(--md-text-secondary, #64748b);
  font-size: 14px;
}

.multi-hint {
  display: block;
  margin-top: 6px;
  font-size: 12px;
  color: #94a3b8;
}

.upload-text em {
  color: var(--md-primary, #5b6ee8);
  font-style: normal;
  font-weight: 600;
}

.upload-tip {
  margin-top: 14px;
  color: #94a3b8;
  font-size: 12px;
}

.html-content {
  padding: 22px;
  background: #fff;
  border: 1px solid var(--md-border, #e4e7ef);
  border-radius: var(--md-radius-sm, 8px);
  max-height: 800px;
  overflow-y: auto;
  line-height: 1.85;
  color: var(--md-text, #1f2937);
}

.html-content :deep(img) {
  max-width: 100%;
  height: 20px;
  vertical-align: middle;
  margin: 10px 0;
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

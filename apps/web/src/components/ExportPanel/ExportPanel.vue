<template>
  <el-card class="export-panel-card">
    <template #header>
      <div class="card-header">
        <span>
          <el-icon><Download /></el-icon>
          组卷导出
        </span>
        <el-tag type="primary">已选 {{ assemblyItems.length }} 题</el-tag>
      </div>
    </template>

    <p class="hint">请先预览并拖动调整顺序，再导出 Word。</p>

    <el-form :model="exportForm" label-position="top">
      <el-form-item label="文档标题">
        <el-input
          v-model="exportForm.title"
          placeholder="请输入文档标题"
          maxlength="50"
          show-word-limit
        />
      </el-form-item>

      <el-form-item label="水印文字">
        <el-input
          v-model="exportForm.watermark"
          placeholder="请输入水印文字（可选）"
          maxlength="20"
          show-word-limit
        />
      </el-form-item>

      <el-form-item>
        <el-checkbox v-model="exportForm.includeAnswer"> 包含答案 </el-checkbox>
      </el-form-item>

      <el-form-item>
        <el-checkbox v-model="exportForm.includeAnalysis"> 包含解析 </el-checkbox>
      </el-form-item>
    </el-form>

    <div class="export-actions">
      <el-button type="primary" size="large" @click="openPreview">
        <el-icon><View /></el-icon>
        预览并调整顺序
      </el-button>
    </div>

    <el-dialog
      v-model="previewVisible"
      title="组卷预览（拖动排序 / 可移除）"
      width="min(920px, 96vw)"
      destroy-on-close
      class="assembly-preview-dialog"
      @opened="onPreviewOpened"
    >
      <p class="preview-tip">
        拖拽左侧手柄调整顺序；不需要的题目可点「移除」，将同步取消勾选。
      </p>
      <el-empty
        v-if="previewRows.length === 0"
        description="预览中暂无题目，关闭后可重新选题加入组卷"
        :image-size="72"
      />
      <div v-else class="preview-list">
        <div
          v-for="(row, index) in previewRows"
          :key="row.clientKey"
          class="preview-row"
          :class="{ 'is-drag-over': dragOverIndex === index }"
          draggable="true"
          @dragstart="onDragStart($event, index)"
          @dragover.prevent="onDragOver($event, index)"
          @drop.prevent="onDrop(index)"
          @dragleave="onDragLeave($event, index)"
          @dragend="onDragEnd"
        >
          <span class="drag-handle" title="拖动排序" aria-hidden="true">
            <el-icon><Rank /></el-icon>
          </span>
          <span class="preview-index">{{ index + 1 }}.</span>
          <div class="preview-body">
            <div class="preview-meta">
              <el-tag size="small" type="info">{{ row.sourceLabel }}</el-tag>
              <span class="preview-type">{{ row.question.type_name }}</span>
            </div>
            <div class="preview-content">
              <math-renderer :content="row.question.content_html" />
            </div>
          </div>
          <div class="preview-row-actions">
            <el-button
              type="danger"
              text
              size="small"
              title="移出本次组卷"
              @click.stop="removePreviewRow(index)"
            >
              <el-icon><Close /></el-icon>
              移除
            </el-button>
          </div>
        </div>
      </div>
      <template #footer>
        <el-button @click="previewVisible = false">取消</el-button>
        <el-button type="primary" :loading="exporting" @click="handleExport">
          <el-icon><Download /></el-icon>
          导出 Word
        </el-button>
      </template>
    </el-dialog>
  </el-card>
</template>

<script setup lang="ts">
import { reactive, ref, nextTick } from 'vue'
import { ElMessage } from 'element-plus'
import { Download, View, Rank, Close } from '@element-plus/icons-vue'
import MathRenderer from '../MathRenderer/MathRenderer.vue'
import type { Question } from '../../types'
import { waitForMathJaxReady } from '../../utils/mathjaxReady'

export interface AssemblyRow {
  fileId: string
  sourceLabel: string
  question: Question
  clientKey: string
}

interface Props {
  assemblyItems: AssemblyRow[]
}

const props = defineProps<Props>()

const emit = defineEmits<{
  (
    e: 'update-assembly',
    payload: Array<{ fileId: string; question: Question; clientKey: string }>
  ): void
}>()

const exportForm = reactive({
  title: '导出的题目',
  watermark: '',
  includeAnswer: false,
  includeAnalysis: false
})

const previewVisible = ref(false)
const previewRows = ref<AssemblyRow[]>([])
const exporting = ref(false)
const dragFromIndex = ref<number | null>(null)
const dragOverIndex = ref<number | null>(null)

function emitAssemblySync() {
  emit(
    'update-assembly',
    previewRows.value.map(({ fileId, question, clientKey }) => ({
      fileId,
      question,
      clientKey
    }))
  )
}

const openPreview = () => {
  if (props.assemblyItems.length === 0) {
    ElMessage.warning('请先选择题目')
    return
  }
  previewVisible.value = true
}

const syncPreviewRows = () => {
  previewRows.value = props.assemblyItems.map((r) => ({ ...r }))
}

const onPreviewOpened = () => {
  syncPreviewRows()
  nextTick(async () => {
    const el = document.querySelector('.assembly-preview-dialog .preview-list')
    if (!el) return
    const ready = await waitForMathJaxReady()
    if (!ready) return
    const mj = (window as unknown as { MathJax?: { typesetPromise?: (nodes: Element[]) => Promise<void> } })
      .MathJax
    if (mj?.typesetPromise) {
      mj.typesetPromise([el]).catch(() => {})
    }
  })
}

const onDragStart = (e: DragEvent, index: number) => {
  dragFromIndex.value = index
  e.dataTransfer?.setData('text/plain', String(index))
  e.dataTransfer!.effectAllowed = 'move'
}

const onDragOver = (e: DragEvent, index: number) => {
  e.dataTransfer!.dropEffect = 'move'
  dragOverIndex.value = index
}

const onDragLeave = (e: DragEvent, index: number) => {
  const related = e.relatedTarget as Node | null
  if (related && (e.currentTarget as HTMLElement).contains(related)) return
  if (dragOverIndex.value === index) dragOverIndex.value = null
}

const onDrop = (toIndex: number) => {
  const from = dragFromIndex.value
  dragOverIndex.value = null
  if (from === null || from === toIndex) return
  const next = [...previewRows.value]
  const [moved] = next.splice(from, 1)
  let insertAt = toIndex
  if (from < insertAt) insertAt -= 1
  next.splice(insertAt, 0, moved)
  previewRows.value = next
  dragFromIndex.value = null
  emitAssemblySync()
}

const onDragEnd = () => {
  dragFromIndex.value = null
  dragOverIndex.value = null
}

function removePreviewRow(index: number) {
  previewRows.value.splice(index, 1)
  emitAssemblySync()
  if (previewRows.value.length === 0) {
    previewVisible.value = false
    ElMessage.info('预览中的题目已全部移除')
  }
}

const handleExport = async () => {
  if (previewRows.value.length === 0) {
    ElMessage.warning('没有可导出的题目')
    return
  }

  exporting.value = true
  try {
    const response = await fetch('/api/v1/export', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        assembly: previewRows.value.map((r) => ({
          file_id: r.fileId,
          question_id: r.question.id
        })),
        title: exportForm.title,
        options: {
          include_answer: exportForm.includeAnswer,
          include_analysis: exportForm.includeAnalysis,
          watermark: exportForm.watermark?.trim() || undefined
        }
      })
    })

    if (!response.ok) {
      let message = '导出失败'
      try {
        const data = await response.json()
        message = data?.message || message
      } catch {
        // ignore
      }
      throw new Error(message)
    }

    const blob = await response.blob()
    const url = URL.createObjectURL(blob)
    const disposition = response.headers.get('content-disposition') || ''
    let filename = '导出的题目.docx'
    const match = disposition.match(/filename\*?=(?:UTF-8''|")?([^;"]+)/i)
    if (match?.[1]) {
      filename = decodeURIComponent(match[1].trim())
    }
    const downloadLink = document.createElement('a')
    downloadLink.href = url
    downloadLink.download = filename
    document.body.appendChild(downloadLink)
    downloadLink.click()
    document.body.removeChild(downloadLink)
    URL.revokeObjectURL(url)

    ElMessage.success('导出成功')
    previewVisible.value = false
  } catch (error) {
    console.error('导出失败:', error)
    ElMessage.error(error instanceof Error ? error.message : '导出失败，请稍后重试')
  } finally {
    exporting.value = false
  }
}
</script>

<style scoped>
.export-panel-card {
  border-radius: var(--md-radius, 12px);
  border: 1px solid var(--md-border, #e4e7ef);
  box-shadow: var(--md-shadow-md, 0 4px 20px rgba(15, 23, 42, 0.06));
}

.card-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  font-size: 15px;
  font-weight: 600;
  color: var(--md-text, #1f2937);
}

.card-header .el-icon {
  margin-right: 5px;
  color: var(--md-primary, #5b6ee8);
}

.hint {
  margin: 0 0 14px;
  font-size: 13px;
  line-height: 1.55;
  color: var(--md-text-secondary, #64748b);
}

.export-actions {
  margin-top: 22px;
  text-align: center;
}

.export-actions .el-button {
  min-width: 200px;
  font-weight: 600;
}

.preview-tip {
  margin: 0 0 12px;
  font-size: 13px;
  color: var(--md-text-secondary, #64748b);
}

.preview-list {
  display: flex;
  flex-direction: column;
  gap: 10px;
  max-height: 62vh;
  overflow-y: auto;
  padding: 4px 6px 4px 0;
}

.preview-row-actions {
  flex-shrink: 0;
  align-self: flex-start;
  padding-top: 2px;
}

.preview-row {
  display: flex;
  align-items: flex-start;
  gap: 10px;
  padding: 14px;
  border: 1px solid var(--md-border, #e4e7ef);
  border-radius: var(--md-radius-sm, 8px);
  background: #fafbfd;
  cursor: grab;
  transition:
    border-color 0.15s ease,
    background 0.15s ease,
    box-shadow 0.15s ease;
}

.preview-row:active {
  cursor: grabbing;
}

.preview-row.is-drag-over {
  border-color: var(--md-primary, #5b6ee8);
  background: var(--md-primary-soft, #eef0fc);
  box-shadow: 0 0 0 2px rgba(91, 110, 232, 0.15);
}

.drag-handle {
  flex-shrink: 0;
  color: var(--md-text-secondary, #64748b);
  user-select: none;
  display: flex;
  align-items: center;
  padding-top: 2px;
  cursor: grab;
}

.drag-handle .el-icon {
  font-size: 18px;
}

.preview-index {
  flex-shrink: 0;
  font-weight: 700;
  color: var(--md-primary, #5b6ee8);
  min-width: 1.5em;
}

.preview-body {
  flex: 1;
  min-width: 0;
}

.preview-meta {
  display: flex;
  align-items: center;
  gap: 8px;
  margin-bottom: 8px;
}

.preview-type {
  font-size: 12px;
  color: var(--md-text-secondary, #64748b);
}

.preview-content {
  line-height: 1.7;
  color: var(--md-text, #1f2937);
}
</style>

<style>
.assembly-preview-dialog .el-dialog {
  border-radius: 12px;
  overflow: hidden;
}

.assembly-preview-dialog .el-dialog__header {
  padding: 16px 20px 12px;
  margin: 0;
  border-bottom: 1px solid var(--md-border, #e4e7ef);
}

.assembly-preview-dialog .el-dialog__body {
  padding: 12px 20px 8px;
}
</style>

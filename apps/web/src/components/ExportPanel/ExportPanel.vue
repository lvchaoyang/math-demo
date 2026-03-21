<template>
  <el-card class="export-panel-card">
    <template #header>
      <div class="card-header">
        <span>
          <el-icon><Download /></el-icon>
          导出设置
        </span>
        <el-tag type="primary">已选择 {{ selectedCount }} 道题目</el-tag>
      </div>
    </template>
    
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
        <el-checkbox v-model="exportForm.includeAnswer">
          包含答案
        </el-checkbox>
      </el-form-item>
      
      <el-form-item>
        <el-checkbox v-model="exportForm.includeAnalysis">
          包含解析
        </el-checkbox>
      </el-form-item>
    </el-form>
    
    <div class="export-actions">
      <el-button
        type="primary"
        size="large"
        :loading="exporting"
        @click="handleExport"
      >
        <el-icon><Download /></el-icon>
        导出 Word 文档
      </el-button>
    </div>
  </el-card>
</template>

<script setup lang="ts">
import { reactive, ref } from 'vue'
import { ElMessage } from 'element-plus'
import { Download } from '@element-plus/icons-vue'

interface Props {
  selectedCount: number
  selectedIds: string[]
}

const props = defineProps<Props>()

const emit = defineEmits<{
  (e: 'export-success'): void
}>()

const exportForm = reactive({
  title: '导出的题目',
  watermark: '',
  includeAnswer: false,
  includeAnalysis: false,
  paperSize: 'A4'
})

const exporting = ref(false)

const handleExport = async () => {
  if (props.selectedIds.length === 0) {
    ElMessage.warning('请先选择要导出的题目')
    return
  }
  
  exporting.value = true
  
  try {
    const response = await fetch('/api/v1/export', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        question_ids: props.selectedIds,
        title: exportForm.title,
        watermark: exportForm.watermark || undefined,
        include_answer: exportForm.includeAnswer,
        include_analysis: exportForm.includeAnalysis,
        paper_size: exportForm.paperSize
      })
    })
    
    const data = await response.json()
    
    if (data.success) {
      // 触发下载
      const downloadLink = document.createElement('a')
      downloadLink.href = data.download_url
      downloadLink.download = '导出的题目.docx'
      document.body.appendChild(downloadLink)
      downloadLink.click()
      document.body.removeChild(downloadLink)
      
      ElMessage.success('导出成功')
      emit('export-success')
    } else {
      ElMessage.error(data.message || '导出失败')
    }
  } catch (error) {
    console.error('导出失败:', error)
    ElMessage.error('导出失败，请稍后重试')
  } finally {
    exporting.value = false
  }
}
</script>

<style scoped>
.export-panel-card {
  position: sticky;
  bottom: 20px;
  box-shadow: 0 -4px 20px rgba(0, 0, 0, 0.1);
}

.card-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  font-size: 16px;
  font-weight: 600;
}

.card-header .el-icon {
  margin-right: 5px;
}

.export-actions {
  margin-top: 20px;
  text-align: center;
}

.export-actions .el-button {
  width: 200px;
}
</style>

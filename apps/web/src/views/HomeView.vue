<template>
  <div class="home-view">
    <el-row :gutter="20">
      <el-col :span="24">
        <el-card class="welcome-card" shadow="never">
          <template #header>
            <div class="card-header">
              <span>欢迎使用数学试卷解析系统</span>
            </div>
          </template>
          <div class="welcome-content">
            <div class="welcome-icon-wrap">
              <el-icon class="welcome-icon"><Document /></el-icon>
            </div>
            <h2>智能试卷解析与题目管理</h2>
            <p class="welcome-desc">
              上传 Word 试卷，自动解析公式并拆分题目；支持多份试卷跨卷选题、预览调序后导出 Word。
            </p>
            <el-button type="primary" size="large" round class="welcome-cta" @click="$router.push('/upload')">
              <el-icon><Upload /></el-icon>
              开始上传试卷
            </el-button>
          </div>
        </el-card>
      </el-col>
    </el-row>

    <el-row :gutter="20" class="feature-row">
      <el-col :xs="24" :sm="8">
        <el-card class="feature-card" shadow="hover">
          <el-icon class="feature-icon"><DocumentChecked /></el-icon>
          <h3>智能解析</h3>
          <p>自动识别 Word 文档中的数学公式（OMML 格式），转换为 LaTeX 格式</p>
        </el-card>
      </el-col>
      <el-col :xs="24" :sm="8">
        <el-card class="feature-card" shadow="hover">
          <el-icon class="feature-icon"><List /></el-icon>
          <h3>题目拆分</h3>
          <p>按题型自动拆分题目，支持选择题、填空题、解答题等多种题型</p>
        </el-card>
      </el-col>
      <el-col :xs="24" :sm="8">
        <el-card class="feature-card" shadow="hover">
          <el-icon class="feature-icon"><Download /></el-icon>
          <h3>灵活导出</h3>
          <p>预览组卷、拖拽排序，导出 Word 并支持水印与答案/解析选项</p>
        </el-card>
      </el-col>
    </el-row>
  </div>
</template>

<script setup lang="ts">
import { Document, Upload, DocumentChecked, List, Download } from '@element-plus/icons-vue'
</script>

<style scoped>
.home-view {
  max-width: 1200px;
  margin: 0 auto;
}

.welcome-card {
  text-align: center;
  margin-bottom: 24px;
  border-radius: var(--md-radius, 12px);
  border: 1px solid var(--md-border, #e4e7ef);
  overflow: hidden;
}

.welcome-card :deep(.el-card__header) {
  background: linear-gradient(180deg, #fafbff 0%, #fff 100%);
  border-bottom: 1px solid var(--md-border, #e4e7ef);
  padding: 16px 20px;
}

.card-header {
  font-size: 16px;
  font-weight: 600;
  color: var(--md-text, #1f2937);
}

.welcome-content {
  padding: 36px 24px 40px;
}

.welcome-icon-wrap {
  width: 88px;
  height: 88px;
  margin: 0 auto 20px;
  border-radius: 24px;
  background: var(--md-primary-soft, #eef0fc);
  display: flex;
  align-items: center;
  justify-content: center;
}

.welcome-icon {
  font-size: 44px;
  color: var(--md-primary, #5b6ee8);
}

.welcome-content h2 {
  margin: 0 0 12px 0;
  font-size: 22px;
  font-weight: 700;
  color: var(--md-text, #1f2937);
  letter-spacing: -0.02em;
}

.welcome-desc {
  color: var(--md-text-secondary, #64748b);
  margin: 0 auto 28px;
  font-size: 15px;
  line-height: 1.65;
  max-width: 520px;
}

.welcome-cta {
  padding: 12px 28px;
  font-weight: 600;
}

.feature-row {
  margin-top: 8px;
}

.feature-row :deep(.el-col) {
  margin-bottom: 16px;
}

.feature-card {
  text-align: center;
  padding: 28px 20px;
  height: 100%;
  border-radius: var(--md-radius, 12px);
  border: 1px solid var(--md-border, #e4e7ef);
  transition:
    transform 0.25s ease,
    box-shadow 0.25s ease,
    border-color 0.25s ease;
}

.feature-card:hover {
  transform: translateY(-4px);
  box-shadow: var(--md-shadow-lg, 0 12px 40px rgba(15, 23, 42, 0.08));
  border-color: var(--md-border-strong, #dcdfe8);
}

.feature-icon {
  font-size: 40px;
  color: var(--md-primary, #5b6ee8);
  margin-bottom: 16px;
}

.feature-card h3 {
  margin: 0 0 10px 0;
  font-size: 17px;
  font-weight: 600;
  color: var(--md-text, #1f2937);
}

.feature-card p {
  color: var(--md-text-secondary, #64748b);
  line-height: 1.65;
  font-size: 14px;
  margin: 0;
}
</style>

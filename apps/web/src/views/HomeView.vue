<template>
  <div class="home-view">
    <el-row :gutter="20">
      <el-col :span="24">
        <el-card class="welcome-card">
          <template #header>
            <div class="card-header">
              <span>欢迎使用数学试卷解析系统</span>
            </div>
          </template>
          <div class="welcome-content">
            <el-icon class="welcome-icon"><Document /></el-icon>
            <h2>智能试卷解析与题目管理</h2>
            <p>上传 Word 格式的数学试卷，系统自动解析数学公式，拆分题目，支持选择性导出</p>
            <el-button type="primary" size="large" @click="$router.push('/upload')">
              <el-icon><Upload /></el-icon>
              开始上传试卷
            </el-button>
          </div>
        </el-card>
      </el-col>
    </el-row>
    
    <el-row :gutter="20" class="feature-row">
      <el-col :span="8">
        <el-card class="feature-card">
          <el-icon class="feature-icon"><DocumentChecked /></el-icon>
          <h3>智能解析</h3>
          <p>自动识别 Word 文档中的数学公式（OMML格式），转换为 LaTeX 格式</p>
        </el-card>
      </el-col>
      <el-col :span="8">
        <el-card class="feature-card">
          <el-icon class="feature-icon"><List /></el-icon>
          <h3>题目拆分</h3>
          <p>按题型自动拆分题目，支持选择题、填空题、解答题等多种题型</p>
        </el-card>
      </el-col>
      <el-col :span="8">
        <el-card class="feature-card">
          <el-icon class="feature-icon"><Download /></el-icon>
          <h3>灵活导出</h3>
          <p>勾选需要的题目，导出为 Word 文档，支持自定义水印</p>
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
  margin-bottom: 30px;
}

.card-header {
  font-size: 18px;
  font-weight: 600;
}

.welcome-content {
  padding: 40px 20px;
}

.welcome-icon {
  font-size: 80px;
  color: #667eea;
  margin-bottom: 20px;
}

.welcome-content h2 {
  margin: 0 0 15px 0;
  color: #303133;
}

.welcome-content p {
  color: #606266;
  margin-bottom: 30px;
  font-size: 16px;
}

.feature-row {
  margin-top: 20px;
}

.feature-card {
  text-align: center;
  padding: 30px 20px;
  height: 100%;
  transition: transform 0.3s, box-shadow 0.3s;
}

.feature-card:hover {
  transform: translateY(-5px);
  box-shadow: 0 10px 30px rgba(0, 0, 0, 0.1);
}

.feature-icon {
  font-size: 50px;
  color: #667eea;
  margin-bottom: 20px;
}

.feature-card h3 {
  margin: 0 0 15px 0;
  color: #303133;
}

.feature-card p {
  color: #606266;
  line-height: 1.6;
}
</style>

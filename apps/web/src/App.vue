<template>
  <div id="app">
    <el-container>
      <el-header class="app-header">
        <div class="header-content">
          <h1 class="app-title">
            <el-icon><Document /></el-icon>
            数学试卷解析系统
          </h1>
          <nav class="nav-menu">
            <router-link to="/" class="nav-link">首页</router-link>
            <router-link to="/upload" class="nav-link">上传试卷</router-link>
          </nav>
        </div>
      </el-header>
      
      <el-main class="app-main">
        <router-view />
      </el-main>
      
      <el-footer class="app-footer">
        <p>数学试卷解析系统 &copy; 2024</p>
      </el-footer>
    </el-container>
  </div>
</template>

<script setup lang="ts">
import { Document } from '@element-plus/icons-vue'
</script>

<style scoped>
.app-header {
  background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
  color: white;
  padding: 0;
  height: 60px;
  line-height: 60px;
}

.header-content {
  max-width: 1200px;
  margin: 0 auto;
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 0 20px;
}

.app-title {
  margin: 0;
  font-size: 24px;
  font-weight: 600;
  display: flex;
  align-items: center;
  gap: 10px;
}

.nav-menu {
  display: flex;
  gap: 30px;
}

.nav-link {
  color: white;
  text-decoration: none;
  font-size: 16px;
  opacity: 0.9;
  transition: opacity 0.3s;
}

.nav-link:hover,
.nav-link.router-link-active {
  opacity: 1;
  font-weight: 500;
}

.app-main {
  min-height: calc(100vh - 120px);
  background: #f5f7fa;
  padding: 20px;
}

.app-footer {
  background: #2c3e50;
  color: white;
  text-align: center;
  height: 60px;
  line-height: 60px;
}

.app-footer p {
  margin: 0;
}
</style>

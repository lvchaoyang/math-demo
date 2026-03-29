<template>
  <div id="app">
    <el-container class="app-shell">
      <el-header class="app-header">
        <div class="header-content">
          <h1 class="app-title">
            <span class="app-title-mark">
              <el-icon><Document /></el-icon>
            </span>
            数学试卷解析系统
          </h1>
          <nav class="nav-menu" aria-label="主导航">
            <router-link to="/" class="nav-link">首页</router-link>
            <router-link to="/upload" class="nav-link">上传试卷</router-link>
          </nav>
        </div>
      </el-header>

      <el-main class="app-main">
        <router-view />
      </el-main>

      <el-footer class="app-footer">
        <p>数学试卷解析系统 &copy; 2026</p>
      </el-footer>
    </el-container>
  </div>
</template>

<script setup lang="ts">
import { Document } from '@element-plus/icons-vue'
</script>

<style scoped>
.app-shell {
  min-height: 100vh;
}

.app-header {
  background: linear-gradient(
    125deg,
    var(--md-primary, #5b6ee8) 0%,
    var(--md-accent, #7c6ae8) 48%,
    #8b5fd4 100%
  );
  color: #fff;
  padding: 0;
  height: 56px;
  line-height: 56px;
  box-shadow: var(--md-shadow-md, 0 4px 20px rgba(15, 23, 42, 0.12));
}

.header-content {
  max-width: 1200px;
  margin: 0 auto;
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 0 22px;
  height: 100%;
}

.app-title {
  margin: 0;
  font-size: 18px;
  font-weight: 600;
  letter-spacing: 0.02em;
  display: flex;
  align-items: center;
  gap: 10px;
}

.app-title-mark {
  display: flex;
  align-items: center;
  justify-content: center;
  width: 34px;
  height: 34px;
  border-radius: 10px;
  background: rgba(255, 255, 255, 0.18);
  backdrop-filter: blur(6px);
}

.app-title-mark .el-icon {
  font-size: 20px;
}

.nav-menu {
  display: flex;
  align-items: center;
  gap: 8px;
}

.nav-link {
  color: rgba(255, 255, 255, 0.92);
  text-decoration: none;
  font-size: 14px;
  padding: 6px 14px;
  border-radius: 999px;
  transition:
    background 0.2s ease,
    color 0.2s ease;
}

.nav-link:hover {
  background: rgba(255, 255, 255, 0.14);
  color: #fff;
}

.nav-link.router-link-active {
  background: rgba(255, 255, 255, 0.22);
  color: #fff;
  font-weight: 600;
}

.app-main {
  min-height: calc(100vh - 112px);
  background: var(--md-page-bg, #f0f2f8);
  padding: 24px 20px 32px;
}

.app-footer {
  background: #1e293b;
  color: rgba(255, 255, 255, 0.72);
  text-align: center;
  height: 56px;
  line-height: 56px;
  font-size: 13px;
}

.app-footer p {
  margin: 0;
}
</style>

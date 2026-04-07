import './mathjax-bootstrap'
/* SVG 输出不依赖 CHTML 的 woff2 字体路径；Vite 打包后 CHTML 常因字体未正确加载出现根号横线错位、竖线像斜杠 */
import 'mathjax/es5/tex-mml-svg.js'

import { createApp } from 'vue'
import { createPinia } from 'pinia'
import ElementPlus from 'element-plus'
import 'element-plus/dist/index.css'
import './styles/theme.css'
import * as ElementPlusIconsVue from '@element-plus/icons-vue'

import App from './App.vue'
import router from './router'

const app = createApp(App)

// 注册所有图标
for (const [key, component] of Object.entries(ElementPlusIconsVue)) {
  app.component(key, component)
}

app.use(createPinia())
app.use(router)
app.use(ElementPlus)

app.mount('#app')

// router/index.js
import { createRouter, createWebHistory } from 'vue-router'
import TwoColumnPage from '@/views/TwoColumnPage.vue'
import ThreeColumnPage from '@/views/ThreeColumnPage.vue'
import EssayScoringPage from '@/views/EssayScoringPage.vue'

const routes = [
  {
    path: '/two-column',
    name: 'TwoColumn',
    component: TwoColumnPage
  },
  {
    path: '/three-column',
    name: 'ThreeColumn',
    component: ThreeColumnPage
  },
  {
    path: '/essay-scoring',
    name: 'EssayScoring',
    component: EssayScoringPage
  },
  {
    path: '/',
    redirect: '/two-column' // 默认跳转到两栏页面
  }
]

const router = createRouter({
  history: createWebHistory(),
  routes
})

export default router
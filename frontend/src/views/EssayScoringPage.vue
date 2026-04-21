<template>
  <div class="essay-scoring-page">
    <iframe
      :src="iframeSrc"
      frameborder="0"
      style="width: 100%; height: 100%; border: none;"
    ></iframe>
  </div>
</template>

<script setup>
import { computed } from 'vue'
import { getStaticUrl } from '@/utils/config'

const iframeSrc = computed(() => {
  return getStaticUrl('essay_scoring_demo.html')
})
</script>

<style scoped>
.essay-scoring-page {
  width: 100%;
  height: calc(100vh - 60px);
  overflow: hidden;
}
</style>

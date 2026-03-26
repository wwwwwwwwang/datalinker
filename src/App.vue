<script setup lang="ts">
import ContrastTab from "./components/ContrastTab.vue";
</script>

<template>
  <div class="app">
    <ContrastTab />
  </div>
</template>

<style>
.app {
  padding: 16px;
}

.form-row {
  display: flex;
  align-items: center;
  flex-wrap: wrap;
  gap: 10px;
  margin-bottom: 12px;
}

.form-row label {
  font-size: 14px;
  color: #333;
}

.path-input {
  width: 240px;
}

.threshold-input {
  width: 90px;
}

.action-row {
  display: flex;
  gap: 10px;
  margin-top: 12px;
}

.context-menu {
  position: fixed;
  z-index: 9999;
  background: #fff;
  border: 1px solid #dcdfe6;
  box-shadow: 0 2px 12px rgba(0, 0, 0, 0.12);
  border-radius: 4px;
  min-width: 120px;
}

.context-item {
  padding: 8px 12px;
  cursor: pointer;
  font-size: 13px;
}

.context-item:hover {
  background: #f5f7fa;
}
</style>

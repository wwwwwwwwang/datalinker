<script lang="ts">
export default {
  name: "ContrastTab"
};
</script>

<script setup lang="ts">
import { onBeforeUnmount, onMounted, reactive, ref } from "vue";
import { invoke } from "@tauri-apps/api/core";
import { open as openDialog } from "@tauri-apps/plugin-dialog";
import { openPath, revealItemInDir } from "@tauri-apps/plugin-opener";
import { ElMessage } from "element-plus";
import "element-plus/es/components/message/style/css";
import type { ContrastRow } from "../types";
import { loadContrastRows, saveContrastRows } from "../services/configStore";

type ContextMenuState = {
  visible: boolean;
  x: number;
  y: number;
};

const contrastForm = reactive({
  standardSamplePath: "",
  samplePath: "",
  analysisResultsPath: "",
  thresholdNumber: 1
});

const contrastRows = ref<ContrastRow[]>([]);
const currentContrastRow = ref<ContrastRow | null>(null);
const contextMenu = reactive<ContextMenuState>({
  visible: false,
  x: 0,
  y: 0
});

function showContextMenu(row: ContrastRow, event: MouseEvent) {
  event.preventDefault();
  currentContrastRow.value = row;
  contextMenu.x = event.clientX;
  contextMenu.y = event.clientY;
  contextMenu.visible = true;
}

function onContrastRowContextMenu(row: ContrastRow, _column: unknown, event: MouseEvent) {
  showContextMenu(row, event);
}

function hideContextMenu() {
  contextMenu.visible = false;
}

async function chooseStandardSamplePath() {
  const file = await openDialog({
    multiple: false,
    directory: false,
    filters: [{ name: "Excel", extensions: ["xlsx", "xls"] }]
  });
  if (typeof file === "string") {
    contrastForm.standardSamplePath = file;
  }
}

async function chooseSamplePath() {
  const file = await openDialog({
    multiple: false,
    directory: false,
    filters: [{ name: "Excel", extensions: ["xlsx", "xls"] }]
  });
  if (typeof file === "string") {
    contrastForm.samplePath = file;
  }
}

async function chooseAnalysisResultsPath() {
  const dir = await openDialog({
    multiple: false,
    directory: true
  });
  if (typeof dir === "string") {
    contrastForm.analysisResultsPath = dir;
  }
}

async function chooseRowStandardSamplePath(row: ContrastRow) {
  const file = await openDialog({
    multiple: false,
    directory: false,
    filters: [{ name: "Excel", extensions: ["xlsx", "xls"] }]
  });
  if (typeof file === "string") {
    row.standardSamplePath = file;
    await onContrastRowEdited();
  }
}

async function chooseRowSamplePath(row: ContrastRow) {
  const file = await openDialog({
    multiple: false,
    directory: false,
    filters: [{ name: "Excel", extensions: ["xlsx", "xls"] }]
  });
  if (typeof file === "string") {
    row.samplePath = file;
    await onContrastRowEdited();
  }
}

async function chooseRowAnalysisResultsPath(row: ContrastRow) {
  const dir = await openDialog({
    multiple: false,
    directory: true
  });
  if (typeof dir === "string") {
    row.analysisResultsPath = dir;
    await onContrastRowEdited();
  }
}

async function saveContrastConfig(options?: { silent?: boolean }) {
  try {
    await saveContrastRows(contrastRows.value);
    if (!options?.silent) {
      ElMessage.success("保存配置成功");
    }
  } catch (error) {
    ElMessage.error(`保存配置失败：${error}`);
  }
}

async function onSaveContrastConfigClick() {
  await saveContrastConfig();
}

async function loadContrastConfig() {
  try {
    contrastRows.value = await loadContrastRows();
  } catch (error) {
    ElMessage.error(`加载配置失败：${error}`);
  }
}

async function onContrastRowEdited() {
  await saveContrastConfig({ silent: true });
}

async function addContrastRow() {
  if (
    !contrastForm.standardSamplePath.trim()
    || !contrastForm.samplePath.trim()
    || !contrastForm.analysisResultsPath.trim()
  ) {
    ElMessage.warning("请先完整选择标准样本路径、样本路径、解析结果路径");
    return;
  }

  contrastRows.value.push({
    standardSamplePath: contrastForm.standardSamplePath,
    samplePath: contrastForm.samplePath,
    analysisResultsPath: contrastForm.analysisResultsPath,
    thresholdNumber: String(contrastForm.thresholdNumber),
    remarks: ""
  });
  await saveContrastConfig({ silent: true });
}

async function deleteSelectedRow() {
  if (!currentContrastRow.value) {
    return;
  }
  const index = contrastRows.value.indexOf(currentContrastRow.value);
  if (index >= 0) {
    contrastRows.value.splice(index, 1);
    currentContrastRow.value = null;
    await saveContrastConfig({ silent: true });
    hideContextMenu();
  }
}

async function copySelectedRow() {
  if (!currentContrastRow.value) {
    return;
  }
  const index = contrastRows.value.indexOf(currentContrastRow.value);
  const copy: ContrastRow = { ...currentContrastRow.value };
  if (index >= 0) {
    contrastRows.value.splice(index, 0, copy);
  } else {
    contrastRows.value.push(copy);
  }
  await saveContrastConfig({ silent: true });
  hideContextMenu();
}

async function deleteAllRows() {
  contrastRows.value = [];
  currentContrastRow.value = null;
  await saveContrastConfig({ silent: true });
  hideContextMenu();
}

async function runContrast(row: ContrastRow) {
  try {
    const outputFilePath = await invoke<string>("run_contrast", { row });
    ElMessage.success(`解析完毕，结果文件：${outputFilePath}（明细见“相同位点/差异位点/缺失位点”sheet）`);
  } catch (error) {
    ElMessage.error(String(error));
  }
}

function resolveParentDir(path: string) {
  const normalized = path.trim().replace(/[\\/]+$/, "");
  const lastSeparator = Math.max(normalized.lastIndexOf("\\"), normalized.lastIndexOf("/"));
  if (lastSeparator < 0) {
    return normalized;
  }

  let parent = normalized.slice(0, lastSeparator);
  if (/^[A-Za-z]:$/.test(parent)) {
    parent += "\\";
  }
  return parent;
}

async function openResultPath(path: string) {
  const target = path.trim();
  if (!target) {
    ElMessage.warning("路径为空");
    return;
  }

  try {
    await revealItemInDir(target);
    return;
  } catch {}

  try {
    await openPath(target);
    return;
  } catch {}

  const parentDir = resolveParentDir(target);
  if (parentDir && parentDir !== target) {
    try {
      await openPath(parentDir);
      return;
    } catch {}
  }

  ElMessage.error("无法打开路径");
}

function updateCurrentRow(row: ContrastRow | null) {
  currentContrastRow.value = row;
}

onMounted(async () => {
  await loadContrastConfig();
  document.addEventListener("click", hideContextMenu);
});

onBeforeUnmount(() => {
  document.removeEventListener("click", hideContextMenu);
});
</script>

<template>
  <section class="contrast-tab">
    <div class="form-row contrast-form-row">
      <label>标准样本路径：</label>
      <el-input v-model="contrastForm.standardSamplePath" readonly class="path-input">
        <template #append>
          <el-button @click="chooseStandardSamplePath">选择</el-button>
        </template>
      </el-input>

      <label>样本路径：</label>
      <el-input v-model="contrastForm.samplePath" readonly class="path-input">
        <template #append>
          <el-button @click="chooseSamplePath">选择</el-button>
        </template>
      </el-input>

      <label>解析结果路径：</label>
      <el-input v-model="contrastForm.analysisResultsPath" readonly class="path-input">
        <template #append>
          <el-button @click="chooseAnalysisResultsPath">选择</el-button>
        </template>
      </el-input>

      <label>阈值：</label>
      <el-input-number v-model="contrastForm.thresholdNumber" :min="1" :step="1" class="threshold-input" />

      <el-button type="primary" @click="addContrastRow">添加</el-button>
    </div>

    <el-table
      :data="contrastRows"
      table-layout="fixed"
      highlight-current-row
      @current-change="updateCurrentRow"
      @row-contextmenu="onContrastRowContextMenu"
    >
      <el-table-column prop="standardSamplePath" label="标准样本路径" min-width="150">
        <template #default="scope">
          <div class="path-cell">
            <el-input v-model="scope.row.standardSamplePath" size="small" readonly />
            <el-button size="small" @click="chooseRowStandardSamplePath(scope.row)">选择</el-button>
          </div>
        </template>
      </el-table-column>
      <el-table-column label="查看" width="64">
        <template #default="scope">
          <el-button size="small" @click="openResultPath(scope.row.standardSamplePath)">查看</el-button>
        </template>
      </el-table-column>

      <el-table-column prop="samplePath" label="样本路径" min-width="150">
        <template #default="scope">
          <div class="path-cell">
            <el-input v-model="scope.row.samplePath" size="small" readonly />
            <el-button size="small" @click="chooseRowSamplePath(scope.row)">选择</el-button>
          </div>
        </template>
      </el-table-column>
      <el-table-column label="查看" width="64">
        <template #default="scope">
          <el-button size="small" @click="openResultPath(scope.row.samplePath)">查看</el-button>
        </template>
      </el-table-column>

      <el-table-column prop="analysisResultsPath" label="解析结果路径" min-width="150">
        <template #default="scope">
          <div class="path-cell">
            <el-input v-model="scope.row.analysisResultsPath" size="small" readonly />
            <el-button size="small" @click="chooseRowAnalysisResultsPath(scope.row)">选择</el-button>
          </div>
        </template>
      </el-table-column>
      <el-table-column label="查看" width="64">
        <template #default="scope">
          <el-button size="small" @click="openResultPath(scope.row.analysisResultsPath)">查看</el-button>
        </template>
      </el-table-column>

      <el-table-column prop="thresholdNumber" label="阈值" width="90">
        <template #default="scope">
          <el-input v-model="scope.row.thresholdNumber" size="small" @change="onContrastRowEdited" />
        </template>
      </el-table-column>
      <el-table-column label="运行" width="72">
        <template #default="scope">
          <el-button size="small" type="primary" @click="runContrast(scope.row)">运行</el-button>
        </template>
      </el-table-column>
      <el-table-column label="备注" min-width="120">
        <template #default="scope">
          <el-input v-model="scope.row.remarks" size="small" @change="onContrastRowEdited" />
        </template>
      </el-table-column>
    </el-table>

    <div class="action-row">
      <el-button @click="onSaveContrastConfigClick">保存配置</el-button>
      <el-button type="danger" @click="deleteSelectedRow">删除选中行</el-button>
    </div>

    <div v-if="contextMenu.visible" class="context-menu" :style="{ top: `${contextMenu.y}px`, left: `${contextMenu.x}px` }">
      <div class="context-item" @click="copySelectedRow">复制选中行</div>
      <div class="context-item" @click="deleteSelectedRow">删除选中行</div>
      <div class="context-item" @click="deleteAllRows">删除所有行</div>
    </div>
  </section>
</template>

<style scoped>
.contrast-form-row {
  flex-wrap: nowrap;
  overflow-x: auto;
  overflow-y: hidden;
  padding-bottom: 4px;
}

.contrast-form-row label,
.contrast-form-row :deep(.path-input),
.contrast-form-row .el-button {
  flex: 0 0 auto;
  white-space: nowrap;
}

.contrast-form-row :deep(.path-input) {
  flex: 1 1 0;
  width: 0;
  min-width: 0;
}

.contrast-form-row .threshold-input {
  flex: 1 1 0;
  width: 0;
  min-width: 0;
}

.path-cell {
  display: flex;
  align-items: center;
  gap: 6px;
}

.path-cell :deep(.el-input) {
  flex: 1;
  min-width: 0;
}
</style>

<script lang="ts">
export default {
  name: "ContrastTab"
};
</script>

<script setup lang="ts">
import { computed, onBeforeUnmount, onMounted, reactive, ref } from "vue";
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

type RunTask = {
  rowKey: string;
  rowSnapshot: ContrastRow;
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

const pendingTasks = ref<RunTask[]>([]);
const activeTask = ref<RunTask | null>(null);
const isDrainingQueue = ref(false);
const queueCount = computed(() => pendingTasks.value.length + (activeTask.value ? 1 : 0));

const rowKeyMap = new WeakMap<ContrastRow, string>();
let rowKeySeed = 0;

function ensureRowKey(row: ContrastRow) {
  const existing = rowKeyMap.get(row);
  if (existing) {
    return existing;
  }
  rowKeySeed += 1;
  const key = `row-${rowKeySeed}`;
  rowKeyMap.set(row, key);
  return key;
}

function cloneRowSnapshot(row: ContrastRow): ContrastRow {
  return {
    standardSamplePath: row.standardSamplePath,
    samplePath: row.samplePath,
    analysisResultsPath: row.analysisResultsPath,
    thresholdNumber: row.thresholdNumber,
    lastResultFilePath: row.lastResultFilePath,
    remarks: row.remarks
  };
}

function findRowByKey(rowKey: string) {
  return contrastRows.value.find((row) => ensureRowKey(row) === rowKey);
}

function isRowRunning(row: ContrastRow) {
  const key = ensureRowKey(row);
  return activeTask.value?.rowKey === key;
}

function isRowQueued(row: ContrastRow) {
  const key = ensureRowKey(row);
  return pendingTasks.value.some((task) => task.rowKey === key);
}

function removePendingTaskByRow(row: ContrastRow) {
  const key = ensureRowKey(row);
  const before = pendingTasks.value.length;
  pendingTasks.value = pendingTasks.value.filter((task) => task.rowKey !== key);
  return before - pendingTasks.value.length;
}

function enqueueRunTask(row: ContrastRow) {
  if (isRowRunning(row) || isRowQueued(row)) {
    return false;
  }

  pendingTasks.value = [
    ...pendingTasks.value,
    {
      rowKey: ensureRowKey(row),
      rowSnapshot: cloneRowSnapshot(row)
    }
  ];
  return true;
}

async function drainQueue() {
  if (isDrainingQueue.value) {
    return;
  }
  isDrainingQueue.value = true;

  try {
    while (pendingTasks.value.length > 0) {
      const [current, ...rest] = pendingTasks.value;
      pendingTasks.value = rest;
      activeTask.value = current;

      try {
        const outputFilePath = await invoke<string>("run_contrast", { row: current.rowSnapshot });
        const row = findRowByKey(current.rowKey);
        if (row) {
          row.lastResultFilePath = outputFilePath;
          await saveContrastConfig({ silent: true });
        }
        ElMessage.success(
          `解析完毕，结果文件：${outputFilePath}（汇总含完全匹配/不完全匹配/完全不同/标样位点缺失的数量与位置）`
        );
      } catch (error) {
        ElMessage.error(String(error));
      } finally {
        activeTask.value = null;
      }
    }
  } finally {
    isDrainingQueue.value = false;
  }
}

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
    row.lastResultFilePath = "";
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
    row.lastResultFilePath = "";
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
    row.lastResultFilePath = "";
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
    lastResultFilePath: "",
    remarks: ""
  });
  await saveContrastConfig({ silent: true });
}

async function deleteSelectedRow() {
  if (!currentContrastRow.value) {
    return;
  }

  const removedTasks = removePendingTaskByRow(currentContrastRow.value);
  if (removedTasks > 0) {
    ElMessage.info("已从队列移除该任务");
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
  if (pendingTasks.value.length > 0) {
    pendingTasks.value = [];
    ElMessage.info("已清空等待队列，当前运行任务将继续执行");
  }

  contrastRows.value = [];
  currentContrastRow.value = null;
  await saveContrastConfig({ silent: true });
  hideContextMenu();
}

async function runContrast(row: ContrastRow) {
  const enqueued = enqueueRunTask(row);
  if (!enqueued) {
    ElMessage.info("该配置已在队列中，请勿重复提交");
    return;
  }

  const tasksAhead = (activeTask.value ? 1 : 0) + pendingTasks.value.length - 1;
  if (tasksAhead > 0) {
    ElMessage.info(`已加入运行队列，前方还有 ${tasksAhead} 个任务`);
  }

  void drainQueue();
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

async function openResultPath(path: string, options?: { suppressError?: boolean }) {
  const target = path.trim();
  if (!target) {
    if (!options?.suppressError) {
      ElMessage.warning("路径为空");
    }
    return false;
  }

  try {
    await revealItemInDir(target);
    return true;
  } catch {}

  try {
    await openPath(target);
    return true;
  } catch {}

  const parentDir = resolveParentDir(target);
  if (parentDir && parentDir !== target) {
    try {
      await openPath(parentDir);
      return true;
    } catch {}
  }

  if (!options?.suppressError) {
    ElMessage.error("无法打开路径");
  }
  return false;
}

async function openAnalysisResultPath(row: ContrastRow) {
  const target = row.analysisResultsPath.trim();
  if (!target) {
    ElMessage.warning("路径为空");
    return;
  }

  const configuredLatest = row.lastResultFilePath.trim();
  if (configuredLatest) {
    const opened = await openResultPath(configuredLatest, { suppressError: true });
    if (opened) {
      return;
    }
  }

  try {
    const latestFile = await invoke<string | null>("find_latest_result_file", { path: target });
    if (latestFile) {
      const opened = await openResultPath(latestFile, { suppressError: true });
      if (opened) {
        if (row.lastResultFilePath !== latestFile) {
          row.lastResultFilePath = latestFile;
          await onContrastRowEdited();
        }
        return;
      }
    }
  } catch {}

  const openedDir = await openResultPath(target, { suppressError: true });
  if (!openedDir) {
    ElMessage.error("无法打开路径");
  }
}

async function onThresholdChanged(row: ContrastRow) {
  row.lastResultFilePath = "";
  await onContrastRowEdited();
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

    <p class="contrast-hint">
      说明：按位点三联值逐位比较并支持阈值；汇总结果会输出完全匹配/不完全匹配/完全不同/标样位点缺失的数量与位置，样本未匹配到的标样位点会计入“标样位点缺失”。
    </p>
    <p v-if="queueCount > 0" class="contrast-queue">
      当前队列：{{ queueCount }}（运行中 {{ activeTask ? 1 : 0 }}，排队 {{ pendingTasks.length }}）
    </p>

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
      <el-table-column label="查看" width="80">
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
      <el-table-column label="查看" width="80">
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
      <el-table-column label="查看" width="80">
        <template #default="scope">
          <el-button size="small" @click="openAnalysisResultPath(scope.row)">查看</el-button>
        </template>
      </el-table-column>

      <el-table-column prop="thresholdNumber" label="阈值" width="90">
        <template #default="scope">
          <el-input v-model="scope.row.thresholdNumber" size="small" @change="onThresholdChanged(scope.row)" />
        </template>
      </el-table-column>
      <el-table-column label="运行" width="84">
        <template #default="scope">
          <el-button
            size="small"
            type="primary"
            :loading="isRowRunning(scope.row)"
            :disabled="isRowRunning(scope.row) || isRowQueued(scope.row)"
            @click="runContrast(scope.row)"
          >
            {{ isRowRunning(scope.row) ? "运行中" : isRowQueued(scope.row) ? "排队中" : "运行" }}
          </el-button>
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

.contrast-hint {
  margin: 4px 0 6px;
  color: #606266;
  font-size: 13px;
  line-height: 1.45;
}

.contrast-queue {
  margin: 0 0 10px;
  color: #409eff;
  font-size: 13px;
  line-height: 1.45;
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

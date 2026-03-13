<script lang="ts">
export default {
  name: "GroupTab"
};
</script>

<script setup lang="ts">
import { nextTick, onMounted, ref } from "vue";
import { invoke } from "@tauri-apps/api/core";
import { open as openDialog } from "@tauri-apps/plugin-dialog";
import { ElMessage, type TableInstance } from "element-plus";
import type { GroupRow } from "../types";
import { loadGroupRows, saveGroupRows } from "../services/configStore";

const groupFilePath = ref("");
const groupRows = ref<GroupRow[]>([]);
const groupTableRef = ref<TableInstance>();
const selectAllChecked = ref(false);

async function chooseGroupFilePath() {
  const file = await openDialog({
    multiple: false,
    directory: false,
    filters: [{ name: "Excel", extensions: ["xlsx", "xls"] }]
  });
  if (!file) {
    return;
  }

  groupFilePath.value = file as string;
  try {
    const rows = await invoke<GroupRow[]>("load_group_data", {
      path: groupFilePath.value
    });
    groupRows.value = rows.map((row) => ({ ...row, group: "", selected: false }));
    selectAllChecked.value = false;
    if (!groupRows.value.length) {
      ElMessage.error("解析Excel异常,请检查Excel数据");
    }
    await nextTick();
    groupTableRef.value?.clearSelection();
  } catch {
    ElMessage.error("解析Excel异常,请检查Excel数据");
  }
}

async function saveGroupConfig() {
  try {
    await saveGroupRows(groupRows.value);
    ElMessage.success("保存配置成功");
  } catch (error) {
    ElMessage.error(`保存配置失败：${error}`);
  }
}

async function loadGroupConfig() {
  try {
    groupRows.value = await loadGroupRows();
    selectAllChecked.value = groupRows.value.length > 0 && groupRows.value.every((row) => row.selected);
    await nextTick();
    groupTableRef.value?.clearSelection();
    groupRows.value.forEach((row) => {
      if (row.selected) {
        groupTableRef.value?.toggleRowSelection(row, true);
      }
    });
  } catch (error) {
    ElMessage.error(`加载配置失败：${error}`);
  }
}

function onGroupSelectionChange(selection: GroupRow[]) {
  const selectedSet = new Set(selection);
  groupRows.value.forEach((row) => {
    row.selected = selectedSet.has(row);
  });
  selectAllChecked.value = groupRows.value.length > 0 && groupRows.value.every((row) => row.selected);
}

function toggleSelectAll() {
  if (selectAllChecked.value) {
    groupRows.value.forEach((row) => groupTableRef.value?.toggleRowSelection(row, true));
  } else {
    groupTableRef.value?.clearSelection();
  }
}

async function doGroup() {
  try {
    const rows = await invoke<GroupRow[]>("do_group", { rows: groupRows.value });
    groupRows.value = rows;
    selectAllChecked.value = rows.length > 0;
    await nextTick();
    groupTableRef.value?.clearSelection();
    rows.forEach((row) => {
      if (row.selected) {
        groupTableRef.value?.toggleRowSelection(row, true);
      }
    });
  } catch (error) {
    const message = String(error);
    if (message.includes("请勾选数据")) {
      ElMessage.warning("请勾选数据");
      return;
    }
    ElMessage.error(message);
  }
}

onMounted(async () => {
  await loadGroupConfig();
});
</script>

<template>
  <el-tab-pane label="分组" name="group">
    <div class="form-row">
      <el-checkbox v-model="selectAllChecked" @change="toggleSelectAll">全选</el-checkbox>
      <label>原始文件路径：</label>
      <el-input v-model="groupFilePath" readonly class="path-input">
        <template #append>
          <el-button @click="chooseGroupFilePath">选择</el-button>
        </template>
      </el-input>
      <el-button @click="saveGroupConfig">保存配置</el-button>
      <el-button @click="loadGroupConfig">加载配置</el-button>
      <el-button type="primary" @click="doGroup">分组勾选数据</el-button>
    </div>

    <el-table ref="groupTableRef" :data="groupRows" @selection-change="onGroupSelectionChange">
      <el-table-column type="selection" width="55" />
      <el-table-column prop="group" label="分组名称" min-width="160" />
      <el-table-column prop="primerNo" label="引物编号" min-width="160" />
      <el-table-column prop="fuel" label="染料" min-width="120" />
      <el-table-column prop="start" label="起始范围" width="120" />
      <el-table-column prop="end" label="终止范围" width="120" />
    </el-table>
  </el-tab-pane>
</template>

import { LazyStore } from "@tauri-apps/plugin-store";
import type { ContrastRow, GroupRow } from "../types";

const STORE_FILE = "datalinker.store.json";
const CONTRAST_STORE_KEY = "contrastRows";
const GROUP_STORE_KEY = "groupRows";

let storeInstance: LazyStore | null = null;
let initPromise: Promise<LazyStore> | null = null;

function normalizeContrastRow(row: Partial<ContrastRow>): ContrastRow {
  return {
    standardSamplePath: row.standardSamplePath ?? "",
    samplePath: row.samplePath ?? "",
    analysisResultsPath: row.analysisResultsPath ?? "",
    thresholdNumber: row.thresholdNumber ?? "1",
    remarks: row.remarks ?? ""
  };
}

function normalizeGroupRow(row: Partial<GroupRow>): GroupRow {
  return {
    selected: row.selected ?? false,
    group: row.group ?? "",
    primerNo: row.primerNo ?? "",
    fuel: row.fuel ?? "",
    start: row.start ?? "",
    end: row.end ?? ""
  };
}

async function getStore() {
  if (storeInstance) {
    return storeInstance;
  }

  if (!initPromise) {
    const store = new LazyStore(STORE_FILE, {
      autoSave: 200,
      defaults: {
        [CONTRAST_STORE_KEY]: [],
        [GROUP_STORE_KEY]: []
      }
    });

    initPromise = store.init().then(() => {
      storeInstance = store;
      return store;
    });
  }

  return initPromise;
}

export async function loadContrastRows() {
  const store = await getStore();
  await store.reload({ ignoreDefaults: true });
  const rows = await store.get<Partial<ContrastRow>[]>(CONTRAST_STORE_KEY);
  return (rows ?? []).map((row) => normalizeContrastRow(row));
}

export async function saveContrastRows(rows: ContrastRow[]) {
  const store = await getStore();
  await store.set(CONTRAST_STORE_KEY, rows);
  await store.save();
}

export async function loadGroupRows() {
  const store = await getStore();
  await store.reload({ ignoreDefaults: true });
  const rows = await store.get<Partial<GroupRow>[]>(GROUP_STORE_KEY);
  return (rows ?? []).map((row) => normalizeGroupRow(row));
}

export async function saveGroupRows(rows: GroupRow[]) {
  const store = await getStore();
  await store.set(GROUP_STORE_KEY, rows);
  await store.save();
}


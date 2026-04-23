import { LazyStore } from "@tauri-apps/plugin-store";
import type { ContrastRow, ContrastSortMode } from "../types";

const STORE_FILE = "datalinker.store.json";
const CONTRAST_STORE_KEY = "contrastRows";
const CONTRAST_SORT_MODE_STORE_KEY = "contrastSortMode";
const CONTRAST_NEXT_ID_STORE_KEY = "contrastNextId";
const LEGACY_GROUP_STORE_KEY = "groupRows";
const DEFAULT_CONTRAST_SORT_MODE: ContrastSortMode = "createdAtDesc";
const MIN_CONTRAST_NEXT_ID = 1;

let storeInstance: LazyStore | null = null;
let initPromise: Promise<LazyStore> | null = null;

function normalizeSortMode(mode: unknown): ContrastSortMode {
  return mode === "updatedAtDesc" ? "updatedAtDesc" : DEFAULT_CONTRAST_SORT_MODE;
}

function normalizeNextId(value: unknown) {
  if (typeof value === "number" && Number.isFinite(value) && value >= MIN_CONTRAST_NEXT_ID) {
    return Math.floor(value);
  }
  return MIN_CONTRAST_NEXT_ID;
}

function normalizeRowId(value: unknown) {
  if (typeof value === "number" && Number.isFinite(value) && value > 0) {
    return Math.floor(value);
  }
  return 0;
}

function normalizeTimestamp(value: unknown) {
  if (typeof value === "number" && Number.isFinite(value) && value > 0) {
    return Math.floor(value);
  }
  return 0;
}

function normalizeThresholdNumber(value: unknown) {
  const text = String(value ?? "").trim();
  return text || "1";
}

function normalizeContrastRow(row: Partial<ContrastRow>): ContrastRow {
  return {
    id: normalizeRowId(row.id),
    standardSamplePath: row.standardSamplePath ?? "",
    samplePath: row.samplePath ?? "",
    analysisResultsPath: row.analysisResultsPath ?? "",
    thresholdNumber: normalizeThresholdNumber(row.thresholdNumber),
    lastResultFilePath: row.lastResultFilePath ?? "",
    remarks: row.remarks ?? "",
    createdAt: normalizeTimestamp(row.createdAt),
    updatedAt: normalizeTimestamp(row.updatedAt)
  };
}

function getMaxRowId(rows: ContrastRow[]) {
  return rows.reduce((max, row) => (row.id > max ? row.id : max), 0);
}

function migrateRows(rows: ContrastRow[], storedNextId: number) {
  const migratedRows = rows.map((row) => ({ ...row }));
  let didMutateRows = false;

  if (migratedRows.some((row) => row.createdAt <= 0 || row.updatedAt <= 0)) {
    const base = Date.now() + migratedRows.length;
    migratedRows.forEach((row, index) => {
      if (row.createdAt <= 0) {
        row.createdAt = base - index;
        didMutateRows = true;
      }
      if (row.updatedAt <= 0) {
        row.updatedAt = row.createdAt;
        didMutateRows = true;
      }
      if (row.updatedAt < row.createdAt) {
        row.updatedAt = row.createdAt;
        didMutateRows = true;
      }
    });
  }

  let idSeed = Math.max(storedNextId - 1, getMaxRowId(migratedRows));
  migratedRows.forEach((row) => {
    if (row.id <= 0) {
      idSeed += 1;
      row.id = idSeed;
      didMutateRows = true;
    }
  });

  const nextId = Math.max(idSeed + 1, MIN_CONTRAST_NEXT_ID);
  const didMutateNextId = nextId !== storedNextId;

  return {
    rows: migratedRows,
    nextId,
    didMutateRows,
    didMutateNextId
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
        [CONTRAST_SORT_MODE_STORE_KEY]: DEFAULT_CONTRAST_SORT_MODE,
        [CONTRAST_NEXT_ID_STORE_KEY]: MIN_CONTRAST_NEXT_ID
      }
    });

    initPromise = store.init().then(async () => {
      await store.reload({ ignoreDefaults: true });
      const removed = await store.delete(LEGACY_GROUP_STORE_KEY);
      if (removed) {
        await store.save();
      }
      storeInstance = store;
      return store;
    });
  }

  return initPromise;
}

export async function loadContrastRows() {
  const store = await getStore();
  await store.reload({ ignoreDefaults: true });

  const storedRows = await store.get<Partial<ContrastRow>[]>(CONTRAST_STORE_KEY);
  const normalizedRows = (storedRows ?? []).map((row) => normalizeContrastRow(row));

  const storedNextId = normalizeNextId(await store.get<number>(CONTRAST_NEXT_ID_STORE_KEY));
  const migration = migrateRows(normalizedRows, storedNextId);

  if (migration.didMutateRows || migration.didMutateNextId) {
    await store.set(CONTRAST_STORE_KEY, migration.rows);
    await store.set(CONTRAST_NEXT_ID_STORE_KEY, migration.nextId);
    await store.save();
  }

  return migration.rows;
}

export async function saveContrastRows(rows: ContrastRow[]) {
  const store = await getStore();
  await store.set(CONTRAST_STORE_KEY, rows);

  const storedNextId = normalizeNextId(await store.get<number>(CONTRAST_NEXT_ID_STORE_KEY));
  const nextId = Math.max(storedNextId, getMaxRowId(rows) + 1, MIN_CONTRAST_NEXT_ID);
  if (nextId !== storedNextId) {
    await store.set(CONTRAST_NEXT_ID_STORE_KEY, nextId);
  }

  await store.save();
}

export async function loadContrastSortMode() {
  const store = await getStore();
  await store.reload({ ignoreDefaults: true });
  const mode = await store.get<ContrastSortMode>(CONTRAST_SORT_MODE_STORE_KEY);
  return normalizeSortMode(mode);
}

export async function saveContrastSortMode(mode: ContrastSortMode) {
  const store = await getStore();
  await store.set(CONTRAST_SORT_MODE_STORE_KEY, normalizeSortMode(mode));
  await store.save();
}

export async function takeNextContrastRowId() {
  const store = await getStore();
  const nextId = normalizeNextId(await store.get<number>(CONTRAST_NEXT_ID_STORE_KEY));
  await store.set(CONTRAST_NEXT_ID_STORE_KEY, nextId + 1);
  await store.save();
  return nextId;
}

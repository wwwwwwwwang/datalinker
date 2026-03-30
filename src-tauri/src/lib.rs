// Prevents additional console window on Windows in release, DO NOT REMOVE!!
#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

use calamine::{open_workbook_auto, Data, Reader};
use serde::{Deserialize, Serialize};
use std::collections::{BTreeSet, HashMap};
use std::fs;
use std::path::{Path, PathBuf};
use std::time::{SystemTime, UNIX_EPOCH};
use tauri::{Manager, PhysicalSize, Size};
use umya_spreadsheet::{
    self as umya, Border, HorizontalAlignmentValues, VerticalAlignmentValues, Worksheet,
};

#[derive(Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct ContrastRow {
    standard_sample_path: String,
    sample_path: String,
    analysis_results_path: String,
    threshold_number: String,
    remarks: String,
}

#[derive(Clone)]
struct DnaData {
    batch_code: String,
    label: String,
    a: f64,
    b: f64,
    c: f64,
}

#[derive(Clone)]
struct LabelResult {
    label: String,
    result: i32,
    standard_a: f64,
    standard_b: f64,
    standard_c: f64,
    sample_a: f64,
    sample_b: f64,
    sample_c: f64,
}

#[derive(Clone)]
struct ContrastResultRow {
    simple_data_batch_code: Option<String>,
    standard_sample_data_batch_code: String,
    count: usize,
    same_number_bits: usize,
    different_bits: usize,
    missing_bits: usize,
}

#[derive(Clone)]
struct ContrastDetailRow {
    contrast_type: String,
    simple_data_batch_code: String,
    standard_sample_data_batch_code: String,
    label: String,
    category: String,
    standard_a: f64,
    standard_b: f64,
    standard_c: f64,
    sample_a: f64,
    sample_b: f64,
    sample_c: f64,
}

struct ContrastProcessResult {
    summary_rows: Vec<ContrastResultRow>,
    detail_rows: Vec<ContrastDetailRow>,
}

type CompareKey = (String, String);
type CompareLocusKey = (String, String, String);

#[derive(Clone, Copy, Debug, Default)]
struct LocusStatus {
    same: bool,
    partial: bool,
    diff: bool,
    missing: bool,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum LocusCategory {
    Same,
    Partial,
    Diff,
    Missing,
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
struct ContrastExportRow {
    sample_batch_code: String,
    standard_batch_code: String,
    count: usize,
    same_count: usize,
    partial_count: usize,
    diff_count: usize,
    missing_count: usize,
    same_positions: String,
    partial_positions: String,
    diff_positions: String,
    missing_positions: String,
}

type BatchDataMap = HashMap<String, Vec<DnaData>>;
type BatchLabelMap = HashMap<String, HashMap<String, Vec<DnaData>>>;

const CONFIG_DIR_NAME: &str = "xJavaFxTool";
const CONTRAST_CONFIG_FILE: &str = "dataProcess.properties";
const LEGACY_GROUP_CONFIG_FILE: &str = "groupProcess.properties";

fn config_dir() -> PathBuf {
    if let Some(home) = dirs::home_dir() {
        return home.join(CONFIG_DIR_NAME);
    }
    PathBuf::from(CONFIG_DIR_NAME)
}

fn config_file(name: &str) -> PathBuf {
    config_dir().join(name)
}

fn cleanup_legacy_group_config() {
    let path = config_file(LEGACY_GROUP_CONFIG_FILE);
    if let Err(error) = fs::remove_file(path) {
        if error.kind() != std::io::ErrorKind::NotFound {
            let _ = error;
        }
    }
}

fn cleanup_legacy_group_store<R: tauri::Runtime>(app: &tauri::App<R>) {
    let mut store_paths = Vec::new();
    if let Ok(app_data_dir) = app.path().app_data_dir() {
        store_paths.push(app_data_dir.join("datalinker.store.json"));
    }
    if let Some(app_data_dir) = std::env::var_os("APPDATA") {
        store_paths.push(
            PathBuf::from(app_data_dir)
                .join("com.admin.datalinker")
                .join("datalinker.store.json"),
        );
    }
    store_paths.sort();
    store_paths.dedup();

    for store_path in store_paths {
        let Ok(content) = fs::read_to_string(&store_path) else {
            continue;
        };
        let content = content.trim_start_matches('\u{feff}');
        let Ok(mut store_json) = serde_json::from_str::<serde_json::Value>(content) else {
            continue;
        };
        let Some(object) = store_json.as_object_mut() else {
            continue;
        };

        if object.remove("groupRows").is_some() {
            if let Ok(serialized) = serde_json::to_string_pretty(&store_json) {
                let _ = fs::write(store_path, serialized);
            }
        }
    }
}

fn ensure_parent_dir(path: &Path) -> Result<(), String> {
    if let Some(parent) = path.parent() {
        fs::create_dir_all(parent).map_err(|e| format!("闂傚倸鍊风粈渚€骞夐敍鍕殰婵°倕鍟伴惌娆撴煙鐎电啸缁惧彞绮欓弻鐔煎礈瑜忕敮娑㈡煟閹绢垰浜鹃梺璇查缁犲秹宕曢崡鐏绘椽濡搁埡浣稿殤闂佺鐬奸崑鐐烘偂閻斿吋鐓冮柛婵嗗瀹搞儵鏌＄€ｎ偆銆掔紒杈ㄦ尭椤撳ジ宕ㄩ鍜冪悼閳ь剙鐏氬妯尖偓姘煎幖椤洩绠涘☉杈ㄦ櫇闂? {}", e))?;
    }
    Ok(())
}

fn escape_value(value: &str) -> String {
    let mut out = String::new();
    for ch in value.chars() {
        match ch {
            '\\' => out.push_str("\\\\"),
            '\n' => out.push_str("\\n"),
            '\r' => out.push_str("\\r"),
            '\t' => out.push_str("\\t"),
            '=' => out.push_str("\\="),
            ':' => out.push_str("\\:"),
            _ => out.push(ch),
        }
    }
    out
}

fn unescape_value(value: &str) -> String {
    let mut out = String::new();
    let mut chars = value.chars().peekable();
    while let Some(ch) = chars.next() {
        if ch != '\\' {
            out.push(ch);
            continue;
        }
        let Some(next) = chars.next() else { break };
        match next {
            'n' => out.push('\n'),
            'r' => out.push('\r'),
            't' => out.push('\t'),
            '\\' => out.push('\\'),
            ':' => out.push(':'),
            '=' => out.push('='),
            'u' => {
                let mut hex = String::new();
                for _ in 0..4 {
                    if let Some(h) = chars.next() {
                        hex.push(h);
                    }
                }
                if let Ok(code) = u16::from_str_radix(&hex, 16) {
                    if let Some(c) = char::from_u32(code as u32) {
                        out.push(c);
                    }
                }
            }
            _ => out.push(next),
        }
    }
    out
}

fn parse_properties(content: &str) -> Vec<(String, String)> {
    let mut entries = Vec::new();
    for raw_line in content.lines() {
        let line = raw_line.trim();
        if line.is_empty() || line.starts_with('#') || line.starts_with('!') {
            continue;
        }
        let mut split_index: Option<usize> = None;
        let mut escaped = false;
        for (idx, ch) in line.char_indices() {
            if escaped {
                escaped = false;
                continue;
            }
            if ch == '\\' {
                escaped = true;
                continue;
            }
            if ch == '=' || ch == ':' {
                split_index = Some(idx);
                break;
            }
        }
        let (key, value) = if let Some(idx) = split_index {
            let (k, v) = line.split_at(idx);
            (k.trim().to_string(), v[1..].trim().to_string())
        } else {
            (line.to_string(), String::new())
        };
        entries.push((key, unescape_value(&value)));
    }
    entries
}

fn load_properties(path: &Path) -> Result<Vec<(String, String)>, String> {
    let content = fs::read_to_string(path).map_err(|e| format!("闂傚倷娴囧畷鍨叏閺夋嚚娲煛閸滀焦鏅悷婊勫灴婵＄敻骞囬弶璺ㄥ€炲銈嗗笂閼冲爼宕㈤幇鐗堝€垫鐐茬仢閸旀碍绻涚仦鍌氬鐎殿喗鎮傚畷鎺戭潩閼测晛鏁搁梻浣筋嚃閸ㄨ鲸绔熼崱娆掑С闁秆勵殕閸? {}", e))?;
    Ok(parse_properties(&content))
}

fn save_properties(path: &Path, entries: &[(String, String)]) -> Result<(), String> {
    ensure_parent_dir(path)?;
    let mut lines = Vec::new();
    for (key, value) in entries {
        lines.push(format!("{}={}", key, escape_value(value)));
    }
    fs::write(path, lines.join("\n")).map_err(|e| format!("濠电姷鏁搁崕鎴犲緤閽樺娲晜閻愵剙搴婇梺绋跨灱閸嬬偤宕戦妶澶嬬厪濠电偟鍋撳▍鍡涙煟閹绢垰浜鹃梺璇查缁犲秹宕曢崡鐏绘椽濡搁埡浣稿殤闂佺鏈崙褰掓偄閸℃稒鐓欐い鏍ㄧ☉椤ュ绱掗妸銉﹀仴闁? {}", e))?;
    Ok(())
}

fn table_bean_index(key: &str) -> usize {
    key.strip_prefix("tableBean")
        .and_then(|suffix| suffix.parse::<usize>().ok())
        .unwrap_or(usize::MAX)
}

fn sort_table_entries(entries: &mut [(String, String)]) {
    entries.sort_by(|a, b| {
        let index_a = table_bean_index(&a.0);
        let index_b = table_bean_index(&b.0);
        index_a.cmp(&index_b).then_with(|| a.0.cmp(&b.0))
    });
}

fn contrast_row_to_property(row: &ContrastRow) -> String {
    format!(
        "{}__{}__{}__{}__{}",
        row.standard_sample_path,
        row.sample_path,
        row.analysis_results_path,
        row.threshold_number,
        row.remarks
    )
}

fn contrast_row_from_property(value: &str) -> ContrastRow {
    let mut parts = value.splitn(5, "__");
    ContrastRow {
        standard_sample_path: parts.next().unwrap_or("").to_string(),
        sample_path: parts.next().unwrap_or("").to_string(),
        analysis_results_path: parts.next().unwrap_or("").to_string(),
        threshold_number: parts.next().unwrap_or("").to_string(),
        remarks: parts.next().unwrap_or("").to_string(),
    }
}

fn cell_to_string(cell: &Data) -> String {
    match cell {
        Data::Empty => String::new(),
        Data::String(s) => s.to_string(),
        _ => cell.to_string(),
    }
}

fn is_blank(value: &str) -> bool {
    value.trim().is_empty()
}

fn convert_data(value: Option<&String>) -> f64 {
    let Some(raw) = value else { return 0.0 };
    if is_blank(raw) {
        return 0.0;
    }
    raw.parse::<f64>().unwrap_or(0.0)
}

fn get_excel_data(path: &Path) -> Result<Vec<DnaData>, String> {
    let mut workbook =
        open_workbook_auto(path).map_err(|e| format!("Parse Excel failed: {}", e))?;
    let range = workbook
        .worksheet_range_at(0)
        .ok_or_else(|| "Parse Excel failed, please check excel data".to_string())?
        .map_err(|e| format!("Parse Excel failed: {}", e))?;

    let mut rows = range.rows();
    let header = rows.next().unwrap_or(&[]);
    let mut tab_map: HashMap<usize, String> = HashMap::new();
    for (idx, cell) in header.iter().enumerate() {
        if (idx + 1) % 3 == 0 {
            tab_map.insert(idx, cell_to_string(cell));
        }
    }

    let mut dna_data_list = Vec::new();
    for row in rows {
        let mut row_map: HashMap<usize, String> = HashMap::new();
        for (idx, cell) in row.iter().enumerate() {
            let value = cell_to_string(cell);
            if is_blank(&value) {
                continue;
            }
            row_map.insert(idx, value);
        }

        let batch_code = row.get(1).map(cell_to_string).unwrap_or_default();

        let mut keys: Vec<usize> = row_map.keys().cloned().collect();
        keys.sort_unstable();
        for idx in keys {
            if (idx + 1) % 3 != 0 {
                continue;
            }
            let a = convert_data(row_map.get(&idx));
            let b = convert_data(row_map.get(&(idx + 1)));
            let c = convert_data(row_map.get(&(idx + 2)));
            let label = tab_map.get(&idx).cloned().unwrap_or_default();
            dna_data_list.push(DnaData {
                batch_code: batch_code.clone(),
                label,
                a,
                b,
                c,
            });
        }
    }

    Ok(dna_data_list)
}

fn build_batch_data_map(data: &[DnaData]) -> BatchDataMap {
    let mut batch_map: BatchDataMap = HashMap::new();
    for item in data {
        batch_map
            .entry(item.batch_code.clone())
            .or_default()
            .push(item.clone());
    }
    batch_map
}

fn build_batch_label_map(data: &[DnaData]) -> BatchLabelMap {
    let mut batch_label_map: BatchLabelMap = HashMap::new();
    for item in data {
        batch_label_map
            .entry(item.batch_code.clone())
            .or_default()
            .entry(item.label.clone())
            .or_default()
            .push(item.clone());
    }
    batch_label_map
}

fn get_compare_a(standard: &DnaData, sample: &DnaData, threshold: i32) -> i32 {
    let a = standard.a;
    let b = standard.b;
    let c = standard.c;
    let a1 = sample.a;
    let b1 = sample.b;
    let c1 = sample.c;
    if a1 + b1 + c1 == 0.0 {
        return 0;
    }
    if a1 >= a - threshold as f64 && a1 <= a + threshold as f64 {
        return 3;
    }
    if a1 >= b - threshold as f64 && a1 <= b + threshold as f64 {
        return 2;
    }
    if a1 >= c - threshold as f64 && a1 <= c + threshold as f64 {
        return 1;
    }
    0
}

fn get_compare_b(standard: &DnaData, sample: &DnaData, threshold: i32) -> i32 {
    let a = standard.a;
    let b = standard.b;
    let c = standard.c;
    let a1 = sample.a;
    let b1 = sample.b;
    let c1 = sample.c;
    if a1 + b1 + c1 == 0.0 {
        return 0;
    }
    if b1 >= a - threshold as f64 && b1 <= a + threshold as f64 {
        return 3;
    }
    if b1 >= b - threshold as f64 && b1 <= b + threshold as f64 {
        return 2;
    }
    if b1 >= c - threshold as f64 && b1 <= c + threshold as f64 {
        return 1;
    }
    0
}

fn get_compare_c(standard: &DnaData, sample: &DnaData, threshold: i32) -> i32 {
    let a = standard.a;
    let b = standard.b;
    let c = standard.c;
    let a1 = sample.a;
    let b1 = sample.b;
    let c1 = sample.c;
    if a1 + b1 + c1 == 0.0 {
        return 0;
    }
    if c1 >= a - threshold as f64 && c1 <= a + threshold as f64 {
        return 3;
    }
    if c1 >= b - threshold as f64 && c1 <= b + threshold as f64 {
        return 2;
    }
    if c1 >= c - threshold as f64 && c1 <= c + threshold as f64 {
        return 1;
    }
    0
}

fn get_label_result(a: i32, b: i32, c: i32) -> i32 {
    let count = a + b + c;
    if count == 0 {
        return 0;
    }
    if count > 1 {
        return 1;
    }
    2
}

fn real_compare(standard: &DnaData, sample: &DnaData) -> i32 {
    if standard.a == 0.0 && standard.b == 0.0 && standard.c == 0.0 {
        return 0;
    }
    let mut standard_arr = vec![standard.a, standard.b, standard.c];
    standard_arr.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
    let mut sample_arr = vec![sample.a, sample.b, sample.c];
    sample_arr.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
    if standard_arr == sample_arr {
        1
    } else {
        2
    }
}

fn result_category(result: i32) -> &'static str {
    match result {
        1 => "\u{76F8}\u{540C}",
        2 => "\u{5DEE}\u{5F02}",
        _ => "\u{7F3A}\u{5931}",
    }
}

fn normal_process(
    contrast_row: &ContrastRow,
    standard_batch_map: &BatchDataMap,
    sample_batch_map: &BatchLabelMap,
    is_paternity: bool,
) -> ContrastProcessResult {
    let threshold_number = contrast_row.threshold_number.parse::<i32>().unwrap_or(0);
    let contrast_type = if is_paternity {
        "\u{4EB2}\u{5B50}\u{9274}\u{5B9A}".to_string()
    } else {
        "\u{771F}\u{5B9E}\u{6027}\u{9274}\u{5B9A}".to_string()
    };

    let mut sample_batch_codes: Vec<String> = sample_batch_map.keys().cloned().collect();
    sample_batch_codes.sort();

    let mut standard_batch_codes: Vec<String> = standard_batch_map.keys().cloned().collect();
    standard_batch_codes.sort();

    let mut contrast_results: Vec<ContrastResultRow> = Vec::new();
    let mut detail_rows: Vec<ContrastDetailRow> = Vec::new();

    for sample_batch_code in sample_batch_codes {
        let sample_label_map = sample_batch_map.get(&sample_batch_code);
        for standard_batch_code in &standard_batch_codes {
            let Some(standard_items) = standard_batch_map.get(standard_batch_code) else {
                continue;
            };
            if standard_items.is_empty() {
                continue;
            }

            let mut items: Vec<LabelResult> = Vec::with_capacity(standard_items.len());
            for standard in standard_items {
                if let Some(sample_rows) =
                    sample_label_map.and_then(|labels| labels.get(&standard.label))
                {
                    for sample_data in sample_rows {
                        let result = if is_paternity {
                            let compare_a = get_compare_a(standard, sample_data, threshold_number);
                            let compare_b = get_compare_b(standard, sample_data, threshold_number);
                            let compare_c = get_compare_c(standard, sample_data, threshold_number);
                            get_label_result(compare_a, compare_b, compare_c)
                        } else {
                            real_compare(standard, sample_data)
                        };
                        items.push(LabelResult {
                            label: standard.label.clone(),
                            result,
                            standard_a: standard.a,
                            standard_b: standard.b,
                            standard_c: standard.c,
                            sample_a: sample_data.a,
                            sample_b: sample_data.b,
                            sample_c: sample_data.c,
                        });
                    }
                    continue;
                }

                items.push(LabelResult {
                    label: standard.label.clone(),
                    result: 0,
                    standard_a: standard.a,
                    standard_b: standard.b,
                    standard_c: standard.c,
                    sample_a: 0.0,
                    sample_b: 0.0,
                    sample_c: 0.0,
                });
            }

            let same_count = items.iter().filter(|r| r.result == 1).count();
            let diff_count = items.iter().filter(|r| r.result == 2).count();
            let missing_count = items.iter().filter(|r| r.result == 0).count();
            contrast_results.push(ContrastResultRow {
                simple_data_batch_code: Some(sample_batch_code.clone()),
                standard_sample_data_batch_code: standard_batch_code.clone(),
                count: items.len(),
                same_number_bits: same_count,
                different_bits: diff_count,
                missing_bits: missing_count,
            });

            for item in items {
                detail_rows.push(ContrastDetailRow {
                    contrast_type: contrast_type.clone(),
                    simple_data_batch_code: sample_batch_code.clone(),
                    standard_sample_data_batch_code: standard_batch_code.clone(),
                    label: item.label.clone(),
                    category: result_category(item.result).to_string(),
                    standard_a: item.standard_a,
                    standard_b: item.standard_b,
                    standard_c: item.standard_c,
                    sample_a: item.sample_a,
                    sample_b: item.sample_b,
                    sample_c: item.sample_c,
                });
            }
        }
    }

    contrast_results.sort_by(|a, b| {
        let a_simple = a.simple_data_batch_code.clone().unwrap_or_default();
        let b_simple = b.simple_data_batch_code.clone().unwrap_or_default();
        let first = a_simple.cmp(&b_simple);
        if first == std::cmp::Ordering::Equal {
            a.standard_sample_data_batch_code
                .cmp(&b.standard_sample_data_batch_code)
        } else {
            first
        }
    });

    detail_rows.sort_by(|a, b| {
        a.contrast_type
            .cmp(&b.contrast_type)
            .then_with(|| a.simple_data_batch_code.cmp(&b.simple_data_batch_code))
            .then_with(|| {
                a.standard_sample_data_batch_code
                    .cmp(&b.standard_sample_data_batch_code)
            })
            .then_with(|| a.label.cmp(&b.label))
    });

    ContrastProcessResult {
        summary_rows: contrast_results,
        detail_rows,
    }
}

fn summary_row_key(row: &ContrastResultRow) -> CompareKey {
    (
        row.simple_data_batch_code.clone().unwrap_or_default(),
        row.standard_sample_data_batch_code.clone(),
    )
}

fn detail_locus_key(row: &ContrastDetailRow) -> Option<CompareLocusKey> {
    let label = row.label.trim();
    if label.is_empty() {
        return None;
    }
    Some((
        row.simple_data_batch_code.clone(),
        row.standard_sample_data_batch_code.clone(),
        label.to_string(),
    ))
}

fn is_sample_missing(row: &ContrastDetailRow) -> bool {
    row.sample_a == 0.0 && row.sample_b == 0.0 && row.sample_c == 0.0
}

fn update_normal_locus_status(status: &mut LocusStatus, row: &ContrastDetailRow) {
    if is_sample_missing(row) {
        status.missing = true;
    }
    if row.category == "\u{5DEE}\u{5F02}" {
        status.partial = true;
    }
}

fn update_real_locus_status(status: &mut LocusStatus, row: &ContrastDetailRow) {
    if is_sample_missing(row) {
        status.missing = true;
    }
    if row.category == "\u{5DEE}\u{5F02}" {
        status.diff = true;
    } else if row.category == "\u{76F8}\u{540C}" {
        status.same = true;
    }
}

fn resolve_locus_category(status: &LocusStatus) -> LocusCategory {
    if status.missing {
        LocusCategory::Missing
    } else if status.diff {
        LocusCategory::Diff
    } else if status.partial {
        LocusCategory::Partial
    } else {
        LocusCategory::Same
    }
}

fn join_positions(labels: &BTreeSet<String>) -> String {
    labels.iter().cloned().collect::<Vec<_>>().join("/")
}

fn build_export_rows(
    normal_result: &ContrastProcessResult,
    real_result: &ContrastProcessResult,
) -> Vec<ContrastExportRow> {
    let mut locus_status_map: HashMap<CompareLocusKey, LocusStatus> = HashMap::new();
    for row in &normal_result.detail_rows {
        let Some(key) = detail_locus_key(row) else {
            continue;
        };
        let status = locus_status_map.entry(key).or_default();
        update_normal_locus_status(status, row);
    }
    for row in &real_result.detail_rows {
        let Some(key) = detail_locus_key(row) else {
            continue;
        };
        let status = locus_status_map.entry(key).or_default();
        update_real_locus_status(status, row);
    }

    let mut keys = BTreeSet::new();
    for row in &normal_result.summary_rows {
        keys.insert(summary_row_key(row));
    }
    for row in &real_result.summary_rows {
        keys.insert(summary_row_key(row));
    }
    for locus_key in locus_status_map.keys() {
        keys.insert((locus_key.0.clone(), locus_key.1.clone()));
    }

    let mut same_positions: HashMap<CompareKey, BTreeSet<String>> = HashMap::new();
    let mut partial_positions: HashMap<CompareKey, BTreeSet<String>> = HashMap::new();
    let mut diff_positions: HashMap<CompareKey, BTreeSet<String>> = HashMap::new();
    let mut missing_positions: HashMap<CompareKey, BTreeSet<String>> = HashMap::new();
    let mut all_positions: HashMap<CompareKey, BTreeSet<String>> = HashMap::new();

    for (locus_key, status) in &locus_status_map {
        let compare_key = (locus_key.0.clone(), locus_key.1.clone());
        let label = locus_key.2.clone();
        all_positions
            .entry(compare_key.clone())
            .or_default()
            .insert(label.clone());

        match resolve_locus_category(status) {
            LocusCategory::Same => {
                same_positions.entry(compare_key).or_default().insert(label);
            }
            LocusCategory::Partial => {
                partial_positions
                    .entry(compare_key)
                    .or_default()
                    .insert(label);
            }
            LocusCategory::Diff => {
                diff_positions.entry(compare_key).or_default().insert(label);
            }
            LocusCategory::Missing => {
                missing_positions
                    .entry(compare_key)
                    .or_default()
                    .insert(label);
            }
        }
    }

    let mut rows = Vec::with_capacity(keys.len());
    for key in keys {
        let (sample_batch_code, standard_batch_code) = key.clone();
        let count = all_positions.get(&key).map_or(0, BTreeSet::len);
        let same_count = same_positions.get(&key).map_or(0, BTreeSet::len);
        let partial_count = partial_positions.get(&key).map_or(0, BTreeSet::len);
        let diff_count = diff_positions.get(&key).map_or(0, BTreeSet::len);
        let missing_count = missing_positions.get(&key).map_or(0, BTreeSet::len);

        rows.push(ContrastExportRow {
            sample_batch_code,
            standard_batch_code,
            count,
            same_count,
            partial_count,
            diff_count,
            missing_count,
            same_positions: same_positions
                .get(&key)
                .map(join_positions)
                .unwrap_or_default(),
            partial_positions: partial_positions
                .get(&key)
                .map(join_positions)
                .unwrap_or_default(),
            diff_positions: diff_positions
                .get(&key)
                .map(join_positions)
                .unwrap_or_default(),
            missing_positions: missing_positions
                .get(&key)
                .map(join_positions)
                .unwrap_or_default(),
        });
    }
    rows
}

fn write_export_sheet(sheet: &mut Worksheet, data: &[ContrastExportRow]) {
    let headers = [
        "\u{6837}\u{54C1}\u{7F16}\u{53F7}",
        "\u{6807}\u{4F4D}\u{7F16}\u{53F7}",
        "\u{6D4B}\u{5B9A}\u{4F4D}\u{70B9}\u{6570}",
        "\u{76F8}\u{540C}\u{4F4D}\u{70B9}\u{6570}",
        "\u{4E0D}\u{5B8C}\u{5168}\u{76F8}\u{540C}\u{4F4D}\u{70B9}",
        "\u{5DEE}\u{5F02}\u{4F4D}\u{70B9}\u{6570}",
        "\u{7F3A}\u{5931}\u{4F4D}\u{70B9}\u{6570}",
        "\u{76F8}\u{540C}\u{4F4D}\u{70B9}\u{4F4D}\u{7F6E}",
        "\u{4E0D}\u{5B8C}\u{5168}\u{76F8}\u{540C}\u{4F4D}\u{7F6E}",
        "\u{5DEE}\u{5F02}\u{4F4D}\u{7F6E}",
        "\u{7F3A}\u{5931}\u{4F4D}\u{7F6E}",
    ];
    let sub_headers = [
        "",
        "",
        "",
        "\u{5B8C}\u{5168}\u{5339}\u{914D}",
        "\u{4E0D}\u{5206}\u{5339}\u{914D}",
        "\u{5B8C}\u{5168}\u{4E0D}\u{540C}",
        "\u{6837}\u{672C}\u{4F4D}\u{70B9}\u{7F3A}\u{5931}",
        "P**/P**\u{2026}",
        "",
        "",
        "",
    ];
    let widths = [
        46.0, 44.0, 10.0, 13.0, 13.0, 13.0, 13.0, 13.0, 13.0, 13.0, 13.0,
    ];

    for (idx, header) in headers.iter().enumerate() {
        let col = (idx + 1) as u32;
        sheet
            .get_cell_mut((col, 1u32))
            .set_value((*header).to_string());
        sheet
            .get_cell_mut((col, 2u32))
            .set_value(sub_headers[idx].to_string());
        apply_export_cell_style(sheet, col, 1, true, true);
        apply_export_cell_style(sheet, col, 2, true, true);
        sheet
            .get_column_dimension_by_number_mut(&col)
            .set_width(widths[idx]);
    }
    sheet.get_row_dimension_mut(&1u32).set_height(56.25);
    sheet.get_row_dimension_mut(&2u32).set_height(37.5);

    for (row_index, row) in data.iter().enumerate() {
        let excel_row = (row_index + 3) as u32;
        sheet
            .get_cell_mut((1u32, excel_row))
            .set_value(row.sample_batch_code.clone());
        sheet
            .get_cell_mut((2u32, excel_row))
            .set_value(row.standard_batch_code.clone());
        sheet
            .get_cell_mut((3u32, excel_row))
            .set_value(row.count.to_string());
        sheet
            .get_cell_mut((4u32, excel_row))
            .set_value(row.same_count.to_string());
        sheet
            .get_cell_mut((5u32, excel_row))
            .set_value(if row.partial_count > 0 {
                row.partial_count.to_string()
            } else {
                String::new()
            });
        sheet
            .get_cell_mut((6u32, excel_row))
            .set_value(row.diff_count.to_string());
        sheet
            .get_cell_mut((7u32, excel_row))
            .set_value(row.missing_count.to_string());
        sheet
            .get_cell_mut((8u32, excel_row))
            .set_value(row.same_positions.clone());
        sheet
            .get_cell_mut((9u32, excel_row))
            .set_value(row.partial_positions.clone());
        sheet
            .get_cell_mut((10u32, excel_row))
            .set_value(row.diff_positions.clone());
        sheet
            .get_cell_mut((11u32, excel_row))
            .set_value(row.missing_positions.clone());
        for col in 1u32..=11u32 {
            apply_export_cell_style(sheet, col, excel_row, false, col >= 8);
        }
    }
}

fn apply_export_cell_style(sheet: &mut Worksheet, col: u32, row: u32, bold: bool, wrap_text: bool) {
    let style = sheet.get_style_mut((col, row));
    style.get_font_mut().set_name("\u{5B8B}\u{4F53}");
    style.get_font_mut().set_size(11.0);
    style.get_font_mut().set_bold(bold);

    let alignment = style.get_alignment_mut();
    alignment.set_horizontal(HorizontalAlignmentValues::Center);
    alignment.set_vertical(VerticalAlignmentValues::Center);
    alignment.set_wrap_text(wrap_text);

    let borders = style.get_borders_mut();
    borders.get_top_mut().set_border_style(Border::BORDER_THIN);
    borders
        .get_bottom_mut()
        .set_border_style(Border::BORDER_THIN);
    borders.get_left_mut().set_border_style(Border::BORDER_THIN);
    borders
        .get_right_mut()
        .set_border_style(Border::BORDER_THIN);
}

fn now_millis() -> u128 {
    SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|d| d.as_millis())
        .unwrap_or(0)
}

#[tauri::command]
fn load_contrast_config(path: Option<String>) -> Result<Vec<ContrastRow>, String> {
    let file = path
        .map(PathBuf::from)
        .unwrap_or_else(|| config_file(CONTRAST_CONFIG_FILE));
    let mut entries = load_properties(&file)?;
    sort_table_entries(&mut entries);
    let rows = entries
        .into_iter()
        .map(|(_, value)| contrast_row_from_property(&value))
        .collect();
    Ok(rows)
}

#[tauri::command]
fn save_contrast_config(rows: Vec<ContrastRow>, path: Option<String>) -> Result<String, String> {
    let file = path
        .map(PathBuf::from)
        .unwrap_or_else(|| config_file(CONTRAST_CONFIG_FILE));
    let entries: Vec<(String, String)> = rows
        .iter()
        .enumerate()
        .map(|(idx, row)| (format!("tableBean{}", idx), contrast_row_to_property(row)))
        .collect();
    save_properties(&file, &entries)?;
    Ok(file.to_string_lossy().to_string())
}

#[tauri::command]
fn run_contrast(row: ContrastRow) -> Result<String, String> {
    let standard_list = get_excel_data(Path::new(&row.standard_sample_path))?;
    let sample_list = get_excel_data(Path::new(&row.sample_path))?;
    if standard_list.is_empty() || sample_list.is_empty() {
        return Err(
            "\u{89E3}\u{6790}Excel\u{5F02}\u{5E38},\u{8BF7}\u{68C0}\u{67E5}Excel\u{6570}\u{636E}"
                .to_string(),
        );
    }

    let standard_batch_map = build_batch_data_map(&standard_list);
    let sample_batch_map = build_batch_label_map(&sample_list);

    let normal_result = normal_process(&row, &standard_batch_map, &sample_batch_map, true);
    let real_result = normal_process(&row, &standard_batch_map, &sample_batch_map, false);
    let export_rows = build_export_rows(&normal_result, &real_result);

    let output_dir = Path::new(&row.analysis_results_path);
    let file_name = format!("\u{89E3}\u{6790}\u{7ED3}\u{679C}_{}.xlsx", now_millis());
    let output_path = output_dir.join(file_name);
    ensure_parent_dir(&output_path)?;

    let mut book = umya::new_file();
    if let Some(sheet) = book.get_sheet_by_name_mut("Sheet1") {
        sheet.set_name("\u{6BD4}\u{5BF9}\u{7ED3}\u{679C}");
        write_export_sheet(sheet, &export_rows);
    }

    umya::writer::xlsx::write(&book, &output_path)
        .map_err(|e| format!("\u{5199}\u{5165}\u{7ED3}\u{679C}\u{5931}\u{8D25}: {}", e))?;

    Ok(output_path.to_string_lossy().to_string())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn dna(batch_code: &str, label: &str, a: f64, b: f64, c: f64) -> DnaData {
        DnaData {
            batch_code: batch_code.to_string(),
            label: label.to_string(),
            a,
            b,
            c,
        }
    }

    fn contrast_row_for_test() -> ContrastRow {
        ContrastRow {
            standard_sample_path: String::new(),
            sample_path: String::new(),
            analysis_results_path: String::new(),
            threshold_number: "1".to_string(),
            remarks: String::new(),
        }
    }

    fn summary_row(
        sample_batch_code: &str,
        standard_batch_code: &str,
        count: usize,
        same_count: usize,
        diff_count: usize,
        missing_count: usize,
    ) -> ContrastResultRow {
        ContrastResultRow {
            simple_data_batch_code: Some(sample_batch_code.to_string()),
            standard_sample_data_batch_code: standard_batch_code.to_string(),
            count,
            same_number_bits: same_count,
            different_bits: diff_count,
            missing_bits: missing_count,
        }
    }

    fn detail_row(
        contrast_type: &str,
        sample_batch_code: &str,
        standard_batch_code: &str,
        label: &str,
        category: &str,
    ) -> ContrastDetailRow {
        detail_row_with_sample(
            contrast_type,
            sample_batch_code,
            standard_batch_code,
            label,
            category,
            (1.0, 1.0, 1.0),
        )
    }

    fn detail_row_with_sample(
        contrast_type: &str,
        sample_batch_code: &str,
        standard_batch_code: &str,
        label: &str,
        category: &str,
        sample_values: (f64, f64, f64),
    ) -> ContrastDetailRow {
        ContrastDetailRow {
            contrast_type: contrast_type.to_string(),
            simple_data_batch_code: sample_batch_code.to_string(),
            standard_sample_data_batch_code: standard_batch_code.to_string(),
            label: label.to_string(),
            category: category.to_string(),
            standard_a: 0.0,
            standard_b: 0.0,
            standard_c: 0.0,
            sample_a: sample_values.0,
            sample_b: sample_values.1,
            sample_c: sample_values.2,
        }
    }

    fn write_test_input_excel(path: &Path, batch_code: &str, loci: &[(&str, (f64, f64, f64))]) {
        let mut book = umya::new_file();
        let sheet = book.get_sheet_by_name_mut("Sheet1").unwrap();

        sheet.get_cell_mut((1u32, 1u32)).set_value("id".to_string());
        sheet
            .get_cell_mut((2u32, 1u32))
            .set_value("batch".to_string());
        for (idx, (label, _)) in loci.iter().enumerate() {
            let start_col = 3u32 + (idx as u32) * 3;
            sheet
                .get_cell_mut((start_col, 1u32))
                .set_value((*label).to_string());
        }

        sheet.get_cell_mut((1u32, 2u32)).set_value("1".to_string());
        sheet
            .get_cell_mut((2u32, 2u32))
            .set_value(batch_code.to_string());
        for (idx, (_, (a, b, c))) in loci.iter().enumerate() {
            let start_col = 3u32 + (idx as u32) * 3;
            sheet
                .get_cell_mut((start_col, 2u32))
                .set_value(a.to_string());
            sheet
                .get_cell_mut((start_col + 1, 2u32))
                .set_value(b.to_string());
            sheet
                .get_cell_mut((start_col + 2, 2u32))
                .set_value(c.to_string());
        }

        umya::writer::xlsx::write(&book, path).expect("failed to write test excel");
    }

    fn write_test_input_excel_with_rows(
        path: &Path,
        loci: &[&str],
        rows: &[(&str, Vec<(f64, f64, f64)>)],
    ) {
        let mut book = umya::new_file();
        let sheet = book.get_sheet_by_name_mut("Sheet1").unwrap();

        sheet.get_cell_mut((1u32, 1u32)).set_value("id".to_string());
        sheet
            .get_cell_mut((2u32, 1u32))
            .set_value("batch".to_string());
        for (idx, label) in loci.iter().enumerate() {
            let start_col = 3u32 + (idx as u32) * 3;
            sheet
                .get_cell_mut((start_col, 1u32))
                .set_value((*label).to_string());
        }

        for (row_index, (batch_code, values)) in rows.iter().enumerate() {
            let excel_row = (row_index + 2) as u32;
            sheet
                .get_cell_mut((1u32, excel_row))
                .set_value((row_index + 1).to_string());
            sheet
                .get_cell_mut((2u32, excel_row))
                .set_value((*batch_code).to_string());
            for (idx, (a, b, c)) in values.iter().enumerate() {
                let start_col = 3u32 + (idx as u32) * 3;
                sheet
                    .get_cell_mut((start_col, excel_row))
                    .set_value(a.to_string());
                sheet
                    .get_cell_mut((start_col + 1, excel_row))
                    .set_value(b.to_string());
                sheet
                    .get_cell_mut((start_col + 2, excel_row))
                    .set_value(c.to_string());
            }
        }

        umya::writer::xlsx::write(&book, path).expect("failed to write test excel");
    }

    #[test]
    fn normal_process_compares_all_sample_rows_with_same_label() {
        let standard_list = vec![dna("STD-1", "D8S1179", 10.0, 11.0, 12.0)];
        let sample_list = vec![
            dna("SAMPLE-1", "D8S1179", 10.0, 11.0, 12.0),
            dna("SAMPLE-1", "D8S1179", 5.0, 6.0, 7.0),
        ];
        let standard_batch_map = build_batch_data_map(&standard_list);
        let sample_batch_map = build_batch_label_map(&sample_list);

        let result = normal_process(
            &contrast_row_for_test(),
            &standard_batch_map,
            &sample_batch_map,
            false,
        );

        assert_eq!(result.summary_rows.len(), 1);
        let row = &result.summary_rows[0];
        assert_eq!(row.count, 2);
        assert_eq!(row.same_number_bits, 1);
        assert_eq!(row.different_bits, 1);
        assert_eq!(row.missing_bits, 0);
        assert_eq!(result.detail_rows.len(), 2);
    }

    #[test]
    fn normal_process_uses_dynamic_locus_count_per_standard_batch() {
        let standard_list = vec![
            dna("STD-A", "L1", 10.0, 11.0, 12.0),
            dna("STD-A", "L2", 20.0, 21.0, 22.0),
            dna("STD-B", "L1", 10.0, 11.0, 12.0),
        ];
        let sample_list = vec![dna("SAMPLE-1", "L1", 10.0, 11.0, 12.0)];
        let standard_batch_map = build_batch_data_map(&standard_list);
        let sample_batch_map = build_batch_label_map(&sample_list);

        let result = normal_process(
            &contrast_row_for_test(),
            &standard_batch_map,
            &sample_batch_map,
            false,
        );

        assert_eq!(result.summary_rows.len(), 2);

        let std_a_row = result
            .summary_rows
            .iter()
            .find(|row| row.standard_sample_data_batch_code == "STD-A")
            .expect("STD-A summary row not found");
        assert_eq!(std_a_row.count, 2);
        assert_eq!(std_a_row.same_number_bits, 1);
        assert_eq!(std_a_row.missing_bits, 1);

        let std_b_row = result
            .summary_rows
            .iter()
            .find(|row| row.standard_sample_data_batch_code == "STD-B")
            .expect("STD-B summary row not found");
        assert_eq!(std_b_row.count, 1);
        assert_eq!(std_b_row.same_number_bits, 1);
        assert_eq!(std_b_row.missing_bits, 0);
    }

    #[test]
    fn build_export_rows_maps_counts_and_positions() {
        let normal = ContrastProcessResult {
            summary_rows: vec![summary_row("S-1", "STD-1", 4, 2, 2, 0)],
            detail_rows: vec![
                detail_row(
                    "\u{4EB2}\u{5B50}\u{9274}\u{5B9A}",
                    "S-1",
                    "STD-1",
                    "L2",
                    "\u{5DEE}\u{5F02}",
                ),
                detail_row(
                    "\u{4EB2}\u{5B50}\u{9274}\u{5B9A}",
                    "S-1",
                    "STD-1",
                    "L4",
                    "\u{5DEE}\u{5F02}",
                ),
                detail_row(
                    "\u{4EB2}\u{5B50}\u{9274}\u{5B9A}",
                    "S-1",
                    "STD-1",
                    "L4",
                    "\u{5DEE}\u{5F02}",
                ),
            ],
        };
        let real = ContrastProcessResult {
            summary_rows: vec![summary_row("S-1", "STD-1", 4, 2, 2, 0)],
            detail_rows: vec![
                detail_row(
                    "\u{771F}\u{5B9E}\u{6027}\u{9274}\u{5B9A}",
                    "S-1",
                    "STD-1",
                    "L1",
                    "\u{76F8}\u{540C}",
                ),
                detail_row(
                    "\u{771F}\u{5B9E}\u{6027}\u{9274}\u{5B9A}",
                    "S-1",
                    "STD-1",
                    "L1",
                    "\u{76F8}\u{540C}",
                ),
                detail_row(
                    "\u{771F}\u{5B9E}\u{6027}\u{9274}\u{5B9A}",
                    "S-1",
                    "STD-1",
                    "L2",
                    "\u{5DEE}\u{5F02}",
                ),
                detail_row_with_sample(
                    "\u{771F}\u{5B9E}\u{6027}\u{9274}\u{5B9A}",
                    "S-1",
                    "STD-1",
                    "L3",
                    "\u{5DEE}\u{5F02}",
                    (0.0, 0.0, 0.0),
                ),
                detail_row(
                    "\u{771F}\u{5B9E}\u{6027}\u{9274}\u{5B9A}",
                    "S-1",
                    "STD-1",
                    "L4",
                    "\u{76F8}\u{540C}",
                ),
            ],
        };

        let rows = build_export_rows(&normal, &real);
        assert_eq!(rows.len(), 1);
        let row = &rows[0];
        assert_eq!(row.sample_batch_code, "S-1");
        assert_eq!(row.standard_batch_code, "STD-1");
        assert_eq!(row.count, 4);
        assert_eq!(row.same_count, 1);
        assert_eq!(row.partial_count, 1);
        assert_eq!(row.diff_count, 1);
        assert_eq!(row.missing_count, 1);
        assert_eq!(row.same_positions, "L1");
        assert_eq!(row.partial_positions, "L4");
        assert_eq!(row.diff_positions, "L2");
        assert_eq!(row.missing_positions, "L3");
    }

    #[test]
    fn build_export_rows_keeps_partial_count_when_real_summary_missing() {
        let normal = ContrastProcessResult {
            summary_rows: vec![summary_row("S-2", "STD-2", 1, 0, 1, 0)],
            detail_rows: vec![detail_row(
                "\u{4EB2}\u{5B50}\u{9274}\u{5B9A}",
                "S-2",
                "STD-2",
                "L3",
                "\u{5DEE}\u{5F02}",
            )],
        };
        let real = ContrastProcessResult {
            summary_rows: vec![],
            detail_rows: vec![],
        };

        let rows = build_export_rows(&normal, &real);
        assert_eq!(rows.len(), 1);
        let row = &rows[0];
        assert_eq!(row.sample_batch_code, "S-2");
        assert_eq!(row.standard_batch_code, "STD-2");
        assert_eq!(row.count, 1);
        assert_eq!(row.same_count, 0);
        assert_eq!(row.partial_count, 1);
        assert_eq!(row.diff_count, 0);
        assert_eq!(row.missing_count, 0);
        assert_eq!(row.partial_positions, "L3");
    }

    #[test]
    fn build_export_rows_prioritizes_worst_result_per_locus() {
        let normal = ContrastProcessResult {
            summary_rows: vec![summary_row("S-3", "STD-3", 1, 0, 1, 0)],
            detail_rows: vec![
                detail_row(
                    "\u{4EB2}\u{5B50}\u{9274}\u{5B9A}",
                    "S-3",
                    "STD-3",
                    "C1",
                    "\u{5DEE}\u{5F02}",
                ),
                detail_row_with_sample(
                    "\u{4EB2}\u{5B50}\u{9274}\u{5B9A}",
                    "S-3",
                    "STD-3",
                    "C1",
                    "\u{7F3A}\u{5931}",
                    (0.0, 0.0, 0.0),
                ),
            ],
        };
        let real = ContrastProcessResult {
            summary_rows: vec![summary_row("S-3", "STD-3", 1, 1, 0, 0)],
            detail_rows: vec![
                detail_row(
                    "\u{771F}\u{5B9E}\u{6027}\u{9274}\u{5B9A}",
                    "S-3",
                    "STD-3",
                    "C1",
                    "\u{76F8}\u{540C}",
                ),
                detail_row(
                    "\u{771F}\u{5B9E}\u{6027}\u{9274}\u{5B9A}",
                    "S-3",
                    "STD-3",
                    "C1",
                    "\u{5DEE}\u{5F02}",
                ),
            ],
        };

        let rows = build_export_rows(&normal, &real);
        assert_eq!(rows.len(), 1);
        let row = &rows[0];
        assert_eq!(row.count, 1);
        assert_eq!(row.same_count, 0);
        assert_eq!(row.partial_count, 0);
        assert_eq!(row.diff_count, 0);
        assert_eq!(row.missing_count, 1);
        assert_eq!(row.missing_positions, "C1");
    }

    #[test]
    fn build_export_rows_treats_zero_and_absent_sample_as_missing() {
        let standard_list = vec![
            dna("STD-1", "L1", 10.0, 11.0, 12.0),
            dna("STD-1", "L2", 20.0, 21.0, 22.0),
        ];
        let sample_list = vec![dna("S-1", "L1", 0.0, 0.0, 0.0)];
        let standard_batch_map = build_batch_data_map(&standard_list);
        let sample_batch_map = build_batch_label_map(&sample_list);

        let normal = normal_process(
            &contrast_row_for_test(),
            &standard_batch_map,
            &sample_batch_map,
            true,
        );
        let real = normal_process(
            &contrast_row_for_test(),
            &standard_batch_map,
            &sample_batch_map,
            false,
        );

        let rows = build_export_rows(&normal, &real);
        assert_eq!(rows.len(), 1);
        let row = &rows[0];
        assert_eq!(row.sample_batch_code, "S-1");
        assert_eq!(row.standard_batch_code, "STD-1");
        assert_eq!(row.count, 2);
        assert_eq!(row.same_count, 0);
        assert_eq!(row.partial_count, 0);
        assert_eq!(row.diff_count, 0);
        assert_eq!(row.missing_count, 2);
        assert_eq!(row.missing_positions, "L1/L2");
    }

    #[test]
    fn run_contrast_exports_single_result_sheet() {
        let temp_dir = std::env::temp_dir().join(format!("datalinker_test_{}", now_millis()));
        fs::create_dir_all(&temp_dir).expect("failed to create temp dir");
        let standard_path = temp_dir.join("standard.xlsx");
        let sample_path = temp_dir.join("sample.xlsx");
        write_test_input_excel(
            &standard_path,
            "STD-1",
            &[("L1", (10.0, 11.0, 12.0)), ("L2", (20.0, 21.0, 22.0))],
        );
        write_test_input_excel_with_rows(
            &sample_path,
            &["L1", "L2"],
            &[
                ("SAMPLE-1", vec![(10.0, 11.0, 12.0), (0.0, 0.0, 0.0)]),
                ("SAMPLE-1", vec![(10.0, 11.0, 12.0), (20.0, 21.0, 22.0)]),
            ],
        );

        let row = ContrastRow {
            standard_sample_path: standard_path.to_string_lossy().to_string(),
            sample_path: sample_path.to_string_lossy().to_string(),
            analysis_results_path: temp_dir.to_string_lossy().to_string(),
            threshold_number: "1".to_string(),
            remarks: String::new(),
        };

        let output = run_contrast(row).expect("run_contrast failed");
        let output_path = Path::new(&output);
        assert!(output_path.exists());

        let mut workbook = open_workbook_auto(output_path).expect("failed to open output excel");
        assert_eq!(
            workbook.sheet_names(),
            vec!["\u{6BD4}\u{5BF9}\u{7ED3}\u{679C}".to_string()]
        );

        let range = workbook
            .worksheet_range_at(0)
            .expect("first sheet missing")
            .expect("sheet parse failed");
        let rows: Vec<Vec<String>> = range
            .rows()
            .take(3)
            .map(|row| row.iter().map(cell_to_string).collect())
            .collect();
        assert_eq!(rows[0][0], "\u{6837}\u{54C1}\u{7F16}\u{53F7}");
        assert_eq!(
            rows[0][8],
            "\u{4E0D}\u{5B8C}\u{5168}\u{76F8}\u{540C}\u{4F4D}\u{7F6E}"
        );
        assert_eq!(rows[0][10], "\u{7F3A}\u{5931}\u{4F4D}\u{7F6E}");
        assert_eq!(rows[1][3], "\u{5B8C}\u{5168}\u{5339}\u{914D}");
        assert_eq!(rows[1][4], "\u{4E0D}\u{5206}\u{5339}\u{914D}");
        assert_eq!(
            rows[1][6],
            "\u{6837}\u{672C}\u{4F4D}\u{70B9}\u{7F3A}\u{5931}"
        );
        assert_eq!(rows[2][0], "SAMPLE-1");
        assert_eq!(rows[2][1], "STD-1");
        assert_eq!(rows[2][2], "2");
        assert_eq!(rows[2][3], "1");
        assert_eq!(rows[2][4], "");
        assert_eq!(rows[2][5], "0");
        assert_eq!(rows[2][6], "1");
        assert_eq!(rows[2][7], "L1");
        assert_eq!(rows[2][10], "L2");

        let _ = fs::remove_file(output_path);
        let _ = fs::remove_file(standard_path);
        let _ = fs::remove_file(sample_path);
        let _ = fs::remove_dir_all(temp_dir);
    }
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_store::Builder::default().build())
        .plugin(tauri_plugin_opener::init())
        .setup(|app| {
            cleanup_legacy_group_config();
            cleanup_legacy_group_store(app);
            if let Some(window) = app.get_webview_window("main") {
                let _ = window.set_title("\u{6570}\u{636E}\u{5904}\u{7406}");
                if let Ok(Some(monitor)) = window.primary_monitor() {
                    let work_area = monitor.work_area();
                    let width = (work_area.size.width as f64 * 0.68).round() as u32;
                    let height = (work_area.size.height as f64 * 0.60).round() as u32;
                    let _ = window.set_size(Size::Physical(PhysicalSize::new(width, height)));
                    let _ = window.center();
                }
            }
            Ok(())
        })
        .invoke_handler(tauri::generate_handler![
            load_contrast_config,
            save_contrast_config,
            run_contrast
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}

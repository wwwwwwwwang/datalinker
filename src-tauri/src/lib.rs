// Prevents additional console window on Windows in release, DO NOT REMOVE!!
#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

use calamine::{open_workbook_auto, Data, Reader};
use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::fs;
use std::path::{Path, PathBuf};
use std::time::{SystemTime, UNIX_EPOCH};
use tauri::{Manager, PhysicalSize, Size};
use umya_spreadsheet::{self as umya, Worksheet};

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
        store_paths.push(PathBuf::from(app_data_dir).join("com.admin.datalinker").join("datalinker.store.json"));
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
    let mut workbook = open_workbook_auto(path).map_err(|e| format!("Parse Excel failed: {}", e))?;
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

        let batch_code = row
            .get(1)
            .map(cell_to_string)
            .unwrap_or_default();

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
    let threshold_number = contrast_row
        .threshold_number
        .parse::<i32>()
        .unwrap_or(0);
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
                if let Some(sample_rows) = sample_label_map.and_then(|labels| labels.get(&standard.label)) {
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

    let mut last_simple = String::new();
    let mut seen_first = false;
    for row in &mut contrast_results {
        let current_simple = row.simple_data_batch_code.clone().unwrap_or_default();
        if seen_first && current_simple == last_simple {
            row.simple_data_batch_code = None;
            continue;
        }
        last_simple = current_simple;
        seen_first = true;
    }

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

fn write_contrast_sheet(sheet: &mut Worksheet, data: &[ContrastResultRow]) {
    let headers = [
        "\u{6837}\u{54C1}\u{7F16}\u{53F7}",
        "\u{6807}\u{4F4D}\u{7F16}\u{53F7}",
        "\u{6D4B}\u{5B9A}\u{4F4D}\u{70B9}\u{6570}",
        "\u{76F8}\u{540C}\u{4F4D}\u{70B9}\u{6570}",
        "\u{5DEE}\u{5F02}\u{4F4D}\u{70B9}\u{6570}",
        "\u{7F3A}\u{5931}\u{4F4D}\u{70B9}\u{6570}",
    ];
    for (idx, header) in headers.iter().enumerate() {
        sheet
            .get_cell_mut(((idx + 1) as u32, 1u32))
            .set_value(header.to_string());
    }
    for (row_index, row) in data.iter().enumerate() {
        let excel_row = (row_index + 2) as u32;
        let simple_code = row.simple_data_batch_code.clone().unwrap_or_default();
        sheet
            .get_cell_mut((1u32, excel_row))
            .set_value(simple_code);
        sheet
            .get_cell_mut((2u32, excel_row))
            .set_value(row.standard_sample_data_batch_code.clone());
        sheet
            .get_cell_mut((3u32, excel_row))
            .set_value(row.count.to_string());
        sheet
            .get_cell_mut((4u32, excel_row))
            .set_value(row.same_number_bits.to_string());
        sheet
            .get_cell_mut((5u32, excel_row))
            .set_value(row.different_bits.to_string());
        sheet
            .get_cell_mut((6u32, excel_row))
            .set_value(row.missing_bits.to_string());
    }
}

fn format_detail_number(value: f64) -> String {
    if (value.fract()).abs() < f64::EPSILON {
        format!("{value:.0}")
    } else {
        value.to_string()
    }
}

fn write_contrast_detail_sheet(sheet: &mut Worksheet, data: &[ContrastDetailRow]) {
    let headers = [
        "\u{5BF9}\u{6BD4}\u{7C7B}\u{578B}",
        "\u{6837}\u{54C1}\u{7F16}\u{53F7}",
        "\u{6807}\u{4F4D}\u{7F16}\u{53F7}",
        "\u{4F4D}\u{70B9}",
        "\u{5206}\u{7C7B}",
        "\u{6807}\u{6837}A",
        "\u{6807}\u{6837}B",
        "\u{6807}\u{6837}C",
        "\u{6837}\u{672C}A",
        "\u{6837}\u{672C}B",
        "\u{6837}\u{672C}C",
    ];
    for (idx, header) in headers.iter().enumerate() {
        sheet
            .get_cell_mut(((idx + 1) as u32, 1u32))
            .set_value(header.to_string());
    }

    for (row_index, row) in data.iter().enumerate() {
        let excel_row = (row_index + 2) as u32;
        sheet
            .get_cell_mut((1u32, excel_row))
            .set_value(row.contrast_type.clone());
        sheet
            .get_cell_mut((2u32, excel_row))
            .set_value(row.simple_data_batch_code.clone());
        sheet
            .get_cell_mut((3u32, excel_row))
            .set_value(row.standard_sample_data_batch_code.clone());
        sheet
            .get_cell_mut((4u32, excel_row))
            .set_value(row.label.clone());
        sheet
            .get_cell_mut((5u32, excel_row))
            .set_value(row.category.clone());
        sheet
            .get_cell_mut((6u32, excel_row))
            .set_value(format_detail_number(row.standard_a));
        sheet
            .get_cell_mut((7u32, excel_row))
            .set_value(format_detail_number(row.standard_b));
        sheet
            .get_cell_mut((8u32, excel_row))
            .set_value(format_detail_number(row.standard_c));
        sheet
            .get_cell_mut((9u32, excel_row))
            .set_value(format_detail_number(row.sample_a));
        sheet
            .get_cell_mut((10u32, excel_row))
            .set_value(format_detail_number(row.sample_b));
        sheet
            .get_cell_mut((11u32, excel_row))
            .set_value(format_detail_number(row.sample_c));
    }
}

fn filter_detail_rows_by_category(data: &[ContrastDetailRow], category: &str) -> Vec<ContrastDetailRow> {
    data.iter()
        .filter(|row| row.category == category)
        .cloned()
        .collect()
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
        return Err("\u{89E3}\u{6790}Excel\u{5F02}\u{5E38},\u{8BF7}\u{68C0}\u{67E5}Excel\u{6570}\u{636E}".to_string());
    }

    let standard_batch_map = build_batch_data_map(&standard_list);
    let sample_batch_map = build_batch_label_map(&sample_list);

    let normal_result = normal_process(&row, &standard_batch_map, &sample_batch_map, true);
    let real_result = normal_process(&row, &standard_batch_map, &sample_batch_map, false);
    let mut detail_rows = normal_result.detail_rows.clone();
    detail_rows.extend(real_result.detail_rows.clone());

    let output_dir = Path::new(&row.analysis_results_path);
    let file_name = format!("\u{89E3}\u{6790}\u{7ED3}\u{679C}_{}.xlsx", now_millis());
    let output_path = output_dir.join(file_name);
    ensure_parent_dir(&output_path)?;

    let mut book = umya::new_file();
    if let Some(sheet) = book.get_sheet_by_name_mut("Sheet1") {
        sheet.set_name("\u{4EB2}\u{5B50}\u{9274}\u{5B9A}");
        write_contrast_sheet(sheet, &normal_result.summary_rows);
    }
    let _ = book.new_sheet("\u{771F}\u{5B9E}\u{6027}\u{9274}\u{5B9A}");
    if let Some(sheet) = book.get_sheet_by_name_mut("\u{771F}\u{5B9E}\u{6027}\u{9274}\u{5B9A}") {
        write_contrast_sheet(sheet, &real_result.summary_rows);
    }
    let same_rows = filter_detail_rows_by_category(&detail_rows, "\u{76F8}\u{540C}");
    let diff_rows = filter_detail_rows_by_category(&detail_rows, "\u{5DEE}\u{5F02}");
    let missing_rows = filter_detail_rows_by_category(&detail_rows, "\u{7F3A}\u{5931}");

    let _ = book.new_sheet("\u{76F8}\u{540C}\u{4F4D}\u{70B9}");
    if let Some(sheet) = book.get_sheet_by_name_mut("\u{76F8}\u{540C}\u{4F4D}\u{70B9}") {
        write_contrast_detail_sheet(sheet, &same_rows);
    }

    let _ = book.new_sheet("\u{5DEE}\u{5F02}\u{4F4D}\u{70B9}");
    if let Some(sheet) = book.get_sheet_by_name_mut("\u{5DEE}\u{5F02}\u{4F4D}\u{70B9}") {
        write_contrast_detail_sheet(sheet, &diff_rows);
    }

    let _ = book.new_sheet("\u{7F3A}\u{5931}\u{4F4D}\u{70B9}");
    if let Some(sheet) = book.get_sheet_by_name_mut("\u{7F3A}\u{5931}\u{4F4D}\u{70B9}") {
        write_contrast_detail_sheet(sheet, &missing_rows);
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


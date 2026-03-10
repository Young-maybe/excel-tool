"""
当当拉销量 - 批量处理 Excel

严格保持原脚本的核心处理逻辑，仅做适配：
- 输入：由 UI 选择 input_dir（原脚本为当前目录/弹窗）
- 输出：统一写入 output_dir（原脚本覆盖原文件并在原目录生成存档）

处理规则（来自需求.md/原脚本）：
- 读取 sheet：供应链订单明细表
- 过滤“部门”为空的行
- 选取列：第三方平台单号、销售码、数量 → 去重
- 在“整理结果”sheet：
  - A-C 写入去重后的三列
  - F-G 写入按“销售码”汇总的数量（透视结果）
- 单元格格式：
  - A、B 文本格式（保留 >15 位精度）
  - 数量列使用整数格式
"""
from __future__ import annotations

import re
import shutil
from pathlib import Path
from typing import Callable, List, Optional

import pandas as pd


TARGET_SHEET = "供应链订单明细表"
RESULT_SHEET = "整理结果"
DEPT_COLUMN = "部门"
NEEDED_COLUMNS = ["第三方平台单号", "销售码", "数量"]
NUMBER_FORMAT = "0"
TEXT_FORMAT = "@"


def normalize_identifier(value):
    """保留标识字段精度：>15位强制文本；≤15位纯数字转整数；其他保持原样去掉空白。"""
    if pd.isna(value):
        return pd.NA
    s = str(value).strip()
    if s == "" or s.lower() in {"nan", "none"}:
        return pd.NA
    s_clean = s.replace(",", "")
    if len(s_clean) > 15:
        return s_clean
    if re.fullmatch(r"-?\d+", s_clean):
        try:
            return int(s_clean)
        except Exception:
            return s_clean
    return s_clean


def normalize_quantity(series: pd.Series) -> pd.Series:
    cleaned = series.astype(str).str.replace(",", "", regex=False).str.strip()
    nums = pd.to_numeric(cleaned, errors="coerce").fillna(0)
    return nums


def filter_department(df: pd.DataFrame) -> pd.DataFrame:
    if DEPT_COLUMN not in df.columns:
        raise ValueError(f"缺少必需列：{DEPT_COLUMN}")
    mask = df[DEPT_COLUMN].apply(lambda x: str(x).strip() not in {"", "nan", "None"})
    return df.loc[mask].copy()


def _copy_to_output(input_file: Path, output_dir: Path) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    out_path = output_dir / input_file.name

    # 如果输出目录已存在同名文件，自动生成不冲突的文件名（避免用户打开了输出文件导致复制失败）
    if out_path.exists():
        for i in range(1, 1000):
            candidate = output_dir / f"{input_file.stem}_处理副本{i}{input_file.suffix}"
            if not candidate.exists():
                out_path = candidate
                break

    try:
        shutil.copy2(input_file, out_path)
    except PermissionError as exc:
        # WinError 32: 文件被占用（Excel/WPS打开）
        raise PermissionError(f"文件正在被占用，请先关闭Excel/WPS后重试：{input_file}") from exc
    return out_path


def _create_backup(processing_file: Path) -> Path:
    backup_path = processing_file.with_name(f"{processing_file.stem}.存档.xlsx")
    shutil.copy2(processing_file, backup_path)
    return backup_path


def process_one_file(
    processing_file: Path,
    progress_cb: Optional[Callable[[int, str], None]] = None,
) -> List[str]:
    """处理单个文件（在 output_dir 内部的副本上操作），返回生成/修改的文件路径列表。"""
    def report(p: int, s: str) -> None:
        if progress_cb:
            progress_cb(p, s)

    report(10, "创建存档备份…")
    backup = _create_backup(processing_file)

    report(18, f"读取Excel：{TARGET_SHEET}")
    df_raw = pd.read_excel(
        processing_file,
        sheet_name=TARGET_SHEET,
        dtype=str,
        na_filter=False,
    )

    report(28, "过滤无效数据（部门为空）…")
    df_filtered = filter_department(df_raw)
    missing_cols = [col for col in NEEDED_COLUMNS if col not in df_filtered.columns]
    if missing_cols:
        raise ValueError(f"{processing_file.name} 缺少列：{', '.join(missing_cols)}")

    report(38, "抽取字段并清洗（单号/销售码/数量）…")
    df_selected = df_filtered[NEEDED_COLUMNS].copy()
    for col in ["第三方平台单号", "销售码"]:
        df_selected[col] = df_selected[col].map(normalize_identifier)
    df_selected["数量"] = normalize_quantity(df_selected["数量"])

    report(48, "去重并生成透视汇总…")
    df_selected = df_selected.drop_duplicates()
    df_selected = df_selected.where(pd.notna(df_selected), None)

    pivot_source = df_selected[["销售码", "数量"]].copy()
    pivot_source["销售码文本"] = pivot_source["销售码"].apply(lambda x: "" if pd.isna(x) else str(x).strip())
    pivot_source = pivot_source[pivot_source["销售码文本"] != ""].drop(columns=["销售码文本"])
    pivot = pivot_source.groupby("销售码", dropna=False)["数量"].sum().reset_index()

    df_filtered_to_save = df_filtered.where(pd.notna(df_filtered), None)

    report(70, "写回Excel（主sheet + 整理结果 + 透视）…")
    with pd.ExcelWriter(processing_file, engine="openpyxl", mode="w") as writer:
        df_filtered_to_save.to_excel(writer, sheet_name=TARGET_SHEET, index=False)
        df_selected.to_excel(writer, sheet_name=RESULT_SHEET, index=False, startrow=0, startcol=0)
        pivot.to_excel(writer, sheet_name=RESULT_SHEET, index=False, startrow=0, startcol=5)

        ws = writer.book[RESULT_SHEET]
        report(85, "应用单元格格式与列宽…")
        for col_letter in ("A", "B"):
            for cell in ws[col_letter]:
                cell.number_format = TEXT_FORMAT
        for col_cells in ws.iter_cols(min_col=3, max_col=7, min_row=1, max_row=ws.max_row):
            for cell in col_cells:
                cell.number_format = NUMBER_FORMAT
        ws.column_dimensions["A"].width = 25
        ws.column_dimensions["B"].width = 25

    report(100, "完成")
    return [str(processing_file), str(backup)]


def process_folder(
    input_dir: str,
    output_dir: str,
    progress_cb: Optional[Callable[[int, str], None]] = None,
) -> List[str]:
    """批量处理 input_dir 下的 .xlsx，输出写入 output_dir（支持可视化进度/状态上报）。"""
    def report(p: int, s: str) -> None:
        if progress_cb:
            progress_cb(p, s)

    in_dir = Path(input_dir)
    out_dir = Path(output_dir)
    if not in_dir.exists():
        raise FileNotFoundError(f"输入目录不存在：{input_dir}")

    report(2, "扫描输入文件…")
    excel_files = sorted(
        [
            p
            for p in in_dir.glob("*.xlsx")
            if (not p.stem.endswith("存档"))
            and (not p.name.startswith("~$"))  # Excel临时锁文件，必须忽略
        ]
    )
    if not excel_files:
        raise FileNotFoundError("输入文件夹中未找到需要处理的 .xlsx 文件。")

    outputs: List[str] = []
    total = len(excel_files)
    for i, src in enumerate(excel_files, start=1):
        base = 5 + int((i - 1) / max(total, 1) * 80)
        report(base, f"处理文件 {i}/{total}：{src.name}")
        processing_file = _copy_to_output(src, out_dir)
        report(min(base + 5, 90), f"准备处理副本：{processing_file.name}")
        # 将单文件内部进度映射到该文件的区间内（base ~ base+20）
        def child_progress(p: int, s: str) -> None:
            mapped = base + int(p * 0.25)  # 0~100 -> base~base+25
            report(min(mapped, 95), s)

        outputs.extend(process_one_file(processing_file, progress_cb=child_progress))
        report(min(base + 25, 95), f"已完成：{src.name}")
    report(100, "全部处理完成")
    return outputs



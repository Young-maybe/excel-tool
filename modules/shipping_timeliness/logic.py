"""
发货时效表处理（从“发货时效初步处理数据2终版.py”迁入）

约束：
- 不改中间处理“强逻辑”，只做 I/O 适配
- 输入：由 UI 选择的单个 CSV 文件（需符合命名规范）
- 输出：写入程序统一输出目录下的子目录，并避免同名/占用冲突
"""

from __future__ import annotations

import logging
import re
from datetime import datetime
from pathlib import Path
from typing import Callable, List, Optional

import pandas as pd
import xlsxwriter

logger = logging.getLogger(__name__)


# 内置表头配置（保持原脚本内容）
HEADER_CONFIG = {
    "columns": [
        "单据编号",
        "审核时间",
        "付款时间",
        "发货时间",
        "外仓出库时间",
        "实际发货时间",
        "24小时发货时效",
        "48小时发货时效",
        "仓库名称",
        "店铺类型",
        "订单类型",
        "店铺名称",
        "平台单号",
        "会员名称",
        "会员代码",
        "商品类别",
        "品牌",
        "商品代码",
        "商品名称",
        "商品简称",
        "规格代码",
        "规格名称",
        "数量",
        "总重量",
        "成本单价",
        "折扣",
        "标准单价",
        "标准金额",
        "实际单价",
        "实际金额",
        "让利金额",
        "让利后金额",
        "成本总价",
        "物流费用",
        "物流成本",
        "商品标准利润",
        "商品实际利润",
        "买家备注",
        "卖家备注",
        "业务员",
        "物流公司",
        "物流单号",
        "收货人",
        "收货人电话",
        "收货人手机",
        "地区信息",
        "收货地址",
        "批次信息",
        "商品单位",
        "商品重量",
        "总体积",
        "发票抬头",
        "发票内容",
        "纳税人识别号",
        "是否货到付款",
        "到付金额",
        "外仓单据",
        "唯一码总实际进价",
        "平台商品名称",
        "平台规格名称",
        "发票类型",
        "明细备注",
        "其他服务费",
        "省",
        "市",
        "区/县",
        "赠品",
        "赠品来源",
        "多包裹物流单号",
        "组合商品名称",
        "其他补贴",
        "平台折扣金额",
        "唯一码",
        "分销商客户名称",
        "平台附加单号",
        "线下备注",
        "作者",
        "促销期段",
        "名人推荐",
        "年龄段",
        "引入/淘汰备注",
        "生产/出品方",
        "获奖信息",
    ],
    "column_widths": {
        "A": 17.3636363636364,
        "B": 17.3636363636364,
        "G": 13.1909090909091,
        "C": 16.73,
        "D": 16.73,
        "E": 16.73,
        "F": 16.73,
        "I": 46.6363636363636,
        "J": 9.81818181818182,
        "M": 21.9,
        "N": 9.81818181818182,
        "O": 12.8181818181818,
        "P": 9.81818181818182,
        "AB": 11.8181818181818,
        "AC": 9.81818181818182,
        "AD": 11.8181818181818,
        "AE": 11.8181818181818,
        "AF": 10.6363636363636,
        "AG": 9.81818181818182,
        "AH": 11.8181818181818,
        "AJ": 12.8181818181818,
        "AK": 9.81818181818182,
    },
}


def _get_current_month() -> int:
    return datetime.now().month


def _extract_month_from_filename(filename: str) -> Optional[str]:
    match = re.search(r"发货商品详情(\d+)月\.csv$", filename)
    if match:
        return match.group(1)
    return None


def _require_valid_input_name(input_csv: Path) -> str:
    month = _extract_month_from_filename(input_csv.name)
    if not month:
        raise ValueError("CSV 文件名需符合命名规范：发货商品详情X月.csv（例如：发货商品详情8月.csv）")
    return month


def _unique_path(path: Path) -> Path:
    """
    避免 WinError 32/同名覆盖：若目标存在则生成 “*_处理副本N.ext”。
    """
    if not path.exists():
        return path

    stem = path.stem
    suffix = path.suffix
    parent = path.parent
    for i in range(1, 1000):
        candidate = parent / f"{stem}_处理副本{i}{suffix}"
        if not candidate.exists():
            return candidate
    raise RuntimeError("输出目录中同名文件过多，请清理后重试。")


def _detect_csv_encoding_and_columns(csv_path: Path) -> tuple[str, list[str]]:
    """
    尝试多种编码读取表头（nrows=0），返回可用编码与列名列表。
    目的：后续用 usecols 只读必要列，提升 40MB+ CSV 的读取速度与内存占用。
    """
    encodings = ["utf-8", "gbk", "gb2312", "utf-8-sig", "cp936"]
    last_error: Exception | None = None

    for encoding in encodings:
        try:
            df0 = pd.read_csv(csv_path, encoding=encoding, nrows=0)
            return encoding, list(df0.columns)
        except UnicodeDecodeError as exc:
            last_error = exc
            continue
        except Exception as exc:
            last_error = exc
            continue

    raise ValueError(f"读取 CSV 表头失败（已尝试多种编码）。文件：{csv_path}\n错误：{last_error}")


def _normalize_col(name: str) -> str:
    """
    归一化列名，解决 CSV 列名包含 BOM/首尾空格/不可见字符导致的“只匹配到第一列”等问题。
    """
    # 去 BOM
    s = name.replace("\ufeff", "")
    # 去首尾空白（含 \t \r \n）
    s = s.strip()
    return s


_COL_KEY_CLEAN_RE = re.compile(r"[\s\u3000\u200b\u200c\u200d\ufeff]+")


def _col_key(name: str) -> str:
    """
    用于“匹配列名”的强归一化 key：
    - 去 BOM
    - 去所有空白（含全角空格、零宽空格等）
    目的：解决“看起来一样但匹配不上”的列名问题。
    """
    s = str(name).replace("\ufeff", "")
    s = _COL_KEY_CLEAN_RE.sub("", s)
    return s


def _canonicalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    将 DataFrame 的列名尽可能映射到“标准列名”（HEADER_CONFIG + 过滤/计算所需列）。
    这样后续逻辑始终只用标准列名访问，不受 CSV 列名脏数据影响。
    """
    want = set(HEADER_CONFIG["columns"]) | {
        "店铺名称",
        "仓库名称",
        "物流公司",
        "审核时间",
        "发货时间",
        "外仓出库时间",
        "单据编号",
    }

    # 建立 key->实际列名映射（如果 key 重复，优先保留“原样 strip 后完全等于标准名”的列）
    key_to_actual: dict[str, str] = {}
    for c in df.columns:
        key = _col_key(c)
        if key not in key_to_actual:
            key_to_actual[key] = c

    rename_map: dict[str, str] = {}
    for canonical in want:
        key = _col_key(canonical)
        actual = key_to_actual.get(key)
        if actual is not None:
            rename_map[actual] = canonical

    # 额外：对剩余列至少做一次轻度 strip/BOM 清理，避免“\ufeff单据编号”之类残留
    cleaned_others = {c: _normalize_col(str(c)) for c in df.columns if c not in rename_map}
    rename_map.update(cleaned_others)

    return df.rename(columns=rename_map)


def _build_usecols(all_columns: list[str]) -> tuple[list[str] | None, dict[str, str]]:
    """
    返回：(usecols_original_names_or_None, normalized_to_original_map)
    - usecols 使用“原始列名”，确保 pandas 能正确选列
    - 读取后我们会把列名统一 rename 成归一化名，后续逻辑都用归一化名匹配
    """
    normalized_to_original: dict[str, str] = {}
    for c in all_columns:
        n = _normalize_col(str(c))
        # 避免同名覆盖：保留首次出现的原始列名
        normalized_to_original.setdefault(n, c)

    # 这里不再强依赖 usecols 子集（真实 CSV 列名可能有全角空格/零宽字符，容易选错导致“只剩第一列”）。
    # 保留 mapping 供后续 rename 使用，但读取时优先保证“读到列”。
    return None, normalized_to_original


def _read_csv_data(csv_path: Path) -> pd.DataFrame:
    encoding, all_columns = _detect_csv_encoding_and_columns(csv_path)
    usecols, normalized_to_original = _build_usecols(all_columns)

    logger.info("Reading CSV: %s (encoding=%s usecols=%s)", csv_path, encoding, "subset" if usecols else "all")
    df = pd.read_csv(csv_path, encoding=encoding, usecols=usecols)
    return _canonicalize_columns(df)


def _iter_csv_chunks(csv_path: Path, chunksize: int = 20000):
    """
    分块读取 CSV（用于大文件流式处理）。
    返回：encoding, iterator[DataFrame]
    """
    encoding, all_columns = _detect_csv_encoding_and_columns(csv_path)
    usecols, normalized_to_original = _build_usecols(all_columns)

    it = pd.read_csv(csv_path, encoding=encoding, usecols=usecols, chunksize=chunksize)

    def _renaming_iter():
        for chunk in it:
            yield _canonicalize_columns(chunk)

    return encoding, _renaming_iter()


def _filter_data(df: pd.DataFrame) -> pd.DataFrame:
    # 原脚本逻辑：删除指定条件的行
    if "店铺名称" in df.columns:
        df = df[~df["店铺名称"].str.contains("光尘|测试", na=False)]

    if "仓库名称" in df.columns:
        warehouses_to_delete = [
            "虚拟仓",
            "测试仓",
            "自营_光尘永清电商仓",
            "自营_光尘电商分销仓",
            "缺货商品虚拟仓",
            "清元-虚拟仓",
        ]
        df = df[~df["仓库名称"].isin(warehouses_to_delete)]

    if "物流公司" in df.columns:
        df = df[~df["物流公司"].str.contains("物流|自提", na=False)]

    return df


def _calculate_actual_ship_time(df: pd.DataFrame) -> pd.Series:
    """
    向量化计算“实际发货时间”（保持原逻辑的优先级与兜底策略）：
    - 两列都空：输出空字符串
    - 只有一列有值：输出原始值（不强制改格式）
    - 两列都有值且都能解析为时间：取更早者，并格式化为 %Y-%m-%d %H:%M:%S
    - 两列都有值但解析失败：优先输出“外仓出库时间”的原始值
    """
    if "发货时间" not in df.columns or "外仓出库时间" not in df.columns:
        return pd.Series([""] * len(df))

    ship_raw = df["发货时间"]
    wh_raw = df["外仓出库时间"]

    ship_dt = pd.to_datetime(ship_raw, errors="coerce")
    wh_dt = pd.to_datetime(wh_raw, errors="coerce")

    both_missing = ship_raw.isna() & wh_raw.isna()
    wh_missing = wh_raw.isna() & ~ship_raw.isna()
    ship_missing = ship_raw.isna() & ~wh_raw.isna()

    # 两个都存在且都能解析
    parse_ok = ship_dt.notna() & wh_dt.notna()
    min_dt = pd.concat([ship_dt, wh_dt], axis=1).min(axis=1)
    min_str = min_dt.dt.strftime("%Y-%m-%d %H:%M:%S")

    # 两个都存在但有解析失败：按原脚本兜底优先外仓出库时间
    parse_fail = (~ship_raw.isna()) & (~wh_raw.isna()) & (~parse_ok)

    out = pd.Series([""] * len(df), index=df.index, dtype="object")
    out.loc[both_missing] = ""
    out.loc[wh_missing] = ship_raw.loc[wh_missing]
    out.loc[ship_missing] = wh_raw.loc[ship_missing]
    out.loc[parse_ok] = min_str.loc[parse_ok]
    out.loc[parse_fail] = wh_raw.loc[parse_fail]
    return out


def _calculate_delivery_efficiency(df: pd.DataFrame, actual_ship_times: pd.Series) -> tuple[pd.Series, pd.Series]:
    """
    向量化计算 24/48 小时时效（与原逻辑一致）：
    - 实际发货时间为空：未发货
    - 审核时间为空：空字符串
    - 任一时间解析失败：空字符串
    """
    if "审核时间" not in df.columns:
        empty = pd.Series([""] * len(df), index=df.index, dtype="object")
        return empty, empty

    audit_raw = df["审核时间"]
    audit_dt = pd.to_datetime(audit_raw, errors="coerce")
    actual_dt = pd.to_datetime(actual_ship_times, errors="coerce")

    actual_empty = actual_ship_times.isna() | (actual_ship_times.astype("string").fillna("") == "")
    audit_missing = audit_raw.isna()

    parse_ok = (~actual_empty) & (~audit_missing) & actual_dt.notna() & audit_dt.notna()
    days_diff = (actual_dt - audit_dt).dt.total_seconds() / (24 * 3600)

    e24 = pd.Series([""] * len(df), index=df.index, dtype="object")
    e48 = pd.Series([""] * len(df), index=df.index, dtype="object")

    e24.loc[actual_empty] = "未发货"
    e48.loc[actual_empty] = "未发货"

    e24.loc[parse_ok] = (days_diff.loc[parse_ok] > 1).map(lambda x: "不满足" if x else "满足")
    e48.loc[parse_ok] = (days_diff.loc[parse_ok] > 2).map(lambda x: "不满足" if x else "满足")

    return e24, e48


def _col_letter_to_index(col: str) -> int:
    """
    Excel 列字母转 0-based index（A->0, B->1, Z->25, AA->26 ...）
    """
    idx = 0
    for ch in col.upper():
        if not ("A" <= ch <= "Z"):
            raise ValueError(f"非法列字母：{col}")
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx - 1


def _create_excel_with_builtin_config(
    csv_path: Path,
    output_xlsx_path: Path,
    progress_cb: Optional[Callable[[int, str], None]] = None,
) -> Path:
    """
    性能优化点：
    - pandas 向量化计算（替代 iterrows）
    - 分块读取 CSV + xlsxwriter 流式写出（避免一次性载入/避免 pandas.to_excel 的默认表头格式）
    """
    header_columns = HEADER_CONFIG["columns"]
    output_xlsx_path = _unique_path(output_xlsx_path)
    month = _extract_month_from_filename(csv_path.name) or str(_get_current_month())
    main_sheet_name = f"发货商品详情{month}月"

    def _estimate_total_rows(p: Path) -> int:
        """
        粗略估算 CSV 数据行数（不含表头）。
        40MB 级别文件该扫描开销通常可接受，换来更“可信”的进度百分比。
        """
        try:
            with p.open("rb") as f:
                count = 0
                for chunk in iter(lambda: f.read(1024 * 1024), b""):
                    count += chunk.count(b"\n")
            # 至少 1 行数据
            return max(count - 1, 1)
        except Exception:
            return 1

    total_rows_est = _estimate_total_rows(csv_path)

    def _write_workbook(path: Path) -> None:
        workbook = xlsxwriter.Workbook(str(path), {"constant_memory": True})
        try:
            ws_main = workbook.add_worksheet(main_sheet_name)
            ws_24 = workbook.add_worksheet("24")
            ws_48 = workbook.add_worksheet("48")

            # 主表列宽（按原配置）
            for col_letter, width in HEADER_CONFIG["column_widths"].items():
                col_idx = _col_letter_to_index(col_letter)
                ws_main.set_column(col_idx, col_idx, width)

            # 24/48 sheet 列宽（与原脚本一致）
            for ws in (ws_24, ws_48):
                ws.set_column(0, 0, 17.36)
                ws.set_column(1, 1, 13.19)
                ws.set_column(2, 2, 46.64)

            # 显式表头格式：无边框（避免你反馈的“第一行黑框线”）
            header_plain = workbook.add_format({"border": 0})
            header_yellow = workbook.add_format({"border": 0, "bg_color": "#FFFF00"})
            yellow_targets = {"仓库名称", "店铺名称", "商品名称", "物流公司"}

            for col_idx, col_name in enumerate(header_columns):
                fmt = header_yellow if col_name in yellow_targets else header_plain
                ws_main.write(0, col_idx, col_name, fmt)

            # 24/48 表头（无边框）
            ws_24.write(0, 0, "单据编号", header_plain)
            ws_24.write(0, 1, "24小时发货时效", header_plain)
            ws_24.write(0, 2, "仓库名称", header_plain)

            ws_48.write(0, 0, "单据编号", header_plain)
            ws_48.write(0, 1, "48小时发货时效", header_plain)
            ws_48.write(0, 2, "仓库名称", header_plain)

            # 分块读取 + 写主表
            _, chunks = _iter_csv_chunks(csv_path, chunksize=20000)
            out_row = 1  # Excel 行号：0 是表头
            processed_rows = 0

            # 维护 24/48 去重集合（保持“首次出现保留”）
            seen24 = set()
            seen48 = set()
            eff24_rows: list[tuple[object, object, object]] = []
            eff48_rows: list[tuple[object, object, object]] = []

            for chunk in chunks:
                chunk = _filter_data(chunk)
                if len(chunk) == 0:
                    continue

                actual_ship_times = _calculate_actual_ship_time(chunk)
                efficiency_24h, efficiency_48h = _calculate_delivery_efficiency(chunk, actual_ship_times)

                # 主表写入：constant_memory 模式要求按“行”顺序写，不能按列批量写，否则会出现“只有第一列有内容”
                col_series: list[pd.Series] = []
                for col_name in header_columns:
                    if col_name in chunk.columns:
                        s = chunk[col_name]
                    elif col_name == "实际发货时间":
                        s = actual_ship_times
                    elif col_name == "24小时发货时效":
                        s = efficiency_24h
                    elif col_name == "48小时发货时效":
                        s = efficiency_48h
                    else:
                        s = pd.Series([""] * len(chunk), index=chunk.index)
                    col_series.append(s.where(pd.notna(s), ""))

                # 逐行写入（itertuples/zip 比 iterrows 更快）
                for row_values in zip(*[s.tolist() for s in col_series]):
                    ws_main.write_row(out_row, 0, list(row_values))
                    out_row += 1
                    processed_rows += 1

                # 构建 24/48 去重数据（三列组合）
                order_no = chunk["单据编号"] if "单据编号" in chunk.columns else pd.Series([""] * len(chunk), index=chunk.index)
                warehouse = chunk["仓库名称"] if "仓库名称" in chunk.columns else pd.Series([""] * len(chunk), index=chunk.index)
                order_no = order_no.where(pd.notna(order_no), "")
                warehouse = warehouse.where(pd.notna(warehouse), "")
                e24 = efficiency_24h.where(pd.notna(efficiency_24h), "")
                e48 = efficiency_48h.where(pd.notna(efficiency_48h), "")

                for a, b, c in zip(order_no.tolist(), e24.tolist(), warehouse.tolist()):
                    key = (a, b, c)
                    if key not in seen24:
                        seen24.add(key)
                        eff24_rows.append(key)

                for a, b, c in zip(order_no.tolist(), e48.tolist(), warehouse.tolist()):
                    key = (a, b, c)
                    if key not in seen48:
                        seen48.add(key)
                        eff48_rows.append(key)

                # 进度：主表行写入占大头（30%~85%）
                # 注意：这里不能太频繁上报，避免拖慢写入；每块上报一次即可。
                if progress_cb:
                    pct = 30 + int(min(1.0, processed_rows / max(total_rows_est, 1)) * 55)
                    progress_cb(pct, f"写入主表… {processed_rows}/{total_rows_est}")

            # 写 24/48 sheet 数据
            for i, (a, b, c) in enumerate(eff24_rows, start=1):
                ws_24.write(i, 0, a)
                ws_24.write(i, 1, b)
                ws_24.write(i, 2, c)

            for i, (a, b, c) in enumerate(eff48_rows, start=1):
                ws_48.write(i, 0, a)
                ws_48.write(i, 1, b)
                ws_48.write(i, 2, c)

        finally:
            workbook.close()

    if progress_cb:
        progress_cb(20, "开始写入 Excel（主表/24/48）…")

    try:
        _write_workbook(output_xlsx_path)
    except PermissionError:
        output_xlsx_path = _unique_path(output_xlsx_path.with_name(output_xlsx_path.stem + "_处理副本.xlsx"))
        _write_workbook(output_xlsx_path)

    return output_xlsx_path


def process_csv(
    input_csv: str,
    output_dir: str,
    progress_cb: Optional[Callable[[int, str], None]] = None,
) -> List[str]:
    """
    输入：单个 CSV 文件路径（必须符合：发货商品详情X月.csv）
    输出：Excel 文件（主表 + 24/48 sheet）
    """
    input_path = Path(input_csv)
    if not input_path.exists():
        raise FileNotFoundError(f"未找到输入 CSV 文件：{input_path}")
    if input_path.suffix.lower() != ".csv":
        raise ValueError("请选择 .csv 文件作为输入。")

    month = _require_valid_input_name(input_path)

    out_root = Path(output_dir)
    if not out_root.exists():
        raise FileNotFoundError(f"未找到输出目录：{out_root}")

    # 统一输出口下的子目录，避免根目录变乱
    out_folder = out_root / "发货时效表处理"
    out_folder.mkdir(parents=True, exist_ok=True)

    output_filename = f"发货商品详情（{month}月）.xlsx"
    output_path = out_folder / output_filename

    logger.info("Shipping timeliness processing: input=%s output=%s", input_path, output_path)
    if progress_cb:
        progress_cb(5, "检查输入与输出路径…")
        progress_cb(15, "准备输出文件…")
    saved_path = _create_excel_with_builtin_config(input_path, output_path, progress_cb=progress_cb)
    if progress_cb:
        progress_cb(100, "完成")
    return [str(saved_path)]



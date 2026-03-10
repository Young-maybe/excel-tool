"""
使用统计（节省时长）服务

需求要点：
- 每个模块有“定义节省时长”（固定或按产出动态计算）
- 每次运行记录：运行耗时、定义节省时长、实际节省时长=定义-运行耗时（最小为0）
- 本地持久化：不随程序关闭重置；删除程序或迁移到新设备才会重置
- 报表：按“大模块”总节省时长降序；同时展示子模块明细；单位小时

实现原则：
- 简单、稳定、易维护（符合 forapp.mdc）
- 统计失败不影响主流程
"""

from __future__ import annotations

import hashlib
import json
import platform
import time
import uuid
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Tuple


def _project_root() -> Path:
    import sys

    if hasattr(sys, "frozen") and getattr(sys, "frozen"):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parents[1]


def _device_fingerprint() -> str:
    raw = f"{platform.node()}|{platform.platform()}|{uuid.getnode()}"
    return hashlib.sha256(raw.encode("utf-8", errors="ignore")).hexdigest()[:16]


def _stats_file_path() -> Path:
    root = _project_root()
    return root / "resources" / "usage_stats.json"


def _now_ts() -> int:
    return int(time.time())


def clamp_non_negative(x: float) -> float:
    return x if x > 0 else 0.0


def hours_str(seconds: float) -> str:
    return f"{float(seconds) / 3600.0:.2f}"


# ----------------- 模块定义（key -> (label, group_label)） -----------------

MODULE_DEFS: Dict[str, Tuple[str, str]] = {
    # 一级（无子项）
    "dangdang": ("当当切出拉销量", "当当切出拉销量"),
    "yituidanweiruku": ("已推单未入库表处理", "已推单未入库表处理"),
    "fahuoshixiao": ("发货时效表处理", "发货时效表处理"),
    "pandianfenxi": ("盘点的分析处理", "盘点的分析处理"),
    # 二级：退货入库时效
    "tuihuo_chubu": ("表的初步处理", "退货入库时效"),
    "tuihuo_jisuan": ("表的时效计算", "退货入库时效"),
    # 二级：盘点的初步处理
    "pandian_guanyi": ("管易基础表处理工具1", "盘点的初步处理"),
    "pandian_baihe": ("百合基础表处理工具2", "盘点的初步处理"),
    "pandian_cangku": ("仓库实盘表处理工具3", "盘点的初步处理"),
    # 二级：B2B
    "b2b_songhuobeihuo": ("送货单、备货单生成", "猴面包树B2B发货"),
    "b2b_mubanpipei": ("送货单与模板匹配（需要备货单的规格箱数）", "猴面包树B2B发货"),
    "b2b_tihuodan": ("提货单生成（记得填sdo和箱数）", "猴面包树B2B发货"),
    "b2b_xiangma": ("箱唛转换", "猴面包树B2B发货"),
    # 二级：库存处理工具（之前未接入统计；现在加入报表）
    "stock_weipeihuo": ("未配货", "库存处理工具"),
    "stock_weifahuo": ("未发货", "库存处理工具"),
    "stock_kucunbaobiao": ("库存报表", "库存处理工具"),
}


def baseline_seconds_fixed(minutes: float = 0.0, hours: float = 0.0) -> float:
    return float(minutes) * 60.0 + float(hours) * 3600.0


BASELINE_FIXED: Dict[str, float] = {
    "dangdang": baseline_seconds_fixed(minutes=10),
    "yituidanweiruku": baseline_seconds_fixed(hours=0.5),
    "fahuoshixiao": baseline_seconds_fixed(hours=0.5),
    "tuihuo_chubu": baseline_seconds_fixed(hours=0.5),
    "tuihuo_jisuan": baseline_seconds_fixed(minutes=10),
    "pandianfenxi": baseline_seconds_fixed(hours=1),
    "pandian_guanyi": baseline_seconds_fixed(minutes=15),
    "pandian_baihe": baseline_seconds_fixed(minutes=15),
    "pandian_cangku": baseline_seconds_fixed(minutes=30),
    "b2b_xiangma": baseline_seconds_fixed(minutes=15),
    # 库存处理工具（如需更精确，可按你的经验调整下面三项）
    "stock_weipeihuo": baseline_seconds_fixed(hours=1),
    "stock_weifahuo": baseline_seconds_fixed(minutes=12),
    "stock_kucunbaobiao": baseline_seconds_fixed(minutes=8),
}


# ----------------- 持久化 -----------------


def _load_state() -> Dict[str, Any]:
    path = _stats_file_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    if not path.exists():
        return {"device": _device_fingerprint(), "events": []}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        if not isinstance(data, dict):
            return {"device": _device_fingerprint(), "events": []}
        if data.get("device") != _device_fingerprint():
            return {"device": _device_fingerprint(), "events": []}
        if "events" not in data or not isinstance(data["events"], list):
            data["events"] = []
        return data
    except Exception:
        return {"device": _device_fingerprint(), "events": []}


def _save_state(state: Dict[str, Any]) -> None:
    path = _stats_file_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")


@dataclass(frozen=True)
class UsageEvent:
    ts: int
    key: str
    label: str
    runtime_sec: float
    baseline_sec: float
    saved_sec: float


def record_event(key: str, runtime_sec: float, baseline_sec: float) -> UsageEvent:
    if key not in MODULE_DEFS:
        raise ValueError(f"未知模块key：{key}")
    label, _group = MODULE_DEFS[key]
    runtime_sec = float(runtime_sec)
    baseline_sec = float(baseline_sec)
    saved_sec = clamp_non_negative(baseline_sec - runtime_sec)

    ev = UsageEvent(
        ts=_now_ts(),
        key=key,
        label=label,
        runtime_sec=runtime_sec,
        baseline_sec=baseline_sec,
        saved_sec=saved_sec,
    )

    state = _load_state()
    events = state.get("events", [])
    events.append(
        {
            "ts": ev.ts,
            "key": ev.key,
            "label": ev.label,
            "runtime_sec": ev.runtime_sec,
            "baseline_sec": ev.baseline_sec,
            "saved_sec": ev.saved_sec,
        }
    )
    state["events"] = events
    state["device"] = _device_fingerprint()
    _save_state(state)
    return ev


def list_events() -> List[UsageEvent]:
    state = _load_state()
    out: List[UsageEvent] = []
    for raw in state.get("events", []):
        try:
            out.append(
                UsageEvent(
                    ts=int(raw.get("ts", 0)),
                    key=str(raw.get("key", "")),
                    label=str(raw.get("label", "")),
                    runtime_sec=float(raw.get("runtime_sec", 0.0)),
                    baseline_sec=float(raw.get("baseline_sec", 0.0)),
                    saved_sec=float(raw.get("saved_sec", 0.0)),
                )
            )
        except Exception:
            continue
    return out


@dataclass(frozen=True)
class ModuleAgg:
    key: str
    label: str
    group: str
    runs: int
    total_saved_sec: float


@dataclass(frozen=True)
class UsageReport:
    total_saved_sec: float
    groups_sorted: List[str]
    group_totals: Dict[str, float]
    modules_by_group: Dict[str, List[ModuleAgg]]


def build_report() -> UsageReport:
    events = list_events()

    by_key: Dict[str, List[UsageEvent]] = {}
    for e in events:
        if e.key not in MODULE_DEFS:
            continue
        by_key.setdefault(e.key, []).append(e)

    modules: List[ModuleAgg] = []
    total_saved = 0.0
    for key, evs in by_key.items():
        label, group = MODULE_DEFS[key]
        saved = sum(x.saved_sec for x in evs)
        total_saved += saved
        modules.append(ModuleAgg(key=key, label=label, group=group, runs=len(evs), total_saved_sec=saved))

    group_totals: Dict[str, float] = {}
    for m in modules:
        group_totals[m.group] = group_totals.get(m.group, 0.0) + m.total_saved_sec

    groups_sorted = sorted(group_totals.keys(), key=lambda g: group_totals[g], reverse=True)

    modules_by_group: Dict[str, List[ModuleAgg]] = {g: [] for g in groups_sorted}
    for m in modules:
        modules_by_group.setdefault(m.group, []).append(m)
    for g in modules_by_group:
        modules_by_group[g] = sorted(modules_by_group[g], key=lambda x: x.total_saved_sec, reverse=True)

    return UsageReport(
        total_saved_sec=total_saved,
        groups_sorted=groups_sorted,
        group_totals=group_totals,
        modules_by_group=modules_by_group,
    )



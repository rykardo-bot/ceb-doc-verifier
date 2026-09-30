# -*- coding: utf-8 -*-
"""双表对账导入引擎（Issue #2：己方台账 × 联运公司对账单）。

职责（任务书§一~§五）：
  1. 解析两份真实结构不同的 Excel（列名容错，不依赖列序）；
  2. 经站编维护表把 发站/口岸/到站 名称解析为站编；查不到的站名收集为
     "待补录"，不瞎猜（沿用编号引擎的保守口径）；
  3. 用内部匹配码（发运日期+发站+口岸+到站）两表配对。己方到站精确到
     具体车站、联运只到国家层级时，按国家层级兜底配对（任务书§四）；
  4. 四状态：已确认（合计一致）/ 待确认（合计不一致）/ 己方缺失 /
     联运缺失。**只以结算合计是否一致作为匹配成败标准；科目细项差异
     （如7-18代理费并入铁路运费）仅展示，不触发异常（任务书§二）。**

纯解析/比对逻辑不写库；入库（train_trips/train_settlements/
train_prepayments）由 apply_preview 调 train_store 完成。
"""

from __future__ import annotations

import io
import re
from datetime import date

from openpyxl import load_workbook

import train_number
import train_recon
import train_store
from datacheck_import import _norm_header, parse_dep_cell

# ---------------------------------------------------------------- 状态

CONFIRMED = "confirmed"            # 配对成功且合计一致
PENDING = "pending"                # 配对成功但合计不一致
OWN_MISSING = "own_missing"        # 联运有、己方无
AGENT_MISSING = "agent_missing"    # 己方有、联运无
UNRESOLVED_NAME = "unresolved_name"  # 站名清洗后仍未登记（去补录，Issue#3 Bug2）

STATUS_LABELS = {
    CONFIRMED: "已确认（合计一致）",
    PENDING: "待确认（合计不一致）",
    OWN_MISSING: "己方缺失",
    AGENT_MISSING: "联运缺失",
    UNRESOLVED_NAME: "站名未登记（待补录）",
}

# ---------------------------------------------------------------- 列名别名

# 己方台账列：字段键 → 表头别名（归一化后比对；归一化剔除非字母数字汉字）
OWN_HEADER_ALIASES = {
    "seq": ["序号", "no"],
    "dep_date": ["发运日期", "日期", "开行日期"],
    "route": ["班列路线", "路线", "路线信息", "班列路线信息"],
    "pay_date": ["付款日期", "支付日期"],
    "prepay": ["预付金额", "预付款", "预付"],
    "freight": ["运费", "铁路运费"],
    "customs": ["报关费", "报关"],
    "agency": ["代理费", "服务费"],
    "total": ["结算金额合计", "合计", "结算合计", "结算金额"],
}

# 联运对账单列
AGENT_HEADER_ALIASES = {
    "year": ["年份", "年"],
    "seq": ["序号"],
    "month": ["月份", "月"],
    "month_seq": ["月序号"],
    "dep_date": ["开行日期", "发运日期", "日期"],
    "station": ["发站", "始发站"],
    "port": ["口岸", "出境口岸"],
    "dest": ["到站", "目的地", "目的国"],
    "goods_name": ["货物品名", "品名"],
    "wagon_count": ["车数", "车皮数"],
    "container_40hd": ["柜量40hd", "40hd", "40柜"],
    "container_20hd": ["柜量20hd", "20hd", "20柜"],
    "teu_total": ["折合teu", "teu"],
    "rail_freight": ["铁路运费"],
    "customs_fee": ["报关费"],
    "service_fee": ["服务费"],
    "other_fee": ["其他", "其他费用"],
    "settle_total": ["合计", "结算合计"],
    "actual_freight": ["实付运费", "实付", "运费"],
    "deposit": ["保证金"],
}


def _locate_headers(ws, alias_map: dict, min_matches: int = 2):
    """前6行内定位单行/双行表头，返回 (字段→列, 数据起始行)。"""
    rows = list(ws.iter_rows(min_row=1, max_row=6, values_only=True))
    width = max((len(r) for r in rows), default=0)
    best = None

    def match_cell(text):
        norm = _norm_header(text)
        if not norm:
            return None
        for field, aliases in alias_map.items():
            if norm in aliases:
                return field
        return None

    for i in range(len(rows)):
        single = {}
        for col, cell in enumerate(rows[i]):
            f = match_cell(cell)
            if f:
                single.setdefault(col, f)
        if single:
            count = len(single)
            if count >= min_matches and (best is None or count > best[0]):
                best = (count, single, i + 2)    # 数据从表头下一行
        if i + 1 < len(rows):
            combined = {}
            for col in range(width):
                g, s = rows[i][col], rows[i + 1][col]
                f = (match_cell(str(g or "") + str(s or ""))
                     or match_cell(s) or match_cell(g))
                if f:
                    combined.setdefault(col, f)
            if len(combined) >= min_matches and (
                    best is None or len(combined) > best[0]):
                best = (len(combined), combined, i + 3)
    if best is None:
        raise ValueError(
            "未识别到有效表头（至少需匹配2个已知列名）。请核对列名"
            "（发运日期/班列路线… 或 开行日期/发站/口岸/到站…）。")
    return best[1], best[2]


# ---------------------------------------------------------------- 名称→站编

# 站名清洗（Issue#3 Bug2）：真实表里到站常带括号备注/修饰词
# （如"莫斯科混编（电煤/谢利）"），不清洗会误报未登记
_PAREN_BLOCK_RE = re.compile(r"[（(【][^）)】]*[）)】]")
_STATION_NOISE_RE = re.compile(r"混编|混拼|拼车|加开|重复|二次上报")


def _clean_station_name(text) -> str:
    """站名清洗：去括号及括号内内容、去"混编"类修饰词、去首尾分隔符。
    清洗后再查站编表；仍查不到才归为未登记站名。"""
    t = str(text or "")
    t = _PAREN_BLOCK_RE.sub("", t)
    t = _STATION_NOISE_RE.sub("", t)
    return t.strip(" \u3000—–—-/，,;；、.。\t\r\n")


def _alias_tokens(aliases) -> list[str]:
    if not aliases:
        return []
    return [a.strip() for a in str(aliases).split(",") if a.strip()]


def _row_country(row: dict) -> str:
    """行的国家层级口径：国家层兜底行（RU/ZY）以自身站名作为国家；
    具体站行用 country 字段。"""
    code = row["code"]
    if code in {"RU", "ZY"}:
        return str(row["name"])
    return row.get("country") or ""


def resolve_name(name: str, category: str, code_rows: list[dict]):
    """名称 → (code, country, level) 或 None。

    level: "exact" 命中具体站名/别名；"country" 命中国家层兜底行。
    匹配顺序：编号直配 → 标准站名 → 别名 → 国家层行（仅 dest）。
    """
    text = "".join(str(name or "").split())
    if not text:
        return None
    upper = text.upper()
    for row in code_rows:
        if row["category"] != category:
            continue
        if row["code"].upper() == upper:
            level = "country" if row["code"] in {"RU", "ZY"} else "exact"
            return row["code"], _row_country(row), level
    for row in code_rows:
        if row["category"] != category:
            continue
        if "".join(str(row["name"]).split()) == text:
            level = "country" if row["code"] in {"RU", "ZY"} else "exact"
            return row["code"], _row_country(row), level
    for row in code_rows:
        if row["category"] != category:
            continue
        for alias in _alias_tokens(row.get("aliases")):
            if "".join(alias.split()) == text:
                return row["code"], _row_country(row), "exact"
    return None


# ---------------------------------------------------------------- 行解析

def _dec(value) -> str:
    """单元格 → 去空白字符串（数值在比对时再转Decimal）。"""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def parse_own_table(data: bytes) -> tuple[list[dict], list[dict]]:
    """己方台账 → (归一化行, 未解析名称清单)。

    班列路线为合并文本（"平旺-满洲里-谢利亚季诺"），按分隔符拆三段。
    """
    wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    ws = wb.active
    mapping, data_start = _locate_headers(ws, OWN_HEADER_ALIASES)
    code_rows = train_store.list_codes()

    rows_out, unknown = [], []
    for row_idx, row in enumerate(
            ws.iter_rows(min_row=data_start, values_only=True),
            start=data_start):
        values = {field: (row[col] if col < len(row) else None)
                  for col, field in mapping.items()}
        dep_date, type_hint = parse_dep_cell(values.get("dep_date"))
        row_text = "".join(_dec(v) for v in (row or []))
        # 合计/汇总行（Issue#3 Bug3）：序号列"合计"字样、或无日期且含合计字样/
        # 无路线信息的行——静默跳过，不进配对也不报解析错误
        if dep_date is None:
            if any(k in row_text for k in ("合计", "总计", "小计"))                     or not _dec(values.get("route")):
                continue
            rows_out.append({"row_index": row_idx,
                             "error": "发运日期无法解析（该行跳过）"})
            continue

        route_text = _dec(values.get("route"))
        parts = [p.strip() for p in re.split(r"[-—–_]", route_text) if p.strip()]
        if len(parts) < 3:
            rows_out.append({"row_index": row_idx,
                             "error": f"班列路线「{route_text}」无法拆出"
                                      " 发站-口岸-到站 三段"})
            continue
        station_name, port_name, dest_name = parts[:3]
        # 站名清洗（Issue#3 Bug2）：去括号备注/混编等修饰词后再查站编
        station_clean = _clean_station_name(station_name)
        port_clean = _clean_station_name(port_name)
        dest_clean = _clean_station_name(dest_name)
        unresolved = []
        st_r = resolve_name(station_clean, "station", code_rows)
        po_r = resolve_name(port_clean, "port", code_rows)
        de_r = resolve_name(dest_clean, "dest", code_rows)
        for label, cat, name, res in (
                ("发站", "station", station_name, st_r),
                ("口岸", "port", port_name, po_r),
                ("到站", "dest", dest_name, de_r)):
            if res is None:
                unresolved.append(label)
                unknown.append({"side": "own", "category": cat,
                                "name": _clean_station_name(name),
                                "row_index": row_idx})
        rec = {
            "row_index": row_idx,
            "dep_date": dep_date.isoformat(),
            "train_type": type_hint or train_number.TYPE_SCHEDULED,
            "route_raw": route_text,
            "station_name": station_name, "port_name": port_name,
            "dest_name": dest_name,
            "station_code": st_r[0] if st_r else None,
            "port_code": po_r[0] if po_r else None,
            "dest_code": de_r[0] if de_r else None,
            "dest_country": (de_r[1] if de_r else ""),
            "unresolved": unresolved,
            "pay_date": _dec(values.get("pay_date")),
            "prepay": _dec(values.get("prepay")),
            "line_items": [
                {"label": "运费", "value": _dec(values.get("freight"))},
                {"label": "报关费", "value": _dec(values.get("customs"))},
                {"label": "代理费", "value": _dec(values.get("agency"))},
            ],
            "total": _dec(values.get("total")),
        }
        rows_out.append(rec)
    wb.close()
    return rows_out, unknown


def parse_agent_table(data: bytes) -> tuple[list[dict], list[dict]]:
    """联运对账单 → (归一化行, 未解析名称清单)。到站为独立字段、通常只到国家层。"""
    wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    ws = wb.active
    mapping, data_start = _locate_headers(ws, AGENT_HEADER_ALIASES,
                                          min_matches=3)
    code_rows = train_store.list_codes()

    rows_out, unknown = [], []
    for row_idx, row in enumerate(
            ws.iter_rows(min_row=data_start, values_only=True),
            start=data_start):
        values = {field: (row[col] if col < len(row) else None)
                  for col, field in mapping.items()}
        dep_date, type_hint = parse_dep_cell(values.get("dep_date"))
        row_text = "".join(_dec(v) for v in (row or []))
        # 合计/汇总行（Issue#3 Bug3）：静默跳过
        if dep_date is None:
            if any(k in row_text for k in ("合计", "总计", "小计"))                     or not _dec(values.get("station")):
                continue
            rows_out.append({"row_index": row_idx,
                             "error": "开行日期无法解析（该行跳过）"})
            continue

        station_name, port_name, dest_name = (
            _dec(values.get("station")), _dec(values.get("port")),
            _dec(values.get("dest")))
        station_clean = _clean_station_name(station_name)
        port_clean = _clean_station_name(port_name)
        dest_clean = _clean_station_name(dest_name)
        unresolved = []
        st_r = resolve_name(station_clean, "station", code_rows)
        po_r = resolve_name(port_clean, "port", code_rows)
        de_r = resolve_name(dest_clean, "dest", code_rows)
        for label, cat, name, res in (
                ("发站", "station", station_name, st_r),
                ("口岸", "port", port_name, po_r),
                ("到站", "dest", dest_name, de_r)):
            if res is None:
                unresolved.append(label)
                unknown.append({"side": "agent", "category": cat,
                                "name": _clean_station_name(name),
                                "row_index": row_idx})
        rec = {
            "row_index": row_idx,
            "dep_date": dep_date.isoformat(),
            "train_type": type_hint or train_number.TYPE_SCHEDULED,
            "station_name": station_name, "port_name": port_name,
            "dest_name": dest_name,
            "station_code": st_r[0] if st_r else None,
            "port_code": po_r[0] if po_r else None,
            "dest_code": de_r[0] if de_r else None,
            "dest_country": (de_r[1] if de_r else ""),
            "dest_level": (de_r[2] if de_r else None),
            "unresolved": unresolved,
            "goods_name": _dec(values.get("goods_name")),
            "wagon_count": _dec(values.get("wagon_count")),
            "container_40hd": _dec(values.get("container_40hd")),
            "container_20hd": _dec(values.get("container_20hd")),
            "teu_total": _dec(values.get("teu_total")),
            "line_items": [
                {"label": "铁路运费", "value": _dec(values.get("rail_freight"))},
                {"label": "报关费", "value": _dec(values.get("customs_fee"))},
                {"label": "服务费", "value": _dec(values.get("service_fee"))},
                {"label": "其他", "value": _dec(values.get("other_fee"))},
            ],
            "total": _dec(values.get("settle_total")),
            "actual_freight": _dec(values.get("actual_freight")),
            "deposit": _dec(values.get("deposit")),
        }
        rows_out.append(rec)
    wb.close()
    return rows_out, unknown


# ---------------------------------------------------------------- 配对

def _base_key(rec: dict) -> tuple:
    return (rec["dep_date"], rec["station_code"], rec["port_code"])


def _pair_match(own: dict, agent: dict) -> bool:
    """两记录是否同一趟车：日期+发站+口岸一致，到站按口径匹配：
    - 双方都有具体到站 → 必须相同；
    - 一方仅国家层（或其具体站所属国家与对方国家一致）→ 国家兜底匹配。
    """
    if _base_key(own) != _base_key(agent):
        return False
    od, ad = own.get("dest_code"), agent.get("dest_code")
    if od and ad and od == ad:
        return True
    # 国家兜底：具体站的 country 与 国家层到站 的名称一致
    own_country = own.get("dest_country", "")
    agent_country = agent.get("dest_country", "")
    if own_country and agent_country and own_country == agent_country:
        return True
    return False


def compare_totals(own_total, agent_total) -> str:
    a, b = train_recon.to_decimal(own_total), train_recon.to_decimal(agent_total)
    if a is None or b is None:
        return PENDING if (a or b) else CONFIRMED
    return CONFIRMED if a == b else PENDING


def _category_diffs(own_items, agent_items, own_total, agent_total):
    """科目细项对照（仅展示；不影响状态）。"""
    def nonzero(items):
        out = []
        for i in items:
            dec_v = train_recon.to_decimal(i["value"])
            out.append({"label": i["label"],
                        "value": str(dec_v) if dec_v not in (None, 0) else ""})
        return out
    return {"own": nonzero(own_items), "agent": nonzero(agent_items),
            "own_total": str(train_recon.to_decimal(own_total) or ""),
            "agent_total": str(train_recon.to_decimal(agent_total) or "")}


def build_preview(own_data: bytes, agent_data: bytes) -> dict:
    """双表解析 + 配对 → 预览（只读不写库）。

    返回 pairs（每元素含状态/两侧行/合计差额/科目对照）+ unknown_names +
    stats + message。
    """
    own_rows, own_unknown = parse_own_table(own_data)
    agent_rows, agent_unknown = parse_agent_table(agent_data)

    # 坏行（日期/路线解析失败）直接呈现为错误条目
    pairs: list[dict] = []
    def _badness(row):
        """None=可配对；"unresolved"=站名未登记（去补录组）；"error"=解析失败。"""
        if "error" in row:
            return "error"
        if row.get("unresolved") or not (row.get("station_code")
                                          and row.get("port_code")
                                          and row.get("dest_code")):
            return "unresolved"
        return None

    own_bad = [r for r in own_rows if _badness(r)]
    agent_bad = [r for r in agent_rows if _badness(r)]
    own_ok = [r for r in own_rows if not _badness(r)]
    agent_ok = [r for r in agent_rows if not _badness(r)]

    matched_own, matched_agent = set(), set()
    # 贪心两两配对（含一方国家层兜底）
    for i, own in enumerate(own_ok):
        for j, agent in enumerate(agent_ok):
            if j in matched_agent:
                continue
            if _pair_match(own, agent):
                matched_own.add(i)
                matched_agent.add(j)
                status = compare_totals(own["total"], agent["total"])
                diff = None
                a, b = (train_recon.to_decimal(own["total"]),
                         train_recon.to_decimal(agent["total"]))
                if a is not None and b is not None:
                    diff = b - a
                pairs.append({
                    "status": status,
                    "match_key": "|".join(_base_key(own)
                                          + ((own.get("dest_code") or "?"),)),
                    "own": _slim_row(own), "agent": _slim_row(agent),
                    "own_total": own["total"], "agent_total": agent["total"],
                    "diff": str(diff) if diff is not None else None,
                    "category_breakdown": _category_diffs(
                        own["line_items"], agent["line_items"],
                        own["total"], agent["total"]),
                    "message": (_status_message(status, own, agent)),
                })
                break

    # 未配对 → 缺失状态
    for j, agent in enumerate(agent_ok):
        if j not in matched_agent:
            pairs.append({
                "status": OWN_MISSING,
                "match_key": "|".join(_base_key(agent)
                                      + ((agent.get("dest_code") or "?"),)),
                "own": None, "agent": _slim_row(agent),
                "own_total": "", "agent_total": agent["total"],
                "diff": None, "category_breakdown": None,
                "message": "联运对账单有此记录，己方台账缺失，请核实是否漏记。",
            })
    for i, own in enumerate(own_ok):
        if i not in matched_own:
            pairs.append({
                "status": AGENT_MISSING,
                "match_key": "|".join(_base_key(own)
                                      + ((own.get("dest_code") or "?"),)),
                "own": _slim_row(own), "agent": None,
                "own_total": own["total"], "agent_total": "",
                "diff": None, "category_breakdown": None,
                "message": "己方台账有此记录，联运对账单缺失，请核实对方是否漏开。",
            })
    for bad in own_bad + agent_bad:
        badness = _badness(bad)
        if badness == "unresolved":
            # 未登记站名：只在此处报告一次（Issue#3 Bug2，不再叠加解析错误）
            pairs.append({
                "status": UNRESOLVED_NAME, "match_key": "",
                "own": bad if bad in own_bad else None,
                "agent": bad if bad in agent_bad else None,
                "message": "站名未在站编表登记："
                           + "、".join(bad.get("unresolved") or [])
                           + "——请到代码字典补录后重新预览。"})
        else:
            pairs.append({"status": "error", "match_key": "",
                          "own": bad if bad in own_bad else None,
                          "agent": bad if bad in agent_bad else None,
                          "message": bad.get("error", "行解析失败")})

    stats = {CONFIRMED: 0, PENDING: 0, OWN_MISSING: 0,
             AGENT_MISSING: 0, UNRESOLVED_NAME: 0, "error": 0}
    for p in pairs:
        stats[p["status"]] = stats.get(p["status"], 0) + 1

    unknown = own_unknown + agent_unknown
    message = (
        f"已确认 {stats[CONFIRMED]}，待确认 {stats[PENDING]}，"
        f"己方缺失 {stats[OWN_MISSING]}，联运缺失 {stats[AGENT_MISSING]}，"
        f"站名未登记 {stats[UNRESOLVED_NAME]}，错误 {stats['error']}。")
    if unknown:
        message += f" 另有 {len(unknown)} 个站名未在站编表登记，补录后重新预览。"
    return {"pairs": pairs, "unknown_names": unknown,
            "stats": stats, "message": message}


def _slim_row(rec: dict) -> dict:
    """配对结果里的行视图（去掉内部辅助字段）。"""
    keys = ("row_index", "dep_date", "train_type", "station_name",
            "port_name", "dest_name",
            "station_code", "port_code", "dest_code", "dest_country",
            "unresolved", "goods_name", "wagon_count",
            "container_40hd", "container_20hd", "teu_total",
            "total", "actual_freight", "pay_date", "prepay", "route_raw")
    return {k: rec.get(k) for k in keys if k in rec}


def _status_message(status: str, own: dict, agent: dict) -> str:
    if status == CONFIRMED:
        return (f"两表结算合计一致（{own['total']}），自动通过；"
                "科目细项如有差异仅为记账口径，不算异常。")
    return (f"两表结算合计不一致：己方 {own['total']} vs 联运 {agent['total']}，"
            "请人工核实后选择以哪方数据入库。")


# ---------------------------------------------------------------- 应用入库

def apply_preview(decisions: list[dict], by: str) -> dict:
    """按人工决定把配对结果写入现有表（复用 train_store，不新建体系）。

    每个 decision: {"match_key", "decision": "use_agent"|"use_own"|"skip",
                    "own": {...}|None, "agent": {...}|None}
    口径：
      - 班列主档优先取联运行（含品名/车数/柜量）；联运缺失时仅建最小主档；
      - 结算明细取联运数据（train_settlements 字段即联运账单结构）；
      - 己方预付金额 → train_prepayments；
      - 待确认(pending)必须显式选择 use_agent/use_own，未选按 skip。
    """
    results = []
    counts = {"created": 0, "updated": 0, "skipped": 0, "error": 0}
    for item in decisions:
        decision = item.get("decision", "skip")
        own, agent = item.get("own"), item.get("agent")
        key = item.get("match_key", "")
        if decision == "skip" or (own and agent and not decision):
            counts["skipped"] += 1
            results.append({"match_key": key, "status": "skipped",
                            "message": "按人工选择跳过"})
            continue
        try:
            # 数据源：use_own 且无 agent 时用己方；其余优先 agent（字段最全）
            src = own if (decision == "use_own" and agent is None) else (
                agent or own)
            dep_date = src["dep_date"]
            station = src.get("station_code")
            port = src.get("port_code")
            dest = src.get("dest_code") or "RU"
            # 已有同班列？
            existing = [t for t in train_store.list_trips(
                date_from=dep_date, date_to=dep_date,
                station=station, port=port, dest=dest)]
            basic = {
                "goods_name": src.get("goods_name", ""),
                "wagon_count": _int_or_zero(src.get("wagon_count")),
                "container_40hd": _int_or_zero(src.get("container_40hd")),
                "container_20hd": _int_or_zero(src.get("container_20hd")),
                "teu_total": (train_recon.to_decimal(src.get("teu_total"))
                              if src.get("teu_total") else None),
            }
            if not existing:
                trip = train_store.create_trip(
                    dep_date, station, port, dest,
                    train_store.train_number.TYPE_SCHEDULED, by, basic)
                trip_no = trip["trip_no"]
                action = "created"
            else:
                trip_no = min(existing, key=lambda t: t["suffix"])["trip_no"]
                if basic:
                    train_store.update_trip_basic(trip_no, basic, by)
                action = "updated"

            # 结算明细：联运结构 → train_settlements
            settle_fields = {}
            if agent:
                mapping = [("rail_freight", None), ("customs_fee", "customs_fee"),
                           ("service_fee", "service_fee"),
                           ("other_fee", None),
                           ("settle_total", "total"),
                           ("actual_freight", "actual_freight")]
                # agent slim 行仅保留部分键；需要的细项已在 pair 阶段展示，
                # 入库以 total/actual 为主（科目细项差异不作为异常）
                for col, src_key in mapping:
                    val = agent.get(src_key) if src_key else None
                    if val not in (None, ""):
                        settle_fields[col] = val
            elif own:
                # 联运缺失：按己方口径写入合计
                if own.get("total"):
                    settle_fields["settle_total"] = own["total"]
            if settle_fields:
                train_store.upsert_settlement(trip_no, settle_fields, by)

            # 己方预付款（付款日期可能是 2025.7.1 等写法，统一解析）。
            # 金额为 0/空/无法解析时不落库：0 元预付款记录是纯噪音，
            # 还会让台账的"预付核对"横幅误算需补款金额（修复验收发现）。
            prepay_amt = 0.0
            if own and own.get("prepay") is not None:
                try:
                    prepay_amt = float(str(own["prepay"]).replace(",", "").strip()
                                       or 0)
                except ValueError:
                    prepay_amt = 0.0
            if own and prepay_amt > 0:
                from datacheck_import import parse_dep_cell
                pay_date, _ = parse_dep_cell(own.get("pay_date"))
                train_store.add_prepayment(trip_no,
                                          pay_date.isoformat()
                                          if pay_date else "",
                                          own["prepay"], "双表对账导入", by)

            counts[action] += 1
            results.append({"match_key": key, "status": action,
                            "trip_no": trip_no,
                            "message": f"已{ '新建' if action=='created' else '更新' }入库"})
        except Exception as exc:
            counts["error"] += 1
            results.append({"match_key": key, "status": "error",
                            "message": str(exc)})

    if counts["created"] or counts["updated"]:
        train_store.refresh_anomalies(by)
    return {"counts": counts, "results": results,
            "message": (f"入库完成：新建 {counts['created']}，"
                        f"更新 {counts['updated']}，跳过 {counts['skipped']}，"
                        f"错误 {counts['error']}")}


def _int_or_zero(value):
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return None

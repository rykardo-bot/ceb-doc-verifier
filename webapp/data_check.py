# -*- coding: utf-8 -*-
"""数据核对页（财务角色入口）——班列统一编号 + 联运结算/补贴对账（任务书v1）。

数据流：
  Excel/手工录入 → /datacheck/* API（PostgreSQL + 审计留痕）
                → 本页展示（编号查询/台账编辑/三方对账/导入预览确认）
  API 不可用时降级为进程内直连 train_store（本地补记审计，口径一致）。

业务场景（引导式入口：先选"要办哪项业务"，再进各自的线性流程页）：
  🔎 查班列编号    按发运日期/线路查正式编号，一键复制（粘贴回Excel台账用）
  ➕ 登记新班列    正式登记（自动编号+冲突后缀）或预估登记（EST-，可锁定映射）
  📋 班列台账维护  基础信息/结算明细/预付款/补贴测算 编辑（四状态+留痕+差异提示）
  ⚖️ 三方对账      大同供应链 vs 联运公司 vs 上级拨付 + 阈值告警 + 重复值检测
  💰 资金/费用批次 一笔资金/费用登记成批次，分摊关联到多列班列
  🧮 双表对账      我方台账与对方账单自动比对，逐项确认差异后入库
  📥 Excel差异导入 差异预览（新增/一致/冲突）→ 人工逐行/批量决定 覆盖/保留，绝不静默覆盖

范围边界（任务书 §本阶段范围）：不含客户报价、客户预付款、票据流（v2 另行下发）。
"""

from __future__ import annotations

import io
import os
from datetime import date, datetime, timedelta

import pandas as pd
import requests
import streamlit as st

import audit
import datacheck_import
import dual_recon
import fund_store
import train_store
from webapp import session

API_URL = os.environ.get("VERIFY_API_URL", "http://localhost:8000")

RECORD_STATUS_LABELS = {
    "confirmed": "已确认", "pending": "待确认",
    "not_found": "未找到", "business_missing": "业务确认缺失",
}
REVIEW_STATUS_LABELS = {
    "normal": "正常", "suspect": "存疑，需人工复核", "reviewed": "已复核",
}
TRAIN_TYPE_LABELS = {"T": "T=计划内（图定）", "L": "L=临时增开"}

_AMOUNT_FIELDS_SETTLE = [
    ("rail_freight", "铁路运费"), ("customs_fee", "报关费"),
    ("service_fee", "服务费"), ("other_fee", "其他费用"),
    ("settle_total", "结算合计"), ("actual_freight", "实付运费"),
]
_AMOUNT_FIELDS_SUBSIDY = [
    ("dt_supply_100", "大同供应链·补贴资料100%"),
    ("dt_supply_70", "大同供应链·补贴70%"),
    ("ly_advance_100", "联运垫付补贴100%"),
    ("ly_recover_70", "需回款给联运的70%"),
    ("auth_confirm_100", "上级·确认补贴100%"),
    ("auth_advance_70", "上级·预拨付70%"),
    ("auth_remain_30", "上级·剩余30%"),
    ("forecast_diff", "预测补贴差额"),
]


# ================================================================ API 调用封装

def _headers() -> dict:
    try:
        return session.auth_headers()
    except Exception:
        return {}


def _api(method: str, path: str, **kwargs):
    """调 /datacheck API（携带登录令牌）；不可用返回 None（走本地降级）。

    404 静默降级：部署过渡期旧版 API 无 /datacheck 路由（或记录确实不存在），
    本地降级查询的是同一数据库，结果一致——不弹错干扰操作。"""
    try:
        resp = requests.request(method, f"{API_URL}{path}", timeout=20,
                                headers=_headers(), **kwargs)
        if resp.status_code == 401:
            session.logout()
            st.warning("登录状态已过期，请重新登录。")
            st.rerun()
        if resp.status_code == 404:
            return None
        if resp.status_code >= 400:
            try:
                detail = resp.json().get("detail", resp.text)
            except Exception:
                detail = resp.text
            st.error(f"接口返回 {resp.status_code}：{detail}")
            return None
        return resp
    except requests.RequestException:
        return None


def _fallback_note() -> None:
    st.caption("⚠️ 在线核验服务未连接，当前在本机完成核对；操作留痕不受影响。")


def _safe_audit(action, object_type, object_id, **kw) -> None:
    """审计兜底：留痕失败不阻断页面（与单据核对页同口径）。"""
    try:
        audit.record(session.current_username(), action, object_type, object_id, **kw)
    except Exception as exc:
        print(f"[data_check] 审计写入失败（{action}）: {exc}")


def dc_lookup(dep_date="", station="", port="", dest="") -> dict | None:
    r = _api("GET", "/datacheck/lookup", params={
        "dep_date": dep_date, "station": station, "port": port, "dest": dest})
    if r is not None:
        return r.json()
    try:
        result = train_store.lookup(dep_date=dep_date or None,
                                    station=station or None, port=port or None,
                                    dest=dest or None)
        _safe_audit(audit.DC_QUERY_NUMBER, "trip_lookup", dep_date or
                    f"{station}-{port}-{dest}".strip("-"),
                    detail={"trips": len(result["trips"]), "via": "local-fallback"})
        return result
    except Exception as exc:
        st.error(str(exc))
        return None


def dc_codes() -> dict[str, dict[str, str]]:
    r = _api("GET", "/datacheck/codes")
    if r is not None:
        code_map: dict[str, dict[str, str]] = {}
        for row in r.json().get("codes", []):
            if row.get("active", True):
                code_map.setdefault(row["category"], {})[row["code"]] = row["name"]
        return code_map
    try:
        return train_store.load_code_map()
    except Exception:
        return {}


def dc_list_trips(date_from, date_to) -> list | None:
    r = _api("GET", "/datacheck/trips", params={
        "date_from": str(date_from), "date_to": str(date_to), "limit": 500})
    if r is not None:
        return r.json().get("trips", [])
    try:
        return train_store.list_trips(date_from, date_to, limit=500)
    except Exception as exc:
        st.error(str(exc))
        return None


def dc_trip_detail(trip_no: str) -> dict | None:
    r = _api("GET", f"/datacheck/trips/{trip_no}")
    if r is not None:
        return r.json()
    try:
        return train_store.trip_detail(trip_no)
    except Exception as exc:
        st.error(str(exc))
        return None


def dc_create_trip(payload: dict) -> dict | None:
    r = _api("POST", "/datacheck/trips", json=payload)
    if r is not None:
        return r.json()
    try:
        trip = train_store.create_trip(
            payload["dep_date"], payload["station_code"], payload["port_code"],
            payload["dest_code"], payload["train_type"],
            session.current_username(),
            {k: payload.get(k) for k in ("goods_name", "wagon_count",
                                         "container_40hd", "container_20hd",
                                         "teu_total", "route_label", "remark")})
        _safe_audit(audit.DC_CREATE_TRIP, "trip", trip["trip_no"],
                    after={"dep_date": payload["dep_date"],
                           "via": "local-fallback"})
        return trip
    except Exception as exc:
        st.error(str(exc))
        return None


def dc_update(trip_no: str, kind: str, fields: dict) -> dict | None:
    r = _api("PUT", f"/datacheck/trips/{trip_no}/{kind}", json={"fields": fields})
    if r is not None:
        return r.json()
    try:
        fn = (train_store.update_trip_basic if kind == "basic"
              else train_store.upsert_settlement if kind == "settlement"
              else train_store.upsert_subsidy)
        _, after = fn(trip_no, fields, session.current_username())
        _safe_audit(audit.EDIT_FIELD, f"trip_{kind}", trip_no,
                    after={k: str(v) for k, v in after.items()})
        return after
    except Exception as exc:
        st.error(str(exc))
        return None


def dc_add_prepay(trip_no: str, paid_at: str, amount, remark: str) -> bool:
    r = _api("POST", f"/datacheck/trips/{trip_no}/prepays",
             json={"paid_at": paid_at, "amount": amount, "remark": remark})
    if r is not None:
        return True
    try:
        row = train_store.add_prepayment(trip_no, paid_at, amount, remark,
                                         session.current_username())
        _safe_audit(audit.EDIT_FIELD, "prepay", f"{trip_no}/新增预付款", after=row)
        return True
    except Exception as exc:
        st.error(str(exc))
        return False


def dc_delete_prepay(pid: int) -> bool:
    r = _api("DELETE", f"/datacheck/prepays/{pid}")
    if r is not None:
        return True
    try:
        removed = train_store.delete_prepayment(pid)
        _safe_audit(audit.EDIT_FIELD, "prepay", str(pid), before=removed, after=None)
        return True
    except Exception as exc:
        st.error(str(exc))
        return False


def dc_suspect(trip_no: str, review_status: str, note: str) -> bool:
    r = _api("POST", f"/datacheck/trips/{trip_no}/suspect",
             json={"review_status": review_status, "note": note})
    if r is not None:
        return True
    try:
        _, after = train_store.mark_suspect(trip_no, review_status, note,
                                            session.current_username())
        _safe_audit(audit.DC_SUSPECT_MARK, "trip_subsidy", trip_no,
                    after={"review_status": after["review_status"],
                           "suspect_note": after["suspect_note"]})
        return True
    except Exception as exc:
        st.error(str(exc))
        return False


def dc_create_est(payload: dict) -> dict | None:
    r = _api("POST", "/datacheck/est", json=payload)
    if r is not None:
        return r.json()
    try:
        est = train_store.create_est(payload["dep_date"], payload["station_code"],
                                     payload["port_code"], payload["dest_code"],
                                     payload["train_type"],
                                     session.current_username())
        _safe_audit(audit.DC_CREATE_TRIP, "trip_est", est["est_no"],
                    after={"dep_date_est": payload["dep_date"],
                           "via": "local-fallback"})
        return est
    except Exception as exc:
        st.error(str(exc))
        return None


def dc_lock_est(est_no: str, dep_date: str) -> dict | None:
    r = _api("POST", f"/datacheck/est/{est_no}/lock", json={"dep_date": dep_date})
    if r is not None:
        return r.json()
    try:
        result = train_store.lock_est(est_no, dep_date, session.current_username())
        _safe_audit(audit.DC_EST_LOCK, "trip_est", est_no,
                    detail={"official_no": result["official_no"],
                            "relinked_prepays": result["relinked_prepays"],
                            "via": "local-fallback"})
        return result
    except Exception as exc:
        st.error(str(exc))
        return None


def dc_recon(date_from, date_to) -> dict | None:
    r = _api("GET", "/datacheck/recon", params={
        "date_from": str(date_from), "date_to": str(date_to)})
    if r is not None:
        return r.json()
    try:
        rows = train_store.trips_with_subsidies(date_from, date_to)
        thr = train_store.thresholds()
        import train_recon
        for row in rows:
            row["three_way"] = train_recon.three_way_check(row, thr["pct"],
                                                           thr["amount"])
        return {"trips": rows,
                "duplicate_flags": train_recon.duplicate_neighbor_flags(rows),
                "thresholds": {"pct": str(thr["pct"]), "amount": str(thr["amount"])}}
    except Exception as exc:
        st.error(str(exc))
        return None


def dc_import_preview(kind: str, file_bytes: bytes, filename: str) -> dict | None:
    r = _api("POST", "/datacheck/import/preview", params={"kind": kind},
             files={"file": (filename, file_bytes)})
    if r is not None:
        return r.json()
    try:
        result = datacheck_import.detect(kind, file_bytes)
        _safe_audit(audit.DC_IMPORT_PREVIEW, "import", kind,
                    detail={**result["stats"], "via": "local-fallback"})
        return result
    except Exception as exc:
        st.error(str(exc))
        return None


def dc_import_apply(kind: str, decisions: list[dict]) -> dict | None:
    r = _api("POST", "/datacheck/import/apply",
             json={"kind": kind, "decisions": decisions})
    if r is not None:
        return r.json()
    try:
        result = datacheck_import.apply(kind, decisions,
                                        session.current_username())
        _safe_audit(audit.DC_IMPORT_APPLY, "import", kind,
                    detail={**result["counts"], "via": "local-fallback"})
        return result
    except Exception as exc:
        st.error(str(exc))
        return None


def dc_template(kind: str) -> bytes | None:
    r = _api("GET", "/datacheck/import/template", params={"kind": kind})
    if r is not None:
        return r.content
    try:
        return datacheck_import.build_template(kind)
    except Exception as exc:
        st.error(str(exc))
        return None


# ================================================================ v1.1 资金/费用批次

def dc_list_fund_batches(batch_type: str = "") -> list | None:
    r = _api("GET", "/datacheck/fund-batches",
              params={"batch_type": batch_type} if batch_type else None)
    if r is not None:
        return r.json().get("batches", [])
    try:
        return fund_store.list_batches(batch_type or None)
    except Exception as exc:
        st.error(str(exc))
        return None


def dc_create_fund_batch(payload: dict) -> dict | None:
    r = _api("POST", "/datacheck/fund-batches", json=payload)
    if r is not None:
        return r.json()
    try:
        return fund_store.create_batch(
            batch_type=payload["batch_type"],
            total_amount=payload["total_amount"],
            paid_at=payload.get("paid_at") or None,
            fund_purpose=payload.get("fund_purpose", "预付运费"),
            counterparty=payload.get("counterparty", ""),
            cost_category=payload.get("cost_category"),
            trip_nos=[t["trip_no"] for t in payload.get("trips", [])],
            allocations={t["trip_no"]: t.get("allocated_amount")
                         for t in payload.get("trips", [])
                         if t.get("allocated_amount") is not None},
            remark=payload.get("remark", ""),
            by=session.current_username())
    except Exception as exc:
        st.error(str(exc))
        return None


def dc_fund_participates(batch_id: str, participates: bool,
                         reason: str) -> bool:
    r = _api("PUT", f"/datacheck/fund-batches/{batch_id}/participates",
             json={"participates": participates, "reason": reason})
    if r is None:
        try:
            fund_store.set_participates(batch_id, participates, reason,
                                       session.current_username())
        except Exception as exc:
            st.error(str(exc))
            return False
    return True


# ================================================================ v1.1 双表对账

def dc_dual_preview(own_bytes: bytes, own_name: str,
                    agent_bytes: bytes, agent_name: str) -> dict | None:
    try:
        resp = requests.post(
            f"{API_URL}/datacheck/dual-recon/preview",
            files={"own_file": (own_name, own_bytes),
                   "agent_file": (agent_name, agent_bytes)},
            headers=_headers(), timeout=60)
        if resp.status_code == 401:
            session.logout()
            st.warning("登录状态已过期，请重新登录。")
            st.rerun()
        if resp.status_code >= 400:
            try:
                detail = resp.json().get("detail", resp.text)
            except Exception:
                detail = resp.text
            st.error(f"接口返回 {resp.status_code}：{detail}")
            return None
        return resp.json()
    except requests.RequestException:
        try:
            result = dual_recon.build_preview(own_bytes, agent_bytes)
            _safe_audit(audit.DC_IMPORT_PREVIEW, "dual_recon", "preview",
                        detail=result["stats"])
            return result
        except Exception as exc:
            st.error(str(exc))
            return None


def dc_dual_apply(decisions: list[dict]) -> dict | None:
    r = _api("POST", "/datacheck/dual-recon/apply", json={"decisions": decisions})
    if r is not None:
        return r.json()
    try:
        result = dual_recon.apply_preview(decisions, session.current_username())
        _safe_audit(audit.DC_IMPORT_APPLY, "dual_recon", "apply",
                    detail=result["counts"])
        return result
    except Exception as exc:
        st.error(str(exc))
        return None


# ================================================================ 展示辅助

def _money(value) -> str:
    if value in (None, ""):
        return "—"
    try:
        return f"{float(value):,.2f}"
    except (TypeError, ValueError):
        return str(value)


# 三方对账的三个100%口径字段（与 train_recon.three_way_check 的入参一致）
_THREE_WAY_VALUE_FIELDS = ("dt_supply_100", "ly_advance_100", "auth_confirm_100")


def _threeway_status(t: dict) -> str:
    """三方对账状态列：三方金额都已录入才允许"一致/需关注"判定；
    缺任一方=尚未录入完整数据，用中性占位——空数据绝不能显示 ✅
    （空≠核对通过，防"假阳性确认"，与Issue #3合计行误判同类风险；
    Issue #5 Bug3回归覆盖）。"""
    if any(t.get(f) in (None, "") for f in _THREE_WAY_VALUE_FIELDS):
        return "— 尚未录入"
    return "❌ 需人工关注" if (t.get("three_way") or {}).get("flag") else "✅"


def _pair_icon(pair: dict) -> str:
    """差异明细对比图标：两值都有才给 ✅/❌，缺值给中性占位（同上口径）。"""
    if (pair.get("a_value") in (None, "") or pair.get("b_value") in (None, "")):
        return "—"
    return "❌" if pair.get("flag") else "✅"


def _num_input(label: str, key: str, value=None) -> float | None:
    """金额输入：空=未填写（None），不强制必填（任务书：允许字段暂缺）。"""
    text = st.text_input(label, value="" if value in (None, "") else str(value),
                         key=key, placeholder="未填写")
    text = text.strip().replace(",", "")
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        st.warning(f"{label} 不是有效数值，已按未填写处理")
        return None


def _route_selectors(code_map: dict, key_prefix: str, defaults: dict | None = None,
                     all_label: str = "（全部）"):
    """发站/口岸/目的地下拉（读代码字典，线路选择）。

    all_label：哨兵选项文案。查询场景用"（全部）"（默认=不筛选）；
    登记/创建场景应传"（请选择）"（默认=未选择，避免误以为已填好）。
    """
    defaults = defaults or {}
    cols = st.columns(3)
    picks = {}
    for col, (cat, label) in zip(cols, (("station", "发站"), ("port", "口岸"),
                                        ("dest", "目的地"))):
        options = sorted((code_map.get(cat) or {}).items())
        labels = [f"{c}={n}" for c, n in options] + [all_label]
        default_idx = len(labels) - 1
        d = defaults.get(cat)
        if d:
            for i, (c, _n) in enumerate(options):
                if c == d:
                    default_idx = i
        with col:
            picks[cat] = st.selectbox(label, labels,
                                      index=default_idx,
                                      key=f"{key_prefix}_{cat}")
    for cat in ("station", "port", "dest"):
        picked = picks[cat]
        # 哨兵选项（（全部）/（请选择））一律视作未选择
        picks[cat] = "" if picked.startswith("（") else picked.split("=")[0]
    return picks


# ================================================================ 页签实现

# 场景入口（Issue #6：先选业务场景，再进各自的线性流程页，替代页签堆叠）。
# 场景值落在普通session键，切页往返不丢（与单据核对同一机制，Issue #5 口径）。
DC_SCENE_KEY = "ceb_dc_scene"
DC_SCENES = [
    ("query", "🔎", "查班列编号",
     "知道发运日期或线路？查出班列统一编号，一键复制粘贴到 Excel 台账"),
    ("register", "➕", "登记新班列",
     "登记一列新班列：直接生成正式编号，或先生成预估编号、事后再锁定"),
    ("ledger", "📋", "班列台账维护",
     "查看某列班列的费用结算、预付款、补贴测算，补录或修改"),
    ("recon", "⚖️", "三方对账",
     "我方、联运公司、上级拨付三方金额放一起，自动找出对不上的班列"),
    ("fund", "💰", "资金/费用批次",
     "把一笔资金或费用登记成批次，分摊关联到多列班列"),
    ("dual", "🧮", "双表对账",
     "我方台账和对方账单两张表自动比对，逐项确认差异后入库"),
    ("import", "📥", "Excel差异导入",
     "把对方 Excel 里的数据批量导进系统：先预览差异，逐行或批量决定"),
]


def render() -> None:
    """数据核对模块入口（Issue #6 引导式重构）。

    第一层只放"您要办理哪项业务？"场景卡片；点击进入各自业务页
    （原页签内的业务内容原样保留，仅外层导航改为卡片引导），
    每页左上角可返回入口层。
    """
    scene_renderers = {
        "query": _tab_query, "register": _tab_register, "ledger": _tab_ledger,
        "recon": _tab_recon, "fund": _tab_fund_batches, "dual": _tab_dual_recon,
        "import": _tab_import,
    }
    scene = st.session_state.get(DC_SCENE_KEY)
    if scene not in scene_renderers:
        _render_scene_entry()
        return
    if st.button("← 返回业务入口", key="dc_scene_back"):
        st.session_state.pop(DC_SCENE_KEY, None)
        st.rerun()
    title = next((t for k, _icon, t, _desc in DC_SCENES if k == scene), "")
    st.subheader(f"📊 数据核对 · {title}")
    scene_renderers[scene]()


def _render_scene_entry() -> None:
    """入口层：业务场景卡片（整卡可点，单击直达，文案大白话）。"""
    st.header("📊 数据核对")
    st.caption("班列统一编号 · 联运结算与补贴对账　|　"
               "编号规则：发运日期-发站-口岸-目的地-L/T，同日同线路同类型自动加 -01 后缀")
    st.markdown("#### 请选择要办理的业务")
    st.markdown('<div class="ceb-cards-anchor"></div>', unsafe_allow_html=True)
    for row in range(0, len(DC_SCENES), 2):
        cols = st.columns(2)
        for col, (key, icon, title, desc) in zip(cols, DC_SCENES[row:row + 2]):
            with col:
                if st.button(f"{icon} **{title}**\n\n{desc}",
                             use_container_width=True, key=f"dc_scene_{key}"):
                    st.session_state[DC_SCENE_KEY] = key
                    st.rerun()


# ---------------------------------------------------------------- 🔎 编号查询

def _tab_query() -> None:
    st.markdown("按 **发运日期** 或 **线路（发站+口岸+目的地）** 查询班列统一编号，"
                "可一键复制粘贴到 Excel 台账的“班列编号”列。")
    code_map = dc_codes()
    today = date.today()
    col_a, col_b = st.columns([1, 2])
    with col_a:
        q_date = st.date_input("按发运日期查询（可留空）", value=None,
                               min_value=date(2020, 1, 1),
                               max_value=today + timedelta(days=366),
                               key="q_date")
    with col_b:
        picks = _route_selectors(code_map, "q_route")
    if st.button("🔍 查询编号", type="primary"):
        st.session_state["q_result"] = {
            "params": (q_date.isoformat() if q_date else "",
                       picks["station"], picks["port"], picks["dest"]),
            "payload": dc_lookup(
                dep_date=q_date.isoformat() if q_date else "",
                station=picks["station"], port=picks["port"], dest=picks["dest"]),
        }
    # 查询条件一变，旧结果自动失效隐藏（防止照着旧条件抄错编号）
    cached = st.session_state.get("q_result")
    cur_params = (q_date.isoformat() if q_date else "",
                  picks["station"], picks["port"], picks["dest"])
    if cached and cached.get("params") != cur_params:
        st.caption("查询条件已变化，请重新点击「🔍 查询编号」获取新结果。")
        return
    result = (cached or {}).get("payload")
    if not result:
        return
    trips = result.get("trips", [])
    ests = result.get("ests", [])
    if not trips and ests:
        msg = result.get("message") or "尚无正式编号，仅有预估编号。"
        st.warning(f"⚠️ {msg}（预估编号可用于先登记预付款，发运日期确定后在"
                   f"“班列台账”或“登记班列”页锁定为正式编号）")
    if not trips:
        st.info("没有匹配的正式编号。" if ests else "没有匹配的班列记录。")
    for t in trips:
        no = t["trip_no"]
        c1, c2, c3 = st.columns([3, 1.2, 1.6])
        with c1:
            st.code(no, language=None)      # 自带一键复制按钮（任务书 §四）
        with c2:
            ttype = "临时(L)" if t.get("train_type") == "L" else "图定(T)"
            st.markdown(f"**{ttype}**")
        with c3:
            st.caption(f"{t.get('route_label') or ''} "
                       f"{t.get('goods_name') or ''}".strip() or "—")
        _safe_audit(audit.DC_COPY_NUMBER, "trip", no)   # 展示即视为可复制，尽力留痕
    locked = [e for e in ests if e.get("official_no")]
    if locked:
        st.markdown("#### 相关预估编号沿革")
        for e in locked:
            st.markdown(f"- `{e['est_no']}` → 已锁定为 "
                        f"`{e['official_no']}`")


# ---------------------------------------------------------------- 📋 班列台账

def _tab_ledger() -> None:
    today = date.today()
    dr = st.date_input("发运日期范围",
                       value=(today - timedelta(days=365), today + timedelta(days=92)),
                       min_value=date(2020, 1, 1),
                       max_value=today + timedelta(days=366),
                       key="ledger_range")
    date_from, date_to = (dr if isinstance(dr, tuple) and len(dr) == 2
                          else (today - timedelta(days=365), today))
    if st.button("📋 刷新台账"):
        st.session_state.pop("ledger_trips", None)
    if "ledger_trips" not in st.session_state:
        st.session_state["ledger_trips"] = dc_list_trips(date_from, date_to) or []
    trips = st.session_state["ledger_trips"]
    if not trips:
        st.info("该日期范围内暂无班列记录。可到“登记班列”页新增，"
                "或通过“Excel导入”批量导入。")
        return
    options = [f"{t['trip_no']}（{t['dep_date']} {'临时L' if t['train_type']=='L' else '图定T'}）"
               for t in trips]
    pick = st.selectbox("选择班列", options, key="ledger_pick")
    trip_no = trips[options.index(pick)]["trip_no"]
    detail = dc_trip_detail(trip_no)
    if not detail:
        st.error("班列详情获取失败。")
        return
    if not _api("GET", "/health"):
        _fallback_note()
    trip = detail["trip"]
    checks = detail["checks"]

    st.markdown(f"### `{trip_no}`")
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("车数", trip.get("wagon_count") if trip.get("wagon_count") is not None else "—")
    with col2:
        st.metric("柜量 40HD/20HD",
                  f"{trip.get('container_40hd') or 0} / {trip.get('container_20hd') or 0}")
    with col3:
        st.metric("折合TEU", trip.get("teu_total") if trip.get("teu_total") is not None else "—")
    with col4:
        status = (detail.get("settlement") or {}).get("record_status", "pending")
        st.metric("结算状态", RECORD_STATUS_LABELS.get(status, status))

    # 四个编辑区默认全部收起（Issue #6：辅助内容默认折叠，点标题才展开，
    # 避免一进来就是一屏表单）
    sec_basic, sec_settle, sec_prepay, sec_subsidy = st.expander(
        "🧾 班列基础信息", expanded=False), st.expander(
        "💰 联运费用结算与实付", expanded=False), st.expander(
        "💳 预付款记录", expanded=False), st.expander(
        "🏛️ 补贴测算与复核", expanded=False)
    with sec_basic:
        _edit_basic(trip)
    with sec_settle:
        _edit_settlement(trip_no, detail)
    with sec_prepay:
        _edit_prepays(detail)
    with sec_subsidy:
        _edit_subsidy(trip_no, detail)

    if detail.get("est_history"):
        st.markdown("#### 📜 编号沿革（预估编号 → 正式编号，留痕）")
        for e in detail["est_history"]:
            st.markdown(f"- `{e['est_no']}` → `{trip_no}`　"
                        f"锁定人 {e.get('locked_by') or '—'}　"
                        f"锁定时间 {str(e.get('locked_at') or '—')[:19]}")


def _edit_basic(trip: dict) -> None:
    with st.form("basic_form", border=True):
        c1, c2, c3 = st.columns(3)
        with c1:
            wagon = st.number_input("车数", min_value=0, step=1,
                                    value=int(trip.get("wagon_count") or 0))
            teu = st.text_input("折合TEU", value=str(trip.get("teu_total") or ""))
        with c2:
            c40 = st.number_input("柜量 40HD", min_value=0, step=1,
                                  value=int(trip.get("container_40hd") or 0))
            goods = st.text_input("货物品名", value=trip.get("goods_name") or "")
        with c3:
            c20 = st.number_input("柜量 20HD", min_value=0, step=1,
                                  value=int(trip.get("container_20hd") or 0))
            route = st.text_input("班列路线（展示用）", value=trip.get("route_label") or "")
        remark = st.text_area("备注", value=trip.get("remark") or "", height=60)
        st.caption("发运日期/发站/口岸/目的地/类型 不可修改——编号即身份；"
                   "登记错误请删除后重新登记。")
        if st.form_submit_button("保存基础信息", type="primary"):
            fields = {
                "wagon_count": wagon or None,
                "container_40hd": c40 or None,
                "container_20hd": c20 or None,
                "teu_total": teu.strip() or None,
                "goods_name": goods, "route_label": route, "remark": remark,
            }
            if dc_update(trip["trip_no"], "basic", fields) is not None:
                st.success("已保存（修改已留痕）")
                st.session_state.pop("ledger_trips", None)


def _edit_settlement(trip_no: str, detail: dict) -> None:
    settle = detail.get("settlement") or {}
    checks = detail["checks"]
    pay = checks["prepay"]
    sa = checks["settle_actual"]
    # 核对提示（任务书 §二：差异原因为空必须提示，不能空着不提示）
    if sa.get("needs_reason"):
        st.warning(f"⚠️ {sa['message']} 请在下方“结算vs实付差异原因”中补充。")
    elif sa.get("has_diff"):
        st.info(f"ℹ️ {sa['message']}")
    st.caption(f"预付核对：{pay['message']}")
    with st.form("settle_form", border=True):
        amounts = {}
        cols = st.columns(3)
        for i, (field, label) in enumerate(_AMOUNT_FIELDS_SETTLE):
            with cols[i % 3]:
                amounts[field] = _num_input(f"{label}（元）", f"st_{field}",
                                            settle.get(field))
        c1, c2 = st.columns(2)
        with c1:
            paid_at = st.date_input("实付时间", value=None, key="st_paid_at",
                                    min_value=date(2020, 1, 1))
        with c2:
            diff_reason = st.text_input("结算vs实付差异原因",
                                        value=settle.get("diff_reason") or "",
                                        placeholder="存在差异时必填（否则系统持续提示）")
        status = st.selectbox("记录状态（四状态口径，允许暂缺不阻断）",
                              list(RECORD_STATUS_LABELS),
                              format_func=RECORD_STATUS_LABELS.get,
                              index=list(RECORD_STATUS_LABELS).index(
                                  settle.get("record_status", "pending")))
        if st.form_submit_button("保存结算明细", type="primary"):
            fields = {k: v for k, v in amounts.items()}
            fields["diff_reason"] = diff_reason
            fields["record_status"] = status
            if paid_at:
                fields["actual_paid_at"] = paid_at.isoformat()
            if dc_update(trip_no, "settlement", fields) is not None:
                st.success("已保存（修改已留痕）；预付/实付核对结论见上方提示。")
                st.session_state.pop("ledger_trips", None)


def _edit_prepays(detail: dict) -> None:
    prepays = detail.get("prepays") or []
    trip_no = detail["trip"]["trip_no"]
    if prepays:
        rows = [{"ID": p["id"], "预付时间": str(p.get("paid_at") or "未填写"),
                 "预付金额(元)": _money(p.get("amount")),
                 "备注": p.get("remark") or "",
                 "原始编号": p.get("est_ref") or ""} for p in prepays]
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
        pid = st.number_input("删除指定预付款（输入ID，0=不删）", min_value=0,
                              step=1, key="prepay_del")
        if pid and st.button("🗑️ 删除该笔预付款"):
            if dc_delete_prepay(int(pid)):
                st.success("已删除（留痕）")
                st.rerun()
    else:
        st.caption("暂无预付款记录。")
    with st.form("prepay_add", border=True):
        c1, c2, c3 = st.columns(3)
        with c1:
            paid_at = st.date_input("预付时间", value=None, key="pp_date",
                                    min_value=date(2020, 1, 1))
        with c2:
            amount = st.text_input("预付金额（元）", key="pp_amount",
                                   placeholder="未填写")
        with c3:
            remark = st.text_input("备注", key="pp_remark")
        submitted = st.form_submit_button("➕ 登记预付款", type="primary")
    if submitted:
        if not amount.strip():
            st.error("预付金额不能为空。")
        elif dc_add_prepay(trip_no, paid_at.isoformat() if paid_at else "",
                           amount.strip(), remark):
            st.success("预付款已登记。")
            st.rerun()


def _edit_subsidy(trip_no: str, detail: dict) -> None:
    subsidy = detail.get("subsidy") or {}
    checks = detail["checks"]
    review = subsidy.get("review_status", "normal")
    if review == "suspect":
        st.error(f"🚩 该补贴记录被标记为 **存疑，需人工复核**——"
                 f"{subsidy.get('suspect_note') or '（未填写存疑原因）'}。"
                 f"原始数值仅作参考，请人工核实后再确认。")
    # 实时重复值提示（任务书 §三：只提示不下结论）
    if detail.get("duplicate_flags"):
        others = {f["other_trip"] for f in detail["duplicate_flags"]}
        fields_txt = "、".join({f["label"] for f in detail["duplicate_flags"]})
        st.warning(f"⚠️ 重复值提示：本记录的 **{fields_txt}** 与相邻记录 "
                   f"{'、'.join(sorted(others))} 数值完全相同——测算表存在复制未更新的可能，"
                   f"请人工复核（系统仅提示，不作结论）。")
    missing = checks.get("missing_subsidy_fields") or []
    if missing:
        st.caption(f"未填写字段：{'、'.join(missing)}（允许暂缺，不强制阻断）")
    tw = checks["three_way"]
    if tw.get("flag"):
        st.error(f"⚖️ 三方对账超阈值：{'；'.join(p['message'] for p in tw['pairs'] if p['flag'])}")
    if subsidy.get("diff_reason"):
        st.info(f"预测差额原因：{subsidy['diff_reason']}")
    elif subsidy.get("forecast_diff") not in (None, ""):
        st.warning("⚠️ 已填写预测补贴差额但未填写原因分析，请补充差异原因后再确认。")

    with st.form("subsidy_form", border=True):
        amounts = {}
        cols = st.columns(2)
        for i, (field, label) in enumerate(_AMOUNT_FIELDS_SUBSIDY):
            with cols[i % 2]:
                amounts[field] = _num_input(f"{label}（元）", f"sub_{field}",
                                            subsidy.get(field))
        diff_reason = st.text_input("预测补贴差额·原因分析",
                                    value=subsidy.get("diff_reason") or "")
        status = st.selectbox("记录状态（四状态口径）", list(RECORD_STATUS_LABELS),
                              format_func=RECORD_STATUS_LABELS.get,
                              index=list(RECORD_STATUS_LABELS).index(
                                  subsidy.get("record_status", "pending")))
        if st.form_submit_button("保存补贴测算", type="primary"):
            fields = dict(amounts)
            fields["diff_reason"] = diff_reason
            fields["record_status"] = status
            if dc_update(trip_no, "subsidy", fields) is not None:
                st.success("已保存（修改已留痕）。")
                st.rerun()

    st.markdown("#### 复核标记")
    c1, c2 = st.columns([2, 1])
    with c1:
        new_review = st.selectbox("复核状态", list(REVIEW_STATUS_LABELS),
                                  format_func=REVIEW_STATUS_LABELS.get,
                                  index=list(REVIEW_STATUS_LABELS).index(review))
        note = st.text_input("复核说明", value=subsidy.get("suspect_note") or "")
    with c2:
        st.write("")
        if st.button("💾 保存复核标记", type="primary"):
            if dc_suspect(trip_no, new_review, note):
                st.success("复核标记已保存（留痕）。")
                st.rerun()


# ---------------------------------------------------------------- ➕ 登记班列

def _tab_register() -> None:
    code_map = dc_codes()
    st.markdown("### 正式登记（发运日期已确定）")
    st.caption("编号 = 发运日期-发站-口岸-目的地-L/T；同日同线路同类型已存在时"
               "自动追加 -01/-02（如 `20251010-PW-MZL-RU-T-01`）。")
    with st.form("reg_form", border=True):
        c1, c2 = st.columns(2)
        with c1:
            dep_date = st.date_input("发运日期 *", value=None, key="reg_date",
                                     min_value=date(2020, 1, 1),
                                     max_value=date.today() + timedelta(days=366))
            ttype = st.radio("车次性质 *", ["T", "L"],
                             format_func=lambda v: TRAIN_TYPE_LABELS[v],
                             horizontal=True)
        with c2:
            picks = _route_selectors(code_map, "reg_route",
                                     all_label="（请选择）")
            goods = st.text_input("货物品名")
        c3, c4, c5 = st.columns(3)
        with c3:
            wagon = st.number_input("车数", min_value=0, step=1, value=0)
        with c4:
            c40 = st.number_input("柜量 40HD", min_value=0, step=1, value=0)
        with c5:
            c20 = st.number_input("柜量 20HD", min_value=0, step=1, value=0)
        if st.form_submit_button("➕ 登记并生成正式编号", type="primary"):
            if dep_date is None:
                st.error("请选择发运日期。")
            else:
                trip = dc_create_trip({
                    "dep_date": dep_date.isoformat(),
                    "station_code": picks["station"], "port_code": picks["port"],
                    "dest_code": picks["dest"], "train_type": ttype,
                    "goods_name": goods, "wagon_count": wagon or None,
                    "container_40hd": c40 or None, "container_20hd": c20 or None,
                    "route_label": "", "remark": "",
                })
                if trip:
                    st.success("登记成功！班列统一编号：")
                    st.code(trip["trip_no"], language=None)
                    st.session_state.pop("ledger_trips", None)
                    if trip.get("matching_est_nos"):
                        st.session_state["reg_matching_ests"] = {
                            "trip_no": trip["trip_no"],
                            "est_nos": trip["matching_est_nos"],
                            "dep_date": dep_date.isoformat()}

    match = st.session_state.get("reg_matching_ests")
    if match:
        st.warning(f"检测到发运要素相同的未锁定预估编号："
                   f"{'、'.join(match['est_nos'])}。若该班列此前用预估编号登记过"
                   f"预付款，请锁定映射，预付款将自动改挂到正式编号。")
        if st.button("🔗 锁定预估编号映射"):
            results = [dc_lock_est(no, match["dep_date"]) for no in match["est_nos"]]
            for res in results:
                if res:
                    st.success(f"`{res['est_no']}` → `{res['official_no']}`，"
                               f"改挂预付款 {res['relinked_prepays']} 笔")
            st.session_state.pop("reg_matching_ests")

    st.divider()
    st.markdown("### 预估登记（发运日期未确定，需先登记预付款等）")
    st.caption("编号格式：`EST-` + 预估日期主干。发运日期确定后在此锁定为正式编号，"
               "映射与预付款改挂自动留痕。")
    with st.form("est_form", border=True):
        c1, c2 = st.columns(2)
        with c1:
            est_date = st.date_input("预估发运日期 *", value=None, key="est_date",
                                     min_value=date(2020, 1, 1),
                                     max_value=date.today() + timedelta(days=366))
            est_type = st.radio("车次性质 *", ["T", "L"],
                                format_func=lambda v: TRAIN_TYPE_LABELS[v],
                                horizontal=True, key="est_type")
        with c2:
            est_picks = _route_selectors(code_map, "est_route",
                                         all_label="（请选择）")
        if st.form_submit_button("🧪 生成预估编号", type="primary"):
            if est_date is None:
                st.error("请选择预估发运日期。")
            else:
                est = dc_create_est({
                    "dep_date": est_date.isoformat(),
                    "station_code": est_picks["station"],
                    "port_code": est_picks["port"], "dest_code": est_picks["dest"],
                    "train_type": est_type})
                if est:
                    st.success("预估编号已生成（可在台账中给该编号登记预付款）：")
                    st.code(est["est_no"], language=None)


# ---------------------------------------------------------------- ⚖️ 三方对账

def _tab_recon() -> None:
    st.markdown("三方对账口径：**大同供应链测算补贴(100%)** vs "
                "**联运公司测算补贴(100%)** vs **上级拨付(确认补贴100%)**；"
                "差异超过阈值（默认 5% 或 5000 元，任一触发，管理员可在配置中调整）"
                "标注“需人工关注”。重复值检测只提示不下结论。")
    today = date.today()
    dr = st.date_input("发运日期范围",
                       value=(today - timedelta(days=365), today + timedelta(days=92)),
                       key="recon_range")
    date_from, date_to = (dr if isinstance(dr, tuple) and len(dr) == 2
                          else (today - timedelta(days=365), today))
    data = dc_recon(date_from, date_to)
    if not data:
        return
    trips = data.get("trips", [])
    if not trips:
        st.info("该日期范围内暂无补贴数据。请先在“班列台账”录入或通过 Excel 导入。")
        return
    dup = data.get("duplicate_flags", {})
    rows = []
    for t in trips:
        review = t.get("review_status") or "normal"
        rows.append({
            "班列编号": t["trip_no"], "发运日期": str(t["dep_date"]),
            "类型": "L" if t["train_type"] == "L" else "T",
            "大同供应链(100%)": _money(t.get("dt_supply_100")),
            "联运公司(100%)": _money(t.get("ly_advance_100")),
            "上级拨付(100%)": _money(t.get("auth_confirm_100")),
            "三方一致": _threeway_status(t),
            "重复值提示": "⚠️ 有" if t["trip_no"] in dup else "—",
            "复核状态": REVIEW_STATUS_LABELS.get(review, review),
        })
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
    st.caption(f"当前阈值：差异比 {data['thresholds']['pct']}% 或 "
               f"{data['thresholds']['amount']} 元（管理员可调）")

    flagged = [t for t in trips if (t.get("three_way") or {}).get("flag")]
    if flagged:
        st.markdown("### ⚖️ 需人工关注的差异明细")
        for t in flagged:
            with st.expander(f"{t['trip_no']}（{t['dep_date']}）"):
                for pair in t["three_way"]["pairs"]:
                    icon = _pair_icon(pair)
                    st.markdown(
                        f"- {icon} {pair['a_label']} {_money(pair['a_value'])} "
                        f"vs {pair['b_label']} {_money(pair['b_value'])}"
                        f"　→ {pair['message']}")

    if dup:
        st.markdown("### ⚠️ 相邻记录重复值检测（简单精确相等，请人工复核）")
        for trip_no, flags in sorted(dup.items()):
            for f in flags:
                st.markdown(f"- `{trip_no}` 的 **{f['label']}** = "
                            f"{_money(f['value'])} 元，与 `{f['other_trip']}`"
                            f"（{f['other_dep_date']}）完全相同")


# ---------------------------------------------------------------- 📥 Excel导入

def _tab_import() -> None:
    st.markdown("从 Excel 批量导入。流程：**下载模板/准备文件 → "
                "上传预览差异 → 逐行选择 覆盖/保留 → 确认应用**。"
                "系统绝不静默覆盖已有记录。")
    kind = st.radio("导入类别", ["trip", "subsidy"],
                    format_func=lambda v: ("班列结算导入（基础信息+联运费用+实付）"
                                           if v == "trip" else "补贴测算导入（三方补贴+差额原因）"),
                    horizontal=True, key="imp_kind")
    tpl = dc_template(kind)
    if tpl:
        st.download_button("⬇️ 下载导入模板（含脱敏示例行）", data=tpl,
                           file_name="班列结算导入模板.xlsx" if kind == "trip"
                           else "补贴测算导入模板.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    up = st.file_uploader("上传 Excel（.xlsx）", type=["xlsx"], key="imp_file")
    st.caption("选择要导入的差异结果 Excel，单个文件最大 200MB。")
    if not up:
        st.session_state.pop("imp_preview", None)
        return
    if st.button("🔎 解析并预览差异", type="primary"):
        file_bytes = up.getvalue()
        st.session_state["imp_preview"] = dc_import_preview(
            kind, file_bytes, up.name)
        st.session_state.pop("imp_decisions", None)
    preview = st.session_state.get("imp_preview")
    if not preview:
        return
    st.info(preview["message"])
    rows = preview.get("rows", [])
    decisions = st.session_state.setdefault("imp_decisions", {})
    action_badge = {"new": "🆕 新增", "same": "✅ 一致", "conflict": "⚔️ 冲突",
                    "error": "⛔ 错误"}
    # 冲突行可逐行展开决定，也可一键批量决定（批量后仍可逐行改）——
    # 避免几十行时逐行展开点选（Issue #6 排查项）
    conflict_rows = [r for r in rows if r["action"] == "conflict"]
    if conflict_rows:
        st.caption(f"共 {len(conflict_rows)} 行与库内冲突，需要决定保留哪一边："
                   "可逐行展开选择，也可一键批量决定。")
        bk1, bk2, _ = st.columns([1.2, 1.4, 3])
        with bk1:
            if st.button("全部保留现有", key="imp_bulk_keep",
                         use_container_width=True):
                for r in conflict_rows:
                    st.session_state[f"imp_dec_{r['row_index']}"] = "保留现有"
                st.rerun()
        with bk2:
            if st.button("全部覆盖为导入值", key="imp_bulk_overwrite",
                         use_container_width=True):
                for r in conflict_rows:
                    st.session_state[f"imp_dec_{r['row_index']}"] = "覆盖为导入值"
                st.rerun()
    for row in rows:
        if row["action"] == "same":
            continue
        idx = row["row_index"]
        with st.expander(f"第{idx}行　{action_badge.get(row['action'], row['action'])}"
                         f"　{row.get('trip_no') or row.get('message','')}"):
            if row["action"] == "error":
                st.error(row.get("message") or "无法处理该行")
                continue
            diffs = row.get("diffs") or {}
            if diffs:
                diff_rows = [{"字段": k,
                              "现有值": v.get("old") or "（空）",
                              "导入值": v.get("new") or "（空）"}
                             for k, v in diffs.items()]
                st.dataframe(pd.DataFrame(diff_rows), use_container_width=True,
                             hide_index=True)
            else:
                st.caption("新记录：将按编号规则自动生成班列编号并写入。")
            if row["action"] == "conflict":
                choice = st.radio("处理方式", ["保留现有", "覆盖为导入值"],
                                  key=f"imp_dec_{idx}", horizontal=True,
                                  index=0)
                decisions[idx] = {"row_index": idx,
                                  "values": row.get("values", {}),
                                  "decision": ("overwrite" if choice == "覆盖为导入值"
                                               else "keep")}
            elif row["action"] == "new":
                decisions[idx] = {"row_index": idx, "values": row.get("values", {}),
                                  "decision": "create"}
    applicable = [d for d in decisions.values()]
    if applicable and st.button(f"✅ 确认应用（{len(applicable)} 行）",
                                type="primary"):
        result = dc_import_apply(kind, applicable)
        if result:
            st.success(result["message"])
            for res in result.get("results", []):
                icon = {"created": "🆕", "updated": "♻️", "kept": "⏸️",
                        "error": "⛔"}.get(res["status"], "•")
                st.markdown(f"- {icon} 第{res.get('row_index')}行 "
                            f"`{res.get('trip_no') or '—'}`：{res['message']}")
            st.session_state.pop("imp_preview", None)
            st.session_state.pop("imp_decisions", None)
            st.session_state.pop("ledger_trips", None)


# ---------------------------------------------------------------- 💰 资金/费用批次（v1.1）

_BATCH_VERDICT_LABEL = {
    "covered": "✅ 批次总额一致",
    "shortage": "⚠️ 口径合计多于批次额",
    "under": "⚠️ 口径合计少于批次额",
    "skipped": "➖ 不参与核对",
    "unknown": "❔ 无法核对",
}


def _tab_fund_batches() -> None:
    st.markdown("### 资金/费用批次（一笔钱覆盖多趟车 / 多笔钱结清同一批车）")
    st.caption("批次总金额是唯一权威数字；只需告诉系统「这张付款单/发票覆盖哪几趟车、"
               "总共多少钱」，每趟车分摊金额可以留空，系统按批次整体核对——"
               "不再因某趟车账面记0就误报亏损。")

    c1, c2 = st.columns([1, 3])
    with c1:
        filter_type = st.radio(
            "批次类型", ["", "prepay", "cost"],
            format_func=lambda v: {"": "全部", "prepay": "预付款批次",
                                   "cost": "费用/账单批次"}[v],
            key="fund_filter")
    batches = dc_list_fund_batches(filter_type)
    if batches is None:
        return

    with st.expander("➕ 创建批次", expanded=not batches):
        _fund_create_form()

    st.divider()
    if not batches:
        st.info("尚无批次记录。")
    else:
        rows = []
        for b in batches:
            cov = b.get("coverage", {})
            rows.append({
                "批次号": b["batch_id"],
                "类型": "预付" if b["batch_type"] == "prepay" else "费用",
                "用途": b.get("fund_purpose", ""),
                "打款/开票": str(b.get("paid_at") or "—"),
                "批次总额": _money(b.get("total_amount")),
                "关联车数": len(b.get("trips", [])),
                "口径合计": _money(cov.get("covered_total")),
                "差额": _money(cov.get("diff")) if cov.get("diff") not in (None, 0) else "0.00",
                "核对": _BATCH_VERDICT_LABEL.get(cov.get("verdict"),
                                                 cov.get("verdict", "")),
            })
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

        for b in batches:
            cov = b.get("coverage", {})
            with st.expander(f"📄 {b['batch_id']} · "
                             f"{_BATCH_VERDICT_LABEL.get(cov.get('verdict'), '')}"):
                st.caption(cov.get("message", ""))
                if b.get("counterparty"):
                    st.caption(f"对方主体：{b['counterparty']}"
                               + (f"　费用类目：{b['cost_category']}"
                                  if b.get("cost_category") else ""))
                link_rows = [{
                    "班列编号": t["trip_no"],
                    "分摊金额": _money(t.get("allocated_amount"))
                                 if t.get("allocated_amount") is not None else "（未拆分）",
                } for t in b.get("trips", [])]
                st.dataframe(pd.DataFrame(link_rows),
                             use_container_width=True, hide_index=True)
                if cov.get("missing_trip_amounts"):
                    st.caption("以下关联班列暂无结算口径数据："
                               + "、".join(cov["missing_trip_amounts"]))
                _fund_participates_editor(b)

    st.divider()
    st.markdown("### 🕸️ 孤儿/重复登记检测（安全网）")
    if st.button("运行检测"):
        r = _api("GET", "/datacheck/fund-batches/orphan-check/run")
        result = r.json() if r is not None else fund_store.run_orphan_check()
        if result.get("has_issue"):
            st.warning(result["message"])
            if result.get("unknown_trip_links"):
                for item in result["unknown_trip_links"]:
                    st.markdown(f"- 批次 `{item['batch_id']}` → 不存在的编号 "
                                f"`{item['trip_no']}`")
            for item in result.get("duplicate_risk", []):
                st.markdown(f"- `{item['trip_no']}` 同时在旧表登记预付款，"
                            f"批次：{'、'.join(item['batches'])}")
        else:
            st.success(result["message"])


def _fund_create_form() -> None:
    with st.form("fund_batch_form", border=True):
        c1, c2 = st.columns(2)
        with c1:
            batch_type = st.radio(
                "批次类型 *", ["prepay", "cost"],
                format_func=lambda v: "预付款（我们付给联运）" if v == "prepay"
                else "费用/账单（联运开给我们）",
                horizontal=True, key="fb_type")
            paid_at = st.date_input("打款/开票日期 *", value=date.today(),
                                    key="fb_date")
        with c2:
            # 用途：常规选项直接展示；保证金收进"更多历史科目"折叠区
            purpose = st.selectbox(
                "资金用途 *",
                options=["预付运费", "尾款结算", "补贴回款", "其他"],
                key="fb_purpose")
            show_hist = st.checkbox("更多/历史科目（保证金）", key="fb_hist")
            if show_hist:
                purpose = st.selectbox("历史科目", ["保证金"],
                                        key="fb_purpose_hist")
            counterparty = st.text_input("对方主体（可留空）", key="fb_cp")
        amount_text = st.text_input("批次总金额（元）*", key="fb_amount",
                                    placeholder="唯一权威金额")
        cost_category = None
        if batch_type == "cost":
            cost_category = st.selectbox(
                "费用类目 *", ["铁路运费", "报关费", "服务费", "其他"],
                key="fb_cost_cat")
        trips_text = st.text_area(
            "关联班列编号 *（每行一个；正式编号或EST-编号均可）",
            key="fb_trips", height=90,
            placeholder="20250808-DT-EL-RU-T\n20250809-ZD-HGS-ZY-T")
        alloc_text = st.text_area(
            "分摊金额（可选；格式：班列编号=金额，每行一条；留空即不拆分）",
            key="fb_allocs", height=70,
            placeholder="20250808-DT-EL-RU-T=2581233.30")
        remark = st.text_input("备注（用途选「其他」时必填具体说明）", key="fb_remark")
        submitted = st.form_submit_button("创建批次", type="primary")

    if not submitted:
        return
    errors = []
    amount = amount_text.strip().replace(",", "")
    if not amount:
        errors.append("请填写批次总金额")
    trip_nos = [t.strip() for t in trips_text.splitlines() if t.strip()]
    if not trip_nos:
        errors.append("请至少填写1趟关联班列编号")
    allocations = {}
    for line in alloc_text.splitlines():
        if "=" not in line:
            continue
        no, val = line.split("=", 1)
        allocations[no.strip()] = val.strip()
    if purpose == "其他" and not remark.strip():
        errors.append("用途为「其他」时必须在备注填写具体说明")
    if errors:
        for e in errors:
            st.error(e)
        return
    payload = {
        "batch_type": batch_type,
        "total_amount": amount,
        "paid_at": paid_at.isoformat(),
        "fund_purpose": purpose,
        "counterparty": counterparty,
        "cost_category": cost_category,
        "trips": [{"trip_no": no,
                   "allocated_amount": allocations.get(no)}
                  for no in trip_nos],
        "remark": remark,
    }
    result = dc_create_fund_batch(payload)
    if result:
        st.success(f"批次已创建：{result['batch_id']}")
        st.rerun()


def _fund_participates_editor(batch: dict) -> None:
    """参与核对标志的人工改写（必须原因+留痕）。"""
    current = batch.get("participates_in_freight_recon", True)
    with st.expander("⚙️ 高级：人工改写「参与运费核对」"):
        new_flag = st.checkbox(
            "该批次参与运费覆盖核对", value=bool(current),
            key=f"fb_rec::{batch['batch_id']}")
        if new_flag != bool(current):
            reason = st.text_input(
                "改写原因（必填，系统留痕）",
                key=f"fb_rec_reason::{batch['batch_id']}")
            if st.button("保存改写", key=f"fb_rec_save::{batch['batch_id']}"):
                if dc_fund_participates(batch["batch_id"], new_flag, reason):
                    st.success("已保存（EDIT_FIELD 留痕）。")
                    st.rerun()
        elif batch.get("recon_override_reason"):
            st.caption(f"历史改写原因：{batch['recon_override_reason']}")


# ---------------------------------------------------------------- 🧮 双表对账（Issue #2）

_DUAL_STATUS_LABEL = {
    "confirmed": "✅ 已确认（合计一致）",
    "pending": "⚠️ 待确认（合计不一致）",
    "own_missing": "🔴 己方缺失",
    "agent_missing": "🔵 联运缺失",
    "error": "⛔ 解析错误",
}


# ---------------------------------------------------------------- 🧮 双表对账（Issue #2/#3：任务清单式呈现）

_DUAL_STATUS_LABEL = {
    "confirmed": "✅ 已确认（合计一致）",
    "pending": "⚠️ 待确认（合计不一致）",
    "own_missing": "🔴 己方缺失",
    "agent_missing": "🔵 联运缺失",
    "unresolved_name": "❓ 站名未登记（待补录）",
    "error": "⛔ 解析错误",
}


def _train_type_label(row: dict | None) -> str:
    """临/图标识（Issue#3：10-10 两趟是不同车次，不是重复统计）。"""
    if not row:
        return ""
    t = row.get("train_type")
    if t == "L":
        return "（临时L）"
    if t == "T":
        return "（图定T）"
    return ""


def _row_brief(side: str, row: dict | None) -> str:
    if not row:
        return f"{side}：—"
    route = row.get("route_raw") or (
        f"{row.get('station_name','')}-{row.get('port_name','')}"
        f"-{row.get('dest_name','')}")
    return (f"{side} {row.get('dep_date','')}{_train_type_label(row)} "
            f"{route} 合计 {row.get('total') or '—'}")


def _tab_dual_recon() -> None:
    st.markdown("### 双表对账导入（己方台账 × 联运公司对账单）")
    st.caption("判定口径：**只看结算合计是否一致**（科目细项差异如代理费并入"
               "铁路运费，仅展示不报警）。已确认的自动入库；其余按下方任务清单处理。")

    c1, c2 = st.columns(2)
    with c1:
        own_file = st.file_uploader("① 己方台账（大同陆港联运费用明细）",
                                    type=["xlsx"], key="dual_own")
    with c2:
        agent_file = st.file_uploader("② 联运公司对账单",
                                      type=["xlsx"], key="dual_agent")
    st.caption("两份文件均为 .xlsx，单个文件最大 200MB。")
    # 按钮常驻 + 未就绪禁用（P2-5：空态也给出"下一步"，不再是只有一行提示）
    ready = bool(own_file and agent_file)
    if st.button("🔎 解析并配对预览", type="primary", disabled=not ready):
        st.session_state["dual_preview"] = dc_dual_preview(
            own_file.getvalue(), own_file.name,
            agent_file.getvalue(), agent_file.name)
        st.session_state.pop("dual_decisions", None)
    if not ready:
        st.info("请同时上传两份文件（① 己方台账 + ② 联运公司对账单），"
                "然后点击上方「🔎 解析并配对预览」。")
        return

    preview = st.session_state.get("dual_preview")
    if not preview:
        return

    pairs = preview.get("pairs", [])
    stats = preview.get("stats", {})
    unknown = preview.get("unknown_names") or []

    # ---- 汇总一行（任务清单总览）
    st.markdown(
        f'<div style="background:#F9FAFB;border:1px solid #E5E7EB;border-radius:10px;'
        f'padding:8px 14px;font-size:13.5px;">'
        f'✅ 已确认 <b>{stats.get("confirmed", 0)}</b>（无需处理）　'
        f'⚠️ 金额对不上 <b style="color:#8D6E00;">{stats.get("pending", 0)}</b>　'
        f'🔴/🔵 只在一方出现 <b style="color:#B71C1C;">'
        f'{stats.get("own_missing", 0) + stats.get("agent_missing", 0)}</b>　'
        f'❓ 站名认不出 <b style="color:#B26A00;">{stats.get("unresolved_name", 0)}</b>　'
        f'⛔ 解析失败 <b>{stats.get("error", 0)}</b></div>',
        unsafe_allow_html=True)

    decisions = st.session_state.setdefault("dual_decisions", {})
    key_of = lambda p: p.get("match_key") or (
        f"idx{p.get('own', {}) and p['own'].get('row_index')}"
        f"{p.get('agent', {}) and p['agent'].get('row_index')}")

    def _decision_widget(pair: dict, options: list[str], codes: list[str],
                         default_index: int = 0):
        key = key_of(pair)
        choice = st.radio("处理方式", options, key=f"dual_dec_{key}",
                          horizontal=True, index=default_index)
        decisions[key] = {"match_key": key,
                          "decision": codes[options.index(choice)],
                          "own": pair.get("own"), "agent": pair.get("agent")}

    # ---- 任务①：站名认不出来 → 去代码字典补录（含未登记行 + 其解析错误行）
    unknown_pairs = [p for p in pairs if p["status"] == "unresolved_name"]
    if unknown or unknown_pairs:
        with st.expander(f"❶ 这些站名系统认不出来，请去代码字典补录"
                         f"（{max(len(unknown), len(unknown_pairs))} 项）",
                         expanded=True):
            seen = set()
            for u in unknown:
                key = (u["category"], u["name"])
                if key in seen:
                    continue
                seen.add(key)
                cat_label = {"station": "发站", "port": "口岸",
                             "dest": "到站"}.get(u["category"], u["category"])
                st.markdown(f"- [{cat_label}] **{u['name']}**"
                            f"（第{u['row_index']}行）")
            for pair in unknown_pairs:
                st.markdown(f"- {pair['message']}")
            st.caption("补录入口：管理员 → 代码字典（登记站编与别名后回来重新预览）。")

    # ---- 任务②：金额对不上 → 判断谁的对
    pending_pairs = [p for p in pairs if p["status"] == "pending"]
    with st.expander(f"❷ 这些记录金额对不上，请判断谁的对（{len(pending_pairs)} 条）",
                     expanded=bool(pending_pairs)):
        if not pending_pairs:
            st.caption("没有金额不一致的记录。")
        for pair in pending_pairs:
            own, agent = pair.get("own"), pair.get("agent")
            st.markdown(f"**{_row_brief('己方', own)}**")
            st.markdown(f"**{_row_brief('联运', agent)}**")
            st.markdown(f"差额：{pair.get('diff') or '—'} 元")
            cbreak = pair.get("category_breakdown")
            if cbreak:
                c1b, c2b = st.columns(2)
                with c1b:
                    st.markdown("己方科目：" + "；".join(
                        f"{i['label']} {i['value'] or '—'}"
                        for i in cbreak["own"]))
                with c2b:
                    st.markdown("联运科目：" + "；".join(
                        f"{i['label']} {i['value'] or '—'}"
                        for i in cbreak["agent"]))
                st.caption("科目差异仅为记账口径展示，不影响合计核对结论。")
            _decision_widget(pair,
                             ["以联运数据为准", "以己方数据为准", "跳过（暂不处理）"],
                             ["use_agent", "use_own", "skip"])
            st.divider()

    # ---- 任务③：只在一方出现 → 核实是否漏记
    missing_pairs = [p for p in pairs
                     if p["status"] in ("own_missing", "agent_missing")]
    with st.expander(f"❸ 这些记录只在一方出现，请核实是否漏记（{len(missing_pairs)} 条）",
                     expanded=bool(missing_pairs)):
        if not missing_pairs:
            st.caption("没有单侧记录。")
        for pair in missing_pairs:
            if pair["status"] == "own_missing":
                st.markdown(f"🔴 **缺己方台账**：{_row_brief('联运', pair.get('agent'))}")
                options = ["按联运数据补录入库", "跳过（先核实）"]
                codes = ["use_agent", "skip"]
            else:
                st.markdown(f"🔵 **缺联运对账单**：{_row_brief('己方', pair.get('own'))}")
                options = ["按己方数据补录入库", "跳过（先核实）"]
                codes = ["use_own", "skip"]
            _decision_widget(pair, options, codes)

    # ---- 任务④：解析失败（真正的格式问题，合计行已过滤）
    error_pairs = [p for p in pairs if p["status"] == "error"]
    with st.expander(f"❹ 这些记录系统解析失败，请检查原始表格式（{len(error_pairs)} 条）",
                     expanded=bool(error_pairs)):
        if not error_pairs:
            st.caption("没有解析失败的记录。")
        for pair in error_pairs:
            st.markdown(f"- {pair.get('message')}　"
                        f"({_row_brief('行', pair.get('own') or pair.get('agent'))})")

    # ---- 已确认：默认折叠，只看汇总数（带临/图标识防"重复"误解）
    confirmed_pairs = [p for p in pairs if p["status"] == "confirmed"]
    for pair in confirmed_pairs:
        key = key_of(pair)
        decisions[key] = {"match_key": key, "decision": "use_agent",
                          "own": pair.get("own"), "agent": pair.get("agent")}
    with st.expander(f"✅ 已确认 {len(confirmed_pairs)} 条，无需处理"
                     "（点开可查看明细）"):
        for pair in confirmed_pairs:
            st.markdown(f"- {_row_brief('己方', pair.get('own'))}　"
                        f"（联运合计 {pair.get('agent_total') or '—'}）")

    applicable = [d for d in decisions.values() if d.get("decision") != "skip"]
    if st.button(f"✅ 确认入库（含自动通过，共 {len(applicable)} 条）",
                 type="primary"):
        result = dc_dual_apply(list(decisions.values()))
        if result:
            st.success(result["message"])
            for res in result.get("results", []):
                icon = {"created": "🆕", "updated": "♻️", "skipped": "⏸️",
                        "error": "⛔"}.get(res["status"], "•")
                st.markdown(f"- {icon} `{res.get('match_key','')[:40]}`："
                            f"{res['message']}")
            st.session_state.pop("dual_preview", None)
            st.session_state.pop("dual_decisions", None)
            st.session_state.pop("ledger_trips", None)

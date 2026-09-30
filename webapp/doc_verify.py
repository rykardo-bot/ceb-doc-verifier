# -*- coding: utf-8 -*-
"""
单据核对页（原单文件网页前端的展示层，webapp 双入口架构下的功能页）。

数据流：sample_data/*.json / 向导上传 / App批次（PostgreSQL库内记录）
        -> verification_engine.run_verification()（核验规则全部在引擎中，本文件不做规则判断）
        -> 页面展示（汇总卡片 / 明细表格 / AI修正建议 / PDF导出 / 手动编辑实时复核）

架构升级后的差异（相对旧 app.py 单体）：
  - 登录后才能进入（见 webapp/Home.py），核验/查看请求携带登录令牌；
  - 向导上传完成的批次持久化到 PostgreSQL（source=web，记录创建人）；
  - 编辑字段/生成邮件/导出PDF等操作写操作日志（audit_logs，只增不改）。
"""

import hashlib
import io
import json
import llm_layer
import re
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

import audit
import chat_assistant
import doc_contract
import doc_rules
import email_generator
import knowledge_base
import mobile_store
import pdf_ingest
import upload_wizard
from llm_endpoint import resolve_endpoint
from webapp import session, waybill_check
from webapp.format import fmt_dt
from webapp.api_client import (api_healthy, fetch_batch_full,
                               run_verification_effective)

from verification_engine import (
    STATUS_FAIL,
    STATUS_PASS,
    STATUS_WARNING,
    run_verification,
    sort_results_by_severity,
)
from risk_model import GRADE_META

SAMPLE_DIR = Path(__file__).resolve().parent.parent / "sample_data"

BATCH_FILES = [
    ("batch_clean.json", "示例批次A：全部通过（干净数据）"),
    ("batch_with_issues.json", "示例批次B：含4类典型问题"),
    ("batch_route_warning.json", "示例批次C：路线合规告警（SMGS运单+土耳其线路）"),
]

SCENE_SENTENCE = (
    "中欧班列欧洲枢纽换装平均耗时38.7小时，单证不一致是主因之一——"
    "本工具在发运前自动核验单证一致性、齐全性与路线合规性，"
    "几秒内发现人工容易漏掉的单证问题，降低边境滞留与退运风险。"
)

STATUS_META = {
    STATUS_PASS: {"label": "✅ PASS", "bg": "#E8F5E9", "fg": "#1B5E20"},
    STATUS_WARNING: {"label": "⚠️ WARNING", "bg": "#FFF8E1", "fg": "#8D6E00"},
    STATUS_FAIL: {"label": "⛔ FAIL", "bg": "#FFEBEE", "fg": "#B71C1C"},
}

# 风险等级配色（唯一事实来源：评分卡/明细/建议/仪表盘统一引用，P1 配色一致性）
GRADE_COLORS = {
    "low": {"fg": "#1B5E20", "bg": "#E8F5E9", "label": "低风险"},
    "medium": {"fg": "#8D6E00", "bg": "#FFF8E1", "label": "中风险"},
    "high": {"fg": "#B71C1C", "bg": "#FFEBEE", "label": "高风险"},
}

# 四步流程（P1 步骤指引）
FLOW_STEPS = ["选择/上传单证", "核对识别结果", "查看核验报告", "生成整改材料"]

from webapp.styles import APP_CSS  # noqa: E402  (页面共用样式)


def render_step_indicator(done_through: int, current: int) -> None:
    """四步流程指示器（P1）：done_through 及之前的步骤显示为已完成，
    current 步高亮为当前所在，其余待办。"""
    spans = []
    for i, name in enumerate(FLOW_STEPS, start=1):
        if i <= done_through and i != current:
            cls, mark = "done", "✓ "
        elif i == current:
            cls, mark = "cur", f"{i}. "
        else:
            cls, mark = "", f"{i}. "
        spans.append(f'<span class="{cls}">{mark}{name}</span>')
    st.markdown(f'<div class="ceb-step">{"".join(spans)}</div>', unsafe_allow_html=True)

# 手动编辑（P1/F05）暴露的字段：key -> (中文名, 控件类型, 附加参数)
# - 全部字段对每份单证渲染（缺失字段可补齐，不再跳过）；
# - route_countries 用列表序列化（顿号/逗号分隔），修复手机端"列表变字符串"；
# - waybill_type 限定受支持枚举（与 F07 契约一致）。
EDITABLE_FIELDS = [
    ("goods_description", "货物描述", "text", None),
    ("total_packages", "件数（箱数）", "int", None),
    ("gross_weight_kg", "毛重（kg）", "float", 10.0),
    ("net_weight_kg", "净重（kg）", "float", 10.0),
    ("consignor_name", "发货人名称", "text", None),
    ("consignee_name", "收货人名称", "text", None),
    ("total_amount", "发票总金额", "float", 100.0),
    ("declared_value", "报关申报金额", "float", 100.0),
    ("currency", "币种（如USD）", "text", None),
    ("waybill_no", "运单号", "text", None),
    ("waybill_type", "运单类型", "select",
     ["SMGS国际货协运单", "CIM国际铁路运单", "CIM/SMGS统一运单"]),
    ("container_no", "集装箱号", "text", None),
    ("departure_country", "起运国", "text", None),
    ("destination_country", "运抵国", "text", None),
    ("departure_station", "起运站（发站）", "text", None),
    ("destination_station", "目的站（到站）", "text", None),
    ("route_countries", "经停国家（顿号/逗号分隔，保存为列表）", "list", None),
]

_WAYBILL_UNSET = "（不设置/留空）"


@st.cache_data(show_spinner=False)
def load_batch(filename: str) -> dict:
    return json.loads((SAMPLE_DIR / filename).read_text(encoding="utf-8"))


# ---------------------------------------------------------------- UI 组件


def render_risk_dashboard(risk: dict, summary: dict | None = None) -> None:
    """风险评分仪表盘（P0 视觉焦点）：大号环形仪表 + 汇总迷你卡 + 可解释分数构成。
    配色统一引用 GRADE_COLORS。"""
    meta = GRADE_COLORS.get(risk["grade"], GRADE_COLORS["low"])
    color, bg = meta["fg"], meta["bg"]
    score = int(risk["score"])
    deg = max(0, min(100, score)) * 3.6

    mini_cards = ""
    if summary:
        mini_cards = "".join(
            f'<div class="ceb-mini" style="flex:1; min-width:110px; background:{m_bg};'
            f' border:1px solid {m_bd};"><div class="v" style="color:{m_fg};">{m_val}</div>'
            f'<div class="k">{m_label}</div></div>'
            for m_label, m_val, m_fg, m_bg, m_bd in [
                ("总检查项", summary["total"], "#111827", "#F3F4F6", "#E5E7EB"),
                ("通过 PASS", summary["pass"], "#1B5E20", "#E8F5E9", "#A5D6A7"),
                ("警告 WARNING", summary["warning"], "#8D6E00", "#FFF8E1", "#FFE082"),
                ("不合格 FAIL", summary["fail"], "#B71C1C", "#FFEBEE", "#F5B4BD"),
            ])

    chips = "".join(
        f'<div style="display:flex; justify-content:space-between; gap:8px; padding:5px 12px;'
        f' margin:4px 0; border-radius:8px; font-size:13px; background:#F9FAFB; border:1px solid #E5E7EB;">'
        f'<span>{"🔴" if item["status"]=="FAIL" else "🟡"} {item["reason"]}</span>'
        f'<span style="font-weight:800; color:{"#B71C1C" if item["status"]=="FAIL" else "#8D6E00"};">+{item["points"]}</span></div>'
        for item in risk["breakdown"]
    ) or '<div style="font-size:13px; color:#6B7280; padding:4px 0;">无扣分项，各检查全部通过。</div>'

    st.markdown(
        f'<div style="display:flex; gap:22px; margin:6px 0 18px; align-items:stretch; flex-wrap:wrap;">'
        # —— 左：环形仪表（视觉焦点） ——
        f'<div style="flex:0 0 260px; display:flex; justify-content:center; align-items:center;">'
        f'<div style="width:248px; height:248px; border-radius:50%;'
        f' background:conic-gradient({color} {deg}deg, #E9EDF3 {deg}deg);'
        f' display:flex; align-items:center; justify-content:center;'
        f' box-shadow:0 4px 18px {color}2E;">'
        f'<div style="width:192px; height:192px; border-radius:50%; background:#FFFFFF;'
        f' display:flex; flex-direction:column; align-items:center; justify-content:center;">'
        f'<div style="font-size:13px; color:#6B7280; font-weight:600;">单证组风险分</div>'
        f'<div style="font-size:66px; font-weight:800; color:{color}; line-height:1.05;">{score}</div>'
        f'<div style="font-size:14px; font-weight:700; color:{color}; background:{bg};'
        f' border:1px solid {color}44; border-radius:999px; padding:2px 14px; margin-top:6px;">{meta["label"]}</div>'
        f'</div></div></div>'
        # —— 右：汇总迷你卡 + 分数构成 ——
        f'<div style="flex:1; min-width:300px; display:flex; flex-direction:column; gap:10px;">'
        + (f'<div style="display:flex; gap:10px; flex-wrap:wrap;">{mini_cards}</div>' if mini_cards else "")
        + f'<div style="flex:1; border-radius:12px; padding:12px 16px; background:#FFFFFF;'
        f' border:1px solid #E5E7EB;">'
        f'<div style="font-size:13px; color:#6B7280; font-weight:700; margin-bottom:6px;">分数构成（可解释分解）</div>'
        f'{chips}'
        f'<div style="font-size:11px; color:#9CA3AF; margin-top:6px;">风险分级：0-20 低 · 21-50 中 · 51-100 高</div>'
        f'</div></div></div>',
        unsafe_allow_html=True,
    )


def render_summary_cards(summary: dict) -> None:
    """汇总迷你卡（独立渲染形态；主流程已并入仪表盘，保留供其他入口使用）。"""
    cells = "".join(
        f'<div class="ceb-mini" style="flex:1; min-width:120px; background:{bg};'
        f' border:1px solid {bd};"><div class="v" style="color:{fg};">{value}</div>'
        f'<div class="k">{label}</div></div>'
        for label, value, fg, bg, bd in [
            ("总检查项", summary["total"], "#111827", "#F3F4F6", "#E5E7EB"),
            ("通过 PASS", summary["pass"], "#1B5E20", "#E8F5E9", "#A5D6A7"),
            ("警告 WARNING", summary["warning"], "#8D6E00", "#FFF8E1", "#FFE082"),
            ("不合格 FAIL", summary["fail"], "#B71C1C", "#FFEBEE", "#F5B4BD"),
        ]
    )
    st.markdown(
        f'<div style="display:flex; gap:10px; margin:0 0 12px; flex-wrap:wrap;">{cells}</div>',
        unsafe_allow_html=True,
    )


def render_detail_table(results: list) -> None:
    rows = []
    for r in sort_results_by_severity(results):
        rows.append({
            "状态": r["status"],
            "类别": r["category"],
            "检查项": r["check_name"],
            "核验说明": r["detail"] or "—",
            "涉及单证": "、".join(r["involved_docs"]) if r["involved_docs"] else "—",
        })
    df = pd.DataFrame(rows)

    def _highlight(series: pd.Series):
        return [
            f"background-color: {STATUS_META[s]['bg']}; color: {STATUS_META[s]['fg']};"
            "font-weight: 700; text-align: center;"
            for s in series
        ]

    styler = (
        df.style
        .apply(_highlight, subset=["状态"])
        .format({"状态": lambda s: STATUS_META[s]["label"]})
    )
    st.dataframe(
        styler,
        hide_index=True,
        width="stretch",
        column_config={
            "状态": st.column_config.Column(width="small"),
            "类别": st.column_config.Column(width="small"),
            "核验说明": st.column_config.Column(width="large"),
            "涉及单证": st.column_config.Column(width="medium"),
        },
    )


def render_document_groups(verification: dict) -> None:
    """按单据实例分组的分层结果（任务书问题四）：
    中层=每份单据一组（运单#1/运单#2/…），底层=每组内的具体问题
    （哪个字段、什么问题、怎么改）。批次级问题（缺单证等）单独一组。
    完整明细表保留在下方折叠视图，两种视图数据同源（document_groups）。"""
    groups = verification.get("document_groups") or []
    batch_issues = verification.get("batch_level_issues") or []
    if not groups:
        return

    st.markdown("#### 按单据查看问题（每份单据的问题、字段与修改建议）")

    def _issue_card(issue: dict) -> None:
        meta = STATUS_META.get(issue["status"], STATUS_META[STATUS_WARNING])
        field_chip = (f'<span style="font-size:11.5px; color:{meta["fg"]}; background:{meta["bg"]};'
                      f' border:1px solid {meta["fg"]}44; border-radius:999px; padding:1px 8px;'
                      f' margin-left:6px;">字段：{issue["field"]}</span>'
                      if issue.get("field") else "")
        st.markdown(
            f'<div class="ceb-problem" style="background:{meta["bg"]}; border-color:{meta["fg"]};">'
            f'<h4 style="color:{meta["fg"]};">{meta["label"]}　{issue["check_name"]}</h4>'
            f'<div class="d">{issue["detail"]}{field_chip}</div>'
            + (f'<div class="s">💡 <b>怎么改</b>：{issue["suggestion"]}</div>'
               if issue.get("suggestion") else "")
            + "</div>",
            unsafe_allow_html=True)

    for g in groups:
        fail_n, warn_n = g.get("fail_count", 0), g.get("warning_count", 0)
        if fail_n:
            # 注意：st.expander 标题不渲染HTML，徽章用纯文字+状态符号表达
            badge = f"⛔ {fail_n} 项不合格"
        elif warn_n:
            badge = f"⚠️ {warn_n} 项待复核"
        else:
            badge = "✅ 本份无问题"
        with st.expander(
                f"📄 {g['label']}　{g.get('title', '')}　（{g.get('doc_id', '')}）　{badge}",
                expanded=bool(fail_n or warn_n)):
            if not g["issues"]:
                st.success("这份单据未发现问题。")
            for issue in g["issues"]:
                _issue_card(issue)

    if batch_issues:
        with st.expander(f"🧾 批次级问题（整套单证共 {len(batch_issues)} 项，"
                         "如缺单证/结构/构成不符）", expanded=True):
            for issue in batch_issues:
                _issue_card(issue)
    st.caption(f"单据级规范规则版本：{verification.get('doc_rules_version', '—')}"
               f"（规则定义见 doc_rules.yaml，业务可编辑）")

    # 全部通过项折叠收纳（分层视图只展开问题，通过项不干扰定位）
    results = verification.get("results") or []
    passes = [r for r in results if r["status"] == STATUS_PASS]
    if passes:
        with st.expander(f"✅ 全部通过项（{len(passes)} 项）——点击展开查看"):
            rows = [{"检查项": r["check_name"], "类别": r["category"],
                     "核验说明": r["detail"] or "—"} for r in passes]
            st.dataframe(pd.DataFrame(rows), hide_index=True, width="stretch",
                         column_config={
                             "检查项": st.column_config.Column(width="small"),
                             "核验说明": st.column_config.Column(width="large"),
                         })
    # 完整明细表（平铺视图，与分层视图数据同源；供逐项巡检与存档对照）
    with st.expander("📋 完整核验明细表（按检查项平铺，含状态底色与涉及单据）", expanded=False):
        render_detail_table(results)


def render_detail_section(results: list) -> None:
    """核验明细区（旧版扁平视图，保留供调试对照；主流程已改用
    render_document_groups 的按单据分层视图，任务书问题四）。"""
    problems = [r for r in sort_results_by_severity(results) if r["status"] != STATUS_PASS]
    passes = [r for r in results if r["status"] == STATUS_PASS]

    st.markdown("#### 核验明细")
    if problems:
        st.markdown(
            f'<span style="font-size:13px;color:#6B7280;">发现 <b style="color:#B71C1C;">'
            f'{sum(1 for r in problems if r["status"] == STATUS_FAIL)}</b> 项不合格、'
            f'<b style="color:#8D6E00;">{sum(1 for r in problems if r["status"] == STATUS_WARNING)}'
            f'</b> 项警告（已展开）；另有 {len(passes)} 项通过已折叠。</span>',
            unsafe_allow_html=True)
        for r in problems:
            meta = STATUS_META[r["status"]]
            badge = STATUS_META[r["status"]]["label"]
            rule_note = ""
            if r.get("rule_version"):
                rule_note = (f'<div style="font-size:11.5px;color:#9CA3AF;margin-top:6px;">'
                             f'规则版本：{r["rule_version"]}　|　适用范围见说明</div>')
            st.markdown(
                f'<div class="ceb-problem" style="background:{meta["bg"]}; border-color:{meta["fg"]};">'
                f'<h4 style="color:{meta["fg"]};">{badge}　{r["check_name"]}'
                f'<span style="font-weight:400;color:#6B7280;font-size:12px;">　·　{r["category"]}</span></h4>'
                f'<div class="d">{r["detail"]}</div>'
                + (f'<div class="s">💡 <b>建议</b>：{r["suggestion"]}</div>' if r.get("suggestion") else "")
                + rule_note +
                f'</div>',
                unsafe_allow_html=True)
    else:
        st.success(f"本批次 {len(passes)} 项检查全部通过，未发现 FAIL/WARNING 问题。")

    if passes:
        with st.expander(f"✅ 全部通过项（{len(passes)} 项）——点击展开查看"):
            rows = [{"检查项": r["check_name"], "类别": r["category"],
                     "核验说明": r["detail"] or "—"} for r in passes]
            st.dataframe(pd.DataFrame(rows), hide_index=True, width="stretch",
                         column_config={
                             "检查项": st.column_config.Column(width="small"),
                             "核验说明": st.column_config.Column(width="large"),
                         })
    with st.expander("📋 完整核验明细表（含全部状态与涉及单证）", expanded=False):
        render_detail_table(results)


def collect_gray_cases(results: list, documents: list) -> list:
    """收集进入LLM二次判断的灰色地带场景（语义存疑/路线合规WARNING）。"""
    import re as _re
    gray_cases = []
    for r in results:
        if r["status"] != STATUS_WARNING:
            continue
        if r["check_id"] == "CONS-001":
            for sim in r.get("similarities") or []:
                if sim.get("grade") == "suspect":
                    m = _re.search(r"基准「(.+?)」", r["detail"])
                    gray_cases.append(("CONS-001", {
                        "baseline": m.group(1) if m else "",
                        "other": sim.get("text") or "",
                        "score": sim.get("score", 0),
                        "doc_name": "、".join(sim.get("docs") or []),
                    }))
        elif r["check_id"] == "ROUTE-001":
            wb = next((d for d in documents if d.get("doc_type") == "railway_waybill"), None)
            if wb:
                gray_cases.append(("ROUTE-001", {
                    "waybill_type": (wb.get("fields") or {}).get("waybill_type", ""),
                    "route": (wb.get("fields") or {}).get("route_countries") or [],
                }))
    return gray_cases


def render_ai_reasoning(results: list, documents: list) -> None:
    """AI协同推理区块（升级任务书·方向三）：对灰色地带WARNING给出LLM二次判断与解释。"""
    gray_cases = collect_gray_cases(results, documents)
    if not gray_cases:
        return

    st.subheader("🤖 AI 推理说明（规则筛选 → LLM二次判断）")
    for check_id, context in gray_cases:
        opinion = llm_layer.get_ai_opinion(check_id, context)
        if opinion:
            source_badge = {"live": "🟢 实时调用", "preset": "📦 预置推理结果（离线生成）"}.get(
                opinion.get("source"), opinion.get("source", ""))
            st.info(
                f"**场景**：{opinion.get('scenario') or context}\n\n"
                f"**AI判断**：{opinion.get('verdict')}（置信度：{opinion.get('confidence')}）\n\n"
                f"**推理过程**：{opinion.get('explanation')}\n\n"
                f"**建议动作**：{opinion.get('action')}\n\n"
                f"<span style='font-size:11px; color:#6B7280;'>来源：{source_badge} · {opinion.get('llm', '')}"
                f" —— 当前版本为保证稳定性采用预置推理结果，配置 ARK_API_KEY 后未命中的场景将实时调用</span>",
                unsafe_allow_html=True,
            )
        else:
            st.caption(f"⚠️ 场景 {context} 暂无预置推理结果；配置 ARK_API_KEY 环境变量后可实时调用大模型分析。")


def render_kb_basis(results: list) -> None:
    """合规依据展示（RAG-lite）：对 FAIL/WARNING 项做知识库向量检索并引用条文。"""
    problems = [r for r in results if r["status"] != STATUS_PASS]
    if not problems:
        return
    st.subheader("📖 合规依据（知识库向量检索）")
    for r in sort_results_by_severity(problems):
        basis = knowledge_base.basis_for_result(r)
        if not basis:
            continue
        with st.expander(f"依据 · {r['check_name']}（{r['status']}）"):
            for hit in basis:
                e = hit["entry"]
                st.markdown(
                    f"**{e['id']}｜{e['title']}**（相关度 {hit['score']:.2f}）\n\n"
                    f"> {e['text']}")
            st.caption("🔍 " + knowledge_base.KB_DISCLAIMER
                       + " 检索为字符n-gram向量余弦相似度（与语义比对同一套技术）。")


def render_email_generator(verification: dict, documents: list) -> None:
    """AI整改邮件（原生AI功能2）：LLM实时生成，无Key时降级为数据填充的结构化草稿。
    缓存按数据版本（内容哈希）关联——字段编辑/换文件后旧草稿失效，需重新生成（修复 F09）。"""
    st.subheader("📧 整改邮件（AI代拟给供应商/货代）")
    problems = [r for r in verification["results"] if r["status"] != STATUS_PASS]
    if not problems:
        st.success("本批次全部检查通过，无需生成整改邮件。")
        return

    batch_id = str(verification.get("batch_id"))
    dv = doc_contract.data_version(documents)
    cache_key = f"email::{batch_id}::{dv}"
    generated_dv = st.session_state.get(f"email_generated_dv::{batch_id}")
    if generated_dv is not None and generated_dv != dv:
        st.warning("⚠️ 单证数据已变化（字段被编辑或文件已更换），此前生成的邮件基于旧数据，"
                   "请点击下方按钮重新生成。")
    if st.button("✉️ 生成中英双语整改邮件草稿", type="primary"):
        with st.spinner("AI正在起草邮件…"):
            st.session_state[cache_key] = email_generator.generate_email(verification, documents)
            st.session_state[f"email_generated_dv::{batch_id}"] = dv
    result = st.session_state.get(cache_key)
    if not result:
        st.caption("草稿基于本批次真实的 FAIL/WARNING 明细生成；生成后可直接编辑文本，"
                   "下载内容为编辑后的最新文本。")
        return

    if result["mode"] == "not_needed":
        st.success(result["message"])
        return
    mode_badge = {"live": "🟢 AI实时生成",
                  "offline_template": "📦 离线模板模式（已按本批次核验明细自动填充；"
                                      "如需AI润色完整商务邮件，请联系管理员开通AI服务）"}.get(result["mode"], result["mode"])
    st.caption(f"生成方式：{mode_badge}　|　发送前请人工审阅编辑")
    tab_zh, tab_en = st.tabs(["中文版", "English"])
    with tab_zh:
        # 下载读取当前编辑后的实际内容（修复 F09：下载不再使用缓存的原始文本）
        zh_text = st.text_area("邮件草稿（可编辑）", value=result["zh"], height=380,
                               key=f"{cache_key}::zh")
        st.download_button("下载中文版 (.txt)", data=zh_text.encode("utf-8"),
                           file_name=f"整改邮件_中文_{dv}.txt", width="stretch")
    with tab_en:
        en_source = result.get("en") or "（English version unavailable）"
        en_text = st.text_area("Email draft (editable)", value=en_source, height=380,
                               key=f"{cache_key}::en")
        st.download_button("Download English (.txt)", data=en_text.encode("utf-8"),
                           file_name=f"remediation_email_en_{dv}.txt", width="stretch")


def render_chat_assistant(verification: dict, documents: list) -> None:
    """对话式核验助手（原生AI功能1）：LLM+工具调用，数值假设会真实重跑风险模型。
    会话按数据版本关联——字段编辑/换文件后旧对话不带入新数据，并明确提示（修复 F09）。"""
    st.subheader("💬 向AI追问（对话式核验助手）")
    st.caption("ℹ️ 回答由AI实时生成，可能存在误差，请以核验报告明细为准。"
               "数值假设类问题会调用核验引擎真实重算并经计算验证，不是模型猜测。")

    if resolve_endpoint() is None:
        st.info("⚠️ 对话式核验助手需配置 LLM API Key 启用（环境变量 ARK_API_KEY，"
                "或 GLM_API_KEY / BIGMODEL_API_KEY）。为保证诚实性，本功能不提供预置问答。",
                icon="🔑")
        return

    batch_id = str(verification.get("batch_id"))
    dv = doc_contract.data_version(documents)
    hist_key = f"chat::{batch_id}::{dv}"
    prev_dv = st.session_state.get(f"chat_dv::{batch_id}")
    if prev_dv is not None and prev_dv != dv and st.session_state.get(f"chat::{batch_id}::{prev_dv}"):
        st.warning("⚠️ 单证数据已变化（字段被编辑或文件已更换），此前对话基于旧数据；"
                   "以下为新数据的全新会话，重要结论请重新提问核对。")
    st.session_state[f"chat_dv::{batch_id}"] = dv
    history = st.session_state.setdefault(hist_key, [])
    for msg in history:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    question = st.chat_input("例如：为什么这批风险打100分？／如果箱数改成480，风险会变成多少？／一句话总结核心问题")
    if not question:
        return
    history.append({"role": "user", "content": question})
    with st.chat_message("user"):
        st.markdown(question)
    with st.chat_message("assistant"):
        with st.spinner("AI思考中…"):
            reply = chat_assistant.answer_question(
                question, verification, documents, history=history[:-1])
        st.markdown(reply["answer"])
        if reply.get("tool_trace"):
            recalcs = [t for t in reply["tool_trace"] if t["tool"] == "simulate_field_change"]
            if recalcs:
                last = recalcs[-1]["result"]
                st.caption(f"🔧 已调用核验引擎真实重算：新风险分 {last['new_risk_score']}"
                           f"（{last['new_risk_grade']}），FAIL {last['new_summary']['fail']} 项")
        history.append({"role": "assistant", "content": reply["answer"]})


def build_pdf(batch: dict, verification: dict, edited_count: int,
              documents: list | None = None) -> bytes:
    """用 reportlab 生成与页面内容一致的 PDF 报告（中文用内置 STSong-Light 字体）。"""
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    from reportlab.platypus import (Paragraph, SimpleDocTemplate, Spacer, Table,
                                    TableStyle)

    pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
    FONT = "STSong-Light"

    # 数据版本+规则版本入报告（F09/F07：导出材料可追溯到当前数据与规则口径）
    dv = doc_contract.data_version(documents) if documents else "-"
    rule_version = verification.get("rule_version", "")
    # 单据实例标签（任务书问题四：报告能定位到"哪份单据"）
    doc_labels = {g.get("doc_id"): g.get("label") for g in (verification.get("document_groups") or [])}

    def _doc_label(doc_id: str) -> str:
        if doc_id in doc_labels:
            return f"{doc_labels[doc_id]}（{doc_id}）"
        return doc_id

    title_style = ParagraphStyle("t", fontName=FONT, fontSize=16, leading=22, spaceAfter=4)
    normal_style = ParagraphStyle("n", fontName=FONT, fontSize=9.5, leading=14)
    small_style = ParagraphStyle("s", fontName=FONT, fontSize=8, leading=12,
                                 textColor=colors.HexColor("#6B7280"))

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, topMargin=1.5 * cm, bottomMargin=1.5 * cm,
                            leftMargin=1.5 * cm, rightMargin=1.5 * cm,
                            title="中欧班列单证核验报告")
    _declared = verification.get("declared_composition")
    _declared_note = ("　|　申报构成：" + "、".join(f"{k}×{v}" for k, v in _declared.items())
                      if _declared else "")
    story = [
        Paragraph("中欧班列单证智能核验报告",
                  ParagraphStyle("h", parent=title_style, fontSize=17)),
        Paragraph(f"批次：{verification['batch_name']}　|　"
                  f"路径：{batch.get('destination_summary', '—')}　|　"
                  f"生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M')}"
                  + (f"　|　⚠️ 含手动修改字段 {edited_count} 处" if edited_count else "")
                  + f"　|　数据版本 {dv}"
                  + (f"　|　规则版本 {rule_version}" if rule_version else "")
                  + _declared_note,
                  normal_style),
        Spacer(1, 6),
        Paragraph(f"背景：{SCENE_SENTENCE}", small_style),
        Spacer(1, 10),
    ]

    s = verification["summary"]
    risk = verification.get("risk", {})
    risk_line = f"单证组风险分：{risk.get('score', 0)}/100（{risk.get('grade_label', '低风险')}）"
    breakdown = risk.get("breakdown", [])
    if breakdown:
        risk_line += " —— 构成：" + "，".join(
            f"{item['reason']} +{item['points']}" for item in breakdown)
    story.append(Paragraph(
        f"<b>汇总：</b>共 {s['total']} 项检查 —— 通过 {s['pass']}，警告 {s['warning']}，不合格 {s['fail']}",
        ParagraphStyle("sum", parent=normal_style, fontSize=11, leading=16),
    ))
    story.append(Paragraph(
        f"<b>{risk_line}</b>",
        ParagraphStyle("risk", parent=normal_style, fontSize=11, leading=16,
                       textColor=colors.HexColor(GRADE_META.get(risk.get("grade", "low"), {}).get("color", "#111827"))),
    ))
    story.append(Spacer(1, 8))

    header = [Paragraph("状态", ParagraphStyle("h1", parent=normal_style, textColor=colors.white)),
              Paragraph("检查项", ParagraphStyle("h2", parent=normal_style, textColor=colors.white)),
              Paragraph("涉及单据", ParagraphStyle("h2b", parent=normal_style, textColor=colors.white)),
              Paragraph("核验说明", ParagraphStyle("h3", parent=normal_style, textColor=colors.white)),
              Paragraph("AI修正建议", ParagraphStyle("h4", parent=normal_style, textColor=colors.white))]
    data = [header]
    status_fills = []
    # PDF 内置中文字体无 emoji 字形，状态列使用纯文字，颜色由底色表达
    pdf_status_label = {STATUS_PASS: "PASS", STATUS_WARNING: "WARNING", STATUS_FAIL: "FAIL"}
    for i, r in enumerate(sort_results_by_severity(verification["results"]), start=1):
        meta = STATUS_META[r["status"]]
        status_fills.append((i, meta["bg"], meta["fg"]))
        involved = "、".join(_doc_label(d) for d in (r.get("involved_docs") or [])) or "—"
        data.append([
            Paragraph(pdf_status_label[r["status"]], normal_style),
            Paragraph(r["check_name"], normal_style),
            Paragraph(involved, normal_style),
            Paragraph(r["detail"] or "—", normal_style),
            Paragraph(r.get("suggestion") or "—", normal_style),
        ])
    table = Table(data, colWidths=[1.8 * cm, 3.4 * cm, 2.6 * cm, 5.1 * cm, 5.1 * cm],
                  repeatRows=1)
    style_cmds = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#374151")),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#D1D5DB")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]
    for row_idx, bg, _fg in status_fills:
        style_cmds.append(("BACKGROUND", (0, row_idx), (0, row_idx), colors.HexColor(bg)))
    table.setStyle(TableStyle(style_cmds))
    story.append(table)
    story.append(Spacer(1, 10))

    if verification["suggestions"]:
        story.append(Paragraph("AI 修正建议汇总", ParagraphStyle("s2", parent=normal_style,
                                                              fontSize=12, spaceAfter=4)))
        for sug in verification["suggestions"]:
            story.append(Paragraph("• " + sug,
                                   ParagraphStyle("li", parent=normal_style, leftIndent=10,
                                                  spaceAfter=3)))
        story.append(Spacer(1, 8))

    story.append(Paragraph(
        "使用说明：本系统为初级版（内部试用），核验规则为简化规则集（路线规则仅覆盖"
        "中欧班列国际铁路联运场景，版本见报告头部），识别结果可能存在误差，"
        "输出供人工复核参考，不构成自动放行或商业/法律依据。系统正基于一线使用反馈持续迭代。",
        small_style))

    doc.build(story)
    return buf.getvalue()


# ---------------------------------------------------------------- 手动编辑（P1）


def _parse_numeric_text(text: str, kind: str):
    """编辑框文本按字段类型解析（F05；契约实现见 doc_contract.parse_edited_number）。"""
    return doc_contract.parse_edited_number(text, as_int=(kind == "int")), True


def _serialize_edited_value(kind: str, text: str):
    """按字段类型序列化编辑值；返回 None 表示该字段应删除/不设置。"""
    if kind == "list":
        return doc_contract.split_route_text(text)
    return text.strip() or None


def _value_changed(new_value, original) -> bool:
    """编辑改动判定：数值按数值比较（int 12300 与 float 123.0 视为相同，
    避免 number_input 把整数字段误报"已修改"）；其余按内容比较。"""
    def _is_num(v):
        return isinstance(v, (int, float)) and not isinstance(v, bool)
    if _is_num(new_value) and _is_num(original):
        return float(new_value) != float(original)
    return doc_contract.canonical_json(new_value) != doc_contract.canonical_json(original)


def render_doc_field_editor(batch_id: str, doc: dict) -> tuple[dict, int]:
    """渲染单份单据的字段编辑控件（F05契约），返回 (合并后的fields, 修改字段数)。
    供示例批次的整批编辑与向导式上传的逐份确认共用。"""
    original_fields = dict(doc.get("fields") or {})
    fields = dict(original_fields)          # 保留全部原始字段（含非编辑字段）
    edited = 0
    for key, label, kind, extra in EDITABLE_FIELDS:
        widget_key = f"fld::{batch_id}::{doc['doc_id']}::{key}"
        has_original = key in original_fields
        original = original_fields.get(key)
        new_value = None
        included = False

        if kind in ("int", "float"):
            try:
                seed = float(original) if has_original else 0.0
                numeric_seed = seed
            except (TypeError, ValueError):
                numeric_seed = None
            if numeric_seed is None:
                # 原值不是纯数字（如"12300kg"）：退化为文本框，解析交给引擎口径
                text = st.text_input(label, value="" if original is None else str(original),
                                     key=widget_key)
                new_value = doc_contract.parse_edited_number(text, as_int=(kind == "int"))
                included = new_value is not None
            else:
                step = extra or 1.0
                num_fmt = "%.0f" if kind == "int" else None   # 整数字段不显示 480.00
                value = st.number_input(label, value=numeric_seed, format=num_fmt,
                                        min_value=0.0, step=step, key=widget_key)
                new_value = value if kind == "float" else round(float(value), 2)
                if kind == "int" and float(new_value).is_integer():
                    new_value = int(new_value)   # 整数字段保持int类型（不产生12300.0）
                included = has_original or new_value not in (0, 0.0)
        elif kind == "select":
            options = list(extra or [])
            if has_original and str(original) not in options:
                options = [str(original)] + options
            if has_original:
                value = st.selectbox(label, options,
                                     index=options.index(str(original)), key=widget_key)
                new_value, included = value, True
            else:
                value = st.selectbox(label, [_WAYBILL_UNSET] + options,
                                     index=0, key=widget_key)
                new_value = None if value == _WAYBILL_UNSET else value
                included = new_value is not None
        elif kind == "list":
            display = "、".join(str(x) for x in original) if isinstance(original, list) \
                else ("" if original is None else str(original))
            text = st.text_input(label, value=display, key=widget_key)
            new_value = _serialize_edited_value("list", text)
            included = new_value is not None
        else:   # text
            text = st.text_input(label, value="" if original is None else str(original),
                                 key=widget_key)
            new_value = _serialize_edited_value("text", text)
            included = new_value is not None

        if included:
            fields[key] = new_value
        elif has_original:
            fields.pop(key, None)     # 用户清空该字段 → 显式删除
        if _value_changed(fields.get(key), original if has_original else None):
            edited += 1
            _audit_field_edit(batch_id, doc['doc_id'], key,
                              original if has_original else None, fields.get(key))
    return fields, edited


def collect_edited_documents(batch: dict) -> tuple[list, int]:
    """渲染编辑控件并返回合并后的单证列表 + 修改字段数。
    F05：编辑表单来自契约字段定义——缺失字段也渲染控件（可补齐）；
    按字段类型序列化（数值保持数值、路线保持字符串列表）。
    未暴露为可编辑字段的原始字段（发票号/日期等）原样保留，不丢失。"""
    batch_id = batch["batch_id"]
    edited = 0
    documents = []
    for doc in batch["documents"]:
        with st.expander(f"📄 {_doc_title(doc)}（{doc.get('doc_id', '')}）"):
            fields, delta = render_doc_field_editor(batch_id, doc)
            edited += delta
        documents.append({**doc, "fields": fields})
    return documents, edited


def _doc_title(doc: dict) -> str:
    return doc.get("title") or doc.get("doc_type", "单证")


def reset_edits(batch: dict) -> None:
    for prefix in (f"fld::{batch['batch_id']}::", f"fld_audit::{batch['batch_id']}::"):
        for key in list(st.session_state.keys()):
            if key.startswith(prefix):
                del st.session_state[key]
    st.rerun()


# ---------------------------------------------------------------- 页面主体



def _file_sig(up_file) -> str:
    """上传文件签名（修复 F09）：文件名 + 内容SHA256。
    同名同大小但内容不同的文件签名不同，不会误命中旧解析缓存。
    （向导式上传的缓存签名见 upload_wizard._file_sig，口径一致）"""
    data = None
    getter = getattr(up_file, "getvalue", None)
    if callable(getter):
        try:
            data = getter()
        except Exception:
            data = None
    if data is not None:
        return f"{up_file.name}:{hashlib.sha256(data).hexdigest()[:16]}"
    return f"{up_file.name}:{getattr(up_file, 'size', '?')}"


# ---------------------------------------------------------------- 辅助：留痕与入库


def _safe_audit(action, object_type, object_id, detail=None) -> None:
    """操作留痕兜底封装：审计失败不阻断页面主流程（DB异常时页面仍可用）。"""
    try:
        audit.record(session.current_username(), action, object_type, object_id,
                     detail=detail)
    except Exception as exc:  # 审计存储不可用时降级为日志输出
        print(f"[audit] 写入失败（{action}）: {exc}")


def _audit_field_edit(batch_id: str, doc_id: str, field: str, before, after) -> None:
    """编辑字段留痕（同一改动去重）：记录修改前后的值（任务书 §3 审计要求）。"""
    state_key = f"fld_audit::{batch_id}::{doc_id}::{field}"
    after_json = doc_contract.canonical_json(after)
    last_json = st.session_state.get(state_key)
    if last_json == after_json:
        return
    st.session_state[state_key] = after_json
    before_obj = before
    if last_json is not None:
        try:
            before_obj = json.loads(last_json)
        except (TypeError, ValueError):
            before_obj = before
    try:
        audit.record(session.current_username(), audit.EDIT_FIELD, "field",
                     f"{batch_id}/{doc_id}/{field}",
                     before=before_obj, after=after)
    except Exception as exc:
        print(f"[audit] 编辑留痕失败: {exc}")


def _persist_web_batch(batch: dict, verification: dict) -> None:
    """向导完成的批次持久化到 PostgreSQL（source=web + 创建人）并留痕上传；
    同一批次只在首次展示结果时入库一次（幂等标记），重复重跑不产生重复日志。"""
    flag = f"web_batch_saved::{batch.get('batch_id')}"
    if st.session_state.get(flag):
        return
    st.session_state[flag] = True
    documents = batch.get("documents") or []
    try:
        mobile_store.save_batch(documents, verification, source="web",
                                created_by=session.current_username())
        _safe_audit(audit.UPLOAD_DOCS, "batch", batch.get("batch_id", ""),
                    detail={"source": "web", "doc_count": len(documents),
                            "risk_score": verification.get("risk", {}).get("score"),
                            "rule_version": verification.get("rule_version", "")})
    except Exception as exc:
        st.session_state[flag] = False
        st.warning(f"批次入库失败（本次查看不受影响，请检查数据库连接）：{exc}")


def render_recognition_summary(batch: dict, documents: list) -> None:
    """识别状态摘要（P1任务书B2-2/B2-4）：与核验结论分离呈现。
    让用户先看到"系统识别得怎么样、哪些待我处理"，而不是把未识别字段当成业务缺失。"""
    totals = {doc_contract.FIELD_RECOGNIZED: 0, doc_contract.FIELD_NEEDS_REVIEW: 0,
              doc_contract.FIELD_NOT_FOUND: 0, doc_contract.FIELD_BUSINESS_MISSING: 0,
              doc_contract.FIELD_NOT_APPLICABLE: 0}
    open_items = []          # (单据标题, 字段中文名, 状态)
    for i, doc in enumerate(documents, start=1):
        summary = doc_contract.summarize_field_status(doc)
        for k, n in summary["counts"].items():
            totals[k] = totals.get(k, 0) + n
        title = _doc_title(doc) or f"单据#{i}"
        for field in summary["open_fields"]:
            open_items.append((title, doc_contract.FIELD_LABELS_ZH.get(field, field),
                               doc_contract.field_status(doc, field)))
    st.markdown("### 🔎 识别状态（系统是否找到字段）")
    st.markdown(
        f'<div style="background:#F9FAFB;border:1px solid #E5E7EB;border-radius:12px;'
        f'padding:10px 16px;font-size:13.5px;line-height:1.9;">'
        f'✅ 已识别 <b style="color:#1B5E20;">{totals[doc_contract.FIELD_RECOGNIZED]}</b>　'
        f'⚠️ 待人工确认 <b style="color:#8D6E00;">{totals[doc_contract.FIELD_NEEDS_REVIEW]}</b>　'
        f'❔ 未找到候选 <b style="color:#5B6470;">{totals[doc_contract.FIELD_NOT_FOUND]}</b>　'
        f'⛔ 业务确认缺失 <b style="color:#B71C1C;">{totals[doc_contract.FIELD_BUSINESS_MISSING]}</b>　'
        f'➖ 不适用 {totals[doc_contract.FIELD_NOT_APPLICABLE]}</div>', unsafe_allow_html=True)
    if totals[doc_contract.FIELD_NOT_FOUND]:
        st.caption("❔“未找到”是系统没有识别到，不等于单据业务缺失；"
                   "可在字段编辑区补录，或人工确认。")
    if open_items:
        with st.expander(f"待处理字段清单（{len(open_items)} 项）"):
            for doc_title, field_zh, status in open_items[:30]:
                icon = "⚠️" if status == doc_contract.FIELD_NEEDS_REVIEW else (
                    "⛔" if status == doc_contract.FIELD_BUSINESS_MISSING else "❔")
                st.markdown(f"{icon} **{doc_title}** · {field_zh}"
                            f"（{doc_contract.FIELD_STATUS_LABELS.get(status, status)}）")
    st.caption("👇 下方为**核验状态**（字段之间是否矛盾、是否满足业务规则）——"
               "识别状态与核验状态是两件事。")


def render_recent_batches() -> None:
    """库内近期批次一览（PostgreSQL）：网页向导批次与App批次同库可见，
    在"手机拍摄批次"模式输入编号即可查看完整报告。"""
    with st.expander("🗂️ 库内近期批次（含网页上传与App现场上传）", expanded=False):
        try:
            rows = mobile_store.recent(10)
        except Exception as exc:
            st.caption(f"数据库暂不可用：{exc}")
            return
        if not rows:
            st.caption("暂无历史批次。上传核验过的批次会自动保存在数据库中。")
            return
        st.dataframe(
            [{"批次编号": r["batch_id"], "时间": r["created_at"],
              "来源": "网页" if r["source"] == "web" else "App",
              "创建人": r.get("created_by") or "—",
              "风险": r.get("risk_label") or r.get("risk_level"),
              "单证数": r.get("doc_count"),
              "一句话结论": r.get("one_line")} for r in rows],
            use_container_width=True, hide_index=True)
        st.caption("查看完整报告：切换到「📱 手机拍摄批次」输入批次编号，"
                   "或直接访问 网页地址/?mobile_batch=批次编号")


# （编辑留痕挂接点由下方文本替换注入 render_doc_field_editor）


def render() -> None:
    """单据核对模块入口（Issue #6 引导式重构）。

    第一层只放"您要做什么"场景卡片（上传核验 / 运单核对 / 手机批次复查），
    示例批次收在次要入口后面（默认不铺开）；点击卡片进入各自的线性流程页
    （上传 → 核对 → 展示 → 导出），每页左上角可返回入口层。
    核验引擎、上传向导、字段编辑、报告渲染等函数全部原样复用，业务逻辑零改动。
    """
    scene = st.session_state.get(DOC_SCENE_KEY)
    if scene not in (SCENE_UPLOAD, SCENE_WAYBILL, SCENE_MOBILE, SCENE_SAMPLE):
        _render_scene_entry()
        return
    if st.button("← 返回业务入口", key="doc_scene_back"):
        st.session_state.pop(DOC_SCENE_KEY, None)
        st.rerun()
    if scene == SCENE_UPLOAD:
        _render_upload_flow()
    elif scene == SCENE_WAYBILL:
        waybill_check.render()
    elif scene == SCENE_MOBILE:
        _render_mobile_flow()
    else:
        _render_sample_flow()


# 单证来源选项（模块级常量，供 _source_mode_radio 与回归测试共用）
SOURCE_SAMPLE = "📁 示例批次（3组预置模拟数据，一键加载，推荐先看）"
SOURCE_UPLOAD = "📎 上传PDF单证（向导式：声明构成→上传拆分→逐份确认→核验）"
SOURCE_MOBILE = "📱 手机拍摄批次（输入App批次编号，查看现场上传的完整报告）"
SOURCE_MODE_OPTIONS = [SOURCE_SAMPLE, SOURCE_UPLOAD, SOURCE_MOBILE]

# 来源选择的普通session键。Streamlit在切页（组件卸载）时会连同显式key一起
# 清理组件状态（1.64实测，test_issue5_fixes.py 覆盖），普通键不受影响——
# 选择值必须落在普通键，否则"切页→切回"后来源选择被静默重置（Issue #5 Bug1）。
SOURCE_MODE_STATE_KEY = "ceb_source_mode"

# 场景入口（Issue #6：先选场景、再走线性流程，替代页签/单选堆叠）。
# 场景值同样落在普通session键（与来源选择同一机制），切页往返不丢。
DOC_SCENE_KEY = "ceb_doc_scene"
SCENE_UPLOAD = "upload"      # 📎 上传PDF单证核验
SCENE_WAYBILL = "waybill"    # 🚋 运单核对（车底号/铅封号）
SCENE_MOBILE = "mobile"      # 📱 手机拍摄批次复查
SCENE_SAMPLE = "sample"      # 📁 查看示例（辅助内容，默认收起）


def _source_mode_radio() -> str:
    """渲染"选择单证来源"单选，并把选择持久化到普通session键。

    Issue #6 后导航改由场景卡片承担，本函数保留供回归测试与机制兜底
    （切页状态保持的修复口径不变，test_issue5_fixes.py 覆盖）。
    """
    prev = st.session_state.get(SOURCE_MODE_STATE_KEY)
    index = SOURCE_MODE_OPTIONS.index(prev) if prev in SOURCE_MODE_OPTIONS else 0
    chosen = st.radio(
        "第一步 · 选择单证来源",
        SOURCE_MODE_OPTIONS,
        index=index,
        label_visibility="collapsed",
        horizontal=True,
    )
    st.session_state[SOURCE_MODE_STATE_KEY] = chosen
    return chosen


def _mobile_batch_unavailable() -> None:
    """手机批次服务不可用提示（业务口径，不暴露启动命令/技术组件）。"""
    if not api_healthy():
        st.warning("手机批次的完整报告保存在核验服务端，当前网页无法连接核验服务。"
                   "请稍后重试，或联系管理员确认核验服务已启动。")
        st.stop()


def _materials_used() -> bool:
    """本批次是否已生成过整改材料（邮件/AI追问）——用于步骤指示器第4步高亮。"""
    return any(k.startswith(("email_generated_dv::", "chat::")) and v
               for k, v in st.session_state.items())


# ---------------------------------------------------------------- 入口层（场景卡片）

def _render_scene_entry() -> None:
    """入口层：只放"请选择要办理的业务"场景卡片；示例内容收在次要按钮后面。

    卡片顺序按业务优先级：运单核对第一（当前主推），其余依次排列。
    """
    st.markdown(APP_CSS, unsafe_allow_html=True)

    st.title("🔍 单据核对")
    st.markdown("#### 请选择要办理的业务")
    st.markdown('<div class="ceb-cards-anchor"></div>', unsafe_allow_html=True)
    col_left, col_right = st.columns(2)
    with col_left:
        if st.button("🚋 **运单核对（车底号/铅封号）**\n\n"
                     "提交订车系统之后、确认运单之前：以现场装车清单为基准，"
                     "标红不一致的车底号/铅封号",
                     use_container_width=True, key="doc_scene_waybill"):
            st.session_state[DOC_SCENE_KEY] = SCENE_WAYBILL
            st.rerun()
        if st.button("📱 **手机拍摄批次复查**\n\n"
                     "现场已用 App 拍照上传？输入批次编号，"
                     "在电脑端查看完整核验报告",
                     use_container_width=True, key="doc_scene_mobile"):
            st.session_state[DOC_SCENE_KEY] = SCENE_MOBILE
            st.rerun()
    with col_right:
        if st.button("📎 **上传PDF单证核验**\n\n"
                     "把手里的 PDF 单证传上来：声明构成 → 上传拆分 → 逐份确认 → "
                     "AI 核验 → 导出报告与整改材料",
                     use_container_width=True, key="doc_scene_upload"):
            st.session_state[DOC_SCENE_KEY] = SCENE_UPLOAD
            st.rerun()
        with st.container(border=True):
            st.markdown('<div class="ceb-secondary-anchor"></div>',
                        unsafe_allow_html=True)
            if st.button("📁 查看示例批次（3 组演示数据，先看看系统能做什么）",
                         use_container_width=True, key="doc_scene_sample"):
                st.session_state[DOC_SCENE_KEY] = SCENE_SAMPLE
                st.rerun()

    render_recent_batches()
    _render_sidebar_guide()


def _render_sidebar_guide() -> None:
    """侧边栏：使用指引（只描述怎么用，不承担导航）。"""
    with st.sidebar:
        st.subheader("🧭 使用指引")
        st.markdown(
            "**运单核对（车底号/铅封号）**\n\n"
            "上传装车清单 + PDF运单 → 开始核对 → 标红不一致项\n\n"
            "**上传PDF单证核验**\n\n"
            "声明单据构成 → 上传拆分 → 逐份确认 → 核验查看结果 → 导出报告/整改材料\n\n"
            "**手机拍摄批次复查**\n\n"
            "现场App拍照上传 → 记下批次编号 → 输入编号查看完整报告\n\n"
            "**查看示例**\n\n"
            "3组演示数据，先了解系统能做什么")
        st.divider()
        st.caption("📱 Android App：现场拍照即传+速查（电脑端是大脑、手机端是触手），"
                   "完整处理与深度分析在本网页完成；安装包请联系管理员获取。"
                   "网页端建议使用 PC 浏览器。")
        st.divider()
        st.caption("内部系统 · 数据保存于公司内网服务器 · 全程操作留痕")


# ---------------------------------------------------------------- 各场景线性流程

def _render_upload_flow() -> None:
    """📎 上传PDF单证核验：向导（声明→上传→逐份确认）→ 核验结果 → 整改材料。"""
    st.markdown(APP_CSS, unsafe_allow_html=True)
    st.subheader("📎 上传PDF单证核验")
    st.caption("按向导一步步来：声明单据构成 → 上传拆分 → 逐份确认 → 查看核验结果。")

    if not pdf_ingest.ocr_available():
        st.warning("未检测到扫描识别组件（OCR）：扫描件/拍照生成的PDF将无法提取文字，"
                   "文本型PDF不受影响。如需核对扫描件，请联系管理员安装识别组件。")
    result = upload_wizard.render_upload_wizard(render_doc_field_editor)
    if result is None:
        st.stop()
    batch, edited_documents, edited_count = result

    # 核验（规则全部来自 verification_engine，本文件只做展示；优先走API，不可用时直连）
    effective_batch = {**batch, "documents": edited_documents}
    with st.spinner("核验计算中…"):
        verification, verify_mode = run_verification_effective(effective_batch)
    _persist_web_batch(effective_batch, verification)

    _render_report_tail(batch, edited_documents, edited_count, verification, verify_mode)


def _render_mobile_flow() -> None:
    """📱 手机拍摄批次复查：输入批次编号 → 查看完整报告（只读，现场纠正回来源重拍）。"""
    st.markdown(APP_CSS, unsafe_allow_html=True)
    st.subheader("📱 手机拍摄批次复查")

    _mobile_batch_unavailable()
    _default_mobile_id = st.query_params.get("mobile_batch", "")
    mobile_id = st.text_input(
        "输入手机App上传后显示的批次编号（形如 MB-20260920-143001-8A3C）",
        value=_default_mobile_id, key="mobile_batch_id", placeholder="MB-…").strip()
    if not mobile_id:
        st.info("现场用手机App拍照上传后，App只返回一句话摘要；在此输入批次编号即可查看"
                "该批次的完整核验报告（明细、分数构成、AI建议都在电脑端看）。")
        st.stop()
    try:
        _resp = fetch_batch_full(mobile_id)
        if _resp.status_code == 401:
            session.logout()
            st.warning("登录状态已过期，请重新登录。")
            st.rerun()
        if _resp.status_code == 404:
            st.error(f"未找到批次 {mobile_id} 的核验记录——请核对编号后重试；"
                     f"若持续出现，请联系管理员确认该批次已上传成功。")
            st.stop()
        _resp.raise_for_status()
        _mobile_record = _resp.json()
    except Exception as e:
        st.error(f"查询批次失败：{e}")
        st.stop()
    if not _mobile_record.get("verification"):
        st.error("该批次记录缺少完整核验报告（可能由旧版本App上传），无法展示。")
        st.stop()
    _verif = _mobile_record["verification"]
    batch = {
        "batch_id": _mobile_record.get("batch_id", mobile_id),
        "batch_name": _verif.get("batch_name") or f"App现场拍摄 {mobile_id}",
        "destination_summary": "（手机App现场拍摄上传）",
        "documents": _mobile_record.get("documents", []),
    }
    _uploader = _mobile_record.get('created_by') or '—'
    st.success(f"已加载手机批次 {_mobile_record.get('batch_id')} —— "
               f"拍摄于 {fmt_dt(_mobile_record.get('created_at'))}（上传人：{_uploader}），"
               f"风险等级：{_mobile_record.get('risk_label', '—')}")
    st.caption("App上的一句话结论：" + _mobile_record.get("one_line", "—"))

    # 手机批次：报告已在拍摄时核验完成，这里只读展示——现场纠正应回单证来源处
    # 重拍/重传，保持结果可追溯。
    edited_documents = batch["documents"]
    _render_report_tail(batch, edited_documents, 0,
                        _mobile_record["verification"], "api-store")


def _render_sample_flow() -> None:
    """📁 示例批次：选择演示批次 → 字段核对（手动修改实时复核）→ 核验结果 → 整改材料。"""
    st.markdown(APP_CSS, unsafe_allow_html=True)
    st.subheader("📁 示例批次（演示数据）")
    st.caption("演示数据已模拟完成上传与识别提取，可直接体验"
               "“核对字段 → 查看报告 → 生成整改材料”的完整流程。")

    labels = [label for _, label in BATCH_FILES]
    batch_ids = [fn.removesuffix(".json") for fn, _ in BATCH_FILES]
    # 支持 URL 参数直达批次（如 ?batch=batch_with_issues），便于分享与培训
    default_index = (
        batch_ids.index(st.query_params["batch"])
        if "batch" in st.query_params and st.query_params["batch"] in batch_ids
        else 0
    )
    chosen_col, desc_col = st.columns([1, 2])
    with chosen_col:
        chosen_index = labels.index(
            st.selectbox("选择示例批次（模拟上传+OCR提取完成）", labels, index=default_index)
        )
    with desc_col:
        batch = load_batch(BATCH_FILES[chosen_index][0])
        st.markdown(
            f'<div style="border:1px solid #E5E7EB; border-radius:12px; padding:10px 16px;'
            f' background:#F9FAFB; font-size:13px; color:#374151;">'
            f'<b>{batch.get("batch_name", "")}</b>　{batch.get("description", "")}'
            f'<br><span style="color:#6B7280;">🚉 运输路径：{batch.get("destination_summary", "—")}'
            f'　|　📎 已提取单证：'
            f'{"、".join(d.get("title", d.get("doc_type", "")) for d in batch["documents"])}</span></div>',
            unsafe_allow_html=True)
    # 流程进度（示例模式沿用四步指示）
    render_step_indicator(2, 3 if not _materials_used() else 4)

    st.subheader("2️⃣ 单证字段（提取结果，可手动修改实时复核）")
    # 示例模式：整批编辑（向导模式的逐份编辑已在向导内完成并快照）
    edited_documents, edited_count = collect_edited_documents(batch)
    c1, c2, _ = st.columns([1, 2, 3])
    with c1:
        if st.button("↺ 重置本批次修改", disabled=edited_count == 0):
            reset_edits(batch)
    with c2:
        if edited_count:
            st.markdown(f"<span style='color:#B26A00;font-weight:600;'>"
                        f"✍️ 已手动修改 {edited_count} 个字段，以下核验结果已实时更新</span>",
                        unsafe_allow_html=True)

    # 核验（规则全部来自 verification_engine，本文件只做展示；优先走API，不可用时直连）
    effective_batch = {**batch, "documents": edited_documents}
    with st.spinner("核验计算中…"):
        verification, verify_mode = run_verification_effective(effective_batch)

    _render_report_tail(batch, edited_documents, edited_count, verification, verify_mode)


def _render_report_tail(batch: dict, edited_documents: list, edited_count: int,
                        verification: dict, verify_mode: str) -> None:
    """各进入方式共用的收尾段：识别状态 → 核验结果 → 导出 → 整改材料。

    渲染组件与原页面完全一致（风险仪表/分组明细/知识库依据/AI建议/邮件/追问），
    未改动任何判定与展示逻辑。
    """
    summary = verification["summary"]
    st.caption("🔌 核验通道：" + (
        "在线核验服务" if verify_mode == "api"
        else "手机批次（读取服务端保存的完整核验报告）" if verify_mode == "api-store"
        else "本机直连（在线核验服务未连接，核验在本网页完成，结果不受影响）"))

    # 识别状态摘要（与核验结论分离，P1任务书B2-4）
    render_recognition_summary(batch, edited_documents)

    # 核验结果汇总（P0：环形风险仪表为全页视觉焦点；导出按钮放标题行右侧）
    head_left, head_right = st.columns([4, 1])
    with head_left:
        st.subheader("3️⃣ 核验结果汇总")
    with head_right:
        pdf_bytes = build_pdf(batch, verification, edited_count, edited_documents)
        ts = datetime.now().strftime("%Y%m%d_%H%M")
        st.download_button(
            "⬇️ 导出PDF报告",
            data=pdf_bytes,
            file_name=f"核验报告_{batch['batch_id']}_{ts}.pdf",
            mime="application/pdf",
            width="stretch",
            type="primary" if summary["fail"] or summary["warning"] else "secondary",
            on_click=_safe_audit, args=(audit.GENERATE_REPORT_PDF, "batch",
                                        batch["batch_id"], None),
        )
    render_risk_dashboard(verification["risk"], summary)

    # 核验结果（任务书问题四：分层结构——总评分 → 按单据实例分组 → 字段级问题+修改建议）
    render_document_groups(verification)
    render_kb_basis(verification["results"])
    render_ai_reasoning(verification["results"], edited_documents)

    # 第四步：生成整改材料（P1 分组导航）
    st.markdown("### 4️⃣ 生成整改材料")
    render_email_generator(verification, edited_documents)
    # 对话助手收进默认折叠的展开器：chat_input 挂载时会自动聚焦并把页面滚到底部，
    # 折叠后首屏保持在顶部；用户点开时再聚焦正合适。
    with st.expander("💬 向AI追问（对话式核验助手）——点击展开", expanded=False):
        render_chat_assistant(verification, edited_documents)

    # 原始单证数据
    with st.expander("🔍 查看原始单证数据（识别提取结果JSON）"):
        st.json(edited_documents)

    st.divider()
    st.caption(
        "本系统为初级版（内部试用）：核验规则为简化规则集，路线规则仅覆盖中欧班列国际铁路联运"
        "场景（每条路线结果附版本与适用范围），识别结果可能存在误差；结论供人工复核参考，"
        "不构成自动放行依据。欢迎通过一线使用反馈问题与需求，推动系统持续迭代。"
    )

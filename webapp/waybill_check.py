# -*- coding: utf-8 -*-
"""运单核对页签（单据核对模块下的新页签，Issue #4 任务二）。

业务窗口：B同事把车底号录入订车系统并提交之后、确认运单之前。
用法：上传 ① A同事现场装车清单Excel（事实来源）② B同事从订车系统批量导出的
加密PDF运单 → 以箱号为主键逐份核对车底号/铅封号 → 不一致的标红置顶，
供B同事当场修改（错了只能回车站重新盖章，改不成发到边境有退运风险）。

本页签为即传即核的轻流程：不落库、不写操作日志（窗口内核对完即走，
历史留痕仍走单证智能核验主流程）。提取与核对规则全部在 waybill_extract.py，
本文件不做规则判断。
"""

from __future__ import annotations

import pandas as pd
import streamlit as st

import waybill_extract as we
from webapp.styles import APP_CSS

_REPORT_KEY = "waybill_check_report"

_STATUS_LABEL = {
    we.ST_MATCH: "✅ 一致",
    we.ST_WAGON_MISMATCH: "⛔ 车底号不一致",
    we.ST_SEAL_MISMATCH: "⛔ 铅封号不一致",
    we.ST_BOTH_MISMATCH: "⛔ 车底号+铅封号均不一致",
    we.ST_PDF_UNMATCHED: "❓ PDF未匹配到清单",
    we.ST_EXCEL_UNMATCHED: "❓ 清单未匹配到PDF",
    we.ST_EXTRACTION_ANOMALY: "⚠️ 提取异常",
    we.ST_DATA_MISSING: "⚠️ 清单字段为空",
}

_FIELD_LABEL = {"wagon_no": "车底号", "seal_no": "铅封号"}

_ANOMALY_LABEL = {
    we.FIELD_NOT_FOUND: "框内无值",
    we.FIELD_ANCHOR_MISSING: "未找到栏位标签（模板可能变化）",
    we.FIELD_FORMAT_ANOMALY: "格式异常",
    we.FIELD_AMBIGUOUS: "多个候选，需人工确认",
    we.FIELD_READ_ERROR: "文件无法解析",
}


def render() -> None:
    st.markdown(APP_CSS, unsafe_allow_html=True)
    st.subheader("🚋 运单核对（车底号 / 铅封号）")
    st.caption("在「提交订车系统之后、确认运单之前」核对B同事的实际录入："
               "以现场装车清单为基准，按箱号匹配，标红不一致的车底号/铅封号。")

    xlsx_file, pdf_files = _upload_section()
    if xlsx_file and pdf_files and st.button("开始核对", type="primary"):
        with st.spinner("正在解密并解析PDF运单（批量约需几十秒）…"):
            report = _run_check(xlsx_file.getvalue(), pdf_files)
        st.session_state[_REPORT_KEY] = report
        st.rerun()

    if _REPORT_KEY in st.session_state:
        _render_report(st.session_state[_REPORT_KEY])


def _template_xlsx() -> bytes:
    """生成现场装车清单模板（含示例行，表头口径与解析器一致）。"""
    import io
    import openpyxl

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "电子版"
    ws.append(["序号", "车种", "车底号", "箱号", "铅封号"])
    ws.append([1, "X70", "5480001", "TCLU1234567", "260001"])
    ws.append([2, "X70", "5480002", "TCLU7654321", "260002"])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _upload_section():
    """两步上传：清单Excel（基准）+ 批量PDF运单（待核对）。"""
    st.download_button(
        "⬇️ 下载现场装车清单模板（含示例行，按模板列名填写）",
        data=_template_xlsx(),
        file_name="现场装车清单模板.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    col_xlsx, col_pdf = st.columns(2)
    with col_xlsx:
        xlsx_file = st.file_uploader(
            "① A同事现场装车清单（Excel，含车种/车底号/箱号/铅封号）",
            type=["xlsx", "xlsm"])
        if xlsx_file is not None:
            st.caption(f"已选择：{xlsx_file.name}")
    with col_pdf:
        pdf_files = st.file_uploader(
            "② 订车系统导出的PDF运单（可一次选多份，一列车五六十份）",
            type=["pdf"], accept_multiple_files=True)
        if pdf_files:
            st.caption(f"已选择 {len(pdf_files)} 份运单")
    return xlsx_file, pdf_files


def _run_check(xlsx_bytes: bytes, pdf_files: list) -> dict:
    """清单解析 + 逐份提取 + 批量核对（提取用safe包装，坏文件不中断批次）。"""
    try:
        rows = we.load_packing_list(xlsx_bytes)
    except Exception as exc:
        st.error(f"现场清单解析失败：{exc}")
        st.stop()
    if not rows:
        st.error("现场清单中没有解析到有效数据行（需含箱号列）。")
        st.stop()

    extractions = []
    progress = st.progress(0.0, text="正在提取PDF运单…")
    for i, up in enumerate(pdf_files, 1):
        extractions.append(we.extract_waybill_safe(up.getvalue(), up.name))
        progress.progress(i / len(pdf_files),
                          text=f"正在提取PDF运单… {i}/{len(pdf_files)}")
    progress.empty()
    return we.compare_batch(rows, extractions)


def _mini_card(label: str, value, bg: str, fg: str) -> str:
    return (f'<div class="ceb-mini" style="background:{bg};">'
            f'<div class="v" style="color:{fg};">{value}</div>'
            f'<div class="k">{label}</div></div>')


def _render_report(report: dict) -> None:
    summary, anomaly = report["summary"], report["anomaly"]

    # ---- 批级合理性告警（任务书：防模板变化静默失效，页面顶部高亮） ----
    if anomaly["batch_warning"]:
        st.markdown(
            f'<div class="ceb-problem" style="border-left-color:#B71C1C;'
            f'background:#FFEBEE;"><h4 style="color:#B71C1C;">'
            f'⚠️ 运单模板可能已变化，建议人工抽查</h4>'
            f'<div class="d">{anomaly["warning_text"]}</div></div>',
            unsafe_allow_html=True)

    mismatch_total = (summary[we.ST_WAGON_MISMATCH] + summary[we.ST_SEAL_MISMATCH]
                      + summary[we.ST_BOTH_MISMATCH])
    unmatched_total = (summary[we.ST_PDF_UNMATCHED]
                       + summary[we.ST_EXCEL_UNMATCHED])
    attention_total = (summary[we.ST_EXTRACTION_ANOMALY]
                       + summary[we.ST_DATA_MISSING])

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    with c1:
        st.markdown(_mini_card("PDF运单", summary["total_pdfs"],
                               "#EFF6FF", "#0B5394"), unsafe_allow_html=True)
    with c2:
        st.markdown(_mini_card("清单行", summary["total_excel_rows"],
                               "#F9FAFB", "#374151"), unsafe_allow_html=True)
    with c3:
        st.markdown(_mini_card("✅ 确认正确", summary[we.ST_MATCH],
                               "#E8F5E9", "#1B5E20"), unsafe_allow_html=True)
    with c4:
        st.markdown(_mini_card("⛔ 不一致", mismatch_total,
                               "#FFEBEE" if mismatch_total else "#F9FAFB",
                               "#B71C1C" if mismatch_total else "#6B7280"),
                    unsafe_allow_html=True)
    with c5:
        st.markdown(_mini_card("❓ 未匹配", unmatched_total,
                               "#FFF8E1" if unmatched_total else "#F9FAFB",
                               "#8D6E00" if unmatched_total else "#6B7280"),
                    unsafe_allow_html=True)
    with c6:
        st.markdown(_mini_card("⚠️ 提取异常", attention_total,
                               "#FFF8E1" if attention_total else "#F9FAFB",
                               "#8D6E00" if attention_total else "#6B7280"),
                    unsafe_allow_html=True)

    results = report["results"]
    _render_mismatches([r for r in results if r["status"] in (
        we.ST_WAGON_MISMATCH, we.ST_SEAL_MISMATCH, we.ST_BOTH_MISMATCH)])
    _render_attention([r for r in results if r["status"] in (
        we.ST_EXTRACTION_ANOMALY, we.ST_DATA_MISSING)])
    _render_unmatched([r for r in results if r["status"] in (
        we.ST_PDF_UNMATCHED, we.ST_EXCEL_UNMATCHED)])
    _render_matches([r for r in results if r["status"] == we.ST_MATCH])


def _pair_line(r: dict, field: str) -> str:
    """单字段"PDF值 vs 清单值"的展示片段（含异常态）。"""
    fr = r["field_results"].get(field, {})
    label = _FIELD_LABEL[field]
    pdf_status = fr.get("pdf_status")
    if pdf_status and pdf_status != we.FIELD_OK:
        detail = _ANOMALY_LABEL.get(pdf_status, pdf_status)
        return (f'{label}：<b>PDF提取异常（{detail}）</b>'
                f'<span style="color:#6B7280;">｜清单值：'
                f'{fr.get("excel_value") or "—"}</span>')
    icon = "✅" if fr.get("result") == "match" else "⛔"
    return (f'{label}：{icon} PDF <b>{fr.get("pdf_value") or "—"}</b>'
            f' vs 清单 <b>{fr.get("excel_value") or "—"}</b>')


def _render_mismatches(rows: list[dict]) -> None:
    if not rows:
        return
    st.markdown(f"#### ⛔ 不一致（{len(rows)}）——请在确认运单前当场修改")
    for r in rows:
        title = (f'{r["pdf_container"]}　'
                 f'（清单序号{r["seq"]}）' if r.get("seq") else
                 f'{r["pdf_container"]}　')
        lines = [_pair_line(r, f) for f in ("wagon_no", "seal_no")]
        st.markdown(
            f'<div class="ceb-problem" style="border-left-color:#B71C1C;'
            f'background:#FFEBEE;"><h4>⛔ {title}</h4>'
            + "".join(f'<div class="d">{ln}</div>' for ln in lines)
            + "</div>", unsafe_allow_html=True)
    st.info("不一致的箱子请B同事在订车系统里修改后重新导出运单再核对；"
            "改不成只能回车站重新盖章，发到边境有退运风险。")


def _render_attention(rows: list[dict]) -> None:
    if not rows:
        return
    st.markdown(f"#### ⚠️ 提取异常 / 待确认（{len(rows)}）——不能自动判定，请人工核对")
    for r in rows:
        who = r["filename"] or (r.get("excel_container")
                                and f"清单序号{r['seq']}（{r['excel_container']}）") or "—"
        bits = []
        for f, fr in r.get("field_results", {}).items():
            label = _FIELD_LABEL[f]
            if fr.get("pdf_status") and fr["pdf_status"] != we.FIELD_OK:
                detail = _ANOMALY_LABEL.get(fr["pdf_status"], fr["pdf_status"])
                bits.append(f"{label}：PDF提取异常（{detail}）"
                            + (f"，清单值 {fr['excel_value']}" if fr.get("excel_value") else ""))
            elif fr.get("result") == "mismatch":
                bits.append(f"{label}：⛔ PDF {fr['pdf_value']} ≠ 清单 {fr['excel_value']}")
            elif fr.get("note"):
                bits.append(f"{label}：{fr['note']}")
        extra = "".join(f'<div class="d">{b}</div>' for b in bits)
        notes = "".join(f'<div class="s">{n}</div>' for n in r.get("notes", [])[:2])
        st.markdown(
            f'<div class="ceb-problem" style="border-left-color:#8D6E00;'
            f'background:#FFF8E1;"><h4>⚠️ {who}</h4>{extra}{notes}</div>',
            unsafe_allow_html=True)


def _render_unmatched(rows: list[dict]) -> None:
    if not rows:
        return
    st.markdown(f"#### ❓ 未匹配（{len(rows)}）——请先人工判断方向，不要当缺失处理")
    pdf_side = [r for r in rows if r["status"] == we.ST_PDF_UNMATCHED]
    excel_side = [r for r in rows if r["status"] == we.ST_EXCEL_UNMATCHED]
    if pdf_side:
        st.markdown("**PDF运单的箱号在现场清单中没找到：**")
        df = pd.DataFrame([{
            "PDF文件": r["filename"],
            "PDF箱号": r["pdf_container"],
            "可能原因": "箱号本身录错，或清单里没有这条",
        } for r in pdf_side])
        st.dataframe(df, use_container_width=True, hide_index=True)
        st.caption(we.PDF_UNMATCHED_TEXT)
    if excel_side:
        st.markdown("**现场清单的箱子没等到对应PDF运单：**")
        df = pd.DataFrame([{
            "清单序号": r.get("seq") or f"第{r['row_no']}行",
            "清单箱号": r.get("excel_container"),
            "车底号": r.get("excel_wagon") or "—",
            "铅封号": r.get("excel_seal") or "—",
            "可能原因": "漏传运单，或PDF中箱号录错",
        } for r in excel_side])
        st.dataframe(df, use_container_width=True, hide_index=True)
        st.caption(we.EXCEL_UNMATCHED_TEXT)


def _render_matches(rows: list[dict]) -> None:
    if not rows:
        return
    with st.expander(f"✅ 一致（{len(rows)}）——无需处理，点击展开明细"):
        df = pd.DataFrame([{
            "清单序号": r.get("seq") or "—",
            "箱号": r["pdf_container"] or r.get("excel_container") or "—",
            "车底号": r["field_results"]["wagon_no"]["pdf_value"],
            "铅封号": r["field_results"]["seal_no"]["pdf_value"],
            "PDF文件": r["filename"],
        } for r in rows if r.get("field_results")])
        st.dataframe(df, use_container_width=True, hide_index=True)

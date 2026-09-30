# -*- coding: utf-8 -*-
"""对话式核验助手（原生AI功能·任务书第1节）。

范式：LLM + 工具调用（function calling）——
  用户提问 → LLM 读懂问题 → 需要数值推演时调用 simulate_field_change 工具
  → 工具在真实数据上重跑核验引擎与风险模型 → LLM 基于真实结果组织回答。

诚实性说明：
  - 本功能**只在配置了 API Key 时可用**（ARK_API_KEY 或 GLM_API_KEY/BIGMODEL_API_KEY），
    无 Key 时显式提示"需配置 API Key 启用"，不预置任何假问答（开放式问答无法穷举预置）。
  - 回答由 LLM 实时生成，页面标注"可能存在误差，请以核验报告明细为准"。
  - 单元测试使用脚本化的假 LLM transport 驱动工具调用循环，验证管线正确性；
    真实回答效果见 live 模式实测（汇报中附实际回答内容）。
"""

from __future__ import annotations

import copy
import json
import re

from llm_endpoint import resolve_endpoint
from verification_engine import run_verification

MAX_TOOL_ROUNDS = 3
REQUEST_TIMEOUT = 40

# 数值假设类问题的识别（F10）：这类问题的回答必须经过程序强制重算验证
_HYPOTHETICAL_RE = re.compile(
    r"(如果|假如|假设|要是|改成|改为|变为|变成|调整成|上调|下调|会怎样|会怎么样|"
    r"会变成|风险会|会到多少|多少分|推演|simulate)")

TOOLS_SCHEMA = [
    {
        "type": "function",
        "function": {
            "name": "simulate_field_change",
            "description": (
                "假设推演：修改单证字段值后重新运行全套核验与风险评分。"
                "凡涉及『如果改成X会怎样/风险会变成多少』类数值假设问题，必须调用本工具获取真实计算结果，"
                "禁止凭空推算数字。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "edits": {
                        "type": "object",
                        "description": (
                            "要修改的字段映射，key 为 『单证类型.字段名』"
                            "（如 export_customs_declaration.total_packages、packing_list.gross_weight_kg），"
                            "value 为新值（数值或字符串）。可一次修改多个字段。"
                        ),
                    }
                },
                "required": ["edits"],
            },
        },
    }
]

SYSTEM_PROMPT = (
    "你是「中欧班列单证智能核验系统」的核验助手，帮助单证员/外贸人员理解核验报告。"
    "回答规则：\n"
    "1. 只依据下面提供的核验结果上下文回答，不得编造不存在的检查项或数字；\n"
    "2. 对『如果…改成…』的数值假设问题，必须调用 simulate_field_change 工具重新计算后再回答，"
    "回答中明确给出计算得到的新风险分与等级，并说明变化原因；\n"
    "3. 用简体中文，简洁分层（可用短列表），总长不超过200字；\n"
    "4. 涉及整改动作时引用上下文中的修正建议。"
)


# ---------------------------------------------------------------- 工具实现


def apply_edits(documents: list, edits: dict) -> list:
    """按 {doc_type.field: new_value} 修改单证（返回深拷贝，不改动原数据）。"""
    docs = copy.deepcopy(documents)
    for key, value in (edits or {}).items():
        if "." not in key:
            continue
        doc_type, field = key.split(".", 1)
        for doc in docs:
            if doc.get("doc_type") == doc_type or doc.get("doc_id") == doc_type:
                doc.setdefault("fields", {})[field] = value
    return docs


def simulate_field_change(documents: list, edits: dict, batch_meta: dict | None = None) -> dict:
    """真实重跑：修改字段 → run_verification → 返回紧凑结果（供LLM阅读）。"""
    batch = {"batch_id": "simulation", "documents": apply_edits(documents, edits)}
    if batch_meta:
        batch.update({k: v for k, v in batch_meta.items() if k != "documents"})
    v = run_verification(batch)
    non_pass = [
        f"{r['check_name']}[{r['status']}]: {r['detail']}"
        for r in v["results"] if r["status"] != "PASS"
    ]
    return {
        "new_risk_score": v["risk"]["score"],
        "new_risk_grade": v["risk"]["grade_label"],
        "new_summary": v["summary"],
        "score_breakdown": [f"{b['reason']} +{b['points']}" for b in v["risk"]["breakdown"]],
        "remaining_issues": non_pass,
    }


# ---------------------------------------------------------------- 上下文构造


def build_context(verification: dict, documents: list) -> str:
    """把核验结果+单证字段压成紧凑上下文（控制在LLM可读范围内）。"""
    parts = ["【核验结果摘要】",
             json.dumps({
                 "批次": verification.get("batch_name", ""),
                 "summary": verification["summary"],
                 "risk": {k: verification["risk"][k] for k in ("score", "grade_label")},
                 "分数构成": [f"{b['reason']} +{b['points']}" for b in verification["risk"]["breakdown"]],
             }, ensure_ascii=False)]
    parts.append("【未通过项明细（已按单据分组，可回答『哪份单据有什么问题』）】")
    groups = verification.get("document_groups") or []
    if groups:
        for g in groups:
            if g.get("issues"):
                parts.append(f"◆ {g.get('label')}（{g.get('doc_id')}）")
                for r in g["issues"]:
                    parts.append(f"- {r['check_name']}[{r['status']}] {r['detail']}")
                    if r.get("suggestion"):
                        parts.append(f"  建议: {str(r['suggestion'])[:120]}")
        for r in verification.get("batch_level_issues") or []:
            parts.append("◆ 批次级（整套单证）")
            parts.append(f"- {r['check_name']}[{r['status']}] {r['detail']}")
            if r.get("suggestion"):
                parts.append(f"  建议: {str(r['suggestion'])[:120]}")
    else:
        for r in verification["results"]:
            if r["status"] != "PASS":
                parts.append(f"- {r['check_name']}[{r['status']}] {r['detail']}")
                if r.get("suggestion"):
                    parts.append(f"  建议: {r['suggestion'][:120]}")
    parts.append("【单证字段值（可被simulate工具修改）】")
    for d in documents:
        parts.append(f"- {d.get('doc_type')} ({d.get('doc_id')}): "
                     + json.dumps(d.get("fields", {}), ensure_ascii=False)[:400])
    parts.append("【可用单证类型】invoice/packing_list/railway_waybill/export_customs_declaration/certificate_of_origin")
    return "\n".join(parts)


# ---------------------------------------------------------------- LLM 调用循环


def _post_chat(endpoint: dict, messages: list, tools: list | None = None) -> dict:
    """POST chat/completions（OpenAI兼容）。独立函数便于测试时打桩。"""
    import urllib.request

    body = {"model": endpoint["model"], "messages": messages, "temperature": 0.2}
    if tools:
        body["tools"] = tools
        body["tool_choice"] = "auto"
    req = urllib.request.Request(
        f"{endpoint['base_url'].rstrip('/')}/chat/completions",
        data=json.dumps(body).encode("utf-8"),
        headers={"Authorization": f"Bearer {endpoint['api_key']}",
                 "Content-Type": "application/json"},
    )
    with urllib.request.urlopen(req, timeout=REQUEST_TIMEOUT) as resp:
        return json.loads(resp.read().decode("utf-8"))


def _verified_block(tool_trace: list) -> str:
    """把核验引擎的重算结果组装成权威结论块（F10）：数值回答以程序重算为准，
    不信任LLM自行报出的数字。"""
    recalcs = [t for t in tool_trace if t["tool"] == "simulate_field_change"]
    if not recalcs:
        return ""
    last = recalcs[-1]["result"]
    breakdown = "；".join(last.get("score_breakdown", [])[:3])
    text = (f"\n\n---\n✅ **计算验证（核验引擎真实重算，非模型生成）**：新风险分 "
            f"{last['new_risk_score']}/100（{last['new_risk_grade']}），"
            f"FAIL {last['new_summary']['fail']} 项。")
    if breakdown:
        text += f"\n扣分构成：{breakdown}"
    return text


def answer_question(question: str, verification: dict, documents: list,
                    history: list | None = None) -> dict:
    """
    对话式问答主入口。

    返回：
      {mode: "live", answer, tool_trace: [...]}             —— 真实调用LLM（含工具调用轨迹）；
                                                                数值假设类回答附带引擎重算验证块
      {mode: "unverified", answer: 拒绝文案}                —— 假设类问题但LLM未触发工具重算：
                                                                拒绝展示未经验证的数值结论（F10）
      {mode: "no_key", answer: 提示文案}                    —— 未配置API Key的显式降级
      {mode: "error", answer: 错误提示}                     —— 调用失败（网络/限流等）
    """
    endpoint = resolve_endpoint()
    if endpoint is None:
        return {
            "mode": "no_key",
            "answer": "⚠️ 对话式核验助手需管理员开通AI服务后使用。"
                      "为保证诚实性，本功能不提供预置问答。",
        }

    is_hypothetical = bool(_HYPOTHETICAL_RE.search(question or ""))
    messages = ([{"role": "system", "content": SYSTEM_PROMPT},
                 {"role": "system", "content": build_context(verification, documents)}]
                + list(history or [])
                + [{"role": "user", "content": question}])
    tool_trace = []
    try:
        for _ in range(MAX_TOOL_ROUNDS):
            data = _post_chat(endpoint, messages, tools=TOOLS_SCHEMA)
            msg = data["choices"][0]["message"]
            tool_calls = msg.get("tool_calls") or []
            if not tool_calls:
                answer = (msg.get("content") or "").strip()
                if is_hypothetical and not tool_trace:
                    # F10：假设推演类问题，LLM未调用工具就给出数字 → 拒绝展示
                    return {"mode": "unverified",
                            "answer": "⚠️ 未能完成计算验证：本次推演未触发核验引擎真实重算，"
                                      "系统拒绝呈现未经重算验证的数值结论（模型给出的数字不作为依据）。"
                                      "请重新提问，例如：『如果报关箱数改成480，风险会变成多少？』",
                            "tool_trace": []}
                answer += _verified_block(tool_trace)
                return {"mode": "live", "answer": answer, "tool_trace": tool_trace}
            messages.append(msg)
            for tc in tool_calls:
                fn = tc["function"]
                try:
                    args = json.loads(fn["arguments"] or "{}")
                except json.JSONDecodeError:
                    args = {}
                result = simulate_field_change(documents, args.get("edits", {}))
                tool_trace.append({"tool": fn["name"], "args": args, "result": result})
                messages.append({
                    "role": "tool",
                    "tool_call_id": tc.get("id", ""),
                    "content": json.dumps(result, ensure_ascii=False),
                })
        # 工具轮次用尽：强制收尾（不带tools再问一次）
        data = _post_chat(endpoint, messages, tools=None)
        answer = (data["choices"][0]["message"]["content"] or "").strip()
        if is_hypothetical and not tool_trace:
            return {"mode": "unverified",
                    "answer": "⚠️ 未能完成计算验证：本次推演未触发核验引擎真实重算，"
                              "系统拒绝呈现未经重算验证的数值结论。请换一种问法重试。",
                    "tool_trace": []}
        answer += _verified_block(tool_trace)
        return {"mode": "live", "answer": answer, "tool_trace": tool_trace}
    except Exception as exc:
        return {"mode": "error",
                "answer": f"⚠️ LLM 调用失败（{type(exc).__name__}: {str(exc)[:120]}）。"
                          f"请检查网络与 API Key 配置；核验报告本身不受影响。"}

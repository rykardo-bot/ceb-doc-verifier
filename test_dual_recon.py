# -*- coding: utf-8 -*-
"""双表对账导入测试（Issue #2）。

核心回归：任务书§二的真实案例——
  序号6（2025.7.18）：两边总数一致（2,793,440.46）但科目分类不同
  （己方代理费 20,623.7 vs 联运并入铁路运费）→ 必须判"已确认"，不报警；
对照 8.8：两边科目逐项相等且合计一致 → 同样已确认。
另覆盖：合计不一致→待确认、双方缺失状态、未登记站名提示、
国家层级兜底配对、应用入库（复用 train_trips/settlements/prepayments）。

Excel 用 openpyxl 在内存中按真实表头生成（脱敏：无企业名称）。
运行：pytest test_dual_recon.py -v
"""

import io
from decimal import Decimal

from openpyxl import Workbook

import dual_recon
import train_store

# ---------------------------------------------------------------- 夹具构造

OWN_HEADERS = ["序号", "发运日期", "班列路线", "付款日期", "预付金额",
               "运费", "报关费", "代理费", "结算金额合计"]
AGENT_HEADERS = ["年份", "序号", "月份", "月序号", "开行日期", "发站", "口岸",
                 "到站", "货物品名", "车数", "柜量40hd", "柜量20hd",
                 "折合teu", "铁路运费", "报关费", "服务费", "其他", "合计",
                 "实付运费", "保证金"]


def _own_row(seq, dep, route, pay, prepay, freight, customs, agency, total):
    return [seq, dep, route, pay, prepay, freight, customs, agency, total]


def _agent_row(seq, dep, station, port, dest, rail, customs, service,
               other, total, goods="汽车成套散件", wagon=50,
               hd40=25, hd20=0, teu=50):
    return ["2025", seq, "7", "1", dep, station, port, dest, goods, wagon,
            hd40, hd20, teu, rail, customs, service, other, total,
            total, 0]


def _wb(headers, rows):
    wb = Workbook()
    ws = wb.active
    ws.append(headers)
    for r in rows:
        ws.append(r)
    return wb


def _own_bytes(rows):
    buf = io.BytesIO()
    _wb(OWN_HEADERS, rows).save(buf)
    return buf.getvalue()


def _agent_bytes(rows):
    buf = io.BytesIO()
    _wb(AGENT_HEADERS, rows).save(buf)
    return buf.getvalue()


# 任务书§二真实数字
OWN_0718 = _own_row(6, "2025.7.18", "平旺-满洲里-谢利亚季诺", "2025.7.1",
                    1426003.11, 2738016.76, 34800, 20623.7, 2793440.46)
AGENT_0718 = _agent_row(6, "2025.7.18", "平旺", "满洲里", "俄罗斯",
                        2758640.46, 34800, 0, 0, 2793440.46)

OWN_0808 = _own_row(12, "2025.8.8", "大同-二连-莫斯科", "2025.8.1",
                    2581233.30, 2581233.30, 0, 0, 2581233.30)
AGENT_0808 = _agent_row(12, "2025.8.8", "大同", "二连浩特", "俄罗斯",
                        2581233.30, 0, 0, 0, 2581233.30)


def _preview(own_rows, agent_rows):
    return dual_recon.build_preview(_own_bytes(own_rows),
                                    _agent_bytes(agent_rows))


# ---------------------------------------------------------------- 7-18 真实案例

def test_july18_same_total_different_categories_is_confirmed():
    """任务书§二核心案例：总数一致但科目分类不同 → 已确认，不报警；
    科目差异仅展示（对照物证：己方代理费 20623.7 / 联运服务费 0）。"""
    result = _preview([OWN_0718], [AGENT_0718])
    assert result["stats"][dual_recon.CONFIRMED] == 1
    assert result["stats"][dual_recon.PENDING] == 0
    pair = result["pairs"][0]
    assert pair["status"] == dual_recon.CONFIRMED
    assert pair["diff"] == "0" or Decimal(pair["diff"]) == 0
    # 科目明细保留展示（不是消失）
    own_items = {i["label"]: i["value"] for i in
                 pair["category_breakdown"]["own"]}
    agent_items = {i["label"]: i["value"] for i in
                   pair["category_breakdown"]["agent"]}
    assert own_items["代理费"] == "20623.7"
    assert own_items["运费"] == "2738016.76"
    assert agent_items["铁路运费"] == "2758640.46"


def test_aug08_matching_categories_confirmed():
    """对照 8.8：科目逐项相等 + 合计一致 → 已确认。"""
    result = _preview([OWN_0808], [AGENT_0808])
    assert result["stats"][dual_recon.CONFIRMED] == 1


def test_total_mismatch_is_pending():
    """合计不一致 → 待确认（哪怕只差一分钱）。"""
    agent = list(AGENT_0808)
    agent[17] = 2581233.31                     # 合计（index 17）差 0.01
    result = _preview([OWN_0808], [agent])
    assert result["stats"][dual_recon.PENDING] == 1
    pair = result["pairs"][0]
    assert pair["status"] == dual_recon.PENDING
    assert pair["diff"] == "0.01"


# ---------------------------------------------------------------- 缺失状态

def test_agent_only_is_own_missing():
    """联运有、己方没有 → 己方缺失。"""
    agent = _agent_row(20, "2025.9.10", "大同", "霍尔果斯", "中亚",
                       1500000, 0, 0, 0, 1500000)
    result = _preview([OWN_0718], [agent])
    assert result["stats"][dual_recon.OWN_MISSING] == 1


def test_own_only_is_agent_missing():
    """己方有、联运没有 → 联运缺失。"""
    own = _own_row(21, "2025.9.12", "中鼎-二连-莫斯科", "2025.9.1",
                   0, 900000, 0, 0, 900000)
    result = _preview([own], [AGENT_0718])
    assert result["stats"][dual_recon.AGENT_MISSING] == 1


# ---------------------------------------------------------------- 站编解析

def test_unknown_station_reported_for_backfill():
    """查不到的站名 → 收集为待补录清单，对应行标为错误不瞎猜。"""
    own = _own_row(30, "2025.9.15", "乌兰-满洲里-莫斯科", "2025.9.1",
                   0, 100000, 0, 0, 100000)
    result = _preview([own], [AGENT_0718])
    names = {(u["category"], u["name"]) for u in result["unknown_names"]}
    assert ("station", "乌兰") in names


def test_alias_and_country_fallback_matching():
    """别名匹配（二连浩特）与国家层兜底（联运到站=俄罗斯 vs 己方到站=谢利亚季诺）。"""
    result = _preview([OWN_0808], [AGENT_0808])
    assert result["stats"][dual_recon.CONFIRMED] == 1   # 二连浩特 别名已命中
    # 国家兜底：己方到站=谢利亚季诺（MSK口径不同），联运到站=俄罗斯
    own_msk = _own_row(6, "2025.7.18", "平旺-满洲里-莫斯科", "2025.7.1",
                       1426003.11, 2793440.46, 0, 0, 2793440.46)
    result2 = _preview([own_msk], [AGENT_0718])
    # 己方=莫斯科(MSK)，联运=俄罗斯(RU 国家层) → 国家兜底配对成功
    assert result2["stats"][dual_recon.CONFIRMED] == 1


def test_resolve_name_levels():
    rows = train_store.list_codes()
    # 别名
    assert dual_recon.resolve_name("二连浩特", "port", rows)[0] == "EL"
    # 国家层
    code, country, level = dual_recon.resolve_name("俄罗斯", "dest", rows)
    assert (code, country, level) == ("RU", "俄罗斯", "country")
    # 具体站
    code, country, _ = dual_recon.resolve_name("谢利亚季诺", "dest", rows)
    assert code == "XLY" and country == "俄罗斯"
    # 未登记
    assert dual_recon.resolve_name("乌兰乌德", "dest", rows) is None


# ---------------------------------------------------------------- 应用入库

def test_apply_creates_trip_settlement_prepayment():
    """已确认对 → 自动入库：新建班列 + 结算明细 + 己方预付款。"""
    own = dict(zip(OWN_HEADERS, OWN_0718))
    agent = dict(zip(AGENT_HEADERS, AGENT_0718))
    own_slim = {"row_index": 2, "dep_date": "2025-07-18",
                "station_name": "平旺", "port_name": "满洲里",
                "dest_name": "谢利亚季诺", "station_code": "PW",
                "port_code": "MZL", "dest_code": "XLY",
                "dest_country": "俄罗斯", "route_raw": "平旺-满洲里-谢利亚季诺",
                "pay_date": "2025.7.1", "prepay": "1426003.11",
                "total": "2793440.46"}
    agent_slim = {"row_index": 2, "dep_date": "2025-07-18",
                  "station_name": "平旺", "port_name": "满洲里",
                  "dest_name": "俄罗斯", "station_code": "PW",
                  "port_code": "MZL", "dest_code": "RU",
                  "goods_name": "汽车成套散件", "wagon_count": "50",
                  "container_40hd": "25", "container_20hd": "0",
                  "teu_total": "50", "total": "2793440.46",
                  "actual_freight": "2793440.46"}
    result = dual_recon.apply_preview([{
        "match_key": "test|july18", "decision": "use_agent",
        "own": own_slim, "agent": agent_slim}], by="t")
    assert result["counts"]["created"] == 1
    trip_no = result["results"][0]["trip_no"]
    assert trip_no == "20250718-PW-MZL-RU-T"
    settlement = train_store.get_settlement(trip_no)
    assert Decimal(str(settlement["settle_total"])) == Decimal("2793440.46")
    prepays = train_store.list_prepayments(trip_no)
    assert len(prepays) == 1
    assert Decimal(str(prepays[0]["amount"])) == Decimal("1426003.11")


def test_apply_pending_requires_explicit_choice():
    """待确认对：skip 则不入库。"""
    result = dual_recon.apply_preview([{
        "match_key": "test|pending", "decision": "skip",
        "own": {"dep_date": "2025-08-08", "station_code": "DT",
                "port_code": "EL", "dest_code": "MSK",
                "total": "2581233.30", "prepay": ""},
        "agent": {"dep_date": "2025-08-08", "station_code": "DT",
                  "port_code": "EL", "dest_code": "RU",
                  "total": "2581233.31"}}], by="t")
    assert result["counts"]["skipped"] == 1
    assert result["counts"]["created"] == 0


# ---------------------------------------------------------------- Issue #3 三个bug回归

def test_sergeevo_region_matching_fixed():
    """Issue#3 Bug1：中鼎-霍尔果斯-谢尔盖利(己方) vs 中鼎-霍尔果斯-中亚(联运)
    ——谢尔盖利 country 修正为'中亚'后，国家层兜底应配对为已确认，
    不得再出现 4 条互不相关的缺失记录。"""
    own_0809 = _own_row(6, "2025.8.9", "中鼎-霍尔果斯-谢尔盖利", "2025.8.1",
                        0, 2034027.60, 0, 0, 2034027.60)
    own_0813 = _own_row(7, "2025.8.13", "中鼎-霍尔果斯-谢尔盖利", "2025.8.1",
                        0, 2031824.40, 0, 0, 2031824.40)
    agent_0809 = _agent_row(6, "2025.8.9", "中鼎", "霍尔果斯", "中亚",
                            2034027.60, 0, 0, 0, 2034027.60)
    agent_0813 = _agent_row(7, "2025.8.13", "中鼎", "霍尔果斯", "中亚",
                            2031824.40, 0, 0, 0, 2031824.40)
    result = _preview([own_0809, own_0813], [agent_0809, agent_0813])
    assert result["stats"][dual_recon.CONFIRMED] == 2
    assert result["stats"][dual_recon.OWN_MISSING] == 0
    assert result["stats"][dual_recon.AGENT_MISSING] == 0
    # 站编表数据核对
    xeg = next(c for c in train_store.list_codes() if c["code"] == "XEG")
    assert xeg["country"] == "中亚"


def test_moscow_mixed_note_cleaned_no_double_report():
    """Issue#3 Bug2：到站"莫斯科混编（电煤/谢利）"清洗后命中莫斯科——
    只产生一条已确认，不再同时出现 未登记站名警告+解析错误 两条。"""
    own = _own_row(6, "2025.7.18", "平旺-满洲里-莫斯科混编（电煤/谢利）",
                   "2025.7.1", 1426003.11, 2793440.46, 0, 0, 2793440.46)
    result = _preview([own], [AGENT_0718])
    assert result["stats"][dual_recon.CONFIRMED] == 1
    assert result["stats"][dual_recon.UNRESOLVED_NAME] == 0
    assert result["stats"]["error"] == 0
    assert result["unknown_names"] == []
    pair = result["pairs"][0]
    assert pair["own"]["dest_code"] == "MSK"


def test_summary_total_rows_skipped():
    """Issue#3 Bug3：表格末尾"合计"汇总行被静默跳过——
    不再出现 全空值的解析错误条目。"""
    own_rows = [OWN_0718,
                _own_row("合计", "", "", "", 1426003.11, 2738016.76,
                         34800, 20623.7, 2793440.46)]
    agent_rows = [AGENT_0718,
                  _agent_row("合计", "", "", "", "", 0, 0, 0, 0,
                             2793440.46)]
    result = _preview(own_rows, agent_rows)
    assert result["stats"]["error"] == 0
    assert result["stats"][dual_recon.CONFIRMED] == 1
    # 汇总行没有作为缺失/错误进入任何状态
    total_statuses = [p["status"] for p in result["pairs"]]
    assert "error" not in total_statuses


def test_lintu_trips_marked_not_duplicates():
    """Issue#3 待标注：10-10 临/图两趟车是正常业务——
    两条已确认记录各自带 train_type（L/T），展示可区分。"""
    own_lin = _own_row(8, "2025/10/10临", "平旺-满洲里-别雷拉斯特",
                       "2025.9.29", 0, 2966062, 0, 0, 2966062)
    own_tu = _own_row(9, "2025/10/10", "平旺-满洲里-别雷拉斯特",
                      "2025.9.29", 0, 2966062, 0, 0, 2966062)
    agent_lin = _agent_row(8, "2025/10/10临", "平旺", "满洲里",
                           "别雷拉斯特", 2966062, 0, 0, 0, 2966062)
    agent_tu = _agent_row(9, "2025/10/10", "平旺", "满洲里",
                          "别雷拉斯特", 2966062, 0, 0, 0, 2966062)
    result = _preview([own_lin, own_tu], [agent_lin, agent_tu])
    assert result["stats"][dual_recon.CONFIRMED] == 2
    types = sorted(p["own"]["train_type"] for p in result["pairs"])
    assert types == ["L", "T"]      # 临/图各自保留，可区分展示


def test_apply_zero_prepay_not_persisted():
    """己方预付为 0/空 → 不落 0 元预付款记录（纯噪音且干扰台账预付核对）。"""
    own_slim = {"row_index": 2, "dep_date": "2025-07-19",
                "station_name": "平旺", "port_name": "满洲里",
                "dest_name": "谢利亚季诺", "station_code": "PW",
                "port_code": "MZL", "dest_code": "XLY",
                "dest_country": "俄罗斯", "route_raw": "平旺-满洲里-谢利亚季诺",
                "pay_date": "2025.7.2", "prepay": "0",
                "total": "2793440.46"}
    agent_slim = {"row_index": 2, "dep_date": "2025-07-19",
                  "station_name": "平旺", "port_name": "满洲里",
                  "dest_name": "俄罗斯", "station_code": "PW",
                  "port_code": "MZL", "dest_code": "RU",
                  "goods_name": "汽车成套散件", "wagon_count": "50",
                  "container_40hd": "25", "container_20hd": "0",
                  "teu_total": "50", "total": "2793440.46",
                  "actual_freight": "2793440.46"}
    result = dual_recon.apply_preview([{
        "match_key": "test|zeroprepay", "decision": "use_agent",
        "own": own_slim, "agent": agent_slim}], by="t")
    assert result["counts"]["created"] == 1
    trip_no = result["results"][0]["trip_no"]
    assert train_store.list_prepayments(trip_no) == []

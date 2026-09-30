# -*- coding: utf-8 -*-
"""数据核对模块v1 API/存储集成测试（班列编号+联运结算/补贴对账）。

覆盖（任务书验收口径）：
  - 角色矩阵：匿名401 / business 403 / finance·admin 200；字典/配置仅 admin；
  - 9条夹具编号回归（走 API 建班列）；10-10临/图专项；-01 兜底去重；
  - EST 预估编号生命周期：创建 → 预付款挂预估编号 → 锁定 → 自动改挂 +
    映射留痕（est_no/official_no/locked_at/locked_by）；
  - 结算编辑：结算≠实付且无差异原因 → needs_reason 提示；
  - 补贴：重复值检测命中夹具 7/8/9 三条；存疑标记留痕；
  - 导入：预览（新增/一致/冲突/坏行）→ keep 不改数据 → overwrite 生效；
  - 三方对账阈值配置读取；审计动作落库断言。
运行：pytest test_datacheck.py -v（需 PostgreSQL，见 conftest）
"""

import io
import json
from decimal import Decimal
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

import audit
import db
import train_store
from api import app

FIXTURE = json.loads((Path(__file__).parent / "sample_data" / "datacheck"
                      / "datacheck_trips_v1.json").read_text(encoding="utf-8"))

client = TestClient(app)


@pytest.fixture
def dc_users():
    """每用例三个角色账号（conftest 每用例清空 users 表，故为 function 级；
    audit_logs 只增，用唯一用户名隔离查询）。"""
    import uuid
    tag = uuid.uuid4().hex[:8]
    password = "Passw0rd1"
    names = {}
    tokens = {}
    for role in ("finance", "admin", "business"):
        name = f"dcapi{tag}_{role}"
        auth_service_create(name, password, role)
        names[role] = name
        resp = client.post("/auth/login",
                           json={"username": name, "password": password})
        assert resp.status_code == 200, resp.text
        tokens[role] = {"Authorization": f"Bearer {resp.json()['access_token']}"}
    return {"names": names, "headers": tokens}


def auth_service_create(name, password, role):
    import auth_service
    auth_service.create_user(name, password, role)


# ---------------------------------------------------------------- 角色矩阵

def test_role_matrix_anonymous_forbidden_allowed(dc_users):
    h = dc_users["headers"]
    assert client.get("/datacheck/trips").status_code == 401
    assert client.get("/datacheck/trips", headers=h["business"]).status_code == 403
    assert client.get("/datacheck/trips", headers=h["finance"]).status_code == 200
    assert client.get("/datacheck/trips", headers=h["admin"]).status_code == 200
    # 字典/配置为管理面：finance 403
    body = {"category": "dest", "code": "MN", "name": "蒙古"}
    assert client.put("/datacheck/codes", headers=h["finance"],
                      json=body).status_code == 403
    assert client.put("/datacheck/codes", headers=h["admin"],
                      json=body).status_code == 200
    db.execute("DELETE FROM train_code_dict WHERE category='dest' AND code='MN'")
    assert client.put("/datacheck/config", headers=h["finance"],
                      json={"fields": {"datacheck.recon_warn_pct": 5}}).status_code == 403


# ---------------------------------------------------------------- 编号回归

def test_create_nine_fixture_trips_via_api(dc_users):
    """9条夹具通过 API 逐一登记，编号与期望完全一致（含10-10临/图不合并）。"""
    h = dc_users["headers"]["finance"]
    created = []
    for rec, expected in zip(FIXTURE["records"], FIXTURE["expected_numbers"]):
        resp = client.post("/datacheck/trips", headers=h, json={
            "dep_date": rec["dep_date"], "station_code": rec["station"],
            "port_code": rec["port"], "dest_code": rec["dest"],
            "train_type": rec["train_type"], "goods_name": rec["goods_name"],
            "wagon_count": rec["wagon_count"],
            "container_40hd": rec["container_40hd"],
            "container_20hd": rec["container_20hd"],
            "teu_total": rec["teu_total"]})
        assert resp.status_code == 200, resp.text
        trip = resp.json()
        assert trip["trip_no"] == expected
        assert trip["suffix"] == 0
        created.append(trip["trip_no"])
    assert len(set(created)) == 9
    assert "20251010-PW-MZL-RU-L" in created and "20251010-PW-MZL-RU-T" in created


def test_dedup_suffix_on_conflicting_registration(dc_users):
    """同日同线路同类型重复登记：第1条无后缀，第2条 -01，第3条 -02。"""
    h = dc_users["headers"]["finance"]
    expected = ["20250901-DT-EL-RU-T", "20250901-DT-EL-RU-T-01",
                "20250901-DT-EL-RU-T-02"]
    for want in expected:
        resp = client.post("/datacheck/trips", headers=h, json={
            "dep_date": "2025-09-01", "station_code": "DT", "port_code": "EL",
            "dest_code": "RU", "train_type": "T"})
        assert resp.status_code == 200
        assert resp.json()["trip_no"] == want
        assert resp.json()["suffix"] == (0 if want.endswith("T") else
                                         int(want.rsplit("-", 1)[1]))


def test_unknown_station_rejected_with_hint(dc_users):
    h = dc_users["headers"]["finance"]
    resp = client.post("/datacheck/trips", headers=h, json={
        "dep_date": "2025-09-02", "station_code": "XX", "port_code": "EL",
        "dest_code": "RU", "train_type": "T"})
    assert resp.status_code == 422
    assert "代码字典" in resp.json()["detail"]


# ---------------------------------------------------------------- 结算与补贴

def test_import_unknown_station_counts_as_error(dc_users):
    """站名未登记的行：action=error 且必须计入 stats["error"]。

    此前只标 error 行、漏计计数——汇总条显示"错误 0"，财务误以为全部可导入。
    """
    h = dc_users["headers"]["finance"]
    data = _xlsx_bytes("trip", [
        ["2025-09-22", "乌兰", "二连", "俄罗斯", "T", "", "", "", "", "",
         "", "", "", "", "2500000", "", "", ""],        # 发站未登记 → error
    ])
    preview = client.post("/datacheck/import/preview", headers=h,
                          params={"kind": "trip"},
                          files={"file": ("t.xlsx", data)}).json()
    assert preview["stats"]["error"] == 1
    row = next(r for r in preview["rows"] if r["action"] == "error")
    assert "代码字典" in row["message"]
    assert "错误 1" in preview["message"]


def test_settlement_diff_reason_prompt(dc_users):
    """结算合计≠实付运费且差异原因为空 → 返回 needs_reason 提示（不阻断）。"""
    h = dc_users["headers"]["finance"]
    client.post("/datacheck/trips", headers=h, json={
        "dep_date": "2025-09-03", "station_code": "ZD", "port_code": "HGS",
        "dest_code": "ZY", "train_type": "T"})
    no = "20250903-ZD-HGS-ZY-T"
    resp = client.put(f"/datacheck/trips/{no}/settlement", headers=h,
                      json={"fields": {"settle_total": "1734027.60",
                                        "actual_freight": 1730000}})
    check = resp.json()["check"]
    assert check["has_diff"] is True and check["needs_reason"] is True
    assert "待人工填写差异原因" in check["message"]
    # 补齐原因后提示消失
    resp = client.put(f"/datacheck/trips/{no}/settlement", headers=h,
                      json={"fields": {"diff_reason": "口岸滞留费另结"}})
    assert resp.json()["check"]["needs_reason"] is False
    # 预付覆盖核对
    client.post(f"/datacheck/trips/{no}/prepays", headers=h,
                json={"paid_at": "2025-08-30", "amount": "1500000",
                      "remark": "首笔预付"})
    detail = client.get(f"/datacheck/trips/{no}", headers=h).json()
    assert detail["checks"]["prepay"]["verdict"] == "shortage"
    # API JSON 序列化 Decimal→float，数值比较前归一
    diff = Decimal(str(detail["checks"]["prepay"]["diff"]))
    assert diff == Decimal("1734027.60") - Decimal("1500000")


def test_subsidy_duplicate_detection_and_suspect_mark(dc_users):
    """夹具 7/8/9（10-05/10-10临/10-10图）的联运垫付补贴100完全相同：
    /datacheck/recon 的重复值检测必须恰好命中这三条；存疑标记留痕。"""
    h = dc_users["headers"]["finance"]
    for rec, expected in zip(FIXTURE["records"], FIXTURE["expected_numbers"]):
        client.post("/datacheck/trips", headers=h, json={
            "dep_date": rec["dep_date"], "station_code": rec["station"],
            "port_code": rec["port"], "dest_code": rec["dest"],
            "train_type": rec["train_type"]})
        sub = {k: str(v) if v is not None else ""
               for k, v in rec["subsidy"].items() if k != "diff_reason"}
        resp = client.put(f"/datacheck/trips/{expected}/subsidy", headers=h,
                          json={"fields": sub})
        assert resp.status_code == 200, resp.text
    recon = client.get("/datacheck/recon", headers=h, params={
        "date_from": "2025-07-01", "date_to": "2025-10-31"}).json()
    dup_ly = {trip for trip, flags in recon["duplicate_flags"].items()
              if any(f["field"] == "ly_advance_100" for f in flags)}
    assert dup_ly == {"20251005-PW-MZL-RU-T", "20251010-PW-MZL-RU-L",
                      "20251010-PW-MZL-RU-T"}
    # 三方一致（合成值 dt=ly=auth）不触发阈值；阈值口径来自配置
    assert recon["thresholds"]["pct"] == "5" and recon["thresholds"]["amount"] == "5000"
    # 存疑标记（任务书 §三：不直接采信原始数值）
    resp = client.post("/datacheck/trips/20251005-PW-MZL-RU-T/suspect",
                       headers=h, json={"review_status": "suspect",
                                        "note": "与相邻记录重复，疑似复制未更新"})
    assert resp.status_code == 200
    assert resp.json()["review_status"] == "suspect"
    rows = audit.query(action=audit.DC_SUSPECT_MARK,
                       username=dc_users["names"]["finance"], limit=5)
    assert rows and rows[0]["object_id"] == "20251005-PW-MZL-RU-T"


# ---------------------------------------------------------------- EST 生命周期

def test_est_number_lifecycle_with_prepay_relink(dc_users):
    """任务书 §一：预估编号 → 锁定正式编号 → 预付款自动改挂，映射留痕。"""
    h = dc_users["headers"]["finance"]
    est = client.post("/datacheck/est", headers=h, json={
        "dep_date": "2025-09-10", "station_code": "PW", "port_code": "MZL",
        "dest_code": "RU", "train_type": "T"}).json()
    assert est["est_no"] == "EST-20250910-PW-MZL-RU-T"
    # 预付款挂预估编号（est_ref 留痕）
    pp = client.post(f"/datacheck/trips/{est['est_no']}/prepays", headers=h,
                     json={"paid_at": "2025-09-01", "amount": "800000",
                           "remark": "按预估编号预付"}).json()
    assert pp["trip_no"] == est["est_no"] and pp["est_ref"] == est["est_no"]
    # 查询提示"尚无正式编号，仅有预估编号"
    lookup = client.get("/datacheck/lookup", headers=h,
                        params={"dep_date": "2025-09-10"}).json()
    assert lookup["trips"] == [] and est["est_no"] in lookup["message"]
    # 锁定：正式发运日期定为 09-12 → 生成正式编号并改挂预付款
    lock = client.post(f"/datacheck/est/{est['est_no']}/lock", headers=h,
                       json={"dep_date": "2025-09-12"}).json()
    assert lock["official_no"] == "20250912-PW-MZL-RU-T"
    assert lock["created_trip"] is True and lock["relinked_prepays"] == 1
    # 映射留痕：est 行 + 详情编号沿革
    est_row = train_store.get_est(est["est_no"])
    assert est_row["official_no"] == "20250912-PW-MZL-RU-T"
    assert est_row["locked_by"] == dc_users["names"]["finance"]
    assert est_row["locked_at"] is not None
    detail = client.get("/datacheck/trips/20250912-PW-MZL-RU-T", headers=h).json()
    assert detail["est_history"][0]["est_no"] == est["est_no"]
    assert detail["prepays"][0]["trip_no"] == "20250912-PW-MZL-RU-T"
    assert detail["prepays"][0]["est_ref"] == est["est_no"]
    # 重复锁定拒绝
    resp = client.post(f"/datacheck/est/{est['est_no']}/lock", headers=h,
                       json={"dep_date": "2025-09-12"})
    assert resp.status_code == 422
    # 锁定审计
    rows = audit.query(action=audit.DC_EST_LOCK,
                       username=dc_users["names"]["finance"], limit=5)
    assert rows and rows[0]["detail"]["relinked_prepays"] == 1


def test_lock_onto_existing_trip_links_without_duplicate(dc_users):
    """锁定时若已存在同要素正式班列 → 直接关联既有编号，不重复建班列。"""
    h = dc_users["headers"]["finance"]
    client.post("/datacheck/trips", headers=h, json={
        "dep_date": "2025-09-15", "station_code": "DT", "port_code": "EL",
        "dest_code": "RU", "train_type": "T"})
    est = client.post("/datacheck/est", headers=h, json={
        "dep_date": "2025-09-15", "station_code": "DT", "port_code": "EL",
        "dest_code": "RU", "train_type": "T"}).json()
    lock = client.post(f"/datacheck/est/{est['est_no']}/lock", headers=h,
                       json={"dep_date": "2025-09-15"}).json()
    assert lock["official_no"] == "20250915-DT-EL-RU-T"
    assert lock["created_trip"] is False


# ---------------------------------------------------------------- Excel 导入

def _xlsx_bytes(kind: str, rows: list[list]) -> bytes:
    from openpyxl import Workbook
    import datacheck_import as imp
    cols = imp.TRIP_COLUMNS if kind == "trip" else imp.SUBSIDY_COLUMNS
    wb = Workbook()
    ws = wb.active
    ws.append([c[1] for c in cols])
    ws.append([c[2] for c in cols])
    for row in rows:
        ws.append(row)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def test_import_preview_conflict_then_keep_and_overwrite(dc_users):
    """导入：预览标注 冲突 → keep 不改数据 → overwrite 覆盖（任务书 §五）。"""
    h = dc_users["headers"]["finance"]
    client.post("/datacheck/trips", headers=h, json={
        "dep_date": "2025-09-20", "station_code": "PW", "port_code": "MZL",
        "dest_code": "RU", "train_type": "T"})
    no = "20250920-PW-MZL-RU-T"
    client.put(f"/datacheck/trips/{no}/settlement", headers=h,
               json={"fields": {"settle_total": "2690000"}})
    data = _xlsx_bytes("trip", [
        ["2025-09-20", "平旺", "满洲里", "俄罗斯", "T", "", "", "", "", "",
         "", "", "", "", "2700000", "", "", ""],        # 结算合计不一致 → 冲突
        ["2025.9.21", "大同", "二连", "俄罗斯", "临", "", "", "", "", "",
         "", "", "", "", "2600000", "", "", ""],        # 新增（临→L）
        ["合计", "", "", "", "", "", "", "", "", "",
         "", "", "", "", "", "", "", ""],               # 坏行 → error 跳过
    ])
    preview = client.post("/datacheck/import/preview", headers=h,
                          params={"kind": "trip"},
                          files={"file": ("t.xlsx", data)}).json()
    assert preview["stats"] == {"new": 1, "same": 0, "conflict": 1, "error": 1}
    by_action = {r["action"]: r for r in preview["rows"]}
    assert by_action["conflict"]["trip_no"] == no
    assert by_action["conflict"]["diffs"]["settle_total"]["old"] == "2690000.00"
    assert by_action["new"]["trip_no"] == "20250921-DT-EL-RU-L"
    # keep：数据不变
    decisions = [{"row_index": by_action["conflict"]["row_index"],
                  "values": by_action["conflict"]["values"], "decision": "keep"}]
    result = client.post("/datacheck/import/apply", headers=h,
                         json={"kind": "trip", "decisions": decisions}).json()
    assert result["counts"]["kept"] == 1
    assert train_store.get_settlement(no)["settle_total"] == Decimal("2690000.00")
    # overwrite 冲突行 + create 新增行
    decisions = [
        {"row_index": by_action["conflict"]["row_index"],
         "values": by_action["conflict"]["values"], "decision": "overwrite"},
        {"row_index": by_action["new"]["row_index"],
         "values": by_action["new"]["values"], "decision": "create"},
    ]
    result = client.post("/datacheck/import/apply", headers=h,
                         json={"kind": "trip", "decisions": decisions}).json()
    assert result["counts"]["updated"] == 1 and result["counts"]["created"] == 1
    assert train_store.get_settlement(no)["settle_total"] == Decimal("2700000.00")
    rows = audit.query(action=audit.DC_IMPORT_APPLY,
                       username=dc_users["names"]["finance"], limit=5)
    assert rows and rows[0]["detail"]["updated"] == 1
    # 二次导入同数据 → 两行均一致（conflict 行覆盖后相同 + new 行已存在）
    preview2 = client.post("/datacheck/import/preview", headers=h,
                           params={"kind": "trip"},
                           files={"file": ("t.xlsx", data)}).json()
    assert preview2["stats"]["same"] == 2 and preview2["stats"]["new"] == 0


def test_subsidy_import_requires_existing_trip(dc_users):
    h = dc_users["headers"]["finance"]
    data = _xlsx_bytes("subsidy", [
        ["2025.12.31", "T", "平旺", "满洲里", "俄罗斯",
         "1", "1", "1", "1", "1", "1", "1", "1", ""],
    ])
    preview = client.post("/datacheck/import/preview", headers=h,
                          params={"kind": "subsidy"},
                          files={"file": ("s.xlsx", data)}).json()
    assert preview["stats"]["new"] == 1
    decisions = [{"row_index": r["row_index"], "values": r["values"],
                  "decision": "overwrite"} for r in preview["rows"]
                 if r["action"] == "new"]
    result = client.post("/datacheck/import/apply", headers=h,
                         json={"kind": "subsidy", "decisions": decisions}).json()
    assert result["counts"]["error"] == 1
    assert "尚未登记" in result["results"][0]["message"]


def test_import_template_downloadable(dc_users):
    h = dc_users["headers"]["finance"]
    for kind in ("trip", "subsidy"):
        resp = client.get("/datacheck/import/template", headers=h,
                          params={"kind": kind})
        assert resp.status_code == 200
        assert resp.content[:2] == b"PK"    # xlsx（zip）魔数


# ---------------------------------------------------------------- 配置阈值

def test_config_threshold_roundtrip(dc_users):
    """阈值可配置（任务书 §三）：admin 修改后三方对账口径随之变化。"""
    ha = dc_users["headers"]["admin"]
    hf = dc_users["headers"]["finance"]
    original = client.get("/datacheck/config", headers=hf).json()["thresholds"]
    try:
        resp = client.put("/datacheck/config", headers=ha,
                          json={"fields": {"datacheck.recon_warn_amount": 999999}})
        assert resp.status_code == 200
        thresholds = client.get("/datacheck/config", headers=hf).json()["thresholds"]
        assert thresholds["amount"] == "999999"
        rows = audit.query(action=audit.DC_CONFIG_UPDATE,
                           username=dc_users["names"]["admin"], limit=5)
        assert rows
    finally:
        client.put("/datacheck/config", headers=ha,
                   json={"fields": {"datacheck.recon_warn_amount":
                                    int(Decimal(original["amount"]))}})


# ---------------------------------------------------------------- 审计留痕

def test_field_edits_audited_with_before_after(dc_users):
    """编辑留痕：逐字段 before/after（复用单据核对 EDIT_FIELD 口径）。"""
    h = dc_users["headers"]["finance"]
    client.post("/datacheck/trips", headers=h, json={
        "dep_date": "2025-09-25", "station_code": "ZD", "port_code": "HGS",
        "dest_code": "ZY", "train_type": "T", "goods_name": "测试品名"})
    no = "20250925-ZD-HGS-ZY-T"
    client.put(f"/datacheck/trips/{no}/basic", headers=h,
               json={"fields": {"goods_name": "光伏组件", "wagon_count": 42}})
    rows = audit.query(username=dc_users["names"]["finance"],
                       action=audit.EDIT_FIELD, limit=100)
    trip_rows = [r for r in rows if r["object_id"] == f"{no}/goods_name"]
    assert trip_rows, [r["object_id"] for r in rows]
    assert trip_rows[0]["before_value"]["value"] == "测试品名"
    assert trip_rows[0]["after_value"]["value"] == "光伏组件"


# ---------------------------------------------------------------- 代码字典维护（补充任务）

def test_code_create_mode_rejects_duplicate(dc_users):
    """新增判重：同类别同缩写已存在（含停用状态）→ 422 明确提示，
    不静默覆盖（补充任务 §6）。"""
    h = dc_users["headers"]["admin"]
    body = {"category": "station", "code": "PW", "name": "平旺（改）",
            "sort": 9, "active": True}
    # PW 是种子数据（启用中）：create 模式必须拒绝
    r = client.put("/datacheck/codes", params={"mode": "create"},
                   headers=h, json=body)
    assert r.status_code == 422
    assert "已登记" in r.json()["detail"]

    # 停用后再 create：仍拒绝（不论启用还是停用状态）
    client.put("/datacheck/codes", headers=h,
               json={"category": "station", "code": "PW", "name": "平旺",
                     "sort": 1, "active": False})
    r2 = client.put("/datacheck/codes", params={"mode": "create"},
                    headers=h, json=body)
    assert r2.status_code == 422
    # 恢复启用，避免影响其他用例
    client.put("/datacheck/codes", headers=h,
               json={"category": "station", "code": "PW", "name": "平旺",
                     "sort": 1, "active": True})


def test_code_audit_actions_distinguished(dc_users):
    """审计动作区分：新增=DC_CODE_CREATE、停用=DC_CODE_DISABLE、
    修改=DC_CODE_UPDATE（补充任务 §5）。"""
    h = dc_users["headers"]["admin"]
    admin_name = dc_users["names"]["admin"]
    db.execute("DELETE FROM train_code_dict WHERE category='dest' AND code='MN'")
    audit.query  # noqa: B018 -- 引用确保模块已加载

    # 新增
    r = client.put("/datacheck/codes", params={"mode": "create"},
                   headers=h, json={"category": "dest", "code": "MN",
                                    "name": "蒙古", "sort": 4, "active": True})
    assert r.status_code == 200
    rows = audit.query(username=admin_name, action=audit.DC_CODE_CREATE)
    assert any(r["object_id"] == "dest/MN" for r in rows)

    # 修改名称/排序（保持启用）
    client.put("/datacheck/codes", headers=h,
               json={"category": "dest", "code": "MN", "name": "蒙古国",
                     "sort": 5, "active": True})
    rows = audit.query(username=admin_name, action=audit.DC_CODE_UPDATE)
    assert any(r["object_id"] == "dest/MN" for r in rows)

    # 停用
    client.put("/datacheck/codes", headers=h,
               json={"category": "dest", "code": "MN", "name": "蒙古国",
                     "sort": 5, "active": False})
    rows = audit.query(username=admin_name, action=audit.DC_CODE_DISABLE)
    assert any(r["object_id"] == "dest/MN" for r in rows)
    # 停用审计含 before/after（软删除语义可追溯）
    row = next(r for r in rows if r["object_id"] == "dest/MN")
    assert row["before_value"].get("active") is True
    assert row["after_value"].get("active") is False
    db.execute("DELETE FROM train_code_dict WHERE category='dest' AND code='MN'")


def test_code_format_validation_matches_engine():
    """网页/引擎校验一致性（补充任务交付要求3）：维护页直接调用
    train_number.is_valid_code，含连字符/超长/小写一律拒绝。"""
    import train_number
    assert train_number.is_valid_code("PW") is True
    assert train_number.is_valid_code("MZL1") is True
    # 连字符是编号分隔符，绝不能出现在缩写里
    assert train_number.is_valid_code("AB-C") is False
    assert train_number.is_valid_code("工具站") is False   # 非字母数字
    assert train_number.is_valid_code("ABCDEFGHI") is False  # 超8位
    # upsert_code 同口径：非法缩写直接报错
    import pytest
    with pytest.raises(train_store.TrainStoreError):
        train_store.upsert_code("station", "AB-C", "测试", by="t")


def test_disabled_code_not_in_new_trip_options_but_history_intact(dc_users):
    """停用语义（补充任务 §3）：软删除——历史编号不受影响，
    新建班列的编号生成不再接受该缩写。"""
    h = dc_users["headers"]["finance"]
    # 先用 KZ 建一趟班列（历史记录）
    r = client.post("/datacheck/trips", headers=h, json={
        "dep_date": "2025-09-26", "station_code": "ZD", "port_code": "HGS",
        "dest_code": "ZY", "train_type": "T"})
    assert r.status_code == 200
    historical_no = r.json()["trip_no"]

    # 停用 ZD 发站
    admin_h = dc_users["headers"]["admin"]
    client.put("/datacheck/codes", headers=admin_h,
               json={"category": "station", "code": "ZD", "name": "中鼎",
                     "sort": 3, "active": False})

    # 历史班列详情仍可读（编号未失效）
    detail = client.get(f"/datacheck/trips/{historical_no}", headers=h)
    assert detail.status_code == 200

    # 新建班列：编号生成（load_code_map active_only）不再接受 ZD
    r2 = client.post("/datacheck/trips", headers=h, json={
        "dep_date": "2025-09-27", "station_code": "ZD", "port_code": "HGS",
        "dest_code": "ZY", "train_type": "T"})
    assert r2.status_code == 422
    assert "ZD" in r2.json()["detail"]

    # 恢复启用
    client.put("/datacheck/codes", headers=admin_h,
               json={"category": "station", "code": "ZD", "name": "中鼎",
                     "sort": 3, "active": True})

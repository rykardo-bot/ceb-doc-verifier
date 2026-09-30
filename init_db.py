# -*- coding: utf-8 -*-
"""
数据库初始化脚本：建表（幂等）+ 首次种子账号（可选：一键演示数据）。

容器启动流程（start.sh）在服务拉起前执行本脚本，保证 schema 与初始账号就绪。
重复执行安全；users 表已有账号时不做任何种子操作（不影响现有账密）。

种子账号（仅 users 表为空时创建，密码取环境变量，未设置则生成随机密码
打印一次——之后无法找回，只能由管理员重置）：
  admin    管理员  ADMIN_USERNAME    / ADMIN_INITIAL_PASSWORD
  business 业务    BUSINESS_USERNAME / BUSINESS_INITIAL_PASSWORD
  finance  财务    FINANCE_USERNAME  / FINANCE_INITIAL_PASSWORD

运行：
  python init_db.py            # 建表 + 种子账号
  python init_db.py --demo     # 追加一套演示业务数据（幂等，培训/验收/演示用；
                               # 也可用环境变量 DEMO_SEED=1）
"""

from __future__ import annotations

import json
import os
import secrets
import sys
from pathlib import Path

import audit
import auth_service
import db
import train_store

SEED_ACCOUNTS = (
    # (用户名env, 密码env, 默认用户名, 角色)
    ("ADMIN_USERNAME", "ADMIN_INITIAL_PASSWORD", "admin", "admin"),
    ("BUSINESS_USERNAME", "BUSINESS_INITIAL_PASSWORD", "business", "business"),
    ("FINANCE_USERNAME", "FINANCE_INITIAL_PASSWORD", "finance", "finance"),
)


def seed_users_if_empty() -> list[tuple[str, str, str]]:
    """users 表为空时创建三个初始账号；返回 (用户名, 角色, 是否随机密码)。"""
    created = []
    if db.query("SELECT count(*) AS c FROM users")[0]["c"] > 0:
        return created
    for username_env, password_env, default_name, role in SEED_ACCOUNTS:
        username = os.environ.get(username_env, "").strip() or default_name
        password = os.environ.get(password_env, "").strip()
        random_password = not password
        if random_password:
            password = (secrets.token_urlsafe(6) + "7a")  # 保底含字母+数字
        auth_service.create_user(username, password, role)
        created.append((username, role, random_password))
        if random_password:
            print(f"[init_db] 初始账号 {username}（{role}）随机密码：{password}"
                  f" —— 仅此一次显示，请立即记录并尽快修改", flush=True)
    return created


def seed_demo_data() -> None:
    """一键演示数据（--demo / DEMO_SEED=1；逐项幂等，可重复跑）。

    空库冷启动时每个页面都是"暂无"，培训/验收/演示观感差（验收发现）。
    种子内容：3 列班列（结算+三方补贴，其一三方一致→对账页有✅示例）
    + 1 个资金批次 + 1 个手机演示批次（带典型问题→复查页有完整报告可看）。
    """
    import fund_store
    import mobile_store
    from verification_engine import run_verification

    by = "demo_seed"
    demo_trips = [
        ("2026-08-18", "PW", "MZL", "RU", 50, 25, 0),
        ("2026-08-25", "ZD", "EL", "MSK", 45, 22, 1),
        ("2026-08-26", "DT", "HGS", "ZY", 55, 27, 1),
    ]
    trip_nos = []
    for dep_date, stn, prt, dst, wagons, c40, c20 in demo_trips:
        existing = train_store.list_trips(date_from=dep_date, date_to=dep_date,
                                          station=stn, port=prt, dest=dst)
        if existing:
            trip_nos.append(existing[0]["trip_no"])
            continue
        trip = train_store.create_trip(
            dep_date, stn, prt, dst, "T", by,
            basic={"wagon_count": wagons, "container_40hd": c40,
                   "container_20hd": c20, "goods_name": "陶瓷卫浴具"})
        trip_nos.append(trip["trip_no"])
        total = 1_500_000 + 100_000 * len(trip_nos)
        train_store.upsert_settlement(
            trip["trip_no"],
            {"rail_freight": total - 50_000, "settle_total": total,
             "actual_freight": total, "record_status": "confirmed"}, by)
        # 三方补贴口径一致 → 三方对账页有 ✅ 示例
        train_store.upsert_subsidy(
            trip["trip_no"],
            {"dt_supply_100": total, "ly_advance_100": total,
             "auth_confirm_100": total, "record_status": "confirmed"}, by)
    if not fund_store.list_batches():
        fund_store.create_batch(batch_type="prepay", total_amount=1_500_000,
                                trip_nos=trip_nos[:1],
                                counterparty="演示-联运公司", by=by)

    sample = Path(__file__).parent / "sample_data" / "batch_with_issues.json"
    if sample.exists() and not mobile_store.recent(1):
        docs = json.loads(sample.read_text(encoding="utf-8"))["documents"]
        verification = run_verification(
            {"batch_id": "MB-DEMO-0001", "documents": docs})
        verification["batch_id"] = "MB-DEMO-0001"
        mobile_store.save_batch(docs, verification,
                                source="mobile_app", created_by="business")
    print(f"[init_db] 演示数据就绪：班列 {trip_nos}")


def main() -> None:
    db.init_schema()
    counts = db.query(
        "SELECT (SELECT count(*) FROM users) AS users,"
        " (SELECT count(*) FROM batches) AS batches")[0]
    print(f"[init_db] schema 就绪：users={counts['users']} batches={counts['batches']}")
    created = seed_users_if_empty()
    if created:
        audit.record("system", audit.CREATE_USER, "system", "seed_accounts",
                     after={"users": [u for u, _, _ in created]})
    for username, role, random_pw in created:
        suffix = "（随机密码已打印一次）" if random_pw else "（密码来自环境变量）"
        print(f"[init_db] 种子账号就绪：{username} / {role}{suffix}")
    # 数据核对模块v1：缩写代码字典 + 核对阈值种子（逐行幂等，不覆盖人工修改）
    seeded = train_store.seed_refs_if_absent()
    print(f"[init_db] 数据核对引用数据就绪：字典新增 {seeded['codes']} 行，"
          f"配置新增 {seeded['config']} 行")
    if "--demo" in sys.argv or os.environ.get("DEMO_SEED", "").strip() == "1":
        seed_demo_data()


if __name__ == "__main__":
    main()

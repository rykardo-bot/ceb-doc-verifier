# -*- coding: utf-8 -*-
"""
网页端主入口（webapp 双入口架构）。

结构：
  1. 登录门禁：未登录显示登录页（用户名+密码，内部系统口径）；
  2. 角色化导航：登录后按角色装配页面（st.navigation）——
       business 业务   : 单据核对
       finance 财务    : 数据核对（班列编号+联运结算/补贴对账）
       admin   管理员  : 单据核对 + 数据核对 + 用户管理 + 操作日志
     财务角色对"单据核对"的可见性待业务侧最终确认，本轮默认不可见
     （任务书 §2/§4；调整只需改下方 ROLE_PAGES）；
  3. 首页：两大功能入口卡片（单据核对 / 数据核对）+ 管理员快捷入口。

运行：streamlit run webapp/Home.py --server.port 8501
"""

from __future__ import annotations

import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import streamlit as st

import audit
import auth_service
from webapp import admin_audit, admin_codes, admin_users, data_check, doc_verify
from webapp import session
from webapp.format import fmt_dt
from webapp.styles import APP_CSS, LOGIN_CSS

st.set_page_config(
    page_title="中欧班列单证智能核验 · 内部系统",
    page_icon="🚂",
    layout="wide",
)
st.markdown(APP_CSS, unsafe_allow_html=True)

# 角色 → 可见页面（初版权限矩阵；后续角色/边界调整只改这里）
ROLE_PAGES = {
    "business": [
        ("单据核对", doc_verify.render, "🔍"),
    ],
    "finance": [
        ("数据核对", data_check.render, "📊"),
    ],
    "admin": [
        ("单据核对", doc_verify.render, "🔍"),
        ("数据核对", data_check.render, "📊"),
        ("用户管理", admin_users.render, "👤"),
        ("代码字典", admin_codes.render, "🔤"),
        ("操作日志", admin_audit.render, "📜"),
    ],
}


# ---------------------------------------------------------------- 登录页

def _client_ip() -> str:
    try:
        return (st.context.headers.get("x-forwarded-for")
                or st.context.headers.get("x-real-ip") or "").split(",")[0].strip()
    except Exception:
        return ""


def render_login() -> None:
    """登录页：左右分栏（品牌区 + 登录表单）。

    只改版式与品牌呈现：认证、失败留痕、令牌签发逻辑与原实现完全一致
    （Issue #5 的登录/登出控制逻辑不动）。
    """
    st.markdown(LOGIN_CSS, unsafe_allow_html=True)
    # 侧栏保持可见（不隐藏）：登录前导航本就只有"登录"一项（P2-9 口径，
    # 不暴露内部页面），放上品牌信息避免"侧栏被藏起来不知从哪展开"。
    with st.sidebar:
        st.subheader("🚂 中欧班列单证智能核验系统")
        st.caption("大同陆港供应链管理有限公司 · 内部系统")
        st.caption("登录后按角色显示可用功能。")
    brand_col, form_col = st.columns([6, 5], gap="large")
    with brand_col:
        st.markdown(
            '<div class="ceb-brand-panel"><div>'
            '<div class="co">大同陆港供应链管理有限公司</div>'
            '<div class="sys">🚂 中欧班列单证智能核验系统</div>'
            '<div class="tag">发运前自动核验单证一致性、齐全性与路线合规，'
            '降低边境滞留与退运风险</div>'
            '</div>'
            '<div class="foot">内部业务系统 · 仅限公司内部网络访问 · 全程操作留痕</div>'
            '</div>', unsafe_allow_html=True)
    with form_col:
        st.markdown('<div class="ceb-login-form-anchor"></div>', unsafe_allow_html=True)
        st.markdown("### 🔑 用户登录")
        st.caption("请使用公司分配的账号登录")

        with st.form("login_form", clear_on_submit=False):
            username = st.text_input("用户名", autocomplete="username")
            password = st.text_input("密码", type="password", autocomplete="current-password")
            submitted = st.form_submit_button("登 录", type="primary", use_container_width=True)

        if submitted:
            try:
                user = auth_service.authenticate(username.strip(), password)
            except auth_service.AuthError as exc:
                try:
                    audit.record(username.strip() or "anonymous", audit.LOGIN_FAILED,
                                 "user", username.strip() or None,
                                 detail={"reason": exc.message, "via": "web"},
                                 ip=_client_ip())
                except Exception:
                    pass
                st.error(exc.message)
            else:
                token = auth_service.create_token(user)
                try:
                    audit.record(user["username"], audit.LOGIN, "user", user["username"],
                                 detail={"via": "web"}, ip=_client_ip())
                except Exception:
                    pass
                session.login(user, token["access_token"], token["expires_at"])
                st.rerun()

        st.caption("忘记密码请联系管理员重置。连续登录失败会记录在操作日志中。")


# ---------------------------------------------------------------- 首页（业务模块入口层）

def render_home() -> None:
    """首页 = "您要办理哪项业务？"模块入口层（Issue #6 引导式首层）。

    整张卡片就是可点链接（page_link 渲染为卡片样式），单击直达模块，
    不再有"卡片纯展示、真正跳转是卡片下方小按钮"的两段式。
    """
    user = session.current_user() or {}
    st.title("🏠 请选择要办理的业务")
    st.caption(f"中欧班列单证智能核验 · 内部系统（初级版）　|　"
               f"当前登录：**{user.get('username', '—')}**（{user.get('role_label', '—')}）")

    role = user.get("role", "")
    st.markdown('<div class="ceb-cards-anchor"></div>', unsafe_allow_html=True)
    # 只渲染当前角色可用的模块入口：不可用模块直接不展示，
    # 避免业务员首页出现永远点不动的"死卡片"（验收建议收敛项）
    entries = []
    if role in ("business", "admin"):
        entries.append((
            PAGE_DOC, "🔍", "**单据核对**",
            "上传单证照片或 PDF，AI 自动核对内容是否一致、齐全、路线合规，"
            "出具风险评分与整改材料"))
    if role in ("finance", "admin"):
        entries.append((
            PAGE_DATA, "📊", "**数据核对**",
            "查班列编号、登记班列、维护结算台账、三方对账、双表对账与 "
            "Excel 差异导入"))
    cols = st.columns(len(entries) or 1)
    for col, (page, icon, title, desc) in zip(cols, entries):
        with col:
            st.page_link(page, icon=icon, use_container_width=True,
                         label=f"{title}\n\n{desc}")

    if role == "admin":
        st.divider()
        st.markdown("### 管理员快捷入口")
        with st.container(border=True):
            st.markdown('<div class="ceb-secondary-anchor"></div>', unsafe_allow_html=True)
            c1, c2, c3, _ = st.columns([1, 1, 1, 2])
            with c1:
                st.page_link(PAGE_USERS, label="👤 用户管理", use_container_width=True)
            with c2:
                st.page_link(PAGE_CODES, label="🔤 代码字典", use_container_width=True)
            with c3:
                st.page_link(PAGE_AUDIT, label="📜 操作日志", use_container_width=True)

    st.divider()
    st.caption("本系统仅限公司内部部署使用，不面向公网提供服务（持续维护原则，见 README）。")


# ---------------------------------------------------------------- 组装导航

# 刷新恢复（P0）：logout 后先清 cookie（本轮完整渲染里执行），否则尝试用
# cookie 里的令牌恢复登录——浏览器刷新/重开不再丢会话。
if not session.process_cookie_clear():
    session.restore_from_cookie()

if not session.is_logged_in():
    # 登录前只注册登录页（P2-9）：导航若沿用上一次运行的页面清单，
    # 退出登录后侧边栏会残留"数据核对/用户管理/操作日志"等内部入口。
    login_page = st.Page(render_login, title="登录", icon="🔑",
                         url_path="login", default=True)
    st.navigation([login_page]).run()
    st.stop()

PAGE_HOME = st.Page(render_home, title="首页", icon="🏠", default=True)

role = session.current_role()
PAGE_DOC = st.Page(doc_verify.render, title="单据核对", icon="🔍", url_path="doc-verify")
PAGE_DATA = st.Page(data_check.render, title="数据核对", icon="📊", url_path="data-check")
PAGE_USERS = st.Page(admin_users.render, title="用户管理", icon="👤", url_path="users")
PAGE_CODES = st.Page(admin_codes.render, title="代码字典", icon="🔤", url_path="codes")
PAGE_AUDIT = st.Page(admin_audit.render, title="操作日志", icon="📜", url_path="audit")
ALL_PAGES = {"单据核对": PAGE_DOC, "数据核对": PAGE_DATA,
             "用户管理": PAGE_USERS, "代码字典": PAGE_CODES,
             "操作日志": PAGE_AUDIT}

pages = [PAGE_HOME] + [ALL_PAGES[name] for name, _, _ in ROLE_PAGES.get(role, [])]

# 侧栏：登录信息与登出（导航链接由 st.navigation 自动生成）
with st.sidebar:
    user = session.current_user() or {}
    st.markdown(f"👤 **{user.get('username', '—')}**（{user.get('role_label', '—')}）")
    expires = st.session_state.get(session.TOKEN_EXPIRES_KEY, "")
    if expires:
        st.caption(f"登录有效期至 {fmt_dt(expires)}")
    if st.button("🚪 退出登录", use_container_width=True):
        try:
            audit.record(session.current_username(), audit.LOGOUT, "user",
                         session.current_username(), detail={"via": "web"})
        except Exception:
            pass
        session.logout()
        st.rerun()

nav = st.navigation(pages)
session.sync_token_cookie()   # 登录态下保证浏览器 cookie 与当前令牌一致（刷新恢复用）
nav.run()

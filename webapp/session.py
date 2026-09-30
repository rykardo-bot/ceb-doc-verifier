# -*- coding: utf-8 -*-
"""网页端登录会话（Streamlit session_state 封装）。

网页登录口径：登录页校验用户名密码成功后，把用户信息与JWT令牌写入
session_state——令牌用于网页服务端调用核验API（Bearer），用户信息用于
角色导航与操作留痕。浏览器关闭/会话过期后需重新登录（内部系统可接受）。
"""

from __future__ import annotations

USER_KEY = "ceb_user"
TOKEN_KEY = "ceb_token"
TOKEN_EXPIRES_KEY = "ceb_token_expires"
COOKIE_KEY = "ceb_session_token"     # 浏览器 cookie 名：存当前JWT，刷新后自动恢复登录
CLEAR_FLAG = "ceb_cookie_clear_pending"   # logout 后下一轮渲染清 cookie 的标记


def login(user: dict, token: str, expires_at: str = "") -> None:
    import streamlit as st
    st.session_state[USER_KEY] = user
    st.session_state[TOKEN_KEY] = token
    st.session_state[TOKEN_EXPIRES_KEY] = expires_at


def logout() -> None:
    import streamlit as st
    for key in (USER_KEY, TOKEN_KEY, TOKEN_EXPIRES_KEY):
        st.session_state.pop(key, None)
    # logout 自带 st.rerun（本轮元素不落地），cookie 清除动作放到下一轮完整渲染
    st.session_state[CLEAR_FLAG] = True


def restore_from_cookie() -> None:
    """浏览器刷新/重开后用 cookie 里的 JWT 自动恢复登录态。

    Streamlit 刷新即新 session，此前的登录与向导进度全部丢失（验收 P0）。
    登录后每轮渲染会通过注入脚本把当前 JWT 写入浏览器 cookie；
    新 session 启动时在此读回并校验（过期/禁用即静默放弃），恢复登录。
    """
    import streamlit as st
    if is_logged_in() or st.session_state.get(CLEAR_FLAG):
        return
    import auth_service
    try:
        token = (st.context.cookies or {}).get(COOKIE_KEY)
    except Exception:
        return
    if not token:
        return
    try:
        view = auth_service.verify_token(token)
    except Exception:
        return
    expires_at = ""
    if view.get("token_exp"):
        from datetime import datetime, timezone
        expires_at = datetime.fromtimestamp(view["token_exp"], tz=timezone.utc) \
            .isoformat(timespec="seconds")
    login(view, token, expires_at)


def process_cookie_clear() -> bool:
    """logout 后的下一轮完整渲染里清除浏览器 cookie。返回是否执行了清除。"""
    import streamlit as st
    import streamlit.components.v1 as components
    if not st.session_state.pop(CLEAR_FLAG, None):
        return False
    components.html(
        f"<script>parent.document.cookie = '{COOKIE_KEY}=; path=/; "
        f"max-age=0; SameSite=Strict';</script>", height=0)
    return True


def sync_token_cookie() -> None:
    """登录态下保证浏览器 cookie 与当前令牌一致（不一致才注入写脚本）。

    必须在完整跑完的渲染轮里调用（st.rerun 会丢弃本轮已生成的元素）。
    """
    import streamlit as st
    import streamlit.components.v1 as components
    token = st.session_state.get(TOKEN_KEY, "")
    if not token:
        return
    try:
        cookie = (st.context.cookies or {}).get(COOKIE_KEY, "")
    except Exception:
        return
    if token == cookie:
        return
    components.html(
        f"<script>parent.document.cookie = '{COOKIE_KEY}={token}; path=/; "
        f"max-age=43200; SameSite=Strict';</script>", height=0)


def is_logged_in() -> bool:
    import streamlit as st
    return bool(st.session_state.get(USER_KEY)) and bool(st.session_state.get(TOKEN_KEY))


def current_user() -> dict | None:
    import streamlit as st
    return st.session_state.get(USER_KEY)


def current_username() -> str:
    user = current_user()
    return (user or {}).get("username", "anonymous")


def current_role() -> str:
    user = current_user()
    return (user or {}).get("role", "")


def auth_headers() -> dict:
    """带令牌的请求头（网页服务端 → 核验API）。未登录返回空（降级直连场景）。"""
    import streamlit as st
    token = st.session_state.get(TOKEN_KEY)
    return {"Authorization": f"Bearer {token}"} if token else {}

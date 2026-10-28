# -*- coding: utf-8 -*-
"""
============================================================
 外贸可行性与风险测算系统（本地 Windows + Streamlit 版）
============================================================
 依据：《中小企业外贸可行性与风险测算指标手册 v1.0》

 功能：
   1. 表单分组录入（企业基础 / 产品财务 / 目标市场 / 订单结算）
   2. 自动计算可行度（盈利35% + 准入25% + 交付20% + 结算20%）与七类风险
   3. Plotly 雷达图（四维可行度）+ 风险矩阵散点图（概率 P × 损失 L）
   4. SQLite 本地存储历史记录（trade_data.db，首次测算时自动生成）
   5. 打包 JSON 调用 Coze 扣子智能体，返回 AI 风险报告
   6. reportlab 导出 PDF 报告（默认保存在本项目文件夹内）

 运行方式（Windows PowerShell，在项目文件夹内）：
   .venv\\Scripts\\python.exe -m streamlit run app.py
   浏览器打开 http://localhost:8501

 Coze 密钥配置（不再写死在代码里）：
   在项目下新建/编辑 .streamlit/secrets.toml，填写 [coze] 的 bot_id / pat_token；
   也可用环境变量 COZE_BOT_ID / COZE_PAT_TOKEN 覆盖。

 注意：本文件不含任何 Colab 相关代码，纯本地运行。
============================================================
"""

# ======================= 配置区 =======================
# 读取顺序：.streamlit/secrets.toml → 环境变量 → 未配置
DEFAULT_COZE_API_BASE = "https://api.coze.cn"

# 四维默认权重（页面“指标权重设置”可实时调整，自动归一化）
DEFAULT_DIM_WEIGHTS = {"盈利可行性": 0.35, "市场准入": 0.25, "供应链交付": 0.20, "资金结算": 0.20}
# =====================================================

import os
import json
import sqlite3
import datetime
import streamlit as st
import pandas as pd
import plotly.graph_objects as go

# 行业基准毛利率参考表（单位 %，用于 P1 满分线自动带出，可手动覆盖）
# 说明：手册未给出分行业基准值，此处为可修改的默认参考值
INDUSTRY_GM_BASE = {
    "消费品": 25.0, "机械": 20.0, "化工": 18.0,
    "电子": 18.0, "纺织": 15.0, "食品": 20.0, "其他": 18.0,
}
INDUSTRY_OPTIONS = list(INDUSTRY_GM_BASE.keys())

# 常见目标市场风险等级参考（用于目的国风险等级自动带出，可手动覆盖）
MARKET_RISK = {
    "美国": "中", "欧盟": "低", "英国": "中", "日本": "低", "韩国": "中",
    "加拿大": "低", "澳大利亚": "低", "新加坡": "低", "马来西亚": "中",
    "越南": "中", "泰国": "中", "印尼": "中", "菲律宾": "中",
    "印度": "高", "巴西": "高", "墨西哥": "中", "俄罗斯": "高",
    "沙特阿拉伯": "中", "阿联酋": "中", "尼日利亚": "高", "肯尼亚": "高",
}
MARKET_OPTIONS = ["美国", "欧盟", "英国", "日本", "韩国", "加拿大", "澳大利亚",
                  "新加坡", "马来西亚", "越南", "泰国", "印尼", "菲律宾",
                  "印度", "巴西", "墨西哥", "俄罗斯", "沙特阿拉伯", "阿联酋",
                  "尼日利亚", "肯尼亚", "其他"]

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "trade_data.db")


# ============================================================
# 一、核心计算引擎（按《指标手册 v1.0》逐条实现）
# ============================================================
def clamp100(x):
    """归一化：限制在 0–100"""
    return min(100.0, max(0.0, x))


def fwd(x, z, f):
    """正向指标：x 当前值，z 零分线，f 满分线。S = CLAMP(100×(x−Z)/(F−Z))"""
    if f <= z:
        return 100.0 if x >= f else 0.0
    return clamp100(100.0 * (x - z) / (f - z))


def rev(x, z, f):
    """负向指标：x 当前值，z 零分线（x≥z 记 0），f 满分线（x≤f 记 100）。S = CLAMP(100×(Z−x)/(Z−F))"""
    if z <= f:
        return 100.0 if x <= f else 0.0
    return clamp100(100.0 * (z - x) / (z - f))


def _pct(v):
    """百分数转小数：32 → 0.32"""
    return float(v) / 100.0


def normalize_weights(weights=None):
    """四维权重归一化：容忍百分数/小数混用，负数按 0 处理，合计为 0 时回退默认权重"""
    src = weights or DEFAULT_DIM_WEIGHTS
    vals = {}
    for k in DEFAULT_DIM_WEIGHTS:
        try:
            v = float(src.get(k, 0.0))
        except (TypeError, ValueError):
            v = 0.0
        vals[k] = max(0.0, v)
    total = sum(vals.values())
    if total <= 0:
        return dict(DEFAULT_DIM_WEIGHTS)
    return {k: v / total for k, v in vals.items()}


def weight_signature(weights):
    """权重的可比较指纹，用于判断页面权重是否被改动"""
    return tuple((k, round(v * 100, 1)) for k, v in weights.items())


# 结算方式基准分（F1）
SETTLE_SCORE = {
    "预付 ≥ 50%": 100.0,
    "即期 LC": 85.0,
    "30% 定金 + 见提单副本": 70.0,
    "D/P": 50.0,
    "D/A": 35.0,
    "O/A ≤ 60 天": 20.0,
    "O/A > 90 天": 5.0,
}

# 七类风险定义（名称, 概率字段, 损失字段）
RISK_DEFS = [
    ("汇率风险", "risk_fx_prob", "risk_fx_loss"),
    ("回款风险", "risk_ar_prob", "risk_ar_loss"),
    ("合规风险", "risk_comp_prob", "risk_comp_loss"),
    ("地缘政策风险", "risk_geo_prob", "risk_geo_loss"),
    ("物流风险", "risk_logi_prob", "risk_logi_loss"),
    ("供应链涨价风险", "risk_cost_prob", "risk_cost_loss"),
    ("知识产权法务风险", "risk_ip_prob", "risk_ip_loss"),
]


def risk_level_of(r):
    """风险值 → 等级（区间左闭右开，5.0 归入高风险）"""
    if r < 3:
        return "低"
    if r < 5:
        return "中"
    if r <= 8:
        return "高"
    return "极高"


def calculate(p, weights=None):
    """
    输入：p 为字段字典（字段 ID → 值，百分数按原样录入如 32 表示 32%）
         weights 为四维权重（可选，缺省用 DEFAULT_DIM_WEIGHTS，内部自动归一化）
    输出：四维得分、二级指标明细、可行度总分与等级、七类风险、决策结论
    """
    w = normalize_weights(weights)
    def g(key, default=0.0):
        v = p.get(key, default)
        if v is None:
            return float(default)
        try:
            return float(v)
        except (TypeError, ValueError):
            return float(default)

    flags = []          # 红牌 / 红线提示
    sub = {}            # 二级指标得分明细

    # ---------- 一、盈利可行性（权重 35%） ----------
    gm = _pct(g("profit_gross_margin"))
    gm0 = _pct(g("profit_gm_breakeven"))
    gmt = _pct(g("profit_gm_target"))
    p1 = fwd(gm, gm0, gmt)
    sub["P1 毛利率"] = round(p1, 1)

    nm = _pct(g("profit_net_margin"))
    p2 = fwd(nm, 0.0, 0.08)
    sub["P2 净利率"] = round(p2, 1)
    if nm <= 0:
        flags.append("盈利红线：净利率 ≤ 0，P2 记 0 分")

    fc = g("profit_fixed_cost_order")       # 单订单固定成本分摊（元）
    uc = g("profit_unit_contribution")      # 单件边际贡献（元）
    q = g("profit_expected_qty")            # 预计可获订单量（件）
    if uc > 0 and q > 0:
        r_bep = (fc / uc) / q               # 盈亏平衡比 r = BEP / Q
        p3 = rev(r_bep, 1.2, 0.6)
    else:
        r_bep = None
        p3 = 0.0
    sub["P3 盈亏平衡比"] = round(p3, 1)

    k = _pct(g("profit_price_transfer_ratio")) * _pct(g("profit_price_hike_accept"))
    p4 = fwd(k, 0.0, 0.7)
    sub["P4 成本转嫁系数"] = round(p4, 1)

    cr1 = _pct(g("profit_top1_customer"))
    p5 = rev(cr1, 0.7, 0.2)
    sub["P5 客户集中度"] = round(p5, 1)

    dim_profit = 0.25 * p1 + 0.20 * p2 + 0.20 * p3 + 0.15 * p4 + 0.20 * p5

    # ---------- 二、市场准入（权重 25%） ----------
    n0 = g("access_certs_required")     # 要求资质数
    n1 = g("access_certs_obtained")     # 已取得资质数
    if n0 > 0:
        a1 = 100.0 * n1 / n0
        if str(p.get("access_key_cert_missing", "无")) != "无":
            a1 = min(a1, 60.0)          # 关键资质缺失，总分封顶 60
    else:
        a1 = 100.0
    sub["A1 资质完备率"] = round(a1, 1)

    t = (_pct(g("access_tariff_rate")) + _pct(g("access_vat_rate"))
         + _pct(g("access_antidumping_rate")))
    a2 = rev(t, 0.30, 0.05)
    sub["A2 综合税负率"] = round(a2, 1)

    nt = g("access_std_total")          # 强制检测项总数
    np_ = g("access_std_passed")        # 已通过检测项数
    if nt > 0:
        a3 = fwd(np_ / nt, 0.5, 0.95)
    else:
        a3 = 100.0
    sub["A3 标准符合率"] = round(a3, 1)

    inq = g("access_inquiries")         # 有效询盘数
    dl = g("access_deals")              # 成交订单数
    if inq > 0:
        a4 = fwd(dl / inq, 0.02, 0.15)
    else:
        a4 = 0.0
    sub["A4 询盘转化率"] = round(a4, 1)

    e5 = _pct(g("access_barrier_exposure"))
    a5 = rev(e5, 0.60, 0.05)
    sub["A5 壁垒暴露度"] = round(a5, 1)

    dim_access = 0.30 * a1 + 0.20 * a2 + 0.20 * a3 + 0.15 * a4 + 0.15 * a5

    # ---------- 三、供应链交付（权重 20%） ----------
    if g("supply_target_qty") > 0:
        cap = g("supply_capacity") / g("supply_target_qty")
        s1 = fwd(cap, 0.8, 1.3)
    else:
        cap = None
        s1 = 0.0
    sub["S1 产能保障"] = round(s1, 1)

    otd = _pct(g("supply_otd"))
    s2 = fwd(otd, 0.80, 0.98)
    sub["S2 准时交付率"] = round(s2, 1)

    crs = _pct(g("supply_top1_supplier_share"))
    n_alt = g("supply_alt_suppliers")
    f1 = min(1.0, max(0.0, (0.8 - crs) / 0.7))
    f2 = min(n_alt, 3) / 3.0
    s3 = 100.0 * (0.6 * f1 + 0.4 * f2)
    sub["S3 供应集中度"] = round(s3, 1)

    sigma = _pct(g("supply_material_volatility"))
    s4 = rev(sigma, 0.35, 0.05)
    sub["S4 原料价格波动"] = round(s4, 1)

    fpy = _pct(g("supply_fpy"))
    s5 = fwd(fpy, 0.90, 0.99)
    sub["S5 一次检验合格率"] = round(s5, 1)

    dim_supply = 0.25 * s1 + 0.25 * s2 + 0.20 * s3 + 0.15 * s4 + 0.15 * s5

    # ---------- 四、资金结算（权重 20%） ----------
    f1 = SETTLE_SCORE.get(str(p.get("cash_settlement", "")), 0.0)
    sub["F1 结算方式"] = round(f1, 1)

    dso = g("cash_dso")
    f2 = rev(dso, 120.0, 30.0)
    sub["F2 回款天数DSO"] = round(f2, 1)

    i3 = _pct(g("cash_financing_rate"))
    f3 = rev(i3, 0.15, 0.04)
    sub["F3 资金成本率"] = round(f3, 1)

    h4 = _pct(g("cash_fx_hedge_ratio"))
    f4 = fwd(h4, 0.0, 0.8)
    sub["F4 汇率敞口对冲"] = round(f4, 1)

    r5 = min(_pct(g("cash_doc_match_rate")), _pct(g("cash_refund_on_time")))
    f5 = fwd(r5, 0.80, 0.98)
    if str(p.get("cash_refund_anomaly", "无")) == "有":
        f5 = 0.0
        flags.append("合规红牌：存在退税异常 / 骗税记录，F5 记 0 分")
    sub["F5 退税合规率"] = round(f5, 1)

    dim_cash = 0.30 * f1 + 0.25 * f2 + 0.20 * f3 + 0.15 * f4 + 0.10 * f5

    # ---------- 可行度总分与等级 ----------
    feas = (w["盈利可行性"] * dim_profit + w["市场准入"] * dim_access
            + w["供应链交付"] * dim_supply + w["资金结算"] * dim_cash)
    if feas >= 80:
        feas_level = "高可行"
    elif feas >= 60:
        feas_level = "条件可行"
    elif feas >= 40:
        feas_level = "谨慎可行"
    else:
        feas_level = "不可行"

    # ---------- 七类风险 ----------
    risks = []
    for name, pk, lk in RISK_DEFS:
        rp = g(pk)
        rl = g(lk)
        rv = rp * rl / 10.0
        risks.append({"name": name, "P": rp, "L": rl, "R": round(rv, 1), "level": risk_level_of(rv)})
    risk_max = max(r["R"] for r in risks)
    risk_avg = sum(r["R"] for r in risks) / len(risks)
    dominant = max(risks, key=lambda x: x["R"])
    high_count = sum(1 for r in risks if r["R"] >= 5)
    if risk_max < 3:
        risk_level = "低风险"
    elif risk_max < 5:
        risk_level = "中风险"
    elif risk_max <= 8:
        risk_level = "高风险"
    else:
        risk_level = "极高风险"

    # ---------- 综合决策矩阵 ----------
    if feas >= 80 and risk_max < 5:
        decision = "通过：可签约推进"
    elif feas >= 80 and risk_max >= 5:
        decision = "有条件通过：先落实缓释再签约"
    elif feas >= 60 and risk_max < 5:
        decision = "条件可行：限期补齐短板后推进"
    elif feas >= 60 and risk_max >= 5:
        decision = "暂缓：短板与风险双项整改后重测"
    elif feas >= 40:
        decision = "谨慎试点：单笔小额、不接受长账期"
    else:
        decision = "不可行：退出或更换目标市场"

    return {
        "dims": {"盈利可行性": round(dim_profit, 1), "市场准入": round(dim_access, 1),
                 "供应链交付": round(dim_supply, 1), "资金结算": round(dim_cash, 1)},
        "sub": sub,
        "weights": {k: round(v * 100, 1) for k, v in w.items()},
        "feas_score": round(feas),       # 总分四舍五入取整
        "feas_level": feas_level,
        "risks": risks,
        "risk_max": risk_max,
        "risk_avg": round(risk_avg, 1),
        "risk_level": risk_level,
        "dominant_risk": dominant["name"],
        "high_risk_count": high_count,
        "decision": decision,
        "flags": flags,
        "bep_ratio": r_bep,
        "cap_ratio": cap,
    }


# ============================================================
# 二、SQLite 本地存储
# ============================================================
def init_db():
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS calc_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            created_at TEXT,
            company TEXT,
            market TEXT,
            payload TEXT,
            result TEXT,
            ai_report TEXT
        )
    """)
    conn.commit()
    conn.close()


def insert_record(company, market, payload, result):
    conn = sqlite3.connect(DB_PATH)
    cur = conn.execute(
        "INSERT INTO calc_records (created_at, company, market, payload, result, ai_report) VALUES (?,?,?,?,?,?)",
        (datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), company, market,
         json.dumps(payload, ensure_ascii=False), json.dumps(result, ensure_ascii=False), None),
    )
    conn.commit()
    rid = cur.lastrowid
    conn.close()
    return rid


def update_ai_report(record_id, report):
    conn = sqlite3.connect(DB_PATH)
    conn.execute("UPDATE calc_records SET ai_report=? WHERE id=?", (report, record_id))
    conn.commit()
    conn.close()


def update_record_result(record_id, payload, result):
    """权重调整后同步刷新该条历史记录的输入包与测算结果"""
    conn = sqlite3.connect(DB_PATH)
    conn.execute("UPDATE calc_records SET payload=?, result=? WHERE id=?",
                 (json.dumps(payload, ensure_ascii=False), json.dumps(result, ensure_ascii=False), record_id))
    conn.commit()
    conn.close()


def fetch_history(limit=200):
    conn = sqlite3.connect(DB_PATH)
    rows = conn.execute(
        "SELECT id, created_at, company, market, result, ai_report FROM calc_records ORDER BY id DESC LIMIT ?",
        (limit,),
    ).fetchall()
    conn.close()
    return rows


def fetch_record(record_id):
    conn = sqlite3.connect(DB_PATH)
    row = conn.execute(
        "SELECT id, created_at, company, market, payload, result, ai_report FROM calc_records WHERE id=?",
        (record_id,),
    ).fetchone()
    conn.close()
    return row


# ============================================================
# 三、Coze 智能体调用（cozepy 官方 SDK）
# ============================================================
def get_coze_config():
    """读取 Coze 配置：.streamlit/secrets.toml 优先，其次环境变量，最后为空（未配置）"""
    bot_id, token = "", ""
    base = os.environ.get("COZE_API_BASE", "").strip() or DEFAULT_COZE_API_BASE
    try:
        sec = st.secrets["coze"]
        bot_id = str(sec.get("bot_id", "") or "")
        token = str(sec.get("pat_token", "") or "")
        base = str(sec.get("api_base", "") or "").strip() or base
    except Exception:
        pass
    bot_id = (bot_id or os.environ.get("COZE_BOT_ID", "")).strip()
    token = (token or os.environ.get("COZE_PAT_TOKEN", "")).strip()
    return bot_id, token, base


def call_coze_ai(payload_text):
    """调用 Coze 智能体，返回 AI 风险报告文本"""
    from cozepy import Coze, TokenAuth, Message, ChatEventType

    bot_id, token, base_url = get_coze_config()
    if not bot_id or not token:
        raise ValueError("未配置 Coze 密钥：请在项目下创建 .streamlit/secrets.toml 并填写 [coze] 的 bot_id / pat_token")

    client = Coze(auth=TokenAuth(token=token), base_url=base_url)
    answer = ""
    for event in client.chat.stream(
        bot_id=bot_id,
        user_id="trade_calc_local_user",
        additional_messages=[Message.build_user_question_text(payload_text)],
        auto_save_history=False,
    ):
        if event.event == ChatEventType.CONVERSATION_MESSAGE_DELTA:
            answer += (event.message.content or "")
    if not answer.strip():
        answer = "（Coze 智能体未返回文本内容，请检查 Bot 是否发布并勾选「API 渠道」）"
    return answer


# ============================================================
# 四、PDF 导出（reportlab，含中文字体处理）
# ============================================================
CN_FONT = "Helvetica"


def register_cn_font():
    """注册中文字体，避免 PDF 中文乱码；Windows/Linux 字体都没有时用内置 CJK CID 字体"""
    global CN_FONT
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    candidates = [
        (r"C:\Windows\Fonts\simhei.ttf", None),          # 黑体
        (r"C:\Windows\Fonts\msyh.ttc", 0),               # 微软雅黑
        (r"C:\Windows\Fonts\simfang.ttf", None),         # 仿宋
        (r"C:\Windows\Fonts\simsun.ttc", 0),             # 宋体
        ("/usr/share/fonts/truetype/wqy/wqy-microhei.ttc", 0),               # Linux 文泉驿
        ("/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc", 0),       # Linux Noto CJK
        ("/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc", 0),
    ]
    for path, idx in candidates:
        if os.path.exists(path):
            try:
                if idx is None:
                    pdfmetrics.registerFont(TTFont("CNFont", path))
                else:
                    pdfmetrics.registerFont(TTFont("CNFont", path, subfontIndex=idx))
                CN_FONT = "CNFont"
                return True
            except Exception:
                continue
    # 云端 Linux 找不到中文字体文件时，用 reportlab 内置 CJK CID 字体（不依赖本地文件）
    try:
        from reportlab.pdfbase.cidfonts import UnicodeCIDFont
        pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
        CN_FONT = "STSong-Light"
        return True
    except Exception:
        return False


def export_pdf(payload, result, ai_report, pdf_path):
    """生成 PDF 报告并保存到 pdf_path"""
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.enums import TA_CENTER
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle

    register_cn_font()

    title_style = ParagraphStyle("t", fontName=CN_FONT, fontSize=16, leading=22, alignment=TA_CENTER)
    h_style = ParagraphStyle("h", fontName=CN_FONT, fontSize=12, leading=18, spaceBefore=10, spaceAfter=4)
    b_style = ParagraphStyle("b", fontName=CN_FONT, fontSize=10, leading=15)
    cell_style = ParagraphStyle("c", fontName=CN_FONT, fontSize=9, leading=13)

    def cell(text):
        return Paragraph(str(text), cell_style)

    # 兼容两种数据包结构：{"meta": {...}, "inputs": {...}} 或扁平字段
    meta = payload.get("meta") or {}
    inputs = payload.get("inputs") or {}
    company = meta.get("company") or payload.get("base_company") or inputs.get("base_company") or "未填写"
    market = meta.get("market") or payload.get("base_market") or inputs.get("base_market") or "—"

    story = []
    story.append(Paragraph("外贸可行性与风险测算报告", title_style))
    story.append(Paragraph(
        f"企业：{company}　|　目标市场：{market}　|　"
        f"生成时间：{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", b_style))
    weights_pct = result.get("weights") or {}
    if weights_pct:
        story.append(Paragraph(
            "维度权重：" + " / ".join(f"{k} {float(v):.0f}%" for k, v in weights_pct.items()), b_style))
    story.append(Spacer(1, 8))

    # 结果摘要
    story.append(Paragraph("一、测算结果摘要", h_style))
    summary = Table([
        [cell("可行度总分"), cell(result["feas_score"]), cell("可行度等级"), cell(result["feas_level"])],
        [cell("最高风险值"), cell(result["risk_max"]), cell("风险等级"), cell(result["risk_level"])],
        [cell("平均风险值"), cell(result["risk_avg"]), cell("主导风险"), cell(result["dominant_risk"])],
        [cell("高风险类数"), cell(result["high_risk_count"]), cell("决策结论"), cell(result["decision"])],
    ], colWidths=[90, 70, 90, 190])
    summary.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("FONT", (0, 0), (-1, -1), CN_FONT, 9),
        ("BACKGROUND", (0, 0), (0, -1), colors.whitesmoke),
        ("BACKGROUND", (2, 0), (2, -1), colors.whitesmoke),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    story.append(summary)

    # 四维得分
    story.append(Paragraph("二、四维可行度得分（雷达图口径）", h_style))
    dims = result["dims"]
    dim_tab = Table([[cell(k), cell(v)] for k, v in dims.items()], colWidths=[150, 150])
    dim_tab.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("FONT", (0, 0), (-1, -1), CN_FONT, 9),
        ("BACKGROUND", (0, 0), (0, -1), colors.whitesmoke),
    ]))
    story.append(dim_tab)

    # 二级指标明细
    story.append(Paragraph("三、二级指标得分明细", h_style))
    sub_rows = [[cell("指标"), cell("得分"), cell("指标"), cell("得分")]]
    items = list(result["sub"].items())
    for i in range(0, len(items), 2):
        row = [cell(items[i][0]), cell(items[i][1])]
        if i + 1 < len(items):
            row += [cell(items[i + 1][0]), cell(items[i + 1][1])]
        else:
            row += [cell(""), cell("")]
        sub_rows.append(row)
    sub_tab = Table(sub_rows, colWidths=[120, 60, 120, 60])
    sub_tab.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("FONT", (0, 0), (-1, -1), CN_FONT, 9),
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
    ]))
    story.append(sub_tab)

    # 七类风险
    story.append(Paragraph("四、七类风险（R = P × L / 10）", h_style))
    risk_rows = [[cell("风险类型"), cell("概率 P"), cell("损失 L"), cell("风险值 R"), cell("等级")]]
    for r in result["risks"]:
        risk_rows.append([cell(r["name"]), cell(r["P"]), cell(r["L"]), cell(r["R"]), cell(r["level"])])
    risk_tab = Table(risk_rows, colWidths=[130, 60, 60, 70, 70])
    risk_tab.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("FONT", (0, 0), (-1, -1), CN_FONT, 9),
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
    ]))
    story.append(risk_tab)

    # 红线 / 红牌提示
    if result["flags"]:
        story.append(Paragraph("五、红线 / 红牌提示", h_style))
        for f in result["flags"]:
            story.append(Paragraph("• " + f, b_style))

    # AI 报告
    if ai_report:
        story.append(Paragraph("六、Coze 智能体风险报告", h_style))
        for line in ai_report.replace("\r\n", "\n").split("\n"):
            if line.strip():
                story.append(Paragraph(line.strip(), b_style))

    story.append(Spacer(1, 12))
    story.append(Paragraph("本报告由「外贸可行性与风险测算系统 v1.0」本地自动生成，仅供参考，不构成法律或财务意见。",
                           ParagraphStyle("f", fontName=CN_FONT, fontSize=8, leading=11, textColor=colors.grey)))

    doc = SimpleDocTemplate(pdf_path, pagesize=A4)
    doc.build(story)


# ============================================================
# 五、图表（Plotly）
# ============================================================
def make_radar(dims):
    labels = list(dims.keys())
    values = [dims[k] for k in labels]
    fig = go.Figure(go.Scatterpolar(
        r=values + [values[0]],
        theta=labels + [labels[0]],
        fill="toself",
        line=dict(color="#1f77b4", width=2),
        marker=dict(size=6),
    ))
    fig.update_layout(
        polar=dict(radialaxis=dict(visible=True, range=[0, 100])),
        showlegend=False, height=420, margin=dict(l=60, r=60, t=30, b=30),
    )
    return fig


def make_risk_matrix(risks):
    colors_map = {"低": "#2ca02c", "中": "#ffc107", "高": "#ff7f0e", "极高": "#d62728"}
    fig = go.Figure()
    for r in risks:
        fig.add_trace(go.Scatter(
            x=[r["L"]], y=[r["P"]], mode="markers+text",
            text=[r["name"]], textposition="top center",
            marker=dict(size=10 + r["R"] * 5, color=colors_map[r["level"]],
                        line=dict(color="black", width=0.8)),
            name=r["name"],
            hovertemplate="%{text}<br>P=%{y}<br>L=%{x}<br>R=" + str(r["R"]) + "<extra></extra>",
        ))
    # 等风险参考线 R = 3 / 5 / 8（L = 10R / P）
    for rv, label in ((3, "R=3 低/中分界"), (5, "R=5 中/高分界"), (8, "R=8 高/极高界")):
        xs, ys = [], []
        for i in range(1, 11):
            l = 10.0 * rv / i
            if 0 <= l <= 10:
                xs.append(l)
                ys.append(i)
        fig.add_trace(go.Scatter(
            x=xs, y=ys, mode="lines",
            line=dict(color="grey", width=1, dash="dot"),
            name=label, hoverinfo="skip",
        ))
    fig.update_layout(
        xaxis=dict(title="损失程度 L（1–10）", range=[0, 10], dtick=1),
        yaxis=dict(title="发生概率 P（1–10）", range=[0, 10], dtick=1),
        height=420, margin=dict(l=50, r=30, t=30, b=40),
        legend=dict(orientation="h", y=-0.2),
    )
    return fig


def make_risk_bar(risks):
    colors_map = {"低": "#2ca02c", "中": "#ffc107", "高": "#ff7f0e", "极高": "#d62728"}
    fig = go.Figure(go.Bar(
        x=[r["name"] for r in risks],
        y=[r["R"] for r in risks],
        marker_color=[colors_map[r["level"]] for r in risks],
        text=[str(r["R"]) for r in risks],
        textposition="outside",
    ))
    fig.add_hline(y=3, line_dash="dot", line_color="grey")
    fig.add_hline(y=5, line_dash="dot", line_color="orange")
    fig.add_hline(y=8, line_dash="dot", line_color="red")
    fig.update_layout(
        yaxis=dict(title="风险值 R", range=[0, 10]),
        xaxis=dict(title=""),
        height=380, margin=dict(l=40, r=30, t=30, b=60),
        showlegend=False,
    )
    return fig


# ============================================================
# 六、Streamlit 界面
# ============================================================
def num(label, key, min_v, max_v, step, fmt="%.1f", help=None, default=0.0):
    # 若该 key 已被写入 session_state（如行业自动带出的满分线），不要再传 value，避免告警
    kwargs = {} if key in st.session_state else {"value": default}
    return st.number_input(label, min_value=min_v, max_value=max_v, step=step,
                           format=fmt, key=key, help=help, **kwargs)


def reset_weights():
    """按钮回调：把四维权重恢复默认（回调在控件实例化前执行，可安全写 session_state）"""
    for dim_name, default_ratio in DEFAULT_DIM_WEIGHTS.items():
        st.session_state[f"w_{dim_name}"] = float(round(default_ratio * 100))


def main():
    st.set_page_config(page_title="外贸可行性与风险测算系统", layout="wide")
    init_db()

    # ---------- 侧边栏 ----------
    with st.sidebar:
        st.header("系统信息")
        st.write("依据：《中小企业外贸可行性与风险测算指标手册 v1.0》")
        st.write("本地存储：trade_data.db（自动生成于项目文件夹）")
        bot_id, token, _base = get_coze_config()
        st.write(f"Bot ID：{bot_id or '未配置'}")
        st.write(f"PAT 令牌：{('••••' + token[-4:]) if token else '未配置'}")
        st.caption("密钥来源：.streamlit/secrets.toml 或环境变量")
        st.divider()
        st.caption("使用方法：录入四个区块字段 → 点击【开始测算】→ 可选【生成AI风险报告】→ 可选【导出PDF报告】")

    # ---------- 主区域：标题 ----------
    st.title("外贸可行性与风险测算系统")
    st.caption("中小企业单笔出口订单 / 单一目标市场 · 本地运行 · 指标手册 v1.0 · 全部数据仅存本机")

    # ============ 区块一：企业基础信息 ============
    st.subheader("① 企业基础信息")
    c = st.columns(4)
    base_company = c[0].text_input("企业名称（报告抬头，可留空）", key="base_company")
    industry = c[1].selectbox("行业品类", INDUSTRY_OPTIONS, key="base_industry")
    base_market = c[2].selectbox("目标市场", MARKET_OPTIONS, key="base_market")
    # 行业 → 基准毛利率自动带出（P1 满分线 = 基准 + 5pp）
    if st.session_state.get("_last_industry") != industry:
        st.session_state["_last_industry"] = industry
        st.session_state["profit_gm_target"] = round(INDUSTRY_GM_BASE[industry] + 5.0, 1)
    # 目标市场 → 国别风险自动带出
    if st.session_state.get("_last_market") != base_market:
        st.session_state["_last_market"] = base_market
        if base_market in MARKET_RISK:
            st.session_state["base_country_risk"] = MARKET_RISK[base_market]
    country_risk = c[3].selectbox("目的国风险等级（按国别自动带出，可改）",
                                  ["低", "中", "高"], key="base_country_risk",
                                  help="选择目标市场后自动带出，可手动覆盖")
    c = st.columns(4)
    base_hs_code = c[0].text_input("HS 编码（6–10 位）", key="base_hs_code")
    base_incoterm = c[1].selectbox("贸易术语", ["EXW", "FOB", "CFR", "CIF", "DDP"], key="base_incoterm")
    base_ship_mode = c[2].selectbox("运输方式", ["海运", "空运", "铁路", "公路", "多式联运"], key="base_ship_mode")
    base_new_customer = c[3].selectbox("是否新客户", ["否", "是"], key="base_new_customer")
    c = st.columns(4)
    base_order_value = num("单笔订单金额（USD）", "base_order_value", 0.0, 1e9, 100.0, "%.0f", "所有比率的分母", 0.0)
    base_annual_revenue = num("年营业收入（万元）", "base_annual_revenue", 0.0, 1e9, 100.0, "%.0f", "用于订单占比与损失口径", 0.0)
    base_export_ratio = num("出口收入占比（%）", "base_export_ratio", 0.0, 100.0, 0.5, "%.1f", "汇率与政策暴露度", 0.0)
    base_contract_months = num("合同 / 账期周期（月）", "base_contract_months", 0.0, 120.0, 0.5, "%.1f", "敞口持续时间", 0.0)

    # ============ 区块二：产品财务信息 ============
    st.subheader("② 产品财务信息")
    st.markdown("**2.1 产品与成本（P1–P5）**")
    c = st.columns(4)
    profit_gross_margin = num("毛利率（%）", "profit_gross_margin", 0.0, 100.0, 0.5, "%.1f", "P1 当前值", 0.0)
    profit_gm_target = num("行业基准毛利率 + 5pp（%）", "profit_gm_target", 0.0, 100.0, 0.5, "%.1f",
                           "P1 满分线，选行业后自动带出，可手动覆盖",
                           st.session_state.get("profit_gm_target", 0.0))
    profit_gm_breakeven = num("盈亏临界毛利率（%）", "profit_gm_breakeven", 0.0, 100.0, 0.5, "%.1f", "P1 零分线", 0.0)
    profit_net_margin = num("净利率（%，−50~50）", "profit_net_margin", -50.0, 50.0, 0.5, "%.1f",
                            "P2，含退税与汇兑损益；≤ 0 触发盈利红线", 0.0)
    c = st.columns(4)
    profit_fixed_cost_order = num("单订单固定成本分摊（元）", "profit_fixed_cost_order", 0.0, 1e9, 100.0, "%.0f", "P3 计算 BEP", 0.0)
    profit_unit_contribution = num("单件边际贡献（元）", "profit_unit_contribution", 0.0, 1e9, 1.0, "%.1f", "P3 计算 BEP", 0.0)
    profit_expected_qty = num("预计可获订单量（件）", "profit_expected_qty", 0.0, 1e9, 10.0, "%.0f", "P3 计算盈亏平衡比 r", 0.0)
    profit_price_transfer_ratio = num("可调价成本占比（%）", "profit_price_transfer_ratio", 0.0, 100.0, 0.5, "%.1f", "P4", 0.0)
    c = st.columns(4)
    profit_price_hike_accept = num("历史涨价接受率（%）", "profit_price_hike_accept", 0.0, 100.0, 0.5, "%.1f", "P4", 0.0)
    profit_top1_customer = num("最大客户销售占比（%）", "profit_top1_customer", 0.0, 100.0, 0.5, "%.1f", "P5，CR1", 0.0)

    st.markdown("**2.2 供应与交付（S1–S5）**")
    c = st.columns(4)
    supply_capacity = num("可用产能（件/月）", "supply_capacity", 0.0, 1e9, 10.0, "%.0f", "S1", 0.0)
    supply_target_qty = num("目标订单量（件）", "supply_target_qty", 0.0, 1e9, 10.0, "%.0f", "S1 分母", 0.0)
    supply_otd = num("准时交付率（%）", "supply_otd", 0.0, 100.0, 0.5, "%.1f", "S2", 0.0)
    supply_top1_supplier_share = num("关键物料单一供应商占比（%）", "supply_top1_supplier_share", 0.0, 100.0, 0.5, "%.1f", "S3", 0.0)
    c = st.columns(4)
    supply_alt_suppliers = num("合格替代供应商数（家）", "supply_alt_suppliers", 0.0, 20.0, 1.0, "%.0f", "S3", 0.0)
    supply_material_volatility = num("原材料价格年化波动率（%）", "supply_material_volatility", 0.0, 100.0, 0.5, "%.1f", "S4", 0.0)
    supply_fpy = num("一次检验合格率（%）", "supply_fpy", 0.0, 100.0, 0.5, "%.1f", "S5", 0.0)

    # ============ 区块三：目标出口国家信息 ============
    st.subheader("③ 目标出口国家信息")
    st.markdown("**3.1 市场准入（A1–A5）**")
    c = st.columns(4)
    access_certs_required = num("目标市场要求资质数（项）", "access_certs_required", 0.0, 50.0, 1.0, "%.0f", "A1", 0.0)
    access_certs_obtained = num("已取得资质数（项）", "access_certs_obtained", 0.0, 50.0, 1.0, "%.0f", "A1", 0.0)
    access_key_cert_missing = c[2].selectbox("关键资质缺失项", ["无", "CE", "FDA", "REACH", "UL", "原产地证", "其他"],
                                             key="access_key_cert_missing", help="A1：关键资质缺失则该项封顶 60 分")
    access_tariff_rate = num("进口关税税率（%）", "access_tariff_rate", 0.0, 100.0, 0.5, "%.1f", "A2", 0.0)
    c = st.columns(4)
    access_vat_rate = num("进口增值税率（%）", "access_vat_rate", 0.0, 30.0, 0.5, "%.1f", "A2，取不可抵扣部分", 0.0)
    access_antidumping_rate = num("反倾销 / 附加税率（%）", "access_antidumping_rate", 0.0, 200.0, 0.5, "%.1f", "A2", 0.0)
    access_std_total = num("强制检测项总数（项）", "access_std_total", 0.0, 100.0, 1.0, "%.0f", "A3", 0.0)
    access_std_passed = num("已通过检测项数（项）", "access_std_passed", 0.0, 100.0, 1.0, "%.0f", "A3", 0.0)
    c = st.columns(4)
    access_inquiries = num("有效询盘数（条）", "access_inquiries", 0.0, 1e6, 1.0, "%.0f", "A4", 0.0)
    access_deals = num("成交订单数（单）", "access_deals", 0.0, 1e6, 1.0, "%.0f", "A4", 0.0)
    access_barrier_exposure = num("受壁垒影响出口额占比（%）", "access_barrier_exposure", 0.0, 100.0, 0.5, "%.1f",
                                  "A5：受反倾销 / 配额 / 制裁 / 加征关税影响占比", 0.0)

    # ============ 区块四：订单与结算信息 ============
    st.subheader("④ 订单与结算信息")
    st.markdown("**4.1 结算与资金（F1–F5）**")
    c = st.columns(4)
    cash_settlement = c[0].selectbox("结算方式", list(SETTLE_SCORE.keys()), key="cash_settlement", help="F1")
    cash_dso = num("平均回款天数（天）", "cash_dso", 0.0, 365.0, 1.0, "%.0f", "F2，DSO", 0.0)
    cash_financing_rate = num("综合资金成本率（%，年化）", "cash_financing_rate", 0.0, 50.0, 0.5, "%.1f", "F3", 0.0)
    cash_fx_exposure = num("外币敞口金额（USD）", "cash_fx_exposure", 0.0, 1e9, 100.0, "%.0f", "F4", 0.0)
    c = st.columns(4)
    cash_fx_hedge_ratio = num("汇率对冲覆盖率（%）", "cash_fx_hedge_ratio", 0.0, 100.0, 0.5, "%.1f", "F4", 0.0)
    cash_doc_match_rate = num("退税单证一致率（%）", "cash_doc_match_rate", 0.0, 100.0, 0.5, "%.1f", "F5", 0.0)
    cash_refund_on_time = num("退税按期申报率（%）", "cash_refund_on_time", 0.0, 100.0, 0.5, "%.1f", "F5", 0.0)
    cash_refund_anomaly = c[3].selectbox("是否存在退税异常记录", ["无", "有"], key="cash_refund_anomaly",
                                         help="F5：选择「有」记 0 分并触发合规红牌")

    st.markdown("**4.2 订单风险打分（1–10，用于七类风险 R = P × L / 10）**")
    c = st.columns(4)
    risk_ui = [
        ("汇率风险", "risk_fx_prob", "risk_fx_loss"),
        ("回款风险", "risk_ar_prob", "risk_ar_loss"),
        ("合规风险", "risk_comp_prob", "risk_comp_loss"),
        ("地缘政策风险", "risk_geo_prob", "risk_geo_loss"),
    ]
    for i, (name, pk, lk) in enumerate(risk_ui):
        c[i % 4].markdown(f"**{name}**")
        c[i % 4].number_input("发生概率 P", 1, 10, 1, key=pk)
        c[i % 4].number_input("损失程度 L", 1, 10, 1, key=lk)
    c = st.columns(3)
    risk_ui2 = [
        ("物流风险", "risk_logi_prob", "risk_logi_loss"),
        ("供应链涨价风险", "risk_cost_prob", "risk_cost_loss"),
        ("知识产权法务风险", "risk_ip_prob", "risk_ip_loss"),
    ]
    for i, (name, pk, lk) in enumerate(risk_ui2):
        c[i].markdown(f"**{name}**")
        c[i].number_input("发生概率 P", 1, 10, 1, key=pk)
        c[i].number_input("损失程度 L", 1, 10, 1, key=lk)

    # ---------- 组装输入字典 ----------
    p = {
        "base_company": base_company, "base_industry": industry, "base_market": base_market,
        "base_hs_code": base_hs_code, "base_incoterm": base_incoterm, "base_ship_mode": base_ship_mode,
        "base_new_customer": base_new_customer, "base_order_value": base_order_value,
        "base_annual_revenue": base_annual_revenue, "base_export_ratio": base_export_ratio,
        "base_contract_months": base_contract_months, "base_country_risk": country_risk,
        "profit_gross_margin": profit_gross_margin, "profit_gm_target": profit_gm_target,
        "profit_gm_breakeven": profit_gm_breakeven, "profit_net_margin": profit_net_margin,
        "profit_fixed_cost_order": profit_fixed_cost_order, "profit_unit_contribution": profit_unit_contribution,
        "profit_expected_qty": profit_expected_qty, "profit_price_transfer_ratio": profit_price_transfer_ratio,
        "profit_price_hike_accept": profit_price_hike_accept, "profit_top1_customer": profit_top1_customer,
        "supply_capacity": supply_capacity, "supply_target_qty": supply_target_qty, "supply_otd": supply_otd,
        "supply_top1_supplier_share": supply_top1_supplier_share, "supply_alt_suppliers": supply_alt_suppliers,
        "supply_material_volatility": supply_material_volatility, "supply_fpy": supply_fpy,
        "access_certs_required": access_certs_required, "access_certs_obtained": access_certs_obtained,
        "access_key_cert_missing": access_key_cert_missing, "access_tariff_rate": access_tariff_rate,
        "access_vat_rate": access_vat_rate, "access_antidumping_rate": access_antidumping_rate,
        "access_std_total": access_std_total, "access_std_passed": access_std_passed,
        "access_inquiries": access_inquiries, "access_deals": access_deals,
        "access_barrier_exposure": access_barrier_exposure,
        "cash_settlement": cash_settlement, "cash_dso": cash_dso, "cash_financing_rate": cash_financing_rate,
        "cash_fx_exposure": cash_fx_exposure, "cash_fx_hedge_ratio": cash_fx_hedge_ratio,
        "cash_doc_match_rate": cash_doc_match_rate, "cash_refund_on_time": cash_refund_on_time,
        "cash_refund_anomaly": cash_refund_anomaly,
        "risk_fx_prob": st.session_state["risk_fx_prob"], "risk_fx_loss": st.session_state["risk_fx_loss"],
        "risk_ar_prob": st.session_state["risk_ar_prob"], "risk_ar_loss": st.session_state["risk_ar_loss"],
        "risk_comp_prob": st.session_state["risk_comp_prob"], "risk_comp_loss": st.session_state["risk_comp_loss"],
        "risk_geo_prob": st.session_state["risk_geo_prob"], "risk_geo_loss": st.session_state["risk_geo_loss"],
        "risk_logi_prob": st.session_state["risk_logi_prob"], "risk_logi_loss": st.session_state["risk_logi_loss"],
        "risk_cost_prob": st.session_state["risk_cost_prob"], "risk_cost_loss": st.session_state["risk_cost_loss"],
        "risk_ip_prob": st.session_state["risk_ip_prob"], "risk_ip_loss": st.session_state["risk_ip_loss"],
    }

    # ---------- 指标权重面板（修改后实时重算） ----------
    with st.expander("⚙ 指标权重设置（默认 35 / 25 / 20 / 20，修改后自动重算）"):
        wc = st.columns(4)
        raw_weights = {}
        for i, (dim_name, default_ratio) in enumerate(DEFAULT_DIM_WEIGHTS.items()):
            wkey = f"w_{dim_name}"
            wkwargs = {} if wkey in st.session_state else {"value": float(round(default_ratio * 100))}
            raw_weights[dim_name] = wc[i].number_input(
                f"{dim_name}（%）", min_value=0.0, max_value=100.0, step=1.0,
                format="%.0f", key=wkey, **wkwargs)
        w_total = sum(raw_weights.values())
        st.caption(f"当前合计 {w_total:.0f}%（系统按合计自动归一化；合计为 0 时回退默认权重）")
        st.button("恢复默认权重", on_click=reset_weights)

    weights = normalize_weights(raw_weights)

    # ---------- 操作按钮 ----------
    b1, b2, b3, b4 = st.columns([1, 1, 1, 2])
    btn_calc = b1.button("🖩 开始测算", type="primary", width="stretch")
    btn_ai = b2.button("🤖 生成AI风险报告", width="stretch")
    btn_pdf = b3.button("📄 导出PDF报告", width="stretch")

    # ---------- 校验核心必填 ----------
    missing = []
    if base_order_value <= 0:
        missing.append("单笔订单金额（USD）")
    if base_annual_revenue <= 0:
        missing.append("年营业收入（万元）")
    if supply_target_qty <= 0:
        missing.append("目标订单量")
    if profit_expected_qty <= 0:
        missing.append("预计可获订单量")
    if profit_unit_contribution <= 0:
        missing.append("单件边际贡献")
    if missing:
        st.warning("请先补全以下核心字段（需 > 0）：" + "、".join(missing))

    # ---------- 开始测算 ----------
    if btn_calc:
        if missing:
            st.error("核心字段未补全，无法测算。")
        else:
            result = calculate(p, weights)
            payload = {
                "meta": {
                    "tool": "外贸可行性与风险测算系统 v1.0",
                    "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "company": base_company, "market": base_market,
                    "industry": industry, "hs_code": base_hs_code,
                },
                "weights": result["weights"],
                "inputs": p,
                "results": result,
            }
            rid = insert_record(base_company, base_market, payload, result)
            st.session_state["_last_result"] = result
            st.session_state["_last_payload"] = payload
            st.session_state["_last_record_id"] = rid
            st.session_state["_last_ai_report"] = None
            st.session_state["_last_weights"] = weight_signature(weights)
            st.success(f"测算完成，已保存历史记录（记录号 #{rid}）。")

    # ---------- 结果展示 ----------
    result = st.session_state.get("_last_result")
    payload = st.session_state.get("_last_payload")
    ai_report = st.session_state.get("_last_ai_report")

    # 权重被调整 → 用当前记录的原输入实时重算，并同步该条历史记录
    if result is not None and payload is not None and st.session_state.get("_last_weights") != weight_signature(weights):
        recalc = calculate(payload.get("inputs", {}), weights)
        payload["results"] = recalc
        payload["weights"] = recalc["weights"]
        st.session_state["_last_result"] = recalc
        st.session_state["_last_payload"] = payload
        st.session_state["_last_weights"] = weight_signature(weights)
        rid_live = st.session_state.get("_last_record_id")
        if rid_live:
            update_record_result(rid_live, payload, recalc)
        result = recalc
        st.caption("已按调整后的权重重算，并同步更新当前历史记录。")

    if result is not None:
        st.divider()
        st.subheader("测算结果")
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("可行度总分", result["feas_score"], result["feas_level"])
        m2.metric("最高风险值", result["risk_max"], result["risk_level"])
        m3.metric("平均风险值", result["risk_avg"])
        m4.metric("主导风险", result["dominant_risk"])
        m5.metric("高风险类数", result["high_risk_count"])
        st.info("**决策结论：** " + result["decision"])
        for f in result["flags"]:
            st.error("⚠ " + f)

        g1, g2 = st.columns(2)
        with g1:
            st.markdown("**四维可行度雷达图**")
            st.plotly_chart(make_radar(result["dims"]), width="stretch")
        with g2:
            st.markdown("**风险矩阵散点图（概率 P × 损失 L，气泡大小 = 风险值 R）**")
            st.plotly_chart(make_risk_matrix(result["risks"]), width="stretch")

        st.markdown("**七类风险值一览**")
        st.plotly_chart(make_risk_bar(result["risks"]), width="stretch")

        # 二级指标明细
        sub_df = pd.DataFrame(
            [{"二级指标": k, "得分": v} for k, v in result["sub"].items()],
            columns=["二级指标", "得分"],
        )
        st.markdown("**二级指标得分明细**")
        st.dataframe(sub_df, width="stretch", hide_index=True)

        # 发送给 Coze 的 JSON 数据包预览
        if payload is not None:
            with st.expander("查看发送给 Coze 的 JSON 数据包（预览）"):
                st.json(payload)

    # ---------- 生成 AI 风险报告 ----------
    if btn_ai:
        if result is None:
            st.error("请先点击【开始测算】。")
        else:
            payload_text = (
                "请基于以下外贸测算JSON数据包，输出一份完整的中文风险评估与决策建议报告，"
                "要求包含：①可行度诊断（四维得分解读）②短板指标分析③七类风险缓释建议④签约决策建议。\n\n"
                + json.dumps(payload, ensure_ascii=False, indent=2)
            )
            try:
                with st.spinner("正在调用 Coze 智能体生成报告，请稍候…"):
                    ai_report = call_coze_ai(payload_text)
                st.session_state["_last_ai_report"] = ai_report
                rid = st.session_state.get("_last_record_id")
                if rid:
                    update_ai_report(rid, ai_report)
                st.success("AI 风险报告已生成并保存。")
            except Exception as e:
                st.error(f"调用 Coze 失败：{e}")
                st.error("排查提示：①检查 app.py 顶部 Bot ID / PAT 是否有空格；②回 Coze 确认 Bot 已发布且勾选「API 渠道」。")

    if ai_report:
        st.divider()
        st.subheader("Coze 智能体风险报告")
        st.markdown(ai_report)

    # ---------- 导出 PDF ----------
    if btn_pdf:
        if result is None:
            st.error("请先点击【开始测算】。")
        else:
            ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            safe_name = "".join(ch for ch in (base_company or "外贸订单") if ch not in '\\/:*?"<>|').strip() or "外贸订单"
            pdf_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                    f"外贸测算报告_{safe_name}_{ts}.pdf")
            try:
                with st.spinner("正在生成 PDF…"):
                    export_pdf(payload, result, ai_report, pdf_path)
                with open(pdf_path, "rb") as f:
                    pdf_bytes = f.read()
                st.success(f"PDF 已生成：{pdf_path}")
                st.download_button("⬇ 下载 PDF 报告", data=pdf_bytes,
                                   file_name=os.path.basename(pdf_path), mime="application/pdf")
            except Exception as e:
                st.error(f"PDF 导出失败：{e}")

    # ---------- 历史记录 ----------
    st.divider()
    st.subheader("历史记录（SQLite 本地库）")
    rows = fetch_history()
    if not rows:
        st.caption("暂无记录，首次【开始测算】后自动生成 trade_data.db。")
    else:
        hist_df = pd.DataFrame(
            [{"记录号": r[0], "时间": r[1], "企业": r[2] or "—", "市场": r[3],
              "可行度": json.loads(r[4]).get("feas_score") if r[4] else None,
              "等级": json.loads(r[4]).get("feas_level") if r[4] else None,
              "最高风险": json.loads(r[4]).get("risk_max") if r[4] else None,
              "决策": json.loads(r[4]).get("decision") if r[4] else None,
              "AI报告": "有" if r[5] else "无"} for r in rows],
            columns=["记录号", "时间", "企业", "市场", "可行度", "等级", "最高风险", "决策", "AI报告"],
        )
        st.dataframe(hist_df, width="stretch", hide_index=True)

        sel_id = st.selectbox("选择记录号查看详情", hist_df["记录号"].tolist())
        if sel_id:
            row = fetch_record(sel_id)
            if row:
                # fetch_record 列顺序：0 id | 1 时间 | 2 企业 | 3 市场 | 4 payload | 5 result | 6 ai_report
                r_pay = json.loads(row[4]) if row[4] else {}
                r_res = json.loads(row[5]) if row[5] else {}
                st.markdown(f"**记录 #{row[0]}**　{row[1]}　企业：{row[2] or '—'}　市场：{row[3]}")
                st.write("可行度：", r_res.get("feas_score"), "（", r_res.get("feas_level"), "）　",
                         "最高风险：", r_res.get("risk_max"), "　决策：", r_res.get("decision"))
                with st.expander("查看该记录的完整测算结果"):
                    st.json(r_res)
                with st.expander("查看该记录的输入数据包"):
                    st.json(r_pay)
                if row[6]:
                    with st.expander("查看该记录的 AI 风险报告"):
                        st.markdown(row[6])
                else:
                    st.caption("该记录未生成 AI 报告。")


if __name__ == "__main__":
    main()

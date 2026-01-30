# -*- coding: utf-8 -*-
"""
整合版 Streamlit App（Excel 全局上传记忆 + 两步模型 + 侧边栏导航 + 首页背景图）

本版本修复点（v4）：
1) 首页不再显示“请先在侧边栏上传...”提示（避免干扰展示）
2) 上传一次 Excel 在同一会话内全局可用：导航改为 st.button 触发的无刷新路由（避免 <a> 导致会话丢失）
3) “定投向”里共享上限滑块：修复“第二个开始重置/无法继续设置”问题（移除 on_change 强改 session_state）
4) 当某个系数=1 时，其它滑块不再报错：自动禁用其它滑块（避免 min=max 异常/组件崩溃）
"""

import base64
import math
import hashlib
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple
from io import BytesIO
from pathlib import Path

import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from sklearn.linear_model import LinearRegression
from scipy.optimize import curve_fit


# =========================
# 全局参数
# =========================
BASE_SAFETY_SHARE = 0.4267   # 安全项目（消除设备安全隐患）基础占比系数
SLIDER_CAP = 1.0            # 所有滑块共享总上限
SLIDER_STEP = 0.001         # 滑块步长

HOME_BG_FILENAME = ".data/home_bg.png"  # 首页背景图（放在脚本同目录下）


# =========================
# 统一样式
# =========================
def inject_global_css():
    st.markdown("""
    <style>
      h1, h2, h3 { color: #505163 !important; }

      /* 指标卡片 */
      div[data-testid="stMetric"] {
        background-color: #f0f8f4 !important;
        border-radius: 12px !important;
        padding: 14px !important;
      }

      /* 首页：带背景的“模型演示”大横框 */
      .home-wrap{
  width:100%;
  display:flex;
  flex-direction:column;
  align-items:center;
  gap:0px;
  margin-top:0px;
}
      .home-slide{
  width: 100%;
  height: 100vh;
  border-radius: 0px;
  border: 0px solid transparent;
  box-shadow: none;
  background-position: center center;
  background-size: cover;
  background-repeat: no-repeat;
  position: relative;
  overflow: hidden;
}
      .home-slide::after{
        content:'';
        position:absolute;
        inset:0;
        background: linear-gradient(180deg, rgba(255,255,255,.00) 0%, rgba(255,255,255,.00) 55%, rgba(255,255,255,.10) 100%);
        pointer-events:none;
      }
      .home-title{
  position:absolute;
  left:50%;
  top: 40%;
  transform:translate(-50%, -50%);
  font-size: 92px;
  font-weight:900;
  letter-spacing:4px;
  color:#505163;
  background: transparent;
  border: none;
  padding: 0;
  border-radius: 0;
  box-shadow: none;
  text-shadow: 0 6px 18px rgba(0,0,0,.18);
}

      /* 关键修复：确保首页也能看到/点击「展开侧边栏」按钮 */
      div[data-testid="collapsedControl"]{
        position: fixed;
        top: 12px;
        left: 12px;
        z-index: 10000 !important;
      }
      div[data-testid="collapsedControl"] button{
        background: rgba(255,255,255,.92) !important;
        border: 2px solid rgba(82,196,26,.55) !important;
        border-radius: 14px !important;
        padding: 6px 10px !important;
        box-shadow: 0 10px 24px rgba(0,0,0,.12) !important;
      }
      button[aria-label="Open sidebar"],
      button[aria-label="Close sidebar"]{
        z-index: 10000 !important;
      }
      .home-slide{ z-index: 1; pointer-events:none; }

      /* 侧边栏导航：大按钮（仅影响 secondary） */
      section[data-testid="stSidebar"] button[kind="secondary"]{
        padding: 14px 14px !important;
        border-radius: 14px !important;
        border: 2px solid rgba(82,196,26,.55) !important;
        background: #ffffff !important;
        color: #505163 !important;
        font-weight: 900 !important;
        font-size: 20px !important;
        box-shadow: 0 8px 18px rgba(0,0,0,.06) !important;
        transition: all .12s ease-in-out;
      }
      section[data-testid="stSidebar"] button[kind="secondary"]:hover{
        transform: translateY(-1px);
        border-color: #389e0d !important;
        box-shadow: 0 12px 26px rgba(0,0,0,.08) !important;
      }

      /* 主按钮 */
      button[data-testid="baseButton-primary"]{
        background-color:#52c41a !important;
        color:#fff !important;
        border:none !important;
        border-radius:10px !important;
        font-weight:800 !important;
      }
      button[data-testid="baseButton-primary"]:hover{
        background-color:#389e0d !important;
      }

      /* 滑块绿色系 */
      div[data-testid="stSlider"] > div > div > div[role="slider"] > div:first-child {
          background-color: #e6f7ef !important;
      }
      div[data-testid="stSlider"] > div > div > div[role="slider"] > div:nth-child(2) {
          background-color: #52c41a !important;
      }
      div[data-testid="stSlider"] > div > div > div[role="slider"] > div[style*="position: absolute"] {
          background-color: #389e0d !important;
          border: 2px solid #136f22 !important;
      }
      div[data-testid="stSlider"] > div > div > div[role="slider"] > div[style*="position: absolute"]:hover {
          background-color: #237804 !important;
          border: 2px solid #0a4706 !important;
      }
    </style>
    """, unsafe_allow_html=True)

    plt.rcParams["font.family"] = ["SimHei", "WenQuanYi Micro Hei", "Heiti TC", "PingFang SC", "Microsoft YaHei"]
    plt.rcParams["axes.unicode_minus"] = False


# =========================
# Query Params 路由（支持浏览器返回）
# =========================
def _get_query_params() -> Dict[str, str]:
    try:
        qp = st.query_params
        out = {}
        for k in qp.keys():
            v = qp.get_all(k)
            out[k] = v[0] if isinstance(v, list) and len(v) else (v if isinstance(v, str) else "")
        return out
    except Exception:
        qp = st.experimental_get_query_params()
        return {k: (v[0] if isinstance(v, list) and len(v) else "") for k, v in qp.items()}


def _set_query_params(**kwargs):
    try:
        st.query_params.clear()
        for k, v in kwargs.items():
            if v is None:
                continue
            st.query_params[k] = str(v)
    except Exception:
        st.experimental_set_query_params(**{k: v for k, v in kwargs.items() if v is not None})


def goto(page: str):
    _set_query_params(page=page)
    st.rerun()


# =========================
# Excel 读取（统一入口）
# =========================
def _read_excel_from_bytes(file_bytes: bytes) -> pd.DataFrame:
    bio = BytesIO(file_bytes)
    df = pd.read_excel(bio)
    df = df.loc[:, ~df.columns.astype(str).str.contains('Unnamed')]
    return df


# =========================
# 首页背景图：读取脚本同目录的 home_bg.png -> base64
# =========================
@st.cache_data
def load_home_bg_b64() -> Optional[str]:
    try:
        base_dir = Path(__file__).resolve().parent
    except Exception:
        base_dir = Path(".").resolve()

    cand = [
        base_dir / HOME_BG_FILENAME,
        Path(".").resolve() / HOME_BG_FILENAME,
    ]
    for p in cand:
        if p.exists() and p.is_file():
            raw = p.read_bytes()
            return base64.b64encode(raw).decode("utf-8")
    return None


# =========================
# Step 1：总额预测（对数回归）
# =========================
@st.cache_data
def load_step1_data_and_fit(file_bytes: bytes) -> Tuple[pd.DataFrame, float, float]:
    df = _read_excel_from_bytes(file_bytes)
    df = df.iloc[6:12].reset_index(drop=True)

    required_cols = {"时间", "供电可靠率", "投入合计"}
    missing = required_cols - set(df.columns)
    if missing:
        raise ValueError(f"Step1 缺少必要字段：{sorted(list(missing))}")

    years = [2021, 2022, 2023, 2024]
    dates = [f"{y}/12/30" for y in years]

    data = pd.DataFrame({
        "年份": years,
        "供电可靠率": [float(df[df["时间"] == d]["供电可靠率"].values[0]) for d in dates],
        "投资金额": [float(df[df["时间"] == d]["投入合计"].values[0]) / 10000 for d in dates],  # 万元
    })

    X = np.log(data["投资金额"].values.reshape(-1, 1))
    y = data["供电可靠率"].values
    model = LinearRegression()
    model.fit(X, y)
    a = float(model.intercept_)
    b = float(model.coef_[0])
    return data, a, b


def render_step1(file_bytes: Optional[bytes]):
    st.subheader("第一步测总额：统计分析2021-2024年历史项目投入与产出指标样本数据，构建对数回归模型，定义项目投入规模与供电可靠率指标的量化函数关系。")
    st.divider()

    if not file_bytes:
        st.warning("请先在侧边栏上传 Excel 数据文件。")
        return

    try:
        data, a, b = load_step1_data_and_fit(file_bytes)
    except Exception as e:
        st.error(f"读取数据/拟合失败：{e}")
        return

    investment_2024 = float(data.loc[data["年份"] == 2024, "投资金额"].values[0])
    reliability_2024 = float(data.loc[data["年份"] == 2024, "供电可靠率"].values[0])

    col1, col2 = st.columns([1, 2])

    with col1:
        st.subheader("预测参数设置")
        target_Y = st.number_input(
            "供电可靠率目标值（%）",
            min_value=99.90,
            max_value=100.00,
            value=st.session_state.get("step1_target_Y", None),
            step=0.001,
            format="%.4f",
            key="step1_target_Y",
        )

        if target_Y is not None:
            if abs(b) < 1e-12:
                st.error("模型系数 b≈0，无法反推投资金额。")
                return

            predicted_investment = float(np.exp((target_Y - a) / b))  # 万元
            investment_diff = predicted_investment - investment_2024
            reliability_diff = float(target_Y - reliability_2024)

            st.subheader("预测结果")
            m1, m2 = st.columns(2)
            with m1:
                st.metric(
                    "2025年预测累计投资金额",
                    f"{predicted_investment:.0f} 万元",
                    delta=f"较2024年需新增投入\n{investment_diff:.0f} 万元",
                    delta_color="inverse" if investment_diff < 0 else "normal",
                )
            with m2:
                st.metric(
                    "供电可靠率提升幅度",
                    f"{reliability_diff:.4f} %",
                    delta="较2024年提升",
                    delta_color="normal",
                )
            st.subheader("阳新县历史数据（2021-2024年）")
            data_display = data.copy()
            data_display["投资金额"] = data_display["投资金额"].round(0).astype(int)
            data_display["供电可靠率"] = data_display["供电可靠率"].round(4)
            st.dataframe(
                data_display,
                hide_index=True,
                use_container_width=True,
                column_config={
                    "年份": st.column_config.NumberColumn("年份", format="%d年"),
                    "投资金额": st.column_config.NumberColumn("投资金额（万元）"),
                    "供电可靠率": st.column_config.NumberColumn("供电可靠率（%）"),
                },
            )
    with col2:
        if target_Y is None:
            st.info("请在左侧输入目标供电可靠率，右侧将自动生成预测图表。")
        else:
            predicted_investment = float(np.exp((target_Y - a) / b))

            fig, ax = plt.subplots(figsize=(12, 8))

            min_inv = float(min(data["投资金额"]) * 0.8)
            max_inv = float(max(max(data["投资金额"]) * 1.3, predicted_investment * 1.2))
            investment_range = np.linspace(min_inv, max_inv, 100)
            Y_pred = a + b * np.log(investment_range)

            ax.plot(investment_range, Y_pred, linewidth=3, label="对数回归拟合线", zorder=3, color='#4AA499')
            ax.scatter(data["投资金额"], data["供电可靠率"], s=120, label="历史数据", zorder=5, linewidth=2)
            ax.scatter(predicted_investment, target_Y, s=500, marker="*", label="2025年预测点", zorder=10, linewidth=1.5, color='orange',edgecolors='#666666')

            for _, row in data.iterrows():
                ax.annotate(
                    f"{int(row['年份'])}年\n{row['投资金额']:.0f}万元",
                    (row["投资金额"], row["供电可靠率"]),
                    xytext=(10, 10),
                    textcoords="offset points",
                    fontsize=10,
                    bbox=dict(boxstyle="round,pad=0.5", fc="white", ec="#52c41a", alpha=0.9),
                    arrowprops=dict(arrowstyle="->", color="#52c41a", lw=1),
                )

            ax.annotate(
                f"2025年预测\n{predicted_investment:.0f}万元",
                (predicted_investment, target_Y),
                xytext=(20, -30),
                textcoords="offset points",
                fontsize=11,
                fontweight="bold",
                bbox=dict(boxstyle="round,pad=0.5", fc="white", ec="red", alpha=0.9),
                arrowprops=dict(arrowstyle="->", color="red", lw=2),
            )

            ax.set_xlabel("投资金额（万元）", fontsize=14, fontweight="bold")
            ax.set_ylabel("供电可靠率（%）", fontsize=14, fontweight="bold")
            ax.legend(fontsize=12, loc="lower right", frameon=True, fancybox=True, shadow=True)
            ax.grid(True, linestyle="--", alpha=0.3, color="gray")
            ax.set_ylim(99.86, 100.00)
            ax.set_xlim(min_inv, max_inv)

            ax.spines["top"].set_visible(False)
            ax.spines["right"].set_visible(False)

            plt.tight_layout()
            st.pyplot(fig)


# =========================
# Step 2：定投向（两层对数拟合 + DP 预算分配）
# =========================
def log_model(x, k, d):
    return k * np.log(x) + d


def fit_mapping(df, mappings, x_label: str):
    results = []
    for (x_col, y_col) in mappings:
        if "投入" in x_label:
            x = df[x_col].values / 10000  # 万元
        else:
            x = df[x_col].values.copy()

        x = np.where(x == 0, 0.0001, x)
        y = df[y_col].values

        mask = x > 0
        x_valid, y_valid = x[mask], y[mask]

        try:
            popt, _ = curve_fit(log_model, x_valid, y_valid, p0=[1.0, float(np.mean(y_valid))], maxfev=10000)
            k, d = float(popt[0]), float(popt[1])

            y_pred = log_model(x_valid, k, d)
            ss_res = float(np.sum((y_valid - y_pred) ** 2))
            ss_tot = float(np.sum((y_valid - np.mean(y_valid)) ** 2))
            r2 = 1 - (ss_res / ss_tot) if ss_tot != 0 else 0.0

            if d >= 0:
                fit_formula = f"{y_col} = {k:.4f}·ln({x_col}) + {d:.4f}"
            else:
                fit_formula = f"{y_col} = {k:.4f}·ln({x_col}) - {abs(d):.4f}"

            results.append({"x变量": x_col, "y变量": y_col, "k": k, "d": d, "模型拟合公式": fit_formula, "R²": r2})
        except Exception as e:
            results.append({"x变量": x_col, "y变量": y_col, "k": None, "d": None, "R²": None, "错误": str(e)})
    return results


@st.cache_data
def fit_log_models(file_bytes: bytes):
    df = _read_excel_from_bytes(file_bytes)
    df = df.iloc[6:12]

    required_cols = {
        "高故障线路改造", "带电作业能力提升", "加强网架结构", "配电自动化", "消除设备安全隐患", "变电站配套送出",
        "绝缘化率", "带电作业率", "N-1通过率", "自动化覆盖率", "中压线路故障率", "电缆化率",
        "供电可靠率"
    }
    missing = required_cols - set(df.columns)
    if missing:
        raise ValueError(f"Step2 缺少必要字段：{sorted(list(missing))}")

    sub_mappings = [
        ("高故障线路改造", "绝缘化率"),
        ("带电作业能力提升", "带电作业率"),
        ("加强网架结构", "N-1通过率"),
        ("配电自动化", "自动化覆盖率"),
        ("消除设备安全隐患", "中压线路故障率"),
        ("变电站配套送出", "电缆化率"),
    ]
    main_mappings = [(x, "供电可靠率") for _, x in sub_mappings]

    sub_results = fit_mapping(df, sub_mappings, "项目投入(万元)")
    main_results = fit_mapping(df, main_mappings, "子指标值")
    return sub_results, main_results


@dataclass
class ProjectConfig:
    name: str
    sub_indicator: str
    min_investment: float
    max_investment: float


def setup_projects(total_budget: float, budget_unit: float, project_bounds: Dict[str, Dict[str, float]]) -> List[ProjectConfig]:
    SUB_INDICATOR_MAP = {
        "高故障线路改造": "绝缘化率",
        "带电作业能力提升": "带电作业率",
        "加强网架结构": "N-1通过率",
        "配电自动化": "自动化覆盖率",
        "消除设备安全隐患": "中压线路故障率",
        "变电站配套送出": "电缆化率",
    }

    total_budget = float(total_budget)
    budget_unit = float(budget_unit)
    if budget_unit <= 0:
        raise ValueError("分配决策单元必须大于0")

    total_budget_eff = math.floor(total_budget / budget_unit + 1e-12) * budget_unit
    total_budget_eff = max(budget_unit, total_budget_eff)

    projects: List[ProjectConfig] = []
    for project_name, sub_indicator in SUB_INDICATOR_MAP.items():
        b = project_bounds[project_name]
        min_ratio = float(b.get("min", 0.0))
        max_ratio = float(b.get("max", 1.0))

        if not (0.0 <= min_ratio <= 1.0 and 0.0 <= max_ratio <= 1.0):
            raise ValueError(f"项目[{project_name}]调整系数必须在[0,1]之间。")
        if min_ratio > max_ratio:
            raise ValueError(f"项目[{project_name}]调整系数 min>max，请检查。")

        min_amt = min_ratio * total_budget_eff
        max_amt = (np.inf if max_ratio >= 1.0 else max_ratio * total_budget_eff)

        projects.append(ProjectConfig(
            name=project_name,
            sub_indicator=sub_indicator,
            min_investment=float(min_amt),
            max_investment=float(max_amt) if np.isfinite(max_amt) else np.inf
        ))
    return projects


def _round_down_to_step(x: float, step: float) -> float:
    if step <= 0:
        return float(x)
    return max(0.0, math.floor((x + 1e-12) / step) * step)


def _snap_to_step(x: float, step: float) -> float:
    if step <= 0:
        return float(x)
    y = round(float(x) / step) * step
    return float(min(1.0, max(0.0, y)))


def calc_dynamic_min_share(project_bounds: Dict[str, Dict[str, float]],
                           base_share: float = BASE_SAFETY_SHARE,
                           safety_project: str = "消除设备安全隐患") -> float:
    sum_other = 0.0
    for name, b in project_bounds.items():
        if name == safety_project:
            continue
        sum_other += float(b.get("min", 0.0))

    min_share = (1.0 - sum_other) * float(base_share)
    return max(0.0, min(float(base_share), float(min_share)))


class BudgetAllocator:
    def __init__(
        self,
        total_budget: float,
        budget_unit: float,
        projects: List[ProjectConfig],
        min_share: float,
        curve_results_sub: List[dict],
        curve_results_main: List[dict],
    ):
        self.total_budget = float(total_budget)
        self.budget_unit = float(budget_unit)
        self.projects = projects
        self.n_projects = len(projects)
        self.curve_results_sub = curve_results_sub
        self.curve_results_main = curve_results_main

        if self.budget_unit <= 0:
            raise ValueError("预算单位必须大于0")

        self.n_budget_units = int(math.floor(self.total_budget / self.budget_unit + 1e-12))
        if self.n_budget_units <= 0:
            raise ValueError("总预算不足以覆盖一个决策单元，请增大总预算或减小决策单元。")
        self.total_budget_eff = self.n_budget_units * self.budget_unit

        self._min_share = float(max(0.0, min(1.0, min_share)))
        self._unit_bounds: Dict[str, Tuple[int, int]] = self._build_feasible_unit_bounds()

    def _build_feasible_unit_bounds(self) -> Dict[str, Tuple[int, int]]:
        total_u = self.n_budget_units
        raw_info = []

        for p in self.projects:
            raw_min = max(0.0, float(p.min_investment) / self.budget_unit)
            ceil_u = int(math.ceil(raw_min - 1e-12))
            ceil_u = max(0, min(ceil_u, total_u))
            penalty = float(ceil_u - raw_min)
            raw_info.append((p.name, raw_min, ceil_u, penalty))

        min_units = {name: ceil_u for (name, _, ceil_u, _) in raw_info}
        sum_min = sum(min_units.values())

        if sum_min > total_u:
            excess = sum_min - total_u
            order = sorted(raw_info, key=lambda t: t[3], reverse=True)

            i = 0
            while excess > 0 and i < len(order):
                name = order[i][0]
                if min_units[name] > 0:
                    min_units[name] -= 1
                    excess -= 1
                i += 1

            if sum(min_units.values()) > total_u:
                raise ValueError("各项目最小投入之和超过总预算。请增大总预算或减小决策单元。")

        bounds = {}
        for p in self.projects:
            if np.isfinite(p.max_investment):
                raw_max = max(0.0, float(p.max_investment) / self.budget_unit)
                mx_u = int(math.floor(raw_max + 1e-12))
            else:
                mx_u = total_u

            mx_u = max(0, min(mx_u, total_u))
            mn_u = min_units[p.name]
            mx_u = max(mx_u, mn_u)
            bounds[p.name] = (mn_u, mx_u)

        return bounds

    def _get_fitted_params(self, results, x_var):
        for r in results:
            if r.get("x变量") == x_var and r.get("k") is not None:
                return float(r["k"]), float(r["d"])
        return None, None

    def _get_project_cfg(self, project_name: str) -> ProjectConfig:
        return next(p for p in self.projects if p.name == project_name)

    def _get_project_unit_bounds(self, project: ProjectConfig) -> Tuple[int, int]:
        return self._unit_bounds[project.name]

    def _calculate_sub_indicator(self, project_name: str, investment: float) -> float:
        if investment <= 0:
            return 0.0

        k, d = self._get_fitted_params(self.curve_results_sub, project_name)
        if k is None:
            return 0.0

        value = k * np.log(investment) + d
        if project_name == "消除设备安全隐患":
            return float(max(0, min(100, 100 - value)))
        return float(max(0, min(100, value)))

    def _calculate_reliability(self, sub_indicator: str, value: float) -> float:
        if value <= 0:
            return 0.0

        k, d = self._get_fitted_params(self.curve_results_main, sub_indicator)
        if k is None:
            return 0.0

        return float(max(0, min(100, k * np.log(value) + d)))

    def _check_safety_constraint(self, safety_amount: float, total_amount: float) -> bool:
        if total_amount <= 0:
            return True
        return (safety_amount / total_amount) >= self._min_share

    def _find_best_project_for_adjustment(self, current_allocation: Dict[str, float], max_adjust: float) -> Optional[str]:
        if max_adjust < self.budget_unit * 0.5:
            return None

        best_project = None
        best_marginal = -1e18

        for project_name, current_amount in current_allocation.items():
            cfg = self._get_project_cfg(project_name)
            _, max_u = self._get_project_unit_bounds(cfg)
            max_amount = max_u * self.budget_unit

            if current_amount + self.budget_unit > max_amount + 1e-9:
                continue

            current_sub = self._calculate_sub_indicator(project_name, current_amount)
            current_rel = self._calculate_reliability(cfg.sub_indicator, current_sub)

            delta = min(self.budget_unit, max_adjust)
            new_amount = current_amount + delta
            new_sub = self._calculate_sub_indicator(project_name, new_amount)
            new_rel = self._calculate_reliability(cfg.sub_indicator, new_sub)

            marginal = (new_rel - current_rel) / delta
            if marginal > best_marginal:
                best_marginal = marginal
                best_project = project_name

        return best_project

    def allocate_budget(self) -> Dict[str, float]:
        safety_project_index = next((i for i, p in enumerate(self.projects) if p.name == "消除设备安全隐患"), -1)
        if safety_project_index == -1:
            raise ValueError("未找到“消除设备安全隐患”项目。")

        dp = np.full((self.n_projects + 1, self.n_budget_units + 1, self.n_budget_units + 1), -np.inf)
        dp[0, 0, 0] = 0.0
        decision = np.zeros((self.n_projects, self.n_budget_units + 1, self.n_budget_units + 1), dtype=int)

        for i in range(1, self.n_projects + 1):
            project = self.projects[i - 1]
            is_safety = (i - 1) == safety_project_index
            p_min_u, p_max_u = self._get_project_unit_bounds(project)

            for b in range(self.n_budget_units + 1):
                for s in range(self.n_budget_units + 1):
                    best_value = -np.inf
                    best_alloc = 0

                    if is_safety:
                        if s > b:
                            continue
                        if not (p_min_u <= s <= p_max_u):
                            continue
                        alloc_range = [s]
                    else:
                        max_alloc = b - s
                        lo = max(0, p_min_u)
                        hi = min(p_max_u, max_alloc)
                        if lo > hi:
                            continue
                        alloc_range = range(lo, hi + 1)

                    for alloc in alloc_range:
                        prev_b = b - alloc
                        prev_s = s - (alloc if is_safety else 0)

                        if prev_s < 0:
                            continue
                        if dp[i - 1, prev_b, prev_s] == -np.inf:
                            continue

                        inv_amount = alloc * self.budget_unit
                        sub_value = self._calculate_sub_indicator(project.name, inv_amount)
                        benefit = self._calculate_reliability(project.sub_indicator, sub_value)
                        total_value = dp[i - 1, prev_b, prev_s] + benefit

                        if total_value > best_value:
                            best_value = total_value
                            best_alloc = alloc

                    if best_value > -np.inf:
                        dp[i, b, s] = best_value
                        decision[i - 1, b, s] = best_alloc

        best_value = -np.inf
        best_b, best_s = 0, 0

        for b in range(self.n_budget_units + 1):
            for s in range(self.n_budget_units + 1):
                val = dp[self.n_projects, b, s]
                if val > best_value:
                    total_amount = b * self.budget_unit
                    safety_amount = s * self.budget_unit
                    if self._check_safety_constraint(safety_amount, total_amount):
                        best_value = val
                        best_b, best_s = b, s

        if best_value == -np.inf:
            raise ValueError("未能找到满足约束条件的分配方案")

        result: Dict[str, float] = {}
        remaining_b, remaining_s = best_b, best_s

        for i in range(self.n_projects - 1, -1, -1):
            alloc = int(decision[i, remaining_b, remaining_s])
            project = self.projects[i]
            amount = alloc * self.budget_unit
            result[project.name] = float(amount)

            remaining_b -= alloc
            if project.name == "消除设备安全隐患":
                remaining_s -= alloc

        allocated = sum(result.values())
        remaining_budget = self.total_budget_eff - allocated

        while remaining_budget >= self.budget_unit * 0.5:
            current_safety_amount = result.get("消除设备安全隐患", 0.0)
            current_total = sum(result.values())

            if self._check_safety_constraint(current_safety_amount, current_total + self.budget_unit):
                best_project = self._find_best_project_for_adjustment(result, remaining_budget)
                if best_project is None:
                    break
            else:
                best_project = "消除设备安全隐患"

            add_amount = min(self.budget_unit, remaining_budget)

            cfg = self._get_project_cfg(best_project)
            _, mx_u = self._get_project_unit_bounds(cfg)
            mx_amt = mx_u * self.budget_unit
            if result[best_project] + add_amount > mx_amt + 1e-9:
                break

            result[best_project] += float(add_amount)
            remaining_budget -= float(add_amount)

        return result


# ================
# Step2：共享上限滑块（稳定版）
# ================
def build_shared_cap_sliders(project_order: List[str], default_bounds: Dict[str, Dict[str, float]]) -> Dict[str, Dict[str, float]]:
    # 初始化（只在 key 不存在时写入），避免“第二个开始被重置”
    for p in project_order:
        key = f"ratio_{p}"
        if key not in st.session_state:
            st.session_state[key] = float(default_bounds[p]["min"])

    # snap 到 step，避免浮点误差造成上限/禁用抖动
    for p in project_order:
        key = f"ratio_{p}"
        st.session_state[key] = _snap_to_step(float(st.session_state.get(key, 0.0)), SLIDER_STEP)

    # 若历史状态 sum>cap，渲染前先收敛到 cap（异常兜底）
    keys = [f"ratio_{p}" for p in project_order]
    total = sum(float(st.session_state[k]) for k in keys)
    if total > SLIDER_CAP + 1e-9:
        excess = total - SLIDER_CAP
        for p in reversed(project_order):
            k = f"ratio_{p}"
            v = float(st.session_state[k])
            dec = min(v, excess)
            st.session_state[k] = _snap_to_step(v - dec, SLIDER_STEP)
            excess -= dec
            if excess <= 1e-9:
                break

    project_bounds: Dict[str, Dict[str, float]] = {}
    total = sum(float(st.session_state[f"ratio_{p}"]) for p in project_order)

    for p in project_order:
        key = f"ratio_{p}"
        cur = float(st.session_state[key])

        sum_others = total - cur
        raw_max = SLIDER_CAP - sum_others
        raw_max = min(1.0, max(0.0, raw_max))
        max_val = _round_down_to_step(raw_max, SLIDER_STEP)

        if max_val < cur:
            max_val = cur

        st.write(f"**{p}**")

        # cap 已满且 cur==0：禁用其它滑块，避免组件报错
        if max_val <= 0.0 and cur <= 0.0:
            st.slider(
                "调整系数",
                min_value=0.0,
                max_value=1.0,
                step=SLIDER_STEP,
                key=key,
                disabled=True,
            )
        else:
            st.slider(
                "调整系数",
                min_value=0.0,
                max_value=float(max_val),
                step=SLIDER_STEP,
                key=key,
            )

        total = sum(float(st.session_state[f"ratio_{pp}"]) for pp in project_order)
        project_bounds[p] = {"min": float(st.session_state[key]), "max": 1.0}

    ratio_sum = sum(float(st.session_state[f"ratio_{p}"]) for p in project_order)
    remaining = max(0.0, SLIDER_CAP - ratio_sum)
    st.caption(f"当前系数合计：{ratio_sum:.3f}，剩余：{remaining:.3f}")

    return project_bounds


def render_step2(file_bytes: Optional[bytes]):
    st.subheader("第二步定投向：在供电可靠率预算分配总额的约束条件下，结合供电可靠率主、子指标的关联关系，从效益最优化出发，运用动态规划模型，制定子指标投资分配方案。")
    st.divider()

    if not file_bytes:
        st.warning("请先在侧边栏上传 Excel 数据文件。")
        return

    try:
        curve_sub, curve_main = fit_log_models(file_bytes)
    except Exception as e:
        st.error(f"读取Excel/拟合失败：{e}")
        return

    if "allocation_result" not in st.session_state:
        st.session_state.allocation_result = None

    col1, col2 = st.columns([1, 2])

    with col1:
        st.subheader("预算分配总额设置")
        total_budget = st.number_input(
            "总预算（万元）",
            min_value=100.0,
            max_value=100000.0,
            value=None,  # 手动输入（不再从「测总额」自动带入）
            step=10.0,
            format="%.0f",
            key="step2_total_budget",
        )

        budget_unit = st.number_input(
            "分配决策单元（万元/单位）",
            min_value=1.0,
            max_value=100.0,
            value=float(st.session_state.get("step2_budget_unit", 10.0)),
            step=1.0,
            format="%.0f",
            key="step2_budget_unit",
        )

        st.divider()
        st.subheader("投入项目调整系数设置")

        default_bounds = {
            "高故障线路改造": {"min": 0.0, "max": 1.0},
            "带电作业能力提升": {"min": 0.0, "max": 1.0},
            "加强网架结构": {"min": 0.0, "max": 1.0},
            "配电自动化": {"min": 0.0, "max": 1.0},
            "消除设备安全隐患": {"min": 0.0, "max": 1.0},
            "变电站配套送出": {"min": 0.0, "max": 1.0},
        }
        project_order = list(default_bounds.keys())
        project_bounds = build_shared_cap_sliders(project_order, default_bounds)

        if st.button("计算最优分配方案", type="primary", use_container_width=True):
            if total_budget is None:
                st.error("请先设置总预算。")
            else:
                with st.spinner("正在执行动态规划预算分配..."):
                    try:
                        projects = setup_projects(float(total_budget), float(budget_unit), project_bounds)

                        dynamic_min_share = calc_dynamic_min_share(
                            project_bounds,
                            base_share=BASE_SAFETY_SHARE,
                            safety_project="消除设备安全隐患",
                        )

                        allocator = BudgetAllocator(
                            total_budget=float(total_budget),
                            budget_unit=float(budget_unit),
                            projects=projects,
                            min_share=float(dynamic_min_share),
                            curve_results_sub=curve_sub,
                            curve_results_main=curve_main,
                        )
                        allocation = allocator.allocate_budget()
                        st.session_state.allocation_result = allocation

                    except Exception as e:
                        st.error(f"计算失败：{e}")

                        # 兜底：按动态安全占比先给安全项目，剩余均分
                        try:
                            dynamic_min_share = calc_dynamic_min_share(
                                project_bounds,
                                base_share=BASE_SAFETY_SHARE,
                                safety_project="消除设备安全隐患",
                            )
                            safety_min_amount = float(total_budget) * float(dynamic_min_share)
                            remaining_budget = float(total_budget) - safety_min_amount
                            other_projects = [p for p in project_order if p != "消除设备安全隐患"]

                            allocation = {"消除设备安全隐患": safety_min_amount}
                            if other_projects:
                                for name in other_projects:
                                    allocation[name] = remaining_budget / len(other_projects)

                            st.session_state.allocation_result = allocation
                            st.info("已启用保障性分配方案")
                        except Exception:
                            st.session_state.allocation_result = None

    with col2:
        st.subheader("预算分配结果")

        if st.session_state.allocation_result is None:
            st.info("请设置基础参数和调整系数，点击【计算最优分配方案】生成结果")
            return

        allocation = st.session_state.allocation_result

        sub_to_project = {
            "N-1通过率": "加强网架结构",
            "电缆化率": "变电站配套送出",
            "绝缘化率": "高故障线路改造",
            "中压线路故障率": "消除设备安全隐患",
            "带电作业率": "带电作业能力提升",
            "自动化覆盖率": "配电自动化",
        }
        project_to_sub = {v: k for k, v in sub_to_project.items()}

        result_data = []
        total_allocated = 0.0
        for project, amount in allocation.items():
            result_data.append({
                "子指标": project_to_sub.get(project, "未知"),
                "项目名称": project,
                "分配金额（万元）": round(float(amount)),
            })
            total_allocated += float(amount)

        result_data.append({
            "子指标": "总计",
            "项目名称": "",
            "分配金额（万元）": f"{round(total_allocated)}",
        })

        df_result = pd.DataFrame(result_data)
        df_total = df_result[df_result["子指标"] == "总计"]
        df_detail = df_result[df_result["子指标"] != "总计"].sort_values(by="分配金额（万元）", ascending=False)
        df_show = pd.concat([df_detail, df_total], ignore_index=True)

        st.dataframe(df_show, use_container_width=True, hide_index=True)

        st.divider()
        col_used, col_total, col_rate = st.columns(3)
        with col_used:
            st.metric("已分配预算", f"{total_allocated:.0f}万元")
        with col_total:
            total_budget_display = st.session_state.get("step2_total_budget", None)
            if total_budget_display is None:
                total_budget_display = total_allocated
            st.metric("总预算", f"{float(total_budget_display):.0f}万元")
        with col_rate:
            tb = float(st.session_state.get("step2_total_budget", total_allocated) or 0.0)
            rate = (total_allocated / tb * 100) if tb > 0 else 0.0
            st.metric("预算使用率", f"{rate:.0f}%")

        st.subheader("预算分配金额可视化")

        plot_data = {k: float(v) for k, v in allocation.items()}
        sorted_pairs = sorted(plot_data.items(), key=lambda x: x[1], reverse=False)
        labels = [p[0] for p in sorted_pairs]
        values = [p[1] for p in sorted_pairs]

        total_v = sum(values) if sum(values) > 0 else 1.0
        percentages = [(v / total_v) * 100 for v in values]

        fig, ax = plt.subplots(figsize=(12, 8))
        colors = plt.cm.tab10(np.linspace(0, 1, len(labels))) if len(labels) else None
        bars = ax.barh(labels, values, color=colors, alpha=0.85, edgecolor="white", linewidth=1)

        max_val = max(values) if values else 1.0
        for bar, value, pct in zip(bars, values, percentages):
            ax.text(
                bar.get_width() + max_val * 0.01,
                bar.get_y() + bar.get_height() / 2,
                # f"{value:.0f}万元 ({pct:.1f}%)",
                f"{value:.0f}万元",
                ha="left", va="center",
                fontsize=10, fontweight="bold",
                color="black",
            )

        ax.set_xlabel("分配金额（万元）", fontsize=14, fontweight="bold")
        ax.set_ylabel("投资项目", fontsize=14, fontweight="bold")
        ax.set_xlim(0, max_val * 1.15)

        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        # ax.grid(axis="x", linestyle="--", alpha=0.3, color="gray")

        plt.tight_layout()
        st.pyplot(fig)


# =========================
# Home：模型演示入口（背景图 + 仅“模型演示”）
# =========================
def render_home():
    # 首页全屏展示：去掉默认容器内边距/顶部空白，让背景铺满主区域
    st.markdown("""
    <style>
      [data-testid='stAppViewContainer'] .main .block-container{padding:0 !important;max-width:100% !important;}
</style>
    """, unsafe_allow_html=True)
    b64 = load_home_bg_b64()
    st.markdown('<div class="home-wrap">', unsafe_allow_html=True)

    if b64:
        st.markdown(
            f'<div class="home-slide" style="background-image:url(\'data:image/png;base64,{b64}\');">'
            f'  <div class="home-title">模型演示</div>'
            f'</div>',
            unsafe_allow_html=True
        )
    else:
        st.markdown(
            '<div class="home-slide" style="background:#ffffff;border:2px solid #52c41a;">'
            '  <div class="home-title">模型演示</div>'
            '</div>',
            unsafe_allow_html=True
        )

    st.markdown('</div>', unsafe_allow_html=True)
    # 按要求：首页不显示“请上传 Excel”提示


# =========================
# Sidebar：全局上传 + 导航（全部在侧边栏）
# =========================
def _hash_bytes(b: bytes) -> str:
    return hashlib.md5(b).hexdigest()


def render_sidebar(page: str):
    with st.sidebar:
        st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)

        # 导航（按钮式路由，避免 <a href> 导致会话丢失/需要重复上传）
        # st.markdown("### 导航")

        def nav_btn(label: str, target: str):
            active = (page == target)
            txt = f"{label}" if active else label
            if st.button(txt, key=f"nav_{target}", use_container_width=True, type="secondary"):
                goto(target)

        nav_btn("模型演示", "home")
        nav_btn("测总额", "step1")
        nav_btn("定投向", "step2")

        st.divider()

        st.subheader("数据文件")
        st.caption("上传一次即可全局共享；如需替换，重新上传即可。")

        uploaded = st.file_uploader(
            "上传 Excel（.xlsx / .xls）",
            type=["xlsx", "xls"],
            key="excel_uploader",
            label_visibility="collapsed",
        )

        if uploaded is not None:
            new_bytes = uploaded.getvalue()
            new_hash = _hash_bytes(new_bytes)

            # 只有文件真的变化才更新/清缓存，避免每次 rerun 重置体验
            if st.session_state.get("excel_hash") != new_hash:
                st.session_state["excel_bytes"] = new_bytes
                st.session_state["excel_name"] = uploaded.name
                st.session_state["excel_hash"] = new_hash

                load_step1_data_and_fit.clear()
                fit_log_models.clear()

                st.success(f"已加载：{uploaded.name}")
            else:
                st.success(f"已加载：{st.session_state.get('excel_name', uploaded.name)}")

        elif st.session_state.get("excel_bytes") is not None:
            st.success(f"已加载：{st.session_state.get('excel_name', '已上传文件')}")

        else:
            st.info("未上传文件。请上传后开始使用。")

        if st.session_state.get("excel_bytes") is not None:
            if st.button("清除已上传文件", use_container_width=True, type="tertiary"):
                st.session_state["excel_bytes"] = None
                st.session_state["excel_name"] = None
                st.session_state["excel_hash"] = None
                load_step1_data_and_fit.clear()
                fit_log_models.clear()
                st.rerun()


# =========================
# App 入口
# =========================
def main():
    st.set_page_config(
        page_title="投入产出模型演示",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    inject_global_css()

    qp = _get_query_params()
    page = qp.get("page", "home")

    # 全局文件记忆
    st.session_state.setdefault("excel_bytes", None)
    st.session_state.setdefault("excel_name", None)
    st.session_state.setdefault("excel_hash", None)

    render_sidebar(page)

    file_bytes = st.session_state.get("excel_bytes", None)

    if page == "home":
        render_home()
    elif page == "step1":
        render_step1(file_bytes)
    elif page == "step2":
        render_step2(file_bytes)
    else:
        _set_query_params(page="home")
        st.rerun()


if __name__ == "__main__":
    main()

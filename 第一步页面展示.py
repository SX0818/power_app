import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from sklearn.linear_model import LinearRegression

st.set_page_config(
    page_title="阳新县投资-可靠率预测系统",
    layout="wide",
    initial_sidebar_state="collapsed"
)

with st.sidebar:
    st.title("文件配置信息")
    st.divider()
    excel_path = "/Users/yuqianxie/Desktop/投入产出黄石资料/投入产出汇报展示版/数据集/xyY.xlsx"
    st.write("**当前读取的Excel文件路径：**")
    st.code(excel_path, language="text")
    # st.caption("如需更换文件，请修改上述路径后重启应用")

# 自定义CSS（统一绿色系、美化按钮/输入框）
st.markdown("""
    <style>
    /* 主按钮样式（绿色） */
    button[data-testid="baseButton-primary"] {
        background-color: #52c41a !important;
        color: white !important;
        border: none !important;
        border-radius: 8px !important;
        font-weight: bold !important;
    }
    button[data-testid="baseButton-primary"]:hover {
        background-color: #389e0d !important;
    }
    /* 输入框标签加粗 */
    div[data-testid="stNumberInput"] label {
        font-weight: bold !important;
        font-size: 12px !important;
    }
    /* 标题样式 */
    h1, h2, h3 {
        color: #505163 !important;
    }
    /* 数据卡片样式 */
    div[data-testid="stMetric"] {
        background-color: #f0f8f4 !important;
        border-radius: 8px !important;
        padding: 12px !important;
    }
    </style>
    """, unsafe_allow_html=True)


plt.rcParams["font.family"] = ["SimHei", "WenQuanYi Micro Hei", "Heiti TC", "PingFang SC", "Microsoft YaHei"]
plt.rcParams["axes.unicode_minus"] = False  # 解决负号显示问题
plt.rcParams['font.sans-serif'] = ["SimHei", "WenQuanYi Micro Hei", "Heiti TC"]


@st.cache_data  
def load_and_preprocess_data():
    # 更新文件路径
    excel_path = "./data/xyY.xlsx"
    df = pd.read_excel(excel_path)
    df = df.loc[:, ~df.columns.str.contains('Unnamed')].iloc[6:12].reset_index(drop=True)
    years = [2021, 2022, 2023, 2024]
    dates = [f"{year}/12/30" for year in years]
    data = pd.DataFrame({
        '年份': years,
        '供电可靠率': [float(df[df['时间'] == date]['供电可靠率'].values[0]) for date in dates],
        '投资金额': [float(df[df['时间'] == date]['投入合计'].values[0]) / 10000 for date in dates]
    })
    # 拟合对数回归模型（Y = a + b*ln(X)）
    X = np.log(data['投资金额'].values.reshape(-1, 1))
    y = data['供电可靠率'].values
    model = LinearRegression()
    model.fit(X, y)
    a = model.intercept_  # 截距
    b = model.coef_[0]    # 系数
    return data, a, b


data, a, b = load_and_preprocess_data()
investment_2024 = data[data['年份'] == 2024]['投资金额'].values[0]
reliability_2024 = data[data['年份'] == 2024]['供电可靠率'].values[0]


# st.title("第一步：测总额")
st.subheader("统计分析2021-2024年历史项目投入与产出指标样本数据，构建对数回归模型，定义项目投入规模与供电可靠率指标的量化函数关系。")
st.divider()

col1, col2 = st.columns([1, 2]) 

with col1:
    # 输入目标供电可靠率
    st.subheader("预测参数设置")
    target_Y = st.number_input(
        "供电可靠率目标值设定（%）",
        min_value=99.90,     
        max_value=100.00,    
        value=None,          
        step=0.001,            
        format="%.4f",        
        help="请输入2025年目标供电可靠率（建议范围：99.90% - 100.00%）"
    )

    # 仅当用户输入目标值后，才计算并显示结果
    if target_Y is not None:
        predicted_investment = np.exp((target_Y - a) / b)  # 万元
        investment_diff = predicted_investment - investment_2024
        reliability_diff = target_Y - reliability_2024

       
        st.subheader("预测结果")
        metric_col1, metric_col2 = st.columns(2)
        with metric_col1:
            st.metric(
                "2025年预测累计投资金额",
                f"{predicted_investment:.0f} 万元",
                delta=f"较2024年需新增投入\n{investment_diff:.0f} 万元",
                delta_color="inverse" if investment_diff < 0 else "normal"
            )
        with metric_col2:
            st.metric(
                "供电可靠率提升幅度",
                f"{reliability_diff:.4f} %",
                delta=f"较2024年提升",
                delta_color="normal"
            )

        # 历史数据表格
        st.subheader("阳新县历史数据（2021-2024年）")
        data_display = data.copy()
        data_display['投资金额'] = data_display['投资金额'].round(0).astype(int)  
        data_display['供电可靠率'] = data_display['供电可靠率'].round(4)          
        st.dataframe(
            data_display,
            hide_index=True,
            use_container_width=True,
            column_config={
                "年份": st.column_config.NumberColumn("年份", format="%d年"),
                "投资金额": st.column_config.NumberColumn("投资金额（万元）"),
                "供电可靠率": st.column_config.NumberColumn("供电可靠率（%）")
            }
        )

        # 模型公式展示（专业标注）
        # st.subheader("模型拟合")
        # st.latex(f"供电可靠率 = {a:.4f} + {b:.4f} \\times \\ln(投资金额)")
        st.caption("注：模型基于2021-2024年历史数据拟合，投资金额单位为万元。")

with col2:
    if target_Y is not None:
        plt.rcParams['figure.dpi'] = 100                             

        fig, ax = plt.subplots(figsize=(12, 8))

        # 绘制拟合曲线
        min_inv = min(data['投资金额']) * 0.8
        max_inv = max(max(data['投资金额']) * 1.3, predicted_investment * 1.2)
        investment_range = np.linspace(min_inv, max_inv, 100)
        Y_pred = a + b * np.log(investment_range)
        ax.plot(
            investment_range, Y_pred,
            color='#52c41a',  
            linewidth=3,
            label='对数回归拟合线',
            zorder=3
        )

        # 2. 绘制历史数据散点
        ax.scatter(
            data['投资金额'], data['供电可靠率'],
            color='blue',  # 蓝色
            s=120,
            label='历史数据',
            zorder=5,
            # edgecolors='black',
            linewidth=2
        )

        # 3. 绘制2025年预测点（突出显示）
        ax.scatter(
            predicted_investment, target_Y,
            color='#ff4d4f',  # 红色
            s=500,
            marker='*',
            label='2025年预测点',
            zorder=10,
            edgecolors='black',
            linewidth=1.5
        )

        # 添加数据标签（历史+预测）
        for i, row in data.iterrows():
            ax.annotate(
                f"{int(row['年份'])}年\n{row['投资金额']:.0f}万元",
                (row['投资金额'], row['供电可靠率']),
                xytext=(10, 10),
                textcoords='offset points',
                fontsize=10,
                bbox=dict(boxstyle="round,pad=0.5", fc="white", ec="gray", alpha=0.9),
                arrowprops=dict(arrowstyle='->', color='gray', lw=1)
            )

        ax.annotate(
            f"2025年预测\n{predicted_investment:.0f}万元",
            (predicted_investment, target_Y),
            xytext=(20, -30),
            textcoords='offset points',
            fontsize=11,
            color='#ff4d4f',
            fontweight='bold',
            bbox=dict(boxstyle="round,pad=0.5", fc="white", ec="#ff4d4f", alpha=0.9),
            arrowprops=dict(arrowstyle='->', color='#ff4d4f', lw=2)
        )

     
        ax.set_xlabel('投资金额（万元）', fontsize=14, fontweight='bold')
        ax.set_ylabel('供电可靠率（%）', fontsize=14, fontweight='bold')
        # ax.set_title('阳新县供电可靠率与投资金额的对数回归关系', fontsize=16, fontweight='bold', pad=20)
        ax.legend(fontsize=12, loc='lower right', frameon=True, fancybox=True, shadow=True)
        ax.grid(True, linestyle='--', alpha=0.3, color='gray')  # 浅色网格
        ax.set_ylim(99.86, 100.00)  # 固定Y轴范围
        ax.set_xlim(min_inv, max_inv)

        # 隐藏顶部和右侧边框
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.spines['left'].set_color('#333333')
        ax.spines['bottom'].set_color('#333333')

        plt.tight_layout()
        st.pyplot(fig)
    else:
        # 未输入目标值时，显示提示
        st.info("请在左侧输入目标供电可靠率，右侧将自动生成预测图表")

# 3. 底部说明
# st.caption("预测说明：本预测基于历史数据的对数回归模型，实际投资需结合项目落地、政策调整等因素综合考量。")

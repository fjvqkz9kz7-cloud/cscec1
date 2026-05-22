"""
业主回款健康度诊断器
====================
单位：中建一局集团建设发展有限公司 财务资金部

功能：导入SAP FBL5N客户行项目数据，自动分析业主回款健康度，
      识别高风险业主，生成催收优先级清单和可视化图表。

使用：部署到Streamlit Cloud后，通过链接 https://xxx.streamlit.app 访问
"""

import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import io
import base64
import os
import sys

# ============================================================
# 动态资源路径（支持本地和云端部署）
# ============================================================
if getattr(sys, 'frozen', False):
    _RESOURCE_BASE = sys._MEIPASS
else:
    _RESOURCE_BASE = os.path.dirname(os.path.abspath(__file__))

ASSETS_DIR = os.path.join(_RESOURCE_BASE, "assets")
EXAMPLE_DIR = os.path.join(_RESOURCE_BASE, "example_data")

LOGO_PATH = os.path.join(ASSETS_DIR, "logo.png")
BANNER_PATH = os.path.join(ASSETS_DIR, "banner.png")

# 将图标转为base64用于HTML内嵌
def img_to_base64(path):
    try:
        with open(path, "rb") as f:
            return base64.b64encode(f.read()).decode()
    except:
        return ""

try:
    logo_b64 = img_to_base64(LOGO_PATH)
    construction_b64 = img_to_base64(os.path.join(ASSETS_DIR, "construction_icon.png"))
except:
    logo_b64 = ""
    construction_b64 = ""

# 导入分析模块
from receivable_analyzer import (
    standardize_fbl5n,
    analyze_fbl5n,
    create_risk_gauge,
    create_owner_comparison,
    create_ar_waterfall,
    create_uncleared_detail,
)

# ============== 页面配置 ==============
st.set_page_config(
    page_title="业主回款健康度诊断器",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============== 自定义样式 ==============
st.markdown("""
<style>
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #003A70 0%, #002850 100%) !important;
    }
    [data-testid="stSidebar"] .stMarkdown { color: white !important; }
    h1 { color: #003A70 !important; font-weight: 700 !important; }
    h2 { color: #003A70 !important; font-weight: 600 !important; border-left: 4px solid #C9A96E !important; padding-left: 12px !important; }
    h3 { color: #004B8D !important; font-weight: 600 !important; }
    .stButton > button {
        background: linear-gradient(135deg, #003A70 0%, #004B8D 100%) !important;
        color: white !important; border: none !important; border-radius: 6px !important;
    }
    .stButton > button:hover {
        background: linear-gradient(135deg, #004B8D 0%, #005CA8 100%) !important;
    }
    .risk-high { background-color: #fdeaea; border-left: 4px solid #e74c3c; padding: 12px; border-radius: 5px; margin: 8px 0; }
    .risk-warning { background-color: #fef5e7; border-left: 4px solid #e67e22; padding: 12px; border-radius: 5px; margin: 8px 0; }
    .risk-watch { background-color: #e8f6f3; border-left: 4px solid #f1c40f; padding: 12px; border-radius: 5px; margin: 8px 0; }
    .risk-normal { background-color: #eafaf1; border-left: 4px solid #27ae60; padding: 12px; border-radius: 5px; margin: 8px 0; }
    .info-box { background-color: #ebf5fb; border: 1px solid #3498db; border-radius: 8px; padding: 12px; margin: 8px 0; }
</style>
""", unsafe_allow_html=True)

# ============== 侧边栏 ==============
with st.sidebar:
    try:
        st.image(LOGO_PATH, width=80)
    except:
        st.markdown("<div style=\"height:80px;\"></div>", unsafe_allow_html=True)
    st.markdown("""
    <div style="text-align: center; margin: 10px 0 20px 0;">
        <div style="font-size: 16px; font-weight: 700; color: white; line-height: 1.4;">
            中建一局集团<br>建设发展有限公司
        </div>
        <div style="font-size: 11px; color: #C9A96E; margin-top: 6px; letter-spacing: 2px;">
            CSCEC · 财务资金部
        </div>
        <div style="width: 40px; height: 2px; background: #C9A96E; margin: 10px auto;"></div>
        <div style="font-size: 13px; color: rgba(255,255,255,0.7);">
            业主回款健康度诊断器
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("---")
    st.markdown("""
    <div style="text-align: center; color: rgba(255,255,255,0.5); font-size: 10px; margin-top: 20px;">
        <div style="color: #C9A96E; margin: 4px 0;">业主回款健康度诊断器</div>
        <div style="margin-top: 8px; color: rgba(255,255,255,0.4);">SAP FBL5N 数据分析工具</div>
    </div>
    """, unsafe_allow_html=True)

# ============== 主页面 ==============
# ============== 顶部Banner ==============
col_banner, col_title = st.columns([1, 2])
with col_banner:
    try:
        st.image(BANNER_PATH, use_column_width=True)
    except:
        st.markdown("<div style=\"height:80px;\"></div>", unsafe_allow_html=True)
with col_title:
    st.markdown("""
    <div style="margin-top: 10px;">
        <div style="font-size: 12px; color: #C9A96E; letter-spacing: 3px; font-weight: 600;">
            CSCEC · 中建一局集团建设发展有限公司
        </div>
        <div style="font-size: 28px; font-weight: 800; color: #003A70; margin: 4px 0;">
            业主回款健康度诊断器
        </div>
        <div style="font-size: 13px; color: #666;">
            SAP FBL5N 数据导入 · 自动分析 · 催收优先级排序 · 可视化图表
        </div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("---")

# ============== 主功能区 ==============
st.markdown(f"""
<div style="display: flex; align-items: center; gap: 12px; margin: 20px 0 10px 0;">
    <img src="data:image/png;base64,{construction_b64}" width="36" style="vertical-align: middle;">
    <span style="font-size: 22px; font-weight: 700; color: #003A70;">业主回款健康度诊断器</span>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div class="info-box">
<b>功能说明：</b>导入SAP FBL5N（客户行项目）导出的Excel数据，自动按客户编码+利润中心+合同维度
识别业主，通过应收总额、收回总额、回款率、未清余额等核心指标分析业主回款健康度，
识别高风险业主，生成催收优先级清单和可视化图表。
<br><b>核心逻辑：</b>FBL5N逐笔明细 → 标准化清洗 → 按业主维度聚合 → 计算回款率与风险等级
→ 催收优先级排序 → 生成诊断报告
</div>
""", unsafe_allow_html=True)
st.markdown(f"""
<div style="display: flex; align-items: center; gap: 12px; margin: 20px 0 10px 0;">
<img src="data:image/png;base64,{construction_b64}" width="36" style="vertical-align: middle;">
<span style="font-size: 22px; font-weight: 700; color: #003A70;">业主回款健康度诊断器</span>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div class="info-box">
<b>功能说明：</b>导入SAP FBL5N（客户行项目）导出的Excel数据，自动按客户编码+利润中心+合同维度
识别业主，通过应收总额、收回总额、回款率、未清余额等核心指标分析业主回款健康度，
识别高风险业主，生成催收优先级清单和可视化图表。
<br><b>核心逻辑：</b>FBL5N逐笔明细 → 标准化清洗 → 按业主维度聚合 → 计算回款率与风险等级
→ 催收优先级排序 → 生成诊断报告
</div>
""", unsafe_allow_html=True)

# ---- 参数设置 ----
col_p1, col_p2 = st.columns(2)
with col_p1:
        default_terms = st.number_input(
            "默认付款条款（天）",
            value=30,
            min_value=0,
            max_value=365,
            help="如果FBL5N数据中没有'到期日'列，将用过账日期+此天数推算到期日"
        )
with col_p2:
        st.info("💡 如需使用真实到期日，请在SAP FBL5N中配置显示'到期日'列后导出")

# ---- 文件上传区 ----
st.subheader("📁 导入SAP FBL5N数据")
uploaded_file = st.file_uploader(
        "上传SAP FBL5N导出的Excel文件（.xlsx/.xls 格式）",
        type=['xlsx', 'xls'],
        help="从SAP FBL5N（客户行项目）事务码导出的明细数据"
)

# 显示数据格式说明
with st.expander("📋 查看期望的数据格式（点击展开）"):
        st.markdown("""
        **数据来源**：SAP → FBL5N（客户行项目）→ 导出为Excel
        
        **必需列**（支持中文/英文列名，系统自动识别）：
        
        | 列名（中文） | 列名（英文） | 说明 | 示例值 |
        |-------------|-------------|------|--------|
        | 已清项目/未清项目符号 | Status symbol | 行状态 | 空 / @5B 等 |
        | 凭证编号 | Document Number | 凭证号 | 5100000001 |
        | 本币金额 | Amount in LC | 金额（元） | 1500000.00 |
        | 科目 | Account | 客户编码 | 0000100001 |
        | 利润中心 | Profit Center | 利润中心 | 1100A001 |
        | 总账科目 | G/L Account | 总账科目 | 1122010100 |
        | 合同 | Contract | 合同号 | HT-2024-001 |
        | 清帐凭证 | Clearing Document | 清账凭证号 | 5200000001 |
        | 清帐日期 | Clearing Date | 清账日期 | 2024-12-31 |
        | 过账日期 | Posting Date | 过账日期 | 2024-06-15 |
        | 到期日 | Net Due Date | 到期日 | 2024-09-15 |
        | 文本 | Text | 摘要 | 工程款结算 |
        | 客户 | Customer | 客户名称 | XX城投公司 |
        
        **处理说明**：
        - 系统自动过滤汇总行（status_symbol含'@'的行）
        - 正金额 = 应收增加（借方），负金额 = 应收减少/收回（贷方）
        - 有清帐凭证/清帐日期的行视为已清项目
        """)

if uploaded_file is not None:
        try:
            # 读取原始数据
            raw_df = pd.read_excel(uploaded_file)
            
            st.success(f"✅ 数据读取成功！原始数据共 **{len(raw_df)}** 行，**{len(raw_df.columns)}** 列")
            
            # 显示原始数据预览
            with st.expander("📋 查看原始数据预览"):
                st.dataframe(raw_df.head(20), use_container_width=True)
                st.caption(f"显示前20行，共{len(raw_df)}行")
            
            # ---- 步骤1：数据标准化 ----
            with st.spinner("🔧 正在标准化数据..."):
                df_std = standardize_fbl5n(raw_df)
            
            if df_std.empty:
                st.error("❌ 标准化后数据为空，请检查上传文件是否为FBL5N格式")
                st.stop()
            
            st.success(f"✅ 标准化完成！有效数据 **{len(df_std)}** 行")
            
            # ---- 步骤2：核心分析 ----
            with st.spinner("🔍 正在分析各业主回款健康度..."):
                result = analyze_fbl5n(df_std, default_payment_terms=int(default_terms))
            
            st.success("✅ 分析完成！")
            
            # 提取结果
            owner_summary = result['owner_summary']
            subject_breakdown = result['subject_breakdown']
            detail = result['detail']
            stats = result['stats']
            
            # ---- 汇总统计卡片 ----
            st.subheader("📈 汇总统计")
            
            col1, col2, col3, col4, col5 = st.columns(5)
            
            with col1:
                st.metric(
                    label="应收总额（万元）",
                    value=f"{stats['total_ar'] / 10000:.2f}"
                )
            with col2:
                st.metric(
                    label="未清余额（万元）",
                    value=f"{stats['uncleared_balance'] / 10000:.2f}"
                )
            with col3:
                st.metric(
                    label="综合回款率（%）",
                    value=f"{stats['collection_rate']:.2f}%"
                )
            with col4:
                st.metric(
                    label="业主数量",
                    value=f"{stats['owner_count']}"
                )
            with col5:
                st.metric(
                    label="高风险业主数",
                    value=f"{stats['high_risk_count']}",
                    delta=f"预警{stats['warning_count']} / 关注{stats['attention_count']} / 正常{stats['normal_count']}",
                    delta_color="inverse"
                )
            
            # ---- 催收优先级清单 ----
            st.subheader("🚨 催收优先级清单")
            
            if not owner_summary.empty:
                # 选择展示列并重命名
                display_cols = [
                    'owner_key', 'account', 'profit_center', 'contract',
                    'ar_total', 'uncleared_balance', 'collection_rate',
                    'max_overdue_days', 'risk_level', 'priority_score'
                ]
                display_df = owner_summary[display_cols].copy()
                
                # 重命名为中文
                display_df.columns = [
                    '业主标识', '客户编码', '利润中心', '合同号',
                    '应收总额(万)', '未清余额(万)', '回款率(%)',
                    '最大逾期天数', '风险等级', '优先级分数'
                ]
                
                # 金额转换为万元
                display_df['应收总额(万)'] = (display_df['应收总额(万)'] / 10000).round(1)
                display_df['未清余额(万)'] = (display_df['未清余额(万)'] / 10000).round(1)
                display_df['回款率(%)'] = display_df['回款率(%)'].round(1)
                display_df['优先级分数'] = display_df['优先级分数'].round(0).astype(int)
                
                # 风险等级颜色标记
                def highlight_risk(val):
                    if val == '高风险':
                        return 'background-color: #fdeaea; color: #c0392b; font-weight: bold'
                    elif val == '预警':
                        return 'background-color: #fef5e7; color: #d68910; font-weight: bold'
                    elif val == '关注':
                        return 'background-color: #fef9e7; color: #b7950b; font-weight: bold'
                    elif val == '正常':
                        return 'background-color: #eafaf1; color: #27ae60'
                    return ''
                
                styled_df = display_df.style.map(highlight_risk, subset=['风险等级'])
                st.dataframe(styled_df, use_container_width=True, height=350)
            else:
                st.info("暂无业主数据")
            
            # ---- 可视化图表区 ----
            st.subheader("📊 可视化分析")
            
            tab1, tab2, tab3, tab4 = st.tabs([
                "📊 业主对比图", "💧 应收瀑布图", "🍕 未清明细", "📅 账龄分析"
            ])
            
            # Tab1: 业主对比图
            with tab1:
                if not owner_summary.empty:
                    fig_comp = create_owner_comparison(owner_summary, top_n=20)
                    st.plotly_chart(fig_comp, use_container_width=True)
                else:
                    st.info("暂无业主数据，无法生成对比图")
            
            # Tab2: 应收瀑布图
            with tab2:
                if not owner_summary.empty:
                    fig_waterfall = create_ar_waterfall(owner_summary)
                    st.plotly_chart(fig_waterfall, use_container_width=True)
                else:
                    st.info("暂无业主数据，无法生成瀑布图")
            
            # Tab3: 未清明细饼图（需选择业主）
            with tab3:
                if not owner_summary.empty and not detail.empty:
                    owner_list = owner_summary['owner_key'].tolist()
                    selected_owner_key = st.selectbox(
                        "选择业主查看未清明细",
                        options=owner_list,
                        format_func=lambda x: f"{x} (未清: {owner_summary[owner_summary['owner_key']==x]['uncleared_balance'].iloc[0]/10000:.2f}万)",
                        key="uncleared_owner_select"
                    )
                    if selected_owner_key:
                        fig_uncleared = create_uncleared_detail(detail, selected_owner_key)
                        st.plotly_chart(fig_uncleared, use_container_width=True)
                else:
                    st.info("暂无数据，无法生成未清明细图")
            
            # Tab4: 账龄分析
            with tab4:
                with st.spinner("🔍 正在计算账龄分析..."):
                    from receivable_analyzer import analyze_aging, create_aging_chart, create_aging_summary_chart
                    aging_result = analyze_aging(detail)
                
                if not aging_result['owner_aging'].empty:
                    # 汇总统计
                    col_a1, col_a2, col_a3, col_a4, col_a5 = st.columns(5)
                    with col_a1:
                        st.metric("0-30天", f"{aging_result['summary'].iloc[0]['amount']/10000:.1f}万")
                    with col_a2:
                        st.metric("31-60天", f"{aging_result['summary'].iloc[1]['amount']/10000:.1f}万")
                    with col_a3:
                        st.metric("61-90天", f"{aging_result['summary'].iloc[2]['amount']/10000:.1f}万")
                    with col_a4:
                        st.metric("91-180天", f"{aging_result['summary'].iloc[3]['amount']/10000:.1f}万")
                    with col_a5:
                        st.metric("180天以上", f"{aging_result['summary'].iloc[4]['amount']/10000:.1f}万")
                    
                    # 全局汇总图
                    fig_aging_summary = create_aging_summary_chart(aging_result['summary'])
                    st.plotly_chart(fig_aging_summary, use_container_width=True)
                    
                    # 业主账龄分布图
                    fig_aging_owner = create_aging_chart(
                        aging_result['owner_aging'],
                        aging_result['aging_labels']
                    )
                    st.plotly_chart(fig_aging_owner, use_container_width=True)
                    
                    # 账龄明细表格
                    with st.expander("📋 查看账龄明细数据"):
                        display_aging = aging_result['owner_aging'].copy()
                        display_aging['amount'] = (display_aging['amount'] / 10000).round(2).astype(str) + " 万"
                        display_aging['percent'] = display_aging['percent'].astype(str) + "%"
                        st.dataframe(display_aging, use_container_width=True)
                else:
                    st.info("📋 数据中没有有效的过账日期，无法进行账龄分析。\n\n" +
                           "💡 解决方法：请在SAP FBL5N中配置显示'过账日期'列后重新导出数据。")
            
            # ---- 诊断报告区 ----
            st.subheader("📋 业主诊断报告")
            
            if not owner_summary.empty:
                # 业主选择器
                owner_list_detail = owner_summary['owner_key'].tolist()
                selected_owner_detail = st.selectbox(
                    "选择业主查看详细诊断报告",
                    options=owner_list_detail,
                    format_func=lambda x: f"{x} (风险: {owner_summary[owner_summary['owner_key']==x]['risk_level'].iloc[0]})",
                    key="detail_owner_select"
                )
                
                if selected_owner_detail:
                    # 获取选中业主的行数据
                    owner_row = owner_summary[owner_summary['owner_key'] == selected_owner_detail].iloc[0]
                    
                    # 计算健康度评分（优先级分数越高越差，所以用100减）
                    priority_score = float(owner_row['priority_score'])
                    health_score = max(0, min(100, 100 - priority_score))
                    
                    col_left, col_right = st.columns([1, 2])
                    
                    with col_left:
                        # 健康度仪表盘
                        fig_gauge = create_risk_gauge(health_score)
                        st.plotly_chart(fig_gauge, use_container_width=True)
                    
                    with col_right:
                        # 风险等级卡片
                        risk = owner_row['risk_level']
                        risk_cls = ""
                        risk_icon = ""
                        if risk == '高风险':
                            risk_cls = "risk-high"
                            risk_icon = "⚠️"
                        elif risk == '预警':
                            risk_cls = "risk-warning"
                            risk_icon = "⚡"
                        elif risk == '关注':
                            risk_cls = "risk-watch"
                            risk_icon = "▶"
                        else:
                            risk_cls = "risk-normal"
                            risk_icon = "✓"
                        
                        st.markdown(f"""
                        <div class="{risk_cls}">
                        <b>{risk_icon} 风险等级：{risk}</b><br>
                        应收总额: {owner_row['ar_total']/10000:.2f}万 | 
                        未清余额: {owner_row['uncleared_balance']/10000:.2f}万 | 
                        回款率: {owner_row['collection_rate']:.2f}% | 
                        优先级: {priority_score:.0f}
                        </div>
                        """, unsafe_allow_html=True)
                        
                        # 诊断报告文字
                        report_text = generate_owner_report(owner_row)
                        st.text_area("诊断报告", report_text, height=250)
            else:
                st.info("暂无业主数据")
            
            # ---- 导出区 ----
            st.subheader("📥 导出报告")
            
            if not owner_summary.empty:
                col_exp1, col_exp2 = st.columns(2)
                
                with col_exp1:
                    # 导出完整Excel报告
                    output_buffer = io.BytesIO()
                    analysis_date = stats.get('analysis_date', datetime.now().strftime("%Y-%m-%d"))
                    
                    with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
                        # Sheet1: 业主汇总
                        export_summary = owner_summary.copy()
                        export_summary.to_excel(writer, sheet_name='业主汇总', index=False)
                        
                        # Sheet2: 总账科目汇总
                        if not subject_breakdown.empty:
                            subject_breakdown.to_excel(writer, sheet_name='科目汇总', index=False)
                        
                        # Sheet3: 逐笔明细
                        if not detail.empty:
                            detail.to_excel(writer, sheet_name='逐笔明细', index=False)
                        
                        # Sheet4: 统计摘要
                        stats_df = pd.DataFrame([stats])
                        stats_df.to_excel(writer, sheet_name='统计摘要', index=False)
                    
                    output_buffer.seek(0)
                    
                    st.download_button(
                        label="📥 下载完整分析报告（Excel）",
                        data=output_buffer.getvalue(),
                        file_name=f"业主回款健康度分析_FBL5N_{analysis_date}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                
                with col_exp2:
                    # 导出催收清单CSV
                    csv_buffer = io.StringIO()
                    owner_summary.to_csv(csv_buffer, index=False, encoding='utf-8-sig')
                    st.download_button(
                        label="📥 下载催收清单（CSV）",
                        data=csv_buffer.getvalue().encode('utf-8-sig'),
                        file_name=f"催收优先级清单_FBL5N_{analysis_date}.csv",
                        mime="text/csv",
                        use_container_width=True
                    )
        
        except Exception as e:
            st.error(f"❌ 分析过程中出现错误: {str(e)}")
            st.info("💡 请检查：1) 上传的文件是否为SAP FBL5N导出的客户行项目数据 2) 是否包含金额列和凭证编号列")

st.markdown("---")
st.markdown(f"""
<div style="background: linear-gradient(135deg, #003A70 0%, #002850 100%);
            padding: 20px; border-radius: 10px; margin-top: 20px; text-align: center;">
    <div style="display: flex; justify-content: center; align-items: center; gap: 15px; margin-bottom: 10px;">
        <img src="data:image/png;base64,{logo_b64}" width="40" style="vertical-align: middle;">
        <div style="text-align: left;">
            <div style="color: white; font-size: 14px; font-weight: 700;">中建一局集团建设发展有限公司</div>
            <div style="color: #C9A96E; font-size: 11px;">CSCEC · 财务资金部</div>
        </div>
    </div>
    <div style="color: rgba(255,255,255,0.5); font-size: 10px; margin-top: 10px;">
        单位：中建一局集团建设发展有限公司 财务资金部 · 财务资金部
    </div>
</div>
""", unsafe_allow_html=True)

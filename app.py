import streamlit as st
import pandas as pd
import io
import os
import openai
from openai import OpenAI
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import matplotlib.pyplot as plt

# ==========================================
# 解决画图中文乱码问题
# ==========================================
plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'Arial Unicode MS']
plt.rcParams['axes.unicode_minus'] = False

# ==========================================
# 🚀 核心大模型双引擎配置 (云端安全保险箱版)
# ==========================================
# 必须从 Streamlit Secrets 安全读取，防止 API Key 被盗用！
VOLC_API_KEY = st.secrets["VOLC_API_KEY"]
VOLC_BASE_URL = "https://ark.cn-beijing.volces.com/api/v3" 
VOLC_MODEL_NAME = "ep-20260503224430-dsq28"

DEEPSEEK_API_KEY = st.secrets["DEEPSEEK_API_KEY"]
DEEPSEEK_BASE_URL = "https://api.deepseek.com/v1"
DEEPSEEK_MODEL_NAME = "deepseek-chat" 

# ==========================================
# 初始化会话状态 
# ==========================================
if 'started' not in st.session_state:
    st.session_state.started = False
if 'processed' not in st.session_state:
    st.session_state.processed = False
if 'selected_model' not in st.session_state:
    st.session_state.selected_model = "豆包大模型 (Volcengine)"

st.set_page_config(page_title="吴江区知识产权简报", layout="wide", page_icon="📈")

# ==========================================
# 💎 Uiverse 高级径向渐变网格 CSS (完美沉底版)
# ==========================================
st.markdown("""
    <style>
    /* 1. 强制透明化可能遮挡网格的默认背景层 */
    .stApp, [data-testid="stAppViewContainer"], .main, [data-testid="stHeader"] {
        background-color: transparent !important;
    }
    
    /* 2. 注入全局浅灰底色 */
    html, body {
        background-color: #f8fafc !important;
    }
    
    /* 3. 核心网格层，使用 z-index: -1 沉底，绝不遮挡文字 */
    .stApp::before {
        content: "";
        position: fixed;
        top: 0;
        right: 0;
        bottom: 0;
        left: 0;
        z-index: -1; 
        background-image: 
            linear-gradient(to right, #e2e8f0 1px, transparent 1px),
            linear-gradient(to bottom, #e2e8f0 1px, transparent 1px);
        background-size: 20px 30px;
        -webkit-mask-image: radial-gradient(ellipse 70% 60% at 50% 0%, #000 60%, transparent 100%);
        mask-image: radial-gradient(ellipse 70% 60% at 50% 0%, #000 60%, transparent 100%);
        pointer-events: none; 
    }

    /* 4. 标题与容器美化，加半透明白底以凸显文字 */
    .flat-title {
        border: 2px solid #e2e8f0; 
        background-color: rgba(255, 255, 255, 0.95); 
        color: #1e293b;
        font-weight: 900; 
        font-size: 20px; 
        text-align: center;
        padding: 15px; 
        border-radius: 8px; 
        margin-bottom: 20px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    }

    /* 5. 按钮立体质感 */
    .stButton > button {
        box-shadow: 0px 4px 0px 0px #94a3b8 !important;
        border-radius: 8px !important;
        transition: all 0.1s !important;
    }
    .stButton > button:active {
        box-shadow: 0px 0px 0px 0px #94a3b8 !important;
        transform: translateY(4px) !important;
    }
    .stButton > button[kind="primary"] {
        background-color: #ef4444 !important;
        box-shadow: 0px 4px 0px 0px #b91c1c !important;
        color: white !important;
    }

    /* 6. 上传区域美化 */
    [data-testid="stFileUploader"] {
        background-color: rgba(255, 255, 255, 0.9);
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 12px;
    }
    
    /* 7. 信息提示框背景优化 */
    .stAlert {
        background-color: rgba(255, 255, 255, 0.85) !important;
        border: 1px solid #e2e8f0 !important;
    }
    </style>
""", unsafe_allow_html=True)

# --- 页面 UI 头部 ---
st.title("📈 吴江区知识产权简报 AI 生成平台")
st.markdown("上传吴江区各板块真实业务清单，Agent 将全自动完成数据透视、公文撰写与多图表大屏排版。")
st.markdown("---")

# ==========================================
# 核心数据处理逻辑函数
# ==========================================
def process_excel_data(f_valid, f_pct, f_auth, f_loss):
    df_valid = pd.read_excel(f_valid)
    df_pct = pd.read_excel(f_pct)
    df_auth = pd.read_excel(f_auth)
    df_loss = pd.read_excel(f_loss)
    
    totals = {'授权总量': len(df_auth), '有效总量': len(df_valid), 'PCT总量': len(df_pct), '失效总量': len(df_loss)}
    
    group_auth = df_auth.groupby('系统划分区县').size()
    group_valid = df_valid.groupby('系统划分区县').size()
    group_pct = df_pct.groupby('系统划分区县').size()
    group_loss = df_loss.groupby('系统划分区县').size()
    
    df_summary = pd.DataFrame({
        '本月授权(件)': group_auth, 
        '有效存量(件)': group_valid, 
        'PCT申请(件)': group_pct, 
        '本月失效(件)': group_loss
    }).fillna(0).astype(int)
    
    df_summary.loc['吴江区总计'] = [totals['授权总量'], totals['有效总量'], totals['PCT总量'], totals['失效总量']]
    df_summary = df_summary.drop(labels=['吴江区', '苏州市吴江区'], errors='ignore')
    
    towns_auth = df_summary['本月授权(件)'].drop(labels=['吴江区总计'], errors='ignore')
    top3_districts = towns_auth.sort_values(ascending=False).head(3).index.tolist()
    
    top10_applicants = df_auth['专利权人名称'].value_counts().head(10)
    applicant_types = df_auth['专利权人类型'].value_counts()
    
    df_auth['技术大类'] = df_auth['主分类号'].astype(str).str[:3]
    top5_ipc = df_auth['技术大类'].value_counts().head(5)
    loss_reasons = df_loss['失效原因'].value_counts()
    
    return totals, df_summary, top3_districts, top10_applicants, applicant_types, top5_ipc, loss_reasons

# ==========================================
# 文档生成：生成 Word
# ==========================================
def generate_official_word(ai_text, df_summary, top10_applicants, loss_reasons):
    doc = Document()
    doc.styles['Normal'].font.name = u'仿宋'
    doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), u'仿宋')
    
    title = doc.add_heading('', level=0)
    run = title.add_run('吴江区知识产权统计深度简报')
    run.font.name = u'方正小标宋简体'
    run.font.color.rgb = RGBColor(255, 0, 0)
    run.font.size = Pt(22)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph('=' * 45).alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_heading('一、本月知识产权宏观产出与特征分析', level=1)
    doc.add_paragraph(ai_text)
    
    doc.add_heading('二、附表：各区镇专利产出明细表', level=2)
    t1 = doc.add_table(rows=1, cols=5)
    t1.style = 'Table Grid'
    t1.rows[0].cells[0].text, t1.rows[0].cells[1].text, t1.rows[0].cells[2].text, t1.rows[0].cells[3].text, t1.rows[0].cells[4].text = \
        '区镇/板块', '本月授权(件)', '有效存量(件)', 'PCT申请(件)', '本月失效(件)'
    for district, row in df_summary.iterrows():
        row_cells = t1.add_row().cells
        row_cells[0].text, row_cells[1].text, row_cells[2].text, row_cells[3].text, row_cells[4].text = \
            str(district), str(row['本月授权(件)']), str(row['有效存量(件)']), str(row['PCT申请(件)']), str(row['本月失效(件)'])

    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

# ==========================================
# 文档生成：生成 PDF 可视化大屏
# ==========================================
def generate_visual_pdf(df_summary, top10_applicants, applicant_types, loss_reasons):
    fig, axes = plt.subplots(2, 2, figsize=(16, 12))
    fig.suptitle('吴江区知识产权深度分析大屏', fontsize=22, fontweight='bold', y=0.98)
    
    # 图 1：各区镇授权量
    ax1 = axes[0, 0]
    df_plot = df_summary.drop(labels=['吴江区总计'], errors='ignore')
    ax1.bar(df_plot.index, df_plot['本月授权(件)'], color='#4C72B0')
    ax1.set_title('吴江各区镇发明专利授权量分布')
    ax1.tick_params(axis='x', rotation=30)
    
    # 图 2：TOP 10 主体
    ax2 = axes[0, 1]
    ax2.barh(top10_applicants.index[::-1], top10_applicants.values[::-1], color='#55A868')
    ax2.set_title('本期授权量 TOP 10 领军主体')
    
    # 图 3：主体类型占比
    ax3 = axes[1, 0]
    ax3.pie(applicant_types.values, labels=applicant_types.index, autopct='%1.1f%%', startangle=140)
    ax3.set_title('获权主体类型结构')
    
    # 图 4：失效原因分析
    ax4 = axes[1, 1]
    if not loss_reasons.empty:
        ax4.pie(loss_reasons.values, labels=loss_reasons.index, autopct='%1.1f%%')
        ax4.set_title('本月专利失效原因分布')
    else:
        ax4.text(0.5, 0.5, '本期无失效数据', ha='center', va='center', fontsize=14)
        ax4.axis('off')

    plt.tight_layout(rect=[0, 0.03, 1, 0.95])
    pdf_io = io.BytesIO()
    plt.savefig(pdf_io, format='pdf')
    plt.close(fig)
    pdf_io.seek(0)
    return pdf_io

# ==========================================
# 文档生成：生成多表单 Excel
# ==========================================
def generate_visual_excel(df_summary, top10_applicants, loss_reasons):
    excel_io = io.BytesIO()
    with pd.ExcelWriter(excel_io, engine='openpyxl') as writer:
        df_summary.to_excel(writer, sheet_name='区镇板块汇总')
        top10_applicants.to_excel(writer, sheet_name='TOP10创新主体')
        loss_reasons.to_excel(writer, sheet_name='失效预警分析')
    excel_io.seek(0)
    return excel_io

# ==========================================
# 布局与交互控制台
# ==========================================
col_left, col_right = st.columns([1, 2])

with col_left:
    st.markdown('<div class="flat-title">上传业务清单</div>', unsafe_allow_html=True)
    
    # 引擎选择
    selected_model_input = st.selectbox(
        "选择大模型驱动引擎", 
        ("豆包大模型 (Volcengine)", "DeepSeek"), 
        label_visibility="collapsed"
    )
    st.markdown("<br>", unsafe_allow_html=True)
    
    # 文件上传
    f_valid = st.file_uploader("《有效专利》", type=["xlsx", "xls"])
    f_pct = st.file_uploader("《PCT申请》", type=["xlsx", "xls"])
    f_auth = st.file_uploader("《发明授权》", type=["xlsx", "xls"])
    f_loss = st.file_uploader("《失效专利》", type=["xlsx", "xls"])
    
    st.markdown("<br>", unsafe_allow_html=True) 
    start_btn = st.button("🚀 开始自动化全区研判", use_container_width=True, type="primary")

# 按钮触发逻辑
if start_btn:
    if not all([f_valid, f_pct, f_auth, f_loss]):
        st.warning("⚠️ 提示：请传齐全部 4 份业务清单后，再启动智能体！")
    else:
        st.session_state.selected_model = selected_model_input
        st.session_state.started = True
        st.session_state.processed = False
        st.rerun() 

with col_right:
    st.markdown('<div class="flat-title">智能体中央控制台</div>', unsafe_allow_html=True)
    
    if not st.session_state.started:
        st.info("👈 请在左侧上传业务清单，配置分析引擎，并点击“开始”启动自动化分析流程。")
        
    if st.session_state.started:
        if not st.session_state.processed:
            with st.status("智能体全链路运作中...", expanded=True) as status:
                # 步骤一：数据透视
                st.write("📥 [1/4] 正在解析表格并执行多维度分析...")
                try:
                    totals, df_summary, top3_dist, top10_app, app_types, top5_ipc, loss_rea = process_excel_data(f_valid, f_pct, f_auth, f_loss)
                except Exception as e:
                    status.update(label="数据解析异常中断！", state="error", expanded=True)
                    st.error(f"🚨 【格式不规范提醒】请检查上传的文件列名是否符合要求。错误详情: {str(e)}")
                    st.stop()
                
                # 步骤二：调用大模型撰写公文
                st.write(f"✍️ [2/4] 正在唤醒 {st.session_state.selected_model} 撰写简报公文...")
                
                # 动态分配 API 密钥和 URL
                if st.session_state.selected_model == "豆包大模型 (Volcengine)":
                    c_key = VOLC_API_KEY
                    c_url = VOLC_BASE_URL
                    c_name = VOLC_MODEL_NAME
                else:
                    c_key = DEEPSEEK_API_KEY
                    c_url = DEEPSEEK_BASE_URL
                    c_name = DEEPSEEK_MODEL_NAME
                
                try:
                    client = OpenAI(api_key=c_key, base_url=c_url)
                    prompt = f"""
                    你是一位吴江区市场监管局的高级知识产权分析师。请根据以下我提供的真实底层核算数据，撰写一段高质量的《知识产权月度分析简报》正文。
                    
                    【核心数据摘要】：
                    - 本月新增授权量为 {totals['授权总量']} 件。
                    - 全区有效发明专利总盘达到 {totals['有效总量']} 件。
                    - 各区镇中，【{', '.join(top3_dist)}】包揽了最多的本月授权量，展现出强劲的区域创新活力。
                    - 创新主体方面，本月火力最猛的企业是“{top10_app.index[0]}”（贡献了 {top10_app.values[0]} 件）。
                    - 核心技术突破主要集中在 IPC 分类号为【{', '.join(top5_ipc.index)}】的技术领域。
                    
                    【排版与语气要求】：
                    1. 请使用“政府官方智库”的口吻，语言精炼、数据准确、避免过度感叹。
                    2. 提炼出核心亮点，自然分成两个段落。
                    3. 总字数控制在 300 到 400 字之间。
                    """
                    
                    response = client.chat.completions.create(
                        model=c_name, 
                        messages=[{"role": "user", "content": prompt}], 
                        temperature=0.3
                    )
                    final_text = response.choices[0].message.content
                
                except Exception as e:
                    error_msg = str(e)
                    status.update(label="API 调用异常", state="error", expanded=True)
                    if "401" in error_msg or "402" in error_msg or "Insufficient Balance" in error_msg:
                        st.error(f"🚨 **【账户异常或余额不足】**：您的 {st.session_state.selected_model} 账户鉴权失败，可能是余额用尽或 Key 有误，请检查。")
                    else:
                        st.error(f"🚨 **【调用失败】**：连接大模型服务出错。详情：{error_msg}")
                    st.stop()
                
                # 步骤三：打包生成各种文件
                st.write("📑 [3/4] 正在将分析数据渲染为可视化大屏，并挂载至 Word 模板...")
                st.session_state.word_file = generate_official_word(final_text, df_summary, top10_app, loss_rea)
                st.session_state.excel_file = generate_visual_excel(df_summary, top10_app, loss_rea)
                st.session_state.pdf_file = generate_visual_pdf(df_summary, top10_app, app_types, loss_rea)
                st.session_state.final_text = final_text
                
                # 步骤四：完成
                st.session_state.processed = True
                status.update(label="处理完成！所有数据已就绪。", state="complete", expanded=False)
                st.rerun() 
                
        if st.session_state.processed:
            st.success("🎉 数据深加工完成！请下载您的多形态简报包。")
            
            c1, c2, c3 = st.columns(3)
            with c1: 
                st.download_button("📝 下载 Word 简报正文", data=st.session_state.word_file, file_name="吴江区知识产权_本期研判.docx", use_container_width=True)
            with c2: 
                st.download_button("📊 下载 Excel 汇总底稿", data=st.session_state.excel_file, file_name="吴江区_多板块汇总底稿.xlsx", use_container_width=True)
            with c3: 
                st.download_button("📈 下载可视化大屏 (PDF)", data=st.session_state.pdf_file, file_name="吴江区_专利结构可视化大屏.pdf", use_container_width=True)
            
            st.markdown('<div class="flat-title" style="margin-top:

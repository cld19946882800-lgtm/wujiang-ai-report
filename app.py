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
# 🚀 核心大模型双引擎配置 (已全部硬编码)
# ==========================================
# --- 1. 豆包大模型配置 ---
VOLC_API_KEY = "ark-2358e229-f997-477c-a223-f71a3c1d67f4-13835"
VOLC_BASE_URL = "https://ark.cn-beijing.volces.com/api/v3"
VOLC_MODEL_NAME = "ep-20260503224430-dsq28"

# --- 2. DeepSeek 大模型配置 ---
DEEPSEEK_API_KEY = "sk-0d889790af4d4eddbb912030535b49a2"
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
# 💎 极简稳定版 CSS (保留科技网格与按钮美化)
# ==========================================
st.markdown("""
    <style>
    /* 1. 科技网格绝对底层背景 */
    [data-testid="stAppViewContainer"] { background-color: transparent !important; }
    .main { background: transparent !important; }
    body, html { background-color: #f8fafc !important; }

    .stApp::before {
        content: ""; position: fixed; top: 0; right: 0; bottom: 0; left: 0; z-index: -1; 
        background-image: linear-gradient(to right, #e2e8f0 1px, transparent 1px), linear-gradient(to bottom, #e2e8f0 1px, transparent 1px);
        background-size: 20px 30px;
        -webkit-mask-image: radial-gradient(ellipse 70% 60% at 50% 0%, #000 60%, transparent 100%);
        mask-image: radial-gradient(ellipse 70% 60% at 50% 0%, #000 60%, transparent 100%);
        pointer-events: none; 
    }

    /* 2. 平面主标题 */
    .flat-title {
        border: 2px solid #e2e8f0; background-color: #ffffff; color: #1e293b;
        font-weight: 900; font-size: 20px; text-align: center;
        padding: 15px; border-radius: 8px; margin: 0 0 20px 0;
        width: 100%; box-sizing: border-box; letter-spacing: 2px;
    }

    /* 3. 按钮 3D 立体按压回弹效果 */
    .stButton > button {
        box-shadow: 0px 5px 0px 0px #94a3b8, 0px 10px 15px rgba(0,0,0,0.1) !important;
        transform: translateY(0px) !important; transition: all 0.1s !important;
        border-radius: 10px !important; font-weight: bold !important;
    }
    .stButton > button:active {
        box-shadow: 0px 0px 0px 0px #94a3b8 !important; transform: translateY(5px) !important;
    }
    .stButton > button[kind="primary"] {
        box-shadow: 0px 5px 0px 0px #b91c1c, 0px 10px 15px rgba(0,0,0,0.1) !important;
        background-color: #ef4444 !important; color: white !important;
    }
    .stButton > button[kind="primary"]:active {
        box-shadow: 0px 0px 0px 0px #b91c1c !important; transform: translateY(5px) !important;
    }

    /* 4. 规范上传组件的外壳，防止拥挤 */
    [data-testid="stFileUploader"] {
        background-color: white; border: 1px solid #e2e8f0; border-radius: 8px; padding: 10px; margin-bottom: 15px;
    }
    </style>
""", unsafe_allow_html=True)

# --- 页面主标题区 ---
st.title("📈 吴江区知识产权简报 AI 生成平台")
st.markdown("上传吴江区各板块真实业务清单，Agent 将全自动完成数据透视、公文撰写与多图表大屏排版。")
st.markdown("---")


# ==========================================
# ⬇️ 模块一：数据处理引擎 ⬇️
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

    df_summary = pd.DataFrame({'本月授权(件)': group_auth, '有效存量(件)': group_valid, 'PCT申请(件)': group_pct,
                               '本月失效(件)': group_loss}).fillna(0).astype(int)

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
# ⬇️ 模块二：排版生成引擎 ⬇️
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
    t1.rows[0].cells[0].text, t1.rows[0].cells[1].text, t1.rows[0].cells[2].text, t1.rows[0].cells[3].text, \
    t1.rows[0].cells[4].text = \
        '区镇/板块', '本月授权(件)', '有效存量(件)', 'PCT申请(件)', '本月失效(件)'
    for district, row in df_summary.iterrows():
        row_cells = t1.add_row().cells
        row_cells[0].text, row_cells[1].text, row_cells[2].text, row_cells[3].text, row_cells[4].text = \
            str(district), str(row['本月授权(件)']), str(row['有效存量(件)']), str(row['PCT申请(件)']), str(
                row['本月失效(件)'])

    doc.add_heading('三、附表：本月授权量 TOP 10 创新主体榜单', level=2)
    t2 = doc.add_table(rows=1, cols=3)
    t2.style = 'Table Grid'
    t2.rows[0].cells[0].text, t2.rows[0].cells[1].text, t2.rows[0].cells[2].text = '排名', '专利权人名称', '授权数量(件)'
    for idx, (name, count) in enumerate(top10_applicants.items(), 1):
        row_cells = t2.add_row().cells
        row_cells[0].text, row_cells[1].text, row_cells[2].text = str(idx), str(name), str(count)

    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io


def generate_visual_pdf(df_summary, top10_applicants, applicant_types, loss_reasons):
    fig, axes = plt.subplots(2, 2, figsize=(16, 12))
    fig.suptitle('吴江区知识产权深度分析大屏', fontsize=22, fontweight='bold', y=0.98)

    ax1 = axes[0, 0]
    df_plot = df_summary.drop(labels=['吴江区总计'], errors='ignore')
    bars = ax1.bar(df_plot.index, df_plot['本月授权(件)'], color='#4C72B0')
    ax1.set_title('吴江各区镇发明专利授权量分布', fontsize=14)
    ax1.tick_params(axis='x', rotation=30)
    for bar in bars:
        ax1.text(bar.get_x() + bar.get_width() / 2., bar.get_height(), int(bar.get_height()), ha='center', va='bottom')

    ax2 = axes[0, 1]
    ax2.barh(top10_applicants.index[::-1], top10_applicants.values[::-1], color='#55A868')
    ax2.set_title('本期授权量 TOP 10 领军主体', fontsize=14)
    for i, v in enumerate(top10_applicants.values[::-1]):
        ax2.text(v + 0.1, i, str(v), va='center')

    ax3 = axes[1, 0]
    ax3.pie(applicant_types.values, labels=applicant_types.index, autopct='%1.1f%%', startangle=140,
            colors=plt.cm.Pastel1.colors)
    ax3.set_title('获权主体类型结构', fontsize=14)

    ax4 = axes[1, 1]
    if not loss_reasons.empty:
        ax4.pie(loss_reasons.values, labels=loss_reasons.index, autopct='%1.1f%%', colors=plt.cm.Set3.colors)
        ax4.set_title('本月专利失效原因分布诊断', fontsize=14)
    else:
        ax4.text(0.5, 0.5, '本月无失效数据', ha='center', va='center')

    plt.tight_layout(rect=[0, 0.03, 1, 0.95])
    pdf_io = io.BytesIO()
    plt.savefig(pdf_io, format='pdf')
    plt.close(fig)
    pdf_io.seek(0)
    return pdf_io


def generate_visual_excel(df_summary, top10_applicants, loss_reasons):
    excel_io = io.BytesIO()
    with pd.ExcelWriter(excel_io, engine='openpyxl') as writer:
        df_summary.to_excel(writer, sheet_name='区镇板块汇总')
        top10_applicants.to_excel(writer, sheet_name='TOP10创新主体')
        loss_reasons.to_excel(writer, sheet_name='失效预警分析')
    excel_io.seek(0)
    return excel_io


# ==========================================
# ⬇️ 模块三：UI 前端绝对稳定原生布局 ⬇️
# ==========================================
# 采用最稳定的 1:2 原生两栏排版，绝不重叠错位
col_left, col_right = st.columns([1, 2])

with col_left:
    st.markdown('<div class="flat-title">上传业务清单</div>', unsafe_allow_html=True)

    st.markdown("##### ⚙️ 核心引擎配置")
    selected_model_input = st.selectbox(
        "选择大模型",
        ("豆包大模型 (Volcengine)", "DeepSeek"),
        label_visibility="collapsed"
    )
    st.markdown("<br>", unsafe_allow_html=True)

    # 原生上传框，稳定第一
    f_valid = st.file_uploader("《有效专利》", type=["xlsx", "xls"])
    f_pct = st.file_uploader("《PCT申请》", type=["xlsx", "xls"])
    f_auth = st.file_uploader("《发明授权》", type=["xlsx", "xls"])
    f_loss = st.file_uploader("《失效专利》", type=["xlsx", "xls"])

    st.markdown("<br>", unsafe_allow_html=True)
    start_btn = st.button("🚀 开始自动化全区研判", use_container_width=True, type="primary")

if start_btn:
    if not all([f_valid, f_pct, f_auth, f_loss]):
        st.warning("⚠️ 提示：请传齐全部 4 份业务清单数据源后再点击开始！")
    else:
        st.session_state.selected_model = selected_model_input
        st.session_state.started = True
        st.session_state.processed = False
        st.rerun()

with col_right:
    st.markdown('<div class="flat-title">智能体中央控制台</div>', unsafe_allow_html=True)

    if not st.session_state.started:
        st.info("👈 请在左侧上传业务清单，选择分析引擎，并点击“开始”按钮进行智能分析。")

    if st.session_state.started:
        if not st.session_state.processed:
            with st.status("智能体全链路运作中...", expanded=True) as status:
                st.write("📥 [1/4] 正在解析表格并执行多维度 GroupBy 聚类计算...")

                # 【防错拦截】拦截错误文件，防止系统崩溃
                try:
                    totals, df_summary, top3_dist, top10_app, app_types, top5_ipc, loss_rea = process_excel_data(
                        f_valid, f_pct, f_auth, f_loss)
                except Exception as e:
                    status.update(label="处理异常中断！", state="error", expanded=True)
                    st.error("🚨 【格式不规范提醒】您上传的文件内容不符合系统要求！请核对是否上传了错误的表格。")
                    st.stop()

                st.write(f"✍️ [2/4] 正在唤醒 {st.session_state.selected_model}，融合高维数据撰写官方研判公文...")

                # 🚀 双引擎调度逻辑 (已填入你的真实 Key)
                if st.session_state.selected_model == "豆包大模型 (Volcengine)":
                    c_key = VOLC_API_KEY
                    c_url = VOLC_BASE_URL
                    c_name = VOLC_MODEL_NAME
                else:
                    c_key = DEEPSEEK_API_KEY
                    c_url = DEEPSEEK_BASE_URL
                    c_name = DEEPSEEK_MODEL_NAME

                # 捕获网络及API调用报错
                try:
                    client = OpenAI(api_key=c_key, base_url=c_url)
                    prompt = f"""
                    你是一位吴江区市场监管局（知识产权局）的高级分析师。请根据以下我算出的吴江区真实底层数据，写一段高质量的《知识产权月度产出研判简报》核心正文。
                    吴江全区本期新增授权 {totals['授权总量']} 件，有效总盘达 {totals['有效总量']} 件。
                    板块创新高地效应显著，【{', '.join(top3_dist)}】包揽了区内最多的授权量。
                    全区本期研发火力最猛的头号领军企业是“{top10_app.index[0]}”（贡献了{top10_app.values[0]}件）。
                    核心技术突破集中在 IPC 分类号为【{', '.join(top5_ipc.index)}】的重点领域。
                    要求：用“政府智库”的口吻，分两个自然段，总字数300-400字左右。
                    """
                    response = client.chat.completions.create(
                        model=c_name,
                        messages=[{"role": "user", "content": prompt}],
                        temperature=0.3
                    )
                    final_text = response.choices[0].message.content
                except openai.APIConnectionError:
                    status.update(label="网络连接失败", state="error", expanded=True)
                    st.error(
                        f"🚨 **【网络连接失败】无法连接到 {st.session_state.selected_model} 的服务器！**\n\n这通常是因为你的电脑开启了 **VPN、加速器或系统代理**，拦截了 Python 的请求。请暂时关闭代理软件，或检查网络连接后再试。")
                    st.stop()
                except openai.AuthenticationError:
                    status.update(label="API 鉴权失败", state="error", expanded=True)
                    st.error(f"🚨 **【API Key 无效或过期】**：您提供的 {st.session_state.selected_model} 密钥无法通过验证。")
                    st.stop()
                except Exception as e:
                    status.update(label="API 调用异常", state="error", expanded=True)
                    st.error(f"🚨 **【调用失败】**：{str(e)}")
                    st.stop()

                st.write("📑 [3/4] 正在将分析结果挂载到 Word 模板与 PDF 可视化渲染引擎...")
                st.session_state.word_file = generate_official_word(final_text, df_summary, top10_app, loss_rea)
                st.session_state.excel_file = generate_visual_excel(df_summary, top10_app, loss_rea)
                st.session_state.pdf_file = generate_visual_pdf(df_summary, top10_app, app_types, loss_rea)
                st.session_state.final_text = final_text

                st.session_state.processed = True
                st.write("✅ [4/4] 全部打包完毕！")
                status.update(label="处理完成，等待用户下载", state="complete", expanded=False)
                st.rerun()

        if st.session_state.processed:
            st.success("🎉 数据深加工完成！")

            btn_col1, btn_col2, btn_col3 = st.columns(3)
            with btn_col1:
                st.download_button("📝 1. 下载吴江区研判简报", data=st.session_state.word_file,
                                   file_name="吴江区知识产权_本期研判.docx", use_container_width=True)
            with btn_col2:
                st.download_button("📊 2. 下载多板块底稿", data=st.session_state.excel_file,
                                   file_name="吴江区数据分析包.xlsx", use_container_width=True)
            with btn_col3:
                st.download_button("📈 3. 下载大屏 (PDF)", data=st.session_state.pdf_file,
                                   file_name="吴江区专利大屏.pdf", use_container_width=True)

            st.markdown('<div class="flat-title" style="margin-top: 30px;">AI 官方智库解读正文</div>',
                        unsafe_allow_html=True)
            st.info(st.session_state.final_text)
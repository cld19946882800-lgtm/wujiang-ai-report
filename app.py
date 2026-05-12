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

plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'Arial Unicode MS']
plt.rcParams['axes.unicode_minus'] = False

VOLC_API_KEY = "ark-2358e229-f997-477c-a223-f71a3c1d67f4-13835"
VOLC_BASE_URL = "https://ark.cn-beijing.volces.com/api/v3"
VOLC_MODEL_NAME = "ep-20260503224430-dsq28"

DEEPSEEK_API_KEY = "sk-0d889790af4d4eddbb912030535b49a2"
DEEPSEEK_BASE_URL = "https://api.deepseek.com/v1"
DEEPSEEK_MODEL_NAME = "deepseek-chat"

DOMAINS = {
    "知识产权": {
        "name": "知识产权",
        "icon": "📈",
        "description": "专利申请、授权、失效、PCT等数据分析",
        "required_files": [
            {"key": "f_valid", "label": "《有效专利清单》", "desc": "包含专利号、专利权人、申请日等"},
            {"key": "f_pct", "label": "《PCT申请清单》", "desc": "PCT国际申请记录"},
            {"key": "f_auth", "label": "《发明授权清单》", "desc": "本月新增授权专利"},
            {"key": "f_loss", "label": "《失效专利清单》", "desc": "本月失效专利清单"}
        ],
        "group_by_field": "系统划分区县",
        "metrics": [
            {"key": "授权总量", "source": "f_auth", "label": "本月授权(件)"},
            {"key": "有效总量", "source": "f_valid", "label": "有效存量(件)"},
            {"key": "PCT总量", "source": "f_pct", "label": "PCT申请(件)"},
            {"key": "失效总量", "source": "f_loss", "label": "本月失效(件)"}
        ],
        "top_field": "专利权人名称",
        "type_field": "专利权人类型",
        "class_field": "主分类号",
        "reason_field": "失效原因",
        "title": "吴江区知识产权统计深度简报",
        "prompt_template": """你是一位吴江区市场监管局（知识产权局）的高级分析师。请根据以下真实底层数据，写一段高质量的《知识产权月度产出研判简报》核心正文。
吴江全区本期新增授权 {授权总量} 件，有效总盘达 {有效总量} 件。
板块创新高地效应显著，【{top3_districts}】包揽了区内最多的授权量。
全区本期研发火力最猛的头号领军企业是"{top1_applicant}"（贡献了{top1_count}件）。
核心技术突破集中在 IPC 分类号为【{top_ipc}】的重点领域。
要求：用"政府智库"的口吻，分两个自然段，总字数300-400字左右。"""
    },
    "投诉举报": {
        "name": "投诉举报",
        "icon": "📞",
        "description": "12345投诉举报数据分析",
        "required_files": [
            {"key": "f_complaint", "label": "《投诉清单》", "desc": "消费者投诉记录"},
            {"key": "f_report", "label": "《举报清单》", "desc": "违法举报记录"},
            {"key": "f_resolve", "label": "《办结清单》", "desc": "已办结案件清单"},
            {"key": "f_satisfaction", "label": "《满意度评价》", "desc": "群众满意度评价"}
        ],
        "group_by_field": "区域",
        "metrics": [
            {"key": "投诉量", "source": "f_complaint", "label": "投诉量(件)"},
            {"key": "举报量", "source": "f_report", "label": "举报量(件)"},
            {"key": "办结量", "source": "f_resolve", "label": "办结量(件)"},
            {"key": "满意度", "source": "f_satisfaction", "label": "满意度(%)"}
        ],
        "top_field": "投诉人",
        "type_field": "投诉类型",
        "class_field": "行业分类",
        "reason_field": "投诉原因",
        "title": "吴江区投诉举报情况分析简报",
        "prompt_template": """你是一位吴江区市场监管局的高级分析师。请根据以下数据，写一段高质量的《投诉举报月度分析简报》。
本月共受理投诉 {投诉量} 件，举报 {举报量} 件。
办结案件 {办结量} 件，办结率较高。
热点投诉领域主要集中在【{top_types}】。
区域分布上，【{top3_districts}】投诉举报量较多。
要求：用"政府智库"的口吻，分两个自然段，总字数300-400字左右。"""
    },
    "食品安全": {
        "name": "食品安全",
        "icon": "🍱",
        "description": "食品安全监管、抽检、处罚数据分析",
        "required_files": [
            {"key": "f_inspection", "label": "《监督检查清单》", "desc": "日常检查记录"},
            {"key": "f_sampling", "label": "《抽检结果清单》", "desc":  "食品安全抽检结果"},
            {"key": "f_penalty", "label": "《行政处罚清单》", "desc": "行政处罚记录"},
            {"key": "f_recall", "label": "《产品召回清单》", "desc": "问题产品召回记录"}
        ],
        "group_by_field": "区域",
        "metrics": [
            {"key": "检查次数", "source": "f_inspection", "label": "检查次数(次)"},
            {"key": "抽检批次", "source": "f_sampling", "label": "抽检批次(批)"},
            {"key": "处罚案件", "source": "f_penalty", "label": "处罚案件(件)"},
            {"key": "召回数量", "source": "f_recall", "label": "召回数量(个)"}
        ],
        "top_field": "经营主体",
        "type_field": "业态类型",
        "class_field": "食品类别",
        "reason_field": "问题类型",
        "title": "吴江区食品安全监管情况分析简报",
        "prompt_template": """你是一位吴江区市场监管局的高级食品安全分析师。请根据以下数据，写一段高质量的《食品安全月度监管简报》。
本月共开展监督检查 {检查次数} 次，完成抽检 {抽检批次} 批次。
查处行政处罚案件 {处罚案件} 件，问题产品召回 {召回数量} 个。
监管重点领域主要集中在【{top_types}】。
区域分布上，【{top3_districts}】监管任务较重。
要求：用"政府智库"的口吻，分两个自然段，总字数300-400字左右。"""
    },
    "医疗器械": {
        "name": "医疗器械",
        "icon": "🏥",
        "description": "医疗器械生产、经营、监管数据分析",
        "required_files": [
            {"key": "f_manufacture", "label": "《生产企业清单》", "desc": "生产企业台账"},
            {"key": "f_operation", "label": "《经营企业清单》", "desc": "经营企业台账"},
            {"key": "f_inspection", "label": "《监督检查记录》", "desc": "日常检查记录"},
            {"key": "f_adverse", "label": "《不良事件报告》", "desc": "医疗器械不良事件"}
        ],
        "group_by_field": "区域",
        "metrics": [
            {"key": "生产企业", "source": "f_manufacture", "label": "生产企业(家)"},
            {"key": "经营企业", "source": "f_operation", "label": "经营企业(家)"},
            {"key": "检查次数", "source": "f_inspection", "label": "检查次数(次)"},
            {"key": "不良事件", "source": "f_adverse", "label": "不良事件(例)"}
        ],
        "top_field": "企业名称",
        "type_field": "企业类型",
        "class_field": "产品分类",
        "reason_field": "问题类别",
        "title": "吴江区医疗器械监管情况分析简报",
        "prompt_template": """你是一位吴江区市场监管局的高级医疗器械监管分析师。请根据以下数据，写一段高质量的《医疗器械月度监管简报》。
全区现有医疗器械生产企业 {生产} 家，经营企业 {经营_value} 家。
本月开展监督检查 {检查次数} 次，收到不良事件报告 {不良事件} 例。
监管重点企业主要集中在【{top_types}】。
区域分布上，【{top3_districts}】企业数量较多。
要求：用"政府智库"的口吻，分两个自然段，总字数300-400字左右。"""
    }
}

if 'started' not in st.session_state:
    st.session_state.started = False
if 'processed' not in st.session_state:
    st.session_state.processed = False
if 'selected_model' not in st.session_state:
    st.session_state.selected_model = "豆包大模型 (Volcengine)"
if 'selected_domain' not in st.session_state:
    st.session_state.selected_domain = "知识产权"

st.set_page_config(page_title="吴江区市场监管数据分析平台", layout="wide", page_icon="📊")

st.markdown("""
    <style>
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

    .flat-title {
        border: 2px solid #e2e8f0; background-color: #ffffff; color: #1e293b;
        font-weight: 900; font-size: 20px; text-align: center;
        padding: 15px; border-radius: 8px; margin: 0 0 20px 0;
        width: 100%; box-sizing: border-box; letter-spacing: 2px;
    }

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

    [data-testid="stFileUploader"] {
        background-color: white; border: 1px solid #e2e8f0; border-radius: 8px; padding: 10px; margin-bottom: 15px;
    }
    </style>
""", unsafe_allow_html=True)

st.title("📊 吴江区市场监管数据分析平台")
st.markdown("选择分析领域，上传业务数据，AI 自动生成多形态分析简报。")
st.markdown("---")

domain_list = [f"{DOMAINS[k]['icon']} {DOMAINS[k]['name']}" for k in DOMAINS.keys()]
domain_keys = list(DOMAINS.keys())

def get_uploaded_files(domain_key):
    files = {}
    domain_config = DOMAINS[domain_key]
    for file_config in domain_config["required_files"]:
        files[file_config["key"]] = None
    return files

def process_domain_data(domain_key, uploaded_files):
    domain = DOMAINS[domain_key]
    group_field = domain["group_by_field"]
    
    dfs = {}
    for file_config in domain["required_files"]:
        key = file_config["key"]
        if uploaded_files.get(key):
            dfs[key] = pd.read_excel(uploaded_files[key])
    
    totals = {}
    for metric in domain["metrics"]:
        source = metric["source"]
        if source in dfs:
            if "率" in metric["label"] or "满意度" in metric["label"]:
                totals[metric["key"]] = round(dfs[source][domain["class_field"]].mean(), 2) if len(dfs[source]) > 0 else 0
            else:
                totals[metric["key"]] = len(dfs[source])
    
    summary_cols = {}
    for metric in domain["metrics"]:
        source = metric["source"]
        if source in dfs and group_field in dfs[source].columns:
            grouped = dfs[source].groupby(group_field).size()
            summary_cols[metric["label"]] = grouped
    
    if summary_cols:
        df_summary = pd.DataFrame(summary_cols).fillna(0).astype(int)
    else:
        df_summary = pd.DataFrame()
    
    top_field = domain.get("top_field")
    type_field = domain.get("type_field")
    class_field = domain.get("class_field")
    reason_field = domain.get("reason_field")
    
    top_data = {}
    type_data = {}
    class_data = {}
    reason_data = {}
    
    first_df = list(dfs.values())[0] if dfs else None
    
    if first_df is not None:
        if top_field and top_field in first_df.columns:
            top_data = first_df[top_field].value_counts().head(10)
        if type_field and type_field in first_df.columns:
            type_data = first_df[type_field].value_counts()
        if class_field and class_field in first_df.columns:
            first_df['大类'] = first_df[class_field].astype(str).str[:3]
            class_data = first_df['大类'].value_counts().head(5)
        if reason_field and reason_field in first_df.columns:
            reason_data = first_df[reason_field].value_counts()
    
    return totals, df_summary, top_data, type_data, class_data, reason_data


def generate_domain_word(domain_key, ai_text, df_summary, top_data, type_data):
    domain = DOMAINS[domain_key]
    doc = Document()
    doc.styles['Normal'].font.name = u'仿宋'
    doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), u'仿宋')

    title = doc.add_heading('', level=0)
    run = title.add_run(domain["title"])
    run.font.name = u'方正小标宋简体'
    run.font.color.rgb = RGBColor(255, 0, 0)
    run.font.size = Pt(22)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph('=' * 45).alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_heading('一、宏观数据分析', level=1)
    doc.add_paragraph(ai_text)

    doc.add_heading('二、区域分布明细表', level=2)
    if not df_summary.empty:
        cols = len(df_summary.columns)
        t1 = doc.add_table(rows=1, cols=cols + 1)
        t1.style = 'Table Grid'
        headers = ['区域'] + list(df_summary.columns)
        for i, h in enumerate(headers):
            t1.rows[0].cells[i].text = h
        for idx, row in df_summary.iterrows():
            row_cells = t1.add_row().cells
            row_cells[0].text = str(idx)
            for i, col in enumerate(df_summary.columns):
                row_cells[i + 1].text = str(row[col])

    if top_data is not None and len(top_data) > 0:
        doc.add_heading(f'三、TOP 10 {domain.get("top_field", "主体")}榜单', level=2)
        t2 = doc.add_table(rows=1, cols=3)
        t2.style = 'Table Grid'
        t2.rows[0].cells[0].text = '排名'
        t2.rows[0].cells[1].text = domain.get("top_field", "名称")
        t2.rows[0].cells[2].text = '数量'
        for idx, (name, count) in enumerate(top_data.items(), 1):
            row_cells = t2.add_row().cells
            row_cells[0].text, row_cells[1].text, row_cells[2].text = str(idx), str(name), str(count)

    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io


def generate_domain_pdf(domain_key, df_summary, top_data, type_data, reason_data):
    domain = DOMAINS[domain_key]
    fig, axes = plt.subplots(2, 2, figsize=(16, 12))
    fig.suptitle(f'{domain["name"]}分析大屏', fontsize=22, fontweight='bold', y=0.98)

    ax1 = axes[0, 0]
    if not df_summary.empty:
        first_col = df_summary.columns[0]
        ax1.bar(df_summary.index, df_summary[first_col], color='#4C72B0')
        ax1.set_title(f'{domain["name"]}各区域分布', fontsize=14)
        ax1.tick_params(axis='x', rotation=30)
    else:
        ax1.text(0.5, 0.5, '暂无数据', ha='center', va='center')

    ax2 = axes[0, 1]
    if top_data is not None and len(top_data) > 0:
        ax2.barh(list(top_data.index)[::-1], list(top_data.values)[::-1], color='#55A868')
        ax2.set_title(f'TOP 10 {domain.get("top_field", "主体")}', fontsize=14)
    else:
        ax2.text(0.5, 0.5, '暂无数据', ha='center', va='center')
        ax2.axis('off')

    ax3 = axes[1, 0]
    if type_data is not None and len(type_data) > 0:
        ax3.pie(type_data.values, labels=type_data.index, autopct='%1.1f%%', startangle=140)
        ax3.set_title(f'{domain.get("type_field", "类型")}分布', fontsize=14)
    else:
        ax3.text(0.5, 0.5, '暂无数据', ha='center', va='center')
        ax3.axis('off')

    ax4 = axes[1, 1]
    if reason_data is not None and len(reason_data) > 0:
        ax4.pie(reason_data.values, labels=reason_data.index, autopct='%1.1f%%')
        ax4.set_title(f'{domain.get("reason_field", "原因")}分析', fontsize=14)
    else:
        ax4.text(0.5, 0.5, '暂无原因数据', ha='center', va='center')
        ax4.axis('off')

    plt.tight_layout(rect=[0, 0.03, 1, 0.95])
    pdf_io = io.BytesIO()
    plt.savefig(pdf_io, format='pdf')
    plt.close(fig)
    pdf_io.seek(0)
    return pdf_io


def generate_domain_excel(domain_key, df_summary, top_data, type_data, reason_data):
    excel_io = io.BytesIO()
    with pd.ExcelWriter(excel_io, engine='openpyxl') as writer:
        if not df_summary.empty:
            df_summary.to_excel(writer, sheet_name='区域汇总')
        if top_data is not None and len(top_data) > 0:
            pd.DataFrame(top_data).to_excel(writer, sheet_name='TOP10')
        if type_data is not None and len(type_data) > 0:
            pd.DataFrame(type_data).to_excel(writer, sheet_name='类型分布')
    excel_io.seek(0)
    return excel_io


col_left, col_right = st.columns([1, 2])

with col_left:
    st.markdown('<div class="flat-title">📁 数据上传与领域选择</div>', unsafe_allow_html=True)

    st.markdown("##### 🏷️ 选择分析领域")
    domain_options = [f"{DOMAINS[k]['icon']} {DOMAINS[k]['name']}" for k in DOMAINS.keys()]
    selected_domain_display = st.selectbox(
        "选择市场监管分析领域",
        domain_options,
        index=domain_keys.index(st.session_state.selected_domain),
        label_visibility="collapsed"
    )
    st.session_state.selected_domain = domain_keys[domain_options.index(selected_domain_display)]

    domain = DOMAINS[st.session_state.selected_domain]
    st.markdown(f"*{domain['description']}*")

    if 'domain_files' not in st.session_state or st.session_state.get('last_domain') != st.session_state.selected_domain:
        st.session_state.domain_files = {}
        st.session_state.last_domain = st.session_state.selected_domain

    st.markdown("##### 📤 上传业务清单")
    uploaded_files = {}
    for file_config in domain["required_files"]:
        key = file_config["key"]
        uploaded_files[key] = st.file_uploader(
            file_config["label"],
            type=["xlsx", "xls"],
            key=f"file_{key}"
        )

    st.markdown("##### ⚙️ 大模型引擎")
    selected_model_input = st.selectbox(
        "选择大模型",
        ("豆包大模型 (Volcengine)", "DeepSeek"),
        label_visibility="collapsed"
    )

    st.markdown("<br>", unsafe_allow_html=True)
    start_btn = st.button("🚀 开始自动化分析", use_container_width=True, type="primary")

if start_btn:
    if not all(uploaded_files.values()):
        st.warning(f"⚠️ 提示：请上传全部 {len(domain['required_files'])} 份业务清单后再点击开始！")
    else:
        st.session_state.selected_model = selected_model_input
        st.session_state.started = True
        st.session_state.processed = False
        st.session_state.uploaded_files = uploaded_files
        st.rerun()

with col_right:
    st.markdown('<div class="flat-title">🎛️ 智能分析控制台</div>', unsafe_allow_html=True)

    if not st.session_state.started:
        st.info("👈 请在左侧选择领域、上传数据，选择分析引擎，点击开始进行智能分析。")

    if st.session_state.started:
        if not st.session_state.processed:
            domain = DOMAINS[st.session_state.selected_domain]
            with st.status("智能体全链路运作中...", expanded=True) as status:
                st.write("📥 [1/4] 正在解析表格并执行多维度分析...")

                try:
                    totals, df_summary, top_data, type_data, class_data, reason_data = process_domain_data(
                        st.session_state.selected_domain, st.session_state.uploaded_files)
                except Exception as e:
                    status.update(label="处理异常中断！", state="error", expanded=True)
                    st.error(f"🚨 【格式不规范提醒】您上传的文件内容不符合系统要求！错误详情: {str(e)}")
                    st.stop()

                st.write(f"✍️ [2/4] 正在唤醒 {st.session_state.selected_model} 撰写分析简报...")

                if st.session_state.selected_model == "豆包大模型 (Volcengine)":
                    c_key = VOLC_API_KEY
                    c_url = VOLC_BASE_URL
                    c_name = VOLC_MODEL_NAME
                else:
                    c_key = DEEPSEEK_API_KEY
                    c_url = DEEPSEEK_BASE_URL
                    c_name = DEEPSEEK_MODEL_NAME

                top3_districts = list(df_summary.index[:3]) if not df_summary.empty else []
                top1_applicant = list(top_data.index)[0] if top_data is not None and len(top_data) > 0 else "无"
                top1_count = list(top_data.values)[0] if top_data is not None and len(top_data) > 0 else 0
                top_types = list(type_data.index[:3]) if type_data is not None and len(type_data) > 0 else []
                top_ipc = list(class_data.index) if class_data is not None and len(class_data) > 0 else []

                prompt_values = {
                    "授权总量": totals.get("授权总量", 0),
                    "有效总量": totals.get("有效总量", 0),
                    "投诉量": totals.get("投诉量", 0),
                    "检查次数": totals.get("检查次数", 0),
                    "top3_districts": ', '.join(top3_districts) if top3_districts else "暂无",
                    "top1_applicant": top1_applicant,
                    "top1_count": top1_count,
                    "top_types": ', '.join(top_types) if top_types else "暂无",
                    "top_ipc": ', '.join(top_ipc) if top_ipc else "暂无",
                    "生产": totals.get("生产企业", 0),
                    "经营_value": totals.get("经营企业", 0),
                    "不良事件": totals.get("不良事件", 0),
                }

                prompt = domain["prompt_template"].format(**prompt_values)

                try:
                    client = OpenAI(api_key=c_key, base_url=c_url)
                    response = client.chat.completions.create(
                        model=c_name,
                        messages=[{"role": "user", "content": prompt}],
                        temperature=0.3
                    )
                    final_text = response.choices[0].message.content
                except openai.APIConnectionError:
                    status.update(label="网络连接失败", state="error", expanded=True)
                    st.error(f"🚨 **【网络连接失败】**无法连接到 {st.session_state.selected_model} 的服务器！请检查网络或代理设置。")
                    st.stop()
                except openai.AuthenticationError:
                    status.update(label="API 鉴权失败", state="error", expanded=True)
                    st.error(f"🚨 **【API Key 无效或过期】**：您提供的 {st.session_state.selected_model} 密钥无法通过验证。")
                    st.stop()
                except Exception as e:
                    status.update(label="API 调用异常", state="error", expanded=True)
                    st.error(f"🚨 **【调用失败】**：{str(e)}")
                    st.stop()

                st.write("📑 [3/4] 正在生成 Word、Excel、PDF 多形态报告...")
                st.session_state.word_file = generate_domain_word(
                    st.session_state.selected_domain, final_text, df_summary, top_data, type_data)
                st.session_state.excel_file = generate_domain_excel(
                    st.session_state.selected_domain, df_summary, top_data, type_data, reason_data)
                st.session_state.pdf_file = generate_domain_pdf(
                    st.session_state.selected_domain, df_summary, top_data, type_data, reason_data)
                st.session_state.final_text = final_text
                st.session_state.domain_title = domain["title"]

                st.session_state.processed = True
                st.write("✅ [4/4] 全部打包完毕！")
                status.update(label="处理完成", state="complete", expanded=False)
                st.rerun()

        if st.session_state.processed:
            st.success("🎉 数据深加工完成！")

            btn_col1, btn_col2, btn_col3 = st.columns(3)
            with btn_col1:
                st.download_button("📝 下载分析简报", data=st.session_state.word_file,
                                   file_name=f"{st.session_state.selected_domain}_分析简报.docx", use_container_width=True)
            with btn_col2:
                st.download_button("📊 下载数据底稿", data=st.session_state.excel_file,
                                   file_name=f"{st.session_state.selected_domain}_数据底稿.xlsx", use_container_width=True)
            with btn_col3:
                st.download_button("📈 下载可视化大屏", data=st.session_state.pdf_file,
                                   file_name=f"{st.session_state.selected_domain}_分析大屏.pdf", use_container_width=True)

            st.markdown('<div class="flat-title" style="margin-top: 30px;">📄 AI 分析报告正文</div>',
                        unsafe_allow_html=True)
            st.info(st.session_state.final_text)

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

DEFAULT_DOMAINS = {
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

def init_custom_domains():
    if 'custom_domains' not in st.session_state:
        st.session_state.custom_domains = {}

def save_custom_domain(name, files_info):
    st.session_state.custom_domains[name] = {
        "name": name,
        "icon": "📋",
        "description": f"自定义分析领域：{name}",
        "required_files": files_info,
        "title": f"{name}分析报告"
    }

init_custom_domains()
DEFAULT_DOMAINS["自定义"] = {
    "name": "自定义",
    "icon": "📋",
    "description": "自定义上传数据并分析",
    "required_files": [],
    "title": "自定义数据分析报告",
    "prompt_template": """你是一位数据分析专家。请根据用户上传的业务数据进行分析，并撰写一份专业的分析简报。

【数据概况】
数据文件数：{数据文件数}
总记录数：{总记录数}

请根据以上数据：
1. 识别数据中的关键指标和趋势
2. 分析主要特征和模式
3. 用"政府智库"的专业口吻撰写一份300-400字的月度分析简报
4. 简报应包含数据解读、问题发现、建议意见三个部分"""
}
DOMAINS = {**DEFAULT_DOMAINS, **st.session_state.custom_domains}

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
    /* 科技网格背景 - 使用 fixed 定位确保不被遮挡 */
    html, body, .stApp, [data-testid="stAppViewContainer"] {
        background-image: 
            linear-gradient(to right, #e2e8f0 1px, transparent 1px),
            linear-gradient(to bottom, #e2e8f0 1px, transparent 1px) !important;
        background-size: 20px 30px !important;
        background-attachment: fixed !important;
        background-color: #f8fafc !important;
    }
    
    /* 内容区域半透明 */
    .main .block-container {
        background: rgba(255,255,255,0.95);
        border-radius: 10px;
        padding: 20px;
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
    
    dfs = {}
    for key, file_obj in uploaded_files.items():
        if file_obj is not None:
            try:
                if file_obj.name.endswith('.csv'):
                    dfs[key] = pd.read_csv(file_obj)
                else:
                    dfs[key] = pd.read_excel(file_obj)
            except:
                pass
    
    all_dataframes = []
    for key, df in dfs.items():
        df['_source_file'] = key
        all_dataframes.append(df)
    
    combined_df = pd.concat(all_dataframes, ignore_index=True) if all_dataframes else pd.DataFrame()
    
    totals = {"数据文件数": len(dfs), "总记录数": len(combined_df)}
    
    numeric_cols = combined_df.select_dtypes(include=['number']).columns.tolist()
    text_cols = combined_df.select_dtypes(include=['object']).columns.tolist()
    
    if numeric_cols:
        for col in numeric_cols[:5]:
            if not combined_df[col].isna().all():
                totals[f"{col}总量"] = int(combined_df[col].sum())
    
    possible_group_fields = ['区域', '系统划分区县', '区县', '板块', '街道', '乡镇']
    group_field = None
    for pf in possible_group_fields:
        if pf in combined_df.columns:
            group_field = pf
            break
    
    df_summary = pd.DataFrame()
    if group_field and numeric_cols:
        summary_cols = {}
        for col in numeric_cols[:4]:
            summary_cols[col] = combined_df.groupby(group_field)[col].sum()
        if summary_cols:
            df_summary = pd.DataFrame(summary_cols).fillna(0).astype(int)
    
    top_data = {}
    type_data = {}
    reason_data = {}
    
    if text_cols:
        for col in text_cols[:3]:
            vc = combined_df[col].value_counts().head(10)
            if len(vc) > 0:
                if not top_data:
                    top_data = vc.to_dict()
                elif not type_data:
                    type_data = vc.to_dict()
                elif not reason_data:
                    reason_data = vc.to_dict()
    
    sample_columns = list(combined_df.columns[:10]) if len(combined_df.columns) > 0 else []
    
    return totals, df_summary, top_data, type_data, reason_data, sample_columns


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
        doc.add_heading('三、数据分布统计', level=2)
        t2 = doc.add_table(rows=1, cols=2)
        t2.style = 'Table Grid'
        t2.rows[0].cells[0].text = '类别'
        t2.rows[0].cells[1].text = '数量'
        for name, count in list(top_data.items())[:10]:
            row_cells = t2.add_row().cells
            row_cells[0].text, row_cells[1].text = str(name), str(count)

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
        items = list(top_data.items())[:10]
        ax2.barh([str(k) for k, v in items][::-1], [v for k, v in items][::-1], color='#55A868')
        ax2.set_title('TOP 10 数据分布', fontsize=14)
    else:
        ax2.text(0.5, 0.5, '暂无数据', ha='center', va='center')
        ax2.axis('off')

    ax3 = axes[1, 0]
    if type_data is not None and len(type_data) > 0:
        ax3.pie(list(type_data.values())[:6], labels=list(type_data.keys())[:6], autopct='%1.1f%%', startangle=140)
        ax3.set_title('类型分布', fontsize=14)
    else:
        ax3.text(0.5, 0.5, '暂无数据', ha='center', va='center')
        ax3.axis('off')

    ax4 = axes[1, 1]
    if reason_data is not None and len(reason_data) > 0:
        ax4.pie(list(reason_data.values())[:6], labels=list(reason_data.keys())[:6], autopct='%1.1f%%')
        ax4.set_title('原因分析', fontsize=14)
    else:
        ax4.text(0.5, 0.5, '暂无数据', ha='center', va='center')
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
            pd.DataFrame(top_data, index=[0]).T.to_excel(writer, sheet_name='数据分布')
        if type_data is not None and len(type_data) > 0:
            pd.DataFrame(type_data, index=[0]).T.to_excel(writer, sheet_name='类型分布')
    excel_io.seek(0)
    return excel_io


col_left, col_right = st.columns([1, 2])

with col_left:
    st.markdown('<div class="flat-title">📁 数据上传与领域选择</div>', unsafe_allow_html=True)

    st.markdown("##### 🏷️ 选择分析领域")
    preset_domains = [
        ("知识产权", "📈", "专利申请、授权、失效、PCT等数据分析"),
        ("投诉举报", "📞", "12345投诉举报数据分析"),
        ("食品安全", "🍱", "食品安全监管、抽检、处罚数据分析"),
        ("医疗器械", "🏥", "医疗器械生产、经营、监管数据分析"),
        ("自定义", "📋", "自定义上传数据并分析")
    ]

    domain_options = [f"{d[1]} {d[0]}" for d in preset_domains]
    domain_keys = [d[0] for d in preset_domains]

    selected_domain_display = st.selectbox(
        "选择市场监管分析领域",
        domain_options,
        index=domain_keys.index(st.session_state.selected_domain) if st.session_state.selected_domain in domain_keys else 0,
        label_visibility="collapsed"
    )

    selected_idx = domain_options.index(selected_domain_display)
    st.session_state.selected_domain = domain_keys[selected_idx]

    domain_desc = ""
    for d in preset_domains:
        if d[0] == st.session_state.selected_domain:
            domain_desc = d[2]
            break
    st.markdown(f"*{domain_desc}*")

    st.markdown("##### 📤 上传业务清单")
    uploaded_files = {}
    
    if st.session_state.selected_domain == "知识产权":
        ip_files = [
            {"key": "f_valid", "label": "《有效专利清单》"},
            {"key": "f_pct", "label": "《PCT申请清单》"},
            {"key": "f_auth", "label": "《发明授权清单》"},
            {"key": "f_loss", "label": "《失效专利清单》"}
        ]
        for f in ip_files:
            uploaded_files[f["key"]] = st.file_uploader(f["label"], type=["xlsx", "xls", "csv"], key=f"file_{f['key']}")
    elif st.session_state.selected_domain == "自定义":
        uploaded = st.file_uploader("上传数据文件（支持多文件）", type=["xlsx", "xls", "csv"], accept_multiple_files=True, key="custom_multi_files")
        if uploaded:
            for i, f in enumerate(uploaded):
                uploaded_files[f"custom_{i}"] = f
    else:
        if 'dynamic_files' not in st.session_state:
            st.session_state.dynamic_files = []
        
        for i, f_info in enumerate(st.session_state.dynamic_files):
            uploaded_files[f"dyn_{i}"] = st.file_uploader(f"《{f_info['label']}》", type=["xlsx", "xls", "csv"], key=f"dyn_file_{i}")
        
        col_f1, col_f2 = st.columns([3, 1])
        with col_f1:
            new_label = st.text_input("添加文件标签", placeholder="如：企业投诉数据", key="new_dynamic_file")
        with col_f2:
            st.markdown("<br>")
            if st.button("➕ 添加", key="add_dynamic"):
                if new_label:
                    st.session_state.dynamic_files.append({"label": new_label, "key": f"dyn_{len(st.session_state.dynamic_files)}"})
                    st.rerun()
        
        if st.session_state.dynamic_files:
            if st.button("清空文件列表", key="clear_dynamic"):
                st.session_state.dynamic_files = []
                st.rerun()

    st.markdown("##### ⚙️ 大模型引擎")
    selected_model_input = st.selectbox(
        "选择大模型",
        ("豆包大模型 (Volcengine)", "DeepSeek"),
        label_visibility="collapsed"
    )

    st.markdown("<br>", unsafe_allow_html=True)
    start_btn = st.button("🚀 开始自动化分析", use_container_width=True, type="primary")

if start_btn:
    uploaded_count = sum(1 for v in uploaded_files.values() if v is not None)
    if uploaded_count == 0:
        st.warning("⚠️ 提示：请至少上传1份业务数据后再点击开始！")
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
                    result = process_domain_data(
                        st.session_state.selected_domain, st.session_state.uploaded_files)
                    if len(result) == 6:
                        totals, df_summary, top_data, type_data, reason_data, sample_columns = result
                        class_data = {}
                    else:
                        totals, df_summary, top_data, type_data, class_data, reason_data = result
                        sample_columns = []
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

                is_custom_domain = st.session_state.selected_domain in st.session_state.custom_domains or st.session_state.selected_domain == "自定义"
                
                if is_custom_domain:
                    data_summary = "\n".join([f"- {k}: {v}" for k, v in totals.items()])
                    top_items = "\n".join([f"- {k}: {v}" for k, v in list(top_data.items())[:5]]) if top_data else "暂无"
                    
                    prompt = f"""你是一位数据分析专家。请根据用户上传的业务数据进行分析，并撰写一份专业的分析简报。

【数据概况】
{data_summary}

【关键数据分布】
{top_items}

请根据以上数据：
1. 识别数据中的关键指标和趋势
2. 分析主要特征和模式
3. 用"政府智库"的专业口吻撰写一份300-400字的月度分析简报
4. 简报应包含数据解读、问题发现、建议意见三个部分

注意：如果某些数据字段不适用于你的分析，可以忽略。"""
                else:
                    top3_districts = list(df_summary.index[:3]) if not df_summary.empty else []
                    top1_applicant = list(top_data.keys())[0] if top_data else "无"
                    top1_count = list(top_data.values())[0] if top_data else 0
                    top_types = list(type_data.keys())[:3] if type_data else []
                    top_ipc = list(class_data.keys()) if class_data else []

                    prompt_values = {
                        "数据文件数": totals.get("数据文件数", 0),
                        "总记录数": totals.get("总记录数", 0),
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

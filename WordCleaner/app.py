import streamlit as st
import re
from docx import Document
from docx.shared import Pt, Inches
from docx.oxml.ns import qn
from io import BytesIO

# 设置页面配置
st.set_page_config(
    page_title="Word文档格式化工具",
    page_icon="📝",
    layout="centered"  # 改为居中布局
)

# 默认配置
DEFAULT_CONFIG = {
    "标题": {
        "各级标题设置": {
            1: {"应用序号": True, "格式": "chinese"},
            2: {"应用序号": True, "格式": "chinese_bracket"},
            3: {"应用序号": True, "格式": "arabic_dot"},
            4: {"应用序号": True, "格式": "arabic_bracket"},
            5: {"应用序号": True, "格式": "arabic_dot"},
            6: {"应用序号": True, "格式": "arabic_bracket"},
            7: {"应用序号": True, "格式": "arabic_dot"},
            8: {"应用序号": True, "格式": "arabic_bracket"},
            9: {"应用序号": True, "格式": "arabic_dot"},
        }
    },
    "正文": {
        "中文字体": "宋体",
        "英文字体": "Times New Roman",
        "字号": 12,
        "段前间距": 12,
        "段后间距": 12,
        "行距": 1.5,
        "首行缩进": 0.5
    },
    "表格": {
        "中文字体": "宋体",
        "英文字体": "Times New Roman",
        "字号": 10,
        "段前间距": 6,
        "段后间距": 6,
        "表格宽度": 6
    }
}

# 初始化session state
if 'config' not in st.session_state:
    st.session_state.config = DEFAULT_CONFIG.copy()
if 'processed' not in st.session_state:
    st.session_state.processed = False

# 样式
st.markdown("""
<style>
    .main-header {
        font-size: 2.2rem;
        font-weight: bold;
        color: #1E3A8A;
        text-align: center;
        margin-bottom: 1.5rem;
        padding-top: 0.5rem;
    }
    .tab-header {
        font-size: 1.4rem;
        font-weight: 600;
        color: #374151;
        margin-top: 0;
        margin-bottom: 1.2rem;
    }
    .upload-box {
        border: 2px dashed #4F46E5;
        border-radius: 10px;
        padding: 2rem;
        text-align: center;
        background: linear-gradient(135deg, #667eea15 0%, #764ba215 100%);
        margin: 1rem 0;
        transition: all 0.3s ease;
    }
    .upload-box:hover {
        border-color: #4338CA;
        background: linear-gradient(135deg, #667eea25 0%, #764ba225 100%);
    }
    .stButton button {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border: none;
        padding: 0.75rem 2rem;
        border-radius: 8px;
        font-weight: 600;
        font-size: 1rem;
        transition: all 0.3s ease;
        width: 100%;
    }
    .stButton button:hover {
        transform: translateY(-2px);
        box-shadow: 0 8px 16px rgba(102, 126, 234, 0.2);
    }
    .success-box {
        background: linear-gradient(135deg, #10B981 0%, #059669 100%);
        color: white;
        padding: 1.5rem;
        border-radius: 10px;
        text-align: center;
        margin: 1rem 0;
        animation: fadeIn 0.5s ease-in;
    }
    .file-info {
        background: linear-gradient(135deg, #3B82F6 0%, #1D4ED8 100%);
        color: white;
        padding: 1.5rem;
        border-radius: 10px;
        margin: 1rem 0;
    }
    .config-table {
        width: 100%;
        margin-bottom: 1.5rem;
    }
    .config-table th {
        background: #F3F4F6;
        padding: 0.75rem;
        text-align: left;
        font-weight: 600;
        color: #374151;
        border-bottom: 2px solid #E5E7EB;
    }
    .config-table td {
        padding: 0.75rem;
        border-bottom: 1px solid #E5E7EB;
    }
    .config-table tr:hover {
        background: #F9FAFB;
    }
    .setting-row {
        display: flex;
        align-items: center;
        gap: 1rem;
        margin-bottom: 1rem;
    }
    .setting-label {
        min-width: 100px;
        font-weight: 500;
        color: #374151;
    }
    .setting-input {
        flex: 1;
    }
    @keyframes fadeIn {
        from { opacity: 0; transform: translateY(-10px); }
        to { opacity: 1; transform: translateY(0); }
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 0.5rem;
        padding: 0;
        margin-bottom: 1rem;
    }
    .stTabs [data-baseweb="tab"] {
        padding: 0.5rem 1.2rem;
        border-radius: 6px;
        background: #F9FAFB;
        border: 1px solid #E5E7EB;
        transition: all 0.3s ease;
        font-weight: 500;
    }
    .stTabs [data-baseweb="tab"][aria-selected="true"] {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border-color: #4F46E5;
    }
    .stTabs [data-baseweb="tab"]:hover {
        background: #F3F4F6;
    }
    .stTabs [aria-selected="true"]:hover {
        background: linear-gradient(135deg, #5a6fd8 0%, #6a4090 100%) !important;
    }
    .help-section {
        background: #F9FAFB;
        padding: 1.2rem;
        border-radius: 8px;
        margin: 0.8rem 0;
        border-left: 4px solid #4F46E5;
    }
    .main-container {
        max-width: 1200px;
        margin: 0 auto;
        padding: 0 1rem;
    }
</style>
""", unsafe_allow_html=True)

def get_outline_level_from_xml(p):
    """从段落的XML中提取大纲级别"""
    xml = p._p.xml
    m = re.search(r'<w:outlineLvl w:val="(\d)"/>', xml)
    level = int(m.group(1)) if m else None
    if level is not None:
        level += 1
    return level

def set_font(run, cz_font_name, font_name):
    """设置字体"""
    rPr = run.element.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn('w:eastAsia'), cz_font_name)
    rFonts.set(qn('w:ascii'), font_name)

def number_to_chinese(number):
    """数字转中文"""
    if number < 0 or number > 100:
        return str(number)
    
    chinese_numbers = ["零", "一", "二", "三", "四", "五", "六", "七", "八", "九"]
    
    if number < 10:
        return chinese_numbers[number]
    elif number < 20:
        return "十" + (chinese_numbers[number - 10] if number != 10 else "")
    elif number < 100:
        tens = number // 10
        ones = number % 10
        return chinese_numbers[tens] + "十" + (chinese_numbers[ones] if ones != 0 else "")
    else:
        return "一百"

def to_roman(num):
    """转换为罗马数字"""
    roman_map = [(1000, 'M'), (900, 'CM'), (500, 'D'), (400, 'CD'),
                 (100, 'C'), (90, 'XC'), (50, 'L'), (40, 'XL'),
                 (10, 'X'), (9, 'IX'), (5, 'V'), (4, 'IV'), (1, 'I')]
    result = ""
    for value, numeral in roman_map:
        while num >= value:
            result += numeral
            num -= value
    return result

def format_number(number, format_type):
    """根据格式类型格式化数字"""
    formats = {
        "chinese": lambda n: f"{number_to_chinese(n)}、",
        "chinese_bracket": lambda n: f"（{number_to_chinese(n)}）",
        "arabic_dot": lambda n: f"{n}.",
        "arabic_bracket": lambda n: f"（{n}）",
        "roman_lower": lambda n: f"{to_roman(n).lower()}.",
        "roman_upper": lambda n: f"{to_roman(n)}.",
        "alphabet_lower": lambda n: f"{chr(96 + n)}." if n <= 26 else f"{n}.",
        "alphabet_upper": lambda n: f"{chr(64 + n)}." if n <= 26 else f"{n}.",
    }
    return formats.get(format_type, lambda n: f"{n}.")(number)

def add_heading_numbers(doc, config):
    """根据配置添加标题序号"""
    heading_numbers = [0] * 9  # 最多9级标题
    heading_settings = config["标题"]["各级标题设置"]
    
    # 匹配常见序号格式
    number_pattern = re.compile(r'^[\d一二三四五六七八九十（）\.、\s]+')

    for paragraph in doc.paragraphs:
        if paragraph.style.name.startswith('Heading'):
            try:
                level = int(paragraph.style.name.split(' ')[1]) - 1
                
                # 检查该级别是否应用序号
                if level + 1 not in heading_settings or not heading_settings[level + 1]["应用序号"]:
                    continue
                
                # 清理原有序号
                paragraph.text = number_pattern.sub('', paragraph.text).strip()
                
                # 更新序号
                heading_numbers[level] += 1
                for i in range(level + 1, 9):
                    heading_numbers[i] = 0
                
                # 获取格式并添加序号
                format_type = heading_settings[level + 1]["格式"]
                number_str = format_number(heading_numbers[level], format_type)
                paragraph.text = number_str + paragraph.text
            except Exception:
                continue

def modify_document_format(doc, config):
    """修改文档格式"""
    # 处理正文
    body_config = config["正文"]
    for paragraph in doc.paragraphs:
        if not paragraph.style.name.startswith("Heading"):
            paragraph.paragraph_format.space_before = Pt(body_config['段前间距'])
            paragraph.paragraph_format.space_after = Pt(body_config['段后间距'])
            paragraph.paragraph_format.line_spacing = body_config['行距']
            paragraph.paragraph_format.first_line_indent = Inches(body_config['首行缩进'])
            for run in paragraph.runs:
                set_font(run, body_config['中文字体'], body_config['英文字体'])
                run.font.size = Pt(body_config['字号'])

    # 处理表格
    table_config = config["表格"]
    for table_obj in doc.tables:
        table_obj.width = Inches(table_config['表格宽度'])
        for row in table_obj.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        set_font(run, table_config['中文字体'], table_config['英文字体'])
                        run.font.size = Pt(table_config['字号'])
                    paragraph.paragraph_format.space_before = Pt(table_config['段前间距'])
                    paragraph.paragraph_format.space_after = Pt(table_config['段后间距'])

def process_document(uploaded_file, config):
    """处理上传的文档"""
    try:
        doc = Document(uploaded_file)
        
        # 转换大纲级别为标题样式
        for para in doc.paragraphs:
            outline_level = get_outline_level_from_xml(para)
            if outline_level is not None and para.style.name == 'Normal':
                if outline_level <= 9:
                    heading_style = f"Heading {outline_level}"
                    if heading_style in doc.styles:
                        para.style = doc.styles[heading_style]
        
        # 添加标题序号
        add_heading_numbers(doc, config)
        
        # 修改格式
        modify_document_format(doc, config)
        
        # 保存到内存
        output = BytesIO()
        doc.save(output)
        output.seek(0)
        return output
    except Exception as e:
        st.error(f"处理失败: {str(e)}")
        return None

def config_main():
    """主配置区域"""
    # 使用tabs组织三大类设置
    tab1, tab2, tab3 = st.tabs(["📝 标题设置", "📄 正文设置", "📊 表格设置"])
    
    with tab1:
        # 标题设置表格
        st.markdown('<div class="config-table">', unsafe_allow_html=True)
        
        # 表头
        st.markdown("""
        <table class="config-table">
            <thead>
                <tr>
                    <th>标题级别</th>
                    <th>应用序号</th>
                    <th>序号格式</th>
                </tr>
            </thead>
            <tbody>
        """, unsafe_allow_html=True)
        
        # 序号格式选项（直接显示效果）
        format_options = {
            "chinese": "一、",
            "chinese_bracket": "（一）",
            "arabic_dot": "1.",
            "arabic_bracket": "（1）",
            "roman_lower": "i.",
            "roman_upper": "I.",
            "alphabet_lower": "a.",
            "alphabet_upper": "A."
        }
        
        # 表格内容
        for level in range(1, 10):
            st.markdown(f"<tr>", unsafe_allow_html=True)
            
            # 标题级别列
            st.markdown(f"<td><strong>{level}级标题</strong></td>", unsafe_allow_html=True)
            
            # 应用序号列
            with st.container():
                col1, col2 = st.columns([1, 1])
                with col1:
                    st.markdown("</td><td>", unsafe_allow_html=True)
                    apply = st.checkbox(
                        "",
                        value=st.session_state.config["标题"]["各级标题设置"][level]["应用序号"],
                        key=f"apply_{level}",
                        label_visibility="collapsed"
                    )
                    st.session_state.config["标题"]["各级标题设置"][level]["应用序号"] = apply
                    st.markdown("</td>", unsafe_allow_html=True)
            
            # 序号格式列
            with st.container():
                with col2:
                    st.markdown("<td>", unsafe_allow_html=True)
                    current_format = st.session_state.config["标题"]["各级标题设置"][level]["格式"]
                    selected = st.selectbox(
                        "",
                        options=list(format_options.keys()),
                        format_func=lambda x: format_options[x],
                        index=list(format_options.keys()).index(current_format) if current_format in format_options else 2,
                        key=f"format_{level}",
                        label_visibility="collapsed"
                    )
                    st.session_state.config["标题"]["各级标题设置"][level]["格式"] = selected
                    st.markdown("</td>", unsafe_allow_html=True)
            
            st.markdown("</tr>", unsafe_allow_html=True)
        
        st.markdown("""
            </tbody>
        </table>
        </div>
        """, unsafe_allow_html=True)
    
    with tab2:
        # 正文设置 - 单行排列
        col1, col2 = st.columns(2)
        
        with col1:
            # 字体设置
            st.markdown('<div class="setting-row">', unsafe_allow_html=True)
            st.markdown('<div class="setting-label">中文字体</div>', unsafe_allow_html=True)
            st.markdown('<div class="setting-input">', unsafe_allow_html=True)
            st.session_state.config["正文"]["中文字体"] = st.text_input(
                "",
                value=st.session_state.config["正文"]["中文字体"],
                key="body_cz_font",
                label_visibility="collapsed"
            )
            st.markdown('</div></div>', unsafe_allow_html=True)
            
            st.markdown('<div class="setting-row">', unsafe_allow_html=True)
            st.markdown('<div class="setting-label">英文字体</div>', unsafe_allow_html=True)
            st.markdown('<div class="setting-input">', unsafe_allow_html=True)
            st.session_state.config["正文"]["英文字体"] = st.text_input(
                "",
                value=st.session_state.config["正文"]["英文字体"],
                key="body_en_font",
                label_visibility="collapsed"
            )
            st.markdown('</div></div>', unsafe_allow_html=True)
            
            # 字号与行距
            st.markdown('<div class="setting-row">', unsafe_allow_html=True)
            st.markdown('<div class="setting-label">字号</div>', unsafe_allow_html=True)
            st.markdown('<div class="setting-input">', unsafe_allow_html=True)
            st.session_state.config["正文"]["字号"] = st.number_input(
                "(pt)",
                min_value=6,
                max_value=72,
                value=int(st.session_state.config["正文"]["字号"]),
                key="body_font_size",
                label_visibility="collapsed"
            )
            st.markdown('</div></div>', unsafe_allow_html=True)
            
            st.markdown('<div class="setting-row">', unsafe_allow_html=True)
            st.markdown('<div class="setting-label">行距</div>', unsafe_allow_html=True)
            st.markdown('<div class="setting-input">', unsafe_allow_html=True)
            st.session_state.config["正文"]["行距"] = st.number_input(
                "",
                min_value=1.0,
                max_value=3.0,
                value=float(st.session_state.config["正文"]["行距"]),
                step=0.1,
                key="body_line_spacing",
                label_visibility="collapsed"
            )
            st.markdown('</div></div>', unsafe_allow_html=True)
        
        with col2:
            # 段落格式
            st.markdown('<div class="setting-row">', unsafe_allow_html=True)
            st.markdown('<div class="setting-label">段前间距</div>', unsafe_allow_html=True)
            st.markdown('<div class="setting-input">', unsafe_allow_html=True)
            st.session_state.config["正文"]["段前间距"] = st.number_input(
                "(pt)",
                min_value=0,
                max_value=100,
                value=int(st.session_state.config["正文"]["段前间距"]),
                key="body_before",
                label_visibility="collapsed"
            )
            st.markdown('</div></div>', unsafe_allow_html=True)
            
            st.markdown('<div class="setting-row">', unsafe_allow_html=True)
            st.markdown('<div class="setting-label">段后间距</div>', unsafe_allow_html=True)
            st.markdown('<div class="setting-input">', unsafe_allow_html=True)
            st.session_state.config["正文"]["段后间距"] = st.number_input(
                "(pt)",
                min_value=0,
                max_value=100,
                value=int(st.session_state.config["正文"]["段后间距"]),
                key="body_after",
                label_visibility="collapsed"
            )
            st.markdown('</div></div>', unsafe_allow_html=True)
            
            st.markdown('<div class="setting-row">', unsafe_allow_html=True)
            st.markdown('<div class="setting-label">首行缩进</div>', unsafe_allow_html=True)
            st.markdown('<div class="setting-input">', unsafe_allow_html=True)
            st.session_state.config["正文"]["首行缩进"] = st.number_input(
                "(英寸)",
                min_value=0.0,
                max_value=2.0,
                value=float(st.session_state.config["正文"]["首行缩进"]),
                step=0.1,
                key="body_indent",
                label_visibility="collapsed"
            )
            st.markdown('</div></div>', unsafe_allow_html=True)
    
    with tab3:
        # 表格设置 - 单行排列
        col1, col2 = st.columns(2)
        
        with col1:
            # 字体设置
            st.markdown('<div class="setting-row">', unsafe_allow_html=True)
            st.markdown('<div class="setting-label">中文字体</div>', unsafe_allow_html=True)
            st.markdown('<div class="setting-input">', unsafe_allow_html=True)
            st.session_state.config["表格"]["中文字体"] = st.text_input(
                "",
                value=st.session_state.config["表格"]["中文字体"],
                key="table_cz_font",
                label_visibility="collapsed"
            )
            st.markdown('</div></div>', unsafe_allow_html=True)
            
            st.markdown('<div class="setting-row">', unsafe_allow_html=True)
            st.markdown('<div class="setting-label">英文字体</div>', unsafe_allow_html=True)
            st.markdown('<div class="setting-input">', unsafe_allow_html=True)
            st.session_state.config["表格"]["英文字体"] = st.text_input(
                "",
                value=st.session_state.config["表格"]["英文字体"],
                key="table_en_font",
                label_visibility="collapsed"
            )
            st.markdown('</div></div>', unsafe_allow_html=True)
            
            # 字号设置
            st.markdown('<div class="setting-row">', unsafe_allow_html=True)
            st.markdown('<div class="setting-label">字号</div>', unsafe_allow_html=True)
            st.markdown('<div class="setting-input">', unsafe_allow_html=True)
            st.session_state.config["表格"]["字号"] = st.number_input(
                "(pt)",
                min_value=6,
                max_value=72,
                value=int(st.session_state.config["表格"]["字号"]),
                key="table_font_size",
                label_visibility="collapsed"
            )
            st.markdown('</div></div>', unsafe_allow_html=True)
        
        with col2:
            # 表格宽度
            st.markdown('<div class="setting-row">', unsafe_allow_html=True)
            st.markdown('<div class="setting-label">表格宽度</div>', unsafe_allow_html=True)
            st.markdown('<div class="setting-input">', unsafe_allow_html=True)
            st.session_state.config["表格"]["表格宽度"] = st.number_input(
                "(英寸)",
                min_value=1,
                max_value=20,
                value=int(st.session_state.config["表格"]["表格宽度"]),
                key="table_width",
                label_visibility="collapsed"
            )
            st.markdown('</div></div>', unsafe_allow_html=True)
            
            # 间距设置
            st.markdown('<div class="setting-row">', unsafe_allow_html=True)
            st.markdown('<div class="setting-label">段前间距</div>', unsafe_allow_html=True)
            st.markdown('<div class="setting-input">', unsafe_allow_html=True)
            st.session_state.config["表格"]["段前间距"] = st.number_input(
                "(pt)",
                min_value=0,
                max_value=100,
                value=int(st.session_state.config["表格"]["段前间距"]),
                key="table_before",
                label_visibility="collapsed"
            )
            st.markdown('</div></div>', unsafe_allow_html=True)
            
            st.markdown('<div class="setting-row">', unsafe_allow_html=True)
            st.markdown('<div class="setting-label">段后间距</div>', unsafe_allow_html=True)
            st.markdown('<div class="setting-input">', unsafe_allow_html=True)
            st.session_state.config["表格"]["段后间距"] = st.number_input(
                "(pt)",
                min_value=0,
                max_value=100,
                value=int(st.session_state.config["表格"]["段后间距"]),
                key="table_after",
                label_visibility="collapsed"
            )
            st.markdown('</div></div>', unsafe_allow_html=True)

def help_sidebar():
    """侧边栏帮助信息"""
    with st.sidebar:
        st.markdown("### 📖 使用说明")
        
        st.markdown('<div class="help-section">', unsafe_allow_html=True)
        st.markdown("**📤 上传文档**")
        st.markdown("""
        1. 点击上传区域选择.docx文件
        2. 文件大小建议不超过50MB
        """)
        st.markdown('</div>', unsafe_allow_html=True)
        
        st.markdown('<div class="help-section">', unsafe_allow_html=True)
        st.markdown("**⚙️ 配置说明**")
        st.markdown("""
        - **标题设置**：控制各级标题的自动编号
        - **正文设置**：调整文档正文的格式样式
        - **表格设置**：设置表格的字体和间距
        """)
        st.markdown('</div>', unsafe_allow_html=True)
        
        st.markdown('<div class="help-section">', unsafe_allow_html=True)
        st.markdown("**🚀 处理流程**")
        st.markdown("""
        1. 上传文档
        2. 设置参数
        3. 点击开始处理
        4. 下载文件
        """)
        st.markdown('</div>', unsafe_allow_html=True)
        
        st.divider()
        
        # 重置按钮
        if st.button("🔄 重置所有设置", use_container_width=True):
            st.session_state.config = DEFAULT_CONFIG.copy()
            st.success("设置已重置！")
            st.rerun()

def main():
    # 主容器
    st.markdown('<div class="main-container">', unsafe_allow_html=True)
    
    # 主标题
    st.markdown('<h1 class="main-header">📝 Word文档格式化工具</h1>', unsafe_allow_html=True)
    
    # 创建两列布局
    col1, col2 = st.columns([3, 1])
    
    with col1:
        # 上传区域
        uploaded_file = st.file_uploader(
            "上传Word文档",
            type=['docx'],
            help="选择需要格式化的.docx文件"
        )
        
        if uploaded_file:
            st.markdown(f'''
            <div class="file-info">
                <div style="font-size: 1.1rem; font-weight: 600; margin-bottom: 0.3rem;">
                    📄 {uploaded_file.name}
                </div>
                <div style="font-size: 0.85rem; opacity: 0.9;">
                    大小: {len(uploaded_file.getvalue()) / 1024:.1f} KB
                </div>
            </div>
            ''', unsafe_allow_html=True)
        
        # 参数设置区域
        config_main()
        
        # 处理按钮区域
        st.markdown("---")
        if uploaded_file:
            if st.button("🚀 开始处理文档", type="primary", use_container_width=True, key="process_button"):
                with st.spinner("正在处理文档，请稍候..."):
                    processed_doc = process_document(uploaded_file, st.session_state.config)
                    
                    if processed_doc:
                        st.session_state.processed = True
                        st.session_state.processed_data = processed_doc
                        st.session_state.output_filename = f"已处理_{uploaded_file.name}"
                        st.rerun()
        
        # 结果展示区域
        if st.session_state.processed:
            st.markdown('<div class="success-box">✅ 文档处理完成！</div>', unsafe_allow_html=True)
            
            col_a, col_b = st.columns([2, 1])
            with col_a:
                st.download_button(
                    label=f"📥 下载 {st.session_state.output_filename}",
                    data=st.session_state.processed_data.getvalue(),
                    file_name=st.session_state.output_filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            with col_b:
                if st.button("🔄 重新处理", use_container_width=True):
                    st.session_state.processed = False
                    st.rerun()
    
    with col2:
        # 侧边栏帮助信息
        help_sidebar()
    
    st.markdown('</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()

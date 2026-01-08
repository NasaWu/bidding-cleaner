import streamlit as st
import pandas as pd
import io
import time

# 引入你的转换逻辑
# 注意：你需要稍微修改一下 yunshu.py 和 general.py，让它们支持传入 DataFrame 或 file object
# 或者直接在这里 import 它们，这里假设我们调用它们的逻辑函数
import yunshu
import general

# ==========================================
# 1. 页面配置与 Apple 风格 CSS 定制
# ==========================================
st.set_page_config(
    page_title="Data Converter Pro",
    page_icon="✨",
    layout="centered"
)

# 自定义 CSS 实现 Apple 风格 (毛玻璃、圆角、阴影、SF字体)
st.markdown("""
<style>
    /* 全局字体设置，模仿 macOS */
    html, body, [class*="css"] {
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
    }
    
    /* 隐藏 Streamlit 默认的菜单和页脚 */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    
    /* 主容器卡片样式 */
    .stApp {
        background-color: #F5F5F7; /* Apple 浅灰背景 */
    }
    
    /* 标题样式 */
    h1 {
        font-weight: 700 !important;
        letter-spacing: -0.02em !important;
        color: #1D1D1F;
    }
    
    /* 上传组件样式优化 */
    .stFileUploader > div > div {
        border-radius: 12px;
        border: 1px dashed #d1d1d6;
        background-color: #ffffff;
        box-shadow: 0 4px 12px rgba(0,0,0,0.03);
    }
    
    /* 按钮通用样式 (Apple Blue) */
    .stButton > button {
        border-radius: 20px !important;
        background-color: #0071e3 !important;
        color: white !important;
        border: none !important;
        padding: 10px 24px !important;
        font-weight: 500 !important;
        box-shadow: 0 4px 6px rgba(0, 113, 227, 0.2);
        transition: all 0.2s ease;
    }
    
    .stButton > button:hover {
        background-color: #0077ED !important;
        box-shadow: 0 6px 12px rgba(0, 113, 227, 0.3);
        transform: scale(1.02);
    }

    /* 识别结果卡片 */
    .type-card {
        padding: 16px;
        border-radius: 12px;
        background: white;
        border: 1px solid #e5e5ea;
        box-shadow: 0 2px 8px rgba(0,0,0,0.04);
        margin-bottom: 20px;
        display: flex;
        align-items: center;
        gap: 10px;
    }
    
    .success-text { color: #34C759; font-weight: 600; }
    .info-text { color: #86868b; font-size: 14px; }
    
</style>
""", unsafe_allow_html=True)

# ==========================================
# 2. 逻辑函数
# ==========================================

def detect_file_type(file_obj, sheet_name):
    """
    读取表头来判断是 运输 还是 通用
    """
    try:
        # 只读取前几行用于判断，节省内存
        # header=[0, 1] 对应你之前的多级表头逻辑
        df = pd.read_excel(file_obj, sheet_name=sheet_name, header=[0, 1], nrows=5)
        
        # 将多级表头展平便于搜索
        # 例如: ('阿里巴巴', '份额比例') -> '阿里巴巴_份额比例'
        # 我们只需要看第二级表头（具体字段名）
        all_sub_columns = [str(col[1]).strip() for col in df.columns]
        
        # 判定逻辑
        # 运输表的特征字段: "车型", "物流组 (LC)"(可能带括号)
        # 通用表的特征字段: "规格型号", "是否租仓类"
        
        is_yunshu = any("车型" in col for col in all_sub_columns)
        is_general = any("规格型号" in col for col in all_sub_columns) or any("是否租仓类" in col for col in all_sub_columns)
        
        if is_yunshu:
            return "transport"
        elif is_general:
            return "general"
        else:
            return "unknown"
            
    except Exception as e:
        return f"error: {str(e)}"

# app.py

def process_file(file_obj, file_type, sheet_name):
    # 1. 临时保存
    temp_input_path = "temp_uploaded.xlsx"
    with open(temp_input_path, "wb") as f:
        f.write(file_obj.getbuffer())

    try:
        # 2. 调用函数时，把 sheet_name 传进去！
        if file_type == "transport":
            # ⬇️ 修改点在这里：增加了 sheet_name
            df_result = yunshu.transform_logistics_table_v3(temp_input_path, sheet_name)
        else:
            # ⬇️ 修改点在这里：增加了 sheet_name
            df_result = general.transform_general_table(temp_input_path, sheet_name)

        # 3. 写入内存 Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_result.to_excel(writer, index=False)
        
        return output.getvalue()

    except Exception as e:
        st.error(f"转换逻辑出错: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

# ==========================================
# 3. 前端 UI 布局
# ==========================================

st.title("招标数据转换中心")
st.markdown("<p class='info-text'>上传系统导出的 Excel 文件，自动识别并清洗格式。</p>", unsafe_allow_html=True)
st.markdown("---")

# A. 文件上传区
uploaded_file = st.file_uploader("拖拽文件到这里 或 点击上传", type=['xlsx', 'xls'])

# B. Sheet 设置区
col1, col2 = st.columns([1, 2])
with col1:
    sheet_name = st.text_input("Sheet 名称", value="Sheet1", help="默认为 Sheet1，如有不同请修改")

# C. 核心交互区
if uploaded_file is not None:
    # 1. 自动识别类型
    file_type = detect_file_type(uploaded_file, sheet_name)
    
    # 显示识别结果
    if file_type == "transport":
        st.markdown(f"""
        <div class="type-card">
            <span style="font-size: 20px;">🚛</span>
            <div>
                <div style="font-weight: 600; color: #1d1d1f;">已识别：运输/物流招标表</div>
                <div class="info-text">将使用 yunshu.py 引擎进行处理</div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
    elif file_type == "general":
        st.markdown(f"""
        <div class="type-card">
            <span style="font-size: 20px;">📦</span>
            <div>
                <div style="font-weight: 600; color: #1d1d1f;">已识别：通用/仓储招标表</div>
                <div class="info-text">将使用 general.py 引擎进行处理</div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
    elif "error" in file_type:
        st.error(f"读取文件失败，请检查 Sheet 名称是否正确。错误信息: {file_type}")
    else:
        st.warning("⚠️ 无法自动识别表格类型，请检查表头格式是否符合规范。")

    # 2. 转换按钮与动画
    if file_type in ["transport", "general"]:
        # 创建一个占位符，用于居中显示按钮
        col_action_1, col_action_2, col_action_3 = st.columns([1, 2, 1])
        
        with col_action_2:
            start_btn = st.button("开始清洗数据 ✨", use_container_width=True)
        
        if start_btn:
            # 进度条/Spinner 动画
            with st.spinner('正在启动 AI 引擎清洗数据...'):
                # 模拟一点点延迟让动画展示一下（更有仪式感）
                time.sleep(0.8) 
                
                # 执行转换
                result_data = process_file(uploaded_file, file_type, sheet_name)
                
            if result_data:
                st.balloons() # 撒花庆祝
                st.success("转换完成！数据已就绪。")
                
                # 3. 下载按钮
                file_label = "运输表" if file_type == "transport" else "通用表"
                st.download_button(
                    label=f"下载清洗后的{file_label} (.xlsx)",
                    data=result_data,
                    file_name=f"清洗结果_{uploaded_file.name}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

# 页脚留白
st.write("")
st.write("")
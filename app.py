import streamlit as st
import os
import tempfile
from pathlib import Path

# 导入你现有的逻辑
from converter import pdf_to_images, create_ppt, cleanup
# 导入 PPTX 转 PDF 的逻辑（直接引用函数比 subprocess 更稳定）
from pptx_to_pdf import extract_images_from_pptx, images_to_pdf

# --- 页面配置 ---
st.set_page_config(page_title="RedecktoPPT - 水印去除工具", page_icon="🚀")

st.title("🚀 RedecktoPPT")
st.markdown("""
通过 AI 检测和图像处理技术，自动去除 PDF 或 PPTX 文档底部的文字水印和 Logo，并生成干净的 PPTX 文件。
""")

# --- 侧边栏设置 ---
st.sidebar.header("配置参数")
bottom_height = st.sidebar.slider(
    "底部检测高度 (px)", 
    min_value=100, 
    max_value=500, 
    value=200, 
    help="用于界定底部水印区域的高度。如果水印较高，请调大此值。"
)
dpi = st.sidebar.select_slider(
    "转换清晰度 (DPI)", 
    options=[72, 100, 150, 200], 
    value=150,
    help="DPI 越高越清晰，但转换速度越慢。"
)

# --- 主界面 ---
uploaded_file = st.file_uploader("选择要处理的 PDF 或 PPTX 文件", type=['pdf', 'pptx'])

if uploaded_file is not None:
    file_ext = Path(uploaded_file.name).suffix.lower()
    
    # 使用 streamlit 的按钮触发转换
    if st.button("开始执行转换", type="primary"):
        # 创建临时工作目录
        with tempfile.TemporaryDirectory() as tmp_dir:
            input_path = os.path.join(tmp_dir, uploaded_file.name)
            output_pptx = os.path.join(tmp_dir, "output_fixed.pptx")
            
            # 保存上传的文件到临时目录
            with open(input_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            
            try:
                with st.status("正在处理中...", expanded=True) as status:
                    curr_input = input_path
                    
                    # 1. 处理 PPTX -> PDF 的预转换阶段
                    if file_ext == '.pptx':
                        st.write("🔄 正在从 PPTX 提取页面...")
                        temp_pdf = os.path.join(tmp_dir, "temp_converted.pdf")
                        
                        # 复用你 pptx_to_pdf.py 的逻辑
                        pptx_images = extract_images_from_pptx(input_path)
                        if not pptx_images:
                            st.error("无法从 PPTX 提取图片，请检查文件是否损坏。")
                            st.stop()
                        
                        images_to_pdf(pptx_images, temp_pdf)
                        # 清理提取出的零散图片
                        for img in pptx_images: 
                            if os.path.exists(img): os.remove(img)
                            
                        curr_input = temp_pdf
                        st.write("✅ PPTX 预处理完成")

                    # 2. 核心处理逻辑：PDF -> 去水印图片
                    st.write("📄 正在分析并去除水印...")
                    # 调用 converter.py 中的逻辑
                    # 注意：converter.py 里的 process_page 默认存入 /tmp/，
                    # 在云端环境是通用的，但这里我们直接调用它的 pdf_to_images
                    processed_images = pdf_to_images(curr_input, dpi=dpi, bottom_height=bottom_height)
                    
                    # 3. 生成最后的 PPTX
                    st.write("🎨 正在生成最终 PPT...")
                    create_ppt(processed_images, output_pptx)
                    
                    status.update(label="转换成功！", state="complete", expanded=False)

                # 4. 展示下载按钮
                with open(output_pptx, "rb") as f:
                    st.success("处理完成！请点击下方按钮下载。")
                    st.download_button(
                        label="📥 下载处理后的 PPTX",
                        data=f,
                        file_name=f"fixed_{Path(uploaded_file.name).stem}.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
                
                # 最后清理 converter 生成的图片
                cleanup(processed_images)

            except Exception as e:
                st.error(f"转换过程中出错: {str(e)}")
                st.info("提示：请确保你的仓库中包含 packages.txt 以安装 tesseract-ocr 系统库。")

else:
    st.info("请上传一个文件以开始。")

# --- 页脚 ---
st.divider()
st.caption("Powered by Streamlit & RedecktoPPT Logic")

# 🚀 RedecktoPPT

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge.svg)](https://redecktoppt.streamlit.app/) 
![GitHub stars](https://img.shields.io/github/stars/dawnguan123/redecktoppt?style=social)
![License](https://img.shields.io/github/license/dawnguan123/redecktoppt)
![Python Version](https://img.shields.io/badge/python-3.8+-blue.svg)

**RedecktoPPT** 是一款基于 AI 视觉和 OCR 技术的自动化办公辅助工具。它能够一键将带有水印或 Logo 的 PDF/PPTX 文档转换为“干净”的 PPTX 文件，特别针对 **NotebookLM** 生成的演示文档进行了深度优化。

> **💡 痛点解决**：AI 生成的文档虽然内容精良，但往往带有无法直接修改的水印。本项目通过计算机视觉技术实现自动识别与无缝覆盖填充。

---

## 🖼️ 效果展示

| 处理前 (带有水印) | 处理后 (无缝覆盖) |
| :--- | :--- |
| ![Original](https://via.placeholder.com/400x225.png?text=Original+with+Watermark) | ![Processed](https://via.placeholder.com/400x225.png?text=Clean+PPTX+Output) |
| *注：此处可替换为您实际处理的对比截图* |

---

## 🌟 核心亮点

- **智能去水印**：利用 **Tesseract OCR** 精确识别文档底部的文字水印（如“Powered by NotebookLM”等区域）。
- **Logo 边缘检测**：通过 **OpenCV** 的 Canny 边缘检测算法，精准定位并覆盖文档角落的各种 Logo。
- **无缝颜色填充**：采用智能采样技术，自动提取覆盖区域周围的颜色进行填充，确保转换后的页面美观自然。
- **全格式支持**：支持 PDF 和 PPTX 格式输入。对于 PPTX，程序会自动提取图片并重组，确保转换闭环。
- **Web 端即开即用**：基于 **Streamlit** 构建的图形界面，支持在线上传、参数微调及处理后文件直接下载。

---

## 🛠️ 技术栈

- **核心引擎**: [PyMuPDF (fitz)](https://github.com/pymupdf/PyMuPDF) —— 高效处理 PDF 渲染与转换。
- **图像识别**: [OpenCV (cv2)](https://opencv.org/) & [pytesseract](https://github.com/madmaze/pytesseract) —— 驱动特征检测与文字识别。
- **演示文稿处理**: [python-pptx](https://python-pptx.readthedocs.io/) —— 生成高质量 PPTX。
- **前端框架**: [Streamlit](https://streamlit.io/) —— 驱动敏捷的 Web 交互。

---

## 🚀 快速开始

### 方式一：直接访问 Web 版（推荐）
点击徽章进入部署在 Streamlit Cloud 上的地址： [https://redecktoppt.streamlit.app/](https://redecktoppt.streamlit.app/)

### 方式二：本地运行

1. **克隆仓库**
   ```bash
   git clone [https://github.com/dawnguan123/redecktoppt.git](https://github.com/dawnguan123/redecktoppt.git)
   cd redecktoppt

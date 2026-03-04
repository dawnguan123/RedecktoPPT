# 🚀 RedecktoPPT

将 PDF 转换为 PPT，自动去除水印和 Logo。

## 安装

```bash
pip install -r requirements.txt
```

## 使用

```bash
python simple.py 输入.pdf 输出.pptx [底部检测高度]
```

- **底部检测高度**：可选，默认 200px，用于检测水印/Logo 区域

```bash
# 示例
python simple.py input.pdf output.pptx
python simple.py input.pdf output.pptx 150  # 检测底部 150px
```

## 依赖

```
pymupdf>=1.23.0
python-pptx>=0.6.23
Pillow>=10.0.0
opencv-python>=4.8.0
pytesseract>=0.3.0
```

## 功能

- **PDF 转 PPT**：将 PDF 每页转换为图片，拼接为 PPT
- **自动去水印**：检测并去除 NotebookLM 等水印
- **自动去 Logo**：检测并覆盖底部 Logo
- **智能填充**：用周围颜色无缝填充覆盖区域

## 技术原理

1. **OCR 文字检测**：用 pytesseract 检测底部文字
2. **边缘检测**：用 OpenCV Canny 边缘检测识别 Logo
3. **智能合并**：只合并距离近的检测区域
4. **颜色填充**：采样周围颜色，无缝覆盖

## 适用场景

- NotebookLM 生成的音频总结 PDF
- 带水印/Logo 的导出文档
- 图片型 PDF（无文字层）

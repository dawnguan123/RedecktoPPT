# 🚀 RedecktoPPT

将 PDF 转换为 PPT，自动去除水印和 Logo。

## 安装

```bash
pip install -r requirements.txt
```

## 使用

支持 PDF 和 PPTX 两种输入格式：

```bash
python simple.py 输入.pdf 输出.pptx [底部检测高度]
python simple.py 输入.pptx 输出.pptx [底部检测高度]
```

- **底部检测高度**：可选，默认 200px，用于检测水印/Logo 区域

```bash
# 示例
python simple.py input.pdf output.pptx
python simple.py input.pptx output.pptx
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

### 1. OCR 文字检测
- 使用 pytesseract 检测底部文字
- 只检测 y > height - 150 的区域（水印区）

### 2. Logo 边缘检测
- 使用 OpenCV Canny 边缘检测
- 检测区域：500px x 150px（右下角）
- 只保留底部附近的检测结果（y1 > height - 50）

### 3. 智能合并
- 右侧 Logo 盒子全部合并
- 按距离合并其他检测结果
- 只有纯 OCR 来源才允许扩展到整行

### 4. 颜色填充
- 优先从 Logo 下方采样
- 用采样颜色填充覆盖区域

## 适用场景

- NotebookLM 生成的音频总结 PDF
- 带水印/Logo 的导出文档
- 图片型 PDF（无文字层）

## 测试通过

- `Data_Synergy.pdf` (5 页) ✅
- `OpenClaw_Autonomous_Digital_Twins.pdf` (5 页) ✅

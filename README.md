# 🚀 RedecktoPPT (简化版)

将 PDF 转换为 PPT 的轻量级工具。

## 安装

```bash
pip install -r requirements.txt
```

## 使用

```bash
python simple.py 输入.pdf 输出.pptx
```

## 依赖

- `pymupdf` - PDF 处理
- `python-pptx` - PPT 生成

## 原理

直接提取 PDF 页面为高清图片，拼接为 PPT。

适合 NotebookLM 生成的图片型 PDF。

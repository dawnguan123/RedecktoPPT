# 🚀 RedecktoPPT

让 NotebookLM 的演示文稿真正回归可编辑。

RedecktoPPT 是一款专为 macOS 设计的智能转换工具，旨在将不可编辑的 NotebookLM PDF 转换为全元素可编辑的高清 PowerPoint 演示文稿。不同于简单的截图拼凑，我们通过深度布局算法重新构建每一张幻灯片。

---

## ✨ 核心特性

### 🧠 深度布局解析 (Deep Layout Analysis)
集成 MinerU (magic-pdf) 引擎，精准识别标题、段落、列表及表格坐标。

### 🍎 Apple Silicon 优化
深度适配 Mac M1/M2/M3 芯片，利用 MPS 硬件加速 OCR 与模型推理。

### 🎨 智能聚类与重构
独创文本聚类合并算法，将零碎的 OCR 行自动组合为逻辑段落。

### 🎭 自动化排版引擎
根据内容密度自动匹配"封面、金句、内容、数据、封底"五大专业布局。

### 📂 高清原图恢复
直接从 PDF 媒体流提取高清资源，彻底告别低清截图。

---

## 📂 项目结构

```
RedecktoPPT/
├── core/                       # 核心引擎层
│   ├── miner_u_wrapper.py      # MinerU 解析包装器
│   ├── layout_refiner.py       # 文本聚类与语义合并
│   ├── coordinate_mapper.py    # PDF 到 PPT 坐标转换公式
│   ├── coordinate_transformer.py  # 坐标转换器 (EMU)
│   └── layout_engine.py        # 智能布局引擎
│
├── rendering/                  # 渲染输出层
│   └── pptx_creator.py       # 基于 python-pptx 的生成器
│
├── api/                       # API 层
│   ├── mineru_client.py       # MinerU API 客户端
│   └── miner_u_handler.py     # 响应处理器
│
├── templates/                  # 商业级 PPT 模版库
│   └── style.json
│
├── utils/                     # 工具模块
│   └── logger.py
│
├── config_check.py            # Mac 环境与 MPS 状态检测
├── main.py                   # 程序入口
└── requirements.txt           # 依赖清单
```

---

## 🚀 快速开始 (macOS)

### 1. 环境准备

```bash
# 克隆项目
git clone https://github.com/your-repo/RedecktoPPT.git
cd RedecktoPPT

# 创建虚拟环境
conda create -n redeckto python=3.10
conda activate redeckto

# 安装依赖
pip install -r requirements.txt
```

### 2. 环境检测

```bash
# 检测 Mac 环境与 MPS 状态
python config_check.py
```

### 3. 本地解析

程序将自动调用本地 MinerU 模型进行解析。若需要使用云端 API，请在 `.env` 中配置 `MINERU_TOKEN`。

```bash
# 运行转换
python main.py --input slides.pdf --output output.pptx
```

---

## 🗺️ 路线图 (Roadmap)

- [ ] **第一阶段**：实现 Mac 本地 OCR 与坐标还原（当前阶段）
- [ ] **第二阶段**：集成 LLM 驱动的语义校对
- [ ] **第三阶段**：开发基于 FastAPI 的 Web 商业版入口

---

## 📄 开源协议

本项目基于 MIT License 开源。

---

## 🙏 致谢

- [NotebookLM2PPT](https://github.com/elliottzheng/NotebookLM2PPT) - 自动化理念启发
- [MinerU](https://mineru.net/) - 深度布局分析能力
- [magic-pdf](https://github.com/T-H-90/magic-pdf) - PDF 解析引擎

# #!/usr/bin/env python3
# """
# RedecktoPPT - Main Entry Point
# 自动化流水线入口

# 作为首席架构师设计的"指挥中心"：
# - 串联解析、聚类、换算和渲染模块
# - 实现一键式 PDF → PPTX 转换
# - 预留商业化扩展空间

# 参考：
# - [cite: 30, 32] MinerU 深度解析
# - [cite: 38, 39] 文本聚类重组
# - [cite: 51, 62] 自动布局选择
# - [cite: 67] 业务约束（200MB/600页）
# """


# import sys

# # --- 终极全空间补丁：修复 Transformers 4.40+ 结构与接口双断代问题 ---
# import sys
# import types
# import transformers

# # 1. 动态构造缺失的模块路径
# def mock_module(path):
#     parts = path.split('.')
#     for i in range(len(parts)):
#         mod_path = '.'.join(parts[:i+1])
#         if mod_path not in sys.modules:
#             sys.modules[mod_path] = types.ModuleType(parts[i])
#     return sys.modules[path]

# # 修复 Roberta Tokenizer 路径丢失
# try:
#     from transformers.models.roberta import RobertaTokenizerFast
#     roberta_mod = mock_module('transformers.models.roberta.tokenization_roberta_fast')
#     roberta_mod.RobertaTokenizerFast = RobertaTokenizerFast
# except ImportError:
#     pass

# # 2. 补齐丢失的核心函数
# def find_pruneable_heads_and_indices(heads, n_heads, head_size, already_pruned_heads):
#     if len(heads) == 0: return [], []
#     mask = [1] * n_heads
#     for head in heads:
#         if head not in already_pruned_heads: mask[head] = 0
#     pruned_heads = sorted(list(already_pruned_heads.union(set(heads))))
#     index = [i for i, m in enumerate(mask) if m == 1]
#     return pruned_heads, index

# def prune_linear_layer(layer, index, dim=0):
#     import torch
#     index = index.to(layer.weight.device)
#     W = layer.weight.index_select(dim, index).clone().detach()
#     if layer.bias is not None:
#         b = layer.bias.index_select(0, index).clone().detach()
#     new_size = list(layer.weight.size())
#     new_size[dim] = len(index)
#     new_layer = torch.nn.Linear(new_size[1], new_size[0], bias=layer.bias is not None).to(layer.weight.device)
#     new_layer.weight.copy_(W.contiguous())
#     if layer.bias is not None:
#         new_layer.bias.copy_(b.contiguous())
#     return new_layer

# # 3. 饱和式注入 (Modeling, PyTorch Utils & 根目录)
# targets = [
#     transformers,
#     transformers.modeling_utils,
#     mock_module('transformers.pytorch_utils')
# ]

# for target in targets:
#     for func in [find_pruneable_heads_and_indices, prune_linear_layer]:
#         if not hasattr(target, func.__name__):
#             setattr(target, func.__name__, func)

# print("   🔧 Transformers 超级补丁已激活：虚拟目录结构已重组")
# # --------------------------------------------------

# # --- 追加补丁：解锁 Detectron2 配置严格模式 ---
# try:
#     from detectron2.config import CfgNode
#     # 核心逻辑：拦截 merge_from_file 方法，强制在加载前开启 set_new_allowed(True)
#     original_merge = CfgNode.merge_from_file
#     def patched_merge(self, cfg_filename, allow_unsafe=False):
#         self.set_new_allowed(True) # 允许 YAML 中出现未知 Key (如 BEIT)
#         return original_merge(self, cfg_filename, allow_unsafe)
#     CfgNode.merge_from_file = patched_merge
#     print("   🔧 Detectron2 配置严格模式已解锁")
# except Exception:
#     pass
# # --------------------------------------------


# # import sys
# import os
# import argparse
# from pathlib import Path
# from typing import Optional

# # 核心模块导入
# from core.miner_u_wrapper import MinerUWrapper
# from core.layout_refiner import LayoutRefiner
# from core.coordinate_mapper import CoordinateMapper
# from core.pptx_renderer import PptxRenderer


# class RedecktoPipeline:
#     """
#     RedecktoPPT 自动化流水线
    
#     全局视角：处理从文件读取到最终 PPT 交付的全生命周期
    
#     工作流程：
#     1. MinerU 深度解析 (PDF → JSON/Images)
#     2. 文本聚类重组 (碎片行 → 逻辑段落)
#     3. 坐标转换 (PDF → PPT EMU)
#     4. 布局选择与渲染 (逻辑块 → PPT 元素)
#     5. 导出交付 (PPTX)
#     """
    
#     def __init__(
#         self,
#         output_dir: str = "output",
#         aspect_ratio: str = "16:9"
#     ):
#         """
#         初始化流水线
        
#         Args:
#             output_dir: 输出目录
#             aspect_ratio: 宽高比 ("16:9" 或 "4:3")
#         """
#         self.output_dir = Path(output_dir)
#         self.output_dir.mkdir(parents=True, exist_ok=True)
        
#         self.aspect_ratio = aspect_ratio
        
#         # 初始化各核心组件
#         self.parser = MinerUWrapper(
#             output_base_dir=str(self.output_dir / "temp")
#         )
#         self.refiner = LayoutRefiner(line_gap_threshold=1.5)
#         self.renderer = PptxRenderer(aspect_ratio=aspect_ratio)
        
#         self.total_slides = 0
    
#     def run(self, pdf_path: str, force_fallback: bool = False) -> Optional[str]:
#         """
#         执行全流程转换
        
#         Args:
#             pdf_path: 输入 PDF 文件路径
            
#         Returns:
#             输出 PPTX 文件路径，失败返回 None
#         """
#         pdf_name = Path(pdf_path).stem
#         print(f"🎬 开始处理项目: {pdf_name}")
#         print("=" * 60)
        
#         try:
#             # ==================== Step 1: 深度解析 ====================
#             # MinerU 提供深度解析，能智能重排文本、统一字体并替换高清图
#             # [cite: 30, 32]
#             print("\n📌 Step 1: MinerU 深度解析")
#             parse_info = self.parser.process_pdf(pdf_path, force_fallback=force_fallback)
            
#             # 显示当前模式
#             mode = getattr(self.parser, 'current_mode', 'unknown')
#             mode_display = {
#                 'deep_parsing': '深度解析 (Deep Parsing)',
#                 'fallback': '备选方案 (Fallback)'
#             }.get(mode, mode)
#             print(f"   🔔 当前模式: {mode_display}")
            
#             json_path = parse_info.get("json_path", "")
#             images_dir = parse_info.get("images_dir", "")
#             page_count = parse_info.get("page_count", 1)
            
#             print(f"   ✅ JSON: {json_path}")
#             print(f"   ✅ Images: {images_dir}")
#             print(f"   ✅ Pages: {page_count}")
            
#             # ==================== Step 2: 逻辑重组与聚类 ====================
#             # 这一步是解决 NotebookLM 导出物为"图片集"痛点的关键
#             # [cite: 38, 39]
#             print("\n📌 Step 2: 文本聚类重组")
            
#             # 解析 JSON 获取数据
#             import json
#             with open(json_path, 'r', encoding='utf-8') as f:
#                 raw_data = json.load(f)
            
#             # 转换为布局引擎格式
#             refined_data = self._prepare_page_data(raw_data, page_count)
            
#             print(f"   ✅ 聚类完成: {len(refined_data)} 页")
            
#             # ==================== Step 3: 渲染生成 ====================
#             # 根据内容自动选择布局，如封面、内容页或封底
#             # [cite: 51, 62]
#             print("\n📌 Step 3: PPTX 渲染")
            
#             # 获取 PDF 尺寸
#             pdf_width = raw_data.get('width', 595)
#             pdf_height = raw_data.get('height', 842)
            
#             # 设置 PDF 源（用于裁剪图片）
#             self.renderer.set_pdf_source(pdf_path)
            
#             # 逐页渲染
#             for idx, page_data in enumerate(refined_data):
#                 # 构建页面数据（组件化格式）
#                 page_dict = {
#                     'blocks': page_data.get('blocks', []),
#                     'images': page_data.get('images', []),
#                     'width': pdf_width,
#                     'height': pdf_height
#                 }
                
#                 self.renderer.create_slide(
#                     page_data=page_dict,
#                     slide_index=idx,
#                     total_pages=len(refined_data),
#                     pdf_path=pdf_path
#                 )
#                 print(f"   📄 Page {idx + 1}/{len(refined_data)}")
            
#             # ==================== Step 4: 导出 ====================
#             output_file = self.output_dir / f"{pdf_name}_Editable.pptx"
#             self.renderer.save(str(output_file))
            
#             self.total_slides = len(refined_data)
            
#             print("\n" + "=" * 60)
#             print(f"🏁 任务圆满完成！")
#             print(f"   📁 文件: {output_file}")
#             print(f"   📊 页数: {self.total_slides}")
#             print("=" * 60)
            
#             return str(output_file)
            
#         except ValueError as ve:
#             # 业务约束错误：如超过 200MB 或 600 页限制
#             # [cite: 67]
#             print(f"\n⚠️ 业务约束错误: {ve}")
#             return None
            
#         except FileNotFoundError as fe:
#             print(f"\n⚠️ 文件未找到: {fe}")
#             return None
            
#         except Exception as e:
#             print(f"\n🔥 系统崩溃异常: {e}")
#             import traceback
#             traceback.print_exc()
#             return None
    
#     def _prepare_page_data(self, raw_data: dict, page_count: int) -> list:
#         """
#         准备页面数据
        
#         将 MinerU JSON 转换为渲染器需要的格式
        
#         Args:
#             raw_data: 原始解析数据
#             page_count: 页数
            
#         Returns:
#             页面数据列表
#         """
#         pages_data = []
        
#         # 尝试多种数据格式
#         pages = raw_data.get('pages', [])
        
#         for page_idx in range(page_count):
#             page_info = {'blocks': [], 'images': []}
            
#             # 如果有分页数据
#             if page_idx < len(pages):
#                 page = pages[page_idx]
                
#                 # 提取文本块
#                 for block in page.get('blocks', []):
#                     text = block.get('text', '').strip()
#                     if not text:
#                         continue
                    
#                     bbox = block.get('bbox', [0, 0, 0, 0])
#                     x1, y1, x2, y2 = bbox
                    
#                     # 估计字体大小
#                     height = y2 - y1
#                     font_size = height / 1.2 if height > 0 else 12
                    
#                     # 判断类型
#                     block_type = 'text'
#                     if block.get('type', '').lower() in ('title', 'heading'):
#                         block_type = 'title'
                    
#                     page_info['blocks'].append({
#                         'type': block_type,
#                         'text': text,
#                         'x': x1,
#                         'y': y1,
#                         'width': x2 - x1,
#                         'height': y2 - y1,
#                         'font_size': font_size
#                     })
                
#                 # 提取图片
#                 for img in page.get('images', []):
#                     bbox = img.get('bbox', [0, 0, 0, 0])
#                     page_info['images'].append({
#                         'path': img.get('path', ''),
#                         'x': bbox[0],
#                         'y': bbox[1],
#                         'width': bbox[2] - bbox[0],
#                         'height': bbox[3] - bbox[1]
#                     })
            
#             pages_data.append(page_info)
        
#         return pages_data


# def main():
#     """主入口"""
#     parser = argparse.ArgumentParser(
#         description="RedecktoPPT - 将 NotebookLM PDF 转换为可编辑 PPT"
#     )
    
#     parser.add_argument(
#         'input',
#         nargs='?',
#         help='输入 PDF 文件路径'
#     )
    
#     parser.add_argument(
#         '-o', '--output',
#         default='output',
#         help='输出目录 (default: output)'
#     )
    
#     parser.add_argument(
#         '-a', '--aspect',
#         choices=['16:9', '4:3'],
#         default='16:9',
#         help='PPT 宽高比 (default: 16:9)'
#     )
    
#     parser.add_argument(
#         '-v', '--verbose',
#         action='store_true',
#         help='详细输出'
#     )
    
#     parser.add_argument(
#         '--fallback',
#         action='store_true',
#         help='强制使用备选方案（当 magic-pdf 配置缺失时）'
#     )
    
#     args = parser.parse_args()
    
#     # 检查参数
#     if not args.input:
#         print("Usage: python main.py <path_to_pdf> [-o output_dir] [-a 16:9|4:3]")
#         print("\n示例:")
#         print("  python main.py slides.pdf")
#         print("  python main.py slides.pdf -o my_output -a 4:3")
#         sys.exit(1)
    
#     # 检查输入文件
#     if not os.path.exists(args.input):
#         print(f"❌ 文件不存在: {args.input}")
#         sys.exit(1)
    
#     # 创建流水线
#     pipeline = RedecktoPipeline(
#         output_dir=args.output,
#         aspect_ratio=args.aspect
#     )
    
#     # 执行转换
#     try:
#         result = pipeline.run(args.input, force_fallback=args.fallback)
        
#         if result:
#             sys.exit(0)
#         else:
#             sys.exit(1)
            
#     except Exception as e:
#         print(f"\n❌ 错误: {e}")
#         sys.exit(1)


# if __name__ == "__main__":
#     main()



#!/usr/bin/env python3
"""
RedecktoPPT - Main Entry Point
GitHub 完美发布版 (V22.0)

核心特性：
- 显式参数桥接 (解决 TypeError: multiple values for detection)
- 动态类继承注入 (解决 AssertionError: Backbone)
- 0.72 DPI 坐标对齐协议
- 全量 Apple Silicon MPS (GPU) 硬件加速
"""

import sys
import os
import types
import argparse
from pathlib import Path
from typing import Optional

# ==============================================================================
# 🧩 阶段一：深度内核补丁 (Kernel Patches)
# ==============================================================================

# 1. Transformers 4.40+ 结构补丁
import transformers
import transformers.modeling_utils

def mock_module(path):
    parts = path.split('.')
    for i in range(len(parts)):
        mod_path = '.'.join(parts[:i+1])
        if mod_path not in sys.modules:
            sys.modules[mod_path] = types.ModuleType(parts[i])
    return sys.modules[path]

try:
    from transformers.models.roberta import RobertaTokenizerFast
    for path in ['transformers.models.roberta.tokenization_roberta_fast', 'transformers.pytorch_utils']:
        m = mock_module(path)
        if 'roberta' in path: m.RobertaTokenizerFast = RobertaTokenizerFast
except ImportError: pass

def find_pruneable_heads_and_indices(heads, n_heads, head_size, already_pruned_heads):
    if len(heads) == 0: return [], []
    mask = [1] * n_heads
    for head in heads:
        if head not in already_pruned_heads: mask[head] = 0
    return sorted(list(already_pruned_heads.union(set(heads)))), [i for i, m in enumerate(mask) if m == 1]

def prune_linear_layer(layer, index, dim=0):
    import torch
    idx = index.to(layer.weight.device)
    W = layer.weight.index_select(dim, idx).clone().detach()
    new_layer = torch.nn.Linear(layer.in_features, len(index) if dim==0 else layer.out_features, 
                                bias=layer.bias is not None).to(layer.weight.device)
    new_layer.weight.copy_(W.contiguous())
    return new_layer

for target in [transformers, transformers.modeling_utils, sys.modules.get('transformers.pytorch_utils')]:
    if target:
        for func in [find_pruneable_heads_and_indices, prune_linear_layer]:
            if not hasattr(target, func.__name__): setattr(target, func.__name__, func)

# 2. LayoutLMv3 内核与权重绑定补丁
try:
    from magic_pdf.model.sub_modules.layout.layoutlmv3.layoutlmft.models.layoutlmv3.modeling_layoutlmv3 import LayoutLMv3Model
    if not hasattr(LayoutLMv3Model, "all_tied_weights_keys"):
        LayoutLMv3Model.all_tied_weights_keys = property(lambda self: {})
except Exception: pass

# 3. Detectron2 视觉全链路协议对齐 (V22.0 核心修复)
try:
    from detectron2.config import CfgNode
    from detectron2.modeling import BACKBONE_REGISTRY, Backbone
    from detectron2.layers import ShapeSpec
    
    # 强制解锁配置并锁定 MPS GPU 设备
    original_merge = CfgNode.merge_from_file
    def patched_merge(self, f, **k):
        self.set_new_allowed(True)
        if hasattr(self, "MODEL"): self.MODEL.DEVICE = "mps"
        return original_merge(self, f, **k)
    CfgNode.merge_from_file = patched_merge

    if "build_beit_backbone" not in BACKBONE_REGISTRY:
        from magic_pdf.model.sub_modules.layout.layoutlmv3.backbone import LayoutLMv3Model as RawModel
        from magic_pdf.model.sub_modules.layout.layoutlmv3.layoutlmft.models.layoutlmv3.configuration_layoutlmv3 import LayoutLMv3Config
        
        # 建立具备 RPN/ROI 期待特征层接口且继承 Backbone 的模型
        class FinalPatchedBackbone(RawModel, Backbone):
            def __init__(self, config, **kwargs):
                super().__init__(config, **kwargs)
            
            def output_shape(self):
                # 协议对齐，解决 KeyError 'res4'
                return {
                    "res2": ShapeSpec(channels=768, stride=4),
                    "res3": ShapeSpec(channels=768, stride=8),
                    "res4": ShapeSpec(channels=768, stride=16),
                    "res5": ShapeSpec(channels=768, stride=32),
                }
            @property
            def size_divisibility(self): return 32

        def bridge_builder(cfg, input_shape):
            # 显式使用关键字参数，解决 TypeError (multiple values for detection)
            hf_config = LayoutLMv3Config(
                max_2d_position_embeddings=1024, coordinate_size=128, shape_size=128,
                hidden_size=768, num_attention_heads=12, num_hidden_layers=12,
                has_relative_attention_bias=True, has_spatial_attention_bias=True
            )
            hf_config.name_or_path = "layoutlmv3-base"
            # 物理实例化，跳过 input_shape 以防位置冲突
            return FinalPatchedBackbone(config=hf_config, detection=True, 
                                        out_features=["layer2", "layer5", "layer8", "layer11"])
        
        BACKBONE_REGISTRY._obj_map["build_beit_backbone"] = bridge_builder
        print("   ✅ V22.0 全链路协议桥接器已激活 (显式参数对齐)")
except Exception as e:
    print(f"   ⚠️ 视觉系统补丁异常: {e}")

print("   🚀 RedecktoPPT 满血版引擎启动")
print("-" * 60)

# ==============================================================================
# 🏗️ 阶段二：业务逻辑 (RedecktoPipeline)
# ==============================================================================

from core.miner_u_wrapper import MinerUWrapper
from core.pptx_renderer import PptxRenderer

class RedecktoPipeline:
    def __init__(self, output_dir: str = "output", aspect_ratio: str = "16:9"):
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.parser = MinerUWrapper(output_base_dir=str(self.output_dir / "temp"))
        self.renderer = PptxRenderer(aspect_ratio=aspect_ratio)
    
    def run(self, pdf_path: str) -> Optional[str]:
        pdf_stem = Path(pdf_path).stem
        print(f"🎬 正在执行全量重构: {pdf_stem}.pdf")
        try:
            # Step 1: AI 深度解析 (触发 V22.0 补丁)
            print("\n📌 Step 1: AI 神经网络版面分析 (GPU 加速)")
            parse_info = self.parser.process_pdf(pdf_path)
            json_path = parse_info.get("json_path", "")
            
            # Step 2: 渲染矢量 PPT
            print("\n📌 Step 2: 执行 PPTX 矢量重构渲染")
            import json
            with open(json_path, 'r', encoding='utf-8') as f:
                raw_data = json.load(f)
            
            self.renderer.set_pdf_source(pdf_path)
            pages = raw_data.get('pages', [])
            for idx, page in enumerate(pages):
                # 坐标系对齐：PPT 720x405 归一化空间
                page_dict = {
                    'blocks': page.get('blocks', []),
                    'images': page.get('images', []),
                    'width': raw_data.get('width', 720),
                    'height': raw_data.get('height', 405)
                }
                self.renderer.create_slide(page_dict, idx, len(pages), pdf_path)
                print(f"   📄 Page {idx + 1}/{len(pages)} 渲染完成")
            
            output_file = self.output_dir / f"{pdf_stem}_Editable.pptx"
            self.renderer.save(str(output_file))
            print(f"\n✨ 恭喜！深度重构成功，原生 PPT 已生成: {output_file}")
            return str(output_file)
            
        except Exception as e:
            print(f"\n🔥 运行时崩溃: {e}")
            import traceback; traceback.print_exc()
            return None

def main():
    parser = argparse.ArgumentParser(description="RedecktoPPT: 原生矢量 PPT 转换引擎")
    parser.add_argument('input', help='PDF 文件路径')
    args = parser.parse_args()
    
    if not os.path.exists(args.input):
        print(f"❌ 找不到文件: {args.input}")
        sys.exit(1)
        
    pipeline = RedecktoPipeline()
    pipeline.run(args.input)

if __name__ == "__main__":
    main()




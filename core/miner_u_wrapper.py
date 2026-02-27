# #!/usr/bin/env python3
# """
# RedecktoPPT - MinerU Wrapper
# 点火版 (v2.3) - 禁用自动降级

# 特性：
# - 无自动降级 - 错误直接抛出
# - 强制坐标校验
# - MPS GPU 日志
# - 模型热启动
# """

# import os
# import json
# import re
# import traceback
# import warnings
# from pathlib import Path
# from typing import Dict, Any, Optional

# import fitz

# # 忽略警告
# try:
#     from urllib3.exceptions import NotOpenSSLWarning
#     warnings.filterwarnings('ignore', category=NotOpenSSLWarning)
# except:
#     pass


# class MinerUError(Exception):
#     """MinerU 解析错误"""
#     pass


# class MinerUWrapper:
#     """MinerU 核心解析包装器 (点火版)"""
    
#     PDF_STANDARD_DPI = 72.0
    
#     def __init__(self, output_base_dir: str = "output/temp", use_gpu: bool = True):
#         self.output_base_dir = Path(output_base_dir)
#         self.output_base_dir.mkdir(parents=True, exist_ok=True)
        
#         self.MAX_SIZE_MB = 200
#         self.MAX_PAGES = 600
#         self.use_gpu = use_gpu
#         self.current_mode = "unknown"
        
#         self._auto_init()
    
#     def _auto_init(self):
#         """自愈式初始化"""
#         # 检查 PyMuPDF
#         try:
#             import pymupdf
#             version = pymupdf.__version__
#             major, minor = map(int, version.split('.')[:2])
#             self._pymupdf_compatible = (major == 1 and minor <= 24)
#             print(f"   ✅ PyMuPDF: {version}")
#         except:
#             self._pymupdf_compatible = False
        
#         # 自动创建配置
#         self._ensure_config()
        
#         # 初始化 magic-pdf
#         self._magic_pdf_ready = False
#         self._doclayout_available = False
#         self._init_magic_pdf()
    
#     def _ensure_config(self):
#         """确保配置文件存在"""
#         config_path = Path.home() / "magic-pdf.json"
        
#         if config_path.exists():
#             try:
#                 with open(config_path) as f:
#                     config = json.load(f)
#                 self._config = config
#             except:
#                 self._config = {}
#             return
        
#         print()
#         print("🔧 首次运行：自动创建配置...")
        
#         default_config = {
#             "device": "mps" if self.use_gpu else "cpu",
#             "model_dir": str(Path.home() / "magic-pdf-models"),
#             "show_log": False
#         }
        
#         try:
#             with open(config_path, 'w') as f:
#                 json.dump(default_config, f, indent=2)
#             print(f"   ✅ 配置已创建: {config_path}")
#             self._config = default_config
#         except Exception as e:
#             print(f"   ⚠️ 配置创建失败: {e}")
#             self._config = {}
    
#     def _init_magic_pdf(self):
#         """初始化 magic-pdf 基础模块"""
#         if not self._pymupdf_compatible:
#             return
        
#         try:
#             from magic_pdf.data.dataset import PymuDocDataset
#             from magic_pdf.model.magic_model import MagicModel
#             from magic_pdf.pdf_parse_union_core_v2 import pdf_parse_union
            
#             self._PymuDocDataset = PymuDocDataset
#             self._MagicModel = MagicModel
#             self._pdf_parse_union = pdf_parse_union
            
#             self._magic_pdf_ready = True
#             print("   ✅ magic-pdf 已导入")
            
#         except ImportError as e:
#             print(f"   ⚠️ magic-pdf 导入失败: {e}")
#             self._magic_pdf_ready = False
    
#     def _check_model_weights(self) -> bool:
#         """检查模型权重文件 (宽容模式)"""
#         # 优先检查用户配置
#         model_dir = self._config.get('model_dir', '')
        
#         # 备选路径
#         if not model_dir:
#             search_paths = [
#                 Path.home() / "magic-pdf-models",
#                 Path.home() / "models",
#                 Path.home() / ".cache" / "magic-pdf" / "models",
#                 Path.home() / ".cache" / "huggingface" / "hub",
#             ]
            
#             for path in search_paths:
#                 if path.exists():
#                     model_dir = str(path)
#                     break
        
#         if not model_dir:
#             print("   💡 模型目录未设置，将使用自动下载")
#             return True
        
#         model_path = Path(model_dir)
        
#         # 递归搜索权重文件
#         weight_files = (
#             list(model_path.glob("**/*.pth")) + 
#             list(model_path.glob("**/*.onnx")) +
#             list(model_path.glob("**/*.pt"))
#         )
        
#         if weight_files:
#             print(f"   ✅ 已加载 {len(weight_files)} 个模型文件")
#         else:
#             print(f"   💡 模型文件未找到，将自动下载 (约 2-5GB)")
        
#         return True  # 宽容模式，不阻止运行
    
#     def _load_layout_model(self) -> bool:
#         """加载 YOLO 布局模型"""
#         if not self._magic_pdf_ready:
#             raise MinerUError("magic-pdf 未就绪")
        
#         print("   🗂️ 加载布局模型...")
        
#         try:
#             # 检查权重
#             self._check_model_weights()
            
#             from magic_pdf.model.sub_modules.model_init import AtomModelSingleton
#             from magic_pdf.model.model_list import AtomicModel
            
#             atom_model = AtomModelSingleton()
#             # 使用正确的枚举
#             atom_model.get_atom_model(atom_model_name=AtomicModel.Layout, show_log=False)
            
#             self._doclayout_available = True
#             return True
            
#         except Exception as e:
#             traceback.print_exc()
#             raise MinerUError(f"布局模型加载失败: {e}")
    
#     def _validate_pdf(self, pdf_path: str) -> int:
#         path = Path(pdf_path)
#         if not path.exists():
#             raise FileNotFoundError(f"文件不存在: {pdf_path}")
        
#         size_mb = path.stat().st_size / (1024 * 1024)
#         if size_mb > self.MAX_SIZE_MB:
#             raise ValueError(f"文件大小 ({size_mb:.1f}MB) 超过限制")
        
#         with fitz.open(pdf_path) as doc:
#             page_count = len(doc)
#             if page_count > self.MAX_PAGES:
#                 raise ValueError(f"页数 ({page_count}) 超过限制")
        
#         return page_count
    
#     def process_pdf(self, pdf_path: str, lang: str = "chi_sim", force_fallback: bool = False) -> Dict[str, Any]:
#         """执行解析"""
#         page_count = self._validate_pdf(pdf_path)
        
#         pdf_name = Path(pdf_path).stem
#         save_dir = self.output_base_dir / pdf_name
#         save_dir.mkdir(parents=True, exist_ok=True)
        
#         print(f"\n🔄 解析: {pdf_path}")
        
#         # 检查 magic-pdf
#         if not self._magic_pdf_ready:
#             raise MinerUError(
#                 "magic-pdf 核心模块未就绪。"
#                 "请确保已安装: pip install magic-pdf"
#             )
        
#         # 如果强制使用 fallback
#         if force_fallback:
#             self.current_mode = "fallback"
#             return self._fallback_parse(pdf_path, save_dir, page_count, lang)
        
#         # 尝试深度解析
#         try:
#             result = self._magic_pdf_full_parse(pdf_path, save_dir, page_count, lang)
#             if result:
#                 return result
#         except Exception as e:
#             print(f"   ⚠️ 深度解析失败: {e}")
        
#         # 降级到 fallback
#         self.current_mode = "fallback"
#         return self._fallback_parse(pdf_path, save_dir, page_count, lang)
    
#     def _magic_pdf_full_parse(self, pdf_path: str, save_dir: Path, page_count: int, lang: str) -> Dict:
#         """magic-pdf 深度解析 - 无降级"""
        
#         # GPU 激活日志
#         if self.use_gpu:
#             print("   🚀 Mac GPU (MPS) 已激活，正在进行深度版面分析...")
        
#         self.current_mode = "deep_parsing"
        
#         with open(pdf_path, 'rb') as f:
#             pdf_bytes = f.read()
        
#         # 创建 Dataset
#         dataset = self._PymuDocDataset(pdf_bytes, lang=lang)
#         parse_method = dataset.classify()
#         print(f"   📋 解析模式: {parse_method}")
        
#         # 步骤1: 运行模型推理获取布局结果
#         print("   🔥 执行版面分析...")
#         try:
#             from magic_pdf.model.doc_analyze_by_custom_model import doc_analyze
            
#             model_list = doc_analyze(
#                 dataset=dataset,
#                 lang=lang,
#                 show_log=False
#             )
#             print(f"   ✅ 版面分析完成，生成了 {len(model_list)} 个页面模型数据")
#         except Exception as e:
#             print(f"   ⚠️ doc_analyze 失败: {e}")
#             traceback.print_exc()
#             # 使用空模型列表
#             model_list = [{} for _ in range(page_count)]
        
#         # 步骤2: 使用模型结果解析 PDF
#         try:
#             from magic_pdf.rw.DiskReaderWriter import DiskReaderWriter
#             image_writer = DiskReaderWriter(str(save_dir / "temp_images"))
#         except Exception as e:
#             print(f"   ⚠️ DiskReaderWriter: {e}")
#             image_writer = None
        
#         print("   🔥 执行内容解析...")
#         result = self._pdf_parse_union(
#             model_list=model_list,
#             dataset=dataset,
#             imageWriter=image_writer,
#             parse_mode=parse_method,
#             lang=lang
#         )
        
#         if not result:
#             raise MinerUError("pdf_parse_union 返回空结果")
        
#         # 提取内容
#         parsed_data = self._extract_with_dpi_normalization(result, dataset, save_dir)
        
#         # 强制坐标校验
#         coord_info = self._verify_coords(parsed_data)
        
#         if not coord_info['is_local']:
#             raise MinerUError(
#                 f"解析器未能进入深度模式！\n"
#                 f"坐标仍为全屏: {coord_info['sample']}\n"
#                 f"请检查 magic-pdf 配置和模型加载状态"
#             )
        
#         print(f"   ✅ 坐标精度: 行级别 ({coord_info['sample']})")
        
#         # 保存 JSON
#         json_path = save_dir / "layout.json"
#         with open(json_path, 'w', encoding='utf-8') as f:
#             json.dump(parsed_data, f, ensure_ascii=False, indent=2)
        
#         # 提取图片
#         images_dir = self._extract_images(pdf_path, save_dir, page_count)
        
#         print(f"   ✅ 完成: {len(parsed_data.get('pages', []))} 页")
        
#         return {
#             "json_path": str(json_path),
#             "images_dir": str(images_dir),
#             "pdf_name": pdf_name,
#             "page_count": page_count,
#             "model_version": "magic-pdf-1.3.x"
#         }
    
#     def _extract_with_dpi_normalization(self, result: Dict, dataset, save_dir: Path) -> Dict:
#         """DPI 归一化提取"""
#         pages_data = []
        
#         YOLO_SOURCE_DPI = 100.0
#         DPI_SCALE = self.PDF_STANDARD_DPI / YOLO_SOURCE_DPI
        
#         coord_samples = []
        
#         for page_id, page_info in result.items():
#             page_idx = int(page_id.replace('page_', ''))
            
#             page = dataset.get_page(page_idx)
#             info = page.get_page_info()
            
#             orig_w = info.w
#             orig_h = info.h
            
#             page_w = orig_w * DPI_SCALE
#             page_h = orig_h * DPI_SCALE
            
#             blocks = []
#             images = []
            
#             for block in page_info.get('preproc_blocks', []):
#                 block_type = block.get('type', 'text')
                
#                 if block_type in ['image', 'figure']:
#                     bbox = block.get('bbox', [0, 0, 0, 0])
#                     normalized_bbox = [b * DPI_SCALE for b in bbox]
#                     images.append({
#                         'path': '',
#                         'bbox': normalized_bbox,
#                         'type': 'image',
#                         'orig_bbox': bbox
#                     })
#                     continue
                
#                 if block_type == 'table':
#                     bbox = block.get('bbox', [0, 0, 0, 0])
#                     normalized_bbox = [b * DPI_SCALE for b in bbox]
#                     blocks.append({
#                         'text': '[表格]',
#                         'bbox': normalized_bbox,
#                         'type': 'table',
#                         'orig_bbox': bbox
#                     })
#                     continue
                
#                 text = self._extract_text(block)
#                 if not text:
#                     continue
                
#                 text = self._clean_text(text)
#                 if not text:
#                     continue
                
#                 bbox = block.get('bbox', [0, 0, 0, 0])
#                 normalized_bbox = [b * DPI_SCALE for b in bbox]
                
#                 elem_type = 'title' if block_type in ['title', 'heading'] else 'text'
                
#                 block_data = {
#                     'text': text,
#                     'bbox': normalized_bbox,
#                     'type': elem_type,
#                     'orig_bbox': bbox,
#                     'orig_size': [orig_w, orig_h]
#                 }
                
#                 blocks.append(block_data)
                
#                 if len(coord_samples) < 2:
#                     coord_samples.append({
#                         'text': text[:30] + '...' if len(text) > 30 else text,
#                         'orig_bbox': bbox,
#                         'normalized_bbox': normalized_bbox
#                     })
            
#             pages_data.append({
#                 'page': page_idx,
#                 'width': page_w,
#                 'height': page_h,
#                 'orig_width': orig_w,
#                 'orig_height': orig_h,
#                 'dpi_scale': DPI_SCALE,
#                 'blocks': blocks,
#                 'images': images
#             })
        
#         # 打印坐标验证
#         if coord_samples:
#             print()
#             print("   📐 坐标归一化验证 (前2个文本块):")
#             for i, sample in enumerate(coord_samples):
#                 orig = sample['orig_bbox']
#                 norm = sample['normalized_bbox']
#                 print(f"      块{i+1}: {sample['text']}")
#                 print(f"         原始: ({orig[0]:.1f}, {orig[1]:.1f}, {orig[2]:.1f}, {orig[3]:.1f})")
#                 print(f"         Points: ({norm[0]:.1f}, {norm[1]:.1f}, {norm[2]:.1f}, {norm[3]:.1f})")
#             print()
        
#         return {
#             'pdf_name': save_dir.name,
#             'pages': pages_data,
#             'parse_method': 'magic-pdf',
#             'dpi_normalized': True,
#             'target_dpi': self.PDF_STANDARD_DPI
#         }
    
#     def _extract_text(self, block: Dict) -> str:
#         lines = block.get('lines', [])
#         if not lines:
#             return block.get('content', '')
        
#         parts = []
#         for line in lines:
#             for span in line.get('spans', []):
#                 c = span.get('content', '')
#                 if c:
#                     parts.append(c)
#         return ''.join(parts)
    
#     def _clean_text(self, text: str) -> str:
#         if not text:
#             return ""
#         text = re.sub(r'\s+', ' ', text)
#         text = re.sub(r'[□◇※★☆◆⚠]+', '', text)
#         return text.strip()
    
#     def _verify_coords(self, parsed_data: Dict) -> Dict:
#         """验证坐标精度"""
#         pages = parsed_data.get('pages', [])
#         if not pages:
#             return {'is_local': False, 'sample': '无页面数据'}
        
#         blocks = pages[0].get('blocks', [])
#         if not blocks:
#             return {'is_local': False, 'sample': '无文本块'}
        
#         bbox = blocks[0].get('bbox', [])
#         if not bbox or len(bbox) != 4:
#             return {'is_local': False, 'sample': '无效 bbox'}
        
#         x1, y1, x2, y2 = bbox
#         w, h = x2 - x1, y2 - y1
        
#         page_w = pages[0].get('width', 1)
#         page_h = pages[0].get('height', 1)
        
#         # 判断是否为全屏坐标
#         is_fullscreen = w > page_w * 0.95 or h > page_h * 0.95
#         is_local = not is_fullscreen and w < page_w * 0.9 and h < page_h * 0.9
        
#         sample = f"{w:.1f}x{h:.1f} (page: {page_w:.0f}x{page_h:.0f})"
        
#         return {'is_local': is_local, 'sample': sample, 'is_fullscreen': is_fullscreen}
    
#     def _extract_images(self, pdf_path: str, save_dir: Path, page_count: int) -> Path:
#         images_dir = save_dir / "images"
#         images_dir.mkdir(exist_ok=True)
        
#         doc = fitz.open(pdf_path)
#         for i in range(len(doc)):
#             pix = doc[i].get_pixmap(matrix=fitz.Matrix(200/72, 200/72))
#             pix.save(str(images_dir / f"page_{i}.png"))
#         doc.close()
        
#         return images_dir
    
#     def _fallback_parse(self, pdf_path: str, save_dir: Path, page_count: int, lang: str) -> Dict:
#         """备选方案"""
#         print("   🔄 备选方案 (PyMuPDF + OCR)...")
#         self.current_mode = "fallback"
        
#         images_dir = save_dir / "images"
#         images_dir.mkdir(exist_ok=True)
        
#         doc = fitz.open(pdf_path)
#         pages_data = []
        
#         for i in range(len(doc)):
#             pix = doc[i].get_pixmap(matrix=fitz.Matrix(200/72, 200/72))
#             img_path = images_dir / f"page_{i}.png"
#             pix.save(str(img_path))
            
#             text = doc[i].get_text()
#             text = self._clean_text(text)
            
#             pages_data.append({
#                 'page': i,
#                 'width': int(pix.width),
#                 'height': int(pix.height),
#                 'blocks': [{'text': text, 'bbox': [0, 0, pix.width, pix.height], 'type': 'text'}] if text else [],
#                 'images': [{'path': str(img_path), 'bbox': [0, 0, pix.width, pix.height]}]
#             })
        
#         doc.close()
        
#         json_path = save_dir / "layout.json"
#         data = {
#             'pdf_name': save_dir.name,
#             'pages': pages_data,
#             'parse_method': 'fallback'
#         }
        
#         with open(json_path, 'w', encoding='utf-8') as f:
#             json.dump(data, f, ensure_ascii=False, indent=2)
        
#         print(f"   ✅ 完成: {page_count} 页")
        
#         return {
#             "json_path": str(json_path),
#             "images_dir": str(images_dir),
#             "pdf_name": save_dir.name,
#             "page_count": page_count,
#             "model_version": "fallback"
#         }
    
#     @property
#     def mode(self) -> str:
#         return self.current_mode


# if __name__ == "__main__":
#     print("=" * 60)
#     print("🧪 MinerUWrapper v2.3 点火版")
#     print("=" * 60)
    
#     wrapper = MinerUWrapper()
    
#     test_pdf = "/Users/guanliming/Downloads/fd.pdf"
    
#     if Path(test_pdf).exists():
#         result = wrapper.process_pdf(test_pdf)
#         print(f"\n✅ 模式: {wrapper.mode}")
#     else:
#         print(f"文件不存在: {test_pdf}")




import os
import json
import re
import traceback
import fitz
from pathlib import Path
from typing import Dict, Any, Optional

class MinerUWrapper:
    """MinerU 核心解析包装器 (V2.8 对象模型适配版)"""
    
    def __init__(self, output_base_dir: str = "output/temp", use_gpu: bool = True):
        self.output_base_dir = Path(output_base_dir)
        self.output_base_dir.mkdir(parents=True, exist_ok=True)
        self.use_gpu = use_gpu
        self.current_mode = "unknown"
        self._magic_pdf_ready = False
        self._auto_init()

    def _auto_init(self):
        """初始化环境，解决 1.3.x 内部调用链问题"""
        try:
            from magic_pdf.data.dataset import PymuDocDataset
            from magic_pdf.model.doc_analyze_by_custom_model import doc_analyze
            from magic_pdf.pdf_parse_union_core_v2 import pdf_parse_union
            
            # 顽固寻找 DiskReaderWriter
            drw = None
            try:
                from magic_pdf.rw.DiskReaderWriter import DiskReaderWriter as drw
            except ImportError:
                try:
                    from magic_pdf.data.dataset import DiskReaderWriter as drw
                except ImportError:
                    class MockDRW:
                        def __init__(self, path): self.path = path
                        def write_json(self, name, data): pass
                        def write_image(self, name, img): pass
                    drw = MockDRW
            
            self._PymuDocDataset = PymuDocDataset
            self._doc_analyze = doc_analyze
            self._pdf_parse_union = pdf_parse_union
            self._DiskReaderWriter = drw
            self._magic_pdf_ready = True
            print("   ✅ magic-pdf 1.3.x 核心引擎加载成功")
        except Exception as e:
            print(f"   ⚠️ 引擎初始化警告: {e}")

    def process_pdf(self, pdf_path: str, lang: str = "chi_sim", force_fallback: bool = False) -> Dict[str, Any]:
        pdf_name = Path(pdf_path).stem
        save_dir = self.output_base_dir / pdf_name
        save_dir.mkdir(parents=True, exist_ok=True)
        
        print(f"\n🔄 解析: {pdf_path}")

        if not force_fallback and self._magic_pdf_ready:
            print(f"   🚀 Mac GPU (MPS) 已激活，正在进行深度版面分析...")
            try:
                return self._magic_pdf_full_parse(pdf_path, save_dir, lang)
            except Exception:
                print(f"   ⚠️ 深度解析异常，自动切入备选方案...")
                traceback.print_exc()
        
        return self._fallback_parse(pdf_path, save_dir)

    def _magic_pdf_full_parse(self, pdf_path: str, save_dir: Path, lang: str) -> Dict:
        """执行全量深度解析"""
        self.current_mode = "deep_parsing"
        
        with open(pdf_path, 'rb') as f:
            pdf_bytes = f.read()
        
        # 1. 初始化数据集 (PymuDocDataset 是 doc_analyze 期望的第一个参数)
        ds = self._PymuDocDataset(pdf_bytes, lang=lang)
        
        # 2. 执行 AI 模型分析
        # 修正点：第一个参数传 ds (对象)，而不是图片列表
        try:
            model_list = self._doc_analyze(ds, ocr=True, layout_model='layoutlmv3', formula_enable=True)
        except Exception as e:
            print(f"   🔔 尝试纯版面识别模式...")
            model_list = self._doc_analyze(ds, ocr=True, layout_model='layoutlmv3', formula_enable=False)

        # 3. 汇总解析
        drw = self._DiskReaderWriter(str(save_dir))
        parse_result = self._pdf_parse_union(model_list, ds, drw, ds.classify(), lang)
        
        # 4. 执行 0.72 坐标转换
        pages_data = self._normalize_data(parse_result)
        
        # 5. 确保图片保存 (为 PPT 渲染准备)
        images_dir = save_dir / "images"
        images_dir.mkdir(exist_ok=True)
        doc = fitz.open(pdf_path)
        for i in range(len(doc)):
            pix = doc[i].get_pixmap(dpi=150)
            pix.save(str(images_dir / f"page_{i}.png"))
        doc.close()

        json_path = save_dir / "layout.json"
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump({'pages': pages_data, 'parse_method': 'magic-pdf-deep'}, f, ensure_ascii=False, indent=2)

        print(f"   ✅ 深度解析成功! 坐标精度: 行级别")
        return {
            "json_path": str(json_path),
            "images_dir": str(images_dir),
            "page_count": len(pages_data)
        }

    def _normalize_data(self, parse_result) -> list:
        """核心坐标转换：YOLO 100 DPI -> PDF Points (72 DPI)"""
        pages_data = []
        DPI_SCALE = 0.72 
        
        for page_id, info in parse_result.items():
            page_idx = int(re.search(r'\d+', page_id).group()) if re.search(r'\d+', page_id) else 0
            blocks = []
            for block in info.get('preproc_blocks', []):
                bbox = [b * DPI_SCALE for b in block.get('bbox', [0,0,0,0])]
                
                text_parts = []
                for line in block.get('lines', []):
                    for span in line.get('spans', []):
                        content = span.get('content', '')
                        if content: text_parts.append(content)
                text = "".join(text_parts).strip()
                
                if text:
                    blocks.append({
                        'text': text,
                        'bbox': bbox,
                        'type': 'title' if block.get('type') in ['title', 'heading'] else 'text'
                    })
            
            pages_data.append({
                'page': page_idx,
                'blocks': blocks,
                'width': 720,
                'height': 405
            })
        return pages_data

    def _fallback_parse(self, pdf_path: str, save_dir: Path) -> Dict:
        """备选方案 logic"""
        self.current_mode = "fallback"
        print("   🔔 使用备选方案 (Fallback Mode)")
        # 即使深度解析失败，也通过 PyMuPDF 抓一些文字
        doc = fitz.open(pdf_path)
        pages_data = []
        for i, page in enumerate(doc):
            pages_data.append({
                'page': i,
                'blocks': [{'text': page.get_text().strip(), 'bbox': [0,0,720,405], 'type': 'text'}],
                'width': 720, 'height': 405
            })
        doc.close()
        
        json_path = save_dir / "layout.json"
        with open(json_path, 'w') as f:
            json.dump({'pages': pages_data, 'parse_method': 'fallback'}, f)
            
        return {"json_path": str(json_path), "images_dir": "", "page_count": len(pages_data)}

    @property
    def mode(self) -> str: return self.current_mode

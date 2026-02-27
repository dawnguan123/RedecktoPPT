"""
RedecktoPPT - MinerU Handler
处理 MinerU API 响应，提取文字、坐标、字体大小等信息
并实现 PDF → PPT 坐标系转换
"""

import json
import os
import zipfile
import tempfile
import shutil
from pathlib import Path
from dataclasses import dataclass, field
from typing import List, Dict, Any, Optional, Tuple
from dataclasses import asdict
import numpy as np

# PPT 尺寸常量（EMU）
# 1 inch = 914400 EMU
# 1 point = 12700 EMU
PPT_WIDTH_16x9 = 10 * 914400    # 10 inches × 914400
PPT_HEIGHT_16x9 = 5.625 * 914400  # 5.625 inches × 914400
PPT_WIDTH_4x3 = 10 * 914400
PPT_HEIGHT_4x3 = 7.5 * 914400

# 标准 PPT 尺寸（points）
PPT_WIDTH_PT = 10 * 72       # 720 points
PPT_HEIGHT_16x9_PT = 5.625 * 72  # 405 points
PPT_HEIGHT_4x3_PT = 7.5 * 72    # 540 points


@dataclass
class TextElement:
    """文本元素"""
    text: str
    x: float          # PDF 坐标 x
    y: float          # PDF 坐标 y
    width: float       # 宽度
    height: float     # 高度
    font_size: float = 0  # 估计字体大小
    font_name: str = ""   # 字体名称
    type: str = "text"    # title, text, caption
    confidence: float = 1.0


@dataclass
class ImageElement:
    """图片元素"""
    name: str
    x: float
    y: float
    width: float
    height: float
    path: str = ""  # 本地路径


@dataclass
class TableElement:
    """表格元素"""
    x: float
    y: float
    width: float
    height: float
    html: str = ""
    rows: List[List[str]] = field(default_factory=list)


@dataclass
class PageLayout:
    """页面布局"""
    page_index: int
    pdf_width: float      # PDF 原始宽度（points）
    pdf_height: float     # PDF 原始高度（points）
    texts: List[TextElement] = field(default_factory=list)
    images: List[ImageElement] = field(default_factory=list)
    tables: List[TableElement] = field(default_factory=list)


class CoordinateConverter:
    """
    坐标转换器：PDF → PPT
    
    支持：
    - 16:9 和 4:3 页面
    - 保持比例缩放
    - 自定义边距
    """
    
    def __init__(
        self,
        target_aspect: str = "16:9",
        margin_pt: float = 20  # 边距 points
    ):
        """
        初始化转换器
        
        Args:
            target_aspect: 目标宽高比 "16:9" 或 "4:3"
            margin_pt: 边距（points）
        """
        self.target_aspect = target_aspect
        self.margin_pt = margin_pt
        
        if target_aspect == "16:9":
            self.target_width = PPT_WIDTH_PT - 2 * margin_pt
            self.target_height = PPT_HEIGHT_16x9_PT - 2 * margin_pt
        else:
            self.target_width = PPT_WIDTH_PT - 2 * margin_pt
            self.target_height = PPT_HEIGHT_4x3_PT - 2 * margin_pt
    
    def compute_scale(self, pdf_width: float, pdf_height: float) -> float:
        """
        计算缩放比例
        
        Args:
            pdf_width: PDF 原始宽度
            pdf_height: PDF 原始高度
            
        Returns:
            缩放因子
        """
        # 计算 PDF 实际内容区域（减去边距）
        pdf_content_w = pdf_width - 2 * self.margin_pt
        pdf_content_h = pdf_height - 2 * self.margin_pt
        
        # 计算缩放比例（取较小的，确保内容完整）
        scale_x = self.target_width / pdf_content_w
        scale_y = self.target_height / pdf_content_h
        
        return min(scale_x, scale_y)
    
    def pdf_to_ppt(
        self,
        x: float,
        y: float,
        width: float,
        height: float,
        pdf_width: float,
        pdf_height: float
    ) -> Tuple[int, int, int, int]:
        """
        PDF 坐标转换为 PPT EMU 坐标
        
        Args:
            x, y, width, height: PDF 坐标
            pdf_width, pdf_height: PDF 页面尺寸
            
        Returns:
            (left, top, width, height) in EMU
        """
        # 计算缩放
        scale = self.compute_scale(pdf_width, pdf_height)
        
        # 居中偏移
        scaled_w = pdf_width * scale
        scaled_h = pdf_height * scale
        offset_x = (self.target_width + 2 * self.margin_pt - scaled_w) / 2
        offset_y = (self.target_height + 2 * self.margin_pt - scaled_h) / 2
        
        # 转换坐标
        # PDF 坐标系：左下角为原点，y 向上
        # PPT 坐标系：左上角为原点，y 向下
        # 需要翻转 Y 轴
        
        left_pt = (x - self.margin_pt) * scale + offset_x
        # PDF y 是从底部开始，PPT 是从顶部开始
        top_pt = (pdf_height - y - height - self.margin_pt) * scale + offset_y
        
        width_pt = width * scale
        height_pt = height * scale
        
        # 转换为 EMU
        left_emu = int(left_pt * 12700)
        top_emu = int(top_pt * 12700)
        width_emu = int(width_pt * 12700)
        height_emu = int(height_pt * 12700)
        
        return left_emu, top_emu, width_emu, height_emu
    
    def pdf_to_ppt_simple(
        self,
        bbox: List[float],
        pdf_width: float,
        pdf_height: float
    ) -> Tuple[int, int, int, int]:
        """
        简化版：从 bbox [x1, y1, x2, y2] 转换
        """
        x1, y1, x2, y2 = bbox
        x = x1
        y = y1
        width = x2 - x1
        height = y2 - y1
        
        return self.pdf_to_ppt(x, y, width, height, pdf_width, pdf_height)


class MinerUHandler:
    """
    MinerU 响应处理器
    
    从 MinerU JSON/zip 中提取：
    - 文字内容
    - 字体大小估计
    - 像素坐标
    """
    
    def __init__(self, output_dir: str = None):
        """
        初始化
        
        Args:
            output_dir: 输出目录（用于解压 zip）
        """
        self.output_dir = output_dir or tempfile.mkdtemp()
        self.json_data: Dict[str, Any] = {}
        self.page_layouts: List[PageLayout] = []
        
        # 提取的图片目录
        self.images_dir = os.path.join(self.output_dir, "images")
        os.makedirs(self.images_dir, exist_ok=True)
    
    def load_from_json(self, json_path: str) -> bool:
        """
        从 JSON 文件加载
        
        Args:
            json_path: JSON 文件路径
            
        Returns:
            是否成功
        """
        try:
            with open(json_path, 'r', encoding='utf-8') as f:
                self.json_data = json.load(f)
            self._parse_layouts()
            return True
        except Exception as e:
            print(f"❌ Failed to load JSON: {e}")
            return False
    
    def load_from_zip(self, zip_path: str) -> bool:
        """
        从 zip 文件加载（MinerU full_zip_url）
        
        Args:
            zip_path: zip 文件路径
            
        Returns:
            是否成功
        """
        try:
            # 解压 zip
            with zipfile.ZipFile(zip_path, 'r') as zf:
                zf.extractall(self.output_dir)
            
            # 查找 JSON 文件
            json_files = list(Path(self.output_dir).rglob("*.json"))
            if not json_files:
                print("❌ No JSON file found in zip")
                return False
            
            # 加载第一个 JSON
            json_path = json_files[0]
            return self.load_from_json(str(json_path))
            
        except Exception as e:
            print(f"❌ Failed to load zip: {e}")
            return False
    
    def load_from_content(self, content: Dict[str, Any]) -> bool:
        """
        从 API content 响应加载
        
        Args:
            content: API 返回的 content 字典
        """
        self.json_data = content
        self._parse_layouts()
        return True
    
    def _parse_layouts(self):
        """解析页面布局"""
        self.page_layouts = []
        
        pages = self.json_data.get('pages', [])
        
        for idx, page in enumerate(pages):
            layout = PageLayout(
                page_index=idx,
                pdf_width=page.get('width', 0),
                pdf_height=page.get('height', 0)
            )
            
            # 解析文本块
            for block in page.get('blocks', []):
                block_type = block.get('type', 'text')
                
                if block_type in ['title', 'text', 'caption', 'figure_caption']:
                    # 提取文本
                    text = block.get('text', '')
                    if not text:
                        continue
                    
                    # 提取坐标
                    bbox = block.get('bbox', [0, 0, 0, 0])
                    x, y, x2, y2 = bbox
                    width = x2 - x
                    height = y2 - y
                    
                    # 估计字体大小
                    # 使用文本框高度作为参考
                    # 假设行高 ≈ 1.2 × 字体大小
                    font_size = height / 1.2 if height > 0 else 12
                    
                    # 检测是否为标题（较大的字体）
                    if block_type == 'title' or font_size > 18:
                        text_type = 'title'
                    elif block_type == 'caption':
                        text_type = 'caption'
                    else:
                        text_type = 'text'
                    
                    layout.texts.append(TextElement(
                        text=text,
                        x=x,
                        y=y,
                        width=width,
                        height=height,
                        font_size=font_size,
                        type=text_type,
                        confidence=block.get('score', 1.0)
                    ))
                
                elif block_type == 'figure':
                    # 图片块
                    bbox = block.get('bbox', [0, 0, 0, 0])
                    x, y, x2, y2 = bbox
                    
                    layout.images.append(ImageElement(
                        name=block.get('name', f'image_{len(layout.images)}'),
                        x=x,
                        y=y,
                        width=x2 - x,
                        height=y2 - y
                    ))
                
                elif block_type == 'table':
                    # 表格块
                    bbox = block.get('bbox', [0, 0, 0, 0])
                    html = block.get('html', '')
                    
                    layout.tables.append(TableElement(
                        x=bbox[0],
                        y=bbox[1],
                        width=bbox[2] - bbox[0],
                        height=bbox[3] - bbox[1],
                        html=html
                    ))
            
            self.page_layouts.append(layout)
    
    def get_page_layout(self, page_idx: int) -> Optional[PageLayout]:
        """获取指定页面布局"""
        if 0 <= page_idx < len(self.page_layouts):
            return self.page_layouts[page_idx]
        return None
    
    def get_all_texts(self, page_idx: int) -> List[TextElement]:
        """获取页面所有文本"""
        layout = self.get_page_layout(page_idx)
        return layout.texts if layout else []
    
    def get_all_images(self, page_idx: int) -> List[ImageElement]:
        """获取页面所有图片"""
        layout = self.get_page_layout(page_idx)
        return layout.images if layout else []
    
    def to_ppt_format(
        self,
        page_idx: int,
        target_aspect: str = "16:9",
        margin_pt: float = 20
    ) -> List[Dict[str, Any]]:
        """
        转换为 PPT 格式
        
        Args:
            page_idx: 页码
            target_aspect: 目标宽高比
            margin_pt: 边距
            
        Returns:
            PPT 元素列表
        """
        layout = self.get_page_layout(page_idx)
        if not layout:
            return []
        
        # 创建转换器
        converter = CoordinateConverter(target_aspect, margin_pt)
        
        ppt_elements = []
        
        # 转换文本
        for text in layout.texts:
            left, top, width, height = converter.pdf_to_ppt(
                text.x, text.y, text.width, text.height,
                layout.pdf_width, layout.pdf_height
            )
            
            ppt_elements.append({
                'type': 'text',
                'text': text.text,
                'left': left,
                'top': top,
                'width': width,
                'height': height,
                'font_size': int(text.font_size * converter.compute_scale(
                    layout.pdf_width, layout.pdf_height
                )),
                'text_type': text.type
            })
        
        # 转换图片
        for img in layout.images:
            left, top, width, height = converter.pdf_to_ppt(
                img.x, img.y, img.width, img.height,
                layout.pdf_width, layout.pdf_height
            )
            
            ppt_elements.append({
                'type': 'image',
                'name': img.name,
                'left': left,
                'top': top,
                'width': width,
                'height': height,
                'path': img.path
            })
        
        return ppt_elements
    
    def get_summary(self) -> Dict[str, Any]:
        """获取解析摘要"""
        total_texts = sum(len(p.texts) for p in self.page_layouts)
        total_images = sum(len(p.images) for p in self.page_layouts)
        total_tables = sum(len(p.tables) for p in self.page_layouts)
        
        return {
            'total_pages': len(self.page_layouts),
            'total_texts': total_texts,
            'total_images': total_images,
            'total_tables': total_tables,
            'pages': [
                {
                    'index': p.page_index,
                    'width': p.pdf_width,
                    'height': p.pdf_height,
                    'texts': len(p.texts),
                    'images': len(p.images),
                    'tables': len(p.tables)
                }
                for p in self.page_layouts
            ]
        }


# ========== 便捷函数 ==========

def load_mineru_result(source: str) -> MinerUHandler:
    """
    加载 MinerU 结果
    
    Args:
        source: JSON 文件路径、zip 文件路径或目录
        
    Returns:
        MinerUHandler 实例
    """
    handler = MinerUHandler()
    
    path = Path(source)
    
    if path.is_file():
        if path.suffix == '.json':
            handler.load_from_json(str(path))
        elif path.suffix == '.zip':
            handler.load_from_zip(str(path))
    elif path.is_dir():
        # 查找 JSON
        json_files = list(path.rglob("*.json"))
        if json_files:
            handler.load_from_json(str(json_files[0]))
    
    return handler


# ========== 示例用法 ==========

if __name__ == "__main__":
    # 测试
    print("MinerUHandler module ready")
    print()
    
    # 坐标转换测试
    converter = CoordinateConverter("16:9", margin_pt=20)
    scale = converter.compute_scale(600, 400)  # PDF 600×400
    print(f"Scale for 600x400 PDF: {scale:.3f}")
    
    # 转换测试
    left, top, w, h = converter.pdf_to_ppt(
        x=100, y=100, width=200, height=30,
        pdf_width=600, pdf_height=400
    )
    print(f"PDF(100,100,200,30) → PPT({left},{top},{w},{h}) EMU")
    
    print()
    print("Usage:")
    print("  handler = load_mineru_result('mineru_output.json')")
    print("  ppt_elements = handler.to_ppt_format(0)")

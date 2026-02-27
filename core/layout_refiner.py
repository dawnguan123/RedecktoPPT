#!/usr/bin/env python3
"""
RedecktoPPT - Layout Refiner
文本聚类与排版优化模块

功能：
1. 预排序：按 Y/X 坐标排序
2. 聚类合并：将相近的行合并为段落
3. 列表识别：识别. 坐标转换 bullet points
4：输出 PPT 格式

NotebookLM 生成的演示文稿目前最大的痛点是其导出的 PDF 本质上是静态图片，
无法直接编辑文字。虽然 MinerU 能识别出文字，但其原始输出往往是细碎的行。
本模块通过空间聚类算法，将这些碎片重新组合为符合人类逻辑的"段落"和"标题"，
这是实现"全可编辑"的关键一步。
"""

import json
import re
from dataclasses import dataclass, field
from typing import List, Dict, Any, Optional, Tuple


@dataclass
class TextBlock:
    """逻辑文本块数据模型"""
    text: str
    x: float
    y: float
    width: float
    height: float
    font_size: float
    block_type: str  # title, paragraph, list


@dataclass
class TextElement:
    """文本元素"""
    text: str
    x: float           # 左上角 x
    y: float           # 左上角 y
    width: float
    height: float
    original_type: str = ""  # 原始类型 from MinerU
    alignment: str = "left"   # left/center/right


@dataclass
class Paragraph:
    """合并后的段落"""
    texts: List[str] = field(default_factory=list)
    x: float = 0
    y: float = 0
    width: float = 0
    height: float = 0
    alignment: str = "left"
    
    @property
    def content(self) -> str:
        return "\n".join(self.texts)
    
    @property
    def bbox(self) -> List[float]:
        return [self.x, self.y, self.x + self.width, self.y + self.height]


@dataclass
class RefinedElement:
    """精炼后的元素"""
    type: str           # title/body/list
    content: str        # 文本内容
    left: float         # PPT 坐标
    top: float
    width: float
    height: float
    alignment: str = "left"


class LayoutRefiner:
    """
    Redeckto-Cluster 算法实现
    
    将 MinerU 的细碎行合并为逻辑段落
    """
    
    # 列表标识符
    BULLET_PATTERNS = [
        r'^[\-\•\*]\s+',      # - • * 
        r'^\d+[\.\)]\s+',     # 1. 2) 
        r'^[a-zA-Z][\.\)]\s+', # a. b)
        r'^\([\d]+\)\s+',    # (1) (2)
    ]
    
    # 标题关键词（辅助判断）
    TITLE_KEYWORDS = [
        '标题', 'title', 'headline', 'chapter',
        '第', '章', '节', 'section'
    ]
    
    def __init__(self, line_gap_threshold: float = 1.5):
        """
        初始化
        
        Args:
            line_gap_threshold: 纵向间距阈值
                               当前行高与下一行高度的倍数
        """
        self.line_gap_threshold = line_gap_threshold
    
    def refine(self, raw_json_path: str) -> List[TextBlock]:
        """
        执行聚类重构逻辑
        
        Args:
            raw_json_path: MinerU 原始 JSON 文件路径
            
        Returns:
            精炼后的文本块列表
        """
        with open(raw_json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        # 1. 提取所有包含文字的元素
        raw_elements = self._extract_elements(data)
        
        # 2. 空间排序：先纵向(Y)，再横向(X)
        sorted_elements = sorted(raw_elements, key=lambda e: (e['y'], e['x']))
        
        # 3. 聚类合并算法
        refined_blocks = self._cluster_elements(sorted_elements)
        
        return refined_blocks
    
    def refine_from_dict(self, data: Dict) -> List[TextBlock]:
        """
        从字典数据执行聚类
        
        Args:
            data: MinerU JSON 数据字典
            
        Returns:
            精炼后的文本块列表
        """
        # 提取元素
        raw_elements = self._extract_elements(data)
        
        # 空间排序
        sorted_elements = sorted(raw_elements, key=lambda e: (e['y'], e['x']))
        
        # 聚类合并
        return self._cluster_elements(sorted_elements)
    
    def _extract_elements(self, data: Dict) -> List[Dict]:
        """
        从 MinerU 数据中提取文字元素
        
        需根据 MinerU 实际输出的 JSON 层级进行解析
        提取 text, x, y, w, h, font_size 等核心字段
        
        Args:
            data: MinerU JSON 数据
            
        Returns:
            元素列表
        """
        elements = []
        
        # 尝试多种 MinerU 输出格式
        
        # 格式1: data['pages'][page_idx]['blocks']
        pages = data.get('pages', [])
        for page in pages:
            blocks = page.get('blocks', [])
            for block in blocks:
                # 提取文本
                text = block.get('text', '').strip()
                if not text:
                    continue
                
                # 提取坐标
                bbox = block.get('bbox', [0, 0, 0, 0])
                x1, y1, x2, y2 = bbox
                
                # 估计字体大小
                height = y2 - y1
                font_size = height / 1.2 if height > 0 else 12
                
                elements.append({
                    'text': text,
                    'x': x1,
                    'y': y1,
                    'width': x2 - x1,
                    'height': height,
                    'font_size': font_size,
                    'type': block.get('type', 'text')
                })
        
        # 格式2: data['elements'] (扁平结构)
        if not elements and 'elements' in data:
            for elem in data['elements']:
                text = elem.get('text', '').strip()
                if not text:
                    continue
                
                bbox = elem.get('bbox', elem.get('box', [0, 0, 0, 0]))
                if len(bbox) == 4:
                    x1, y1, x2, y2 = bbox
                    elements.append({
                        'text': text,
                        'x': x1,
                        'y': y1,
                        'width': x2 - x1,
                        'height': y2 - y1,
                        'font_size': elem.get('font_size', 12),
                        'type': elem.get('type', 'text')
                    })
        
        return elements
    
    def _cluster_elements(self, elements: List[Dict]) -> List[TextBlock]:
        """
        聚类合并算法
        
        将空间上相近且字体大小相似的行合并为同一段落
        
        Args:
            elements: 排序后的元素列表
            
        Returns:
            合并后的文本块列表
        """
        if not elements:
            return []
        
        clusters = []
        current_block = None
        
        for elem in elements:
            if current_block is None:
                current_block = self._init_new_block(elem)
                continue
            
            # 计算纵向间距
            dy = elem['y'] - (current_block.y + current_block.height)
            
            # 逻辑判定：间距相近且字体大小相近，则视为同一段落
            same_paragraph = (
                dy < (elem['height'] * self.line_gap_threshold) and 
                abs(elem['font_size'] - current_block.font_size) < 2
            )
            
            if same_paragraph:
                # 更新文本和边界框
                current_block.text += " " + elem['text']
                current_block.width = max(
                    current_block.width, 
                    elem['x'] + elem['width'] - current_block.x
                )
                current_block.height = (elem['y'] + elem['height']) - current_block.y
            else:
                # 新段落
                clusters.append(current_block)
                current_block = self._init_new_block(elem)
        
        # 保存最后一个块
        if current_block:
            clusters.append(current_block)
        
        return clusters
    
    def _init_new_block(self, elem: Dict) -> TextBlock:
        """
        初始化新的文本块
        
        初步判定类型：根据高度或 MinerU 标签
        
        Args:
            elem: 原始元素
            
        Returns:
            TextBlock 对象
        """
        # 判断类型
        font_size = elem.get('font_size', 12)
        
        # 优先使用 MinerU 原生类型
        orig_type = elem.get('type', '').lower()
        if 'title' in orig_type:
            b_type = 'title'
        elif 'caption' in orig_type:
            b_type = 'list'
        elif font_size > 18:
            b_type = 'title'
        else:
            b_type = 'paragraph'
        
        return TextBlock(
            text=elem['text'],
            x=elem['x'],
            y=elem['y'],
            width=elem['width'],
            height=elem['height'],
            font_size=font_size,
            block_type=b_type
        )
    
    def _detect_bullet(self, text: str) -> bool:
        """
        检测是否为列表项
        
        Args:
            text: 文本内容
            
        Returns:
            是否为列表项
        """
        for pattern in self.BULLET_PATTERNS:
            if re.match(pattern, text):
                return True
        return False
    
    def to_ppt_format(
        self,
        refined_blocks: List[TextBlock],
        pdf_width: float,
        pdf_height: float,
        target_width: float = 720,   # PPT 10英寸 = 720pt
        target_height: float = 405,   # PPT 16:9 5.625英寸 = 405pt
        margin: float = 20
    ) -> List[Dict[str, Any]]:
        """
        转换为 PPT 坐标格式
        
        Args:
            refined_blocks: 精炼后的文本块
            pdf_width: PDF 宽度
            pdf_height: PDF 高度
            target_width: 目标宽度 (points)
            target_height: 目标高度 (points)
            margin: 边距 (points)
            
        Returns:
            PPT 格式的元素列表
        """
        # 计算缩放比例
        scale_x = (target_width - 2 * margin) / pdf_width
        scale_y = (target_height - 2 * margin) / pdf_height
        scale = min(scale_x, scale_y)  # 保持比例
        
        # 计算偏移（居中）
        scaled_w = pdf_width * scale
        scaled_h = pdf_height * scale
        offset_x = (target_width - scaled_w) / 2
        offset_y = (target_height - scaled_h) / 2
        
        result = []
        
        for block in refined_blocks:
            # 转换坐标
            # PDF Y 坐标从底部开始，PPT 从顶部开始，需要翻转
            left = (block.x - margin) * scale + offset_x
            top = (pdf_height - block.y - block.height - margin) * scale + offset_y
            width = block.width * scale
            height = block.height * scale
            
            # 检测列表
            block_type = block.block_type
            if self._detect_bullet(block.text):
                block_type = 'list'
            
            result.append({
                'type': block_type,
                'content': block.text,
                'left': int(left),
                'top': int(top),
                'width': int(width),
                'height': int(height),
                'font_size': int(block.font_size * scale)
            })
        
        return result


# ========== 便捷函数 ==========

def refine_layout(raw_json_path: str, **kwargs) -> List[Dict[str, Any]]:
    """
    便捷函数：一站式精炼+坐标转换
    
    Args:
        raw_json_path: MinerU JSON 文件路径
        **kwargs: 其他参数
        
    Returns:
        PPT 格式的元素列表
    """
    refiner = LayoutRefiner(**kwargs)
    blocks = refiner.refine(raw_json_path)
    
    # 尝试获取 PDF 尺寸
    with open(raw_json_path, 'r') as f:
        data = json.load(f)
    
    pdf_width = data.get('width', 600)
    pdf_height = data.get('height', 800)
    
    return refiner.to_ppt_format(blocks, pdf_width, pdf_height)


if __name__ == "__main__":
    # 测试
    print("LayoutRefiner module ready")
    print()
    
    # 模拟数据测试
    refiner = LayoutRefiner(line_gap_threshold=1.5)
    
    test_data = {
        'pages': [{
            'blocks': [
                {'text': '深度学习简介', 'bbox': [50, 50, 300, 90], 'type': 'title', 'font_size': 24},
                {'text': '第一章', 'bbox': [50, 100, 150, 130], 'type': 'text', 'font_size': 14},
                {'text': '机器学习是人工智能的一个分支', 'bbox': [50, 150, 400, 180], 'type': 'text', 'font_size': 12},
                {'text': '它使用多层神经网络进行学习', 'bbox': [50, 185, 380, 215], 'type': 'text', 'font_size': 12},
            ]
        }],
        'width': 600,
        'height': 800
    }
    
    blocks = refiner.refine_from_dict(test_data)
    
    print("Refined blocks:")
    for b in blocks:
        print(f"  [{b.block_type}] {b.text[:30]}...")
        print(f"      ({b.x}, {b.y}) {b.width}x{b.height}")

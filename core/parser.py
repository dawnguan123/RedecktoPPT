"""
RedecktoPPT - Core Parser Interface
PDF Parser 接口定义，用于对接 MinerU API 获取解析后的 JSON 数据
"""

from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from typing import List, Dict, Any, Optional, Tuple
from pathlib import Path
import numpy as np


@dataclass
class Block:
    """文本块基类"""
    text: str
    bbox: Tuple[float, float, float, float]  # (x1, y1, x2, y2)
    type: str  # 'title', 'text', 'figure', 'table', etc.
    score: float = 1.0


@dataclass
class ImageBlock:
    """图片块"""
    bbox: Tuple[float, float, float, float]
    url: str = ""
    name: str = ""
    pixels: Optional[np.ndarray] = None


@dataclass
class TableBlock:
    """表格块"""
    bbox: Tuple[float, float, float, float]
    html: str = ""
    rows: List[List[str]] = field(default_factory=list)


@dataclass
class PageData:
    """单页解析结果"""
    page_index: int
    width: float
    height: float
    blocks: List[Block] = field(default_factory=list)
    images: List[ImageBlock] = field(default_factory=list)
    tables: List[TableBlock] = field(default_factory=list)
    layout: Dict[str, Any] = field(default_factory=dict)


@dataclass
class PDFParseResult:
    """完整解析结果"""
    source_file: str
    total_pages: int
    pages: List[PageData] = field(default_factory=list)
    metadata: Dict[str, Any] = field(default_factory=dict)


class PDFParser(ABC):
    """
    PDF 解析器抽象基类
    
    所有解析器必须实现以下接口：
    1. parse() - 解析 PDF 文件
    2. parse_page() - 解析单页
    3. get_page_image() - 获取页面图片
    """
    
    def __init__(self, config: Optional[Dict[str, Any]] = None):
        """
        初始化解析器
        
        Args:
            config: 配置字典
        """
        self.config = config or {}
        self._result: Optional[PDFParseResult] = None
    
    @abstractmethod
    def parse(self, pdf_path: str) -> PDFParseResult:
        """
        解析 PDF 文件
        
        Args:
            pdf_path: PDF 文件路径
            
        Returns:
            PDFParseResult: 解析结果
        """
        pass
    
    @abstractmethod
    def parse_page(self, pdf_path: str, page_idx: int) -> PageData:
        """
        解析单页
        
        Args:
            pdf_path: PDF 文件路径
            page_idx: 页码（从 0 开始）
            
        Returns:
            PageData: 单页解析结果
        """
        pass
    
    @abstractmethod
    def get_page_image(self, pdf_path: str, page_idx: int, dpi: int = 150) -> np.ndarray:
        """
        获取页面图片
        
        Args:
            pdf_path: PDF 文件路径
            page_idx: 页码
            dpi: 分辨率
            
        Returns:
            np.ndarray: RGB 图片
        """
        pass
    
    @property
    def result(self) -> Optional[PDFParseResult]:
        """获取解析结果"""
        return self._result
    
    def __repr__(self) -> str:
        return f"{self.__class__.__name__}()"


class MinerUParser(PDFParser):
    """
    MinerU 解析器
    
    使用 MinerU API 进行深度版面分析
    """
    
    def __init__(self, api_key: str = None, config: Optional[Dict[str, Any]] = None):
        super().__init__(config)
        self.api_key = api_key or self.config.get('api_key')
        self.api_url = self.config.get('api_url', 'https://api.mineru.cn/v1')
    
    def parse(self, pdf_path: str) -> PDFParseResult:
        """调用 MinerU API 解析 PDF"""
        # TODO: 实现 API 调用
        raise NotImplementedError("MinerU API integration pending")
    
    def parse_page(self, pdf_path: str, page_idx: int) -> PageData:
        """解析单页"""
        raise NotImplementedError()
    
    def get_page_image(self, pdf_path: str, page_idx: int, dpi: int = 150) -> np.ndarray:
        """获取页面图片"""
        raise NotImplementedError()


class LocalParser(PDFParser):
    """
    本地解析器（备用方案）
    
    使用 PyMuPDF + Tesseract 进行本地解析
    """
    
    def __init__(self, config: Optional[Dict[str, Any]] = None):
        super().__init__(config)
        self.dpi = self.config.get('dpi', 150)
    
    def parse(self, pdf_path: str) -> PDFParseResult:
        """本地解析 PDF"""
        # TODO: 实现本地解析
        raise NotImplementedError()
    
    def parse_page(self, pdf_path: str, page_idx: int) -> PageData:
        """解析单页"""
        raise NotImplementedError()
    
    def get_page_image(self, pdf_path: str, page_idx: int, dpi: int = 150) -> np.ndarray:
        """获取页面图片（使用 PyMuPDF）"""
        import fitz
        doc = fitz.open(pdf_path)
        page = doc[page_idx]
        zoom = dpi / 72.0
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
        img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(
            pix.height, pix.width, pix.n
        )
        doc.close()
        if pix.n == 4:
            img = img[:, :, :3]
        return img


def create_parser(parser_type: str = 'local', **kwargs) -> PDFParser:
    """
    工厂函数：创建解析器
    
    Args:
        parser_type: 'mineru' | 'local'
        **kwargs: 解析器配置
        
    Returns:
        PDFParser 实例
    """
    parsers = {
        'mineru': MinerUParser,
        'local': LocalParser,
    }
    
    if parser_type not in parsers:
        raise ValueError(f"Unknown parser: {parser_type}. Available: {list(parsers.keys())}")
    
    return parsers[parser_type](kwargs)


if __name__ == "__main__":
    # 测试
    parser = create_parser('local', dpi=150)
    print(f"Created parser: {parser}")

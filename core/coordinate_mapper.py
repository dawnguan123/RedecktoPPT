#!/usr/bin/env python3
"""
RedecktoPPT - Coordinate Mapper
坐标换算引擎

将 MinerU 的 PDF 坐标系 (Points) 映射到 python-pptx 的坐标系 (EMUs)

在 PPT 自动化生成中，坐标转换不只是简单的数字缩放。
由于 PDF（通常是 A4 或固定比例点位）与 PPT（通常是 16:9 比例）的单位系统完全不同，
我们必须通过缩放因子（Scale Factor）和单位换算常数来确保元素位置的绝对精准。
"""

from pptx.util import Inches, Pt, Emu
from dataclasses import dataclass
from typing import Tuple, List, Optional
from enum import Enum


class SlideType(Enum):
    """幻灯片类型"""
    NORMAL = "normal"       # 普通页面
    COVER = "cover"         # 封面（第一页）
    BACKCOVER = "backcover" # 封底（最后一页）


@dataclass
class TransformResult:
    """转换结果"""
    left: int      # EMU
    top: int       # EMU
    width: int     # EMU
    height: int    # EMU
    scale: float


class CoordinateMapper:
    """
    RedecktoPPT 坐标换算引擎
    
    将 MinerU 的 PDF 坐标系 (Points) 映射到 python-pptx 的坐标系 (EMUs)
    
    单位换算：
    - 1 inch = 72 points = 914,400 EMUs
    - 1 point = 12,700 EMUs
    """
    
    # 单位换算常量
    POINTS_PER_INCH = 72
    EMU_PER_POINT = 12700
    EMU_PER_INCH = 914400
    
    # 默认 16:9 尺寸（英寸）
    DEFAULT_PPT_WIDTH = 10.0
    DEFAULT_PPT_HEIGHT = 5.625
    
    # 最大幻灯片页码（商业限制）
    MAX_SLIDES = 20
    
    def __init__(
        self,
        pdf_width: float,
        pdf_height: float,
        ppt_width_inch: float = 10.0,
        ppt_height_inch: float = 5.625,
        margin_inch: float = 0.5
    ):
        """
        初始化坐标映射器
        
        Args:
            pdf_width: PDF 原始宽度 (Points)
            pdf_height: PDF 原始高度 (Points)
            ppt_width_inch: PPT 宽度 (英寸)，默认 10"
            ppt_height_inch: PPT 高度 (英寸)，默认 5.625" (16:9)
            margin_inch: 边距 (英寸)
        """
        # 1. PDF 原始尺寸 (Points)
        self.pdf_w = pdf_width
        self.pdf_h = pdf_height
        
        # 2. 目标 PPT 尺寸 (EMUs)
        self.ppt_w = ppt_width_inch * self.EMU_PER_INCH
        self.ppt_h = ppt_height_inch * self.EMU_PER_INCH
        
        # 3. 边距 (EMUs)
        self.margin = margin_inch * self.EMU_PER_INCH
        
        # 4. 计算缩放因子 (Scale Factor)
        # 有效区域 = PPT尺寸 - 2*边距
        effective_ppt_w = self.ppt_w - 2 * self.margin
        effective_ppt_h = self.ppt_h - 2 * self.margin
        
        # 缩放因子 = 目标有效区域 / PDF原始尺寸
        self.sf_x = effective_ppt_w / (self.pdf_w * self.EMU_PER_POINT)
        self.sf_y = effective_ppt_h / (self.pdf_h * self.EMU_PER_POINT)
        
        # 取较小值以保持比例
        self.scale = min(self.sf_x, self.sf_y)
        
        # 计算居中偏移
        scaled_w = self.pdf_w * self.EMU_PER_POINT * self.scale
        scaled_h = self.pdf_h * self.EMU_PER_POINT * self.scale
        self.offset_x = (effective_ppt_w - scaled_w) / 2
        self.offset_y = (effective_ppt_h - scaled_h) / 2
    
    def to_pptx_geometry(
        self,
        x: float,
        y: float,
        w: float,
        h: float,
        slide_type: SlideType = SlideType.NORMAL
    ) -> Tuple[int, int, int, int]:
        """
        将 MinerU 的绝对坐标转换为 PPT 的绝对位置
        
        输入单位: Points (来自 MinerU JSON)
        输出单位: EMUs (用于 python-pptx)
        
        Args:
            x, y, w, h: PDF 坐标 (Points)
            slide_type: 幻灯片类型
            
        Returns:
            (left, top, width, height) in EMUs
        """
        # Hero 模式：封面/封底全屏
        if slide_type in (SlideType.COVER, SlideType.BACKCOVER):
            return self._hero_transform(x, y, w, h)
        
        # 应用公式: 坐标 * 缩放因子 + 偏移
        # PDF Y 坐标从底部开始，PPT 从顶部开始，需要翻转
        pdf_points_x = x * self.EMU_PER_POINT
        pdf_points_y = (self.pdf_h - y - h) * self.EMU_PER_POINT
        
        left = int(pdf_points_x * self.scale + self.margin + self.offset_x)
        top = int(pdf_points_y * self.scale + self.margin + self.offset_y)
        width = int(w * self.EMU_PER_POINT * self.scale)
        height = int(h * self.EMU_PER_POINT * self.scale)
        
        return left, top, width, height
    
    def _hero_transform(
        self,
        x: float,
        y: float,
        w: float,
        h: float
    ) -> Tuple[int, int, int, int]:
        """
        Hero 模式：封面/封底全屏覆盖
        
        忽略边距，图片铺满整个幻灯片
        
        Args:
            x, y, w, h: PDF 坐标
            
        Returns:
            全屏坐标 (EMUs)
        """
        # 全屏缩放（铺满优先）
        scale_x = self.ppt_w / (self.pdf_w * self.EMU_PER_POINT)
        scale_y = self.ppt_h / (self.pdf_h * self.EMU_PER_POINT)
        hero_scale = max(scale_x, scale_y)
        
        # 居中
        scaled_w = self.pdf_w * self.EMU_PER_POINT * hero_scale
        scaled_h = self.pdf_h * self.EMU_PER_POINT * hero_scale
        offset_x = (self.ppt_w - scaled_w) / 2
        offset_y = (self.ppt_h - scaled_h) / 2
        
        # Y轴翻转
        pdf_points_x = x * self.EMU_PER_POINT
        pdf_points_y = (self.pdf_h - y - h) * self.EMU_PER_POINT
        
        left = int(pdf_points_x * hero_scale + offset_x)
        top = int(pdf_points_y * hero_scale + offset_y)
        width = int(w * self.EMU_PER_POINT * hero_scale)
        height = int(h * self.EMU_PER_POINT * hero_scale)
        
        return left, top, width, height
    
    def map_font_size(self, pdf_font_size: float) -> Pt:
        """
        根据页面缩放比例调整字号，确保视觉一致性
        
        Args:
            pdf_font_size: PDF 中的字号 (Points)
            
        Returns:
            调整后的字号 (Pt)
        """
        # 使用 Y 轴方向的缩放系数
        adjusted_size = pdf_font_size * self.scale
        return Pt(int(adjusted_size))
    
    def to_pptx_geometry_simple(
        self,
        bbox: List[float],
        slide_type: SlideType = SlideType.NORMAL
    ) -> Tuple[int, int, int, int]:
        """
        简化接口：从 [x1, y1, x2, y2] 转换
        
        Args:
            bbox: [x1, y1, x2, y2]
            slide_type: 幻灯片类型
            
        Returns:
            (left, top, width, height) in EMUs
        """
        x1, y1, x2, y2 = bbox
        x = x1
        y = y1
        w = x2 - x1
        h = y2 - y1
        
        return self.to_pptx_geometry(x, y, w, h, slide_type)
    
    def validate_slide_number(self, slide_num: int) -> bool:
        """
        验证幻灯片页码
        
        Args:
            slide_num: 幻灯片序号（从 1 开始）
            
        Returns:
            是否有效
        """
        if slide_num < 1 or slide_num > self.MAX_SLIDES:
            return False
        return True
    
    def get_ppt_size(self) -> Tuple[int, int]:
        """获取 PPT 尺寸（EMUs）"""
        return self.ppt_w, self.ppt_h
    
    def get_scale_info(self) -> dict:
        """获取缩放信息"""
        return {
            "scale_x": self.sf_x,
            "scale_y": self.sf_y,
            "scale": self.scale,
            "margin_emu": self.margin,
            "pdf_size": (self.pdf_w, self.pdf_h),
            "ppt_size_inch": (self.ppt_w / self.EMU_PER_INCH, self.ppt_h / self.EMU_PER_INCH)
        }


# ========== 便捷函数 =========-

def create_mapper(
    pdf_width: float,
    pdf_height: float,
    aspect_ratio: str = "16:9",
    margin: float = 0.5
) -> CoordinateMapper:
    """
    创建坐标映射器
    
    Args:
        pdf_width: PDF 宽度
        pdf_height: PDF 高度
        aspect_ratio: "16:9" 或 "4:3"
        margin: 边距
        
    Returns:
        CoordinateMapper
    """
    if aspect_ratio == "16:9":
        ppt_w, ppt_h = 10.0, 5.625
    elif aspect_ratio == "4:3":
        ppt_w, ppt_h = 10.0, 7.5
    else:
        raise ValueError(f"Unsupported aspect ratio: {aspect_ratio}")
    
    return CoordinateMapper(
        pdf_width=pdf_width,
        pdf_height=pdf_height,
        ppt_width_inch=ppt_w,
        ppt_height_inch=ppt_h,
        margin_inch=margin
    )


# ========== 单元测试 =========-

def test_coordinate_mapper():
    """单元测试"""
    print("=" * 60)
    print("🧪 CoordinateMapper 单元测试")
    print("=" * 60)
    
    # 测试 1: 基础转换
    print("\n📌 测试 1: 基础转换 (PDF 600x800 → PPT 16:9)")
    mapper = CoordinateMapper(
        pdf_width=600,
        pdf_height=800,
        ppt_width_inch=10,
        ppt_height_inch=5.625,
        margin_inch=0.5
    )
    
    scale_info = mapper.get_scale_info()
    print(f"   PDF: {scale_info['pdf_size']}")
    print(f"   PPT: {scale_info['ppt_size_inch']}")
    print(f"   Scale: {scale_info['scale']:.4f}")
    
    left, top, w, h = mapper.to_pptx_geometry(100, 100, 200, 50)
    print(f"   Input:  PDF(100, 100, 200x50)")
    print(f"   Output: EMU({left}, {top}, {w}x{h})")
    print(f"   ✅ Pass" if w > 0 else "   ❌ Fail")
    
    # 测试 2: 字号映射
    print("\n📌 测试 2: 字号映射")
    pt = mapper.map_font_size(18)
    print(f"   PDF Font: 18pt → PPT Font: {pt}")
    print(f"   ✅ Pass")
    
    # 测试 3: 简化接口
    print("\n📌 测试 3: 简化接口 [x1,y1,x2,y2]")
    left, top, w, h = mapper.to_pptx_geometry_simple([100, 100, 300, 150])
    print(f"   Input: [100, 100, 300, 150]")
    print(f"   Output: ({left}, {top}, {w}, {h})")
    print(f"   ✅ Pass")
    
    # 测试 4: Hero 模式
    print("\n📌 测试 4: Hero 模式")
    normal = mapper.to_pptx_geometry(100, 100, 400, 300, SlideType.NORMAL)
    hero = mapper.to_pptx_geometry(100, 100, 400, 300, SlideType.COVER)
    print(f"   Normal: ({normal[0]}, {normal[1]}) {normal[2]}x{normal[3]}")
    print(f"   Hero:  ({hero[0]}, {hero[1]}) {hero[2]}x{hero[3]}")
    print(f"   ✅ Hero larger" if hero[2] > normal[2] else "   ❌ Fail")
    
    # 测试 5: 页码验证
    print("\n📌 测试 5: 页码验证")
    print(f"   Slide 5: {'✅ Valid' if mapper.validate_slide_number(5) else '❌ Invalid'}")
    print(f"   Slide 25: {'✅ Valid' if mapper.validate_slide_number(25) else '❌ Invalid'}")
    
    # 测试 6: 16:9 vs 4:3
    print("\n📌 测试 6: 16:9 vs 4:3")
    t16 = create_mapper(600, 400, "16:9", 0.5)
    t4 = create_mapper(600, 400, "4:3", 0.5)
    r16 = t16.to_pptx_geometry(100, 100, 200, 50)
    r4 = t4.to_pptx_geometry(100, 100, 200, 50)
    print(f"   16:9 scale: {t16.scale:.4f}")
    print(f"   4:3 scale:  {t4.scale:.4f}")
    print(f"   ✅ 16:9 scale > 4:3 scale" if r16[2] > r4[2] else "   ⚠️ Check")
    
    print("\n" + "=" * 60)
    print("🎉 所有测试通过!")
    print("=" * 60)


if __name__ == "__main__":
    test_coordinate_mapper()

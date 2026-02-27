"""
RedecktoPPT - Coordinate Transformer
PDF 坐标 → PPT EMU 转换器

功能：
1. PDF 坐标 → PPT EMU 转换
2. 缩放比例计算
3. 边界处理
4. Hero 模式（封面/封底全屏）
"""

import warnings
from dataclasses import dataclass
from typing import Tuple, Optional, List
from enum import Enum


class SlideType(Enum):
    """幻灯片类型"""
    NORMAL = "normal"       # 普通页面
    COVER = "cover"         # 封面（第一页）
    BACKCOVER = "backcover" # 封底（最后一页）
    SECTION = "section"     # 章节页


@dataclass
class TransformConfig:
    """转换配置"""
    # PDF 尺寸（points，1 inch = 72 points）
    pdf_width: float      # PDF 宽度
    pdf_height: float     # PDF 高度
    
    # PPT 尺寸（inches）
    ppt_width: float = 10.0      # 10 inches
    ppt_height: float = 5.625    # 16:9 = 5.625 inches
    
    # 边距（inches）
    margin_left: float = 0.5
    margin_right: float = 0.5
    margin_top: float = 0.5
    margin_bottom: float = 0.5
    
    # 转换常量
    POINTS_PER_INCH: float = 72.0
    EMUS_PER_POINT: int = 12700


@dataclass
class TransformResult:
    """转换结果"""
    left: int      # EMU
    top: int       # EMU
    width: int     # EMU
    height: int    # EMU
    scale: float
    was_truncated: bool = False
    warnings: List[str] = None
    
    def __post_init__(self):
        if self.warnings is None:
            self.warnings = []


class CoordinateTransformer:
    """
    坐标转换器：PDF → PPT EMU
    
    转换公式：
    1. 计算有效区域：ppt_size - 2 * margin
    2. 缩放比例 = min(ppt有效宽 / pdf宽, ppt有效高 / pdf高)
    3. 居中偏移 = (ppt有效区 - pdf实际区 * scale) / 2
    4. PPT坐标 = (PDF坐标 - margin) * scale + 偏移
    5. Y轴翻转：PDF原点在左下，PPT在左上
    6. EMU = points * 12700
    """
    
    # 默认 16:9 尺寸
    DEFAULT_PPT_WIDTH = 10.0    # inches
    DEFAULT_PPT_HEIGHT = 5.625  # inches
    
    # 最大幻灯片页码（商业限制）
    MAX_SLIDES = 20
    
    def __init__(
        self,
        pdf_width: float,
        pdf_height: float,
        ppt_width: float = None,
        ppt_height: float = None,
        margin: float = 0.5,
        hero_mode: bool = False
    ):
        """
        初始化转换器
        
        Args:
            pdf_width: PDF 宽度（points）
            pdf_height: PDF 高度（points）
            ppt_width: PPT 宽度（inches），默认 10"
            ppt_height: PPT 高度（inches），默认 5.625" (16:9)
            margin: 边距（inches）
            hero_mode: 是否启用 Hero 模式（封面全屏）
        """
        self.pdf_width = pdf_width
        self.pdf_height = pdf_height
        
        self.ppt_width = ppt_width or self.DEFAULT_PPT_WIDTH
        self.ppt_height = ppt_height or self.DEFAULT_PPT_HEIGHT
        
        # 转换常量
        self.points_per_inch = 72.0
        self.emus_per_point = 12700
        
        # 边距（转换为 points）
        self.margin = margin
        self.margin_pt = margin * self.points_per_inch
        
        # Hero 模式
        self.hero_mode = hero_mode
        
        # 计算基础参数
        self._compute_scale()
    
    def _compute_scale(self):
        """计算缩放比例和偏移"""
        # PPT 有效区域（减去边距）
        ppt_content_w = (self.ppt_width - 2 * self.margin) * self.points_per_inch
        ppt_content_h = (self.ppt_height - 2 * self.margin) * self.points_per_inch
        
        # 计算缩放比例（保持比例）
        self.scale_x = ppt_content_w / self.pdf_width
        self.scale_y = ppt_content_h / self.pdf_height
        self.scale = min(self.scale_x, self.scale_y)
        
        # 计算居中偏移
        scaled_w = self.pdf_width * self.scale
        scaled_h = self.pdf_height * self.scale
        
        self.offset_x = (ppt_content_w - scaled_w) / 2
        self.offset_y = (ppt_content_h - scaled_h) / 2
    
    def to_emu(
        self,
        x: float,
        y: float,
        width: float,
        height: float,
        slide_type: SlideType = SlideType.NORMAL,
        enable_padding: bool = True
    ) -> TransformResult:
        """
        将 PDF 坐标转换为 PPT EMU
        
        Args:
            x, y, width, height: PDF 坐标（points）
            slide_type: 幻灯片类型
            enable_padding: 是否启用边距
            
        Returns:
            TransformResult: 包含 left, top, width, height (EMU) 及元数据
        """
        warnings = []
        was_truncated = False
        
        # Hero 模式处理
        if slide_type in (SlideType.COVER, SlideType.BACKCOVER) and self.hero_mode:
            return self._hero_transform(x, y, width, height)
        
        # 应用边距
        margin_pt = self.margin_pt if enable_padding else 0
        
        # 计算 PPT 坐标（points）
        ppt_x = (x - margin_pt) * self.scale + self.offset_x + margin_pt * self.scale
        ppt_y = (self.pdf_height - y - height - margin_pt) * self.scale + self.margin_pt * self.scale + self.offset_y
        
        ppt_width = width * self.scale
        ppt_height = height * self.scale
        
        # 转换为 EMU
        left = int(ppt_x * self.emus_per_point)
        top = int(ppt_y * self.emus_per_point)
        width_emu = int(ppt_width * self.emus_per_point)
        height_emu = int(ppt_height * self.emus_per_point)
        
        # 边界检查
        ppt_width_pt = self.ppt_width * self.points_per_inch
        ppt_height_pt = self.ppt_height * self.points_per_inch
        
        if left < 0:
            warnings.append(f"左侧超出边界: {left} < 0")
            left = 0
            was_truncated = True
        
        if top < 0:
            warnings.append(f"顶部超出边界: {top} < 0")
            top = 0
            was_truncated = True
        
        max_left = int(ppt_width_pt * self.emus_per_point) - width_emu
        if left > max_left:
            warnings.append(f"右侧超出边界: {left} > {max_left}")
            left = max_left
            was_truncated = True
        
        max_top = int(ppt_height_pt * self.emus_per_point) - height_emu
        if top > max_top:
            warnings.append(f"底部超出边界: {top} > {max_top}")
            top = max_top
            was_truncated = True
        
        return TransformResult(
            left=left,
            top=top,
            width=width_emu,
            height=height_emu,
            scale=self.scale,
            was_truncated=was_truncated,
            warnings=warnings
        )
    
    def _hero_transform(
        self,
        x: float,
        y: float,
        width: float,
        height: float
    ) -> TransformResult:
        """
        Hero 模式：全屏覆盖
        
        用于封面/封底，忽略边距，图片铺满整个幻灯片
        """
        # 计算全屏缩放
        scale_x = (self.ppt_width * self.points_per_inch) / self.pdf_width
        scale_y = (self.ppt_height * self.points_per_inch) / self.pdf_height
        hero_scale = max(scale_x, scale_y)  # 铺满优先
        
        # 居中
        scaled_w = self.pdf_width * hero_scale
        scaled_h = self.pdf_height * hero_scale
        offset_x = (self.ppt_width * self.points_per_inch - scaled_w) / 2
        offset_y = (self.ppt_height * self.points_per_inch - scaled_h) / 2
        
        # Y轴翻转
        ppt_x = x * hero_scale
        ppt_y = (self.pdf_height - y - height) * hero_scale
        
        left = int(ppt_x * self.emus_per_point)
        top = int(ppt_y * self.emus_per_point)
        width_emu = int(width * hero_scale * self.emus_per_point)
        height_emu = int(height * hero_scale * self.emus_per_point)
        
        return TransformResult(
            left=left,
            top=top,
            width=width_emu,
            height=height_emu,
            scale=hero_scale,
            was_truncated=False,
            warnings=["Hero mode: full-bleed cover"]
        )
    
    def to_emu_simple(
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
            (left, top, width, height) in EMU
        """
        x1, y1, x2, y2 = bbox
        x = x1
        y = y1
        width = x2 - x1
        height = y2 - y1
        
        result = self.to_emu(x, y, width, height, slide_type)
        return result.left, result.top, result.width, result.height
    
    def get_ppt_size(self) -> Tuple[int, int]:
        """获取 PPT 尺寸（EMU）"""
        width = int(self.ppt_width * self.points_per_inch * self.emus_per_point)
        height = int(self.ppt_height * self.points_per_inch * self.emus_per_point)
        return width, height
    
    def validate_slide_number(self, slide_num: int) -> Tuple[bool, Optional[str]]:
        """
        验证幻灯片页码
        
        Args:
            slide_num: 幻灯片序号（从 1 开始）
            
        Returns:
            (是否有效, 警告信息)
        """
        if slide_num < 1:
            return False, f"幻灯片序号不能小于 1"
        
        if slide_num > self.MAX_SLIDES:
            msg = f"⚠️ 超过商业限制 ({self.MAX_SLIDES} 页): 第 {slide_num} 页"
            warnings.warn(msg)
            return False, msg
        
        return True, None


# ========== 便捷函数 =========-

def create_transformer(
    pdf_width: float,
    pdf_height: float,
    aspect_ratio: str = "16:9",
    margin: float = 0.5
) -> CoordinateTransformer:
    """
    创建转换器
    
    Args:
        pdf_width: PDF 宽度
        pdf_height: PDF 高度
        aspect_ratio: "16:9" 或 "4:3"
        margin: 边距
        
    Returns:
        CoordinateTransformer
    """
    if aspect_ratio == "16:9":
        ppt_w, ppt_h = 10.0, 5.625
    elif aspect_ratio == "4:3":
        ppt_w, ppt_h = 10.0, 7.5
    else:
        raise ValueError(f"Unsupported aspect ratio: {aspect_ratio}")
    
    return CoordinateTransformer(
        pdf_width=pdf_width,
        pdf_height=pdf_height,
        ppt_width=ppt_w,
        ppt_height=ppt_h,
        margin=margin
    )


# ========== 单元测试 =========-

def test_coordinate_transformer():
    """单元测试"""
    print("=" * 60)
    print("🧪 CoordinateTransformer 单元测试")
    print("=" * 60)
    
    # 测试 1: 基础转换
    print("\n📌 测试 1: 基础转换 (PDF 600x800 → PPT 16:9)")
    transformer = CoordinateTransformer(
        pdf_width=600,
        pdf_height=800,
        margin=0.5
    )
    
    print(f"   PDF: {transformer.pdf_width}x{transformer.pdf_height} points")
    print(f"   PPT: {transformer.ppt_width}x{transformer.ppt_height} inches")
    print(f"   Scale: {transformer.scale:.4f}")
    
    result = transformer.to_emu(100, 100, 200, 50)
    print(f"   Input:  PDF(100, 100, 200x50)")
    print(f"   Output: EMU({result.left}, {result.top}, {result.width}x{result.height})")
    print(f"   ✅ Pass" if result.width > 0 else "   ❌ Fail")
    
    # 测试 2: 边距效果
    print("\n📌 测试 2: 边距效果 (margin=0.5 vs margin=0)")
    result_no_margin = transformer.to_emu(100, 100, 200, 50, enable_padding=False)
    print(f"   With padding:  ({result.left}, {result.top})")
    print(f"   Without padding: ({result_no_margin.left}, {result_no_margin.top})")
    print(f"   ✅ Padding works" if result.left != result_no_margin.left else "   ❌ Fail")
    
    # 测试 3: 边界截断
    print("\n📌 测试 3: 边界截断")
    transformer2 = CoordinateTransformer(
        pdf_width=100,
        pdf_height=100,
        margin=0.1
    )
    result = transformer2.to_emu(50, 50, 500, 500)  # 超大元素
    print(f"   Input: PDF(50, 50, 500x500) - 超出边界")
    print(f"   Output: EMU({result.left}, {result.top}, {result.width}x{result.height})")
    print(f"   Warnings: {result.warnings}")
    print(f"   Was truncated: {result.was_truncated}")
    print(f"   ✅ Pass" if result.was_truncated else "   ⚠️  Should truncate")
    
    # 测试 4: Hero 模式
    print("\n📌 测试 4: Hero 模式（封面全屏）")
    transformer3 = CoordinateTransformer(
        pdf_width=600,
        pdf_height=800,
        margin=0.5,
        hero_mode=True
    )
    
    # 普通转换
    normal = transformer3.to_emu(100, 100, 400, 300, SlideType.NORMAL)
    # Hero 转换
    hero = transformer3.to_emu(100, 100, 400, 300, SlideType.COVER)
    
    print(f"   Normal:  ({normal.left}, {normal.top}) {normal.width}x{normal.height}")
    print(f"   Hero:   ({hero.left}, {hero.top}) {hero.width}x{hero.height}")
    print(f"   Scale:  Normal={normal.scale:.4f}, Hero={hero.scale:.4f}")
    print(f"   ✅ Hero mode larger" if hero.width > normal.width else "   ❌ Fail")
    
    # 测试 5: 页码验证
    print("\n📌 测试 5: 页码验证")
    valid, msg = transformer.validate_slide_number(5)
    print(f"   Slide 5: {'✅ Valid' if valid else '❌ ' + msg}")
    
    valid, msg = transformer.validate_slide_number(25)
    print(f"   Slide 25: {'✅ Valid' if valid else '⚠️  ' + msg}")
    
    # 测试 6: 简化接口
    print("\n📌 测试 6: 简化接口 to_emu_simple([x1, y1, x2, y2])")
    left, top, w, h = transformer.to_emu_simple([100, 100, 300, 150])
    print(f"   Input: [100, 100, 300, 150]")
    print(f"   Output: ({left}, {top}, {w}, {h})")
    print(f"   ✅ Pass")
    
    # 测试 7: 16:9 vs 4:3
    print("\n📌 测试 7: 16:9 vs 4:3")
    t_16by9 = create_transformer(600, 400, "16:9", 0.5)
    t_4by3 = create_transformer(600, 400, "4:3", 0.5)
    
    r16 = t_16by9.to_emu(100, 100, 200, 50)
    r4 = t_4by3.to_emu(100, 100, 200, 50)
    
    print(f"   16:9 scale: {r16.scale:.4f}")
    print(f"   4:3 scale:  {r4.scale:.4f}")
    print(f"   ✅ 16:9 scale > 4:3 scale" if r16.scale > r4.scale else "   ⚠️  Check logic")
    
    print("\n" + "=" * 60)
    print("🎉 所有测试完成!")
    print("=" * 60)


if __name__ == "__main__":
    test_coordinate_transformer()

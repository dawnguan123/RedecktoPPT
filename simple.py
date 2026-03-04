#!/usr/bin/env python3
"""
RedecktoPPT - 简化版
直接从 PDF 提取图片，转换为 PPT (去水印版)

用法:
    python simple.py input.pdf output.pptx
"""

import sys
import os
from pathlib import Path

import fitz  # PyMuPDF
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor


def pdf_to_images(pdf_path: str, dpi: int = 150) -> list:
    """从 PDF 提取页面为图片"""
    images = []
    
    doc = fitz.open(pdf_path)
    for page_num, page in enumerate(doc):
        # 渲染页面为图片
        zoom = dpi / 72
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat)
        
        # 保存到临时文件
        img_path = f"/tmp/redeck_page_{page_num}.png"
        pix.save(img_path)
        images.append(img_path)
        print(f"   📄 提取第 {page_num + 1} 页")
    
    doc.close()
    return images


def create_ppt(images: list, output_path: str, watermark_height: float = 0.15):
    """
    创建 PPT，每页一张图片
    
    Args:
        images: 图片路径列表
        output_path: 输出路径
        watermark_height: 底部水印区域高度（英寸），默认 0.15
    """
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(5.625)  # 16:9
    
    for img_path in images:
        slide = prs.slides.add_slide(prs.slide_layouts[6])  # 空白布局
        
        # 添加图片，铺满整个幻灯片
        slide.shapes.add_picture(
            img_path,
            Inches(0), Inches(0),
            width=Inches(10), height=Inches(5.625)
        )
        
        # 覆盖底部水印区域（白色矩形）
        shape = slide.shapes.add_shape(
            1,  # msoShapeRectangle
            Inches(0),
            Inches(5.625 - watermark_height),
            Inches(10),
            Inches(watermark_height)
        )
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(255, 255, 255)  # 白色
        shape.line.fill.background()  # 无边框
        
        print(f"   🖼️ 添加图片: {Path(img_path).name}")
    
    prs.save(output_path)
    print(f"\n✨ PPT 生成成功: {output_path}")


def cleanup(images: list):
    """清理临时文件"""
    for img in images:
        try:
            os.remove(img)
        except:
            pass


def main():
    if len(sys.argv) < 3:
        print("用法: python simple.py 输入.pdf 输出.pptx")
        print("       (自动覆盖底部水印区域)")
        sys.exit(1)
    
    pdf_path = sys.argv[1]
    output_path = sys.argv[2]
    
    if not os.path.exists(pdf_path):
        print(f"❌ 文件不存在: {pdf_path}")
        sys.exit(1)
    
    print(f"🎬 开始转换: {Path(pdf_path).name}")
    print("-" * 40)
    
    # 提取图片
    print("📄 提取页面...")
    images = pdf_to_images(pdf_path)
    print(f"   ✅ 提取了 {len(images)} 页")
    
    # 创建 PPT (覆盖底部水印)
    print("\n🎨 生成 PPT (覆盖水印区域)...")
    create_ppt(images, output_path)
    
    # 清理
    cleanup(images)
    print("\n🎉 完成!")


if __name__ == "__main__":
    main()

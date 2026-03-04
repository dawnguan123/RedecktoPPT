#!/usr/bin/env python3
"""
PPTX 转 PDF（通过提取图片）
用于 NotebookLM 生成的 PPTX
"""

import sys
import os
import zipfile
from pathlib import Path
import shutil
import fitz  # PyMuPDF
import re


def natural_sort_key(s):
    """自然排序 key: image1.png -> (1,), image10.png -> (10,)"""
    return [int(c) if c.isdigit() else c.lower() for c in re.split(r'(\d+)', s)]


def extract_images_from_pptx(pptx_path: str) -> list:
    """从 PPTX 中提取图片"""
    images = []
    temp_dir = f"/tmp/pptx_extract_{os.getpid()}"
    
    try:
        # 解压 PPTX
        with zipfile.ZipFile(pptx_path, 'r') as z:
            z.extractall(temp_dir)
        
        # 查找 media 文件夹
        media_dir = os.path.join(temp_dir, 'ppt', 'media')
        if not os.path.exists(media_dir):
            print(f"❌ 未找到 media 文件夹")
            return []
        
        # 获取所有图片（自然排序）
        img_files = sorted([f for f in os.listdir(media_dir) 
                          if f.endswith(('.png', '.jpg', '.jpeg'))],
                          key=natural_sort_key)
        
        for i, img_file in enumerate(img_files):
            src = os.path.join(media_dir, img_file)
            dst = f"/tmp/pptx_page_{i:03d}.png"
            shutil.copy(src, dst)
            images.append(dst)
            print(f"   📄 提取第 {i+1} 页: {img_file}")
        
    finally:
        # 清理临时文件
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)
    
    return images


def images_to_pdf(images: list, output_pdf: str):
    """图片转 PDF"""
    doc = fitz.open()
    
    for img_path in images:
        img = fitz.open(img_path)
        page = doc.new_page(width=img[0].rect.width, height=img[0].rect.height)
        page.insert_image(img[0].rect, filename=img_path)
        img.close()
    
    doc.save(output_pdf)
    doc.close()


def main():
    if len(sys.argv) < 3:
        print("用法: python pptx_to_pdf.py 输入.pptx 输出.pdf")
        sys.exit(1)
    
    pptx_path = sys.argv[1]
    output_pdf = sys.argv[2]
    
    if not os.path.exists(pptx_path):
        print(f"❌ 文件不存在: {pptx_path}")
        sys.exit(1)
    
    print(f"🎬 PPTX 转 PDF: {Path(pptx_path).name}")
    print("-" * 40)
    
    # 提取图片
    print("📄 提取图片...")
    images = extract_images_from_pptx(pptx_path)
    
    if not images:
        print("❌ 无法提取图片")
        sys.exit(1)
    
    print(f"   ✅ 提取了 {len(images)} 张图片")
    
    # 转 PDF
    print("\n📦 生成 PDF...")
    images_to_pdf(images, output_pdf)
    
    # 清理
    for img in images:
        os.remove(img)
    
    print(f"✨ PDF 生成成功: {output_pdf}")


if __name__ == "__main__":
    main()

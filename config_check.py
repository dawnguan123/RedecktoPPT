#!/usr/bin/env python3
"""
RedecktoPPT - 环境自检脚本
自动识别 Mac 芯片架构并给出优化建议
"""

import sys
import os
import platform
import subprocess


def check_status(name, condition, fix_msg):
    """检查状态并打印结果"""
    status = "✅ [PASS]" if condition else "❌ [FAIL]"
    print(f"{status} {name}")
    if not condition:
        print(f"   💡 建议: {fix_msg}")
    return condition


def run_check():
    """运行所有检测"""
    print("=" * 60)
    print("🚀 RedecktoPPT - macOS 环境初始化检测")
    print("=" * 60)
    print()
    
    # ==================== 第一部分 ====================
    
    # 1. 系统架构检测
    is_mac = platform.system() == "Darwin"
    is_arm = platform.machine() == "arm64"
    
    check_status(
        "macOS 系统检测",
        is_mac,
        "此项目目前主要针对 macOS 设计。"
    )
    check_status(
        "Apple Silicon (M1/M2/M3) 芯片检测",
        is_arm,
        "未检测到 ARM 架构，性能可能受限。"
    )
    
    # 2. Python 版本检测
    py_version = sys.version_info
    check_status(
        f"Python 版本 ({py_version.major}.{py_version.minor})",
        py_version.major == 3 and py_version.minor >= 10,
        "推荐使用 Python 3.10+ 以获得最佳兼容性。"
    )
    
    # 3. 深度学习后端与 MPS 加速检测
    try:
        import torch
        has_mps = torch.backends.mps.is_available()
        
        check_status("PyTorch 安装", True, "")
        check_status(
            "MPS (Metal) 硬件加速可用性",
            has_mps,
            "请安装最新版 PyTorch 并确保系统版本为 macOS 12.3+。"
        )
        
        if has_mps:
            print("   ⚡ 加速设备: Apple Silicon GPU (MPS)")
            # 测试 MPS 计算
            try:
                x = torch.randn(3, 3, device='mps')
                y = x + x
                print(f"   ✅ MPS 计算测试通过: {y.shape}")
            except Exception as e:
                print(f"   ⚠️ MPS 计算测试失败: {e}")
                
    except ImportError:
        check_status("PyTorch 安装", False, "请运行 'pip install torch torchvision'")
    
    # 4. MinerU (magic-pdf) 核心引擎检测
    try:
        import magic_pdf
        check_status("MinerU (magic-pdf) 核心库", True, "")
        
        # 检查 magic-pdf 版本
        version = getattr(magic_pdf, '__version__', 'unknown')
        print(f"   📦 版本: {version}")
        
    except ImportError:
        check_status(
            "MinerU (magic-pdf) 核心库",
            False,
            "请运行 'pip install magic-pdf[full]'"
        )
    
    print()
    print("-" * 60)
    
    # ==================== 第二部分 ====================
    
    # 5. PPT 渲染引擎检测
    try:
        import pptx
        from pptx import __version__
        check_status("python-pptx 渲染引擎", True, "")
        print(f"   📦 版本: {__version__}")
    except ImportError:
        check_status(
            "python-pptx 渲染引擎",
            False,
            "请运行 'pip install python-pptx'"
        )
    
    # 6. OCR 引擎检测
    try:
        import paddleocr
        check_status("PaddleOCR 引擎", True, "")
    except ImportError:
        check_status(
            "PaddleOCR 引擎",
            False,
            "请运行 'pip install paddleocr'"
        )
    
    # 7. 其他依赖检测
    print()
    print("-" * 60)
    
    # PyMuPDF
    try:
        import fitz
        check_status("PyMuPDF", True, "")
    except ImportError:
        check_status("PyMuPDF", False, "pip install pymupdf")
    
    # OpenCV
    try:
        import cv2
        check_status("OpenCV", True, "")
    except ImportError:
        check_status("OpenCV", False, "pip install opencv-python")
    
    # SciPy
    try:
        import scipy
        check_status("SciPy", True, "")
    except ImportError:
        check_status("SciPy", False, "pip install scipy")
    
    # Pydantic
    try:
        import pydantic
        check_status("Pydantic", True, "")
    except ImportError:
        check_status("Pydantic", False, "pip install pydantic")
    
    # 8. 环境优化建议
    print()
    print("=" * 60)
    print("📍 检测完成")
    print("=" * 60)
    print()
    
    # 汇总
    checks = [
        ("macOS", is_mac),
        ("Apple Silicon", is_arm),
        ("Python 3.10+", py_version.major == 3 and py_version.minor >= 10),
    ]
    
    try:
        import torch
        checks.append(("PyTorch", True))
        checks.append(("MPS", has_mps))
    except:
        checks.append(("PyTorch", False))
    
    try:
        import magic_pdf
        checks.append(("magic-pdf", True))
    except:
        checks.append(("magic-pdf", False))
    
    try:
        import pptx
        checks.append(("python-pptx", True))
    except:
        checks.append(("python-pptx", False))
    
    try:
        import paddleocr
        checks.append(("PaddleOCR", True))
    except:
        checks.append(("PaddleOCR", False))
    
    passed = sum(1 for _, ok in checks if ok)
    total = len(checks)
    
    print(f"   通过: {passed}/{total}")
    
    if passed == total:
        print()
        print("🎉 所有检测通过！您可以开始开发核心逻辑。")
        print()
    else:
        print()
        print("⚠️  请安装缺失的依赖后重新检测")
        print("   pip install -r requirements.txt")
        print()
    
    # Apple Silicon 优化建议
    if is_arm and has_mps:
        print("🍎 Apple Silicon 优化建议:")
        print("   • 使用 MPS 后端加速模型推理")
        print("   • 推荐使用 PyTorch 2.2+ 以获得最佳性能")
        print("   • 确保 macOS 版本 >= 12.3")
    
    print()


if __name__ == "__main__":
    run_check()

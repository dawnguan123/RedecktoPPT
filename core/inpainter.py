"""
RedecktoPPT - Inpainter
文字移除模块：使用 scipy.griddata 进行高级修复
"""

import cv2
import numpy as np
from scipy import interpolate


class Inpainter:
    """
    高级 Inpainting
    
    使用 scipy.griddata 进行插值修复
    配合边缘羽化实现自然过渡
    """
    
    def __init__(self):
        self.method = 'linear'
    
    def inpaint(
        self,
        image: np.ndarray,
        mask: np.ndarray,
        dilate: int = 3
    ) -> np.ndarray:
        """
        执行 inpainting
        
        Args:
            image: BGR 图片
            mask: 掩码（255 = 需要修复，0 = 保留）
            dilate: 膨胀迭代次数
            
        Returns:
            修复后的图片
        """
        # 膨胀掩码，确保边缘被覆盖
        if dilate > 0:
            kernel = np.ones((dilate, dilate), np.uint8)
            mask = cv2.dilate(mask, kernel, iterations=1)
        
        # 转换为 RGB（OpenCV BGR → NumPy RGB）
        img_rgb = cv2.cvtColor(image, cv2.COLOR_BGR2RGB)
        
        # 执行修复
        result = self._inpaint_via_griddata(img_rgb, mask)
        
        # 转回 BGR
        result = cv2.cvtColor(result, cv2.COLOR_RGB2BGR)
        
        return result
    
    def _inpaint_via_griddata(
        self,
        image: np.ndarray,
        mask: np.ndarray
    ) -> np.ndarray:
        """
        使用 scipy.griddata 进行插值修复
        
        原理：
        - 识别掩码区域（需要修复的部分）
        - 使用周围非掩码区域的像素值
        - 通过插值拟合平滑背景曲面
        """
        h, w = image.shape[:2]
        
        # 创建坐标网格
        x, y = np.meshgrid(np.arange(w), np.arange(h))
        
        # 找到有效点和需要插值的点
        valid_mask = (mask == 0)
        fill_mask = (mask > 0)
        
        # 有效像素坐标和值
        valid_coords = np.column_stack([
            x[valid_mask],
            y[valid_mask]
        ])
        valid_values = image[valid_mask]
        
        # 需要填充的坐标
        fill_coords = np.column_stack([
            x[fill_mask],
            y[fill_mask]
        ])
        
        if len(fill_coords) == 0:
            return image
        
        if len(valid_coords) == 0:
            # 没有有效像素，返回原图
            return image
        
        # 使用 griddata 插值
        # method='linear' 对渐变背景效果最好
        try:
            interp_r = interpolate.griddata(
                valid_coords,
                valid_values[:, 0],
                fill_coords,
                method=self.method,
                fill_value=0
            )
            interp_g = interpolate.griddata(
                valid_coords,
                valid_values[:, 1],
                fill_coords,
                method=self.method,
                fill_value=0
            )
            interp_b = interpolate.griddata(
                valid_coords,
                valid_values[:, 2],
                fill_coords,
                method=self.method,
                fill_value=0
            )
        except Exception as e:
            print(f"⚠️  Griddata failed: {e}, using cv2.inpaint")
            # 回退到 cv2.inpaint
            mask_uint8 = (mask > 0).astype(np.uint8) * 255
            return cv2.inpaint(image, mask_uint8, 3, cv2.INPAINT_TELEA)
        
        # 创建结果图像
        result = image.copy()
        
        # 填充插值结果
        result[fill_mask, 0] = interp_r.astype(np.uint8)
        result[fill_mask, 1] = interp_g.astype(np.uint8)
        result[fill_mask, 2] = interp_b.astype(np.uint8)
        
        # 边缘羽化
        result = self._apply_edge_feathering(result, mask, image)
        
        return result
    
    def _apply_edge_feathering(
        self,
        result: np.ndarray,
        mask: np.ndarray,
        original: np.ndarray,
        iterations: int = 15
    ) -> np.ndarray:
        """
        边缘羽化：拉普拉斯平滑
        
        让修复区域与原图边缘自然过渡
        """
        # 创建边缘掩码
        kernel = np.ones((5, 5), np.uint8)
        dilated = cv2.dilate(mask, kernel, iterations=2)
        eroded = cv2.erode(mask, kernel, iterations=2)
        edge_mask = dilated - eroded
        
        # 边缘区域为需要平滑的区域
        edge_mask = (edge_mask > 0).astype(np.uint8)
        
        # 迭代平滑
        for _ in range(iterations):
            # 对边缘区域进行均值滤波
            blurred = cv2.blur(result, (3, 3))
            
            # 混合
            mask_3ch = np.stack([edge_mask] * 3, axis=-1)
            result = np.where(mask_3ch > 0, blurred, result)
        
        return result


# ========== 便捷函数 ==========

def remove_text(image: np.ndarray, text_boxes: List) -> np.ndarray:
    """
    便捷函数：移除文字
    
    Args:
        image: 输入图片
        text_boxes: 文字区域列表 [(x1, y1, x2, y2), ...]
        
    Returns:
        清理后的图片
    """
    h, w = image.shape[:2]
    mask = np.zeros((h, w), dtype=np.uint8)
    
    for bbox in text_boxes:
        x1, y1, x2, y2 = [int(v) for v in bbox]
        mask[y1:y2, x1:x2] = 255
    
    inpainter = Inpainter()
    return inpainter.inpaint(image, mask)


if __name__ == "__main__":
    print("Inpainter module ready")
    print("Usage:")
    print('  inpainter = Inpainter()')
    print('  result = inpainter.inpaint(image, mask)')

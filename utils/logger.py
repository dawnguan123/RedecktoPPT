"""
RedecktoPPT - Logger Utility
日志配置
"""

import sys
from loguru import logger


def setup_logger(verbose: bool = False) -> logger:
    """
    配置日志
    
    Args:
        verbose: 是否详细输出
        
    Returns:
        logger 实例
    """
    # 移除默认处理器
    logger.remove()
    
    # 添加控制台处理器
    logger.add(
        sys.stderr,
        format="<green>{time:YYYY-MM-DD HH:mm:ss}</green> | <level>{level: <8}</level> | <level>{message}</level>",
        level="DEBUG" if verbose else "INFO",
        colorize=True
    )
    
    return logger


# 导出默认 logger
default_logger = logger

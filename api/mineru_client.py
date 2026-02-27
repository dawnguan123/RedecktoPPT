"""
RedecktoPPT - MinerU API Client
MinerU API 客户端，用于获取 PDF 深度解析结果
"""

import os
import json
import time
import requests
from typing import Optional, Dict, Any, List
from pathlib import Path


class MinerUClient:
    """MinerU API 客户端"""
    
    BASE_URL = "https://api.mineru.cn/v1"
    
    def __init__(self, api_key: str = None):
        self.api_key = api_key or os.environ.get('MINERU_API_KEY', '')
        if not self.api_key:
            raise ValueError("MINERU_API_KEY is required")
        
        self.session = requests.Session()
        self.session.headers.update({
            'Authorization': f'Bearer {self.api_key}',
            'User-Agent': 'RedecktoPPT/1.0'
        })
    
    def upload_file(self, file_path: str) -> Dict[str, Any]:
        """
        上传 PDF 文件
        
        Args:
            file_path: PDF 文件路径
            
        Returns:
            文件信息字典 (包含 file_id)
        """
        url = f"{self.BASE_URL}/files/upload"
        
        with open(file_path, 'rb') as f:
            files = {'file': (Path(file_path).name, f, 'application/pdf')}
            response = self.session.post(url, files=files, timeout=60)
        
        response.raise_for_status()
        return response.json()
    
    def parse(self, file_id: str, options: Dict[str, Any] = None) -> str:
        """
        提交解析任务
        
        Args:
            file_id: 文件 ID
            options: 解析选项
            
        Returns:
            task_id
        """
        url = f"{self.BASE_URL}/parse"
        
        data = {
            'file_id': file_id,
            'options': options or {
                'layout_analysis': True,
                'extract_images': True,
                'extract_tables': True,
            }
        }
        
        response = self.session.post(url, json=data, timeout=30)
        response.raise_for_status()
        
        result = response.json()
        return result.get('task_id')
    
    def get_result(self, task_id: str, timeout: int = 300) -> Dict[str, Any]:
        """
        获取解析结果
        
        Args:
            task_id: 任务 ID
            timeout: 超时时间（秒）
            
        Returns:
            解析结果字典
        """
        url = f"{self.BASE_URL}/tasks/{task_id}"
        
        start_time = time.time()
        while time.time() - start_time < timeout:
            response = self.session.get(url, timeout=30)
            response.raise_for_status()
            
            result = response.json()
            status = result.get('status')
            
            if status == 'completed':
                return result.get('data', {})
            elif status == 'failed':
                raise RuntimeError(f"Parse failed: {result.get('error')}")
            
            time.sleep(2)
        
        raise TimeoutError(f"Parse timeout after {timeout}s")
    
    def parse_pdf(self, pdf_path: str, output_dir: str = None) -> str:
        """
        一站式解析 PDF
        
        Args:
            pdf_path: PDF 文件路径
            output_dir: 输出目录（可选）
            
        Returns:
            JSON 文件路径
        """
        # 1. 上传
        print(f"📤 Uploading {pdf_path}...")
        file_info = self.upload_file(pdf_path)
        file_id = file_info['id']
        print(f"   File ID: {file_id}")
        
        # 2. 解析
        print(f"🔄 Parsing...")
        task_id = self.parse(file_id)
        print(f"   Task ID: {task_id}")
        
        # 3. 获取结果
        print(f"📥 Fetching result...")
        result = self.get_result(task_id)
        
        # 4. 保存
        if output_dir is None:
            output_dir = Path(pdf_path).parent
        
        output_path = Path(output_dir) / f"{Path(pdf_path).stem}_mineru.json"
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        
        print(f"✅ Saved to: {output_path}")
        return str(output_path)


# 便捷函数
def quick_parse(pdf_path: str, api_key: str = None) -> str:
    """
    快速解析 PDF
    
    Args:
        pdf_path: PDF 文件路径
        api_key: API 密钥
        
    Returns:
        JSON 文件路径
    """
    client = MinerUClient(api_key)
    return client.parse_pdf(pdf_path)


if __name__ == "__main__":
    import sys
    
    if len(sys.argv) < 2:
        print("Usage: python mineru_client.py <pdf_file> [api_key]")
        sys.exit(1)
    
    pdf_file = sys.argv[1]
    api_key = sys.argv[2] if len(sys.argv) > 2 else None
    
    result = quick_parse(pdf_file, api_key)
    print(f"\n✅ Result: {result}")

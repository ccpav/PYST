#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
DOCX到PDF转换工具 (静默模式)
此脚本将当前目录下的所有DOCX文件转换为PDF格式，确保最大兼容性和格式正确性。
整个过程完全在后台运行，不会显示任何Office应用程序窗口，不会影响用户操作。
"""

import os
import sys
import time
import argparse
import logging
import win32com.client
import pythoncom
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed

try:
    from tqdm import tqdm
except ImportError:
    print("正在安装所需依赖库...")
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "tqdm", "pywin32"])
    print("依赖库安装完成，重新导入...")
    from tqdm import tqdm

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("docx_to_pdf_conversion.log"),
        logging.StreamHandler()
    ]
)

logger = logging.getLogger(__name__)

def convert_single_file(docx_path, output_dir=None):
    """
    转换单个DOCX文件为PDF，完全静默操作
    
    Args:
        docx_path: DOCX文件路径
        output_dir: 输出目录，默认为None（与原文件相同目录）
    
    Returns:
        bool: 转换是否成功
    """
    try:
        # 在每个线程中初始化COM库
        pythoncom.CoInitialize()
        
        docx_path = Path(docx_path).resolve()
        
        # 确定输出路径
        if output_dir:
            output_path = Path(output_dir) / f"{docx_path.stem}.pdf"
        else:
            output_path = docx_path.with_suffix('.pdf')
        
        # 检查输出文件是否已存在
        if output_path.exists():
            logger.info(f"PDF文件已存在: {output_path}，跳过转换")
            return True
        
        # 确保输出目录存在
        if output_dir:
            Path(output_dir).mkdir(parents=True, exist_ok=True)
        
        logger.info(f"正在转换: {docx_path}")
        
        # 直接使用Word COM对象，而不是docx2pdf库，以获得更多控制
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False  # 不显示Word窗口
        word.DisplayAlerts = False  # 不显示任何警告或对话框
        
        try:
            # 尝试打开文档
            doc = word.Documents.Open(str(docx_path), ReadOnly=True)
            
            # 保存为PDF (17是PDF格式的WdSaveFormat常量)
            doc.SaveAs(str(output_path), FileFormat=17)
            doc.Close(False)  # 关闭文档不保存更改
            
            logger.info(f"转换成功: {output_path}")
            return True
        
        except Exception as inner_e:
            logger.error(f"转换文档时出错: {str(inner_e)}")
            return False
        
        finally:
            # 确保无论如何都关闭Word应用
            try:
                word.Quit()
            except:
                pass
    
    except Exception as e:
        logger.error(f"转换 {docx_path} 失败: {str(e)}")
        return False
    
    finally:
        # 释放COM库
        pythoncom.CoUninitialize()

def find_docx_files(directory, recursive=False):
    """
    查找指定目录下的所有DOCX文件
    
    Args:
        directory: 要搜索的目录
        recursive: 是否递归搜索子目录
    
    Returns:
        list: DOCX文件路径列表
    """
    directory = Path(directory)
    if recursive:
        docx_files = list(directory.glob("**/*.docx"))
    else:
        docx_files = list(directory.glob("*.docx"))
    return docx_files

def convert_all_files(directory, output_dir=None, max_workers=None, recursive=False):
    """
    转换指定目录下的所有DOCX文件为PDF
    
    Args:
        directory: 源目录
        output_dir: 输出目录，默认为None（与原文件相同目录）
        max_workers: 最大工作线程数
        recursive: 是否递归处理子目录
    """
    directory = Path(directory)
    
    # 确保输出目录存在
    if output_dir:
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
    
    # 查找所有DOCX文件
    docx_files = find_docx_files(directory, recursive)
    
    if not docx_files:
        logger.warning(f"在 {directory} 中没有找到DOCX文件")
        print(f"\n在 {directory} 中没有找到DOCX文件")
        return
    
    logger.info(f"找到 {len(docx_files)} 个DOCX文件")
    print(f"找到 {len(docx_files)} 个DOCX文件")
    
    # 使用线程池并行处理文件
    successful = 0
    failed = 0
    
    # 控制并发数，避免创建过多Word实例
    if max_workers is None:
        max_workers = min(4, os.cpu_count() or 4)
    
    # 创建进度条
    with tqdm(total=len(docx_files), desc="转换进度", unit="文件") as pbar:
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            # 提交所有转换任务
            future_to_file = {executor.submit(convert_single_file, docx_file, output_dir): docx_file 
                             for docx_file in docx_files}
            
            # 处理完成的任务
            for future in as_completed(future_to_file):
                docx_file = future_to_file[future]
                try:
                    if future.result():
                        successful += 1
                    else:
                        failed += 1
                except Exception as e:
                    logger.error(f"处理 {docx_file} 时发生异常: {str(e)}")
                    failed += 1
                pbar.update(1)
    
    logger.info(f"转换完成: 成功 {successful} 个, 失败 {failed} 个")
    print(f"\n转换完成: 成功 {successful} 个, 失败 {failed} 个")
    
    if failed > 0:
        logger.warning("部分文件转换失败，请查看日志了解详细信息")
        print("部分文件转换失败，请查看日志了解详细信息")

def main():
    """主函数"""
    parser = argparse.ArgumentParser(description='将DOCX文件转换为PDF格式 (静默模式)')
    parser.add_argument('-d', '--directory', default='.', 
                        help='要处理的目录 (默认: 当前目录)')
    parser.add_argument('-o', '--output', default=None,
                        help='输出目录 (默认: 与原文件相同目录)')
    parser.add_argument('-w', '--workers', type=int, default=2,
                        help='并行处理的最大线程数 (默认: 2)')
    parser.add_argument('-r', '--recursive', action='store_true',
                        help='递归处理子目录')
    parser.add_argument('-s', '--silent', action='store_true', default=True,
                        help='静默运行，不显示Word窗口 (默认: 开启)')
    
    args = parser.parse_args()
    
    print("DOCX到PDF转换工具 (静默模式) 启动")
    print(f"处理目录: {args.directory}")
    if args.output:
        print(f"输出目录: {args.output}")
    if args.recursive:
        print("将递归处理所有子目录")
    print(f"并行转换线程数: {args.workers}")
    
    # 开始转换
    convert_all_files(args.directory, args.output, args.workers, args.recursive)

if __name__ == "__main__":
    main() 
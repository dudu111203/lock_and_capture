#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Windows锁定并拍照工具
运行后立即通过摄像头拍照，然后锁定Windows系统
单文件设计，双击即可使用
"""

import os
import sys
import time
import datetime
import cv2
import ctypes
from ctypes import wintypes
import logging

# 应用信息
APP_NAME = "LockAndCapture"
VERSION = "1.0"

# 配置日志
def setup_logger():
    """设置日志记录器"""
    logger = logging.getLogger(APP_NAME)
    logger.setLevel(logging.INFO)
    
    # 清除已有的处理器
    logger.handlers.clear()
    
    # 创建格式化器
    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    
    # 创建控制台处理器
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    
    # 创建文件处理器
    try:
        # 确定日志文件路径（支持打包后）
        if getattr(sys, 'frozen', False):
            log_dir = os.path.dirname(os.path.abspath(sys.executable))
        else:
            log_dir = os.path.dirname(os.path.abspath(__file__))
        log_path = os.path.join(log_dir, f"{APP_NAME}.log")
        file_handler = logging.FileHandler(log_path, encoding='utf-8')
        file_handler.setLevel(logging.INFO)
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
    except Exception as e:
        logger.error(f"无法创建日志文件: {str(e)}")
    
    return logger

# 全局日志记录器
logger = setup_logger()

class LockAndCaptureApp:
    """锁定并拍照应用"""
    
    def __init__(self):
        """初始化应用"""
        self.logger = logging.getLogger(f"{APP_NAME}.App")
        
        # 获取应用目录（支持打包后）
        if getattr(sys, 'frozen', False):
            # 打包后的可执行文件路径
            self.app_dir = os.path.dirname(os.path.abspath(sys.executable))
        else:
            # 开发环境下的脚本路径
            self.app_dir = os.path.dirname(os.path.abspath(__file__))
        
        # 加载Windows API
        self._load_windows_api()
    
    def _load_windows_api(self):
        """加载Windows API"""
        try:
            # 加载User32.dll
            self.user32 = ctypes.WinDLL('User32.dll')
            
            # 锁定工作站函数
            self.user32.LockWorkStation.argtypes = []
            self.user32.LockWorkStation.restype = wintypes.BOOL
            
            # 消息框函数
            self.user32.MessageBoxW.argtypes = [
                wintypes.HWND, wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.UINT
            ]
            self.user32.MessageBoxW.restype = wintypes.INT
            
            # 消息框样式
            self.MB_OK = 0x00000000
            self.MB_OKCANCEL = 0x00000001
            self.MB_ICONINFORMATION = 0x00000040
            self.MB_ICONWARNING = 0x00000030
            self.MB_ICONERROR = 0x00000010
            
        except Exception as e:
            self.logger.error(f"加载Windows API失败: {str(e)}")
            raise
    
    def show_message(self, message, title=APP_NAME, icon_type=None):
        """显示Windows消息框
        
        Args:
            message: 消息内容
            title: 标题
            icon_type: 图标类型
            
        Returns:
            int: 消息框返回值
        """
        if icon_type is None:
            icon_type = self.MB_ICONINFORMATION
            
        return self.user32.MessageBoxW(None, message, title, icon_type)
    
    def detect_camera(self):
        """检测可用的摄像头
        
        Returns:
            int: 第一个可用摄像头的ID，如果没有找到返回-1
        """
        try:
            self.logger.info("开始检测摄像头...")
            
            # 尝试检测前5个摄像头
            for i in range(5):
                cap = cv2.VideoCapture(i)
                if cap.isOpened():
                    # 尝试读取一帧以确认摄像头正常工作
                    ret, _ = cap.read()
                    if ret:
                        cap.release()
                        self.logger.info(f"检测到摄像头: ID {i}")
                        return i
                    cap.release()
            
            self.logger.warning("未检测到可用摄像头")
            return -1
            
        except Exception as e:
            self.logger.error(f"检测摄像头时发生错误: {str(e)}")
            return -1
    
    def capture_image(self, camera_id):
        """使用指定摄像头捕获图像
        
        Args:
            camera_id: 摄像头ID
            
        Returns:
            str: 保存的图像路径，如果失败返回None
        """
        cap = None
        try:
            # 打开摄像头
            cap = cv2.VideoCapture(camera_id)
            
            if not cap.isOpened():
                self.logger.error(f"无法打开摄像头 {camera_id}")
                return None
            
            # 设置摄像头参数
            cap.set(cv2.CAP_PROP_FRAME_WIDTH, 1280)
            cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 720)
            
            # 等待摄像头初始化
            self.logger.info("摄像头初始化中...")
            time.sleep(0.5)
            
            # 连续捕获几帧，取最后一帧（让摄像头有时间调整）
            for i in range(3):
                ret, frame = cap.read()
                if not ret:
                    self.logger.error(f"第 {i+1} 次读取图像失败")
                    continue
            
            if frame is None:
                self.logger.error("无法捕获图像")
                return None
            
            # 生成文件名
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"capture_{timestamp}.jpg"
            filepath = os.path.join(self.app_dir, filename)
            
            # 保存图像
            success = cv2.imwrite(filepath, frame)
            
            if success:
                self.logger.info(f"成功捕获图像并保存至: {filepath}")
                return filepath
            else:
                self.logger.error(f"保存图像失败: {filepath}")
                return None
                
        except Exception as e:
            self.logger.error(f"捕获图像时发生错误: {str(e)}")
            return None
        finally:
            if cap is not None:
                cap.release()
    
    def lock_workstation(self):
        """锁定Windows工作站"""
        try:
            result = self.user32.LockWorkStation()
            
            if result:
                self.logger.info("系统已锁定")
                return True
            else:
                self.logger.error("锁定系统失败")
                return False
                
        except Exception as e:
            self.logger.error(f"锁定系统时发生错误: {str(e)}")
            return False
    
    def run(self):
        """运行应用主流程"""
        try:
            self.logger.info(f"启动 {APP_NAME} v{VERSION}")
            
            # 显示提示消息
            self.show_message(
                "系统即将锁定并拍照，请准备...",
                APP_NAME,
                self.MB_ICONINFORMATION
            )
            
            # 检测摄像头
            camera_id = self.detect_camera()
            
            if camera_id == -1:
                # 未检测到摄像头，询问是否继续
                result = self.show_message(
                    "未检测到可用摄像头，是否仍要锁定系统？",
                    APP_NAME,
                    self.MB_ICONWARNING | self.MB_OKCANCEL
                )
                
                if result != 1:  # OK按钮返回值为1
                    self.logger.info("用户取消操作")
                    return
            
            # 捕获图像
            if camera_id != -1:
                image_path = self.capture_image(camera_id)
                
                if image_path:
                    self.logger.info(f"图像已保存: {image_path}")
                else:
                    self.show_message(
                        "拍照失败，但系统仍将锁定",
                        APP_NAME,
                        self.MB_ICONWARNING
                    )
            
            # 锁定系统
            self.logger.info("准备锁定系统...")
            time.sleep(1)  # 短暂延迟，让用户看到提示
            
            success = self.lock_workstation()
            
            if not success:
                self.show_message(
                    "锁定系统失败",
                    APP_NAME,
                    self.MB_ICONERROR
                )
                
        except KeyboardInterrupt:
            self.logger.info("用户中断操作")
        except Exception as e:
            self.logger.error(f"应用运行时发生错误: {str(e)}")
            self.show_message(
                f"应用运行时发生错误:\n{str(e)}",
                APP_NAME,
                self.MB_ICONERROR
            )
        finally:
            self.logger.info("应用退出")

def main():
    """应用入口"""
    try:
        # 创建并运行应用
        app = LockAndCaptureApp()
        app.run()
    except Exception as e:
        # 如果无法创建应用实例，直接显示错误消息
        try:
            ctypes.windll.user32.MessageBoxW(
                None,
                f"应用启动失败:\n{str(e)}",
                APP_NAME,
                0x00000010  # MB_ICONERROR
            )
        except:
            print(f"应用启动失败: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()
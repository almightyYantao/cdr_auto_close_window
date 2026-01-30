#!/usr/bin/env python3
"""
CorelDRAW 弹窗自动处理工具
自动检测并处理 CorelDRAW 打开文件时的各种错误弹窗
"""

import time
import sys
import ctypes
from ctypes import wintypes

# Windows API
user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

# 常量
WM_CLOSE = 0x0010
BM_CLICK = 0x00F5
CB_SETCURSEL = 0x014E
CB_GETCURSEL = 0x0147
WM_COMMAND = 0x0111
BN_CLICKED = 0

# 回调函数类型
EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
EnumChildProc = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)


def get_window_text(hwnd):
    """获取窗口标题"""
    length = user32.GetWindowTextLengthW(hwnd) + 1
    buffer = ctypes.create_unicode_buffer(length)
    user32.GetWindowTextW(hwnd, buffer, length)
    return buffer.value


def get_class_name(hwnd):
    """获取窗口类名"""
    buffer = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(hwnd, buffer, 256)
    return buffer.value


def find_child_windows(parent_hwnd):
    """查找所有子窗口"""
    children = []
    
    def callback(hwnd, lparam):
        children.append(hwnd)
        return True
    
    user32.EnumChildWindows(parent_hwnd, EnumChildProc(callback), 0)
    return children


def click_button(hwnd):
    """点击按钮"""
    user32.SendMessageW(hwnd, BM_CLICK, 0, 0)


def find_and_click_button(parent_hwnd, button_text):
    """在父窗口中查找并点击指定按钮"""
    children = find_child_windows(parent_hwnd)
    for child in children:
        text = get_window_text(child)
        class_name = get_class_name(child)
        if button_text.lower() in text.lower() and 'Button' in class_name:
            click_button(child)
            return True
    return False


def select_radio_button(parent_hwnd, radio_text):
    """选择单选按钮"""
    children = find_child_windows(parent_hwnd)
    for child in children:
        text = get_window_text(child)
        class_name = get_class_name(child)
        if radio_text.lower() in text.lower() and 'Button' in class_name:
            click_button(child)
            return True
    return False


def get_dialog_text(hwnd):
    """获取对话框中的所有文本"""
    texts = []
    children = find_child_windows(hwnd)
    for child in children:
        text = get_window_text(child)
        if text:
            texts.append(text)
    return ' '.join(texts)


def handle_popup(hwnd, title, content):
    """根据弹窗内容处理"""
    
    # 1. 无效的轮廓 ID - 点击忽略
    if '无效的轮廓' in content or '轮廓 ID' in content:
        if find_and_click_button(hwnd, '忽略'):
            print(f"[✅] 处理: 无效的轮廓 ID -> 点击忽略")
            return True
    
    # 2. 无效标头 / 无法打开文件 - 点击 OK
    if '无法打开文件' in content or '无效标头' in content:
        if find_and_click_button(hwnd, 'OK'):
            print(f"[✅] 处理: 无效标头 -> 点击 OK")
            return True
    
    # 3. 文件被损坏 - 点击 OK
    if '文件被损坏' in content or '损坏' in content:
        if find_and_click_button(hwnd, 'OK'):
            print(f"[✅] 处理: 文件被损坏 -> 点击 OK")
            return True
    
    # 4. 导入 PS/PRN - 选择曲线，点击 OK
    if 'PS/PRN' in content or 'PS/PRN' in title:
        # 先选择曲线
        if select_radio_button(hwnd, '曲线'):
            print(f"[✅] 处理: 导入 PS/PRN -> 选择曲线")
            time.sleep(0.2)
        if find_and_click_button(hwnd, 'OK'):
            print(f"[✅] 处理: 导入 PS/PRN -> 点击 OK")
            return True
    
    return False


def find_coreldraw_popups():
    """查找 CorelDRAW 相关的弹窗"""
    popups = []
    
    def callback(hwnd, lparam):
        if user32.IsWindowVisible(hwnd):
            title = get_window_text(hwnd)
            class_name = get_class_name(hwnd)
            
            # 查找 CorelDRAW 的对话框
            if 'CorelDRAW' in title or (class_name == '#32770' and title):
                # #32770 是 Windows 对话框的类名
                popups.append((hwnd, title, class_name))
        return True
    
    user32.EnumWindows(EnumWindowsProc(callback), 0)
    return popups


def main():
    print("=" * 50)
    print("CorelDRAW 弹窗自动处理工具")
    print("=" * 50)
    print("正在监控 CorelDRAW 弹窗...")
    print("按 Ctrl+C 退出")
    print()
    print("支持的弹窗类型:")
    print("  1. 无效的轮廓 ID -> 点击忽略")
    print("  2. 无效标头/无法打开文件 -> 点击 OK")
    print("  3. 文件被损坏 -> 点击 OK")
    print("  4. 导入 PS/PRN -> 选择曲线 + 点击 OK")
    print()
    print("-" * 50)
    
    handled_count = 0
    
    try:
        while True:
            popups = find_coreldraw_popups()
            
            for hwnd, title, class_name in popups:
                content = get_dialog_text(hwnd)
                
                if handle_popup(hwnd, title, content):
                    handled_count += 1
                    print(f"    (已处理 {handled_count} 个弹窗)")
            
            time.sleep(0.3)  # 每 300ms 检查一次
            
    except KeyboardInterrupt:
        print()
        print("-" * 50)
        print(f"程序已退出，共处理 {handled_count} 个弹窗")


if __name__ == "__main__":
    # 检查是否在 Windows 上运行
    if sys.platform != 'win32':
        print("❌ 此程序只能在 Windows 上运行！")
        sys.exit(1)
    
    main()

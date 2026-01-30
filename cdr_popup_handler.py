#!/usr/bin/env python3
"""
CorelDRAW 弹窗自动处理工具
自动检测并处理 CorelDRAW 打开文件时的各种错误弹窗
"""

import time
import sys
import ctypes
from ctypes import wintypes
from datetime import datetime

# Windows API
user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

# 常量
WM_CLOSE = 0x0010
BM_CLICK = 0x00F5
WM_LBUTTONDOWN = 0x0201
WM_LBUTTONUP = 0x0202

# 回调函数类型
EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
EnumChildProc = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)


def log(msg):
    """打印带时间戳的日志"""
    timestamp = datetime.now().strftime("%H:%M:%S")
    print(f"[{timestamp}] {msg}")


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


def is_window_visible(hwnd):
    """检查窗口是否可见"""
    return user32.IsWindowVisible(hwnd)


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
    log(f"    -> 尝试点击按钮 hwnd={hwnd}")
    # 方法1: BM_CLICK
    result = user32.SendMessageW(hwnd, BM_CLICK, 0, 0)
    log(f"    -> BM_CLICK 结果: {result}")
    return True


def find_and_click_button(parent_hwnd, button_texts):
    """在父窗口中查找并点击指定按钮"""
    if isinstance(button_texts, str):
        button_texts = [button_texts]
    
    children = find_child_windows(parent_hwnd)
    log(f"  找到 {len(children)} 个子控件")
    
    for child in children:
        text = get_window_text(child)
        class_name = get_class_name(child)
        
        if text:
            log(f"    子控件: '{text}' (类: {class_name}, hwnd: {child})")
        
        for btn_text in button_texts:
            if btn_text.lower() in text.lower():
                if 'Button' in class_name:
                    log(f"  >>> 找到目标按钮: '{text}'")
                    click_button(child)
                    return True
    
    return False


def select_radio_and_click_ok(parent_hwnd, radio_text):
    """选择单选按钮并点击OK"""
    children = find_child_windows(parent_hwnd)
    
    # 先找并点击单选按钮
    for child in children:
        text = get_window_text(child)
        class_name = get_class_name(child)
        
        if radio_text.lower() in text.lower() and 'Button' in class_name:
            log(f"  >>> 选择单选按钮: '{text}'")
            click_button(child)
            time.sleep(0.3)
            break
    
    # 再点击 OK
    time.sleep(0.2)
    return find_and_click_button(parent_hwnd, ['OK', '确定'])


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
    log(f"处理弹窗: title='{title}'")
    log(f"  内容: {content[:200]}...")
    
    # 1. 无效的轮廓 ID - 点击忽略
    if '无效的轮廓' in content or '轮廓 ID' in content or '轮廓ID' in content:
        log("  -> 匹配: 无效的轮廓 ID")
        if find_and_click_button(hwnd, ['忽略', 'Ignore', '忽略(I)', '忽略(&I)']):
            log("  ✅ 成功点击忽略")
            return True
        else:
            log("  ❌ 未找到忽略按钮")
    
    # 2. 无效标头 / 无法打开文件 - 点击 OK
    if '无法打开文件' in content or '无效标头' in content:
        log("  -> 匹配: 无效标头")
        if find_and_click_button(hwnd, ['OK', '确定']):
            log("  ✅ 成功点击 OK")
            return True
    
    # 3. 文件被损坏 - 点击 OK
    if '文件被损坏' in content or '损坏' in content:
        log("  -> 匹配: 文件被损坏")
        if find_and_click_button(hwnd, ['OK', '确定']):
            log("  ✅ 成功点击 OK")
            return True
    
    # 4. 导入 PS/PRN - 选择曲线，点击 OK
    if 'PS/PRN' in content or 'PS/PRN' in title:
        log("  -> 匹配: 导入 PS/PRN")
        if select_radio_and_click_ok(hwnd, '曲线'):
            log("  ✅ 成功选择曲线并点击 OK")
            return True
    
    log("  -> 未匹配任何规则")
    return False


def find_all_windows():
    """查找所有顶层窗口"""
    windows = []
    
    def callback(hwnd, lparam):
        if is_window_visible(hwnd):
            title = get_window_text(hwnd)
            class_name = get_class_name(hwnd)
            if title or class_name == '#32770':  # #32770 是对话框类
                windows.append((hwnd, title, class_name))
        return True
    
    user32.EnumWindows(EnumWindowsProc(callback), 0)
    return windows


def find_coreldraw_popups():
    """查找 CorelDRAW 相关的弹窗"""
    popups = []
    all_windows = find_all_windows()
    
    for hwnd, title, class_name in all_windows:
        # CorelDRAW 对话框通常标题包含 CorelDRAW 或者是 #32770 类
        if 'CorelDRAW' in title:
            content = get_dialog_text(hwnd)
            popups.append((hwnd, title, class_name, content))
        elif class_name == '#32770':
            # 检查内容是否与 CorelDRAW 相关
            content = get_dialog_text(hwnd)
            if any(kw in content for kw in ['轮廓', '无效', 'CorelDRAW', 'CDR', '文件', 'PS/PRN']):
                popups.append((hwnd, title, class_name, content))
    
    return popups


def main():
    print("=" * 60)
    print("CorelDRAW 弹窗自动处理工具 v2.0")
    print("=" * 60)
    print()
    log("程序启动")
    log("正在监控 CorelDRAW 弹窗...")
    log("按 Ctrl+C 退出")
    print()
    print("支持的弹窗类型:")
    print("  1. 无效的轮廓 ID -> 点击忽略")
    print("  2. 无效标头/无法打开文件 -> 点击 OK")
    print("  3. 文件被损坏 -> 点击 OK")
    print("  4. 导入 PS/PRN -> 选择曲线 + 点击 OK")
    print()
    print("-" * 60)
    
    handled_count = 0
    scan_count = 0
    handled_hwnds = set()  # 记录已处理的窗口，避免重复处理
    
    try:
        while True:
            scan_count += 1
            if scan_count % 10 == 0:  # 每10次扫描输出一次状态
                log(f"扫描中... (已扫描 {scan_count} 次, 已处理 {handled_count} 个弹窗)")
            
            popups = find_coreldraw_popups()
            
            if popups:
                log(f"发现 {len(popups)} 个可能的弹窗")
            
            for hwnd, title, class_name, content in popups:
                if hwnd in handled_hwnds:
                    continue
                
                log(f"检测到弹窗: hwnd={hwnd}, title='{title}', class='{class_name}'")
                
                if handle_popup(hwnd, title, content):
                    handled_count += 1
                    handled_hwnds.add(hwnd)
                    log(f"✅ 已处理 {handled_count} 个弹窗")
            
            # 清理已关闭的窗口记录
            handled_hwnds = {h for h in handled_hwnds if user32.IsWindow(h)}
            
            time.sleep(0.5)  # 每 500ms 检查一次
            
    except KeyboardInterrupt:
        print()
        log("-" * 60)
        log(f"程序已退出，共处理 {handled_count} 个弹窗")


if __name__ == "__main__":
    # 检查是否在 Windows 上运行
    if sys.platform != 'win32':
        print("❌ 此程序只能在 Windows 上运行！")
        sys.exit(1)
    
    main()

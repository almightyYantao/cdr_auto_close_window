#!/usr/bin/env python3
"""
CorelDRAW 弹窗自动处理工具 v3.0
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

# 消息常量
BM_CLICK = 0x00F5
WM_GETTEXT = 0x000D
WM_GETTEXTLENGTH = 0x000E

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


def get_control_text(hwnd):
    """获取控件文本（包括静态文本控件）"""
    # 先尝试 GetWindowText
    text = get_window_text(hwnd)
    if text:
        return text
    
    # 再尝试 WM_GETTEXT（对静态控件更有效）
    length = user32.SendMessageW(hwnd, WM_GETTEXTLENGTH, 0, 0)
    if length > 0:
        buffer = ctypes.create_unicode_buffer(length + 1)
        user32.SendMessageW(hwnd, WM_GETTEXT, length + 1, buffer)
        return buffer.value
    
    return ""


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
    log(f"    -> 点击按钮 hwnd={hwnd}")
    result = user32.SendMessageW(hwnd, BM_CLICK, 0, 0)
    time.sleep(0.1)
    return True


def get_all_dialog_content(hwnd):
    """获取对话框中的所有文本内容（包括静态文本）"""
    all_texts = []
    children = find_child_windows(hwnd)
    
    for child in children:
        class_name = get_class_name(child)
        text = get_control_text(child)
        
        if text:
            all_texts.append(text)
        
        # 特别处理静态文本控件
        if class_name == 'Static' or 'STATIC' in class_name.upper():
            # 确保获取静态文本
            static_text = get_control_text(child)
            if static_text and static_text not in all_texts:
                all_texts.append(static_text)
    
    return ' '.join(all_texts)


def find_and_click_button_by_text(parent_hwnd, button_texts):
    """根据按钮文本查找并点击按钮"""
    if isinstance(button_texts, str):
        button_texts = [button_texts]
    
    children = find_child_windows(parent_hwnd)
    
    for child in children:
        text = get_control_text(child)
        class_name = get_class_name(child)
        
        # 检查是否是按钮
        if 'Button' in class_name or 'BUTTON' in class_name.upper():
            for btn_text in button_texts:
                # 支持多种匹配方式
                if (btn_text.lower() in text.lower() or 
                    btn_text.replace('&', '') in text or
                    text.replace('&', '') in btn_text or
                    btn_text == text):
                    log(f"  >>> 找到按钮: '{text}' (匹配: '{btn_text}')")
                    click_button(child)
                    return True
    
    return False


def select_radio_and_click_ok(parent_hwnd, radio_text):
    """选择单选按钮并点击OK"""
    children = find_child_windows(parent_hwnd)
    
    # 先找并点击单选按钮
    for child in children:
        text = get_control_text(child)
        class_name = get_class_name(child)
        
        if 'Button' in class_name and radio_text.lower() in text.lower():
            log(f"  >>> 选择单选按钮: '{text}'")
            click_button(child)
            time.sleep(0.3)
            break
    
    # 再点击 OK
    time.sleep(0.2)
    return find_and_click_button_by_text(parent_hwnd, ['OK', '确定'])


def handle_popup(hwnd, title):
    """根据弹窗处理"""
    # 获取完整的对话框内容
    content = get_all_dialog_content(hwnd)
    
    log(f"弹窗标题: '{title}'")
    log(f"弹窗内容: {content[:300]}...")
    
    # 列出所有子控件（调试用）
    children = find_child_windows(hwnd)
    log(f"  子控件数量: {len(children)}")
    for child in children:
        child_text = get_control_text(child)
        child_class = get_class_name(child)
        if child_text:
            log(f"    - [{child_class}] '{child_text}'")
    
    # === 规则匹配 ===
    
    # 1. 无效的轮廓 ID - 点击忽略
    # 关键词匹配：检查内容或按钮是否包含相关关键词
    if ('无效' in content and '轮廓' in content) or '轮廓 ID' in content or '轮廓ID' in content:
        log("  -> 匹配规则: 无效的轮廓 ID")
        if find_and_click_button_by_text(hwnd, ['忽略(&I)', '忽略', 'Ignore']):
            log("  ✅ 成功点击忽略")
            return True
    
    # 2. 检查按钮组合：如果有"关于、重试、忽略"三个按钮，通常是轮廓ID错误
    if '关于' in content and '重试' in content and '忽略' in content:
        log("  -> 匹配规则: 关于/重试/忽略 按钮组合（可能是轮廓ID错误）")
        if find_and_click_button_by_text(hwnd, ['忽略(&I)', '忽略', 'Ignore']):
            log("  ✅ 成功点击忽略")
            return True
    
    # 3. 无效标头 / 无法打开文件 - 点击 OK
    if '无法打开文件' in content or '无效标头' in content:
        log("  -> 匹配规则: 无效标头")
        if find_and_click_button_by_text(hwnd, ['OK', '确定']):
            log("  ✅ 成功点击 OK")
            return True
    
    # 4. 文件被损坏 - 点击 OK
    if '文件被损坏' in content or ('文件' in content and '损坏' in content):
        log("  -> 匹配规则: 文件被损坏")
        if find_and_click_button_by_text(hwnd, ['OK', '确定']):
            log("  ✅ 成功点击 OK")
            return True
    
    # 5. 导入 PS/PRN - 选择曲线，点击 OK
    if 'PS/PRN' in content or 'PS/PRN' in title:
        log("  -> 匹配规则: 导入 PS/PRN")
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
            windows.append((hwnd, title, class_name))
        return True
    
    user32.EnumWindows(EnumWindowsProc(callback), 0)
    return windows


def find_coreldraw_dialogs():
    """查找 CorelDRAW 相关的对话框"""
    dialogs = []
    all_windows = find_all_windows()
    
    for hwnd, title, class_name in all_windows:
        # 只关注 #32770 类（标准对话框）
        if class_name == '#32770':
            # 检查标题是否包含 CorelDRAW
            if 'CorelDRAW' in title or 'Corel' in title:
                dialogs.append((hwnd, title))
            else:
                # 检查内容是否相关
                content = get_all_dialog_content(hwnd)
                if any(kw in content for kw in ['轮廓', 'CorelDRAW', '文件', 'PS/PRN', '损坏']):
                    dialogs.append((hwnd, title))
    
    return dialogs


def main():
    print("=" * 60)
    print("CorelDRAW 弹窗自动处理工具 v3.0")
    print("=" * 60)
    print()
    log("程序启动")
    log("正在监控 CorelDRAW 弹窗...")
    log("按 Ctrl+C 退出")
    print()
    print("支持的弹窗类型:")
    print("  1. 无效的轮廓 ID -> 点击忽略")
    print("  2. 关于/重试/忽略 按钮组 -> 点击忽略")
    print("  3. 无效标头/无法打开文件 -> 点击 OK")
    print("  4. 文件被损坏 -> 点击 OK")
    print("  5. 导入 PS/PRN -> 选择曲线 + 点击 OK")
    print()
    print("-" * 60)
    
    handled_count = 0
    scan_count = 0
    handled_hwnds = set()
    
    try:
        while True:
            scan_count += 1
            if scan_count % 20 == 0:
                log(f"扫描中... (已扫描 {scan_count} 次, 已处理 {handled_count} 个弹窗)")
            
            dialogs = find_coreldraw_dialogs()
            
            for hwnd, title in dialogs:
                if hwnd in handled_hwnds:
                    continue
                
                log(f"=" * 50)
                log(f"检测到对话框: hwnd={hwnd}")
                
                if handle_popup(hwnd, title):
                    handled_count += 1
                    handled_hwnds.add(hwnd)
                    log(f"✅ 已处理 {handled_count} 个弹窗")
                    log(f"=" * 50)
            
            # 清理已关闭的窗口记录
            handled_hwnds = {h for h in handled_hwnds if user32.IsWindow(h)}
            
            time.sleep(0.5)
            
    except KeyboardInterrupt:
        print()
        log("-" * 60)
        log(f"程序已退出，共处理 {handled_count} 个弹窗")


if __name__ == "__main__":
    if sys.platform != 'win32':
        print("❌ 此程序只能在 Windows 上运行！")
        sys.exit(1)
    
    main()

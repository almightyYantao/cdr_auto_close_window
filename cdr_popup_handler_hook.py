#!/usr/bin/env python3
"""
CorelDRAW 弹窗自动处理工具 v5.0
结合 Hook 获取 GDI 文本 + 精确获取当前弹窗的按钮
"""

import time
import sys
import os
import ctypes
from ctypes import wintypes
from datetime import datetime

# Windows API
user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

# 常量
BM_CLICK = 0x00F5
WM_GETTEXT = 0x000D
WM_GETTEXTLENGTH = 0x000E
GWL_STYLE = -16
BS_DEFPUSHBUTTON = 0x0001
BS_PUSHBUTTON = 0x0000

PROCESS_ALL_ACCESS = 0x1F0FFF
MEM_COMMIT = 0x1000
MEM_RESERVE = 0x2000
PAGE_READWRITE = 0x04
MEM_RELEASE = 0x8000

# 共享内存
SHARED_MEM_NAME = "CDRPopupHandlerSharedMem"
MAX_TEXT_LENGTH = 4096
MAX_TEXT_COUNT = 100

# 回调类型
EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
EnumChildProc = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)


def log(msg):
    timestamp = datetime.now().strftime("%H:%M:%S")
    print(f"[{timestamp}] {msg}")


def get_window_text(hwnd):
    """获取窗口标题"""
    length = user32.GetWindowTextLengthW(hwnd) + 1
    buffer = ctypes.create_unicode_buffer(length)
    user32.GetWindowTextW(hwnd, buffer, length)
    return buffer.value


def get_control_text(hwnd):
    """获取控件文本（包括用 WM_GETTEXT）"""
    # 先用 GetWindowText
    text = get_window_text(hwnd)
    if text:
        return text
    
    # 再用 WM_GETTEXT
    length = user32.SendMessageW(hwnd, WM_GETTEXTLENGTH, 0, 0)
    if length > 0:
        buffer = ctypes.create_unicode_buffer(length + 1)
        user32.SendMessageW(hwnd, WM_GETTEXT, length + 1, buffer)
        return buffer.value
    
    return ""


def get_class_name(hwnd):
    """获取类名"""
    buffer = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(hwnd, buffer, 256)
    return buffer.value


def is_button(hwnd):
    """检查是否是按钮"""
    class_name = get_class_name(hwnd)
    return 'Button' in class_name or 'BUTTON' in class_name.upper()


def find_child_windows(parent_hwnd):
    """获取所有子窗口"""
    children = []
    def callback(hwnd, lparam):
        children.append(hwnd)
        return True
    user32.EnumChildWindows(parent_hwnd, EnumChildProc(callback), 0)
    return children


def get_dialog_info(hwnd):
    """获取对话框的所有信息"""
    info = {
        'title': get_window_text(hwnd),
        'buttons': [],
        'texts': [],
        'all_content': []
    }
    
    children = find_child_windows(hwnd)
    
    for child in children:
        text = get_control_text(child)
        class_name = get_class_name(child)
        
        if not text:
            continue
        
        info['all_content'].append(text)
        
        if is_button(child):
            info['buttons'].append({
                'hwnd': child,
                'text': text,
                'class': class_name
            })
        else:
            info['texts'].append(text)
    
    return info


def click_button(hwnd):
    """点击按钮"""
    user32.SendMessageW(hwnd, BM_CLICK, 0, 0)
    time.sleep(0.1)
    return True


def click_button_by_text(dialog_info, button_texts):
    """根据文本点击按钮"""
    if isinstance(button_texts, str):
        button_texts = [button_texts]
    
    for btn in dialog_info['buttons']:
        btn_text = btn['text']
        for target in button_texts:
            # 多种匹配方式
            if (target.lower() == btn_text.lower() or
                target.lower() in btn_text.lower() or
                target.replace('&', '') == btn_text.replace('&', '') or
                target in btn_text):
                log(f"  >>> 点击按钮: '{btn_text}'")
                click_button(btn['hwnd'])
                return True
    
    return False


class SharedMemory:
    """共享内存读取"""
    
    def __init__(self):
        self.handle = None
        self.size = 4 + 8 + (MAX_TEXT_COUNT * MAX_TEXT_LENGTH * 2) + 4
    
    def create(self):
        self.handle = kernel32.CreateFileMappingW(
            ctypes.c_void_p(-1), None, PAGE_READWRITE, 0, self.size, SHARED_MEM_NAME
        )
        return self.handle is not None
    
    def read_texts(self):
        if not self.handle:
            return []
        
        pBuf = kernel32.MapViewOfFile(self.handle, 0xF001F, 0, 0, self.size)
        if not pBuf:
            return []
        
        try:
            import struct
            count_buf = (ctypes.c_char * 4)()
            ctypes.memmove(count_buf, pBuf, 4)
            text_count = struct.unpack('I', bytes(count_buf))[0]
            
            if text_count > MAX_TEXT_COUNT:
                text_count = MAX_TEXT_COUNT
            
            texts = []
            offset = 4 + 8
            
            for i in range(text_count):
                text_buf = (ctypes.c_wchar * MAX_TEXT_LENGTH)()
                ctypes.memmove(text_buf, pBuf + offset, MAX_TEXT_LENGTH * 2)
                text = text_buf.value
                if text and len(text) > 1:  # 过滤单字符噪音
                    texts.append(text)
                offset += MAX_TEXT_LENGTH * 2
            
            return texts
        finally:
            kernel32.UnmapViewOfFile(pBuf)
    
    def clear(self):
        if not self.handle:
            return
        pBuf = kernel32.MapViewOfFile(self.handle, 0xF001F, 0, 0, self.size)
        if pBuf:
            zero = (ctypes.c_char * 4)()
            ctypes.memmove(pBuf, zero, 4)
            kernel32.UnmapViewOfFile(pBuf)
    
    def close(self):
        if self.handle:
            kernel32.CloseHandle(self.handle)
            self.handle = None


class DLLInjector:
    """DLL 注入器"""
    
    def __init__(self, dll_path):
        self.dll_path = os.path.abspath(dll_path)
        self.injected_pids = set()
    
    def inject(self, pid):
        if pid in self.injected_pids:
            return True
        
        hProcess = kernel32.OpenProcess(PROCESS_ALL_ACCESS, False, pid)
        if not hProcess:
            return False
        
        try:
            dll_path_bytes = (self.dll_path + '\0').encode('utf-16-le')
            dll_path_len = len(dll_path_bytes)
            
            pRemoteMem = kernel32.VirtualAllocEx(
                hProcess, None, dll_path_len, MEM_COMMIT | MEM_RESERVE, PAGE_READWRITE
            )
            if not pRemoteMem:
                return False
            
            written = ctypes.c_size_t()
            if not kernel32.WriteProcessMemory(hProcess, pRemoteMem, dll_path_bytes, dll_path_len, ctypes.byref(written)):
                kernel32.VirtualFreeEx(hProcess, pRemoteMem, 0, MEM_RELEASE)
                return False
            
            hKernel32 = kernel32.GetModuleHandleW("kernel32.dll")
            pLoadLibraryW = kernel32.GetProcAddress(hKernel32, b"LoadLibraryW")
            
            hThread = kernel32.CreateRemoteThread(
                hProcess, None, 0, pLoadLibraryW, pRemoteMem, 0, None
            )
            if not hThread:
                kernel32.VirtualFreeEx(hProcess, pRemoteMem, 0, MEM_RELEASE)
                return False
            
            kernel32.WaitForSingleObject(hThread, 5000)
            kernel32.CloseHandle(hThread)
            kernel32.VirtualFreeEx(hProcess, pRemoteMem, 0, MEM_RELEASE)
            
            self.injected_pids.add(pid)
            log(f"✅ DLL 注入成功: PID {pid}")
            return True
        finally:
            kernel32.CloseHandle(hProcess)
    
    def inject_coreldraw(self):
        def callback(hwnd, lparam):
            if user32.IsWindowVisible(hwnd):
                title = get_window_text(hwnd)
                if 'CorelDRAW' in title:
                    pid = ctypes.c_ulong()
                    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
                    if pid.value and pid.value not in self.injected_pids:
                        log(f"发现 CorelDRAW: PID {pid.value}")
                        self.inject(pid.value)
            return True
        user32.EnumWindows(EnumWindowsProc(callback), 0)


def handle_popup(hwnd, dialog_info, hook_texts):
    """处理弹窗"""
    title = dialog_info['title']
    buttons = [b['text'] for b in dialog_info['buttons']]
    texts = dialog_info['texts']
    
    # 合并所有文本来源
    all_text = ' '.join(dialog_info['all_content'] + hook_texts)
    
    log(f"=" * 50)
    log(f"弹窗标题: '{title}'")
    log(f"按钮列表: {buttons}")
    log(f"静态文本: {texts}")
    log(f"Hook文本: {hook_texts[:5]}...")  # 只显示前5个
    log(f"合并内容: {all_text[:200]}...")
    
    # ========== 规则匹配 ==========
    
    # 规则1: 无效的轮廓 ID -> 点击忽略
    if '无效' in all_text and '轮廓' in all_text:
        log("  -> 匹配: 无效的轮廓 ID")
        if click_button_by_text(dialog_info, ['忽略', '忽略(&I)', 'Ignore']):
            return True
    
    # 规则2: 如果只有一个 OK 按钮，直接点击
    if len(buttons) == 1 and ('OK' in buttons[0] or '确定' in buttons[0]):
        log("  -> 匹配: 单个 OK 按钮")
        if click_button_by_text(dialog_info, ['OK', '确定']):
            return True
    
    # 规则3: 无效标头 / 无法打开 -> OK
    if '无法打开' in all_text or '无效标头' in all_text or '无效的' in all_text:
        log("  -> 匹配: 无效/无法打开")
        if click_button_by_text(dialog_info, ['OK', '确定', '忽略', '忽略(&I)']):
            return True
    
    # 规则4: 文件损坏 -> OK
    if '损坏' in all_text:
        log("  -> 匹配: 文件损坏")
        if click_button_by_text(dialog_info, ['OK', '确定']):
            return True
    
    # 规则5: PS/PRN 导入
    if 'PS/PRN' in all_text:
        log("  -> 匹配: PS/PRN")
        click_button_by_text(dialog_info, ['曲线', '曲线(&C)'])
        time.sleep(0.2)
        if click_button_by_text(dialog_info, ['OK', '确定']):
            return True
    
    # 规则6: 有 忽略 按钮且包含错误关键词
    if any('忽略' in b for b in buttons):
        if any(kw in all_text for kw in ['错误', '无效', '失败', '问题', 'error', 'invalid']):
            log("  -> 匹配: 错误 + 忽略按钮")
            if click_button_by_text(dialog_info, ['忽略', '忽略(&I)']):
                return True
    
    # 规则7: 通用 - 如果有 OK/确定 按钮且是 CorelDRAW 弹窗
    if 'CorelDRAW' in title:
        if click_button_by_text(dialog_info, ['OK', '确定', '是', '是(&Y)', 'Yes']):
            log("  -> 匹配: CorelDRAW 通用弹窗")
            return True
    
    log("  -> 未匹配任何规则")
    return False


def find_coreldraw_dialogs():
    """查找 CorelDRAW 对话框"""
    dialogs = []
    
    def callback(hwnd, lparam):
        if user32.IsWindowVisible(hwnd):
            title = get_window_text(hwnd)
            class_name = get_class_name(hwnd)
            # #32770 是标准对话框类
            if class_name == '#32770' and ('CorelDRAW' in title or 'Corel' in title):
                dialogs.append(hwnd)
        return True
    
    user32.EnumWindows(EnumWindowsProc(callback), 0)
    return dialogs


def main():
    print("=" * 60)
    print("CorelDRAW 弹窗自动处理工具 v5.0")
    print("=" * 60)
    print()
    
    # 检查 DLL
    script_dir = os.path.dirname(os.path.abspath(__file__))
    dll_path = os.path.join(script_dir, "gdi_hook.dll")
    
    # 也检查 PyInstaller 打包后的路径
    if hasattr(sys, '_MEIPASS'):
        dll_path_temp = os.path.join(sys._MEIPASS, "gdi_hook.dll")
        if os.path.exists(dll_path_temp):
            dll_path = dll_path_temp
    
    use_hook = os.path.exists(dll_path)
    
    if use_hook:
        log(f"Hook DLL: {dll_path}")
        shared_mem = SharedMemory()
        shared_mem.create()
        injector = DLLInjector(dll_path)
    else:
        log("⚠️ 未找到 gdi_hook.dll，将只使用标准 API")
        shared_mem = None
        injector = None
    
    log("程序启动")
    log("正在监控 CorelDRAW 弹窗...")
    log("按 Ctrl+C 退出")
    print("-" * 60)
    
    handled_count = 0
    handled_hwnds = set()
    
    try:
        while True:
            # 注入 DLL
            if injector:
                injector.inject_coreldraw()
            
            # 查找对话框
            dialogs = find_coreldraw_dialogs()
            
            for hwnd in dialogs:
                if hwnd in handled_hwnds:
                    continue
                
                # 等待一下让内容稳定
                time.sleep(0.3)
                
                # 获取对话框信息
                dialog_info = get_dialog_info(hwnd)
                
                # 获取 Hook 文本
                hook_texts = []
                if shared_mem:
                    hook_texts = shared_mem.read_texts()
                
                if handle_popup(hwnd, dialog_info, hook_texts):
                    handled_count += 1
                    handled_hwnds.add(hwnd)
                    log(f"✅ 已处理 {handled_count} 个弹窗")
                
                # 清空 Hook 缓存
                if shared_mem:
                    shared_mem.clear()
            
            # 清理已关闭的窗口
            handled_hwnds = {h for h in handled_hwnds if user32.IsWindow(h)}
            
            time.sleep(0.5)
            
    except KeyboardInterrupt:
        print()
        log(f"程序退出，共处理 {handled_count} 个弹窗")
    finally:
        if shared_mem:
            shared_mem.close()


if __name__ == "__main__":
    if sys.platform != 'win32':
        print("❌ 此程序只能在 Windows 上运行！")
        sys.exit(1)
    
    if ctypes.windll.shell32.IsUserAnAdmin() == 0:
        log("⚠️ 建议以管理员权限运行")
    
    main()

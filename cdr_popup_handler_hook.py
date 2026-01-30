#!/usr/bin/env python3
"""
CorelDRAW 弹窗自动处理工具 v4.0 (Hook 版本)
使用 DLL 注入拦截 GDI 文本绘制函数，获取自绘控件的文本
"""

import time
import sys
import os
import ctypes
from ctypes import wintypes
from datetime import datetime
import mmap
import struct

# Windows API
user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32
psapi = ctypes.windll.psapi

# 常量
BM_CLICK = 0x00F5
PROCESS_ALL_ACCESS = 0x1F0FFF
MEM_COMMIT = 0x1000
MEM_RESERVE = 0x2000
PAGE_READWRITE = 0x04
MEM_RELEASE = 0x8000

# 共享内存常量
SHARED_MEM_NAME = "CDRPopupHandlerSharedMem"
MAX_TEXT_LENGTH = 4096
MAX_TEXT_COUNT = 100

# 回调函数类型
EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
EnumChildProc = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)


def log(msg):
    """打印带时间戳的日志"""
    timestamp = datetime.now().strftime("%H:%M:%S")
    print(f"[{timestamp}] {msg}")


class SharedMemory:
    """共享内存管理类"""
    
    def __init__(self):
        self.handle = None
        self.size = 4 + 8 + (MAX_TEXT_COUNT * MAX_TEXT_LENGTH * 2) + 4  # textCount + hwnd + texts + ready
        
    def create(self):
        """创建或打开共享内存"""
        self.handle = kernel32.CreateFileMappingW(
            ctypes.c_void_p(-1),  # INVALID_HANDLE_VALUE
            None,
            PAGE_READWRITE,
            0,
            self.size,
            SHARED_MEM_NAME
        )
        if not self.handle:
            log(f"创建共享内存失败: {kernel32.GetLastError()}")
            return False
        return True
    
    def read_texts(self):
        """读取共享内存中的文本"""
        if not self.handle:
            return []
        
        pBuf = kernel32.MapViewOfFile(self.handle, 0xF001F, 0, 0, self.size)
        if not pBuf:
            return []
        
        try:
            # 读取 textCount
            count_buf = (ctypes.c_char * 4)()
            ctypes.memmove(count_buf, pBuf, 4)
            text_count = struct.unpack('I', bytes(count_buf))[0]
            
            if text_count > MAX_TEXT_COUNT:
                text_count = MAX_TEXT_COUNT
            
            texts = []
            offset = 4 + 8  # Skip textCount and hwnd
            
            for i in range(text_count):
                text_buf = (ctypes.c_wchar * MAX_TEXT_LENGTH)()
                ctypes.memmove(text_buf, pBuf + offset, MAX_TEXT_LENGTH * 2)
                text = text_buf.value
                if text:
                    texts.append(text)
                offset += MAX_TEXT_LENGTH * 2
            
            return texts
        finally:
            kernel32.UnmapViewOfFile(pBuf)
    
    def clear(self):
        """清空共享内存中的文本"""
        if not self.handle:
            return
        
        pBuf = kernel32.MapViewOfFile(self.handle, 0xF001F, 0, 0, self.size)
        if pBuf:
            # 设置 textCount = 0
            zero = (ctypes.c_char * 4)()
            ctypes.memmove(pBuf, zero, 4)
            kernel32.UnmapViewOfFile(pBuf)
    
    def close(self):
        """关闭共享内存"""
        if self.handle:
            kernel32.CloseHandle(self.handle)
            self.handle = None


class DLLInjector:
    """DLL 注入器"""
    
    def __init__(self, dll_path):
        self.dll_path = os.path.abspath(dll_path)
        self.injected_pids = set()
    
    def inject(self, pid):
        """注入 DLL 到目标进程"""
        if pid in self.injected_pids:
            return True
        
        # 打开目标进程
        hProcess = kernel32.OpenProcess(PROCESS_ALL_ACCESS, False, pid)
        if not hProcess:
            log(f"无法打开进程 {pid}: {kernel32.GetLastError()}")
            return False
        
        try:
            # 在目标进程分配内存
            dll_path_bytes = (self.dll_path + '\0').encode('utf-16-le')
            dll_path_len = len(dll_path_bytes)
            
            pRemoteMem = kernel32.VirtualAllocEx(
                hProcess, None, dll_path_len, MEM_COMMIT | MEM_RESERVE, PAGE_READWRITE
            )
            if not pRemoteMem:
                log(f"VirtualAllocEx 失败: {kernel32.GetLastError()}")
                return False
            
            # 写入 DLL 路径
            written = ctypes.c_size_t()
            if not kernel32.WriteProcessMemory(hProcess, pRemoteMem, dll_path_bytes, dll_path_len, ctypes.byref(written)):
                log(f"WriteProcessMemory 失败: {kernel32.GetLastError()}")
                kernel32.VirtualFreeEx(hProcess, pRemoteMem, 0, MEM_RELEASE)
                return False
            
            # 获取 LoadLibraryW 地址
            hKernel32 = kernel32.GetModuleHandleW("kernel32.dll")
            pLoadLibraryW = kernel32.GetProcAddress(hKernel32, b"LoadLibraryW")
            
            # 创建远程线程加载 DLL
            hThread = kernel32.CreateRemoteThread(
                hProcess, None, 0, pLoadLibraryW, pRemoteMem, 0, None
            )
            if not hThread:
                log(f"CreateRemoteThread 失败: {kernel32.GetLastError()}")
                kernel32.VirtualFreeEx(hProcess, pRemoteMem, 0, MEM_RELEASE)
                return False
            
            # 等待线程完成
            kernel32.WaitForSingleObject(hThread, 5000)
            kernel32.CloseHandle(hThread)
            
            # 清理
            kernel32.VirtualFreeEx(hProcess, pRemoteMem, 0, MEM_RELEASE)
            
            self.injected_pids.add(pid)
            log(f"✅ DLL 注入成功: PID {pid}")
            return True
            
        finally:
            kernel32.CloseHandle(hProcess)
    
    def inject_coreldraw(self):
        """注入到所有 CorelDRAW 进程"""
        # 查找 CorelDRAW 进程
        def callback(hwnd, lparam):
            if user32.IsWindowVisible(hwnd):
                title = get_window_text(hwnd)
                if 'CorelDRAW' in title:
                    pid = ctypes.c_ulong()
                    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
                    if pid.value and pid.value not in self.injected_pids:
                        log(f"发现 CorelDRAW 进程: PID {pid.value}")
                        self.inject(pid.value)
            return True
        
        user32.EnumWindows(EnumWindowsProc(callback), 0)


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


def click_button_by_text(parent_hwnd, button_texts):
    """根据按钮文本点击按钮"""
    if isinstance(button_texts, str):
        button_texts = [button_texts]
    
    children = find_child_windows(parent_hwnd)
    for child in children:
        text = get_window_text(child)
        class_name = get_class_name(child)
        
        if 'Button' in class_name:
            for btn_text in button_texts:
                if btn_text.lower() in text.lower() or btn_text in text:
                    log(f"  >>> 点击按钮: '{text}'")
                    user32.SendMessageW(child, BM_CLICK, 0, 0)
                    return True
    return False


def handle_popup_with_hook_data(hwnd, title, hook_texts):
    """根据 Hook 获取的文本处理弹窗"""
    content = ' '.join(hook_texts)
    log(f"弹窗标题: '{title}'")
    log(f"Hook 捕获的文本: {content[:500]}...")
    
    # 规则匹配
    # 1. 无效的轮廓 ID
    if '无效' in content and '轮廓' in content:
        log("  -> 匹配规则: 无效的轮廓 ID")
        if click_button_by_text(hwnd, ['忽略', '忽略(&I)', 'Ignore']):
            return True
    
    # 2. 无效标头
    if '无法打开' in content or '无效标头' in content:
        log("  -> 匹配规则: 无效标头")
        if click_button_by_text(hwnd, ['OK', '确定']):
            return True
    
    # 3. 文件损坏
    if '损坏' in content:
        log("  -> 匹配规则: 文件损坏")
        if click_button_by_text(hwnd, ['OK', '确定']):
            return True
    
    # 4. PS/PRN 导入
    if 'PS/PRN' in content:
        log("  -> 匹配规则: PS/PRN 导入")
        # 先选曲线
        click_button_by_text(hwnd, ['曲线', '曲线(&C)'])
        time.sleep(0.2)
        if click_button_by_text(hwnd, ['OK', '确定']):
            return True
    
    # 5. 按钮组合匹配
    if '关于' in content and '重试' in content and '忽略' in content:
        log("  -> 匹配规则: 关于/重试/忽略 按钮组合")
        if click_button_by_text(hwnd, ['忽略', '忽略(&I)']):
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
            if class_name == '#32770' and 'CorelDRAW' in title:
                dialogs.append((hwnd, title))
        return True
    
    user32.EnumWindows(EnumWindowsProc(callback), 0)
    return dialogs


def main():
    print("=" * 60)
    print("CorelDRAW 弹窗自动处理工具 v4.0 (Hook 版本)")
    print("=" * 60)
    print()
    
    # 检查 DLL 是否存在
    dll_path = os.path.join(os.path.dirname(__file__), "gdi_hook.dll")
    if not os.path.exists(dll_path):
        log(f"❌ 找不到 Hook DLL: {dll_path}")
        log("请先编译 gdi_hook.cpp")
        return
    
    log("初始化共享内存...")
    shared_mem = SharedMemory()
    if not shared_mem.create():
        log("❌ 共享内存初始化失败")
        return
    
    log("初始化 DLL 注入器...")
    injector = DLLInjector(dll_path)
    
    log("程序启动完成")
    log("正在监控 CorelDRAW 弹窗...")
    log("按 Ctrl+C 退出")
    print("-" * 60)
    
    handled_count = 0
    handled_hwnds = set()
    
    try:
        while True:
            # 尝试注入到 CorelDRAW
            injector.inject_coreldraw()
            
            # 查找对话框
            dialogs = find_coreldraw_dialogs()
            
            for hwnd, title in dialogs:
                if hwnd in handled_hwnds:
                    continue
                
                # 等待一下让 Hook 捕获文本
                time.sleep(0.3)
                
                # 读取 Hook 捕获的文本
                hook_texts = shared_mem.read_texts()
                
                log(f"检测到对话框: {title}")
                log(f"  Hook 捕获了 {len(hook_texts)} 个文本")
                
                if handle_popup_with_hook_data(hwnd, title, hook_texts):
                    handled_count += 1
                    handled_hwnds.add(hwnd)
                    log(f"✅ 已处理 {handled_count} 个弹窗")
                
                # 清空捕获的文本
                shared_mem.clear()
            
            # 清理已关闭的窗口
            handled_hwnds = {h for h in handled_hwnds if user32.IsWindow(h)}
            
            time.sleep(0.5)
            
    except KeyboardInterrupt:
        print()
        log(f"程序退出，共处理 {handled_count} 个弹窗")
    finally:
        shared_mem.close()


if __name__ == "__main__":
    if sys.platform != 'win32':
        print("❌ 此程序只能在 Windows 上运行！")
        sys.exit(1)
    
    # 需要管理员权限
    if ctypes.windll.shell32.IsUserAnAdmin() == 0:
        log("⚠️ 建议以管理员权限运行以确保 DLL 注入成功")
    
    main()

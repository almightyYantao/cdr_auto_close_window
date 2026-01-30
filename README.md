# CDR 弹窗自动处理工具

自动检测并处理 CorelDRAW 打开文件时的各种错误弹窗。

## 版本说明

### 标准版 (CDR-Popup-Handler.exe)
- 使用 Windows API 获取控件文本
- 适用于使用标准 Windows 控件的弹窗

### Hook 版 (CDR-Popup-Handler-Hook.exe)
- 使用 DLL 注入 + GDI Hook 技术
- 拦截 TextOut/DrawText 等函数
- 能够获取自绘控件的文本（GDI 绘制的文字）
- **需要管理员权限运行**
- 需要 gdi_hook.dll 在同一目录

## 支持的弹窗类型

| 错误类型 | 处理方式 |
|---------|----------|
| 无效的轮廓 ID | 点击 **忽略** |
| 无效标头/无法打开文件 | 点击 **OK** |
| 文件被损坏 | 点击 **OK** |
| 导入 PS/PRN | 选择 **曲线** → 点击 **OK** |

## 下载

前往 [Actions](../../actions) 页面，点击最新的构建，下载 artifact。

## 使用方法

### 标准版
1. 下载 `CDR-Popup-Handler-Standard`
2. 解压后双击运行 `CDR-Popup-Handler.exe`

### Hook 版
1. 下载 `CDR-Popup-Handler-Hook`
2. 确保 `gdi_hook.dll` 和 `CDR-Popup-Handler-Hook.exe` 在同一目录
3. **右键以管理员身份运行** `CDR-Popup-Handler-Hook.exe`

## 技术原理

### Hook 版工作流程
1. 创建共享内存
2. 注入 gdi_hook.dll 到 CorelDRAW 进程
3. DLL Hook 住 TextOutW/DrawTextW 等 GDI 函数
4. 捕获所有文本绘制内容写入共享内存
5. Python 程序读取共享内存，匹配规则，点击按钮

## 自行编译

### 编译 DLL (需要 Visual Studio)
```cmd
cd hook
cl /LD /EHsc gdi_hook.cpp user32.lib gdi32.lib /Fe:gdi_hook.dll
```

### 编译 EXE
```cmd
pip install pyinstaller
pyinstaller --onefile --console --name "CDR-Popup-Handler-Hook" --add-data "gdi_hook.dll;." cdr_popup_handler_hook.py
```

## License

MIT

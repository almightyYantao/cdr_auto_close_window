# CDR 弹窗自动处理工具

自动检测并处理 CorelDRAW 打开文件时的各种错误弹窗。

## 支持的弹窗类型

| 错误类型 | 处理方式 |
|---------|----------|
| 无效的轮廓 ID | 点击 **忽略** |
| 无效标头/无法打开文件 | 点击 **OK** |
| 文件被损坏 | 点击 **OK** |
| 导入 PS/PRN | 选择 **曲线** → 点击 **OK** |

## 下载

前往 [Actions](../../actions) 页面，点击最新的构建，下载 `CDR弹窗处理工具` artifact。

## 使用方法

1. 下载并解压 `CDR弹窗处理工具.exe`
2. 双击运行程序
3. 保持程序在后台运行
4. 正常使用 CorelDRAW 打开文件
5. 弹窗会被自动处理
6. 按 `Ctrl+C` 退出程序

## 从源码运行

```bash
# 需要 Python 3.8+ 和 Windows
python cdr_popup_handler.py
```

## 自行打包

```bash
pip install pyinstaller
pyinstaller --onefile --console --name "CDR弹窗处理工具" cdr_popup_handler.py
```

## License

MIT

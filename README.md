# PyKeyPresser

一个基于dm.dll（版本: 3.1233） COM接口的Python自动化库，提供鼠标键盘模拟、图像识别、窗口操作等功能。

**语言:** [中文](README.md) | [English](README_EN.md)

## 功能特点

- 🖱️ **鼠标操作**: 移动、点击、拖拽等完整的鼠标控制
- ⌨️ **键盘操作**: 按键按下、释放、字符串输入等键盘控制
- 🎨 **颜色识别**: 获取屏幕颜色、比较颜色、查找指定颜色
- 📷 **图像处理**: 截图、图像查找、OCR文字识别
- 🪟 **窗口操作**: 获取窗口信息、移动窗口、设置窗口状态
- 🔧 **内存操作**: 读取和写入进程内存
- 🎵 **多媒体**: 播放声音、蜂鸣器控制

## 安装要求

- Python 3.6+ **32位版本** （64位Python不支持dm.dll）
- pywin32库
- dm.dll文件

### 安装依赖

```bash
pip install pywin32
```

## 快速开始

### 使用包装类（推荐）

```python
from PyKeyPresser import PyKeyPresser

# 创建PyKeyPresser对象
kp = PyKeyPresser()

# 获取版本
print(kp.Ver())

# 鼠标操作
kp.MoveTo(100, 100)
kp.LeftClick()

# 键盘操作
kp.KeyPress(65)  # A键
```

## 主要功能

### 鼠标操作

```python
kp.MoveTo(x, y)              # 移动鼠标到指定位置
kp.LeftClick()               # 左键单击
kp.RightClick()              # 右键单击
kp.LeftDoubleClick()         # 左键双击
kp.LeftDown()/LeftUp()       # 左键按下/释放
kp.MoveR(rx, ry)             # 相对移动鼠标
kp.WheelUp()/WheelDown()     # 鼠标滚轮操作
```

### 键盘操作

```python
kp.KeyPress(vk_code)         # 按键
kp.KeyDown(vk_code)          # 按键按下
kp.KeyUp(vk_code)            # 按键释放
kp.SendString(hwnd, text)    # 发送字符串
```

### 颜色操作

```python
kp.GetColor(x, y)            # 获取指定位置颜色
kp.CmpColor(x, y, color, sim) # 比较颜色
kp.FindColor(x1, y1, x2, y2, color, sim, dir) # 查找颜色
```

### 图像识别

```python
kp.Ocr(x1, y1, x2, y2, color, sim)                    # OCR文字识别
kp.FindStr(x1, y1, x2, y2, str, color, sim)           # 查找字符串
kp.Capture(x1, y1, x2, y2, file)                       # 截图
kp.FindPic(x1, y1, x2, y2, pic_name, delta_color, sim) # 查找图片
```

### 窗口操作

```python
kp.GetForegroundWindow()     # 获取前台窗口
kp.GetWindowTitle(hwnd)     # 获取窗口标题
kp.GetWindowRect(hwnd)       # 获取窗口矩形
kp.EnumWindow(parent, title, class_name, filter) # 枚举窗口
kp.SetWindowState(hwnd, flag) # 设置窗口状态
```

## 常用虚拟键码

| 键名 | 键码 | 说明 |
|------|------|------|
| MOUSE_LEFT | 0x01 | 左鼠标按钮 |
| MOUSE_RIGHT | 0x02 | 右鼠标按钮 |
| BACK | 0x08 | Backspace |
| TAB | 0x09 | Tab |
| RETURN | 0x0D | Enter |
| SHIFT | 0x10 | Shift |
| CONTROL | 0x11 | Ctrl |
| MENU | 0x12 | Alt |
| ESCAPE | 0x1B | Esc |
| SPACE | 0x20 | 空格 |
| F1-F12 | 0x70-0x7B | 功能键 |

## 示例代码

### 自动点击示例

```python
from PyKeyPresser import PyKeyPresser
import time

kp = PyKeyPresser()

# 移动到指定位置并点击
kp.MoveTo(500, 300)
time.sleep(0.5)
kp.LeftClick()
```

### 颜色查找示例

```python
from PyKeyPresser import PyKeyPresser

kp = PyKeyPresser()

# 查找红色
result, x, y = kp.FindColor(0, 0, 1000, 1000, "ff0000", 0.9, 0)
if result == 1:
    print(f"找到红色，位置: ({x}, {y})")
    kp.MoveTo(x, y)
    kp.LeftClick()
else:
    print("未找到红色")
```

### OCR文字识别示例

```python
from PyKeyPresser import PyKeyPresser

kp = PyKeyPresser()

# 在指定区域进行OCR识别
ocr_result = kp.Ocr(0, 0, 500, 200, "000000-000000", 0.8)
if ocr_result:
    print(f"识别结果: {ocr_result}")
```

## 运行示例

项目包含完整的示例代码，可以直接运行：

```bash
python kp_example.py
```

## 项目结构

```
PyKeyPresser/
├── PyKeyPresser.py      # PyKeyPresser类，kp COM接口的Python包装类
├── kp_example.py      # 使用示例和测试代码
├── CKeyPresser.h      # C++头文件，定义COM接口
├── dm.dll            # dm COM组件
├── README.md         # 本文件
└── 大漠接口说明.chm 
```

## 注意事项

1. **权限问题**: 某些操作需要管理员权限
2. **坐标系统**: 屏幕坐标原点在左上角
3. **颜色格式**: 使用十六进制格式，如"ffffff"表示白色
4. **错误处理**: 建议使用try-catch处理COM异常
5. **免注册调用**: 支持免注册调用dm.dll，详见PyKeyPresser类实现

## 故障排除

1. **COM对象创建失败**: 检查dm.dll是否正确注册
2. **权限不足**: 以管理员身份运行
3. **找不到窗口**: 检查窗口标题或类名是否正确
4. **颜色识别失败**: 调整相似度参数

## 扩展功能

基于PyKeyPresser类，你可以轻松扩展更多功能：

1. **图像匹配和识别**
2. **自动化脚本录制**
3. **游戏辅助工具**
4. **UI自动化测试**

## 许可证

本项目基于现有的dm.dll COM接口开发，请遵循相关许可证要求。

## 贡献

欢迎提交问题报告和功能请求。如果您想贡献代码，请先创建issue讨论您的想法。

## 相关资源

- [PyWin32帮助文档](.venv/Lib/site-packages/PyWin32.chm) - 包含完整的PyWin32 API参考
- [pywin32文档](https://docs.python.org/3/library/win32com.html)
- [comtypes文档](https://comtypes.readthedocs.io/)
- 大漠接口说明.chm 
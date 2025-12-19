# PyKeyPresser

A Python automation library based on dm.dll (version: 3.1233) COM interface, providing mouse/keyboard simulation, image recognition, window operations, and more.

**Language:** [English](README_EN.md) | [中文](README.md)

## Features

- 🖱️ **Mouse Operations**: Complete mouse control including movement, clicking, and dragging
- ⌨️ **Keyboard Operations**: Keyboard control including key press, release, and string input
- 🎨 **Color Recognition**: Get screen colors, compare colors, find specified colors
- 📷 **Image Processing**: Screenshots, image search, OCR text recognition
- 🪟 **Window Operations**: Get window information, move windows, set window states
- 🔧 **Memory Operations**: Read and write process memory
- 🎵 **Multimedia**: Play sounds, control beeper

## Requirements

- Python 3.6+ **32-bit version** (64-bit Python doesn't support dm.dll)
- pywin32 library
- dm.dll file

### Install Dependencies

```bash
pip install pywin32
```

## Quick Start

### Using the Wrapper Class (Recommended)

```python
from PyKeyPresser import PyKeyPresser

# Create PyKeyPresser object
kp = PyKeyPresser()

# Get version
print(kp.Ver())

# Mouse operations
kp.MoveTo(100, 100)
kp.LeftClick()

# Keyboard operations
kp.KeyPress(65)  # A key
```

## Main Features

### Mouse Operations

```python
kp.MoveTo(x, y)              # Move mouse to specified position
kp.LeftClick()               # Left click
kp.RightClick()              # Right click
kp.LeftDoubleClick()         # Left double click
kp.LeftDown()/LeftUp()       # Left button press/release
kp.MoveR(rx, ry)             # Relative mouse movement
kp.WheelUp()/WheelDown()     # Mouse wheel operations
```

### Keyboard Operations

```python
kp.KeyPress(vk_code)         # Key press
kp.KeyDown(vk_code)          # Key down
kp.KeyUp(vk_code)            # Key up
kp.SendString(hwnd, text)    # Send string
```

### Color Operations

```python
kp.GetColor(x, y)            # Get color at specified position
kp.CmpColor(x, y, color, sim) # Compare colors
kp.FindColor(x1, y1, x2, y2, color, sim, dir) # Find color
```

### Image Recognition

```python
kp.Ocr(x1, y1, x2, y2, color, sim)                    # OCR text recognition
kp.FindStr(x1, y1, x2, y2, str, color, sim)           # Find string
kp.Capture(x1, y1, x2, y2, file)                       # Screenshot
kp.FindPic(x1, y1, x2, y2, pic_name, delta_color, sim) # Find picture
```

### Window Operations

```python
kp.GetForegroundWindow()     # Get foreground window
kp.GetWindowTitle(hwnd)     # Get window title
kp.GetWindowRect(hwnd)       # Get window rectangle
kp.EnumWindow(parent, title, class_name, filter) # Enumerate windows
kp.SetWindowState(hwnd, flag) # Set window state
```

## Common Virtual Key Codes

| Key Name | Key Code | Description |
|----------|----------|-------------|
| MOUSE_LEFT | 0x01 | Left mouse button |
| MOUSE_RIGHT | 0x02 | Right mouse button |
| BACK | 0x08 | Backspace |
| TAB | 0x09 | Tab |
| RETURN | 0x0D | Enter |
| SHIFT | 0x10 | Shift |
| CONTROL | 0x11 | Ctrl |
| MENU | 0x12 | Alt |
| ESCAPE | 0x1B | Esc |
| SPACE | 0x20 | Space |
| F1-F12 | 0x70-0x7B | Function keys |

## Example Code

### Auto Click Example

```python
from PyKeyPresser import PyKeyPresser
import time

kp = PyKeyPresser()

# Move to specified position and click
kp.MoveTo(500, 300)
time.sleep(0.5)
kp.LeftClick()
```

### Color Search Example

```python
from PyKeyPresser import PyKeyPresser

kp = PyKeyPresser()

# Find red color
result, x, y = kp.FindColor(0, 0, 1000, 1000, "ff0000", 0.9, 0)
if result == 1:
    print(f"Found red color at: ({x}, {y})")
    kp.MoveTo(x, y)
    kp.LeftClick()
else:
    print("Red color not found")
```

### OCR Text Recognition Example

```python
from PyKeyPresser import PyKeyPresser

kp = PyKeyPresser()

# OCR recognition in specified area
ocr_result = kp.Ocr(0, 0, 500, 200, "000000-000000", 0.8)
if ocr_result:
    print(f"Recognition result: {ocr_result}")
```

## Running Examples

The project includes complete example code that can be run directly:

```bash
python example.py
```

## Project Structure

```
PyKeyPresser/
├── PyKeyPresser.py      # PyKeyPresser class, Python wrapper for kp COM interface
├── example.py      # Usage examples and test code
├── CKeyPresser.h      # C++ header file defining COM interface
├── dm.dll            # dm COM component
├── README.md         # Chinese documentation
└── 大漠接口说明.chm 
```

## Notes

1. **Permission Issues**: Some operations require administrator privileges
2. **Coordinate System**: Screen coordinate origin is at the top-left corner
3. **Color Format**: Uses hexadecimal format, e.g., "ffffff" for white
4. **Error Handling**: It's recommended to use try-catch for COM exceptions
5. **Registration-Free Call**: Supports registration-free calling of dm.dll, see PyKeyPresser class implementation for details

## Troubleshooting

1. **COM Object Creation Failed**: Check if dm.dll is correctly registered
2. **Insufficient Permissions**: Run as administrator
3. **Window Not Found**: Check if window title or class name is correct
4. **Color Recognition Failed**: Adjust similarity parameters

## Extended Features

Based on the PyKeyPresser class, you can easily extend more functionality:

1. **Image matching and recognition**
2. **Automated script recording**
3. **Game assistance tools**
4. **UI automation testing**

## License

This project is developed based on the existing dm.dll COM interface, please follow relevant license requirements.

## Contributing

Issue reports and feature requests are welcome. If you want to contribute code, please create an issue first to discuss your ideas.

## Related Resources

- [PyWin32 Help Documentation](.venv/Lib/site-packages/PyWin32.chm) - Contains complete PyWin32 API reference
- [pywin32 documentation](https://docs.python.org/3/library/win32com.html)
- [comtypes documentation](https://comtypes.readthedocs.io/)
- 大漠接口说明.chm (Chinese documentation)
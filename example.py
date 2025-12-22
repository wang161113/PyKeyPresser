"""
dm COM接口使用示例
演示各种常用功能
"""
from PyKeyPresser import PyKeyPresser
import time

def main():
    # 创建kp对象
    kp = PyKeyPresser()
    
    print(f"版本: {kp.Ver()}")
    
    # 设置工作目录
    kp.SetPath(".")
    
    # 1. 鼠标操作示例
    print("\n=== 鼠标操作示例 ===")
    current_x, current_y = 500, 300
    
    # 移动鼠标到指定位置
    kp.MoveTo(current_x, current_y)
    print(f"鼠标移动到: ({current_x}, {current_y})")
    time.sleep(1)
    
    # 左键单击
    kp.LeftClick()
    print("左键单击")
    time.sleep(0.5)
    
    # 左键双击
    kp.LeftDoubleClick()
    print("左键双击")
    time.sleep(0.5)
    
    # 右键单击
    kp.RightClick()
    print("右键单击")
    time.sleep(0.5)
    
    # 2. 键盘操作示例
    print("\n=== 键盘操作示例 ===")
    
    # 输入文字 "Hello"
    hello_keys = [72, 101, 108, 108, 111]  # H e l l o
    for key_code in hello_keys:
        kp.KeyPress(key_code)
        time.sleep(0.1)
    print("输入: Hello")
    
    # 按下回车键
    kp.KeyPress(13)  # Enter键
    print("按下回车键")
    time.sleep(1)
    
    # 3. 颜色操作示例
    print("\n=== 颜色操作示例 ===")
    
    # 获取指定位置的颜色
    test_x, test_y = 100, 100
    color = kp.GetColor(test_x, test_y)
    print(f"坐标({test_x}, {test_y})的颜色: {color}")
    
    # 比较颜色
    target_color = "ffffff"  # 白色
    similarity = 0.9
    is_match = kp.CmpColor(test_x, test_y, target_color, similarity)
    print(f"颜色匹配结果: {is_match}")
    
    # 查找颜色
    find_result, x, y = kp.FindColor(0, 0, 1000, 1000, "ff0000", 0.9, 0)
    if find_result == 1:
        print(f"找到红色，位置: ({x}, {y})")
    else:
        print("未找到指定颜色")
    
    # 4. 窗口操作示例
    print("\n=== 窗口操作示例 ===")
    
    # 获取前台窗口
    foreground_hwnd = kp.GetForegroundWindow()
    print(f"前台窗口句柄: {foreground_hwnd}")
    
    if foreground_hwnd:
        # 获取窗口标题
        title = kp.GetWindowTitle(foreground_hwnd)
        print(f"窗口标题: {title}")
        
        # 获取窗口类名
        class_name = kp.GetWindowClass(foreground_hwnd)
        print(f"窗口类名: {class_name}")
        
        # 获取窗口矩形
        rect_result, x1, y1, x2, y2 = kp.GetWindowRect(foreground_hwnd)
        if rect_result:
            print(f"窗口位置: ({x1}, {y1}) - ({x2}, {y2})")
            width = x2 - x1
            height = y2 - y1
            print(f"窗口大小: {width} x {height}")
    
    # 5. 截图示例
    print("\n=== 截图示例 ===")
    screenshot_path = "test_screenshot.bmp"
    capture_result = kp.Capture(0, 0, 800, 600, screenshot_path)
    if capture_result == 1:
        print(f"截图成功，保存到: {screenshot_path}")
    else:
        print("截图失败")
    
    # 6. OCR识别示例
    print("\n=== OCR识别示例 ===")
    try:
        # 在指定区域进行OCR识别
        ocr_result = kp.Ocr(0, 0, 500, 200, "000000-000000", 0.8)
        if ocr_result:
            print(f"OCR识别结果: {ocr_result}")
        else:
            print("OCR识别失败或未找到文字")
    except Exception as e:
        print(f"OCR识别出错: {e}")
    
    # 7. 文字查找示例
    print("\n=== 文字查找示例 ===")
    try:
        # 查找指定文字
        find_result, x, y = kp.FindStr(0, 0, 1000, 1000, "确定", "000000-000000", 0.8)
        if find_result == 1:
            print(f"找到文字'确定'，位置: ({x}, {y})")
        else:
            print("未找到指定文字")
    except Exception as e:
        print(f"文字查找出错: {e}")
    
    # 8. 声音提示示例
    print("\n=== 声音提示示例 ===")
    # 发出蜂鸣声
    kp.Beep(1000, 500)  # 1000Hz, 500ms
    print("发出蜂鸣声")
    
    print("\n=== 所有测试完成 ===")

def advanced_example():
    """高级功能示例"""
    kp = PyKeyPresser()
    
    print("\n=== 高级功能示例 ===")
    
    # 1. 枚举窗口
    print("枚举所有窗口:")
    windows = kp.EnumWindow(0, "", "", 0)
    if windows:
        window_list = windows.split(",")
        for i, hwnd in enumerate(window_list): 
            title = kp.GetWindowTitle(int(hwnd))
            if title:
                print(f"  {i+1}. 句柄: {hwnd}, 标题: {title}")
    
    # 2. 连续查找多个颜色点
    print("\n颜色查找模式:")
    colors = ["ff0000", "00ff00", "0000ff"]  # 红、绿、蓝
    for i, color in enumerate(colors):
        result, x, y = kp.FindColor(0, 0, 1000, 1000, color, 0.9, 0)
        if result == 1:
            print(f"  找到颜色 {color}: ({x}, {y})")
        else:
            print(f"  未找到颜色 {color}")
    
    # 3. 相对移动鼠标
    print("\n相对移动鼠标:")
    kp.MoveTo(400, 300)
    print("移动到中心位置")
    time.sleep(1)
    
    # 画一个正方形
    for i in range(4):
        kp.MoveR(50, 0)  # 向右移动50像素
        time.sleep(0.2)
        kp.MoveR(0, 50)  # 向下移动50像素
        time.sleep(0.2)
        kp.MoveR(-50, 0)  # 向左移动50像素
        time.sleep(0.2)
        kp.MoveR(0, -50)  # 向上移动50像素
        time.sleep(0.2)
    print("画了一个正方形")

if __name__ == "__main__":
    try:
        main()
        advanced_example()
    except Exception as e:
        print(f"程序运行出错: {e}")
        import traceback
        traceback.print_exc()
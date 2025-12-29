"""
DM COM接口Python包装类
基于CKeyPresser.h自动生成，函数顺序与CKeyPresser.h保持一致
支持多种创建方式，包括免注册调用
"""
import win32com.client
import time
import os
import sys

class PyKeyPresser:
    def __init__(self, dll_path=None, use_registration_fallback=True):
        """
        初始化DM COM对象
        Args:
            dll_path: dm.dll文件路径，如果为None则自动查找
            use_registration_fallback: 如果注册失败，是否尝试免注册方法
        """
        self.dm = None
        self.creation_method = None
        
        # 尝试多种创建方法
        creation_methods = [
            self._create_with_registration,
            self._create_with_path,
            self._create_with_alternative_progids,
            self._create_with_dll_direct
        ]
        
        for method in creation_methods:
            try:
                self.dm = method(dll_path)
                if self.dm:
                    self.creation_method = method.__name__
                    print(f"成功使用 {method.__name__} 创建DM对象")
                    break
            except Exception as e:
                print(f"{method.__name__} 失败: {e}")
                continue
        
        if not self.dm:
            if use_registration_fallback:
                # 尝试免注册方法
                try:
                    from dm_unregistered import create_dm_from_dll
                    self.dm = create_dm_from_dll()
                    if self.dm:
                        self.creation_method = "unregistered_com"
                        print("使用免注册COM方法创建成功")
                except Exception as e:
                    print(f"免注册方法失败: {e}")
            
            if not self.dm:
                raise Exception("所有DM对象创建方法都失败了")
    
    def _create_with_registration(self, dll_path):
        """标准注册方式创建"""
        return win32com.client.Dispatch("dm.dmsoft")
    
    def _create_with_path(self, dll_path):
        """设置DLL路径后创建"""
        if dll_path and os.path.exists(dll_path):
            dll_dir = os.path.dirname(dll_path)
            if dll_dir not in os.environ['PATH']:
                os.environ['PATH'] = dll_dir + ';' + os.environ['PATH']
        return win32com.client.Dispatch("dm.dmsoft")
    
    def _create_with_alternative_progids(self, dll_path):
        """尝试不同的ProgID"""
        prog_ids = [
            "dm.dmsoft",
            "dmsoft.dmsoft", 
            "DmSoft.DmSoft",
            "dm.soft"
        ]
        
        for prog_id in prog_ids:
            try:
                return win32com.client.Dispatch(prog_id)
            except:
                continue
        return None
    
    def _create_with_dll_direct(self, dll_path):
        """直接从DLL创建"""
        if not dll_path:
            current_dir = os.path.dirname(os.path.abspath(__file__))
            dll_path = os.path.join(current_dir, "dm.dll")
        
        if os.path.exists(dll_path):
            # 添加DLL目录到PATH
            dll_dir = os.path.dirname(dll_path)
            if dll_dir not in os.environ['PATH']:
                os.environ['PATH'] = dll_dir + ';' + os.environ['PATH']
            
            # 尝试创建
            return win32com.client.Dispatch("dm.dmsoft")
        
        return None

    # 按照CKeyPresser.h顺序排列的函数
    
    def Ver(self):
        """获取版本"""
        return self.dm.Ver()
    
    def SetPath(self, path):
        """设置路径"""
        return self.dm.SetPath(path)
    
    def Ocr(self, x1, y1, x2, y2, color, sim):
        """OCR识别"""
        return self.dm.Ocr(x1, y1, x2, y2, color, sim)
    
    def FindStr(self, x1, y1, x2, y2, str_text, color, sim):
        """查找字符串"""
        x, y = 0, 0
        result = self.dm.FindStr(x1, y1, x2, y2, str_text, color, sim, x, y)
        return result, x, y
    
    def GetResultCount(self, str_data):
        """获取结果数量"""
        return self.dm.GetResultCount(str_data)
    
    def GetResultPos(self, str_data, index):
        """获取结果位置"""
        x, y = 0, 0
        result = self.dm.GetResultPos(str_data, index, x, y)
        return result, x, y
    
    def StrStr(self, s, str_text):
        """字符串查找"""
        return self.dm.StrStr(s, str_text)
    
    def SendCommand(self, cmd):
        """发送命令"""
        return self.dm.SendCommand(cmd)
    
    def UseDict(self, index):
        """使用字典"""
        return self.dm.UseDict(index)
    
    def GetBasePath(self):
        """获取基础路径"""
        return self.dm.GetBasePath()
    
    def SetDictPwd(self, pwd):
        """设置字典密码"""
        return self.dm.SetDictPwd(pwd)
    
    def OcrInFile(self, x1, y1, x2, y2, pic_name, color, sim):
        """从文件中进行OCR识别"""
        return self.dm.OcrInFile(x1, y1, x2, y2, pic_name, color, sim)
    
    def Capture(self, x1, y1, x2, y2, file_path):
        """截取屏幕"""
        return self.dm.Capture(x1, y1, x2, y2, file_path)
    
    def KeyPress(self, vk_code):
        """按键按下"""
        return self.dm.KeyPress(vk_code)
    
    def KeyDown(self, vk_code):
        """按键按下"""
        return self.dm.KeyDown(vk_code)
    
    def KeyUp(self, vk_code):
        """按键释放"""
        return self.dm.KeyUp(vk_code)
    
    def KeyDownChar(self, key_str):
        """按下字符键"""
        return self.dm.KeyDownChar(key_str)
    
    def KeyUpChar(self, key_str):
        """释放字符键"""
        return self.dm.KeyUpChar(key_str)
    
    def KeyPressChar(self, key_str):
        """按下字符键"""
        return self.dm.KeyPressChar(key_str)
    
    def KeyPressStr(self, key_str, delay):
        """按下字符串键"""
        return self.dm.KeyPressStr(key_str, delay)
    
    def KeyCombo(self, modifier_key, key_code, delay=0.1):
        """发送组合键
        Args:
            modifier_key: 修饰键的虚拟键码 (如Ctrl=17, Alt=18, Shift=16)
            key_code: 目标键的虚拟键码
            delay: 按键之间的延迟(秒)
        """
        self.KeyDown(modifier_key)
        time.sleep(delay)
        self.KeyPress(key_code)
        time.sleep(delay)
        self.KeyUp(modifier_key)
    
    def CtrlA(self):
        """Ctrl+A 全选"""
        self.KeyCombo(17, 65)  # Ctrl=17, A=65
    
    def CtrlC(self):
        """Ctrl+C 复制"""
        self.KeyCombo(17, 67)  # Ctrl=17, C=67
    
    def CtrlV(self):
        """Ctrl+V 粘贴"""
        self.KeyCombo(17, 86)  # Ctrl=17, V=86
    
    def CtrlX(self):
        """Ctrl+X 剪切"""
        self.KeyCombo(17, 88)  # Ctrl=17, X=88
    
    def CtrlZ(self):
        """Ctrl+Z 撤销"""
        self.KeyCombo(17, 90)  # Ctrl=17, Z=90
    
    def CtrlS(self):
        """Ctrl+S 保存"""
        self.KeyCombo(17, 83)  # Ctrl=17, S=83
    
    def LeftClick(self):
        """左键单击"""
        return self.dm.LeftClick()
    
    def RightClick(self):
        """右键单击"""
        return self.dm.RightClick()
    
    def MiddleClick(self):
        """中键单击"""
        return self.dm.MiddleClick()
    
    def LeftDoubleClick(self):
        """左键双击"""
        return self.dm.LeftDoubleClick()
    
    def LeftDown(self):
        """左键按下"""
        return self.dm.LeftDown()
    
    def LeftUp(self):
        """左键释放"""
        return self.dm.LeftUp()
    
    def RightDown(self):
        """右键按下"""
        return self.dm.RightDown()
    
    def RightUp(self):
        """右键释放"""
        return self.dm.RightUp()
    
    def MoveTo(self, x, y):
        """移动鼠标"""
        return self.dm.MoveTo(x, y)
    
    def MoveR(self, rx, ry):
        """相对移动鼠标"""
        return self.dm.MoveR(rx, ry)
    
    def GetColor(self, x, y):
        """获取指定坐标颜色"""
        return self.dm.GetColor(x, y)
    
    def GetColorBGR(self, x, y):
        """获取指定坐标BGR颜色"""
        return self.dm.GetColorBGR(x, y)
    
    def RGB2BGR(self, rgb_color):
        """RGB颜色转BGR颜色"""
        return self.dm.RGB2BGR(rgb_color)
    
    def BGR2RGB(self, bgr_color):
        """BGR颜色转RGB颜色"""
        return self.dm.BGR2RGB(bgr_color)
    
    def UnBindWindow(self):
        """解除窗口绑定"""
        return self.dm.UnBindWindow()
    
    def CmpColor(self, x, y, color, sim):
        """比较颜色"""
        return self.dm.CmpColor(x, y, color, sim)
    
    def ClientToScreen(self, hwnd, x, y):
        """客户区坐标转屏幕坐标"""
        result = self.dm.ClientToScreen(hwnd, x, y)
        return result, x, y
    
    def ScreenToClient(self, hwnd, x, y):
        """屏幕坐标转客户区坐标"""
        result = self.dm.ScreenToClient(hwnd, x, y)
        return result, x, y
    
    def ShowScrMsg(self, x1, y1, x2, y2, msg, color):
        """显示屏幕消息"""
        return self.dm.ShowScrMsg(x1, y1, x2, y2, msg, color)
    
    def SetMinRowGap(self, row_gap):
        """设置最小行间距"""
        return self.dm.SetMinRowGap(row_gap)
    
    def SetMinColGap(self, col_gap):
        """设置最小列间距"""
        return self.dm.SetMinColGap(col_gap)
    
    def FindColor(self, x1, y1, x2, y2, color, sim, dir=0):
        """查找颜色"""
        x, y = 0, 0
        result = self.dm.FindColor(x1, y1, x2, y2, color, sim, dir, x, y)
        return result, x, y
    
    def FindColorEx(self, x1, y1, x2, y2, color, sim, dir=0):
        """查找颜色(扩展)"""
        return self.dm.FindColorEx(x1, y1, x2, y2, color, sim, dir)
    
    def SetWordLineHeight(self, line_height):
        """设置字行高度"""
        return self.dm.SetWordLineHeight(line_height)
    
    def SetWordGap(self, word_gap):
        """设置字间距"""
        return self.dm.SetWordGap(word_gap)
    
    def SetRowGapNoDict(self, row_gap):
        """设置行间距(不使用字典)"""
        return self.dm.SetRowGapNoDict(row_gap)
    
    def SetColGapNoDict(self, col_gap):
        """设置列间距(不使用字典)"""
        return self.dm.SetColGapNoDict(col_gap)
    
    def SetWordLineHeightNoDict(self, line_height):
        """设置字行高度(不使用字典)"""
        return self.dm.SetWordLineHeightNoDict(line_height)
    
    def SetWordGapNoDict(self, word_gap):
        """设置字间距(不使用字典)"""
        return self.dm.SetWordGapNoDict(word_gap)
    
    def GetWordResultCount(self, str_data):
        """获取文字识别结果数量"""
        return self.dm.GetWordResultCount(str_data)
    
    def GetWordResultPos(self, str_data, index):
        """获取文字识别结果位置"""
        x, y = 0, 0
        result = self.dm.GetWordResultPos(str_data, index, x, y)
        return result, x, y
    
    def GetWordResultStr(self, str_data, index):
        """获取文字识别结果字符串"""
        return self.dm.GetWordResultStr(str_data, index)
    
    def GetWords(self, x1, y1, x2, y2, color, sim):
        """获取识别到的所有文字"""
        return self.dm.GetWords(x1, y1, x2, y2, color, sim)
    
    def GetWordsNoDict(self, x1, y1, x2, y2, color):
        """获取识别到的所有文字(不使用字典)"""
        return self.dm.GetWordsNoDict(x1, y1, x2, y2, color)
    
    def SetShowErrorMsg(self, show):
        """设置是否显示错误消息"""
        return self.dm.SetShowErrorMsg(show)
    
    def GetClientSize(self, hwnd):
        """获取窗口客户区大小"""
        width, height = 0, 0
        result = self.dm.GetClientSize(hwnd, width, height)
        return result, width, height
    
    def MoveWindow(self, hwnd, x, y):
        """移动窗口"""
        return self.dm.MoveWindow(hwnd, x, y)
    
    def GetColorHSV(self, x, y):
        """获取指定坐标HSV颜色"""
        return self.dm.GetColorHSV(x, y)
    
    def GetAveRGB(self, x1, y1, x2, y2):
        """获取区域平均RGB颜色"""
        return self.dm.GetAveRGB(x1, y1, x2, y2)
    
    def GetAveHSV(self, x1, y1, x2, y2):
        """获取区域平均HSV颜色"""
        return self.dm.GetAveHSV(x1, y1, x2, y2)
    
    def GetForegroundWindow(self):
        """获取前台窗口"""
        return self.dm.GetForegroundWindow()
    
    def GetForegroundFocus(self):
        """获取前台焦点窗口"""
        return self.dm.GetForegroundFocus()
    
    def GetMousePointWindow(self):
        """获取鼠标指向的窗口"""
        return self.dm.GetMousePointWindow()
    
    def GetPointWindow(self, x, y):
        """获取指定坐标的窗口"""
        return self.dm.GetPointWindow(x, y)
    
    def EnumWindow(self, parent, title, class_name, filter):
        """枚举窗口"""
        return self.dm.EnumWindow(parent, title, class_name, filter)
    
    def GetWindowState(self, hwnd, flag):
        """获取窗口状态"""
        return self.dm.GetWindowState(hwnd, flag)
    
    def GetWindow(self, hwnd, flag):
        """获取窗口"""
        return self.dm.GetWindow(hwnd, flag)
    
    def GetSpecialWindow(self, flag):
        """获取特殊窗口"""
        return self.dm.GetSpecialWindow(flag)
    
    def SetWindowText(self, hwnd, text):
        """设置窗口文本"""
        return self.dm.SetWindowText(hwnd, text)
    
    def SetWindowSize(self, hwnd, width, height):
        """设置窗口大小"""
        return self.dm.SetWindowSize(hwnd, width, height)
    
    def GetWindowRect(self, hwnd):
        """获取窗口矩形"""
        x1, y1, x2, y2 = 0, 0, 0, 0
        result = self.dm.GetWindowRect(hwnd, x1, y1, x2, y2)
        return result, x1, y1, x2, y2
    
    def GetWindowTitle(self, hwnd):
        """获取窗口标题"""
        return self.dm.GetWindowTitle(hwnd)
    
    def GetWindowClass(self, hwnd):
        """获取窗口类名"""
        return self.dm.GetWindowClass(hwnd)
    
    def SetWindowState(self, hwnd, flag):
        """设置窗口状态"""
        return self.dm.SetWindowState(hwnd, flag)
    
    def CreateFoobarRect(self, hwnd, x, y, w, h):
        """创建矩形Foobar"""
        return self.dm.CreateFoobarRect(hwnd, x, y, w, h)
    
    def CreateFoobarRoundRect(self, hwnd, x, y, w, h, rw, rh):
        """创建圆角矩形Foobar"""
        return self.dm.CreateFoobarRoundRect(hwnd, x, y, w, h, rw, rh)
    
    def CreateFoobarEllipse(self, hwnd, x, y, w, h):
        """创建椭圆形Foobar"""
        return self.dm.CreateFoobarEllipse(hwnd, x, y, w, h)
    
    def CreateFoobarCustom(self, hwnd, x, y, pic, trans_color, sim):
        """创建自定义形状Foobar"""
        return self.dm.CreateFoobarCustom(hwnd, x, y, pic, trans_color, sim)
    
    def FoobarFillRect(self, hwnd, x1, y1, x2, y2, color):
        """Foobar填充矩形"""
        return self.dm.FoobarFillRect(hwnd, x1, y1, x2, y2, color)
    
    def FoobarDrawText(self, hwnd, x, y, w, h, text, color, align):
        """Foobar绘制文本"""
        return self.dm.FoobarDrawText(hwnd, x, y, w, h, text, color, align)
    
    def FoobarDrawPic(self, hwnd, x, y, pic, trans_color):
        """Foobar绘制图片"""
        return self.dm.FoobarDrawPic(hwnd, x, y, pic, trans_color)
    
    def FoobarUpdate(self, hwnd):
        """更新Foobar显示"""
        return self.dm.FoobarUpdate(hwnd)
    
    def FoobarLock(self, hwnd):
        """锁定Foobar"""
        return self.dm.FoobarLock(hwnd)
    
    def FoobarUnlock(self, hwnd):
        """解锁Foobar"""
        return self.dm.FoobarUnlock(hwnd)
    
    def FoobarSetFont(self, hwnd, font_name, size, flag):
        """设置Foobar字体"""
        return self.dm.FoobarSetFont(hwnd, font_name, size, flag)
    
    def FoobarTextRect(self, hwnd, x, y, w, h):
        """设置Foobar文本区域"""
        return self.dm.FoobarTextRect(hwnd, x, y, w, h)
    
    def FoobarPrintText(self, hwnd, text, color):
        """Foobar打印文本"""
        return self.dm.FoobarPrintText(hwnd, text, color)
    
    def FoobarClearText(self, hwnd):
        """清除Foobar文本"""
        return self.dm.FoobarClearText(hwnd)
    
    def FoobarTextLineGap(self, hwnd, gap):
        """设置Foobar文本行间距"""
        return self.dm.FoobarTextLineGap(hwnd, gap)
    
    def Play(self, file):
        """播放声音"""
        return self.dm.Play(file)
    
    def FaqCapture(self, x1, y1, x2, y2, quality, delay, time):
        """FAQ截图"""
        return self.dm.FaqCapture(x1, y1, x2, y2, quality, delay, time)
    
    def FaqRelease(self, handle):
        """释放FAQ资源"""
        return self.dm.FaqRelease(handle)
    
    def FaqSend(self, server, handle, request_type, time_out):
        """FAQ发送"""
        return self.dm.FaqSend(server, handle, request_type, time_out)
    
    def Beep(self, frequency, delay):
        """蜂鸣器"""
        return self.dm.Beep(frequency, delay)
    
    def FoobarClose(self, hwnd):
        """关闭Foobar"""
        return self.dm.FoobarClose(hwnd)
    
    def MoveDD(self, dx, dy):
        """移动DD"""
        return self.dm.MoveDD(dx, dy)
    
    def FaqGetSize(self, handle):
        """获取FAQ大小"""
        return self.dm.FaqGetSize(handle)
    
    def LoadPic(self, pic_name):
        """加载图片到内存"""
        return self.dm.LoadPic(pic_name)
    
    def FreePic(self, pic_name):
        """释放内存中的图片"""
        return self.dm.FreePic(pic_name)
    
    def GetScreenData(self, x1, y1, x2, y2):
        """获取屏幕数据到内存"""
        return self.dm.GetScreenData(x1, y1, x2, y2)
    
    def FreeScreenData(self, handle):
        """释放屏幕数据内存"""
        return self.dm.FreeScreenData(handle)
    
    def WheelUp(self):
        """鼠标滚轮向上"""
        return self.dm.WheelUp()
    
    def WheelDown(self):
        """鼠标滚轮向下"""
        return self.dm.WheelDown()
    
    def SetMouseDelay(self, type, delay):
        """设置鼠标延迟"""
        return self.dm.SetMouseDelay(type, delay)
    
    def SetKeypadDelay(self, type, delay):
        """设置键盘延迟"""
        return self.dm.SetKeypadDelay(type, delay)
    
    def GetEnv(self, index, name):
        """获取环境变量"""
        return self.dm.GetEnv(index, name)
    
    def SetEnv(self, index, name, value):
        """设置环境变量"""
        return self.dm.SetEnv(index, name, value)
    
    def SendString(self, hwnd, str):
        """发送字符串"""
        return self.dm.SendString(hwnd, str)
    
    def DelEnv(self, index, name):
        """删除环境变量"""
        return self.dm.DelEnv(index, name)
    
    def GetPath(self):
        """获取路径"""
        return self.dm.GetPath()
    
    def SetDict(self, index, dict_name):
        """设置字典"""
        return self.dm.SetDict(index, dict_name)
    
    def FindPic(self, x1, y1, x2, y2, pic_name, delta_color, sim, dir=0):
        """查找图片"""
        x, y = 0, 0
        result = self.dm.FindPic(x1, y1, x2, y2, pic_name, delta_color, sim, dir, x, y)
        return result, x, y
    
    def FindPicEx(self, x1, y1, x2, y2, pic_name, delta_color, sim, dir=0):
        """查找图片(返回所有匹配点)"""
        return self.dm.FindPicEx(x1, y1, x2, y2, pic_name, delta_color, sim, dir)
    
    def SetClientSize(self, hwnd, width, height):
        """设置窗口客户区大小"""
        return self.dm.SetClientSize(hwnd, width, height)
    
    def ReadInt(self, hwnd, addr, type):
        """读取整数"""
        return self.dm.ReadInt(hwnd, addr, type)
    
    def ReadFloat(self, hwnd, addr):
        """读取浮点数"""
        return self.dm.ReadFloat(hwnd, addr)
    
    def ReadDouble(self, hwnd, addr):
        """读取双精度浮点数"""
        return self.dm.ReadDouble(hwnd, addr)
    
    def FindInt(self, hwnd, addr_range, int_value_min, int_value_max, type):
        """查找整数"""
        return self.dm.FindInt(hwnd, addr_range, int_value_min, int_value_max, type)
    
    def FindFloat(self, hwnd, addr_range, float_value_min, float_value_max):
        """查找浮点数"""
        return self.dm.FindFloat(hwnd, addr_range, float_value_min, float_value_max)
    
    def FindDouble(self, hwnd, addr_range, double_value_min, double_value_max):
        """查找双精度浮点数"""
        return self.dm.FindDouble(hwnd, addr_range, double_value_min, double_value_max)
    
    def FindString(self, hwnd, addr_range, string_value, type):
        """查找字符串"""
        return self.dm.FindString(hwnd, addr_range, string_value, type)
    
    def GetModuleBaseAddr(self, hwnd, module_name):
        """获取模块基地址"""
        return self.dm.GetModuleBaseAddr(hwnd, module_name)
    
    def MoveToEx(self, x, y, w, h):
        """扩展移动"""
        return self.dm.MoveToEx(x, y, w, h)
    
    def MatchPicName(self, pic_name):
        """匹配图片名称"""
        return self.dm.MatchPicName(pic_name)
    
    def AddDict(self, index, dict_info):
        """添加字典"""
        return self.dm.AddDict(index, dict_info)
    
    def EnterCri(self):
        """进入临界区"""
        return self.dm.EnterCri()
    
    def LeaveCri(self):
        """离开临界区"""
        return self.dm.LeaveCri()
    
    def WriteInt(self, hwnd, addr, type, value):
        """写入整数"""
        return self.dm.WriteInt(hwnd, addr, type, value)
    
    def WriteFloat(self, hwnd, addr, value):
        """写入浮点数"""
        return self.dm.WriteFloat(hwnd, addr, value)
    
    def WriteDouble(self, hwnd, addr, value):
        """写入双精度浮点数"""
        return self.dm.WriteDouble(hwnd, addr, value)
    
    def WriteString(self, hwnd, addr, type, value):
        """写入字符串"""
        return self.dm.WriteString(hwnd, addr, type, value)
    
    def ReadString(self, hwnd, addr, type, length):
        """读取字符串"""
        return self.dm.ReadString(hwnd, addr, type, length)
    
    def AsmCall(self, hwnd, addr, type1, param1, type2=None, param2=None, type3=None, param3=None, type4=None, param4=None):
        """汇编调用"""
        return self.dm.AsmCall(hwnd, addr, type1, param1, type2, param2, type3, param3, type4, param4)
    
    def AsmCode(self, code):
        """汇编代码"""
        return self.dm.AsmCode(code)
    
    def SetAsmRef(self, hwnd, addr, type1, param1, type2=None, param2=None, type3=None, param3=None, type4=None, param4=None):
        """设置汇编引用"""
        return self.dm.SetAsmRef(hwnd, addr, type1, param1, type2, param2, type3, param3, type4, param4)
    
    def GetAsmRef(self, hwnd, addr, type1, param1, type2=None, param2=None, type3=None, param3=None, type4=None, param4=None):
        """获取汇编引用"""
        return self.dm.GetAsmRef(hwnd, addr, type1, param1, type2, param2, type3, param3, type4, param4)
    
    def FreeAsmRef(self, hwnd, addr):
        """释放汇编引用"""
        return self.dm.FreeAsmRef(hwnd, addr)
    
    def GetDmCount(self):
        """获取DM数量"""
        return self.dm.GetDmCount()
    
    def GetDmPath(self, index):
        """获取DM路径"""
        return self.dm.GetDmPath(index)
    
    def FindWindow(self, class_name, title):
        """查找窗口，返回窗口句柄
        Args:
            class_name: 窗口类名
            title: 窗口标题
        Returns:
            窗口句柄，如果未找到返回0
        """
        return self.dm.FindWindow(class_name, title)
    
    def FindWindowEx(self, parent, class_name, title):
        """查找窗口(扩展)，返回窗口句柄
        Args:
            parent: 父窗口句柄，0表示桌面
            class_name: 窗口类名
            title: 窗口标题
        Returns:
            窗口句柄，如果未找到返回0
        """
        return self.dm.FindWindowEx(parent, class_name, title)
    
    def FindWindowByProcess(self, process_name, class_name, title):
        """根据进程名查找窗口，返回窗口句柄
        Args:
            process_name: 进程名
            class_name: 窗口类名
            title: 窗口标题
        Returns:
            窗口句柄，如果未找到返回0
        """
        return self.dm.FindWindowByProcess(process_name, class_name, title)
    
    def FindWindowByProcessId(self, process_id, class_name, title):
        """根据进程ID查找窗口，返回窗口句柄
        Args:
            process_id: 进程ID
            class_name: 窗口类名
            title: 窗口标题
        Returns:
            窗口句柄，如果未找到返回0
        """
        return self.dm.FindWindowByProcessId(process_id, class_name, title)
    
    def GetWindowThreadProcessId(self, hwnd):
        """获取窗口线程进程ID"""
        return self.dm.GetWindowThreadProcessId(hwnd)
    
    def GetWindowProcessId(self, hwnd):
        """获取窗口所属进程ID
        Args:
            hwnd: 窗口句柄
        Returns:
            进程ID
        """
        return self.dm.GetWindowProcessId(hwnd)
    
    def GetWindowProcessPath(self, hwnd):
        """获取窗口所属进程路径
        Args:
            hwnd: 窗口句柄
        Returns:
            进程路径
        """
        return self.dm.GetWindowProcessPath(hwnd)
    
    def LockInput(self, lock):
        """锁定输入
        Args:
            lock: 1锁定，0解锁
        Returns:
            操作结果，0表示失败，1表示成功
        """
        return self.dm.LockInput(lock)
    
    def GetWindowDC(self, hwnd):
        """获取窗口设备上下文"""
        return self.dm.GetWindowDC(hwnd)
    
    def ReleaseDC(self, hwnd, hdc):
        """释放设备上下文"""
        return self.dm.ReleaseDC(hwnd, hdc)
    
    def GetPixel(self, hdc, x, y):
        """获取像素颜色"""
        return self.dm.GetPixel(hdc, x, y)
    
    def SetPixel(self, hdc, x, y, color):
        """设置像素颜色"""
        return self.dm.SetPixel(hdc, x, y, color)
    
    def GetCursorPos(self):
        """获取鼠标位置"""
        x, y = 0, 0
        result = self.dm.GetCursorPos(x, y)
        return result, x, y
    
    def SetCursorPos(self, x, y):
        """设置鼠标位置"""
        return self.dm.SetCursorPos(x, y)
    
    def GetKeyState(self, vk_code):
        """获取按键状态"""
        return self.dm.GetKeyState(vk_code)
    
    def GetAsyncKeyState(self, vk_code):
        """获取异步按键状态"""
        return self.dm.GetAsyncKeyState(vk_code)
    
    def GetClipboard(self):
        """获取剪贴板内容"""
        return self.dm.GetClipboard()
    
    def SetClipboard(self, text):
        """设置剪贴板内容"""
        return self.dm.SetClipboard(text)
    
    def IsDisplay(self, x, y):
        """判断坐标是否在显示区域"""
        return self.dm.IsDisplay(x, y)
    
    def GetDisplay(self):
        """获取显示信息"""
        return self.dm.GetDisplay()
    
    def GetScreenDepth(self):
        """获取屏幕色深"""
        return self.dm.GetScreenDepth()
    
    def GetScRect(self):
        """获取屏幕矩形"""
        return self.dm.GetScRect()
    
    def GetScSize(self):
        """获取屏幕大小"""
        return self.dm.GetScSize()
    
    def GetMouseRect(self):
        """获取鼠标矩形"""
        return self.dm.GetMouseRect()
    
    def GetDPI(self):
        """获取DPI"""
        return self.dm.GetDPI()
    
    def GetSystemInfo(self, info_type):
        """获取系统信息"""
        return self.dm.GetSystemInfo(info_type)
    
    def GetOsType(self):
        """获取操作系统类型"""
        return self.dm.GetOsType()
    
    def GetOsBuildNumber(self):
        """获取操作系统版本号"""
        return self.dm.GetOsBuildNumber()
    
    def GetOsServicePack(self):
        """获取操作系统服务包"""
        return self.dm.GetOsServicePack()
    
    def GetDiskInfo(self, disk):
        """获取磁盘信息"""
        return self.dm.GetDiskInfo(disk)
    
    def GetMemoryInfo(self):
        """获取内存信息"""
        return self.dm.GetMemoryInfo()
    
    def GetCpuInfo(self):
        """获取CPU信息"""
        return self.dm.GetCpuInfo()
    
    def GetNetCard(self):
        """获取网卡信息"""
        return self.dm.GetNetCard()
    
    def GetMac(self):
        """获取MAC地址"""
        return self.dm.GetMac()
    
    def GetID(self):
        """获取ID"""
        return self.dm.GetID()
    
    def CapturePng(self, x1, y1, x2, y2, file):
        """PNG格式截图
        Args:
            x1, y1: 左上角坐标
            x2, y2: 右下角坐标
            file: 保存文件路径
        Returns:
            操作结果，0表示失败，1表示成功
        """
        return self.dm.CapturePng(x1, y1, x2, y2, file)
    
    def CaptureGif(self, x1, y1, x2, y2, file, delay, time):
        """GIF格式截图
        Args:
            x1, y1: 左上角坐标
            x2, y2: 右下角坐标
            file: 保存文件路径
            delay: 帧延迟(毫秒)
            time: 录制时间(毫秒)
        Returns:
            操作结果，0表示失败，1表示成功
        """
        return self.dm.CaptureGif(x1, y1, x2, y2, file, delay, time)
    
    def CaptureJpg(self, x1, y1, x2, y2, file, quality):
        """JPG格式截图
        Args:
            x1, y1: 左上角坐标
            x2, y2: 右下角坐标
            file: 保存文件路径
            quality: 图片质量(1-100)
        Returns:
            操作结果，0表示失败，1表示成功
        """
        return self.dm.CaptureJpg(x1, y1, x2, y2, file, quality)
    
    def FindStrWithFont(self, x1, y1, x2, y2, str_text, color, sim, font_name, font_size, flag):
        """查找字符串(带字体)
        Args:
            x1, y1: 左上角坐标
            x2, y2: 右下角坐标
            str_text: 要查找的字符串
            color: 字体颜色
            sim: 相似度
            font_name: 字体名称
            font_size: 字体大小
            flag: 查找标志
        Returns:
            (result, x, y) - 查找结果和坐标
        """
        x, y = 0, 0
        result = self.dm.FindStrWithFont(x1, y1, x2, y2, str_text, color, sim, font_name, font_size, flag, x, y)
        return result, x, y
    
    def GetMachineCode(self):
        """获取机器码"""
        return self.dm.GetMachineCode()
    
    def GetDiskSerial(self):
        """获取磁盘序列号"""
        return self.dm.GetDiskSerial()
    
    def GetHardDiskSerial(self, disk):
        """获取硬盘序列号"""
        return self.dm.GetHardDiskSerial(disk)
    
    def GetCpuID(self):
        """获取CPU ID"""
        return self.dm.GetCpuID()
    
    def GetBiosID(self):
        """获取BIOS ID"""
        return self.dm.GetBiosID()
    
    def GetBiosSerial(self):
        """获取BIOS序列号"""
        return self.dm.GetBiosSerial()
    
    def GetBiosVersion(self):
        """获取BIOS版本"""
        return self.dm.GetBiosVersion()
    
    def GetBiosDate(self):
        """获取BIOS日期"""
        return self.dm.GetBiosDate()
    
    def GetBoardID(self):
        """获取主板ID"""
        return self.dm.GetBoardID()
    
    def GetBoardSerial(self):
        """获取主板序列号"""
        return self.dm.GetBoardSerial()
    
    def GetBoardVersion(self):
        """获取主板版本"""
        return self.dm.GetBoardVersion()
    
    def GetBoardDate(self):
        """获取主板日期"""
        return self.dm.GetBoardDate()
    
    def GetBoardMaker(self):
        """获取主板制造商"""
        return self.dm.GetBoardMaker()
    
    def GetBoardProduct(self):
        """获取主板产品"""
        return self.dm.GetBoardProduct()
    
    def GetBoardModel(self):
        """获取主板型号"""
        return self.dm.GetBoardModel()
    
    def BindWindow(self, hwnd, display, mouse, keypad, mode):
        """绑定窗口"""
        return self.dm.BindWindow(hwnd, display, mouse, keypad, mode)
    
    def BindWindowEx(self, hwnd, display, mouse, keypad, public_desc, mode):
        """绑定窗口(扩展)"""
        return self.dm.BindWindowEx(hwnd, display, mouse, keypad, public_desc, mode)
    
    def FindStrEx(self, x1, y1, x2, y2, str_text, color, sim):
        """查找字符串(扩展)"""
        return self.dm.FindStrEx(x1, y1, x2, y2, str_text, color, sim)
    
    def FindStrFast(self, x1, y1, x2, y2, str_text, color, sim):
        """快速查找字符串"""
        x, y = 0, 0
        result = self.dm.FindStrFast(x1, y1, x2, y2, str_text, color, sim, x, y)
        return result, x, y
    
    def FindStrFastEx(self, x1, y1, x2, y2, str_text, color, sim):
        """快速查找字符串(扩展)"""
        return self.dm.FindStrFastEx(x1, y1, x2, y2, str_text, color, sim)
    
    def FindPicE(self, x1, y1, x2, y2, pic_name, delta_color, sim, dir=0):
        """查找图片(增强)"""
        x, y = 0, 0
        result = self.dm.FindPicE(x1, y1, x2, y2, pic_name, delta_color, sim, dir, x, y)
        return result, x, y
    
    def FindPicEEx(self, x1, y1, x2, y2, pic_name, delta_color, sim, dir=0):
        """查找图片(增强扩展)"""
        return self.dm.FindPicEEx(x1, y1, x2, y2, pic_name, delta_color, sim, dir)
    
    def FindMultiColor(self, x1, y1, x2, y2, first_color, offset_color, sim, dir=0):
        """查找多点颜色
        Args:
            x1, y1: 左上角坐标
            x2, y2: 右下角坐标
            first_color: 第一个点的颜色
            offset_color: 偏移颜色点格式，如"00ff00|2,3|0000ff|5,5"
            sim: 相似度
            dir: 查找方向
        Returns:
            (result, x, y) - 查找结果和坐标
        """
        x, y = 0, 0
        result = self.dm.FindMultiColor(x1, y1, x2, y2, first_color, offset_color, sim, dir, x, y)
        return result, x, y
    
    def FindMultiColorEx(self, x1, y1, x2, y2, first_color, offset_color, sim, dir=0):
        """查找多点颜色(扩展)
        Args:
            x1, y1: 左上角坐标
            x2, y2: 右下角坐标
            first_color: 第一个点的颜色
            offset_color: 偏移颜色点格式，如"00ff00|2,3|0000ff|5,5"
            sim: 相似度
            dir: 查找方向
        Returns:
            所有匹配点的坐标字符串
        """
        return self.dm.FindMultiColorEx(x1, y1, x2, y2, first_color, offset_color, sim, dir)
    
    def SetWindowTransparent(self, hwnd, v):
        """设置窗口透明度
        Args:
            hwnd: 窗口句柄
            v: 透明度值(0-255)
        Returns:
            操作结果，0表示失败，1表示成功
        """
        return self.dm.SetWindowTransparent(hwnd, v)
    
    def ReadData(self, hwnd, addr, len):
        """读取进程内存数据
        Args:
            hwnd: 窗口句柄
            addr: 内存地址
            len: 读取长度
        Returns:
            读取的数据(十六进制字符串)
        """
        return self.dm.ReadData(hwnd, addr, len)
    
    def WriteData(self, hwnd, addr, data):
        """写入进程内存数据
        Args:
            hwnd: 窗口句柄
            addr: 内存地址
            data: 要写入的数据(十六进制字符串)
        Returns:
            操作结果，0表示失败，1表示成功
        """
        return self.dm.WriteData(hwnd, addr, data)
    
    def FindData(self, hwnd, addr_range, data):
        """在进程内存中查找数据
        Args:
            hwnd: 窗口句柄
            addr_range: 地址范围，如"400000-401000"
            data: 要查找的数据(十六进制字符串)
        Returns:
            找到的地址字符串，多个地址用逗号分隔
        """
        return self.dm.FindData(hwnd, addr_range, data)
    
    def FindPicMem(self, x1, y1, x2, y2, pic_info, delta_color, sim, dir=0):
        """在内存中查找图片"""
        x, y = 0, 0
        result = self.dm.FindPicMem(x1, y1, x2, y2, pic_info, delta_color, sim, dir, x, y)
        return result, x, y
    
    def FindPicMemEx(self, x1, y1, x2, y2, pic_info, delta_color, sim, dir=0):
        """在内存中查找图片(扩展)"""
        return self.dm.FindPicMemEx(x1, y1, x2, y2, pic_info, delta_color, sim, dir)
    
    def UseDict2(self, index, dict_name):
        """使用字典2"""
        return self.dm.UseDict2(index, dict_name)
    
    def ReadIni(self, section, key, file, default):
        """读取INI配置"""
        return self.dm.ReadIni(section, key, file, default)
    
    def WriteIni(self, section, key, value, file):
        """写入INI配置"""
        return self.dm.WriteIni(section, key, value, file)
    
    def ReadIniPwd(self, section, key, file, pwd, default):
        """读取加密INI配置"""
        return self.dm.ReadIniPwd(section, key, file, pwd, default)
    
    def WriteIniPwd(self, section, key, value, file, pwd):
        """写入加密INI配置"""
        return self.dm.WriteIniPwd(section, key, value, file, pwd)
    
    def DeleteIni(self, section, key, file):
        """删除INI项"""
        return self.dm.DeleteIni(section, key, file)
    
    def DeleteIniPwd(self, section, key, file, pwd):
        """删除加密INI项"""
        return self.dm.DeleteIniPwd(section, key, file, pwd)
    
    def EnableSpeedDx(self, en):
        """启用速度加速"""
        return self.dm.EnableSpeedDx(en)
    
    def EnableIme(self, en):
        """启用输入法"""
        return self.dm.EnableIme(en)
    
    def Reg(self, code, Ver):
        """注册"""
        return self.dm.Reg(code, Ver)
    
    def RegNoMac(self, code, Ver):
        """注册(无MAC)"""
        return self.dm.RegNoMac(code, Ver)
    
    def RegExNoMac(self, code, Ver, ip):
        """注册(扩展无MAC)"""
        return self.dm.RegExNoMac(code, Ver, ip)
    
    def SetEnumWindowDelay(self, delay):
        """设置枚举窗口延迟"""
        return self.dm.SetEnumWindowDelay(delay)
    
    def FindMulColor(self, x1, y1, x2, y2, color, sim):
        """查找多点颜色"""
        return self.dm.FindMulColor(x1, y1, x2, y2, color, sim)
    
    def EnableMouseMsg(self, en):
        """启用鼠标消息"""
        return self.dm.EnableMouseMsg(en)
    
    def FileExist(self, file):
        """检查文件是否存在"""
        return self.dm.FileExist(file)
    
    def CopyFile(self, src_file, dst_file):
        """复制文件"""
        return self.dm.CopyFile(src_file, dst_file)
    
    def MoveFile(self, src_file, dst_file):
        """移动文件"""
        return self.dm.MoveFile(src_file, dst_file)
    
    def CreateFolder(self, folder_name):
        """创建文件夹"""
        return self.dm.CreateFolder(folder_name)
    
    def DeleteFolder(self, folder_name):
        """删除文件夹"""
        return self.dm.DeleteFolder(folder_name)
    
    def GetFileLength(self, file):
        """获取文件长度"""
        return self.dm.GetFileLength(file)
    
    def ReadFile(self, file):
        """读取文件"""
        return self.dm.ReadFile(file)
    
    def WaitKey(self, key_code, time_out):
        """等待按键"""
        return self.dm.WaitKey(key_code, time_out)
    
    def SelectFile(self):
        """选择文件"""
        return self.dm.SelectFile()
    
    def SelectDirectory(self):
        """选择目录"""
        return self.dm.SelectDirectory()
import os
import json
import re
import sys
import csv
import time
import configparser
import threading
import win32api
import win32con
import win32gui
import shutil
import base64
import ctypes
import tkinter as tk
from datetime import datetime
from tkinter import Tk, Label, ttk, StringVar
import tkinter.filedialog as filedialog

import pyautogui
import keyboard
import pyperclip
import queue

# 修复 keyboard 库在 Windows 下热键注册了却不响应的问题:
# 该库把 GetModuleHandleW(None)(即 python.exe) 当作 hMod 传给 SetWindowsHookExW,
# 但钩子回调是 ctypes 回调、不属于任何 DLL, Windows 会返回错误126(找不到指定的模块),
# 导致开始/停止热键完全无效。对 WH_KEYBOARD_LL 钩子 hMod 传 NULL 是合法的, 强制覆盖即可。
try:
    import keyboard._winkeyboard as _winkeyboard
    _orig_set_hook = _winkeyboard.SetWindowsHookEx

    def _fixed_set_hook(idHook, lpfn, hMod, dwThreadId):
        return _orig_set_hook(idHook, lpfn, None, dwThreadId)

    _winkeyboard.SetWindowsHookEx = _fixed_set_hook
except Exception:
    pass


pyautogui.FAILSAFE = True  # 保护措施,避免失控`
pyautogui.PAUSE = 0  # 默认最小操作周期

#################################################################
# 全功能版本
#################################################################
# 运行路径(绝对, 保证任何工作目录启动都读到同一配置)
# 打包(OneDir)后: DIR 为 exe 所在目录(可写, 用于 rpa_data 运行时数据),
# 只读资源(主题/图标)从 sys._MEIPASS 解压目录读取; 源码运行时两者相同。
if getattr(sys, 'frozen', False):
    DIR = os.path.dirname(os.path.abspath(sys.executable))
    RES_DIR = sys._MEIPASS
else:
    DIR = os.path.dirname(os.path.abspath(__file__))
    RES_DIR = DIR
CfgFile = os.path.join(DIR, "rpa_data", "pyrpa.ini")
log_file = os.path.join(DIR, "rpa_data", "pyrpa.log")
IconPath = os.path.join(DIR, "rpa_data", "pyrpa.ico")

config = configparser.ConfigParser()
config.read(CfgFile, encoding='utf-8')
# 配置文件缺失/损坏时(如从其它目录启动)用默认值兜底, 避免界面读取到空值或崩溃
if not config.has_section('SAVE'):
    config.add_section('SAVE')
    config.set('SAVE', 'optionselect', '')
    config.set('SAVE', 'loopcounter', '-1')
    config.set('SAVE', 'starthotkey', 'ctrl+7')
    config.set('SAVE', 'stophotkey', 'ctrl+8')
    config.set('SAVE', 'logmethod', '0')
    config.set('SAVE', 'theme', '0')
    config.set('SAVE', 'datacsv', '')
if not config.has_section('TASKCFG'):
    config.add_section('TASKCFG')
    config.set('TASKCFG', 'autoruntaskdir', '')

LogQueue = queue.Queue()  # 用于跨线程向界面日志框推送日志
LogText = None  # 主窗口底部日志框控件, 初始化后非空

mutex = threading.Lock()
ClassWindow = 'TkTopLevel'
WindowName = 'pyrpa'
MSGWindowName = 'AutoWorkMessage'
running = -1  # 1为运行 0 为停止 停止时判断越密集 退出越及时
offseted = False  # 之前是否使用偏移
moved = False  # 之前是否使用移动
JumpLine = -1  # 行跳转标识  可实现某些行间的循环 跳转后继续顺序执行
theme = 0  # 主题
# 主题颜色(随 theme 配置切换, 在 UI 初始化时赋值)
ROW_BG = '#ffffff'      # 数据行背景
ROW_SELECT = '#dce7ff'  # 选中行背景
HEAD_BG = '#e2e2e2'     # 表头背景
FG = '#000000'          # 文字颜色

# ============ 配置载体相关常量 ============
RPA_DATA_DIR = os.path.join(DIR, 'rpa_data')            # 运行时数据目录(绝对, 保证从任何工作目录启动一致)
IMG_DIR = os.path.join(RPA_DATA_DIR, 'img')           # 截图图片目录
TASK_JSON_NAME = 'task.json'                          # 每个任务文件夹下的配置文件名
# 主界面 B 列执行动作下拉框的动作列表(与 README 指令一致)
ACTION_LIST = ['左键', '右键', '输入', '等待', '热键', '中键', '偏移', '移动', '弹窗',
               '左键按下', '左键释放', '右键按下', '右键释放'
               '按键', '滚动', '命令', '截屏', '鼠标拖拽', '相对拖拽']
# 主界面 F 列超时行为下拉框选项
TIMEOUT_CHOICES = ['弹窗', '跳过', '退出', '跳转到序号']
# 执行动作默认参数(选中动作未填参数时自动补齐,保证引擎能识别 指令=参数)
ACTION_PARAM_DEFAULT = {'左键': '1', '右键': '1', '中键': '1', '截屏': '1',
                        '左键按下': '1', '左键释放': '1', '右键按下': '1', '右键释放': '1'}
# 表格各列像素宽(表头与数据行共用,保证纵向对齐)
# 0序号 1执行动作 2参数 3启用 4图片名 5截图 6超时 7超时行为 8跳转序号 9间隔 10删除
COL_PIX = {0: 38, 1: 70, 2: 80, 3: 48, 4: 80,
           5: 58, 6: 62, 7: 88, 8: 56, 9: 52, 10: 44}
COL_WEIGHT = {1: 1, 2: 3, 4: 3}  # 可拉伸列及权重: 执行动作少量 / 参数与图片名同权重(等宽同步伸缩)
COL_HEADERS = {0: '序号', 1: '执行动作', 2: '参数', 3: '启用', 4: '图片名',
               5: ' ', 6: '超时时间', 7: '超时行为', 8: '跳转序号', 9: '找图间隔', 10: ' '}
# 找图默认置信度(与 gui_editor.CONFIDENCE 一致)
CONF2890655891bbbIDENCE = 0.9
# 工作数据默认名(输入为空或非法时使用); 名称仅允许 中文/字母/数字/_/-, 最长10字符
WORK_NAME_DEFAULT = 'task1'
WORK_NAME_MAX = 10


def sanitize_workname(name):
    """规范化工作数据名: 只保留合法字符, 截断到最长10字符; 为空时回退默认名."""
    name = ''.join(c for c in str(name or '').strip()
                   if c.isalnum() or c in '_-')
    return (name or WORK_NAME_DEFAULT)[:WORK_NAME_MAX]


def work_data_path(name):
    """工作数据目录: rpa_data/<工作数据名>/"""
    return os.path.abspath(os.path.join(RPA_DATA_DIR, sanitize_workname(name)))


def apply_theme_colors():
    """按全局 theme(0=亮 1=暗)更新表格主题颜色变量, 供初始化与运行时切换共用."""
    global ROW_BG, ROW_SELECT, HEAD_BG, FG
    if theme == 0:
        ROW_BG, ROW_SELECT, HEAD_BG, FG = '#ffffff', '#dce7ff', '#e2e2e2', '#000000'
    else:
        ROW_BG, ROW_SELECT, HEAD_BG, FG = '#2d2d2d', '#3f4f6d', '#1f1f1f', '#e8e8e8'


def resource_path(relative_path):
    return os.path.join(RES_DIR, relative_path)


#  @ 功能: 调用系统命令的线程
#  @ 参数: [I] : InputCmd 输入的参数
#  @ 备注: 针对后面在"命令"更换subprocess.Popen后控制台版本正常 但是普通版本报错的问题
def threadSysCMD(InputCmd):
    mylog('调用系统CMD 执行系统命令-->', InputCmd)
    ret = os.system(InputCmd)  # 打包后运行普通版本有窗口
    mylog("CMD 线程退出码: ", ret)


#  @ 功能: 分析要做什么
#  @ 参数: [I] : PicName 图片名字  location 找到的图片位置
#  @ 备注: PicName用于防止传进来的位置为空的情况进行重找(小概率)
#         重新找3次 moveTo读不到位置会崩溃
def Analysis(PicName, location):
    global offseted, moved, JumpLine

    def ClickFilter():
        if PicName != 'None' and location is not None:
            pyautogui.moveTo(location.x, location.y, 0)

    mylog('-----> Analysis NowRowKey:', NowRowKey)
    mylog('-----> Analysis NowRowValue:', NowRowValue)
    local = 0

    while local < Key_Value_pair and running == 1:
        mylog('CMD:', NowRowKey[local], 'Value:', NowRowValue[local])
        if NowRowKey[local] == '左键':
            if offseted is True or moved is True:
                offseted = moved = False
                for i in range(0, int(NowRowValue[local])):  # 配合相对偏移点击
                    mylog("右键点击")
                    pyautogui.leftClick()
            else:
                ClickFilter()  # 偏移和移动都没使用过 在点击前判断图片坐标是否有效 否则盲点无意义
                for i in range(0, int(NowRowValue[local])):
                    mylog("右键点击")
                    pyautogui.leftClick()

        elif NowRowKey[local] == '右键':
            if offseted is True or moved is True:
                offseted = moved = False
                for i in range(0, int(NowRowValue[local])):  # 配合相对偏移点击
                    mylog("右键点击")
                    pyautogui.rightClick()
            else:
                ClickFilter()  # 偏移和移动都没使用过 在点击前判断图片坐标是否有效 否则盲点无意义
                for i in range(0, int(NowRowValue[local])):
                    mylog("右键点击")
                    pyautogui.rightClick()

        elif NowRowKey[local] == '等待' or NowRowKey[local] == '延时':
            deadline = time.time() + float(NowRowValue[local])
            while running == 1 and time.time() < deadline:  # 等待可被停止热键打断
                time.sleep(0.1)
        elif NowRowKey[local] == '输入':
            strtemp = pyperclip.paste()
            if CurrentROW == 5:
                fp = open(r'./Source/qq_number.txt', 'r', encoding='utf-8')
                line_txt = fp.readlines()
                lines = [line.strip('\n') for line in line_txt]
                mylog(f'第 {times+1} 次: {lines[times]}')
                pyperclip.copy(str(lines[times]))
                time.sleep(0.2)
                pyautogui.hotkey('ctrl', 'v')
                time.sleep(0.2)
                pyperclip.copy(strtemp)
            else:
                pyperclip.copy(str(NowRowValue[local]))
                time.sleep(0.2)
                pyautogui.hotkey('ctrl', 'v')
                time.sleep(0.2)
                pyperclip.copy(strtemp)
        elif NowRowKey[local] == '按键':
            pyautogui.press(str(NowRowValue[local]))
        elif NowRowKey[local] == '滚动':
            pyautogui.scroll(int(NowRowValue[local]))
        elif NowRowKey[local] == '滑动':
            ClickFilter()
            Split = re.split('/', NowRowValue[local])
            mylog('滑动', Split)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
            time.sleep(1)
            pyautogui.moveRel(xOffset=int(Split[0]), yOffset=int(
                Split[1]), tween=pyautogui.linear)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
        elif NowRowKey[local] == '截屏':
            if os.path.exists('Screenshot') is not True:
                os.mkdir('Screenshot')
            ShotImgpath = 'Screenshot/Shot_' + \
                f'{time.strftime("%m%d%H%M%S ")}.png'
            pyautogui.screenshot().save(ShotImgpath)
        elif NowRowKey[local] == '热键':
            ReplaceStr = NowRowValue[local].replace('=', '+')
            ReplaceStr = ReplaceStr.replace('+', '-')
            Split = re.split('-', ReplaceStr)
            if len(Split) == 2:
                pyautogui.hotkey(Split[0], Split[1])
            elif len(Split) == 3:
                pyautogui.hotkey(Split[0], Split[1], Split[2])
        elif NowRowKey[local] == '命令':
            threading.Thread(target=threadSysCMD, args=(
                NowRowValue[local],)).start()
        elif NowRowKey[local] == '中键':
            pyautogui.middleClick()
        elif NowRowKey[local] == '移动':
            moved = True
            Split = re.split('/', NowRowValue[local])
            mylog('移动鼠标到', Split)
            pyautogui.moveTo(int(Split[0]), int(Split[1]), 0)
        elif NowRowKey[local] == '偏移':  # 相对位移 +X向右 +Y向下  负值相反
            ClickFilter()
            offseted = True
            Split = re.split('/', NowRowValue[local])
            mylog('鼠标相对移动', Split)
            pyautogui.moveRel(xOffset=int(Split[0]), yOffset=int(
                Split[1]), tween=pyautogui.linear)
        elif NowRowKey[local] == '鼠标拖拽':  # 状态栏大多数情况不需要偏移拖拽
            ClickFilter()
            Split = re.split('/', NowRowValue[local])
            mylog('鼠标拖拽', Split)
            pyautogui.dragTo(x=int(Split[0]), y=int(
                Split[1]), duration=3, button='left')
        elif NowRowKey[local] == '相对拖拽':
            Split = re.split('/', NowRowValue[local])
            mylog('相对拖拽', Split)
            pyautogui.dragRel(xOffset=int(Split[0]), yOffset=int(Split[1]), duration=0.11, button='left',
                              mouseDownUp=True)
        elif NowRowKey[local] == '弹窗' or NowRowKey[local] == '提示':
            pyautogui.alert(text=NowRowValue[local], title=MSGWindowName)
        elif NowRowKey[local] == '左键按下':
            if offseted is True:
                offseted = False
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        elif NowRowKey[local] == '左键释放':
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
        elif NowRowKey[local] == '右键按下':
            if offseted is True:
                offseted = False
            win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0)
        elif NowRowKey[local] == '右键释放':
            win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0)
        elif NowRowKey[local] == '跳转':
            JumpLine = int(NowRowValue[local])
            break
        else:
            mylog('CMD:', NowRowKey[local], '!! 未知指令', NowRowKey[local])
            pyautogui.alert(
                text='CMD: ' + NowRowKey[local] + '!! 未知指令', title=MSGWindowName)
        local += 1
        time.sleep(0.01)


#  @ 功能: 找图并执行动作
#  @ 参数: [I] :PicName 图片名字 timeout没找到循环找图的超时时间  interval下次找图的时间间隔
#  @ 备注: timeout 为0表示只找一次
def FindPicAndClick(PicName, timeout, outmethod, interval):
    ImgPath = (IMG_DIR + '\\' + PicName)
    if PicName != '' and os.path.exists(ImgPath) is True and running == 1:
        mylog(ImgPath, '图片有效')
        location = pyautogui.locateCenterOnScreen(ImgPath, confidence=0.9)
        ViewLog = True
        if location is not None:
            mylog(ImgPath, 'location is not None, Quick run')
            Analysis(ImgPath, location)
        else:
            BeginTime = time.time()
            while timeout >= 0 and location is None and running == 1:
                if ViewLog:
                    mylog(ImgPath, 'is not appear,waiting..(timeout > 0)')
                    ViewLog = False

                location = pyautogui.locateCenterOnScreen(
                    ImgPath, confidence=0.9)
                time.sleep(interval)
                if time.time() - BeginTime > timeout:
                    mylog(ImgPath, 'waiting timeout !!!!')
                    mylog('超时方法:  ' + outmethod)
                    if outmethod == '弹窗':  # pyautogui.alert和通知有冲突 通知后无法弹窗
                        pyautogui.alert(text=ImgPath + '查找超时',
                                        title=MSGWindowName)
                        return outmethod
                    else:
                        return outmethod
            while timeout == -1 and location is None and running == 1:  # 一直找图 热键停止
                if ViewLog:
                    mylog(
                        ImgPath, 'timeout = -1, is not appear,waiting..(timeout = -1)')
                    ViewLog = False

                location = pyautogui.locateCenterOnScreen(
                    ImgPath, confidence=0.9)
                time.sleep(interval)

            if location is not None and running == 1:
                mylog(ImgPath, 'appear,waiting succecs,run')
                Analysis(ImgPath, location)
            else:
                mylog(ImgPath, '找图被停止, 本步不再执行动作')
    elif PicName == '':
        mylog(WorkPath, ' Excel中的图片名为空\n[以非找图模式运行]')
        Analysis('None', None)
    else:
        mylog(ImgPath, '！！图片无效,无法继续运行')
        pyautogui.alert(text=ImgPath + ' ！！图片无效,无法继续运行', title=MSGWindowName)


#  @ 功能: 日志记录 调试时可以选择输出控制台
#  @ 参数: [I] :*BUF 输入的内容
LogOutMethod = 0


def mylog(*BUF):
    #  输出到文件
    if LogOutMethod == 1:
        with open(log_file, 'a') as log:
            print(datetime.now().strftime('%F %T:%f'), file=log, end=' ')
            for i in BUF:
                print(i, file=log, end=' ')
            print('', file=log)

    # 输出到控制台
    elif LogOutMethod == 2:
        print(datetime.now().strftime('%F %T:%f'), end=' ')
        for i in BUF:
            print(i, end=' ')
        print(end='\n')

    # 输出到界面日志框(线程安全: 仅写入队列, 由主线程定时取走)
    if LogText is not None:
        line = datetime.now().strftime(
            '%F %T:%f') + ' ' + ' '.join(str(i) for i in BUF)
        LogQueue.put(line)


Key_Value_pair = 0  # 键值对数
NowRowKey = []
NowRowValue = []
CurrentROW = 1


#  @ 功能: json 步骤数据的读写与规范化
def task_json_path(work_path):
    return os.path.join(work_path, TASK_JSON_NAME)


def load_loop_data(work_path):
    """读取数据 CSV 的运行时数据行(跳过表头, 剔除空行); 无文件返回空列表.
    优先使用用户选定的 DataCsvPath, 未选择时回退 <工作数据目录>/data.csv."""
    path = DataCsvPath or os.path.join(work_path, 'data.csv')
    if not os.path.exists(path):
        return []
    try:
        with open(path, 'r', encoding='utf-8-sig') as f:
            rows = list(csv.reader(f))
    except Exception as e:
        mylog('读取数据CSV失败: ', e)
        return []
    return [r for r in rows[1:] if any(c.strip() for c in r)]


def load_task_steps(work_path):
    """读取任务 json 的 steps 列表;文件不存在或异常时返回空列表."""
    p = task_json_path(work_path)
    if not os.path.exists(p):
        return []
    try:
        with open(p, 'r', encoding='utf-8') as f:
            data = json.load(f)
        if isinstance(data, list):
            return data
        return data.get('steps', [])
    except Exception as e:
        mylog('读取任务配置失败: ', e)
        return []


def save_task_steps(work_path, steps):
    """把 steps 列表写回任务 json(含 task 名)."""
    os.makedirs(work_path, exist_ok=True)
    p = task_json_path(work_path)
    with open(p, 'w', encoding='utf-8') as f:
        json.dump({'task': os.path.basename(work_path.rstrip('\\/')), 'steps': steps},
                  f, ensure_ascii=False, indent=2)
    return p


def normalize_step(step):
    """把一步(dict,可能缺字段)规范化为引擎可用的字段."""
    return {
        'seq': int(step.get('seq', 0) or 0),
        'action': str(step.get('action', '') or ''),
        'enabled': int(step.get('enabled', 1) or 1),
        'image': str(step.get('image', '') or '').strip(),
        'timeout': step.get('timeout', 2),
        'timeout_action': str(step.get('timeout_action', '跳过') or '跳过'),
        'jump_to': str(step.get('jump_to', '') or '').strip(),
        'interval': step.get('interval', 0.5),
    }


def parse_action(source_str):
    """把"动作=参数,动作=参数"解析为 (Key_Value_pair, keys[], values[])."""
    ReplaceStr = (source_str or '').replace(',', ',').replace(',', '=')
    Split = re.split('=', ReplaceStr)
    keys = []
    vals = []
    i = 0
    while i + 1 < len(Split):
        keys.append(Split[i])
        vals.append(Split[i + 1])
        i += 2
    return len(keys), keys, vals


def is_number(text):
    try:
        float(text)
        return True
    except (TypeError, ValueError):
        return False


def validate_steps(steps):
    """校验 json steps 数据,返回 (是否通过, 错误详情列表)."""
    errs = []
    for i, step in enumerate(steps):
        no = i + 1
        s = normalize_step(step)
        if not is_number(s['timeout']):
            errs.append('第 %d 行: E 超时时间必须为数字(可为 1、1.5、-1、0)' % no)
        if not is_number(s['interval']):
            errs.append('第 %d 行: G 找图间隔必须为数字' % no)
        if not s['action']:
            errs.append('第 %d 行: B 执行动作不能为空' % no)
        if s['timeout_action'] == '跳转到序号' and not str(s['jump_to']).isdigit():
            errs.append('第 %d 行: F 超时行为为「跳转到序号」时必须填写目标序号' % no)
    return len(errs) == 0, errs


def find_seq_index(steps, seq):
    """按 seq 序号查找对应步骤的下标,找不到返回 -1."""
    try:
        target = int(seq)
    except (TypeError, ValueError):
        return -1
    for i, s in enumerate(steps):
        if normalize_step(s)['seq'] == target:
            return i
    return -1


#  @ 功能: 主要用于找图前的参数输入
#  @ 参数: [I] :steps 任务步骤列表(来自 json)
#  @ 备注: 跳转以「序号 seq」定位(区别于原 xls 的行号)
def run_steps_once(steps, values=None):
    """执行一遍完整流程; values 为当前数据行(列表)时, 把步骤动作参数里的 $1/$2/... 替换为对应列数据."""
    global NowRowKey, NowRowValue, Key_Value_pair, JumpLine, CurrentROW
    CurrentROW = 0
    while CurrentROW < len(steps) and running == 1:
        step = steps[CurrentROW]
        if step['enabled'] != 1:
            mylog("STEP", step['seq'], '未启用操作')
        else:
            mylog('--------------work start--------------')
            mylog('STEP', step['seq'])
            action = step['action']
            if values:  # 数据驱动: 替换占位符 $1/$2/...
                for k, v in enumerate(values, 1):
                    action = action.replace('$%d' % k, v.strip())
            mylog('STEP Str: ', action)
            Key_Value_pair, NowRowKey, NowRowValue = parse_action(action)
            # 超时行为: 跳转到序号 -> 转成引擎可识别的 "跳转=N"
            outmethod = step['timeout_action']
            if outmethod == '跳转到序号':
                outmethod = '跳转=' + (step['jump_to'] or '0')
            ret = ''
            try:
                timeout = float(step['timeout'])
                interval = float(step['interval'])
            except (TypeError, ValueError):
                timeout, interval = 2, 0.5
            if step['image']:
                mylog('找图模式,图片: ', step['image'])
                ret = str(FindPicAndClick(PicName=step['image'], timeout=timeout,
                                          outmethod=outmethod, interval=interval))
            else:
                mylog('图片名为空,非找图模式运行')
                Analysis('None', None)
            mylog("FindPicAndClick ret=", ret)
            NowRowKey.clear()
            NowRowValue.clear()
            if ret == '退出':
                mylog('查找超时,退出整个查找')
                return ret
            if ret.find("跳转") != -1:
                Templist = re.split('=', ret)
                target = int(Templist[1]) if len(Templist) > 1 else 0
                idx = find_seq_index(steps, target)
                mylog("由超时行为触发的跳转到序号 ", target)
                if idx < 0:
                    mylog("请检查跳转参数")
                    pyautogui.alert(text='请检查跳转参数', title=MSGWindowName)
                    return -1
                CurrentROW = idx - 1  # 外层 +1 回到目标行

        if JumpLine != -1:
            mylog("由动作触发的跳转到序号 ", JumpLine)
            idx = find_seq_index(steps, JumpLine)
            JumpLine = -1
            if idx < 0:
                mylog("请检查跳转参数")
                pyautogui.alert(text='请检查跳转参数', title=MSGWindowName)
                return -1
            CurrentROW = idx - 1  # 外层 +1 回到目标行
        CurrentROW += 1
    return ''


def workspace(steps, values=None):
    """执行一遍完整流程.
    values: 当前要使用的数据行(list)或 None(无数据). 第 i 次循环传入第 i 行,由外层控制."""
    global StatusText, JumpLine
    StatusText = '工作'
    steps = [normalize_step(s) for s in (steps or [])]
    ok, errs = validate_steps(steps)
    if not ok:
        mylog('配置校验失败: ')
        for e in errs:
            mylog(e)
        pyautogui.alert(text='配置校验失败: \n' + '\n'.join(errs),
                        title=MSGWindowName)
        return
    if values:
        mylog('数据驱动: 当前数据行 %s' % (values,))
    else:
        mylog('本次无数据行,以原占位符执行')
    JumpLine = -1
    ret = run_steps_once(steps, values)
    mylog('works end')
    return ret


#  @ 功能: 窗口控制
#  @ 参数: [I] :wClassName 窗口类名字 wCaption窗口名
#              action = -1 关闭窗口并结束所有任务 action=1显示  action=0最小化
#              已知bug: 如果最小化窗口开始,将导致运行结束不能正常还原窗口
def WindowCtrl(wClassName, wCaption, action):
    hwnd = win32gui.FindWindow(wClassName, wCaption)
    if hwnd != 0:
        if action == -1:  # 暂不使用
            # mylog('执行窗口摧毁')
            win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
        elif action == 0:
            if win32gui.IsIconic(hwnd) is not True:
                # mylog('执行窗口最小化')
                win32gui.ShowWindow(hwnd, win32con.SW_SHOWMINIMIZED)
        elif action == 1:
            if win32gui.IsIconic(hwnd):
                # mylog('执行窗口恢复')
                win32gui.ShowWindow(hwnd, win32con.SW_SHOWNORMAL)


#  @ 功能: 在独立线程里执行窗口操作
#  @ 备注: 开始/停止热键回调运行在 keyboard 钩子线程, 若直接 ShowWindow 会等待主窗口
#          线程应答而互相死锁, 因此所有热键触发的窗口操作都放进后台线程执行.
def window_ctrl_threaded(ops):
    for cls, name, action in ops:
        try:
            WindowCtrl(cls, name, action)
        except Exception:
            pass


#  @ 功能: 开始热键绑定的事件
def begin_working():
    global running, LoopData, LpCounter, _last_begin_ts, _last_stop_ts
    if running == 1:  # 运行进行中, 忽略重复触发, 避免叠加重跑
        return
    now = time.time()
    if now - _last_begin_ts < 1.0:  # 防抖: 抑制 OS 自动重复/连击造成的重复启动
        return
    _last_begin_ts = now
    # mylog('热键按下 : begin_working')
    threading.Thread(target=window_ctrl_threaded,
                     args=([(ClassWindow, WindowName, 0)],), daemon=True).start()
    # 循环次数已由界面的 StringVar 实时同步到全局, 直接使用最新值, 不做跨线程 UI 访问
    mutex.acquire()
    running = 1
    mutex.release()
    LoopData = load_loop_data(WorkPath)  # 热键开始同样支持数据驱动
    _write_host_state('1')
    _last_stop_ts = 0.0  # 解除停止热键防抖: 本次运行第一次按停止必定生效


#  @ 功能: 结束热键绑定的事件
def finished_working():
    global running, _last_stop_ts, _last_begin_ts
    if running != 1:  # 空闲/已停止时重复触发不再切换, 防止主循环空转刷屏
        return
    now = time.time()
    if now - _last_stop_ts < 1.0:  # 防抖: 抑制自动重复/连击(仅在真正运行中记录, 避免吞掉有效停止)
        return
    _last_stop_ts = now
    # mylog('热键按下 : finished_working')
    # 恢复窗口/关闭消息弹窗放进后台线程, 避免在钩子线程中 ShowWindow 死锁
    threading.Thread(target=window_ctrl_threaded,
                     args=([(ClassWindow, WindowName, 1), (None, MSGWindowName, -1)],),
                     daemon=True).start()
    mutex.acquire()
    running = 0
    mutex.release()
    _write_host_state('0')
    _last_begin_ts = 0.0  # 解除开始热键防抖: 停止后允许立即再次开始


#  @ 功能: 把运行状态写入临时文件,供 Web 编辑器子进程跨进程读取
def _write_host_state(val):
    try:
        p = os.path.join(os.path.dirname(__file__), '.host_running')
        with open(p, 'w', encoding='utf-8') as f:
            f.write(val)
    except Exception:
        pass


StatusText = ''
_last_begin_ts = 0.0  # 开始热键防抖: 距上次触发不足阈值视为自动重复/连击
_last_stop_ts = 0.0  # 停止热键防抖


#  @ 功能: 定期刷新显示左上角状态标签
class ViewSta(tk.Frame):
    msec = 100  # 标签更新频率

    def __init__(self, parent=None, **kw):
        tk.Frame.__init__(self, parent, kw)
        mutex.acquire()
        self._running = False
        mutex.release()
        self.str1 = StringVar()
        Lab = Label(self, textvariable=self.str1,  # 设置文本内容
                    width=0,  # 设置label的宽度
                    height=0,  # 设置label的高度
                    justify='left',  # 设置文本对齐方式: 左对齐
                    anchor='nw',  # 设置文本在label的方位: 西北方位
                    font=('宋体', 8),  # 设置字体,字号
                    fg='red',  # 设置前景色
                    bg='white',  # 设置背景色
                    padx=0,  # 设置x方向内边距
                    pady=0)  # 设置y方向内边距
        Lab.pack()
        self.flag = True

    def _update(self):
        self._setstr()
        self.timer = self.after(self.msec, self._update)

    def _setstr(self):
        self.str1.set(StatusText)

    def start(self):
        self._update()
        self.pack(side=tk.TOP)


#  @ 功能: 维持左上角标签的线程(基于窗口)
def ThreadShowLabelWindow():
    mylog("ThreadShowLabelWindow")
    root = Tk()
    root.overrideredirect(True)
    t = ViewSta(root)
    t.start()
    root.mainloop()


ETLoop = None
ETStart = None
ETStop = None
LpCounter = 1
StartKey = ''
StopKey = ''
ListCfg = ['loopcounter', 'starthotkey', 'stophotkey']  # 下拉栏是独立的
XlsSource = None  # 当前任务 json 的 steps 列表(运行引擎读取)
WorkPath = ''
LoopData = []  # 运行时数据行列表(来自 CSV, 每行跑一遍流程; 为空则按循环次数跑)
DataCsvPath = ''  # 用户选定的数据 CSV 文件路径; 为空时回退 <工作数据目录>/data.csv


def KillSelf():
    TempPath = os.path.dirname(DIR)
    mylog("TempPath: ", TempPath)
    for root, dirs, files in os.walk(TempPath):
        if "_MEI" in root and DIR not in root:
            try:
                mylog("删除", root)
                shutil.rmtree(root)
            except:
                pass
        else:
            pass


#  @ 功能: 显示主界面和处理事件
class TaskRow(tk.Frame):
    """主界面操作项表格的一行: 序号/动作/参数/启用/图片名+截图/超时/超时行为/跳转序号/间隔/删除."""

    def __init__(self, parent, table, step):
        super().__init__(parent)
        self.table = table
        self.configure(bg=ROW_BG)
        step = normalize_step(step)
        self.seq = step['seq']
        # 解析 step['action'](形如 "左键=3,输入=hi"),提取动作名与首个参数
        act_name, act_param = self._split_action(step['action'])
        # 0 序号
        self.seq_lab = tk.Label(self, text=str(
            self.seq), bg=ROW_BG, fg=FG, anchor='e')
        self.seq_lab.grid(row=0, column=0, sticky='nsew', padx=2)
        # 1 执行动作(下拉选动作)
        self.act = ttk.Combobox(self, values=ACTION_LIST, width=3)
        self.act.set(act_name)
        self.act.grid(row=0, column=1, sticky='nsew', padx=2)
        self.act.bind('<<ComboboxSelected>>', self._on_act_selected)
        self.act.bind('<KeyRelease>', self._chg_single)
        # 2 参数(动作入参: 左键=点击次数、输入=文本、移动=x/y 等)
        self.pval = ttk.Entry(self, width=5)
        self.pval.insert(0, act_param)
        self.pval.grid(row=0, column=2, sticky='nsew', padx=2)
        self.pval.bind('<KeyRelease>', self._chg_single)
        # 3 启用(仅 1/0)
        self.enabled = ttk.Combobox(self, values=(
            '1', '0'), width=1, state='readonly')
        self.enabled.set(str(step['enabled']))
        self.enabled.grid(row=0, column=3, sticky='nsew', padx=2)
        self.enabled.bind('<<ComboboxSelected>>', self._chg)
        # 4 图片名(可空)+ 5 截图按钮
        self.image = ttk.Entry(self, width=8)
        self.image.insert(0, step['image'])
        self.image.grid(row=0, column=4, sticky='nsew', padx=2)
        self.image.bind('<KeyRelease>', self._chg_single)
        self.snip = ttk.Button(self, text='截图', width=4,
                               command=lambda: table.pick_image(self))
        self.snip.grid(row=0, column=5, sticky='nsew', padx=2)
        # 6 超时时间
        self.timeout = ttk.Entry(self, width=5)
        self.timeout.insert(0, str(step['timeout']))
        self.timeout.grid(row=0, column=6, sticky='nsew', padx=2)
        self.timeout.bind('<KeyRelease>', self._chg_single)
        # 7 超时行为
        self.tma = ttk.Combobox(
            self, values=TIMEOUT_CHOICES, width=6, state='readonly')
        self.tma.set(step['timeout_action'])
        self.tma.grid(row=0, column=7, sticky='nsew', padx=2)
        self.tma.bind('<<ComboboxSelected>>', self._chg)
        # 8 跳转序号(始终占位保证对齐,仅“跳转到序号”时可编辑)
        self.jump = ttk.Entry(self, width=5)
        self.jump.insert(0, step['jump_to'])
        self.jump.grid(row=0, column=8, sticky='nsew', padx=2)
        self.jump.bind('<KeyRelease>', self._chg_single)
        # 9 找图间隔
        self.interval = ttk.Entry(self, width=4)
        self.interval.insert(0, str(step['interval']))
        self.interval.grid(row=0, column=9, sticky='nsew', padx=2)
        self.interval.bind('<KeyRelease>', self._chg_single)
        # 10 删除
        self.del_btn = ttk.Button(
            self, text='✕', width=2, command=lambda: table.delete_row(self))
        self.del_btn.grid(row=0, column=10, sticky='nsew', padx=2)
        # 统一像素列宽(与表头一致,保证纵向对齐)
        for col, w in COL_PIX.items():
            self.grid_columnconfigure(col, minsize=w,
                                      weight=(1 if col in COL_WEIGHT else 0))
        self._toggle_jump()
        for w in (self, self.seq_lab):
            w.bind('<Button-1>', lambda e, r=self: table.select_row(r))

    @staticmethod
    def _split_action(action):
        """把"动作=参数"拆成 (动作名, 参数);无 = 时返回 (整串, '')."""
        action = str(action or '')
        if '=' in action:
            name, _, param = action.partition('=')
            return name.strip(), param.strip()
        return action.strip(), ''

    def _on_act_selected(self, _e=None):
        name = self.act.get().strip()
        if name in ACTION_PARAM_DEFAULT and not self.pval.get().strip():
            self.pval.delete(0, tk.END)
            self.pval.insert(0, ACTION_PARAM_DEFAULT[name])
        self._chg()

    def _toggle_jump(self):
        jump_ok = self.tma.get() == '跳转到序号'
        self.jump.configure(state='normal' if jump_ok else 'disabled')

    def _chg(self, _e=None):
        self._toggle_jump()
        self.table.changed()

    def _chg_single(self, _e=None):
        self.table.changed()

    def set_seq(self, n):
        self.seq = n
        self.seq_lab.configure(text=str(n))

    def set_selected(self, sel):
        bg = ROW_SELECT if sel else ROW_BG
        self.configure(bg=bg)
        self.seq_lab.configure(bg=bg)


class TaskTable(ttk.Frame):
    """主界面操作项表格: A序号 / B执行动作 / C启用 / D图片名 / E超时 / F超时行为 / G间隔."""

    def __init__(self, master, top_master, on_change=None, pick_csv_cb=None):
        super().__init__(master)
        self.top = top_master          # 主窗口,截图时用于最小化/还原
        self.on_change = on_change
        self.pick_csv_cb = pick_csv_cb  # 选择数据CSV的回调(由界面层提供)
        self.rows = []
        self.selected = None
        self._build()

    def _build(self):
        bar = ttk.Frame(self)
        bar.grid(row=0, column=0, sticky='ew')
        if self.pick_csv_cb:
            ttk.Button(bar, text='数据CSV', command=self.pick_csv_cb).pack(
                side=tk.LEFT, padx=2)
        ttk.Button(bar, text='添加行', command=self.add_row).pack(
            side=tk.LEFT, padx=2)
        ttk.Button(bar, text='删除选中行', command=self.delete_selected).pack(
            side=tk.LEFT, padx=2)
        ttk.Button(bar, text='↑ 上移选中行', command=self.move_up).pack(
            side=tk.LEFT, padx=2)
        ttk.Button(bar, text='↓ 下移选中行', command=self.move_down).pack(
            side=tk.LEFT, padx=2)

        # 表头与数据体用同一套网格列宽,滚动条独立占一列(与表头同高),保证纵向对齐
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)
        head = tk.Frame(self, bg=HEAD_BG, height=26)
        head.pack_propagate(False)
        head.grid(row=1, column=0, sticky='ew')
        self.head = head
        # 列宽基线：预设像素宽，若表头文字本身更宽则按文字实际宽度提升（保证表头与数据列对齐）
        self._col_min = dict(COL_PIX)
        for col in range(len(COL_PIX)):
            head.grid_columnconfigure(col, minsize=COL_PIX[col])
            lab = tk.Label(head, text=COL_HEADERS[col], bg=HEAD_BG, fg=FG,
                           font=('TkDefaultFont', 9, 'bold'), anchor='w')
            # 图片名列(4)表头文字略微右移,对齐下方 Entry 文字的内缩
            padx = (4, 2) if col == 4 else 2
            lab.grid(row=0, column=col, sticky='nsew', padx=padx)
        head.update_idletasks()
        for col, lab in enumerate(head.winfo_children()):
            self._col_min[col] = max(COL_PIX[col], lab.winfo_reqwidth() + 6)

        body = ttk.Frame(self)
        body.grid(row=2, column=0, sticky='nsew')
        self.canvas = tk.Canvas(body, highlightthickness=0)
        self.vsb = ttk.Scrollbar(
            self, orient='vertical', command=self.canvas.yview)
        self.vsb.grid(row=1, column=1, rowspan=2, sticky='ns')
        self.vsb.grid_remove()
        self.inner = tk.Frame(self.canvas, bg=ROW_BG)
        self.inner.bind('<Configure>',
                        lambda e: self.canvas.configure(scrollregion=self.canvas.bbox('all')))
        self.canvas.create_window(
            (0, 0), window=self.inner, anchor='nw', width=900)
        self.canvas.configure(yscrollcommand=self.vsb.set)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self._vsb_shown = False
        self.canvas.bind('<Configure>', self._on_resize)
        self.bind_all('<MouseWheel>', self._on_wheel)
        self._apply_widths(900)

    def _column_widths(self, total):
        """按可用总宽计算各列像素宽: 固定列取基线宽,可拉伸列按权重分配余量."""
        base = sum(self._col_min.values())
        extra = max(0, int(total) - base)
        widths = {col: w for col, w in self._col_min.items()}
        if extra and COL_WEIGHT:
            total_w = sum(COL_WEIGHT.values())
            given = 0
            for col, wt in COL_WEIGHT.items():
                share = extra * wt // total_w
                widths[col] += share
                given += share
            # 余数补给权重最大的列; 若最大权重有多个(等宽组), 补给组外权重最大的列, 保证组内严格等宽
            rest = extra - given
            if rest:
                max_wt = max(COL_WEIGHT.values())
                top = [c for c, wt in COL_WEIGHT.items() if wt == max_wt]
                if len(top) == 1:
                    widths[top[0]] += rest
                else:
                    others = [(c, wt)
                              for c, wt in COL_WEIGHT.items() if c not in top]
                    target = max(others, key=lambda kv: kv[1])[
                        0] if others else top[0]
                    widths[target] += rest
        return widths

    def _apply_widths(self, total):
        """把同一套列宽同时应用到表头与所有数据行,保证各列像素宽完全一致."""
        widths = self._column_widths(total)
        for frame in [self.head] + list(self.rows):
            for col, w in widths.items():
                frame.grid_columnconfigure(col, minsize=w, weight=0)

    def _on_resize(self, e):
        self.canvas.itemconfigure(1, width=e.width)
        self._apply_widths(e.width)
        self._update_scrollbar()

    def _update_scrollbar(self, _e=None):
        needs = self.inner.winfo_reqheight() > self.canvas.winfo_height()
        if needs and not self._vsb_shown:
            self.vsb.grid()
            self._vsb_shown = True
        elif not needs and self._vsb_shown:
            self.vsb.grid_remove()
            self._vsb_shown = False
            self.canvas.yview_moveto(0)
        self._apply_widths(self.canvas.winfo_width())

    def _on_wheel(self, event):
        # 仅接管本窗口内、非下拉弹层的滚轮(下拉弹层是独立 Toplevel)
        if getattr(event.widget, 'winfo_toplevel', None) is None:
            return
        if event.widget.winfo_toplevel() is not self.winfo_toplevel():
            return
        # 仅当指针位于表格画布内时才滚动表格, 避免滚轮连动底部日志框等其他控件
        w = event.widget
        while w is not None:
            if w is self.canvas or w is self.inner:
                self.canvas.yview_scroll(-1 * (event.delta // 120), 'units')
                return
            w = getattr(w, 'master', None)

    def set_steps(self, steps):
        for r in self.rows:
            r.destroy()
        self.rows = []
        self.selected = None
        for i, s in enumerate(steps):
            self._add_row(s, i + 1)
        if not self.rows:
            self.add_row()
        self._update_scrollbar()

    def _add_row(self, step=None, seq=None):
        if seq is not None and step is not None:
            step = dict(step)
            step['seq'] = seq
        if step is None:
            step = {'seq': len(self.rows) + 1, 'action': '左键=1', 'enabled': 1,
                    'image': '', 'timeout': 2, 'timeout_action': '跳过',
                    'jump_to': '', 'interval': 0.5}
        row = TaskRow(self.inner, self, step)
        row.configure(padx=0)
        row.pack(side=tk.TOP, fill=tk.X, pady=1)
        self.rows.append(row)
        self._renumber()
        self._apply_widths(self.canvas.winfo_width())
        self._update_scrollbar()
        return row

    def add_row(self):
        """添加新行: 选中某行时插入到该行下方, 未选中时追加到末尾"""
        sel = self.selected
        row = self._add_row()
        if sel is not None:
            i = self.rows.index(sel)
            self.rows.remove(row)
            self.rows.insert(i + 1, row)
            self._repack_rows()
            self._renumber()
        self.select_row(row)

    def delete_row(self, row):
        if row in self.rows:
            self.rows.remove(row)
            row.destroy()
        if self.selected is row:
            self.selected = None
        self._renumber()
        self._update_scrollbar()

    def delete_selected(self):
        if self.selected is not None:
            self.delete_row(self.selected)

    def _repack_rows(self):
        """按 self.rows 当前顺序重新 pack 所有行(pack 顺序即显示顺序)"""
        for r in self.rows:
            r.pack_forget()
        for r in self.rows:
            r.pack(side=tk.TOP, fill=tk.X, pady=1)

    def move_up(self):
        """选中行与上一行交换位置"""
        if self.selected is None:
            return
        i = self.rows.index(self.selected)
        if i <= 0:
            return
        self.rows[i], self.rows[i - 1] = self.rows[i - 1], self.rows[i]
        self._repack_rows()
        self._renumber()
        self._update_scrollbar()
        self.changed()

    def move_down(self):
        """选中行与下一行交换位置"""
        if self.selected is None:
            return
        i = self.rows.index(self.selected)
        if i >= len(self.rows) - 1:
            return
        self.rows[i], self.rows[i + 1] = self.rows[i + 1], self.rows[i]
        self._repack_rows()
        self._renumber()
        self._update_scrollbar()
        self.changed()

    def select_row(self, row):
        for r in self.rows:
            r.set_selected(r is row)
        self.selected = row

    def _renumber(self):
        for i, r in enumerate(self.rows):
            r.set_seq(i + 1)

    def apply_theme(self):
        """按当前主题变量就地刷新表头/画布/所有行颜色(切换主题时调用)"""
        self.head.configure(bg=HEAD_BG)
        for lab in self.head.winfo_children():
            lab.configure(bg=HEAD_BG, fg=FG)
        self.inner.configure(bg=ROW_BG)
        for r in self.rows:
            bg = ROW_SELECT if r is self.selected else ROW_BG
            r.configure(bg=bg)
            r.seq_lab.configure(bg=bg, fg=FG)

    def pick_image(self, row):
        try:
            import gui_editor
        except ImportError as e:
            pyautogui.alert(text='gui_editor 模块加载失败: %s' %
                            e, title=MSGWindowName)
            return
        WindowCtrl(ClassWindow, WindowName, 0)  # 最小化主窗口

        def on_saved(path, name):
            WindowCtrl(ClassWindow, WindowName, 1)  # 还原主窗口
            row.image.delete(0, tk.END)
            row.image.insert(0, name)
            self.changed()

        try:
            os.makedirs(IMG_DIR, exist_ok=True)
            gui_editor.SnipWindow(self.top, IMG_DIR, on_saved)
        except Exception as e:
            WindowCtrl(ClassWindow, WindowName, 1)
            pyautogui.alert(text='截图异常: %s' % e, title=MSGWindowName)

    def changed(self):
        if self.on_change:
            self.on_change()

    def get_steps(self):
        out = []
        for r in self.rows:
            tma = r.tma.get()
            name = r.act.get().strip()
            param = r.pval.get().strip()
            if name:
                if param:
                    action = name + '=' + param
                elif name in ACTION_PARAM_DEFAULT:
                    action = name + '=' + ACTION_PARAM_DEFAULT[name]
                else:
                    action = name
            else:
                action = ''
            out.append({
                'seq': r.seq,
                'action': action,
                'enabled': int(r.enabled.get() or 0),
                'image': r.image.get().strip(),
                'timeout': r.timeout.get().strip(),
                'timeout_action': tma,
                'jump_to': r.jump.get().strip() if tma == '跳转到序号' else '',
                'interval': r.interval.get().strip(),
            })
        return out


#  @ 功能: 显示主界面和处理事件(json 配置版: 主界面内置 A-G 列操作项表格)
def ThreadShowUIAndManageEvent():
    global g_fg, ETLoop, ETStart, ETStop, LpCounter, StartKey, StopKey, XlsSource, theme, WorkPath, DataCsvPath
    mylog("ThreadShowUIAndManageEvent")
    Top = tk.Tk()
    Top.title(WindowName)
    Top.tk.call("source", os.path.join(RES_DIR, "sun-valley.tcl"))
    Top.geometry("940x580+40+20")  # 默认尺寸 = 最小尺寸
    Top.minsize(940, 580)
    Top.iconbitmap(IconPath)
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
    ScaleFactor = ctypes.windll.shcore.GetScaleFactorForDevice(0)
    mylog('当前系统缩放: ', ScaleFactor, ' %')
    Top.tk.call('tk', 'scaling', ScaleFactor / 85)

    theme = int(config.get("SAVE", 'theme'))
    Top.tk.call("set_theme", "light" if theme == 0 else "dark")
    g_fg = "#000000" if theme == 0 else "#E8E8E8"
    apply_theme_colors()

    os.makedirs(IMG_DIR, exist_ok=True)  # 确保截图目录存在
    DataCsvPath = config.get(
        'SAVE', 'datacsv', fallback='').strip()  # 恢复上次选定的数据CSV

    # ---------- 顶部工具栏 ----------
    bar = ttk.Frame(Top, padding=6)
    bar.pack(side=tk.TOP, fill=tk.X)
    ttk.Label(bar, text='工作数据:', font=('宋体', 10)).pack(side=tk.LEFT)
    # 工作数据: 输入框(仅中文/字母/数字/_/-, 最长10字符), 数据保存在 rpa_data/<名称>/

    def validate_workname(new):
        return len(new) <= WORK_NAME_MAX and all(c.isalnum() or c in '_-' for c in new)

    vcmd = (Top.register(validate_workname), '%P')
    Entry_1 = ttk.Entry(bar, width=10, validate='key', validatecommand=vcmd)
    Entry_1.pack(side=tk.LEFT, padx=(2, 10))
    try:
        saved_name = config.get('SAVE', 'optionselect', fallback='').strip()
    except Exception:
        saved_name = ''
    if not saved_name or saved_name.isdigit():  # 旧版本保存的是数字索引, 视为未保存
        saved_name = WORK_NAME_DEFAULT
    Entry_1.insert(0, sanitize_workname(saved_name))

    ttk.Label(bar, text='日志:', font=('宋体', 10)).pack(side=tk.LEFT)
    LogMethodList = ['不记录', '写文件', 'Debug']
    Combobox_2 = ttk.Combobox(bar, values=LogMethodList, width=10)
    Combobox_2.pack(side=tk.LEFT, padx=(2, 10))
    try:
        lopt = int(config.get('SAVE', 'logmethod'))
    except Exception:
        lopt = 0
    Combobox_2.current(lopt if lopt < 3 else 0)

    ttk.Label(bar, text='主题:', font=('宋体', 10)).pack(side=tk.LEFT)
    ThemeList = ['亮色', '暗色']
    Combobox_3 = ttk.Combobox(bar, values=ThemeList, width=4, state='readonly')
    Combobox_3.current(theme)
    Combobox_3.pack(side=tk.LEFT, padx=(2, 10))

    ttk.Label(bar, text='循环次数:', font=('宋体', 10)).pack(side=tk.LEFT)
    # 循环次数通过 StringVar 与全局 LpCounter 实时同步: 任何线程读 LpCounter 都是界面最新值,
    # 避免热键启动时跨线程读 tkinter 控件读到旧值(如仍为 -1)导致无限循环.
    # StringVar 必须显式绑定主窗口 Top: 左上角状态小窗(TopLevel)会先创建 Tk 根,
    # 不带 master 的 StringVar 会绑到它的 Tcl 解释器, 导致循环次数输入框永远空白.
    LoopVar = tk.StringVar(master=Top)
    ETLoop = ttk.Entry(bar, width=5, textvariable=LoopVar)
    ETLoop.pack(side=tk.LEFT, padx=(2, 10))
    ttk.Label(bar, text='开始热键:', font=('宋体', 10)).pack(side=tk.LEFT)
    ETStart = ttk.Entry(bar, width=10)
    ETStart.pack(side=tk.LEFT, padx=(2, 6))
    ttk.Label(bar, text='停止热键:', font=('宋体', 10)).pack(side=tk.LEFT)
    ETStop = ttk.Entry(bar, width=10)
    ETStop.pack(side=tk.LEFT, padx=(2, 10))
    ttk.Button(bar, text='保存', command=lambda: UpdataCfg()
               ).pack(side=tk.LEFT, padx=2)
    ttk.Button(bar, text='导入', command=lambda: import_task()
               ).pack(side=tk.LEFT, padx=2)
    ttk.Button(bar, text='开始', command=lambda: Bbegin()
               ).pack(side=tk.LEFT, padx=2)

    LpCounter = config.get("SAVE", ListCfg[0]).strip()
    if not LpCounter.lstrip('-').isdigit():  # 配置为空/非法(如上次被清空保存)时回退默认 1, 避免打开空白或运行崩溃
        LpCounter = '1'
        config.set("SAVE", ListCfg[0], LpCounter)
    StartKey = config.get("SAVE", ListCfg[1])
    StopKey = config.get("SAVE", ListCfg[2])

    def _sync_loopcounter(*_):
        """界面循环次数变化时立即同步全局 LpCounter(仅在 UI 线程触发,线程安全)"""
        global LpCounter
        v = LoopVar.get().strip()
        if v.lstrip('-').isdigit():
            LpCounter = v

    LoopVar.set(LpCounter)
    LoopVar.trace_add('write', _sync_loopcounter)
    ETStart.insert(0, StartKey)
    ETStop.insert(0, StopKey)
    # 热键必须 suppress=True + trigger_on_release=True:
    # keyboard 0.13.5 在 suppress=False 时 trigger_on_release 的热键永远不会触发
    # (松开事件的键已从按下集合移除, 组合键匹配不上), 会导致开始/停止热键完全无响应.
    keyboard.add_hotkey(StartKey, begin_working, suppress=True, trigger_on_release=True)
    keyboard.add_hotkey(StopKey, finished_working, suppress=True, trigger_on_release=True)

    # ---------- 操作项表格 ----------
    def sync_work_name():
        """把输入框内容规范化后写回输入框、更新 WorkPath 与配置(不重载表格)"""
        global WorkPath
        name = sanitize_workname(Entry_1.get())
        if Entry_1.get() != name:
            Entry_1.delete(0, tk.END)
            Entry_1.insert(0, name)
        WorkPath = work_data_path(name)
        config.set('SAVE', 'optionselect', name)
        with open(CfgFile, 'w', encoding='utf-8') as file:
            config.write(file)

    def on_work_name_confirm(_e=None):
        """回车/失焦时切换工作数据: 同步名称并重载该目录下的任务配置"""
        sync_work_name()
        table.set_steps(load_task_steps(WorkPath))

    sync_work_name()
    table = TaskTable(Top, Top, pick_csv_cb=lambda: on_pick_csv())
    table.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=6, pady=4)
    table.set_steps(load_task_steps(WorkPath))

    # ---------- 底部日志输出框 ----------
    global LogText
    log_frame = ttk.Frame(Top, padding=(6, 0, 6, 6))
    log_frame.pack(side=tk.BOTTOM, fill=tk.X)
    log_text = tk.Text(log_frame, height=10, wrap='word',
                       bg=ROW_BG, fg=FG, state='disabled', font=('Consolas', 9))
    log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    log_sb = ttk.Scrollbar(log_frame, command=log_text.yview)
    log_sb.pack(side=tk.RIGHT, fill=tk.Y)
    log_text.configure(yscrollcommand=log_sb.set)
    LogText = log_text

    def pump_log():
        """定时从队列取出日志写入界面, 避免跨线程直接操作 tk 控件"""
        try:
            while True:
                line = LogQueue.get_nowait()
                log_text.configure(state='normal')
                log_text.insert(tk.END, line + '\n')
                log_text.configure(state='disabled')
                log_text.see(tk.END)
        except queue.Empty:
            pass
        Top.after(200, pump_log)

    Top.after(200, pump_log)

    def save_all(quiet=False):
        sync_work_name()
        steps = table.get_steps()
        ok, errs = validate_steps(steps)
        if not ok:
            pyautogui.alert(text='配置校验失败: \n' +
                            '\n'.join(errs), title=MSGWindowName)
            return False
        try:
            save_task_steps(WorkPath, steps)
        except Exception as e:
            pyautogui.alert(text='保存失败: %s' % e, title=MSGWindowName)
            return False
        global XlsSource
        XlsSource = steps
        mylog('已保存任务配置 ->', task_json_path(WorkPath))
        return True

    def import_task():
        """从外部 task.json 文件导入步骤到当前表格(载入后可编辑, 再点保存写入当前工作数据)"""
        path = filedialog.askopenfilename(
            title='选择要导入的 task.json',
            filetypes=[('任务配置', '*.json'), ('所有文件', '*.*')])
        if not path:
            return
        try:
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
        except Exception as e:
            pyautogui.alert(text='导入失败: %s' % e, title=MSGWindowName)
            return
        steps = data if isinstance(data, list) else data.get('steps', [])
        table.set_steps(steps)
        mylog('已导入任务配置 ->', path, ' 步骤数:', len(steps))

    def on_pick_csv():
        """选择运行数据 CSV 文件(替代默认的 data.csv), 路径保存到配置"""
        global DataCsvPath
        p = filedialog.askopenfilename(
            title='选择运行数据 CSV 文件(第一行为表头, 自动跳过)',
            filetypes=[('CSV 文件', '*.csv'), ('所有文件', '*.*')])
        if not p:
            return
        DataCsvPath = os.path.abspath(p)
        config.set('SAVE', 'datacsv', DataCsvPath)
        with open(CfgFile, 'w', encoding='utf-8') as file:
            config.write(file)
        mylog('已选择数据CSV ->', DataCsvPath)

    def on_theme_selected(_e=None):
        """切换亮/暗主题: 更新配置、ttk 主题与表格颜色, 并立即刷新界面"""
        global theme, g_fg
        theme = ThemeList.index(Combobox_3.get())
        config.set('SAVE', 'theme', str(theme))
        with open(CfgFile, 'w', encoding='utf-8') as file:
            config.write(file)
        Top.tk.call('set_theme', 'light' if theme == 0 else 'dark')
        g_fg = '#000000' if theme == 0 else '#E8E8E8'
        apply_theme_colors()
        table.apply_theme()
        log_text.configure(bg=ROW_BG, fg=FG)

    def Bbegin():
        global running, LoopData, LpCounter
        if running == 1:  # 运行进行中, 提示先停止再改循环次数, 避免用户改了次数却不生效
            pyautogui.alert(text='任务正在运行,请先按停止热键停止后再开始',
                            title=MSGWindowName)
            return
        if not save_all(quiet=True):
            return
        LoopData = load_loop_data(WorkPath)  # 数据驱动: 每行数据跑一遍流程
        running = 1
        WindowCtrl(ClassWindow, WindowName, 0)

    def UpdataCfg():
        global LpCounter, StartKey, StopKey
        v = ETLoop.get().strip()
        if v.lstrip('-').isdigit():
            LpCounter = v
        if not LpCounter.lstrip('-').isdigit():  # 空/非法输入不写入配置, 回退默认 1
            LpCounter = '1'
        ETLoop.delete(0, tk.END)  # 回填有效值, 保证下次打开不空白
        ETLoop.insert(0, LpCounter)
        new_start = ETStart.get().strip()
        new_stop = ETStop.get().strip()
        if new_start != StartKey:
            if StartKey:
                keyboard.remove_hotkey(StartKey)
            keyboard.add_hotkey(new_start, begin_working, suppress=True, trigger_on_release=True)
        if new_stop != StopKey:
            if StopKey:
                keyboard.remove_hotkey(StopKey)
            keyboard.add_hotkey(new_stop, finished_working, suppress=True, trigger_on_release=True)
        StartKey = new_start
        StopKey = new_stop
        for j in range(3):
            config.set("SAVE", ListCfg[j], [LpCounter, StartKey, StopKey][j])
        with open(CfgFile, 'w', encoding='utf-8') as file:
            config.write(file)
        save_all(quiet=True)

    def on_log_selected(event):
        global LogOutMethod
        LogOutMethod = LogMethodList.index(Combobox_2.get())
        config.set('SAVE', 'logmethod', str(LogOutMethod))
        with open(CfgFile, 'w', encoding='utf-8') as file:
            config.write(file)

    Entry_1.bind('<Return>', on_work_name_confirm)
    Entry_1.bind('<FocusOut>', on_work_name_confirm)
    Combobox_2.bind('<<ComboboxSelected>>', on_log_selected)
    Combobox_3.bind('<<ComboboxSelected>>', on_theme_selected)

    def close():
        # UI 运行在独立线程, 仅 sys.exit()/quit() 无法让整个进程退出:
        # 主线程 main() 的 while 1 死循环、状态标签线程的 root.mainloop、
        # keyboard 库的非守护钩子线程都会继续占据进程, 导致点关闭后"挂起".
        # 因此这里做最小清理后强制结束进程, 保证确实关闭。
        global running
        running = 0
        try:
            keyboard.unhook_all()  # 移除热键与钩子线程
        except Exception:
            pass
        try:
            Top.quit()  # 结束当前 Tk mainloop
        except Exception:
            pass
        try:
            Top.destroy()
        except Exception:
            pass
        os._exit(0)

    Top.protocol("WM_DELETE_WINDOW", close)
    Top.mainloop()


#  @ 功能: 解码base64图标
def WriteIcon():
    b64encodeIcon = "AAABAAEAgIAAAAEAIAAoCAEAFgAAACgAAACAAAAAAAEAAAEAIAAAAAAAAAABAMMOAADDDgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAC3g24At4NuALeDbgC3g24At4NuALeDbgC3g24At4NuALeDbgC3g24At4NuALeDbgC3g24At4NuAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAtoZtALeDbgC3g24Wt4NudbeDboC3g26At4NugLeDboC3g26At4NugLeDboC3g26At4NucLeDbg+3g24AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAC4hG0At4NuALeDblC3g278t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27kt4NuHreDbgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAALiEbgC2gm4At4NujLeDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbuG3g24et4NuAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAt4NuALeDbg63g27Et4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu4beDbh63g24AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAALmEbwC3g24At4NuMbeDbuu3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27ht4NuHreDbgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAwIpwALeDbgC3g25nt4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbuG3g24et4NuAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAC3g24At4NuA7eDbqO3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu4beDbh63g24AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAALeDbgC3g24at4Nu1reDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27ht4NuHreDbgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAt4NwALaDbgC2g24At4NuALeDbgC3g24At4NuALeDbgC3g24At4JuALiCbgC1hG8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAC3g24At4NuALeDbkW3g271t4Nu/7eD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bv+3g27/t4Nu/7eDbv+3g27/t4Nu4beDbh63g24AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAALeDbgC3g255t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g255t4NuAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAt4NuALeDbgC2gm0Ft4NuereDbve3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbuW3g26Ft4NuILeDbgC3g24AuIFvALaEbQAAAAAAAAAAAAAAAAAAAAAAAAAAALiFbAC3g24At4Nuc7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27ht4NuHreDbgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAt4NuALeDbnm3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbnm3g24AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAALeDbgC3g24AtoJtBbeDbnq3g273t4Nu/7eDbv+3g27/t4Nu/7eD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bv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbtG3g24nt4NuALaDbgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAC3g24At4NuALaCbQW3g256t4Nu97eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu4beDbh63g24AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAALiDbQC6hGcAt4NvALeDbjy3g27Rt4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27Rt4NuPbeDbgC1iG8At4NwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAt4NuALeDbgC2gm0Ft4NuereDbve3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eD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5P/i0sv/0LKm/8GVg/+4hXH/toFs/7eCbf+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7aBbP/IopL/8O3r//Hv7f/x7+3/8e/vfvbitgD71oAA+9aAAPvWgAD71oAA+9aAAPvWgAD71oAA+9aAAPvWgAD71oAA+9aAAPvWgAD71oAA+9aAAPvWgAD71oAA+9aAAPvWgAD71oAA+9aAAPvWgADgtnkAs39tKLeDbuW3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27lt4NuKreDbgAAAAAAAAAAALeDbgC3g26At4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7iEb/+hc1v/iWBG/4lgRv+JYEb/iWBG/4lgRv+JYEb/iWBG/4lgRv+IX0X/h15F/4deRf+HXkX/h15F/4ZcQ/+8p5n/8/Hw//Hv7f/w7ev/yKKS/7aBbP+3g27/t4Nu/7eCbf+8jHn/w5mI/7qIdP+2gmz/toJt/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/toFs/8iikv/w7ev/8e/t//Hv7v/058m/+9V9fvvWgID71oCA+9aAgPvWgID71oCA+9aAgPvWgID71oCA+9aAgPvWgID71oCA+9aAgPvWgID71oCA+9aAgPvWgID71oCA+9aAgPvWgID71oCA+9aAgP3ZgX3WqHamt4Nu/beDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv23g25bt4NuAAAAAAAAAAAAt4NuALeDboC3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/uIRv/6ByWv+IYEX/iWBG/4lgRv+JYEb/iWBG/4lgRv+JYEb/iWBG/4lgRv+JYEb/iWBG/4lgRv+JYEb/h15D/72omv/z8fD/8e/t//Dt6//IopL/toFs/7eDbv+3g27/t4Nu/7eDbf+2gm3/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+2gWz/yKKS//Dt6//x7+3/8e/v//bitv/71n7/+9aA//vWgP/71oD/+9aA//vWgP/71oD/+9aA//vWgP/71oD/+9aA//vWgP/71oD/+9aA//vWgP/71oD/+9aA//vWgP/71oD/+9aA//vWgP/71oD//NeA/9uveP+2gm7/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbnm3g24AAAAAAAAAAAC3g24At4NugLeDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+4hG//oHJa/4hgRf+JYEb/iWBG/4lgRv+JYEb/iWBG/4lgRv+JYEb/iWBG/4lgRv+JYEb/iWBG/4lgRv+HXkP/vaia//Px8P/x7+3/8O3r/8iikv+2gWz/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+2gm3/t4Nt/7eDbv+3g27/t4Nu/7aBbP/IopL/8O3r//Hv7f/x7+//9uK2//vWfv/71oD/+9aA//vWgP/71oD/+9aA//vWgP/71oD/+9aA//vWgP/71oD/+9aA//vWgP/71oD/+9aA//vWgP/71oD/+9aA//vWgP/71oD/+9aA//vWgP/814D/2694/7aCbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4NuebeDbgAAAAAAAAAAALeDbgC3g26At4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7iEb/+hc1v/iWBG/4lgRv+JYEb/iWBG/4lgRv+JYEb/iWBG/4lgRv+IX0b/iF9F/4hfRf+IX0X/iF9F/4ZdQ/+8p5n/8/Hw//Hv7f/w7ev/yKKS/7aBbP+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7aCbf+2gmz/uYdz/8KYh/+8jHn/t4Jt/7eDbv+3g27/toFs/8iikv/w7ev/8e/t//Hv7v/058m/+9V9fvvWgID71oCA+9aAgPvWgID71oCA+9aAgPvWgID71oCA+9aAgPvWgID71oCA+9aAgPvWgID71oCA+9aAgPvWgID71oCA+9aAgPvWgID71oCA+9aAgP3ZgX3WqHamt4Nu/beDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv23g25bt4NuAAAAAAAAAAAAt4NuALeDboC3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/uIRv/6R0Xf+JYEb/iWBG/4lgRv+JYEb/iWBG/4lgRv+JYEb/imFG/5hvTv+ddVH/nXVQ/511UP+ddVD/nHNO/8eynv/y8e//8e/t//Dt6//IopL/toFs/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Jt/7aBbP+4hXD/wJSC/9CxpP/h0cr/7efk/8+vov+2gmz/t4Nu/7eDbv+2gWz/yKKS//Dt6//x7+3/8e/t//Hv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AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAC3g24At4NuALeDbwW3g256t4Nu97eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu5beDboW3g24gt4NuALeDbgC4gW8AtoRtAAAAAAAAAAAAAAAAAAAAAAAAAAAAuIVsALeDbgC3g25zt4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbuG3g24et4NuAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAC3g24At4NuebeDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4NuebeDbgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAC3g24At4NuALeDbwW3g256t4Nu97eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu8beDbp63g24ywX5zALeDbgC4hG0AuINtAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAC3g24At4NuALeDbju3g27xt4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu4beDbh63g24AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAALeDbgC3g255t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g255t4NuAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAC3g24At4NuALeDbwW3g256t4Nu97eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu+reDbri3g25ItoNtBbeDbgC4g24AuINuAICAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAC3g24At4NuE7eDbs23g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27ht4NuHreDbgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAt4NuALeDbly3g279t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/beDblu3g24AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAC3g24At4NuALeDbwW3g256t4Nu97eDbv+3g27/t4Nu/7eDbs63g25ht4NuDreDbgC3g24At4NtALeEbwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAALeDbgC5hnAAt4NumLeDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbuG3g24et4NuAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAC3g24At4NuKreDbuW3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27lt4NuKreDbgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAC3g24At4NuALeDbwW3g256t4Nu9reDbuC3g257t4NuG7eDbgC3g24At4JvALeEbQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAt35oALeDbgC3g25bt4Nu/LeDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu4beDbh63g24AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAALeDbgC4gm4Dt4NumLeDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbpe2g24DtoNuAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA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AAAAAAAAAAAAAAAAAAC4g20AuoNqALeDbgC3gm4Dt4NuKreDbly3g255t4NuebeDbly3g24qt4JvA7eDbgC3jWgAt4VtAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAt4NuALeDbhq3g27Wt4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbuG3g24et4NuAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAC4g20At4NuALeDbgC3g24At4NuALeDbgC3g24At4NuALeDbgC3g28At4JvALaDbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAC3g24At4NuA7eDbqO3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu4beDbh63g24AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAMCKcAC3g24At4NuZ7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27ht4NuHreDbgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAuYRvALeDbgC3g24xt4Nu67eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbuG3g24et4NuAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAt4NuALeDbg63g27Et4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu4beDbh63g24AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAC4hG4AtoJuALeDboy3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27ht4NuHreDbgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAALiEbQC3g24At4NuULeDbvy3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbv+3g27/t4Nu/7eDbuS3g24et4NuAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAtoZtALeDbgC3g24Wt4NudbeDboC3g26At4NugLeDboC3g26At4NugLeDboC3g26At4NucLeDbg+3g24AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAt4NuALeDbgC3g24At4NuALeDbgC3g24At4NuALeDbgC3g24At4NuALeDbgC3g24At4NuALeDbgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA//////////////////////////////AAP//////////////////gAD//////////////////4AA//////////////////+AAP//////////////////gAD//////////////////wAA//////////////////8AAP//////////////////AAD//////////////////wAA///8AD////////////4AAP//+AAf////////D//+AAD///AAD////////gf//gAA///gAAf///////wB//wAAP//4AAH///////4AH/8AAD//+AAB///////8AAf/AAA///gAAf//////+AAB/wAAP//4AAH///////AAAP4AAD//+AAB///////gAAA+AAA///gAAf//////wAAAAAAAP//4AAH//////4AAAAAAAD//+AAB//////8AAAAAAAA///gAAf/////+AAAAAAAAP//4AAH//////AAAAAAAAD//+AAB//////gAAAAAAAA///wAA//////wAAAAAAAAP//+AAf/////4AAAAAAAAD///wAP/////+AAAAAAAAA////gf//////gAAAAAAAAP///4H//////4AAAAAAAAD///+B///////AAAAAAAAA////gf//////4AAAAAAAAP///4H//////+AAAAAAAAD///+B/gB////wAAAAAAAAAwH/gfwAP///8AAAAAAAAAMB/4HwAA////gAAAAAAAADAf+B8AAP///4AAAAAAAAAwH/geAAB////AAAAAAAAAMB/4HgAAf///wAAAAAAAADAf+B4AAH///+AAAAAAAAAwH/geAAB////wAAAAAAAAAAA4HgAAf///8AAAAAAAAAAAOB4AAH////gAAAAAAAAAADgeAAB////4AAAAAAAAAAAAHgAAf///+AAAAAAAAAAAABwAAH////gAAAAAAAAAAAAYAAB////4AAAAAAAAAAAAEAAAf///4AAAAAAAAAAAAAAAAP///gAAAAAAAAAAAAAAAAD///AAAAAAAAAAAAAAAAAD//8AAAAAAAAAAAAAPgAAB//wAAAAAAAAAAAAADwAH///4AAAAAAAAAAAAAAAAD///+AAAAAAAAAAAAAAAAB////gAAAAAAAAAAAAAAAA////4AAAAAAAAAAAAAAAA//AA+AAAAAAAAAAAAAAAAf/gAHgAAAAAAAAAAAAAAAP/wAA4AAAAAAAAAAAAAAAH/4AAGAAAAAAAAAAAAAAP//+AABgAAAAAAAAAAAAAD///gAAYAAAAAAAAAAAAAAAAAAAAGAAAAAAAAAAAAAAAAAAAABgAAAAAAAAAAAAAAAAAAAAYAAAAAAAAAAAAAAAAAAAAGAAAAAAAAAAAAAAAAAAAABgAAAAAAAAAAAAAAAAAAAAYAAAAAAAAAAAAAA///4AAGAAAAAAAAAAAAAAP//+AABgAAAAAAAAAAAAAAAf/gAAYAAAAAAAAAAAAAAAD/8AAOAAAAAAAAAAAAAAAAf/gAHgAAAAAAAAAAAAAAAD/8AD4AAAAAAAAAAAAAAAAP///+AAAAAAAAAAAAAAAAB////gAAAAAAAAAAAAAAAAP///8AAAAAAAAAAAAAA8AB////8AAAAAAAAAAAAAPgAAB///8AAAAAAAAAAAAAAAAAP///4AAAAAAAAAAAAAAAAA////4AAAAAAAAAAAAAAAAP////gAAAAAAAAAAAAQAAB////4AAAAAAAAAAAAGAAAf///+AAAAAAAAAAAABwAAH////gAAAAAAAAAAAAeAAB////4AAAAAAAAAAA4HgAAf///8AAAAAAAAAAAOB4AAH////AAAAAAAAAAADgeAAB////gAAAAAAAAMB/4HgAAf///wAAAAAAAADAf+B4AAH///8AAAAAAAAAwH/geAAB///+AAAAAAAAAMB/4HgAAf///gAAAAAAAADAf+B8AAP///wAAAAAAAAAwH/gfAAD///8AAAAAAAAAMB/4H8AD///+AAAAAAAAD///+B/gB////gAAAAAAAA////gf//////wAAAAAAAAP///4H//////4AAAAAAAAD///+B//////+AAAAAAAAA////gf//////gAAAAAAAAP///4H//////4AAAAAAAAD///wAP//////AAAAAAAAA///4AB//////4AAAAAAAAP//8AAP//////AAAAAAAAD//+AAB//////4AAAAAAAA///gAAf//////AAAAAAAAP//4AAH//////4AAAAAAAD//+AAB///////AAAAAAAA///gAAf//////4AAAPgAAP//4AAH///////AAAP4AAD//+AAB///////4AAH/AAA///gAAf///////AAH/wAAP//4AAH///////4AH/8AAD//+AAB////////AH//AAA///gAAf///////4H//4AAP//4AAH////////D//+AAD///AAD////////////gAA///4AB////////////8AAP///AA/////////////AAD//////////////////wAA//////////////////8AAP//////////////////gAD//////////////////4AA//////////////////+AAP//////////////////gAD//////////////////8AA///////////////////////////////8= "
    img = base64.b64decode(b64encodeIcon)
    file = open(IconPath, 'wb')
    file.write(img)
    file.close()


#  @ 功能: 全局初始化
def Initial():
    global LogOutMethod
    global StatusText
    global autoruntaskdir

    LogOutMethod = int(config.get('SAVE', 'logmethod'))
    autoruntaskdir = str(config.get('TASKCFG', 'autoruntaskdir'))
    mylog('Run path:', DIR)
    mylog('Execute File:', sys.argv[0])
    mylog('autoruntaskdir: ', autoruntaskdir)
    # 删除上次的运行文件放到程序关闭时  但-c版本需要关闭窗口才能删除文件 关闭控制台时文件将不会被清除 但下次正常关闭时可以删除之前运行的所有垃圾
    StatusText = '启动'
    if os.path.exists(IconPath) is not True:
        WriteIcon()
    os.makedirs(RPA_DATA_DIR, exist_ok=True)  # 确保工作数据目录存在
    mylog('工作数据目录:', RPA_DATA_DIR)


RunCounter = 0
times = 0


#  很多警告都是拼写相关 建议关掉这些不必要的警告
def MainWork():
    global StatusText
    global RunCounter
    global times
    global running
    mylog('等待热键按下,或点击开始')

    # 空闲(-1)与停止(0)都继续原地等待, 直到收到开始信号(running==1)
    # 若停止键把 running 置 0 就返回, 会让 main() 主循环立即重置 running=-1 并重启本函数,
    # 停止键重复触发时会在 0/-1 之间空转, 疯狂刷日志并卡死界面
    while running != 1:
        time.sleep(0.1)
        StatusText = '准备'

    time.sleep(0.5)  # 等待窗口退出
    RunCounter = int(LpCounter)
    mylog('本次运行循环次数: ', RunCounter)
    n_data = len(LoopData)
    if LoopData:
        mylog('检测到 data.csv 数据 %d 行 → 第 i 次循环使用第 i 行数据' % n_data)
    if RunCounter == -1:
        mylog('进入一直循环')
        it = 1
        while running == 1:
            # 一直循环: 有数据时按行循环取模复用; 无数据则 None
            row = LoopData[(it - 1) % n_data] if LoopData else None
            it += 1
            if workspace(XlsSource, row) == '退出':
                break
    else:
        numCounter = 0
        times = -1
        totalCounter = RunCounter
        for it in range(1, RunCounter + 1):
            if running != 1:
                break
            numCounter += 1
            times += 1
            # 以循环次数为准: 第 it 次用第 it 行; 超出数据行数的次数不再替换占位符
            row = LoopData[it - 1] if LoopData and it <= n_data else None
            mylog('\n【运行', numCounter, '/', totalCounter, '次 ↓】')
            if workspace(XlsSource, row) == '退出':
                break
    mylog('EXCEL遍历结束')
    # 一次运行结束(自然跑完或被停止热键打断)必须复位 running,
    # 否则 main() 的 while 1 会立刻再次调用本函数重跑任务 → 循环次数形同虚设、永远循环
    mutex.acquire()
    running = 0
    mutex.release()
    _write_host_state('0')


def main():
    """程序入口: 初始化并运行主循环(供 pyproject [project.scripts] 调用)."""
    Initial()
    threading.Thread(target=ThreadShowLabelWindow).start()
    threading.Thread(target=ThreadShowUIAndManageEvent).start()

    while 1:
        mylog('*********************主循环*********************')
        mutex.acquire()
        mutex.release()
        MainWork()


if __name__ == '__main__':
    main()

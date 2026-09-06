"""
PyRPA GUI 截图工具
====================================
提供:
- SnipWindow   全屏框选截图,保存到当前任务文件夹

由 main.py 在按钮事件中调用;窗口运行在 tkinter 主线程内.
(旧版基于 .xls 的任务编辑器 TaskEditorWindow 及其 xlrd/xlwt 读写已删除,
 任务存储已改用 main.py 的 json 方案.)
"""

import os
import time
import tkinter as tk
from tkinter import messagebox

CONFIDENCE = 0.9  # 与 main.py FindPicAndClick 的找图置信度保持一致


# ---------------------------------------------------------------- 截图器
class SnipWindow:
    """全屏遮罩框选截图.

    流程: 全屏半透明遮罩 → 拖拽画框(Esc 取消)→ 隐藏遮罩 → ImageGrab 截取
    → 时间戳命名存入任务文件夹 → 回调 (png绝对路径, 文件名).
    """

    def __init__(self, master, work_path, on_saved):
        self.work_path = work_path
        self.on_saved = on_saved
        self.start_x = self.start_y = 0
        self.rect_id = None

        self.top = tk.Toplevel(master)
        self.top.overrideredirect(True)
        self.top.attributes('-topmost', True)
        sw = self.top.winfo_screenwidth()
        sh = self.top.winfo_screenheight()
        self.top.geometry('%dx%d+0+0' % (sw, sh))
        self.top.attributes('-alpha', 0.3)
        self.top.configure(bg='black')

        self.canvas = tk.Canvas(self.top, cursor='cross',
                                bg='grey', highlightthickness=0)
        self.canvas.pack(fill=tk.BOTH, expand=True)
        self.canvas.bind('<ButtonPress-1>', self._on_press)
        self.canvas.bind('<B1-Motion>', self._on_drag)
        self.canvas.bind('<ButtonRelease-1>', self._on_release)
        self.top.bind('<Escape>', lambda e: self.top.destroy())
        self.top.focus_force()
        self.canvas.create_text(sw // 2, 30, text='拖拽框选截图区域,Esc 取消',
                                fill='yellow', font=('宋体', 14, 'bold'))

    def _on_press(self, event):
        self.start_x, self.start_y = event.x, event.y
        if self.rect_id:
            self.canvas.delete(self.rect_id)
        self.rect_id = self.canvas.create_rectangle(self.start_x, self.start_y,
                                                    self.start_x, self.start_y,
                                                    outline='red', width=2)

    def _on_drag(self, event):
        if self.rect_id:
            self.canvas.coords(self.rect_id, self.start_x,
                               self.start_y, event.x, event.y)

    def _on_release(self, event):
        x1, y1 = min(self.start_x, event.x), min(self.start_y, event.y)
        x2, y2 = max(self.start_x, event.x), max(self.start_y, event.y)
        self.top.withdraw()  # 截屏前隐藏遮罩,避免遮罩被截进图片
        self.top.after(200, lambda: self._grab(x1, y1, x2, y2))

    def _grab(self, x1, y1, x2, y2):
        try:
            if x2 - x1 < 5 or y2 - y1 < 5:
                messagebox.showwarning('框选截图', '选区太小(<5 像素),已取消')
                self.top.destroy()
                return
            from PIL import ImageGrab
            img = ImageGrab.grab(bbox=(x1, y1, x2, y2))
            os.makedirs(self.work_path, exist_ok=True)
            name = 'img_' + time.strftime('%Y%m%d_%H%M%S') + '.png'
            path = os.path.join(self.work_path, name)
            img.save(path)
            self.top.destroy()
            self.on_saved(path, name)
        except Exception as exc:
            self.top.destroy()
            messagebox.showerror('框选截图', '截图失败: %s' % exc)

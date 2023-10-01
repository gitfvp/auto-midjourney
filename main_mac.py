import threading
import time
import random
import pandas as pd
from AppKit import NSScreen, NSEvent
from Foundation import NSAppleScript
from tkinter import Tk, Label, Entry, Button, StringVar

class App:
    def __init__(self, root):
        self.root = root
        root.title("midjourney自动产出器 by 老陆 vx:laolu2045") 
        screen = NSScreen.mainScreen()
        root.geometry("{}x{}".format(int(screen.frame().size.width), int(screen.frame().size.height)))

        self.excel_path = StringVar()
        self.wait_time_after_paste = StringVar()
        self.cmd_sum = StringVar()
        self.wait_time_after_onece = StringVar()

        Label(root, text="Excel路径:").grid(row=0)
        Label(root, text="粘贴后等待时间(秒):").grid(row=1)
        Label(root, text="一轮发几条:").grid(row=2)
        Label(root, text="每轮等待时间(秒):").grid(row=3)
        
        Entry(root, textvariable=self.excel_path).grid(row=0, column=1)
        Entry(root, textvariable=self.wait_time_after_paste).grid(row=1, column=1)
        Entry(root, textvariable=self.cmd_sum).grid(row=2, column=1)
        Entry(root, textvariable=self.wait_time_after_onece).grid(row=3, column=1)

        self.start_button = Button(root, text="开始", command=self.start)
        self.start_button.grid(row=4, column=0)

        self.stop_button = Button(root, text="停止", command=self.stop, state="disabled")
        self.stop_button.grid(row=4, column=1)

        self.running = False

    def start(self):
        self.running = True
        self.start_button.config(state="disabled")
        self.stop_button.config(state="normal")

        threading.Thread(target=self.run).start()
        
        NSEvent.addGlobalMonitorForEventsMatchingMask(self.handler, NSEventMask.keyDown)

    def handler(self, event):
        if event.characters == '\x1b': # ESC键
            self.stop()
        return event

    def run(self):
        df = pd.read_excel(self.excel_path.get(), header=None)
        wait_time_after_paste = int(self.wait_time_after_paste.get())
        cmd_sum = int(self.cmd_sum.get())
        wait_time_after_onece = int(self.wait_time_after_onece.get())

        try:
            for i, row in df.iterrows():
                if not self.running:
                    break

                for cell in row:
                    cmd = f'echo "{cell}" | pbcopy'
                    NSAppleScript(source=cmd).executeAndReturnError_(None)
                    time.sleep(0.5)
                    
                    cmd = 'paste'
                    NSAppleScript(source=cmd).executeAndReturnError_(None)
                    time.sleep(wait_time_after_paste)

                    cmd = 'key code 36' # enter键
                    NSAppleScript(source=cmd).executeAndReturnError_(None)
                    time.sleep(1)

                    if (i + 1) % cmd_sum == 0:
                        time.sleep(wait_time_after_onece)

        finally:
            self.running = False
            self.start_button.config(state="normal")
            self.stop_button.config(state="disabled")

    def stop(self):
        self.running = False
        
root = Tk()
app = App(root)
root.mainloop()
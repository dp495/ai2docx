import tkinter as tk
from tkinter import font
import os
import re
import ctypes
import datetime
import atexit
import subprocess

temp_files = []

def print_log(msg,err=True):

    now = datetime.datetime.now().strftime("%X")
    log_entry.config(state='normal')
    log_entry.delete(0, tk.END)
    if err:
        log=now+' 错误：'+msg
    else:
        log=now+' '+msg
    log_entry.insert(0, log)
    log_entry.config(state='readonly')


def delete_temp_files():
    for file in temp_files:
        try:
            os.remove(file)
        except Exception as e:
            print_log(f'删除文件失败-{e}')

# 注册退出时的清理函数
atexit.register(delete_temp_files)

def convert_to_word():

    md=text_area.get("1.0", tk.END)
    text_area.delete("1.0", tk.END)

    if re.fullmatch(r'^\s*$', md):
        print_log('输入内容为空')
        return
    md=re.sub(r'\\\( *','$',md)
    md=re.sub(r' *\\\)','$',md)
    md=md.replace('\\[','$$')
    md=md.replace('\\]','$$')
    #md=re.sub(r'\\\[\s*','$$',md)
    #md=re.sub(r'\s*\\\]','$$',md)

    filename = re.sub(r'[<>:"/\\|?*\x00-\x1F]+', '_', re.sub(r'\s+', ' ', md[0:11]))
    
    try:
        with open(f'{filename}.md', 'w', encoding='utf-8') as f:
            f.write(md)    
    except Exception as e:
        print_log(f'写文件失败-{e}')
        return 
    
    try:
        pcp=subprocess.run(f'pandoc -f commonmark_x -t docx "{filename}.md" -o "{filename}.docx"',bufsize=0,stderr=subprocess.STDOUT)
    except Exception as e:
        print_log(f'调用pandoc失败-{e}')
        return
    if pcp.returncode != 0:
        print_log(f'转换失败：{re.sub(r'\s', ' ', pcp.stdout)}')
        return
    
    temp_files.append(f'{filename}.md')
    temp_files.append(f'{filename}.docx')
    try:
        os.startfile(f'{filename}.docx')
        print_log(f'转换成功：{filename}.md 与{filename}.docx',0)
    except Exception as e:
        print_log(f'文件打开失败-{e}')
        return

# 创建主窗口

#告诉操作系统使用程序自身的dpi适配
ctypes.windll.shcore.SetProcessDpiAwareness(1)

#获取屏幕的缩放因子
scale_factor=ctypes.windll.shcore.GetScaleFactorForDevice(0)
scale_factor/=50
pad0=int(3*scale_factor)
pad1=int(5*scale_factor)

root = tk.Tk()
root.tk.call('tk', 'scaling', scale_factor)
root.title("语言模型输出转为 Word")
root.geometry(f"{int(320*scale_factor)}x{int(240*scale_factor)}")

# 标题
# 创建一个框架来容纳标题
title_frame = tk.Frame(root)
title_frame.pack(fill=tk.X)
# 添加标题
instruction_label = tk.Label(title_frame, text="请粘贴复制按钮复制的文字：")
instruction_label.pack(side=tk.LEFT,padx=pad1,pady=pad1)

# 输入
# 创建一个框架来容纳多行输入框和滚动条
input_frame = tk.Frame(root)
input_frame.pack(fill=tk.BOTH,expand=True)
# 添加一个多行输入框
text_area = tk.Text(input_frame, height=0, width=0, borderwidth=0, font='TkFixedFont')
text_area.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
# 添加一个垂直滚动条
scrollbar = tk.Scrollbar(input_frame, command=text_area.yview)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
text_area['yscrollcommand']=scrollbar.set

# 日志
# 添加一个只读的单行输入框来显示日志
SmallFixed=font.nametofont('TkFixedFont').copy()
SmallFixed['size'] = int(font.nametofont('TkFixedFont').cget('size')*0.8)
log_entry =tk.Entry(root,font=SmallFixed,state='readonly')
log_entry.pack(fill=tk.X)

# 添加一个按钮
read_button = tk.Button(root, text="转换", command=convert_to_word)
read_button.pack(pady=pad0)

# 运行主循环，显示窗口
root.mainloop()
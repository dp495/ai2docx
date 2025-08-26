import tkinter as tk
from tkinter import font
import os
import platform
import subprocess
import datetime
import atexit
import re


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

    filename = re.sub(r'[<>:"/\\|?*\x00-\x1F]+', '_', re.sub(r'\s+', ' ', md[0:11].rstrip()))
    
    try:
        with open(f'{filename}.md', 'w', encoding='utf-8') as f:
            f.write(md)    
    except Exception as e:
        print_log(f'写文件失败-{e}')
        return 
    
    try:
        pcp=subprocess.run(f'pandoc -f commonmark_x -t docx "{filename}.md" -o "{filename}.docx"'\
                           ,bufsize=0,stderr=subprocess.STDOUT,creationflags=subprocess.CREATE_NO_WINDOW)
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


#DPI适配
os_name = platform.system()
scale_factor = 1.0
if os_name == "Windows":
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
        scale_factor=ctypes.windll.shcore.GetScaleFactorForDevice(0)
        scale_factor*=0.02
    except Exception as e:
        print(f"适配DPI失败-{e}")
        scale_factor=1.0

pad0=int(3*scale_factor)
pad1=int(5*scale_factor)

# 创建主窗口
root = tk.Tk()
root.tk.call('tk', 'scaling', scale_factor)
root.title("语言模型输出转为 Word")
root.geometry(f"{int(320*scale_factor)}x{int(256*scale_factor)}")


# 创建一个较小的等宽字体
SmallFixed = font.nametofont('TkFixedFont').copy()
SmallFixed['size'] = int(font.nametofont('TkFixedFont').cget('size') * 0.8)

SmallFont = font.nametofont('TkDefaultFont').copy()
SmallFont['size'] = int(font.nametofont('TkDefaultFont').cget('size') * 0.7)


# 使用 grid 布局
root.grid_columnconfigure(0, weight=1)
root.grid_rowconfigure(2, weight=1)

# 标题
instruction_label = tk.Label(root, text="请粘贴复制按钮复制的文字：")
instruction_label.grid(row=0, column=0, sticky='w', padx=pad1, pady=pad1)

# 提示
hint_label = tk.Label(root, text="注：转换按钮无条件清空输入框，退出程序无条件删除所有输出文件。修改输出文件后请备份！",
                      font=SmallFont, wraplength=int(320*scale_factor) - 2 * pad1, justify="left")
hint_label.grid(row=1, column=0, sticky='w', padx=pad1)

def update_warplength(event):
    instruction_label.config(wraplength=root.winfo_width() - 2 * pad1)
    hint_label.config(wraplength=root.winfo_width() - 2 * pad1)
root.bind("<Configure>", update_warplength)

# 输入框和滚动条
# 创建一个框架来容纳多行输入框和滚动条
input_frame = tk.Frame(root)
input_frame.grid(row=2, column=0, sticky='nsew')
input_frame.grid_columnconfigure(0, weight=1)
input_frame.grid_rowconfigure(0, weight=1)

# 添加一个多行输入框
text_area = tk.Text(input_frame, height=0, width=0, borderwidth=0, font='TkFixedFont')
text_area.grid(row=0, column=0, sticky='nsew')

# 添加一个垂直滚动条
scrollbar = tk.Scrollbar(input_frame, command=text_area.yview)
scrollbar.grid(row=0, column=1, sticky='ns')
text_area['yscrollcommand'] = scrollbar.set

# 日志
log_entry = tk.Entry(root, font=SmallFixed, state='readonly')
log_entry.grid(row=3, column=0, sticky='ew')

# 按钮
read_button = tk.Button(root, text="转换", command=convert_to_word)
read_button.grid(row=4, column=0, pady=pad0)


# 运行主循环，显示窗口
root.mainloop()

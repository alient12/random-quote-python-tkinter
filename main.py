from turtle import window_width
import win32com.client
import pandas as pd
import random

import tkinter as tk
from tkinter import ttk
from tkinter.filedialog import asksaveasfilename, askopenfilename

import threading
import time

bg_color = "#202020"
fg_color = "#3bc9ff"
# bg_color = "#E0C9A6"
# fg_color = "#654321"

_window_x = 500
_window_y = 300

def animate_window(window, window_height, close=False):
    timer = 15
    frames_num = 20
    global _window_x, _window_y
    wh = 1
    if close:
        wh = window_height
    slide = window_height // frames_num
    window.geometry(f"550x{wh}+{_window_x}+{_window_y}")
    window.overrideredirect(True)
    window.resizable(width=False, height=False)
    window.attributes("-alpha", 0.85)
    window.focus()
    window.configure(bg=bg_color)
    def update():
        nonlocal wh
        if close:
            wh -= slide
        else:
            wh += slide
        
        if close:
            window.geometry(f"550x{wh}+{_window_x}+{_window_y}")
            if wh > 0:
                window.after(timer, lambda:update())
            if wh < 5:
                window.destroy()
        else:
            if wh > window_height:
                wh = window_height
            window.geometry(f"550x{wh}+{_window_x}+{_window_y}")
            if wh < window_height:
                window.after(timer, lambda:update())
    window.after(timer, lambda:update())
        
def gui(text, author):
    window = tk.Tk()

    _offsetx = 0
    _offsety = 0
    global _window_x, _window_y

    qoute_str = text
    cnt = 0
    row_cnt = 0
    while True:
        cnt += 69
        row_cnt += 1
        if cnt >= len(qoute_str):
            break
        qoute_str = qoute_str[:cnt]+ "\n" + qoute_str[cnt:]
    qoute_str = qoute_str + "\n" + 75 * " " + author

    window_height = 69 + 20 * row_cnt
    animate_window(window, window_height)

    window.title("Quote of the day")
    # window.geometry(f"550x{window_height}+{_window_x}+{_window_y}")
    window.overrideredirect(True)
    window.resizable(width=False, height=False)
    window.attributes("-alpha", 0.85)
    window.focus()
    window.configure(bg=bg_color)


    def dragwin(event):
        nonlocal _offsetx, _offsety
        global _window_x, _window_y
        delta_x = window.winfo_pointerx() - _offsetx
        delta_y = window.winfo_pointery() - _offsety
        x = _window_x + delta_x
        y = _window_y + delta_y
        window.geometry("+{x}+{y}".format(x=x, y=y))
        _offsetx = window.winfo_pointerx()
        _offsety = window.winfo_pointery()
        _window_x = x
        _window_y = y

    def clickwin(event):
        nonlocal _offsetx, _offsety
        _offsetx = window.winfo_pointerx()
        _offsety = window.winfo_pointery()

    def close_window():
        # speaker.pause()
        speaker.Skip("Sentence", 5)
        speaker.Skip("Sentence", 5)
        animate_window(window, window_height, close=True)
    
    window.bind("<Button-1>", clickwin)
    window.bind("<B1-Motion>", dragwin)

    qoute_lbl = tk.Label(
        window,
        text=qoute_str,
        justify=tk.LEFT,
        font=("Arial", 13),
        fg=fg_color,
        bg=bg_color,
    )
    qoute_lbl.place(relx=0.02, rely=0.2, anchor="nw")

    close_btn = tk.Button(
        window, text="OK", command=close_window, fg=fg_color, bg=bg_color, borderwidth=0
    )
    close_btn.place(relx=0.45, y=window_height - 36, width=50, height=35, anchor="nw")

    top_line_btn = tk.Button(
        window, text="", fg=bg_color, bg=fg_color, state=tk.DISABLED
    )
    top_line_btn.place(relx=0, y=0, width=550, height=5, anchor="nw")

    bot_line_btn = tk.Button(
        window, text="", fg=bg_color, bg=fg_color, state=tk.DISABLED
    )
    bot_line_btn.place(relx=0, rely=0.96, width=550, height=5, anchor="nw")

    window.mainloop()

def speak(text, author):
    speaker.Speak("Quote of the day")
    speaker.Speak(text)
    speaker.Speak(author)

df = pd.read_excel('Motivational Quotes Database.xlsx')
author_df = df["Author"]
quotes_df = df["Quotes"]
rand_indx = random.randint(0, len(df) - 1)
text = quotes_df[rand_indx]
author = author_df[rand_indx]
print(text, author)
speaker = win32com.client.Dispatch("SAPI.SpVoice")
vcs = speaker.GetVoices()
speaker.Voice
speaker.SetVoice(vcs.Item(1))

speak_thread = threading.Thread(target=speak, args=[text, author])
gui_thread = threading.Thread(target=gui, args=[text, author])

speak_thread.start()
gui_thread.start()
import tkinter as tk

def create_entries(num_entries):
    global current_row
    entries = []
    for i in range(num_entries):
        entry = tk.Entry(inner_frame)
        entry.grid(row=current_row, column=i, padx=5, pady=5)
        entries.append(entry)
    
    current_row += 1
    add_button.grid(row=current_row, column=0, columnspan=5, pady=10)

    # 更新 Scrollbar
    inner_frame.update_idletasks()
    scrollbar_frame.config(scrollregion=scrollbar_frame.bbox("all"))

def add_new_row():
    create_entries(5)

# 建立主視窗
root = tk.Tk()
root.title("不使用 Canvas 的滾動條")

# 外框架與滾動條
outer_frame = tk.Frame(root)
outer_frame.pack(fill=tk.BOTH, expand=True)

# 設置內部滾動 Frame
scrollbar = tk.Scrollbar(outer_frame)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

scrollbar_frame = tk.Canvas(outer_frame, yscrollcommand=scrollbar.set)
scrollbar_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

scrollbar.config(command=scrollbar_frame.yview)

# 在 Canvas 中放置 Frame 來放 Entry
inner_frame = tk.Frame(scrollbar_frame)
scrollbar_frame.create_window((0, 0), window=inner_frame, anchor="nw")

# 初始化按鈕，放置在內部 Frame
current_row = 1
add_button = tk.Button(inner_frame, text="新增一行", command=add_new_row)
add_button.grid(row=0, column=0, columnspan=5, pady=10)

# 更新 Scrollbar 區域
inner_frame.update_idletasks()
scrollbar_frame.config(scrollregion=scrollbar_frame.bbox("all"))

# 啟動主循環
root.mainloop()

import tkinter as tk
from tkinter import ttk, messagebox
from PIL import Image, ImageTk

from excel import ExcelCMS

class App(tk.Tk):
    def __init__(self):
        super().__init__()

        # Title, Icon, size
        self.title('德龍企業社')
        self.iconbitmap('./image/icon.ico')
        self.geometry('1024x768')

        # 創建 IndexPage 和 CMSPage 兩個頁面
        self.index_page = IndexPage(self)
        self.cms_page = CMSPage(self)
        self.quote_page = QuotePage(self)
        self.report_page = ReportPage(self)

        # 預設顯示 IndexPage
        self.index_page.pack(fill='both', expand=True)

    def change_page(self, page):
        # 隱藏當前顯示的頁面
        for widget in self.winfo_children():
            widget.pack_forget()

        # 顯示傳入的目標頁面
        page.pack(fill='both', expand=True)

class IndexPage(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.resize_ids = {}  # 用於跟踪各個Label的after()事件ID

        self.create_page()  # 在初始化時創建頁面佈局

    def create_page(self):
        # index Frame
        self.frame_container = tk.Frame(self)
        self.frame_container.pack(fill='both', expand=True)
        
        # index image
        img = Image.open('./image/logo.png')
        self.label_img = tk.Label(self.frame_container)
        self.label_img.place(relx=0.5, rely=0.45, relwidth=0.55, relheight=0.5, anchor='center')
        
        # title
        self.title = tk.Label(self.frame_container, text='德龍企業社', font=('標楷體', 30, 'bold'), fg='#000000')
        self.title.place(relx=0.5, rely=0.1, relheight=0.1, relwidth=0.5, anchor='center')
        
        # buttons
        self.btn_quote = tk.Button(self.frame_container, text='報價單', font=('標楷體', 20, 'bold'), fg='#000000', command=lambda: self.master.change_page(self.master.quote_page))
        self.btn_quote.place(relx=0.25, rely=0.85, relwidth=0.2, relheight=0.1, anchor='center')
        
        self.btn_cms = tk.Button(self.frame_container, text='客戶管理', font=('標楷體', 20, 'bold'), fg='#000000', command=lambda: self.master.change_page(self.master.cms_page))
        self.btn_cms.place(relx=0.5, rely=0.85, relwidth=0.2, relheight=0.1, anchor='center')
        
        self.btn_report = tk.Button(self.frame_container, text='營收報表', font=('標楷體', 20, 'bold'), fg='#000000', command=lambda: self.master.change_page(self.master.report_page))
        self.btn_report.place(relx=0.75, rely=0.85, relwidth=0.2, relheight=0.1, anchor='center')

        self.bind_resize_event(self.label_img, img)

    def bind_resize_event(self, label, image):
        label.bind("<Configure>", lambda event: self.delayed_resize_image(event, label, image))

    def delayed_resize_image(self, event, label, image):
        # 如果該Label已有延迟更新的計劃，取消它
        if label in self.resize_ids:
            self.after_cancel(self.resize_ids[label])

        # 延迟 100 毫秒后调用真正的图片更新函数
        self.resize_ids[label] = self.after(100, lambda: self.resize_image(label, image))

    def resize_image(self, label, img):
        # 获取 Label 的宽度和高度
        width, height = label.winfo_width(), label.winfo_height()

        if width > 0 and height > 0:  # 确保宽高有效
            # 按照 Label 尺寸调整图片大小
            resized_image = img.resize((width, height))
            tk_image = ImageTk.PhotoImage(resized_image)

            # 更新 Label 中的图片
            label.config(image=tk_image)
            label.image = tk_image  # 保留引用，防止图片被垃圾回收

class CMSPage(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)

        self.cms_datas = []

        style = ttk.Style()
        style.configure("Treeview.Heading", font=('標楷體', 14, 'bold'), foreground='black')
        style.map("Treeview.Heading", background=[('active', '#FFD700')])
        style.configure("Treeview.Heading", background='#FFD700')

        self.create_page()  # 初始化時創建佈局

        try:
            self.excel_cms = ExcelCMS('customers.xlsx')
            self.cms_datas = self.excel_cms.read_all_datas()
            if not self.cms_datas:
                print("Warning: No data found in Excel file.")
        except Exception as e:
            print(f"Error loading data: {e}")
            messagebox.showerror("錯誤", "無法讀取客戶資料")

        self.update_treeview()

    def create_page(self):
        # page Frame
        self.frame_container = tk.Frame(self)
        self.frame_container.pack(fill='both', expand=True)
        # title
        self.title = tk.Label(self.frame_container, text='客戶管理', font=('標楷體', 30, 'bold'), fg='#000000')
        self.title.place(relx=0.5, rely=0.08, relheight=0.1, relwidth=0.5, anchor='center')
        # button - back 
        self.btn_back = tk.Button(self.frame_container, text='回首頁', font=('標楷體', 16, 'bold'), fg='#000000', command=lambda: self.master.change_page(self.master.index_page))
        self.btn_back.place(relx=0.9, rely=0.1, relwidth=0.1, relheight=0.05, anchor='center')
        # treeview 
        self.tree_cms = ttk.Treeview(self.frame_container)
        self.tree_yscroll = tk.Scrollbar(self.tree_cms, orient="vertical")
        self.tree_yscroll.pack(side="right", fill="y")
        self.tree_yscroll.config(command=self.tree_cms.yview) 
        self.tree_cms.configure(yscrollcommand=self.tree_yscroll.set)
        self.tree_cms.place(relx=0.5, rely=0.4, relheight=0.5, relwidth=0.9, anchor='center')
        
        self.tree_cms['columns'] = ('公司名稱', '統編', '公司地址', '聯絡人', '聯絡電話')
        self.tree_cms.column('#0', width=0, stretch=False)  # 隱藏 #0 列
        for column in self.tree_cms['columns']:
            self.tree_cms.heading(column, text=column)
            self.tree_cms.column(column, anchor='center')
        
        # Entrys Custom
        self.label_custom = tk.Label(self.frame_container, text='公司名稱:', font=('標楷體', 14, 'bold'), fg='#000000')
        self.label_custom.place(relx=0.15, rely=0.75, relwidth=0.1, relheight=0.05, anchor='center')
        self.entry_custom = tk.Entry(self.frame_container, font=('標楷體', 14, 'bold'))
        self.entry_custom.place(relx=0.32, rely=0.75, relwidth=0.25, relheight=0.04, anchor='center')
        self.label_compiled = tk.Label(self.frame_container, text='統編:', font=('標楷體', 14, 'bold'), fg='#000000')
        self.label_compiled.place(relx=0.5, rely=0.75, relwidth=0.1, relheight=0.05, anchor='center')
        self.entry_compiled = tk.Entry(self.frame_container, font=('標楷體', 14, 'bold'))
        self.entry_compiled.place(relx=0.67, rely=0.75, relwidth=0.25, relheight=0.04, anchor='center')
        self.label_address = tk.Label(self.frame_container, text='公司地址:', font=('標楷體', 14, 'bold'), fg='#000000')
        self.label_address.place(relx=0.15, rely=0.8, relwidth=0.3, relheight=0.05, anchor='center')
        self.entry_address = tk.Entry(self.frame_container, font=('標楷體', 14, 'bold'))
        self.entry_address.place(relx=0.5, rely=0.8, relwidth=0.6, relheight=0.04, anchor='center')
        self.label_contact = tk.Label(self.frame_container, text='聯絡人:', font=('標楷體', 14, 'bold'), fg='#000000')
        self.label_contact.place(relx=0.15, rely=0.85, relwidth=0.1, relheight=0.05, anchor='center')
        self.entry_contact = tk.Entry(self.frame_container, font=('標楷體', 14, 'bold'))
        self.entry_contact.place(relx=0.32, rely=0.85, relwidth=0.25, relheight=0.04, anchor='center')
        self.label_phone = tk.Label(self.frame_container, text='聯絡電話:', font=('標楷體', 14, 'bold'), fg='#000000')
        self.label_phone.place(relx=0.5, rely=0.85, relwidth=0.1, relheight=0.05, anchor='center')
        self.entry_phone = tk.Entry(self.frame_container, font=('標楷體', 14, 'bold'))
        self.entry_phone.place(relx=0.67, rely=0.85, relwidth=0.25, relheight=0.04, anchor='center')
        # button - create
        self.btn_create = tk.Button(self.frame_container, text='新增', font=('標楷體', 14, 'bold'), fg='#000000', command=self.create_custom)
        self.btn_create.place(relx=0.8, rely=0.9, relwidth=0.08, relheight=0.04, anchor='center')
        # button - clear
        self.btn_clear = tk.Button(self.frame_container, text='清空', font=('標楷體', 14, 'bold'), fg='#000000', command=self.clear)
        self.btn_clear.place(relx=0.7, rely=0.9, relwidth=0.08, relheight=0.04, anchor='center')

    def update_treeview(self):
        # 先清空現有的內容
        for item in self.tree_cms.get_children():
            self.tree_cms.delete(item)
        
        # 插入資料到 Treeview
        for item in self.cms_datas:
            # print(f"Inserting item: {item}")  # 檢查item的內容是否正確
            self.tree_cms.insert('', 'end', values=item)

    def create_custom(self):
        if self.entry_custom.get()!='' and self.entry_compiled.get()!='' and self.entry_address.get()!='' and self.entry_contact.get()!='' and self.entry_phone.get()!='':
            custom = self.entry_custom.get()
            compiled = self.entry_compiled.get()
            address = self.entry_address.get()
            contact = self.entry_contact.get()
            phone = self.entry_phone.get()
            list_custom = [custom, compiled, address, contact, phone]
            self.excel_cms.create(*list_custom)
            # clear entrys
            self.entry_custom.delete(0, 'end')
            self.entry_compiled.delete(0, 'end')
            self.entry_address.delete(0, 'end')
            self.entry_contact.delete(0, 'end')
            self.entry_phone.delete(0, 'end')
            # update treeview
            self.cms_datas = self.excel_cms.read_all_datas()
            self.update_treeview()
        else:
            messagebox.showwarning(title='警告', message='客戶資料不能空白')
    
    def clear(self):
        self.entry_custom.delete(0, 'end')
        self.entry_compiled.delete(0, 'end')
        self.entry_address.delete(0, 'end')
        self.entry_contact.delete(0, 'end')
        self.entry_phone.delete(0, 'end')

class QuotePage(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)

        self.create_page()  # 初始化時創建佈局

    def create_page(self):
        # page Frame
        self.frame_container = tk.Frame(self)
        self.frame_container.pack(fill='both', expand=True)
        
        # title
        self.title = tk.Label(self.frame_container, text='報價單', font=('標楷體', 30, 'bold'), fg='#000000')
        self.title.place(relx=0.5, rely=0.1, relheight=0.1, relwidth=0.5, anchor='center')
        
        # buttons
        self.btn_back = tk.Button(self.frame_container, text='回首頁', font=('標楷體', 16, 'bold'), fg='#000000', command=lambda: self.master.change_page(self.master.index_page))
        self.btn_back.place(relx=0.9, rely=0.1, relwidth=0.1, relheight=0.05, anchor='center')

class ReportPage(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)

        self.create_page()  # 初始化時創建佈局

    def create_page(self):
        # page Frame
        self.frame_container = tk.Frame(self)
        self.frame_container.pack(fill='both', expand=True)
        
        # title
        self.title = tk.Label(self.frame_container, text='營收報表', font=('標楷體', 30, 'bold'), fg='#000000')
        self.title.place(relx=0.5, rely=0.1, relheight=0.1, relwidth=0.5, anchor='center')
        
        # buttons
        self.btn_back = tk.Button(self.frame_container, text='回首頁', font=('標楷體', 16, 'bold'), fg='#000000', command=lambda: self.master.change_page(self.master.index_page))
        self.btn_back.place(relx=0.9, rely=0.1, relwidth=0.1, relheight=0.05, anchor='center')


if __name__ == '__main__':
    app = App()
    app.mainloop()

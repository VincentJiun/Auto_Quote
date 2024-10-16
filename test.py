import tkinter as tk
from PIL import Image, ImageTk

class App(tk.Tk):
    def __init__(self):
        super().__init__()

        # Title, Icon, size
        self.title('德龍企業社')
        self.iconbitmap('./image/icon.ico')
        self.geometry('700x800')

        self.index_page = IndexPage(self)
        self.index_page.pack(fill='both', expand=True)
        self.index_page.create_index_page()


class IndexPage(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.resize_ids = {}  # 用於跟踪各個Label的after()事件ID

    def create_index_page(self):
        # index Frame
        self.frame_container = tk.Frame(self, background='Yellow')
        self.frame_container.pack(fill='both', expand=True)

        # 加載兩張圖片
        img1 = Image.open('./image/logo.png')
        img2 = Image.open('./image/logo.png')

        # 創建兩個不同的 Label，分別顯示不同的圖片
        self.label_img1 = tk.Label(self.frame_container)
        self.label_img1.place(relx=0.3, rely=0.3, relwidth=0.4, relheight=0.4, anchor='center')

        self.label_img2 = tk.Label(self.frame_container)
        self.label_img2.place(relx=0.7, rely=0.7, relwidth=0.4, relheight=0.4, anchor='center')

        # 將兩個圖片和 Label 傳遞到動態調整大小的功能中
        self.bind_resize_event(self.label_img1, img1)
        self.bind_resize_event(self.label_img2, img2)

    def bind_resize_event(self, label, image):
        # 綁定 <Configure> 事件，當窗口大小改變時觸發
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


if __name__ == '__main__':
    app = App()
    app.mainloop()

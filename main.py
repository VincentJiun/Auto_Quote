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

    def create_index_page(self):
        # index Frame
        self.frame_container = tk.Frame(self, background='Yellow')
        self.frame_container.pack(fill='both', expand=True)
        # index image
        self.img = Image.open('./image/logo.png')
        # self.logo_img = ImageTk.PhotoImage(self.img)
        self.label_img = tk.Label(self.frame_container)
        self.label_img.place(relx=0.5, rely=0.3, relwidth=0.5, relheight=0.5, anchor='center')
        self.label_img.bind("<Configure>", lambda event: self.image_on_configure(event, self.label_img, self.img))

    def image_on_configure(self, enevt, object, img):
        width, height = object.winfo_width(), object.winfo_height()
        resized_image = img.resize((width, height))
        tk_image = ImageTk.PhotoImage(resized_image)
        object.config(image=tk_image)    
        object.image = tk_image

        



if __name__ == '__main__':
    app = App()
    app.mainloop()



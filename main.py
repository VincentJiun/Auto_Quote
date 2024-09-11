import tkinter as tk

class App(tk.Tk):
    def __init__(self):
        super().__init__()

        # Title, Icon, size
        self.title('德龍企業社')
        self.iconbitmap('./image/icon.ico')
        self.geometry('700x800')



if __name__ == '__main__':
    app = App()
    app.mainloop()



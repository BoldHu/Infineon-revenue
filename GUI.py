import tkinter as tk

class main(object):
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Test")
        self.root.geometry("500x500")
        self.root.resizable(0, 0)

        self.label = tk.Label(self.root, text="Hello World")
        self.label.pack()

        self.root.mainloop()
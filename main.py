import tkinter as tk
from src.gui import WhatsAppSenderApp

if __name__ == "__main__":
    root = tk.Tk()
    app = WhatsAppSenderApp(root)
    root.mainloop()
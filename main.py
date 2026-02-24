import tkinter as tk
from gui import MailerGUI

def main():
    root = tk.Tk()
    app = MailerGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()

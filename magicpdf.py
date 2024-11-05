import datetime
import getpass
import logging
import os
import pathlib
import sys
import tkinter as tk
from tkinter import ttk, messagebox

import Merger, Extractor, Deleter


class MainWindow(tk.Tk):
    def __init__(self, filelist=[]):
        super().__init__()
        self.app_path = os.path.normpath(os.getcwd())

        self.title('Magic PDF')
        try:
            icon_img = tk.PhotoImage(file=os.path.normpath(os.path.join(self.app_path,'images/main_icon_32.png')))
            self.iconphoto(True, icon_img)
        except tk.TclError as err:
            logging.error(err)
        self.minsize(width=480, height=360)
        self.geometry('640x480')
        self.attributes('-alpha', 1)
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        
        self.create_menu()
        self.create_widgets(filelist)

    def create_menu(self):
        # Create a menubar
        menubar = tk.Menu(self)
        self.config(menu=menubar)
 
        # Create a File menu
        file_menu = tk.Menu(menubar, tearoff=False)
        # Create a submenu of File menu
        file_menu.add_command(label='Exit', command=self.destroy)

        # Create the Help menu
        help_menu = tk.Menu(menubar, tearoff=False)
        # Add the Help menu items to the menu
        help_menu.add_command(label='About...', command=self.show_about)
 
        # Add the File menu to the menubar
        menubar.add_cascade(label='File', menu=file_menu, underline=0)
        # Add the Help menu to the menubar
        menubar.add_cascade(label='Help', menu=help_menu, underline=0)

    def create_widgets(self, filelist=[]):
        notebook = ttk.Notebook()
        notebook.columnconfigure(0, weight=1)
        notebook.rowconfigure(0, weight=1)
        notebook.grid(column=0, row=0, sticky=tk.NSEW)
        if len(filelist) != 1:
            merger = Merger.Merger(notebook, filelist)
        else:
            merger = Merger.Merger(notebook)
        merger.pack(fill='both', expand=True)
        notebook.add(merger, text='Merge')
        if len(filelist) == 1:
            extractor = Extractor.Extractor(notebook, filelist[0])
        else:
            extractor = Extractor.Extractor(notebook)
        extractor.pack(fill='both', expand=True)
        notebook.add(extractor, text='Extract')
        if len(filelist) == 1:
            deleter = Deleter.Deleter(notebook, filelist[0])
        else:
            deleter = Deleter.Deleter(notebook)
        deleter.pack(fill='both', expand=True)
        notebook.add(deleter, text='Delete')
        if len(filelist) == 1:
            notebook.select(1)

    def show_about(self):
        about_message = 'MagicPDF ver. 0.6\n\nDesign and development by Yevhen E.\n\nUsing the pypdf library\n\n\u2764\ufe0f For Dashuta Funtik \u2764\ufe0f'
        messagebox.showinfo(title='About...',
                            message=about_message)


if __name__ == '__main__':
    curr_datetime = datetime.datetime.today().strftime('%Y-%m-%d')
    LOG_DIR = os.path.join(pathlib.Path.home(), 'magicpdf', 'logs')
    LOG_FILENAME = '{}.log'.format(curr_datetime)
    if not os.path.exists(LOG_DIR):
        os.makedirs(LOG_DIR)
    LOG_PATH = os.path.join(LOG_DIR, LOG_FILENAME)
    FORMAT = '%(asctime)s %(levelname)s:{}: %(message)s'.format(getpass.getuser())
    logging.basicConfig(
        format=FORMAT,
        datefmt='%Y-%m-%d %H:%M:%S',
        filename=LOG_PATH,
        level=logging.DEBUG,
    )
    logging.info('****Start Program****')

    file_list = []
    if len(sys.argv) > 1:
        for file in sys.argv[1:]:
            if file.lower().endswith('.pdf'):
                file_list.append(file)
    app = MainWindow(file_list)
    app.mainloop()
    logging.info('****End Program****\n')


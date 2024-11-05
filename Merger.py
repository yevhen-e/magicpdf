import datetime
import logging
import os
import shutil
import tempfile
import threading
import time
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pypdf import PdfWriter


class PdfMergerThread(threading.Thread):
    def __init__(self, in_file_list, result_file_path, is_outlines):
        super().__init__()
        self.in_files_list = in_file_list
        self.result_file_path = result_file_path
        self.is_outlines = is_outlines
        self.warning_message = ''

    def run(self):
        tmp_file_path = None
        pdf_writer = PdfWriter()
        try:
            logging.info('****Beginning merging session...****')
            logging.info('Start appending...')
            start_time = time.time()
            for file_path in self.in_files_list:
                if not os.path.exists(path=file_path):
                    logging.warning('The file {} does not exist. The merging was not completed!'.format(file_path))
                    logging.info('Stop appending')
                    self.set_message('The file {} does not exist.\nThe merging was not completed!'.format(file_path))
                    return
                if self.is_outlines:
                    pdf_writer.append(file_path, os.path.splitext(os.path.basename(file_path))[0])
                else:
                    pdf_writer.append(file_path)

            stop_time = time.time()
            logging.info('Stop appending')
            logging.info('Append time: {}'.format(stop_time - start_time))

            tmp_file = tempfile.TemporaryFile()
            tmp_file_path = os.path.join(tempfile.gettempdir(), str(tmp_file.name))
            tmp_file.close()

            logging.info('Start writing...')
            start_time = time.time()
            pdf_writer.write(tmp_file_path)

            stop_time = time.time()
            logging.info('Stop writing')
            logging.info('Write time: {}'.format(stop_time - start_time))

            logging.info('Start copying file...')
            start_time = time.time()
            shutil.copyfile(src=tmp_file_path, dst=self.result_file_path)
            stop_time = time.time()
            logging.info('Stop copying file')
            logging.info('Copying time: {}'.format(stop_time - start_time))
        except Exception as ex:
            logging.error(ex)
            self.set_message('Something went wrong...')
        finally:
            pdf_writer.close()
            try:
                if os.path.exists(tmp_file_path):
                    os.remove(tmp_file_path)
            except Exception as ex:
                logging.error(ex)
            logging.info('****End merging session****')

    def set_message(self, message):
        self.warning_message = message

    def get_message(self):
        return self.warning_message


class Merger(ttk.Frame):
    def __init__(self, container, filelist=[]):
        super().__init__(container)
        self.is_outlines = None
        self.clear_items_button = None
        self.del_items_button = None
        self.add_items_button = None
        self.files_listbox = None

        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        self.listbox_items = []
        if filelist:
            self.listbox_items = filelist
            self.listbox_items.sort()

        self.create_widgets()

    def create_widgets(self):
        merger_frame = ttk.Frame(self)

        merger_frame.columnconfigure(0, weight=1)

        self.files_listbox = tk.Listbox(merger_frame, relief=tk.FLAT, selectmode=tk.EXTENDED)
        self.files_listbox.delete(0, tk.END)
        if self.listbox_items:
            self.files_listbox.insert(tk.END, *[os.path.basename(item) for item in self.listbox_items])
            self.files_listbox.focus()
            self.files_listbox.select_set(0)
            self.files_listbox.selection_anchor(0)
            self.files_listbox.activate(0)
        self.files_listbox.grid(column=0, row=0, rowspan=8, sticky=tk.NSEW, padx=(5, 0), pady=5)

        self.add_items_button = ttk.Button(merger_frame, text='Add', command=self.add_items_to_listbox)
        self.add_items_button.grid(column=1, row=0, sticky=tk.EW, padx=5, pady=5)

        self.del_items_button = ttk.Button(merger_frame, text='Delete', command=self.del_items_from_listbox)
        self.del_items_button.grid(column=1, row=1, sticky=tk.EW, padx=5)

        self.clear_items_button = ttk.Button(merger_frame, text='Clear', command=self.del_all_items_from_listbox)
        self.clear_items_button.grid(column=1, row=2, sticky=tk.EW, padx=5, pady=5)

        self.is_outlines = tk.BooleanVar(value=True)
        self.outlines_checkbox = ttk.Checkbutton(merger_frame, text='New Bookmarks', variable=self.is_outlines)
        self.outlines_checkbox.grid(column=1, row=3, sticky=tk.NW, padx=5, pady=5)

        merger_frame.rowconfigure(3, weight=1)

        self.move_up_button = ttk.Button(merger_frame, text='Up', command=self.move_up_listbox_item)
        self.move_up_button.grid(column=1, row=4, padx=5, sticky=tk.EW)

        self.move_down_button = ttk.Button(merger_frame, text='Down', command=self.move_down_listbox_item)
        self.move_down_button.grid(column=1, row=5, padx=5, pady=5, sticky=tk.EW)

        self.move_top_button = ttk.Button(merger_frame, text='Top', command=self.move_top_listbox_item)
        self.move_top_button.grid(column=1, row=6, padx=5, sticky=tk.EW)

        self.move_bottom_button = ttk.Button(merger_frame, text='Bottom', command=self.move_bottom_listbox_item)
        self.move_bottom_button.grid(column=1, row=7, padx=5, pady=5, sticky=tk.EW)

        separator_top = ttk.Separator(merger_frame, orient='horizontal')
        separator_top.grid(columnspan=2, column=0, row=8, sticky=tk.EW, pady=(0, 5), padx=5)

        merge_pbar_empty_frame = ttk.Frame(merger_frame)
        merge_pbar_empty_frame.columnconfigure(0, weight=1)

        empty_label = ttk.Label(merge_pbar_empty_frame)
        empty_label.grid(column=0, row=0, sticky=tk.EW)

        merge_pbar_empty_frame.grid(column=0, row=9, sticky=tk.EW, padx=5, pady=0)

        self.merge_pbar_frame = ttk.Frame(merger_frame)
        self.merge_pbar_frame.columnconfigure(1, weight=1)

        merge_label_status = ttk.Label(self.merge_pbar_frame, text="Please wait...")
        merge_label_status.grid(column=0, row=0, sticky=tk.W, padx=5)
        self.merge_pbar = ttk.Progressbar(self.merge_pbar_frame, orient=tk.HORIZONTAL, mode='indeterminate')
        self.merge_pbar.grid(column=1, row=0, sticky=tk.EW, padx=0)

        self.merge_button = ttk.Button(merger_frame, text='Merge', command=self.merge_files)
        self.merge_button.grid(column=1, row=9, sticky=tk.EW, padx=5)

        separator_bottom = ttk.Separator(self, orient='horizontal')
        separator_bottom.grid(columnspan=2, column=0, row=10, sticky=tk.EW, pady=(5, 5), padx=5)

        merger_frame.grid(column=0, row=0, sticky=tk.NSEW)

    def move_up_listbox_item(self):
        if not self.listbox_items or not self.files_listbox.curselection():
            messagebox.showinfo(title='Information', message='Nothing to move...')
            return
        indexes = self.files_listbox.curselection()
        if not indexes or self.files_listbox.selection_includes(0):
            return
        for pos in indexes:
            text = self.files_listbox.get(pos)
            self.files_listbox.delete(pos)
            self.files_listbox.insert(pos - 1, text)
            text_item = self.listbox_items[pos]
            self.listbox_items.pop(pos)
            self.listbox_items.insert(pos - 1, text_item)
        for index in indexes:
            self.files_listbox.select_set(index - 1)

    def move_down_listbox_item(self):
        if not self.listbox_items or not self.files_listbox.curselection():
            messagebox.showinfo(title='Information', message='Nothing to move...')
            return
        indexes = self.files_listbox.curselection()
        if not indexes or self.files_listbox.selection_includes(tk.END):
            return
        for pos in indexes[::-1]:
            text = self.files_listbox.get(pos)
            self.files_listbox.delete(pos)
            self.files_listbox.insert(pos + 1, text)
            text_item = self.listbox_items[pos]
            self.listbox_items.pop(pos)
            self.listbox_items.insert(pos + 1, text_item)
        for index in indexes:
            self.files_listbox.select_set(index + 1)

    def move_top_listbox_item(self):
        if not self.listbox_items or not self.files_listbox.curselection():
            messagebox.showinfo(title='Information', message='Nothing to move...')
            return
        indexes = self.files_listbox.curselection()
        if not indexes or self.files_listbox.selection_includes(0):
            return
        delta_index = indexes[0:1:][0]
        for pos in indexes:
            text = self.files_listbox.get(pos)
            self.files_listbox.delete(pos)
            self.files_listbox.insert(pos - delta_index, text)
            text_item = self.listbox_items[pos]
            self.listbox_items.pop(pos)
            self.listbox_items.insert(pos - delta_index, text_item)
        for index in indexes:
            self.files_listbox.select_set(index - delta_index)

    def move_bottom_listbox_item(self):
        if not self.listbox_items or not self.files_listbox.curselection():
            messagebox.showinfo(title='Information', message='Nothing to move...')
            return
        indexes = self.files_listbox.curselection()
        if not indexes or self.files_listbox.selection_includes(tk.END):
            return
        delta_index = (len(self.listbox_items) - 1 - indexes[::-1][0])
        for pos in indexes[::-1]:
            text = self.files_listbox.get(pos)
            self.files_listbox.delete(pos)
            self.files_listbox.insert(pos + delta_index, text)
            text_item = self.listbox_items[pos]
            self.listbox_items.pop(pos)
            self.listbox_items.insert(pos + delta_index, text_item)
        for index in indexes:
            self.files_listbox.select_set(index + delta_index)

    def merge_files(self):
        if not self.listbox_items or len(self.listbox_items) <= 1:
            messagebox.showinfo(title='Information', message='Nothing to merge...')
            return
        curr_datetime = datetime.datetime.today().strftime('%Y%m%d%H%M%S')
        initial_dir = os.path.dirname(self.listbox_items[0])
        output_path = filedialog.asksaveasfilename(title='Save As...',
                                                   filetypes=(('PDF Files', '*.pdf'), ('All Files', '*.*')),
                                                   initialdir=initial_dir,
                                                   initialfile='result_{}'.format(curr_datetime),
                                                   defaultextension='.pdf')
        try:
            if output_path:
                if repr(os.path.normpath(output_path)) in repr(self.listbox_items)[:]:
                    messagebox.showinfo(title='Information', message='The file cannot be written to itself...')
                    return
                self.start_merge()
                merger_thread = PdfMergerThread(in_file_list=self.listbox_items, result_file_path=output_path,
                                                is_outlines=self.is_outlines.get())
                merger_thread.start()
                self.merger_thread_monitor(merger_thread, in_file_list=self.listbox_items, result_file_path=output_path,
                                           is_outlines=self.is_outlines.get())
        except Exception as ex:
            logging.error(ex)
            self.stop_merge()
            messagebox.showwarning(title='Warning!', message='Something went wrong...')

    def merger_thread_monitor(self, thread, in_file_list, result_file_path, is_outlines):
        if thread.is_alive():
            self.after(100, lambda: self.merger_thread_monitor(thread=thread, in_file_list=in_file_list,
                                                               result_file_path=result_file_path,
                                                               is_outlines=is_outlines))
        else:
            self.stop_merge()
            if thread.get_message():
                messagebox.showwarning(title='Warning!', message=thread.get_message())

    def start_merge(self):
        self.merge_button['state'] = tk.DISABLED
        self.add_items_button['state'] = tk.DISABLED
        self.del_items_button['state'] = tk.DISABLED
        self.clear_items_button['state'] = tk.DISABLED
        self.outlines_checkbox['state'] = tk.DISABLED
        self.move_up_button['state'] = tk.DISABLED
        self.move_down_button['state'] = tk.DISABLED
        self.move_top_button['state'] = tk.DISABLED
        self.move_bottom_button['state'] = tk.DISABLED

        self.merge_pbar_frame.grid(column=0, row=9, sticky=tk.EW, padx=0)
        self.merge_pbar.start(10)

    def stop_merge(self):
        self.merge_button['state'] = tk.NORMAL
        self.add_items_button['state'] = tk.NORMAL
        self.del_items_button['state'] = tk.NORMAL
        self.clear_items_button['state'] = tk.NORMAL
        self.outlines_checkbox['state'] = tk.NORMAL
        self.move_up_button['state'] = tk.NORMAL
        self.move_down_button['state'] = tk.NORMAL
        self.move_top_button['state'] = tk.NORMAL
        self.move_bottom_button['state'] = tk.NORMAL

        self.merge_pbar.stop()
        self.merge_pbar_frame.grid_remove()

    def listbox_filling(self):
        files = list(
            filedialog.askopenfilenames(title='Open File(s)', filetypes=(('PDF Files', '*.pdf'),)))
        if files:
            if self.listbox_items:
                self.del_all_items_from_listbox()
            for file in files:
                if file.lower().endswith('.pdf'):
                    self.files_listbox.insert(tk.END, os.path.basename(file))
                    self.listbox_items.append(os.path.normpath(file))
            if self.listbox_items: self.files_listbox.select_set(0)

    def add_items_to_listbox(self):
        files = list(
            filedialog.askopenfilenames(title='Open File(s)', filetypes=(('PDF Files', '*.pdf'),)))
        if files:
            for file in files:
                if file.lower().endswith('.pdf'):
                    if os.path.basename(file) in set(self.files_listbox.get(0, tk.END)):
                        ask_message = 'The file "{}" is exist in the list\n\nDo you want to add it?'.format(os.path.basename(file))
                        if not messagebox.askyesno(title='Information',
                                                   message=ask_message):
                            continue
                    self.files_listbox.insert(tk.END, os.path.basename(file))
                    self.listbox_items.append(os.path.normpath(file))

    def del_items_from_listbox(self):
        indexes = self.files_listbox.curselection()
        if not indexes:
            messagebox.showinfo(title='Information', message='Nothing to delete...')
            return
        for index in indexes[::-1]:
            self.files_listbox.delete(index)
            self.listbox_items.pop(index)

    def del_all_items_from_listbox(self):
        if not self.listbox_items:
            messagebox.showinfo(title='Information', message='Nothing to delete...')
            return
        self.files_listbox.delete(0, tk.END)
        self.listbox_items.clear()

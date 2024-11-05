import datetime
import logging
import os
import shutil
import tempfile
import threading
import time
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from pypdf import PdfWriter, PdfReader


class RangeDeleteThread(threading.Thread):
    def __init__(self, pdf_reader, output_path, page_range, status_label=None):
        super().__init__()
        self.pdf_reader = pdf_reader
        self.output_path = output_path
        self.page_range = page_range
        self.status_label = status_label
        self.warning_message = ''

    def run(self):
        tmp_file_path = None
        pdf_writer = PdfWriter()
        try:
            logging.info('**** The page range deleting session is started... ****')
            logging.info('Deleting pages is begun...')
            start_time = time.time()
            i = 0
            for page_numder in range(len(self.pdf_reader.pages)):
                if page_numder + 1 not in self.page_range:
                    pdf_writer.add_page(self.pdf_reader.pages[page_numder])
                else:
                    if self.status_label:
                        i += 1
                        self.status_label['text'] = 'Deleting page progress: deleted {} of {}...'.format(i, len(self.page_range))

            stop_time = time.time()
            logging.info('Deleting pages is finished')
            logging.info('Deleting pages time: {}'.format(stop_time - start_time))

            tmp_file = tempfile.TemporaryFile()
            tmp_file_path = os.path.join(tempfile.gettempdir(), str(tmp_file.name))
            tmp_file.close()

            logging.info('The writing to the temporary file is started...')
            start_time = time.time()
            if self.status_label:
                self.status_label['text'] = 'Deleting page progress: writing result file...'

            pdf_writer.write(tmp_file_path)

            stop_time = time.time()
            logging.info('The writing to the temporary file is finished')
            logging.info('Writing to tmp file time: {}'.format(stop_time - start_time))

            logging.info('Starting copying file...')
            start_time = time.time()
            shutil.copyfile(src=tmp_file_path, dst=self.output_path)
            stop_time = time.time()
            logging.info('The copying file is finished')
            logging.info('Copying time: {}'.format(stop_time - start_time))

        except Exception as ex:
            logging.error(ex)
            self.set_message(str(ex))
            return
        finally:
            pdf_writer.close()
            try:
                if os.path.exists(tmp_file_path):
                    os.remove(tmp_file_path)
            except Exception as ex:
                logging.error(ex)

        logging.info('**** The page range deleting session is finished ****')

    def set_message(self, message):
        self.warning_message = message

    def get_message(self):
        return self.warning_message


class OpenPDFFileThread(threading.Thread):
    def __init__(self, source_file_path):
        super().__init__()
        self.source_file_path = source_file_path
        self.pdf_reader = None
        self.warning_message = ''

    def run(self):
        logging.info('**** Starting the loading pdf file session... ****')
        start_time = time.time()
        logging.info('Start loading...')
        try:
            pdf_reader = PdfReader(self.source_file_path)
            self.set_pdf_reader(pdf_reader)
        except Exception as ex:
            logging.error(ex)
            self.set_message(str(ex))
        stop_time = time.time()
        logging.info('Loading is finished')
        logging.info('Loading time: {}'.format(stop_time - start_time))
        logging.info('**** The loading pdf file session is finished ****')

    def set_pdf_reader(self, pdf_reader):
        self.pdf_reader = pdf_reader

    def get_pdf_reader(self):
        return self.pdf_reader

    def set_message(self, message):
        self.warning_message = message

    def get_message(self):
        return self.warning_message


class Deleter(ttk.Frame):
    def __init__(self, container, input_file=''):
        super().__init__(container)
        self.deleting_page_progress = None
        self.pages_number_title = None
        self.pdf_reader = None
        self.page_range_example_title = None
        self.source_file_path = None
        self.page_range_entry = None
        self.delete_button = None
        self.delete_pbar = None
        self.delete_pbar_frame = None
        self.input_file_name = None
        self.open_file_button = None

        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        self.create_widgets()
        if input_file:
            self.open_file(input_file)

    def create_widgets(self):
        deleter_frame = ttk.Frame(self)

        deleter_frame.columnconfigure(0, weight=1)

        top_sep = ttk.Separator(deleter_frame, orient='horizontal')
        top_sep.grid(columnspan=2, column=0, row=0, sticky=tk.EW, pady=(5, 5), padx=(5, 5))

        # Create Open file Frame
        file_frame = ttk.Frame(deleter_frame)
        file_frame.columnconfigure(1, weight=1)
        file_frame.rowconfigure(0, weight=1)
        file_frame.rowconfigure(1, weight=1)

        input_file_title = ttk.Label(file_frame, text='Source file:')
        input_file_title.grid(column=0, row=0, sticky=tk.E, padx=(5, 0), pady=(0, 0))

        self.input_file_name = ttk.Label(file_frame)
        self.input_file_name.grid(column=1, row=0, sticky=tk.EW, padx=(5, 5), pady=(0, 0))

        self.pages_number_title = ttk.Label(file_frame, text='')
        self.pages_number_title.grid(columnspan=2, column=0, row=1, sticky=tk.W, padx=(5, 0), pady=(0, 0))

        file_frame.grid(column=0, row=1, sticky=tk.EW, padx=(5, 5))

        # Create Open file Button
        self.open_file_button = ttk.Button(deleter_frame, text='Open File', command=self.open_file)
        self.open_file_button.grid(rowspan=2, column=1, row=1, sticky=tk.NW, padx=5)

        delete_top_sep = ttk.Separator(deleter_frame, orient='horizontal')
        delete_top_sep.grid(columnspan=2, column=0, row=2, sticky=tk.EW, pady=(5, 5), padx=(5, 5))

        # Create delete page range Frame
        page_range_frame = ttk.Frame(deleter_frame)

        # Extraction type
        page_range_title = ttk.Label(page_range_frame, text='Page range:')
        page_range_title.grid(column=0, row=0, sticky=tk.E, padx=(5, 0), pady=(0, 0))

        self.page_range_entry = ttk.Entry(page_range_frame)
        self.page_range_entry.grid(column=1, row=0, sticky=tk.EW, padx=(5, 5), pady=(0, 0))

        self.page_range_example_title = ttk.Label(page_range_frame, text='e.g. 3-7,9,14-17', font=('', 7))
        self.page_range_example_title.grid(column=1, row=1, sticky=tk.EW, padx=(5, 5), pady=(0, 0))

        self.deleting_page_progress = ttk.Label(page_range_frame, text='')

        page_range_frame.columnconfigure(1, weight=1)

        page_range_frame.grid(column=0, row=3, sticky=tk.EW, padx=(5, 5), pady=(15, 0))

        deleter_frame.rowconfigure(5, weight=1)

        delete_pbar_top_sep = ttk.Separator(deleter_frame, orient='horizontal')
        delete_pbar_top_sep.grid(columnspan=2, column=0, row=6, sticky=tk.EW, pady=(0, 5), padx=5)

        # Create Progressbar Frame
        pbar_frame = ttk.Frame(deleter_frame)
        delete_pbar_empty_frame = ttk.Frame(pbar_frame)
        delete_pbar_empty_frame.columnconfigure(0, weight=1)

        empty_label = ttk.Label(delete_pbar_empty_frame)
        empty_label.grid(column=0, row=0, sticky=tk.EW)

        delete_pbar_empty_frame.grid(column=0, row=0, sticky=tk.EW, padx=5, pady=0)

        self.delete_pbar_frame = ttk.Frame(pbar_frame)
        self.delete_pbar_frame.columnconfigure(1, weight=1)

        delete_label_status = ttk.Label(self.delete_pbar_frame, text="Please wait...")
        delete_label_status.grid(column=0, row=0, sticky=tk.W, padx=5)

        self.delete_pbar = ttk.Progressbar(self.delete_pbar_frame, orient=tk.HORIZONTAL, mode='indeterminate')
        self.delete_pbar.grid(column=1, row=0, sticky=tk.EW, padx=0)

        pbar_frame.columnconfigure(0, weight=1)
        pbar_frame.grid(column=0, row=7, sticky=tk.EW, padx=5)

        self.delete_button = ttk.Button(deleter_frame, text='Delete', command=self.delete_pages)
        self.delete_button.grid(column=1, row=7, sticky=tk.EW, padx=5)

        separator_bottom = ttk.Separator(self, orient='horizontal')
        separator_bottom.grid(columnspan=2, column=0, row=8, sticky=tk.EW, pady=(5, 5), padx=5)

        deleter_frame.grid(column=0, row=0, sticky=tk.NSEW)

    def delete_pages(self):
        message_nothing_to_do = 'Nothing to do...\nPlease choose a Source File'
        if not self.input_file_name['text']:
            tk.messagebox.showinfo('Information...', message_nothing_to_do)
            return
        output_file_name = self.update_output_file_name(self.source_file_path)
        output_path = os.path.dirname(self.source_file_path)
        output_path = filedialog.asksaveasfilename(title='Save As...',
                                                   filetypes=(('PDF Files', '*.pdf'),),
                                                   initialdir=output_path,
                                                   initialfile=output_file_name,
                                                   defaultextension='.pdf')
        try:
            if output_path:
                pages_range = self.parse_pages_range(self.page_range_entry.get())
                if max(pages_range) > len(self.pdf_reader.pages):
                    messagebox.showwarning(title='Warning!', message='Page range is not correct!')
                    return
                self.start_thread()
                range_delete_thread = RangeDeleteThread(self.pdf_reader, output_path, pages_range,
                                                        self.deleting_page_progress)
                range_delete_thread.start()
                self.range_delete_thread_monitor(thread=range_delete_thread, pdf_reader=self.pdf_reader,
                                                 output_path=output_path, pages_range=pages_range)
        except ValueError as ex:
            logging.error(ex)
            messagebox.showwarning(title='Warning!', message='Invalid range format!\n{}'.format(str(ex)))
            return
        except Exception as ex:
            logging.error(ex)
            self.stop_thread()
            self.deleting_page_progress.grid_remove()
            messagebox.showwarning(title='Warning!', message='Something went wrong...')

    def range_delete_thread_monitor(self, thread, pdf_reader, output_path, pages_range):
        self.deleting_page_progress.grid(columnspan=3, column=0, row=2, sticky=tk.EW, padx=(5, 5), pady=(10, 0))
        if thread.is_alive():
            self.after(100, lambda: self.range_delete_thread_monitor(thread=thread, pdf_reader=pdf_reader,
                                                                     output_path=output_path, pages_range=pages_range))
        else:
            self.deleting_page_progress.grid_remove()
            self.stop_thread()
            if thread.get_message():
                messagebox.showwarning(title='Warning!', message=thread.get_message())

    def open_pdf_file_thread_monitor(self, thread, source_file_path):
        if thread.is_alive():
            self.after(100, lambda: self.open_pdf_file_thread_monitor(thread=thread, source_file_path=source_file_path))
        else:
            self.stop_thread()
            self.pdf_reader = thread.get_pdf_reader()
            self.pages_number_title['text'] = 'The number of pages is {}'.format(len(self.pdf_reader.pages))
            self.input_file_name['text'] = os.path.basename(self.source_file_path)
            if thread.get_message():
                messagebox.showwarning(title='Warning!', message=thread.get_message())

    def start_thread(self):
        self.delete_button['state'] = tk.DISABLED
        self.open_file_button['state'] = tk.DISABLED
        self.page_range_entry['state'] = tk.DISABLED
        self.delete_pbar_frame.grid(column=0, row=0, sticky=tk.EW, padx=0)
        self.delete_pbar.start(10)

    def stop_thread(self):
        self.delete_button['state'] = tk.NORMAL
        self.open_file_button['state'] = tk.NORMAL
        self.page_range_entry['state'] = tk.NORMAL
        self.delete_pbar.stop()
        self.delete_pbar_frame.grid_remove()

    def open_file(self, file_path=''):
        if not file_path:
            file_path = filedialog.askopenfilename(title='Open File', filetypes=(('PDF Files', '*.pdf'),))
        if not file_path:
            return
        self.source_file_path = os.path.normpath(file_path)
        self.page_range_entry.delete(0, tk.END)
        self.pages_number_title['text'] = ''
        self.input_file_name['text'] = ''
        if self.pdf_reader:
            self.pdf_reader.stream.close()
        try:
            self.start_thread()
            open_pdf_file_thread = OpenPDFFileThread(self.source_file_path)
            open_pdf_file_thread.start()
            self.open_pdf_file_thread_monitor(open_pdf_file_thread, self.source_file_path)
        except Exception as ex:
            logging.error(ex)
            self.stop_thread()
            messagebox.showwarning(title='Warning!', message=str(ex))
            return

    def update_output_file_name(self, source_file):
        curr_datetime = datetime.datetime.today().strftime('%Y%m%d%H%M%S')
        source_file_name, source_file_extension = os.path.splitext(os.path.basename(source_file))
        file_name = '{}_del_result_{}{}'.format(source_file_name, curr_datetime, source_file_extension)
        return file_name

    def parse_pages_range(self, parse_string):
        result = []
        for part in parse_string.split(','):
            x = part.split('-')
            result.extend(range(int(x[0]), int(x[-1]) + 1))
        return result

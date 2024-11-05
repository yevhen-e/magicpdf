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


class PbPExtractThread(threading.Thread):
    def __init__(self, source_file_path, pdf_reader, output_path, output_dir_name, status_label=None):
        super().__init__()
        self.source_file_path = source_file_path
        self.pdf_reader = pdf_reader
        self.output_path = output_path
        self.output_dir_name = output_dir_name
        self.status_label = status_label
        self.warning_message = ''

    def run(self):
        logging.info('**** The page by page extraction session is started... ****')
        try:
            logging.info('Creating directory is started...')
            os.makedirs(os.path.join(self.output_path, self.output_dir_name))
            logging.info(
                'The directory \"{}\" was created'.format(os.path.join(self.output_path, self.output_dir_name)))
        except Exception as ex:
            logging.error(ex)
            self.set_message(str(ex))
            return
        output_filename = os.path.splitext(os.path.basename(self.source_file_path))[0]
        logging.info('Extraction pages is begun...')
        start_time = time.time()
        page_count = len(self.pdf_reader.pages)
        tmp_file_path = None
        for i, page in enumerate(self.pdf_reader.pages):
            pdf_writer = PdfWriter()
            try:
                tmp_file = tempfile.TemporaryFile()
                tmp_file_path = os.path.join(tempfile.gettempdir(), str(tmp_file.name))
                tmp_file.close()

                if self.status_label:
                    self.status_label['text'] = 'Extracting page progress: extracted {} of {}...'.format(i+1, page_count)
                pdf_writer.add_page(page)
                pdf_writer.write(tmp_file_path )

                shutil.copyfile(src=tmp_file_path, dst=os.path.join(self.output_path,
                                                                    self.output_dir_name,
                                                                    "Page {} - {}.pdf".format(i + 1, output_filename)))

            except Exception as ex:
                logging.error(ex)
                self.set_message(str(ex))
                try:
                    logging.info('Starting a deleting directory...')
                    shutil.rmtree(self.output_path)
                    logging.info('The directory \"{}\" was deleted'.format(self.output_path))
                except Exception as ex:
                    logging.error(ex)
                    self.set_message(str(ex))
            finally:
                pdf_writer.close()
                try:
                    if os.path.exists(tmp_file_path):
                        os.remove(tmp_file_path)
                except Exception as ex:
                    logging.error(ex)

        stop_time = time.time()
        logging.info('Extraction pages is finished')
        logging.info('Extraction pages time: {}'.format(stop_time - start_time))
        logging.info('**** The page by page extraction session is finished ****')

    def set_message(self, message):
        self.warning_message = message

    def get_message(self):
        return self.warning_message


class RangeExtractThread(threading.Thread):
    def __init__(self, pdf_reader, output_path, pages_range, status_label=None):
        super().__init__()
        self.pdf_reader = pdf_reader
        self.output_path = output_path
        self.pages_range = pages_range
        self.status_label = status_label
        self.warning_message = ''

    def run(self):
        tmp_file_path = None
        pdf_writer = PdfWriter()
        try:
            logging.info('**** The page range extraction session is started... ****')
            logging.info('Extraction pages is begun...')
            start_time = time.time()
            page_count = len(self.pages_range)
            for i, page_number in enumerate(self.pages_range):
                pdf_writer.add_page(self.pdf_reader.pages[page_number - 1])
                if self.status_label:
                    self.status_label['text'] = 'Extracting page progress: extracted {} of {}...'.format(i+1, page_count)

            stop_time = time.time()
            logging.info('Extraction pages is finished')
            logging.info('Extraction pages time: {}'.format(stop_time - start_time))

            tmp_file = tempfile.TemporaryFile()
            tmp_file_path = os.path.join(tempfile.gettempdir(), str(tmp_file.name))
            tmp_file.close()

            logging.info('The writing to the temporary file is started...')
            start_time = time.time()
            if self.status_label:
                self.status_label['text'] = 'Extracting page progress: writing result file...'

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
        logging.info('**** The page range extraction session is finished ****')

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


class Extractor(ttk.Frame):
    def __init__(self, container, input_file=''):
        super().__init__(container)
        self.extracting_page_progress = None
        self.pages_number_title = None
        self.pdf_reader = None
        self.page_range_example_title = None
        self.source_file_path = None
        self.page_range_entry = None
        self.extr_type_combobox_values = None
        self.extr_type_combobox = None
        self.extract_button = None
        self.extract_pbar = None
        self.extract_pbar_frame = None
        self.input_file_name = None
        self.open_file_button = None

        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        self.create_widgets()
        if input_file:
            self.open_file(input_file)

    def create_widgets(self):
        extractor_frame = ttk.Frame(self)

        extractor_frame.columnconfigure(0, weight=1)

        top_sep = ttk.Separator(extractor_frame, orient='horizontal')
        top_sep.grid(columnspan=2, column=0, row=0, sticky=tk.EW, pady=(5, 5), padx=(5, 5))

        # Create Open file Frame
        file_frame = ttk.Frame(extractor_frame)
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
        self.open_file_button = ttk.Button(extractor_frame, text='Open File', command=self.open_file)
        self.open_file_button.grid(rowspan=2, column=1, row=1, sticky=tk.NW, padx=5)

        extr_type_top_sep = ttk.Separator(extractor_frame, orient='horizontal')
        extr_type_top_sep.grid(columnspan=2, column=0, row=2, sticky=tk.EW, pady=(5, 5), padx=(5, 5))

        # Create Extraction Type Frame
        extr_type_frame = ttk.Frame(extractor_frame)

        # Extraction type
        extr_type_title = ttk.Label(extr_type_frame, text='Extraction type:')
        extr_type_title.grid(column=0, row=0, sticky=tk.E, padx=(5, 0), pady=(0, 0))

        self.extr_type_combobox_values = ('Page by Page', 'Pages range')
        self.extr_type_combobox = ttk.Combobox(extr_type_frame)
        self.extr_type_combobox['values'] = self.extr_type_combobox_values
        self.extr_type_combobox['state'] = 'readonly'
        self.extr_type_combobox.current(0)

        def extr_type_combobox_change_item(event):
            self.page_range_entry.grid_remove()
            self.page_range_example_title.grid_remove()
            if self.extr_type_combobox.get() == self.extr_type_combobox_values[1]:
                self.page_range_entry.grid(column=2, row=0, sticky=tk.EW, padx=(5, 5), pady=(0, 0))
                self.page_range_example_title.grid(column=2, row=1, sticky=tk.EW, padx=(5, 5), pady=(0, 0))

        self.extr_type_combobox.bind('<<ComboboxSelected>>', extr_type_combobox_change_item)
        self.extr_type_combobox.grid(column=1, row=0, sticky=tk.EW, padx=(5, 0), pady=(0, 0))

        self.page_range_entry = ttk.Entry(extr_type_frame)

        self.page_range_example_title = ttk.Label(extr_type_frame, text='e.g. 3-7,9,14-17', font=('', 7))

        self.extracting_page_progress = ttk.Label(extr_type_frame)

        extr_type_frame.columnconfigure(2, weight=1)
        extr_type_frame.grid(column=0, row=3, sticky=tk.EW, padx=(5, 5), pady=(15, 0))

        extractor_frame.rowconfigure(5, weight=1)

        extract_pbar_top_sep = ttk.Separator(extractor_frame, orient='horizontal')
        extract_pbar_top_sep.grid(columnspan=2, column=0, row=6, sticky=tk.EW, pady=(0, 5), padx=5)

        # Create Progressbar Frame
        pbar_frame = ttk.Frame(extractor_frame)
        extract_pbar_empty_frame = ttk.Frame(pbar_frame)
        extract_pbar_empty_frame.columnconfigure(0, weight=1)

        empty_label = ttk.Label(extract_pbar_empty_frame)
        empty_label.grid(column=0, row=0, sticky=tk.EW)

        extract_pbar_empty_frame.grid(column=0, row=0, sticky=tk.EW, padx=5, pady=0)

        self.extract_pbar_frame = ttk.Frame(pbar_frame)
        self.extract_pbar_frame.columnconfigure(1, weight=1)

        extract_label_status = ttk.Label(self.extract_pbar_frame, text="Please wait...")
        extract_label_status.grid(column=0, row=0, sticky=tk.W, padx=5)

        self.extract_pbar = ttk.Progressbar(self.extract_pbar_frame, orient=tk.HORIZONTAL, mode='indeterminate')
        self.extract_pbar.grid(column=1, row=0, sticky=tk.EW, padx=0)

        pbar_frame.columnconfigure(0, weight=1)
        pbar_frame.grid(column=0, row=7, sticky=tk.EW, padx=5)

        self.extract_button = ttk.Button(extractor_frame, text='Extract', command=self.extract_pages)
        self.extract_button.grid(column=1, row=7, sticky=tk.EW, padx=5)

        separator_bottom = ttk.Separator(self, orient='horizontal')
        separator_bottom.grid(columnspan=2, column=0, row=8, sticky=tk.EW, pady=(5, 5), padx=5)

        extractor_frame.grid(column=0, row=0, sticky=tk.NSEW)

    def extract_pages(self):
        message_nothing_to_do = 'Nothing to do...\nPlease choose a Source File'
        if self.extr_type_combobox.get() == self.extr_type_combobox_values[1]:
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
                    range_extract_thread = RangeExtractThread(self.pdf_reader, output_path, pages_range,
                                                              self.extracting_page_progress)
                    range_extract_thread.start()
                    self.range_extract_thread_monitor(thread=range_extract_thread, pdf_reader=self.pdf_reader,
                                                      output_path=output_path, pages_range=pages_range)
            except ValueError as ex:
                logging.error(ex)
                messagebox.showwarning(title='Warning!', message='Invalid range format!\n{}'.format(str(ex)))
                return
            except Exception as ex:
                logging.error(ex)
                self.stop_thread()
                self.extracting_page_progress.grid_remove()
                messagebox.showwarning(title='Warning!', message='Something went wrong...')

        elif self.extr_type_combobox.get() == self.extr_type_combobox_values[0]:
            if not self.input_file_name['text']:
                tk.messagebox.showinfo('Information...', message_nothing_to_do)
                return
            output_dir_name = self.update_output_dir_name(self.source_file_path)
            output_path = os.path.dirname(self.source_file_path)
            output_path = filedialog.askdirectory(title='Save to...', initialdir=output_path)
            try:
                if output_path:
                    self.start_thread()
                    pbp_extract_thread = PbPExtractThread(self.source_file_path, self.pdf_reader, output_path,
                                                          output_dir_name, self.extracting_page_progress)
                    pbp_extract_thread.start()
                    self.pbp_extract_thread_monitor(thread=pbp_extract_thread, source_file_path=self.source_file_path,
                                                    pdf_reader=self.pdf_reader, output_path=output_path,
                                                    output_dir_name=output_dir_name)
            except Exception as ex:
                logging.error(ex)
                self.stop_thread()
                self.extracting_page_progress.grid_remove()
                messagebox.showwarning(title='Warning!', message='Something went wrong...')

    def range_extract_thread_monitor(self, thread, pdf_reader, output_path, pages_range):
        self.extracting_page_progress.grid(columnspan=3, column=0, row=2, sticky=tk.EW, padx=(5, 5), pady=(10, 0))
        if thread.is_alive():
            self.after(100, lambda: self.range_extract_thread_monitor(thread=thread, pdf_reader=pdf_reader,
                                                                      output_path=output_path, pages_range=pages_range))
        else:
            self.extracting_page_progress.grid_remove()
            self.stop_thread()
            if thread.get_message():
                messagebox.showwarning(title='Warning!', message=thread.get_message())

    def pbp_extract_thread_monitor(self, thread, source_file_path, pdf_reader, output_path, output_dir_name):
        self.extracting_page_progress.grid(columnspan=3, column=0, row=2, sticky=tk.EW, padx=(5, 5), pady=(10, 0))
        if thread.is_alive():
            self.after(100, lambda: self.pbp_extract_thread_monitor(thread=thread, source_file_path=source_file_path,
                                                                    pdf_reader=pdf_reader, output_path=output_path,
                                                                    output_dir_name=output_dir_name))
        else:
            self.extracting_page_progress.grid_remove()
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
        self.extract_button['state'] = tk.DISABLED
        self.open_file_button['state'] = tk.DISABLED
        self.extr_type_combobox['state'] = tk.DISABLED
        self.page_range_entry['state'] = tk.DISABLED
        self.extract_pbar_frame.grid(column=0, row=0, sticky=tk.EW, padx=0)
        self.extract_pbar.start(10)

    def stop_thread(self):
        self.extract_button['state'] = tk.NORMAL
        self.open_file_button['state'] = tk.NORMAL
        self.extr_type_combobox['state'] = tk.NORMAL
        self.page_range_entry['state'] = tk.NORMAL
        self.extract_pbar.stop()
        self.extract_pbar_frame.grid_remove()

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

    def update_output_dir_name(self, source_file):
        curr_datetime = datetime.datetime.today().strftime('%Y%m%d%H%M%S')
        source_file_name, source_file_extension = os.path.splitext(os.path.basename(source_file))
        dir_name = '{}_extr_result_{}'.format(source_file_name, curr_datetime)
        return dir_name

    def update_output_file_name(self, source_file):
        curr_datetime = datetime.datetime.today().strftime('%Y%m%d%H%M%S')
        source_file_name, source_file_extension = os.path.splitext(os.path.basename(source_file))
        file_name = '{}_extr_result_{}{}'.format(source_file_name, curr_datetime, source_file_extension)
        return file_name

    def parse_pages_range(self, parse_string):
        result = []
        for part in parse_string.split(','):
            x = part.split('-')
            result.extend(range(int(x[0]), int(x[-1]) + 1))
        return result

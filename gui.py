#!/usr/bin/env python
# -*- coding: utf-8 -*-
# cython: language_level=3
"""
    Convert RFP from Excel to Word.
"""
import os
import sys
import time
import re
import traceback
#from collections import OrderedDict

import tkinter as tk
from tkinter import filedialog
from tkinter import font
import tkinter.ttk as ttk
from tkinter import messagebox

import xtow
import config
import version

WHATS_NEW = [
        ('0.9', 'Command line release'),
        ('1.0', 'Added GUI interface'),
        ('1.1', 'Status line added providing ability to see why next button is in disabled mode'),
        ('1.2', 'Added quit option during generation process'),
        ('1.3', "Added What's New pupup window"),
        ('1.4', 'Added counter for generated document features'),
        ('1.5', 'Added skip ordering phase in case only one source file is selected'),
        ('1.6', 'Added message on error in cell'),
]

version.add_release('0.9', 'Command line release')
version.add_release('1.0', 'Added GUI interface')
version.add_release('1.1', 'Status line added providing ability to see why next button is in disabled mode')
version.add_release('1.2', 'Added quit option during generation process')
version.add_release('1.3', "Added What's New pupup window")
version.add_release('1.4', 'Added counter for generated document features')
version.add_release('1.5', 'Added skip ordering phase in case only one source file is selected')
version.add_release('1.6', 'Dramatically improved performance')
version.add_release('1.7', 'Added message on error in cell')
version.add_release('1.8', 'Updated python to 3.9.1 and rebuild to support Big Sur')


def whats_new_text(last_version):
    return '\n\n'.join(['Version %s:\n\t%s' % (v, text)
                        for v, text in version.iterate_whats_new(last_version)])


# Geometry constants
WIDTH = 500
HEIGHT = 440
HSTATUS = 40
SLMARGIN=5
WLEFT = 100
WLOGO = 64
HSPACE = 40
HSSPACE = 28
#HMSPACE = 22
HMSPACE = 18
HBUTTON = 40
WBUTTON = 100
MARGIN = (WLEFT - WLOGO) // 2


DEFAULT_DESTINATION = 'no default'


class SetupPhase(object):
    phases = list()
    conf = dict(
        title='Converter',
        options=list(),
        sheets=list()
    )
    last_conf = dict()
    root = tk.Tk()
    status_var = tk.StringVar()
    default_status = ''
    destination_folder = DEFAULT_DESTINATION

    next_btn = None

    current_phase = 0

    y = MARGIN
    stop_flag = False

    zero_len = font.Font(family=tk.NORMAL, size=100, weight="normal").measure("0")
    macos_len = 61
    LABEL_FONT = (None, int(12 / macos_len * zero_len), font.NORMAL)
    BUTTON_FONT = LABEL_FONT
    SMALL_FONT = (None, int(10 / macos_len * zero_len), font.NORMAL)
    TITLE_FONT = (None, int(14 / macos_len * zero_len), font.BOLD)
    TIP_FONT = (None, int(9 / macos_len * zero_len), font.NORMAL)
    PHASES_FONT_SIZE = int(12 / macos_len * zero_len)
    STATUS_FONT = (None, int(9 / macos_len * zero_len), font.NORMAL)

    @classmethod
    def mainloop(cls):
        cls.root.mainloop()

    @classmethod
    def cancel(cls):
        answer = messagebox.askyesno(
            title='Quit',
            message='Are you sure?'
        )
        if answer:
            # destroy
            cls.stop_flag = True
            cls.root.quit()

    @classmethod
    def show_error(self, *args):
        err = traceback.format_exception(*args)
        messagebox.showerror('Exception', ''.join(err))

    @classmethod
    def acquire_data(cls):
        pass

    @classmethod
    def next(cls):
        cls.change_phase(1)

    @classmethod
    def prev(cls):
        cls.change_phase(-1)

    @classmethod
    def change_phase(cls, delta):
        cls.acquire_data()
        SetupPhase.current_phase += delta
        cls.cleanup()
        cls.phases[SetupPhase.current_phase].operate()

    @staticmethod
    def icon_path():
        file_name = 'media/bullet-icon.gif'
        if hasattr(sys, '_MEIPASS'):
            return os.path.join(sys._MEIPASS, file_name)
        else:
            return os.path.join(os.path.dirname(sys.argv[0]), file_name)

    @classmethod
    def add_left_pane(cls):
        frame = tk.Frame(height=HEIGHT - HSTATUS - MARGIN * 2, width=2, bd=1, relief=tk.SUNKEN)
        frame.place(x=WLEFT, y=MARGIN)

    @classmethod
    def add_image(cls):
        cls.image = tk.PhotoImage(file=cls.icon_path())
        tk.Label(cls.root, image=cls.image).place(x=MARGIN, y=MARGIN)

    @classmethod
    def progress_list(cls):
        y = MARGIN + WLOGO + HSPACE
        for phase in SetupPhase.phases:
            if phase is cls:
                weight = font.NORMAL
                text = '\u25B6 %s' % phase.name
            else:
                weight = font.NORMAL
                text = "   %s" % phase.name
            label = tk.Label(cls.root, text=text, font=(None, cls.PHASES_FONT_SIZE, weight))
            label.place(x=MARGIN - 10, y=y)
            y += HMSPACE

    @classmethod
    def add_cancel(cls):
        btn = tk.Button(cls.root, text=" Cancel ", command=cls.cancel, font=cls.BUTTON_FONT)
        btn.place(x=WIDTH - MARGIN, y=HEIGHT - HSTATUS - MARGIN, anchor=tk.SE)

    @classmethod
    def next_button_label(cls):
        return " Next > "

    @classmethod
    def add_next(cls):
        cls.next_btn = tk.Button(cls.root, text=cls.next_button_label(), command=cls.next, font=cls.BUTTON_FONT)
        cls.next_btn.config(default=tk.ACTIVE)
        cls.next_btn.place(x=WLEFT + MARGIN + WBUTTON, y=HEIGHT - HSTATUS - MARGIN, anchor=tk.SW)
        cls.run_fields_check()

    @classmethod
    def add_prev(cls):
        btn = tk.Button(cls.root, text=" < Previous ", command=cls.prev, font=cls.BUTTON_FONT)
        btn.place(x=WLEFT + MARGIN, y=HEIGHT - HSTATUS - MARGIN, anchor=tk.SW)

    @classmethod
    def add_status_line(cls):
        frame = tk.Frame(height=HSTATUS-MARGIN, width=WIDTH-MARGIN*2,
                         bd=1, relief=tk.SUNKEN)
        frame.place(x=MARGIN, y=HEIGHT-MARGIN, anchor=tk.SW)
        label = tk.Label(frame, textvariable=cls.status_var, font=cls.STATUS_FONT)
        label.place(x=0, y=(HSTATUS-MARGIN)/2-2, anchor=tk.W)
        cls.set_status(cls.default_status)

    @classmethod
    def set_status(cls, text):
        cls.status_var.set(text)

    @classmethod
    def add_title(cls, text):
#        label = tk.Label(self.root, text=text, font=TITLE_FONT)
        label = tk.Message(cls.root, text=text, width=WIDTH - WLEFT - MARGIN * 2, font=cls.TITLE_FONT)
        label.place(x=WLEFT + MARGIN, y=SetupPhase.y)
        SetupPhase.y += HSPACE

    @classmethod
    def add_pick_folder_file(cls, prompt, variable, callback, tip=None):
        label = tk.Label(cls.root, text=prompt, font=cls.LABEL_FONT)
        label.place(x=WLEFT + MARGIN, y=SetupPhase.y)
        SetupPhase.y += HSPACE

        vcmd = (cls.root.register(cls.validate), '%P')
        entry = tk.Entry(cls.root, textvariable=variable, width=34, font=cls.LABEL_FONT,
                         validate="key", validatecommand=vcmd)
        entry.place(x=WLEFT + MARGIN, y=SetupPhase.y, anchor=tk.W)

        button = tk.Button(cls.root, text="Choose", command=callback, font=cls.BUTTON_FONT)
        button.place(x=WIDTH - MARGIN, y=cls.y, anchor=tk.E)

        if tip is not None:
            SetupPhase.y += HMSPACE
            message = tk.Message(cls.root, width=WIDTH - WLEFT - MARGIN * 2 - 70, text=tip, font=cls.LABEL_FONT)
            message.place(x=WLEFT + MARGIN, y=SetupPhase.y)
            # label = tk.Label(self.root, text=tip, font=TIP_FONT)
            # label.place(x=WLEFT + MARGINE, y=SetupPhase.y)
            SetupPhase.y += HSPACE
        else:
            SetupPhase.y += HSPACE

    @classmethod
    def add_pick_folder(cls, prompt, variable, tip=None):
        def pick_folder_callback():
            cls.pick_folder(variable)
            cls.run_fields_check()
        cls.add_pick_folder_file(prompt, variable, pick_folder_callback, tip=tip)

    @classmethod
    def add_pick_files(cls, prompt, variable, ftypes, tip=None):
        def pick_file_callback():
            cls.pick_files(variable, ftypes)
            cls.run_fields_check()
        cls.add_pick_folder_file(prompt, variable, pick_file_callback, tip=tip)


    @classmethod
    def add_input(cls, prompt, width, variable, tip=None):
        label = tk.Label(cls.root, text=prompt, font=LABEL_FONT)
        label.place(x=WLEFT + MARGIN, y=SetupPhase.y)
        SetupPhase.y += HMSPACE

        vcmd = (cls.root.register(cls.validate), '%P')
        entry = tk.Entry(cls.root, width=width, textvariable=variable, validate="key", validatecommand=vcmd)
        entry.place(x=WLEFT + MARGIN, y=SetupPhase.y)
        if tip is not None:
            SetupPhase.y += HMSPACE
            message = tk.Message(cls.root, text=tip, width=WIDTH - WLEFT - MARGIN * 2 - 70, font=cls.TIP_FONT)
            message.place(x=WLEFT + MARGIN, y=SetupPhase.y)
            SetupPhase.y += HSPACE
        else:
            SetupPhase.y += HSSPACE

    @classmethod
    def add_button(cls, text, command):
        btn = tk.Button(cls.root, text=text, command=command, font=cls.BUTTON_FONT)
        btn.place(x=WLEFT + MARGIN, y=SetupPhase.y)
        SetupPhase.y += HMSPACE

    @classmethod
    def add_checkbutton(cls, text, variable, tip=None):
        check = tk.Checkbutton(cls.root,
                               text=text,
                               variable=variable,
                               font=LABEL_FONT)

        check.place(x=WLEFT + MARGIN, y=SetupPhase.y)
        if tip is not None:
            SetupPhase.y += HMSPACE
            label = tk.Label(cls.root, font=cls.TIP_FONT, text=tip)
            label.place(x=WLEFT + MARGIN, y=SetupPhase.y)
            SetupPhase.y += HSPACE
        else:
            SetupPhase.y += HSSPACE
        return check

    @classmethod
    def add_progress(cls, text, variable, maximum):
        label = tk.Label(cls.root, text=text)
        label.place(x=WLEFT + MARGIN, y=SetupPhase.y)
        SetupPhase.y += HSPACE
        bar = ttk.Progressbar(cls.root,
                              orient="horizontal",
                              length=WIDTH - WLEFT - MARGIN * 3,
                              mode="determinate",
                              variable=variable,
                              maximum=maximum)
        bar.place(x=WLEFT + MARGIN, y=SetupPhase.y)
        SetupPhase.y += HSPACE
        return bar, label

    @classmethod
    def add_message(cls, text, **kwargs):
        options = dict(font=cls.LABEL_FONT)
        options.update(kwargs)
        message = tk.Message(cls.root,
                             width=WIDTH - WLEFT - MARGIN * 3,
                             text=text,
                             **options)
        message.place(x=WLEFT + MARGIN, y=SetupPhase.y)
        SetupPhase.y += HSPACE

    @classmethod
    def add_dropdown(cls, name, var, choices):
        SetupPhase.y += HSSPACE // 2
        label = tk.Label(cls.root, text=name)
        label.place(x=WLEFT + MARGIN, y=SetupPhase.y, anchor=tk.W)

        menu = tk.OptionMenu(cls.root, var, *choices)
        menu.config(width=10)
        menu.place(x=WLEFT+(WIDTH-WLEFT)//2, y=SetupPhase.y, anchor=tk.W)
        SetupPhase.y += HMSPACE

    @classmethod
    def add_listbox(cls, elements):
        frame = tk.Frame(cls.root)
        frame.place(x=WLEFT + MARGIN, y=SetupPhase.y)

        listbox = tk.Listbox(frame, selectmode=tk.MULTIPLE,
                             width=36, height=16)#, font=("Helvetica", 12))
        listbox.pack(side="left", fill="y")

        scrollbar = tk.Scrollbar(frame, orient="vertical")
        scrollbar.config(command=listbox.yview)
        scrollbar.pack(side="right", fill="y")

        listbox.config(yscrollcommand=scrollbar.set)

        for each in elements:
            listbox.insert(tk.END, each)

        return listbox

    @classmethod
    def init(cls):
        tk.Tk.report_callback_exception = cls.show_error

        ws = cls.root.winfo_screenwidth()
        hs = cls.root.winfo_screenheight()

        cls.win_pos_x = (ws - WIDTH) // 2
        cls.win_pos_y = (hs - HEIGHT) // 2

        cls.root.geometry("{}x{}+{}+{}".format(WIDTH, HEIGHT, cls.win_pos_x, cls.win_pos_y))
        cls.root.resizable(
            width=True,
            height=True
        )
        #cls.root.iconbitmap('te_mac.icns')
        cls.root.title("RFP Converter")
        cls.root.protocol("WM_DELETE_WINDOW", cls.cancel)
        #cls.root.attributes("-topmost", True)

        #cls.root.lift()
        cls.root.attributes('-topmost', True)
        cls.root.update()
        cls.root.attributes('-topmost', False)

        cls.phases[0].operate()


    @classmethod
    def cleanup(cls):
        for each in list(cls.root.children):
            cls.root.children[each].destroy()

    @classmethod
    def basic_widgets(cls):
        cls.add_image()
        cls.progress_list()
        cls.add_left_pane()
        cls.add_next()
        cls.add_prev()
        cls.add_cancel()
        cls.add_status_line()
        cls.run_fields_check()

    @classmethod
    def operate(cls):
        cls.basic_widgets()
        SetupPhase.y = MARGIN
        cls.decorate()

    @classmethod
    def decorate(cls):
        pass

    @classmethod
    def pick_folder(cls, var):
        initialdir = config.get('pick_folder')
        dir_name = filedialog.askdirectory(initialdir=initialdir, #os.path.expanduser('~'),#initialdir=var.get(),
                                             title="Select folder")
        if dir_name == '':
            return
        config.set('pick_folder', dir_name)
        var.set(dir_name)

    @classmethod
    def pick_files(cls, var, ftypes):
        initialdir = config.get('pick_files')
        files = filedialog.askopenfilenames(initialdir=initialdir, #initialdir=var.get().split(',')[0].strip(),
                                            title="Select files",
                                            filetypes=ftypes)
        if not files:
            return
        config.set('pick_files', os.path.dirname(files[0]))
        var.set(', '.join(files))

    @staticmethod
    def validate_host(string):
        ip = '^(([0-9]|[1-9][0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])\.){3}([0-9]|[1-9][0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])$'
        hostname = '^(([a-zA-Z0-9]|[a-zA-Z0-9][a-zA-Z0-9\-]*[a-zA-Z0-9])\.)*([A-Za-z0-9]|[A-Za-z0-9][A-Za-z0-9\-]*[A-Za-z0-9])$'
        return re.match(hostname, string) is not None or re.match(ip, string) is not None

    @staticmethod
    def validate_uuid(string):
        uuid = '^[A-Za-z0-9]{8}-[A-Za-z0-9]{4}-[A-Za-z0-9]{4}-[A-Za-z0-9]{4}-[A-Za-z0-9]{12}$'
        return re.match(uuid, string) is not None

    @staticmethod
    def validate_email(string):
        email = '^([a-zA-Z0-9_\-\.]+)@([a-zA-Z0-9_\-\.]+)\.([a-zA-Z]{2,5})$'
        return re.match(email, string) is not None

    @staticmethod
    def validate_port(string):
        try:
            v = int(string)
        except ValueError:
            return False
        return 0 <= v <= 0xFFFF

    @classmethod
    def validate(cls, P=''):
        cls.root.after(0, cls.run_fields_check)
        return True

    @classmethod
    def run_fields_check(cls):
        if cls.check_fields():
            cls.next_btn.config(state=tk.NORMAL)
        else:
            cls.next_btn.config(state=tk.DISABLED)

    @classmethod
    def check_fields(cls):
        return True

    @staticmethod
    def yesno(title, message):
        return messagebox.askyesno(title, message)

    @classmethod
    def output_file_name(cls):
        return cls.output_file_name_for(
            SetupPhase.conf['target_folder'],
            SetupPhase.conf['sources']
        )

    @classmethod
    def output_file_name_for(cls, folder, sources):
        prefix = os.path.basename(folder)
        name = '_'.join(os.path.splitext(os.path.basename(fname))[0]
                        for fname in sources)
        today = time.strftime('%d%m%Y')
        file_name = '{}_{}_RFP_{}.docx'.format(prefix, name, today)
        return os.path.join(folder, file_name)

    @classmethod
    def open_workbook(cls, format=None):
        SetupPhase.conf['output_file_name'] = cls.output_file_name()
        #print('Output file name: {}'.format(
        #    SetupPhase.conf['output_file_name']
        #))
        if format is None:
            converter = xtow.ExcelToWord
        else:
            converter = xtow.available_converters[format]
        return converter(
            SetupPhase.conf['sources'],
            SetupPhase.conf['output_file_name'],
            criteria=SetupPhase.conf['options'],
            sheets=SetupPhase.conf['sheets']
        )

    @classmethod
    def font_size(cls, size):
        macos_len = 61.0
        return int(size/macos_len*cls.zero_len)


class IntroPhase(SetupPhase):
    name = 'Intro'
    default_status = 'Check your XLS files for requirements and press [ Next > ] button'

    @classmethod
    def add_prev(cls):
        pass

    @classmethod
    def pop_up_whats_new(cls, text):
        WHATS_NEW_WIDTH=600
        window = tk.Toplevel()

        window.wm_title("What's New")

        label = tk.Label(window, text="List of new features:", anchor=tk.W)
        label.pack(fill=tk.BOTH)

        message = tk.Message(window, text=text, width=WHATS_NEW_WIDTH)
        message.pack()

        button = tk.Button(window, text=' Ok  ', command=window.destroy)
        button.pack()
        window.update()
        sw = cls.root.winfo_screenwidth()
        sh = cls.root.winfo_screenheight()
        ww = window.winfo_width()
        wh = window.winfo_height()
        print(wh, ww)
        win_pos_x = (sw - ww) // 2
        win_pos_y = (sh - wh) // 2
        window.geometry("{}x{}+{}+{}".format(ww, wh+20, win_pos_x, win_pos_y))
        window.attributes('-topmost', True)
        window.update()
        window.attributes('-topmost', False)

        return window

    @classmethod
    def show_whats_new(cls):
        config_version = config.get('version')
        """
        if config_version == '':
            print('First run')
            config.set('version', VERSION)
        elif float(VERSION) > float(config_version):
            print('Upgrade')
            config.set('version', VERSION)
        else:
            print('Second run and more')
        """
        #whats_new = collect_whats_new(config_version)
        if version.have_new_release(config_version):
            config.set('version', version.version())
            text = whats_new_text(config_version)
            window = cls.pop_up_whats_new(text)
            #cls.root.wait_window(window)
            #messagebox.showinfo(
            #    title="What's New",
            #    message=text
            #)

    @classmethod
    def decorate(cls):
        cls.add_message(xtow.description)
        cls.add_message(text=xtow.requirements_en, font=cls.SMALL_FONT)
        cls.root.after(0, cls.show_whats_new)


SetupPhase.phases.append(IntroPhase)


class SourcesPhase(SetupPhase):
    name = 'Sources'
    default_status = 'Press [ Next > ] button'

    sources_var = tk.StringVar()
    target_var = tk.StringVar()

    @classmethod
    def sources_list(cls):
        return [fname.strip() for fname in cls.sources_var.get().split(',')]

    @classmethod
    def check_fields(cls):
        for path in cls.sources_list():
            if not os.path.isfile(path):
                cls.set_status('Provide source file(s)')
                return False
        if not os.path.isdir(cls.target_var.get()):
            cls.set_status('Target folder does not exist')
            return False
        output = cls.output_file_name_for(cls.target_var.get(), cls.sources_list())
        if os.path.exists(output):
            cls.set_status('Output file already exists')
            return False
        cls.set_status(cls.default_status)
        return True

    @classmethod
    def decorate(cls):
        cls.add_pick_files("Source Files", cls.sources_var, (("MS Excel file", "*.xls"),),
                           tip='(Multiply files can be selected)')
        cls.add_pick_folder("Target Folder", cls.target_var)

    @classmethod
    def acquire_data(cls):
        SetupPhase.conf['sources'] = [s.strip() for s in cls.sources_var.get().split(',')]
        SetupPhase.conf['target_folder'] = cls.target_var.get()

    @classmethod
    def next(cls):
        if len(cls.sources_var.get().split(',')) == 1:
            cls.change_phase(2)
        else:
            cls.change_phase(1)


SetupPhase.phases.append(SourcesPhase)


class OrderPhase(SetupPhase):
    name = 'Order'
    default_status = 'Order requirements'

    @classmethod
    def acquire_data(cls):
        pass

    @classmethod
    def add_source(cls, n, file_name):
        def fixup(file_name):
            max_length = 33
            if len(file_name) < max_length:
                return file_name
            return '...' + file_name[-max_length-3:]

        def move_up():
            s = SetupPhase.conf['sources']
            s[n], s[n-1] = s[n-1], s[n]
            SetupPhase.y = MARGIN
            cls.decorate()

        def move_down():
            s = SetupPhase.conf['sources']
            s[n+1], s[n] = s[n], s[n+1]
            SetupPhase.y = MARGIN
            cls.decorate()

        BUTTON_SIZE = 42
        label = tk.Label(cls.root, text='%d.' % (n+1))
        label.place(x=WLEFT + MARGIN, y=SetupPhase.y, anchor=tk.W)
        label = tk.Label(cls.root, text=fixup(file_name))
        label.place(x=WIDTH - MARGIN - 2*BUTTON_SIZE, y=SetupPhase.y, anchor=tk.E)
        up_btn = tk.Button(cls.root, text='\u2b06', command=move_up, font=BUTTON_FONT)
        up_btn.place(x=WIDTH - MARGIN - BUTTON_SIZE, y=SetupPhase.y, anchor=tk.E)

        if n == 0:
            up_btn.config(state="disabled")
        down_btn = tk.Button(cls.root, text='\u2b07', command=move_down, font=BUTTON_FONT)
        down_btn.place(x=WIDTH - MARGIN, y=SetupPhase.y, anchor=tk.E)
        if n == len(SetupPhase.conf['sources'])-1:
            down_btn.config(state="disabled")
        SetupPhase.y += HSPACE

    @classmethod
    def decorate(cls):
        cls.add_title("Order Excel files")
        for n, source in enumerate(SetupPhase.conf['sources']):
            cls.add_source(n, source)
        if len(SetupPhase.conf['sources']) == 1:
            cls.set_status('Press [ Next > ] button')


SetupPhase.phases.append(OrderPhase)


class SuitesPhase(SetupPhase):
    name = 'Suites'
    default_status = 'Pick up required software suites'
    listbox = None

    @classmethod
    def acquire_data(cls):
        SetupPhase.conf['options'] = []
        for index in cls.listbox.curselection():
            SetupPhase.conf['options'].append(cls.listbox.get(index))


    @classmethod
    def decorate(cls):
        etow = cls.open_workbook()
        full_list = etow.list_packages()

        cls.add_title('Pick required product suites')
        cls.listbox = cls.add_listbox(full_list)
        if SetupPhase.conf['options']:
            for number, name in enumerate(full_list):
                if name in SetupPhase.conf['options']:
                    cls.listbox.select_set(number)


SetupPhase.phases.append(SuitesPhase)

class SheetsPhase(SetupPhase):
    name = 'Sheets'
    default_status = 'Remove selection from unnecessary sheets'
    sheets_listbox = None

    @classmethod
    def acquire_data(cls):
        SetupPhase.conf['sheets'] = []
        for index in cls.sheets_listbox.curselection():
            SetupPhase.conf['sheets'].append(cls.sheets_listbox.get(index))


    @classmethod
    def decorate(cls):
        etow = cls.open_workbook()
        sheets = etow.list_sheets()
        cls.add_title('Unpick unnecessary sheets')
        cls.sheets_listbox = cls.add_listbox(sheets)
        if SetupPhase.conf['sheets']:
            for number, name in enumerate(sheets):
                if name in SetupPhase.conf['sheets']:
                    cls.sheets_listbox.select_set(number)
        else:
            cls.sheets_listbox.select_set(0, tk.END)



class FormatPhase(SetupPhase):
    name = 'Format'
    default_status = 'Pick up required output file format'
    default_format = xtow.default_converter.name
    format_var = tk.StringVar(value=default_format)

    @classmethod
    def check_fields(cls):
        return True

    @classmethod
    def acquire_data(cls):
        SetupPhase.conf['format'] = cls.format_var.get()

    @classmethod
    def decorate(cls):
        cls.add_title('Pick output file format')
        cls.add_dropdown(
            name='Format',
            var=cls.format_var,
            choices=xtow.available_converters.keys()
        )

    @classmethod
    def next_button_label(cls):
        return " Generate > "


SetupPhase.phases.append(FormatPhase)


class GeneratePhase(SetupPhase):
    name = 'Generate'
    default_status = 'Generating output file...'

    progress_var = tk.IntVar()
    open_result_var = tk.IntVar(value=1)

    @classmethod
    def reveal_result(cls):
        if cls.open_result_var.get() == 0:
            return
        #command = 'open "%s"' % os.path.dirname(SetupPhase.conf['output_file_name'])
        #os.system(command)
        import subprocess
        command = [
            'open',
            '-R',
            SetupPhase.conf['output_file_name']
        ]
        #print(' '.join(command))
        subprocess.call(command)

    @classmethod
    def add_option(cls, name, var, y):
        # unused
        label = tk.Label(cls.root, text=name)
        label.place(x=WLEFT + MARGIN, y=y)
        entry = tk.Entry(cls.root, textvariable=var)
        entry.place(x=WLEFT + MARGIN, y=y + HSSPACE)

    @classmethod
    def check_fields(cls):
        return True # Check exceptions from processing

    @classmethod
    def decorate(cls):
        cls.next_btn.config(state=tk.DISABLED)

        cls.add_title("Processing")

        counter = xtow.Counter(SetupPhase.conf['sources'],
                               SetupPhase.conf['output_file_name'],
                                criteria=SetupPhase.conf['options'],
                                sheets=SetupPhase.conf['sheets'])

        lines = counter.count_scope_lines()
        print('lines', lines)
        print('count', counter.counter)
        bar, sheet_name_label = cls.add_progress(
            text="Progress",
            variable=cls.progress_var,
            maximum=counter.count_scope_lines()
        )

        cls.add_message('Output file: %s' % os.path.basename(SetupPhase.conf['output_file_name']))

        cls.add_checkbutton(
            text='Reveal result in Finder',
            variable=cls.open_result_var
        )

        etow = cls.open_workbook(SetupPhase.conf['format'])
        requirements_count = 0
        start_time = 0;
        for file_name, sheet, line in etow.run():
            if cls.stop_flag:
                break
            #print('line: {}'.format(line))
            requirements_count += 1
            if time.time() - start_time > 0.2:
                start_time = time.time();
                phase = '%s: %s' % (file_name, sheet)
                sheet_name_label.config(text=phase)
                cls.progress_var.set(line)
                cls.root.update()
        else:
            cls.next_btn.config(state=tk.NORMAL)
            cls.set_status('Generated {} requirements. Press [{}] button to exit'.format(
                requirements_count,
                cls.next_button_label())
            )

    @classmethod
    def next(cls):
        cls.reveal_result()
        cls.root.quit()

    @classmethod
    def next_button_label(cls):
        return " Finish "

    @classmethod
    def add_prev(cls):
        pass

#    @classmethod
#    def add_cancel(cls):
#        pass


SetupPhase.phases.append(GeneratePhase)


def main():
    SetupPhase.init()
    SetupPhase.mainloop()


if __name__ == '__main__':
    main()

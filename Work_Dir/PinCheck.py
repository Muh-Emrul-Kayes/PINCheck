from tkinter import *
from tkinter import filedialog
from tkinter import scrolledtext
from tkinter import messagebox
import sys
import os
import threading
import subprocess
import psutil
import shutil


default__Win__Width = 800
default__Win__Height = 700

min__Win__Width = 700
min__win__Height = 600

gui_title = "GF Library QA"
gui_icon = 'logo.ico'


document_path = r"H:\Ulkasemi\Tkinter\TU1\pin_check\pin_check.pdf"
script_path = r"H:\Ulkasemi\Tkinter\TU1\pin_check\pinCheckFE.py"
report_path = r'H:\Ulkasemi\Tkinter\TU1\outputfile\FE_CHECK\report\pin_check.xlsx'


current_path = os.getcwd()
# path = os.path.abspath(os.path.join(path, os.pardir))  # parent path
# print(path)
# path = f'{path}\pin_check\pinCheckFE.py'
# print(path)

# ================================================================#
#                       Main GUI Window                           #
# ================================================================#
# Create Main Window
root = Tk()

# Set GUI Icon
root.iconbitmap(gui_icon)

# Set GUI Title
root.title(gui_title)


# ================================================================#
#                   Geometry Management                           #
# ================================================================#

# Define Default Window Size('Width x Height')
root.geometry('%dx%d' % (default__Win__Width, default__Win__Height))

# Define Minimum Window Size
root.minsize(min__Win__Width, min__win__Height)


# ================================================================#
#                           Intro Frame                           #
# ================================================================#

# Create Frame
intro_frame = Frame(root)

# Add Frame to main Window
intro_frame.grid(row=0, column=0)


#########################
#         Label         #
#########################

# Create Label
intro_label = Label(
    intro_frame, text='Std Cell Library Pin Consistency Check and Datasheet Generator',
    font=('comicsansms', 15, 'bold')
)

designer_label = Label(
    root, text='By Automation Team, Ulkasemi',
    font=('comicsansms', 10, 'bold')
)

# Add Label to Intro Frame
intro_label.grid(row=0, column=0, padx=10)
designer_label.grid(row=1, column=0, padx=10)


# ================================================================#
#                            Input Frame                          #
# ================================================================#

# create Frame
inpt_frame = LabelFrame(
    root,
    text='Pin Check',
    pady=5
)

# Add frame to main window
inpt_frame.grid(row=3, column=0, padx=20, pady=20, sticky=W+E)


#########################
#        Dropdown       #
#########################

# Create dropdow
options = ['MEMORY', 'IO', 'LOGIC']
option_name = StringVar()
option_name.set(options[0])

drop = OptionMenu(inpt_frame, option_name, *options)

# Add dropdown to input frame
drop.grid(row=3, column=0, sticky=W, pady=10, columnspan=2)


#########################
#         Label         #
#########################

# Create Label
fr4_file1_label = Label(inpt_frame, text='Library')
fr4_file3_label = Label(inpt_frame, text='Output')

# Add Label to input frame
fr4_file1_label.grid(row=4, column=0, sticky=W)
fr4_file3_label.grid(row=5, column=0, sticky=W)


#########################
#         Entry         #
#########################

# Create Entry
fr4_Library_Enrty = Entry(inpt_frame)
fr4_output_Entry = Entry(inpt_frame)

# Add Entry to input frame
fr4_Library_Enrty.grid(row=4, column=1, padx=5, sticky=W+E)
fr4_output_Entry.grid(row=5, column=1, padx=5, sticky=W+E)

#########################
#     configuration     #
#########################

inpt_frame.columnconfigure(1, weight=1)


fr4_entry_dic = [
    {'entry': fr4_Library_Enrty,    'tag': 'Library'},
    {'entry': fr4_output_Entry,     'tag': 'Output'},
]


#-------------------------#
#    Command Functions    #
#-------------------------#

def create_cmd():
    cmd = f'python {script_path} -i {option_name.get()} -l {fr4_Library_Enrty.get()}'
    return cmd


global pid
pid = 1234


def run_command(run_button, stop_button):
    global pid
    cmd = create_cmd()
    print(cmd)
    p = subprocess.Popen(cmd.split(),
                         stdout=subprocess.PIPE, stderr=subprocess.STDOUT, bufsize=1, text=True)

    run_button.grid_forget()
    stop_button.grid(row=6, column=2, pady=5)
    pid = p.pid
    while True:
        msg = p.stdout.readline().strip()
        if msg:
            print(msg)
        if not msg:
            break
    while p.poll() is None:
        msg = p.stdout.readline().strip()
        if msg:
            print(msg)
    stop_button.grid_forget()
    run_button.grid(row=6, column=2, pady=5)

    if p.returncode == 0:
        if os.path.exists(report_path):
            messagebox.showinfo(
                title='Run Complte', message='Script has run successfully!')
            fr4_open_report_button['state'] = NORMAL


def kill_process(pid, run_button, stop_button):
    if psutil.pid_exists(pid):
        p = psutil.Process(pid)
        p.kill()
        txtbox.insert('end', 'Process has been stoped!\n')
        run_button.grid(row=6, column=2, pady=5)
        stop_button.grid_forget()
        run_button.grid(row=6, column=2, pady=5)


def run(run_button, stop_button, entry_list_dic):
    fr4_open_report_button['state'] = DISABLED
    for dic in entry_list_dic:
        if not dic['entry'].get():
            ent = dic['tag']
            messagebox.showerror("error", f'{ent} Entry is Empty!')
            return
        if os.path.exists(dic['entry'].get()):
            continue
        else:
            msg = dic['tag']
            messagebox.showerror("error", f'{msg} does not exist!')
            return

    for index, info in enumerate(entry_list_dic):
        if info['tag'] == 'Output':
            output_dir = info['entry'].get()
            break
    os.chdir(output_dir)
    lib_setup_path = os.path.abspath("FE_CHECK")
    if os.path.exists(lib_setup_path):
        status = messagebox.askyesno(
            title='LIBRARY SETUP FILES ALREADY EXISTS', message='DO YOU WANT TO DELETE THE SETUP FILES ?')
        if not status:
            return

        for filename in os.listdir(output_dir):
            file_path = os.path.join(output_dir, filename)
            try:
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    os.unlink(file_path)
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)
            except Exception as e:
                print('Failed to delete %s. Reason: %s' % (file_path, e))
                return

    threading.Thread(target=lambda: run_command(
        run_button, stop_button)).start()


def dummyCommand():
    pass


'''
psutil.pid_exists(pid):
        p = psutil.Process(pid)
        p.kill()
'''

global proc1, proc2
proc1, proc2 = 1235, 1236


def show(file_path):
    global proc1, proc2
    cmd = f'start {file_path}'
    subproc = subprocess.Popen(
        cmd.split(),
        stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=True
    )
    subproc.wait()


def showFile(file_path, name):
    threading.Thread(target=lambda: show(file_path), name=name).start()


def showAbout():
    messagebox.showinfo(
        title='About', message='Std Cell Library Pin Consistency Check\n and Datasheet Generator')


def exit():
    print(threading.enumerate())
    status = messagebox.askyesno(
        title='Exit Window', message='Do You Want to Exit the Programme?')

    if status:
        root.destroy()
    return


def clear_console():
    txtbox.delete('1.0', END)


def OpenFile(fileType, entry):
    filename = filedialog.askopenfilename(initialdir=current_path, title='Select A File', filetypes=(
        (f'{fileType} files', f'*.{fileType.lower()}'), ('all files', '*.*')))
    if filename:
        entry.delete(0, END)
    entry.insert(END, filename)


def Opendir(entry):
    folderName = filedialog.askdirectory(
        initialdir=current_path, title='Select A Folder')
    if folderName:
        entry.delete(0, END)
    entry.insert(END, folderName)


#########################
#           Button      #
#########################

# Create Button
Library_select = Button(inpt_frame, text='SELECT',
                        command=lambda: Opendir(fr4_Library_Enrty))

Output_folder_select = Button(inpt_frame, text='SELECT',
                              command=lambda: Opendir(fr4_output_Entry))

fr4_run_button = Button(inpt_frame, text="Run", padx=8, font=('comicsansms', 10, 'bold'),
                        command=lambda: run(fr4_run_button, fr4_stop_button, fr4_entry_dic))

fr4_stop_button = Button(inpt_frame, text="Stop", padx=8, font=('comicsansms', 10, 'bold'),
                         command=lambda: kill_process(pid, fr4_run_button, fr4_stop_button), bg='#f37373')

fr4_open_report_button = Button(inpt_frame, text="Open Report", padx=8, font=('comicsansms', 10, 'bold'),
                                command=lambda: showFile(report_path, 'ExcelFile'))


# Add button to input frame
Library_select.grid(row=4, column=2, padx=5, pady=5)
Output_folder_select.grid(row=5, column=2, padx=5, pady=5)
fr4_run_button.grid(row=6, column=2, pady=5)
fr4_open_report_button.grid(row=7, column=2, pady=5, padx=5)

fr4_open_report_button['state'] = DISABLED


# ================================================================#
#                          Console Frame                          #
# ================================================================#

# Create Console Frame
console_frame = LabelFrame(root, text='Console', padx=5, pady=5)

# Add console Frame to main window
console_frame.grid(row=8, column=0, padx=20, pady=20, sticky=E+W+N+S)


#########################
#   Scrolled Textbox    #
#########################

# Create textbox
txtbox = scrolledtext.ScrolledText(console_frame, width=40, height=10)

# Add textbox to console frame
txtbox.grid(row=8, column=0, sticky=E+W+N+S)


class Redirect():

    def __init__(self, text):
        self.widget = text

    def write(self, text):
        self.widget.insert('end', text)
        self.widget.see(END)

    def flush(self):
        pass


sys.stdout = Redirect(txtbox)

#########################
#     configuration     #
#########################

console_frame.rowconfigure(8, weight=1)
console_frame.columnconfigure(0, weight=1)


#-------------------------#
#    Root configuration   #
#-------------------------#

root.columnconfigure(0, weight=1)
root.rowconfigure(8, weight=1)


# ================================================================#
#                           Menu Option                           #
# ================================================================#

# Create Programme menu
prog_menu = Menu(root)
root.config(menu=prog_menu)

# Create File menu
file_menu = Menu(prog_menu, tearoff=0)
prog_menu.add_cascade(label='File', menu=file_menu)
file_menu.add_command(label='New...', command=dummyCommand)
file_menu.add_separator()
file_menu.add_command(label='Exit', command=exit)

# Create Edit menu
edit_menu = Menu(prog_menu, tearoff=0)
prog_menu.add_cascade(label='Edit', menu=edit_menu)
edit_menu.add_command(label='Cut', command=dummyCommand)
edit_menu.add_command(label='Copy', command=dummyCommand)
edit_menu.add_command(label='Clear Console', command=clear_console)

# Create run menu
run_menu = Menu(prog_menu, tearoff=0)
prog_menu.add_cascade(label='Run', menu=run_menu)
run_menu.add_command(label='Run', command=lambda: run(
    fr4_run_button, fr4_stop_button, fr4_entry_dic))
run_menu.add_command(label='Start Debugging', command=dummyCommand)
run_menu.add_separator()
run_menu.add_command(label='Add Configuration', command=dummyCommand)

# Create Help menu
help_menu = Menu(prog_menu, tearoff=0)
prog_menu.add_cascade(label='Help', menu=help_menu)
help_menu.add_command(label='Welcome', command=dummyCommand)
help_menu.add_command(label='Documentation',
                      command=lambda: showFile(document_path, 'pdfFile'))
help_menu.add_separator()
help_menu.add_command(label='About', command=showAbout)


root.mainloop()

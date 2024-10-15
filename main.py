import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd
import modify_excel as mod
import output_excel as out

excel_file_path = 'empty'
root = tk.Tk()
root.geometry("800x500")
root.title("Process Shelving Excel file")
label = tk.Label(root, text="Set Excel File", font="'Arial', 16")
label.pack(padx=10, pady=10)

def show_success_popup(message):
    popup = tk.Toplevel()
    popup.title(message)
    popup.geometry("200x100")

    label = tk.Label(popup, text=message)
    label.pack(pady=20)

    button = tk.Button(popup, text="OK", command=popup.destroy)
    button.pack(pady=10)

    popup.mainloop()


def modify_call_number_function():
    global excel_file_path
    if excel_file_path != 'empty':
        mod.run_modify_excel(excel_file_path)
        show_success_popup('Excel Modified')


def output_errors_function():
    global excel_file_path
    if excel_file_path != 'empty':
        out.run_out_excel(excel_file_path)
        show_success_popup('Miss-shelved Created')


def select_file():
    filetypes = (
        ('Excel File', '*.xlsx'),
        ('Raw Text', '*.txt')
    )

    filename = fd.askopenfilename(
        title='Open a file',
        initialdir='/',
        filetypes=filetypes)

    label_set_file = tk.Label(root, text=filename, font="'Arial', 16")
    label_set_file.pack(padx=10, pady=10)
    global excel_file_path
    excel_file_path = filename


# open button
select_excel = ttk.Button(
    root,
    text='Select Excel File',
    command=select_file
)
select_excel.pack()


modify_call_numbers = ttk.Button(
    root,
    text='Create Normalized Call Numbers',
    command=modify_call_number_function
)
modify_call_numbers.pack()

output_errors = ttk.Button(
    root,
    text='Output Miss-Shelved',
    command=output_errors_function
)
output_errors.pack()

root.mainloop()

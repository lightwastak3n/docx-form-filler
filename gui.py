import sys
import tkinter as tk
from tkinter import filedialog as fd
from docx_editor import write_documents


window=tk.Tk()

window.title('Docx form filler')

instructions_text = """Select .docx template, .csv file that will be used for filling out the template and a folder in which files will be created. \
Write the formula to be used for file name (Example: $name_$lastname will produce $name_$lastname.docx files)."""
instructions = tk.Label(window, text=instructions_text, font="helvetica 13", wraplength=780)
instructions.grid(row=0, column=0, columnspan=3, padx=10, pady=(10, 10))


line1 = tk.Frame(window, height=2, width=750, bg="grey80", relief='groove')
line1.grid(row=1, column=0, columnspan=3, pady=5)

# .docx template
def select_docx_template():
    filetypes = (
        ('docx files', '*.docx'),
        ('All files', '*.*')
    )
    filename = fd.askopenfilename(
        title='Select a docx template',
        initialdir='/',
        filetypes=filetypes)
    docx_entry.delete(1, tk.END)
    docx_entry.insert(0, filename)

docx_path = tk.Label(text="Input .docx template file path", font="helvetica 12")
docx_path.grid(row=2, column=0, columnspan=2, pady=10)
docx_entry = tk.Entry(text="", width=80)
docx_entry.grid(row=3, column=0, padx=10)
docx_browse = tk.Button(text="Browse", width=20, command=select_docx_template)
docx_browse.grid(row=3, column=1)

line2 = tk.Frame(window, height=2, width=750, bg="grey80", relief='groove')
line2.grid(row=4, column=0, columnspan=3, pady=5)


# .csv data
def select_csv_file():
    filetypes = (
        ('csv files', '*.csv'),
        ('All files', '*.*')
    )
    filename = fd.askopenfilename(
        title='Select a csv data file',
        initialdir='/',
        filetypes=filetypes)
    csv_entry.delete(1, tk.END)
    csv_entry.insert(0, filename)

csv_path = tk.Label(text="Input .csv file path", font="helvetica 12")
csv_path.grid(row=5, column=0, columnspan=2, pady=10)
csv_entry = tk.Entry(text="", width=80)
csv_entry.grid(row=6, column=0, padx=10)
csv_browse = tk.Button(text="Browse", width=20, command=select_csv_file)
csv_browse.grid(row=6, column=1)

line3 = tk.Frame(window, height=2, width=750, bg="grey80", relief='groove')
line3.grid(row=7, column=0, columnspan=3, pady=5)

# Output folder
def select_output_folder():
    directory = fd.askdirectory(
        title='Select output directory',
        initialdir='/')
    output_entry.delete(1, tk.END)
    output_entry.insert(0, directory)

output_path = tk.Label(text="Output folder path", font="helvetica 12")
output_path.grid(row=8, column=0, columnspan=2, pady=10)
output_entry = tk.Entry(text="", width=80)
output_entry.grid(row=9, column=0, padx=10)
output_browse = tk.Button(text="Browse", width=20, command=select_output_folder)
output_browse.grid(row=9, column=1)

line4 = tk.Frame(window, height=2, width=750, bg="grey80", relief='groove')
line4.grid(row=10, column=0, columnspan=3, pady=5)

# Placeholders to use for name
name_label = tk.Label(text="Placeholders to use for file name", font="helvetica 12")
name_label.grid(row=11, column=0)
name_entry = tk.Entry(text="", width=80)
name_entry.grid(row=12, column=0, padx=10, pady=(5,10))

# Process data
def process_data():
    docx_template = docx_entry.get()
    csv_data = csv_entry.get()
    output_dir = output_entry.get()
    name_formula = name_entry.get()
    write_documents(docx_template, csv_data, name_formula, output_dir)


run_button = tk.Button(text ="RUN", command = process_data, width=30, font="helvetica 13", bg="blue", fg="yellow")
run_button.grid(row=12, column=1, pady=20)

output_files = tk.Label(text="List of output files", font="helvetica 12")
output_files.grid(row=13, column=0, columnspan=2)
textbox= tk.Text(window, height=5)
textbox.grid(row=14, column=0, columnspan=2, pady=(20,0))


def redirector(inputStr):
    textbox.insert(tk.INSERT, inputStr)
sys.stdout.write = redirector


window.mainloop()

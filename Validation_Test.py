import pandas as pd
import requests
from tkinter import Tk, Button, Label, filedialog, messagebox, StringVar
from tkinter.ttk import Progressbar
import threading

def link_validator(url):
    try:
        response = requests.head(url, allow_redirects=True)
        if response.status_code == 200:
            return "active"
        else:
            return "error"
    except requests.RequestException:
        return "Invalid URL / Connection Error"
        
    
def read_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])

    if file_path:
        global df
        try:
            df = pd.read_excel(file_path)
            messagebox.showinfo("Success", "File loaded successfully")
            
            threading.Thread(target=update_file).start()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file: {e}")

def update_file():
    if df is not None:
        link_column = 'Link'
        total_links = len(df)
        validated_links = 0

        for index,row in df.iterrows():
            df.at[index, 'Status'] = link_validator(row[link_column])
            validated_links += 1
            update_progress_bar(validated_links, total_links)

        messagebox.showinfo("Success", "Link validation completed")
        save_file()
    else:
        messagebox.showerror("Error", "No file loaded")
    
def update_progress_bar(validated, total):
    progress_var.set(f"Validated {validated}/{total} links")
    progress_bar['value'] = (validated / total) * 100


def save_file():
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        try:
            df.to_excel(file_path, index=False)
            messagebox.showinfo("Success", f"File saved successfully to {file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save file: {e}")




root = Tk()
root.title("Excel Hyperlink Validator")
root.geometry('400x200')

df = None

load_button = Button(root, text="Load and Validate File", command=read_file)
load_button.pack(pady=10)

progress_var = StringVar()
progress_label = Label(root, textvariable=progress_var)
progress_label.pack(pady=10)

progress_bar = Progressbar(root, orient="horizontal", length=300, mode="determinate")
progress_bar.pack(pady=10)

root.mainloop()

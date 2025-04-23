import re
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, Listbox, END

def extract_data_from_excel(file_path):
    try:
        df = pd.read_excel(file_path, engine='openpyxl')
        text = df.astype(str).apply(lambda x: ' '.join(x), axis=1).str.cat(sep=' ')
        
        # Extract all emails with .com, .org, .in, etc.
        emails = re.findall(r'\b[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}\b', text)
        
        # Phone number pattern remains the same
        phones = re.findall(r'(?:\+91[\-\s]?)?[6-9]\d{9}', text)
        
        return list(set(emails)), list(set(phones))
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read file:\n{e}")
        return [], []


# Open file dialog and extract info
def open_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        emails, phones = extract_data_from_excel(file_path)
        display_result(emails, phones)

# Display output in GUI
def display_result(emails, phones):
    email_listbox.delete(0, END)
    phone_output_box.delete(1.0, END)

    for email in emails:
        email_listbox.insert(END, email)

    phone_output_box.insert(tk.END, "📞 Contact Numbers:\n")
    phone_output_box.insert(tk.END, '\n'.join(phones))

# GUI setup
root = tk.Tk()
root.title("Excel Gmail & Contact Extractor")
root.geometry("700x550")

frame = tk.Frame(root)
frame.pack(pady=10)

upload_btn = tk.Button(frame, text="📂 Upload Excel File", command=open_file, font=('Arial', 12))
upload_btn.pack()

# Label and Listbox for Emails
email_label = tk.Label(root, text="📧 Gmail IDs:", font=('Arial', 12, 'bold'))
email_label.pack()

email_listbox = Listbox(root, width=80, height=10, font=('Courier', 10))
email_listbox.pack(pady=5)

# ScrolledText for Contact Numbers
phone_output_box = scrolledtext.ScrolledText(root, width=80, height=10, font=('Courier', 10))
phone_output_box.pack(pady=10)

root.mainloop()

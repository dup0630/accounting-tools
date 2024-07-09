import tkinter as tk
from tkinter import messagebox

def extract_action():
    file_name = entry.get()
    # You can add your file extraction logic here
    messagebox.showinfo("Action", f"Extracting {file_name}")

def cancel_action():
    app.destroy()

app = tk.Tk()
app.title("File Extractor")
app.geometry('300x150')

# Label
label = tk.Label(app, text="Enter file name:")
label.pack(pady=(20, 5))

# Entry widget for file name
entry = tk.Entry(app, width=30)
entry.pack(pady=(0, 20))

# Extract button
extract_button = tk.Button(app, text="Extract", command=extract_action)
extract_button.pack(side=tk.LEFT, padx=(50, 10))

# Cancel button
cancel_button = tk.Button(app, text="Cancel", command=cancel_action)
cancel_button.pack(side=tk.RIGHT, padx=(10, 50))

app.mainloop()

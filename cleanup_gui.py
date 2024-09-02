import tkinter as tk
from tkinter import messagebox
from cleanup import clean_outlook

def start_cleaning():
    email = email_var.get()
    folder_name = folder_var.get()
    days_old = int(days_var.get())
    subject_keyword = keyword_var.get()

    deleted_count = clean_outlook(email, folder_name, days_old, subject_keyword)
    
    messagebox.showinfo("Cleanup Complete", f"Deleted {deleted_count} emails from the {folder_name} folder.")

# Create the main window
root = tk.Tk()
root.title("Outlook Folder Cleaner")

# Create and place the input fields and labels
tk.Label(root, text="Your Email:").grid(row=0, column=0, padx=10, pady=10)
email_var = tk.StringVar()
email_entry = tk.Entry(root, textvariable=email_var)
email_entry.grid(row=0, column=1, padx=10, pady=10)

tk.Label(root, text="Folder Name:").grid(row=1, column=0, padx=10, pady=10)
folder_var = tk.StringVar()
folder_entry = tk.Entry(root, textvariable=folder_var)
folder_entry.grid(row=1, column=1, padx=10, pady=10)

tk.Label(root, text="Days Old:").grid(row=2, column=0, padx=10, pady=10)
days_var = tk.StringVar(value="30")  # Default value is 30 days
days_entry = tk.Entry(root, textvariable=days_var)
days_entry.grid(row=2, column=1, padx=10, pady=10)

tk.Label(root, text="Subject Keyword:").grid(row=3, column=0, padx=10, pady=10)
keyword_var = tk.StringVar()
keyword_entry = tk.Entry(root, textvariable=keyword_var)
keyword_entry.grid(row=3, column=1, padx=10, pady=10)

# Create and place the 'Clean' button
clean_button = tk.Button(root, text="Clean", command=start_cleaning)
clean_button.grid(row=4, column=0, columnspan=2, pady=20)

# Start the GUI event loop
root.mainloop()

import os
import comtypes.client
import tkinter as tk
from tkinter import filedialog, messagebox
import sys

def convert_pptx_to_pdf(input_folder, output_folder=None):
    if not output_folder:
        output_folder = input_folder

    # Initialize PowerPoint application
    powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
    powerpoint.Visible = 1  # PowerPoint must be visible for some versions to work

    # Get all files in the directory
    files = [f for f in os.listdir(input_folder) if f.endswith((".ppt", ".pptx"))]

    print(f"Found {len(files)} presentation files...")

    for filename in files:
        input_path = os.path.abspath(os.path.join(input_folder, filename))
        output_path = os.path.abspath(os.path.join(output_folder, os.path.splitext(filename)[0] + ".pdf"))
        
        # Debug: Print the full path to check for issues
        print(f"Processing: {input_path}")

        # Check if PDF already exists to avoid re-doing it
        if os.path.exists(output_path):
            print(f"Skipping {filename} (PDF already exists)")
            continue

        try:
            print(f"Converting: {filename}")
            deck = powerpoint.Presentations.Open(input_path)
            # 32 is the integer value for the ppSaveAsPDF format
            deck.SaveAs(output_path, 32)
            deck.Close()
        except Exception as e:
            print(f"Failed to convert {filename}: {e}")

    powerpoint.Quit()
    print("Batch conversion complete!")

if __name__ == "__main__":
    def start_gui():
        root = tk.Tk()
        root.title("PPTX to PDF Converter")
        root.geometry("300x150")

        def select_folder():
            folder_selected = filedialog.askdirectory()
            if folder_selected:
                try:
                    convert_pptx_to_pdf(folder_selected)
                    messagebox.showinfo("Success", "Batch conversion complete!")
                except Exception as e:
                    messagebox.showerror("Error", f"An error occurred:\n{str(e)}")

        label = tk.Label(root, text="Select a folder to convert pptx files to pdf")
        label.pack(pady=20)

        btn = tk.Button(root, text="Select Folder", command=select_folder)
        btn.pack(pady=10)

        root.mainloop()

    start_gui()
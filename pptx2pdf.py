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


def start_gui():
    root = tk.Tk()
    root.title("PPTX to PDF Converter")
    root.geometry("400x350")

    # Variables to store paths
    input_folder_var = tk.StringVar()
    output_folder_var = tk.StringVar()

    def select_input_folder():
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            input_folder_var.set(folder_selected)

    def select_output_folder():
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            output_folder_var.set(folder_selected)

    def run_conversion():
        input_folder = input_folder_var.get()
        output_folder = output_folder_var.get()

        if not input_folder:
            messagebox.showwarning("Warning", "Please select an input folder first.")
            return

        try:
            # If no output folder is selected, pass None to use input folder
            target_output = output_folder if output_folder else None
            convert_pptx_to_pdf(input_folder, target_output)
            messagebox.showinfo("Success", "Batch conversion complete!")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred:\n{str(e)}")

    # UI Layout
    # Input Section
    input_frame = tk.LabelFrame(root, text="Input", padx=10, pady=10)
    input_frame.pack(fill="x", padx=10, pady=5)

    tk.Label(input_frame, text="Source Folder containing PPTX files:").pack(anchor="w")
    tk.Entry(input_frame, textvariable=input_folder_var, width=50).pack(pady=5)
    tk.Button(input_frame, text="Select Input Folder", command=select_input_folder).pack(anchor="e")

    # Output Section
    output_frame = tk.LabelFrame(root, text="Output (Optional)", padx=10, pady=10)
    output_frame.pack(fill="x", padx=10, pady=5)

    tk.Label(output_frame, text="Destination Folder (leave empty for same as input):").pack(anchor="w")
    tk.Entry(output_frame, textvariable=output_folder_var, width=50).pack(pady=5)
    tk.Button(output_frame, text="Select Output Folder", command=select_output_folder).pack(anchor="e")

    # Action Section
    tk.Button(root, text="Convert to PDF", command=run_conversion, bg="#4CAF50", fg="white", font=("Arial", 10, "bold"), height=2).pack(fill="x", padx=20, pady=20)

    root.mainloop()

if __name__ == "__main__":
    start_gui()
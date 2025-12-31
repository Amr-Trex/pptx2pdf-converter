<h1 style="display: flex; align-items: center; gap: 10px; box-shadow: 0 4px 8px rgba(0, 0, 0, 0.3); padding: 10px; border-radius: 5px;"><img src="app_icon.ico" alt="PPTX to PDF Converter" width="64" height="64"> PPTX to PDF Converter</h1>

A simple Python utility with a Graphical User Interface (GUI) to batch convert PowerPoint presentations (`.pptx`, `.ppt`) into PDF format using the Microsoft PowerPoint COM interface.

***(NOTE: Project may be upgraded soon to be a fully fledged multi file type converter.).***

## How It Works

This script leverages the automation capabilities of Windows and Microsoft Office to perform high-fidelity conversions.

1.  **GUI Selection**: The application uses `tkinter` to launch a small window where users can select a target directory containing PowerPoint files.
2.  **COM Automation**: It uses the `comtypes` library to create an instance of the `PowerPoint.Application` object. This effectively "talks" to the installed PowerPoint application on your computer.
3.  **Batch Processing**:
    *   The script scans the selected folder for `.ppt` and `.pptx` files.
    *   It opens each presentation in PowerPoint (in the background).
    *   It invokes the built-in `SaveAs` method with the PDF file format code (File Format `32`).
    *   It saves the generated PDF in the same directory as the source file.
4.  **Cleanup**: Once all files are processed, it cleanly closes the PowerPoint application instance.

## Prerequisites

*   **Operating System**: Windows (required for COM interface).
*   **Software**: Microsoft PowerPoint must be installed on the machine.
*   **Python**: Python 3.x installed.

## Dependencies

You need to install the `comtypes` library to allow Python to interact with Windows applications.

```bash
pip install comtypes
```

*(Note: `tkinter` is included with standard Python installations).*

## Usage
### Using the Executable ***`MOST-RECOMMENDED`***
If you are using the standalone version, simply double-click `pptx2pdf.exe` from the `dist` folder to launch the application. This version does not require Python or any dependencies to be installed on your system.

### Running from Source


1.  **Run the script**:
    ```bash
    python pptx2pdf.py
    ```
2.  **Select Folders**:
    *   **Input**: Click **Select Input Folder** and choose the directory with your PPTX files.
    *   **Output (Optional)**: Click **Select Output Folder** to choose where to save the PDFs. If skipped, PDFs will be saved in the input folder.
    *   Click **Convert to PDF** to start the process.
3.  **Conversion**:
    *   The script will begin processing the files. You might see PowerPoint open/close briefly or run in the background.
    *   A popup message will verify when the batch conversion is complete.

## Troubleshooting

*   **"PowerPoint could not be found"**: Ensure Microsoft Office is installed and you are running on Windows.
*   **Permissions**: Make sure you have write permissions in the selected folder, as the PDFs are saved there.

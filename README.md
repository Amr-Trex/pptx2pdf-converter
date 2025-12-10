# PPTX to PDF Converter

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

1.  **Run the script**:
    ```bash
    python pptx2pdf.py
    ```
2.  **Select Folder**:
    *   A window titled "PPTX to PDF Converter" will appear.
    *   Click the **Select Folder** button.
    *   Navigate to and choose the folder containing your PowerPoint files.
3.  **Conversion**:
    *   The script will begin processing the files. You might see PowerPoint open/close briefly or run in the background.
    *   A popup message will verify when the batch conversion is complete.

## Troubleshooting

*   **"PowerPoint could not be found"**: Ensure Microsoft Office is installed and you are running on Windows.
*   **Permissions**: Make sure you have write permissions in the selected folder, as the PDFs are saved there.

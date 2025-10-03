# script-launcher-app/script-launcher-app/README.md

# Script Launcher App

This project provides a graphical user interface (GUI) for launching various Python scripts used for the various tasks of a Servipoli internship worker in the UPV (used the Computer Science Faculty as reference). The application is built using Tkinter and allows users to easily execute scripts from a simple interface.

## Project Structure

```
script-launcher-app
├── _internal
├── scripts
│   ├── blog
│   │   ├── BlogOfertasWordPrep.py
│   │   └── PDFFromMsgDownloader.py
│   ├── convenios
│   │   └── CreaCarpListenerReady.py
│   └── utilidades
│       ├── comprobar_letra_DNI.py
│       └── ExtractPDFtext.py
├── UPE Script Master.exe
├── requirements.txt
└── README.md
```

## Scripts Overview

- **BlogOfertasWordPrep.py**: Extracts job offers from PDF files and generates a Word document with the details.
- **PDFFromMsgDownloader.py**: Allows users to drag and drop .msg files to extract PDF attachments and save them to a specified directory.
- **CreaCarpListenerReady.py**: Monitors a specified folder for new PDF files, extracts relevant information, and organizes them into folders for each student.
- **comprobar_letra_DNI.py**: Provides a GUI to calculate the letter corresponding to a Spanish DNI number.
- **ExtractPDFtext.py**: Allows users to drag and drop PDF files to extract text and save it to a text file.

## Installation

1. Extract the _internal folder.
2. Clone the repository:

   ```
   git clone <(https://github.com/CarlosRzUb/UPE_Script_Master)>
   ```
3. Install the required dependencies:

   ```
   pip install -r requirements.txt
   ```

## Usage

To run the application, execute the following command:

```
python launcher/main_gui.py
```

This will open the GUI where you can select and execute the various scripts available in the `scripts` directory.

## Requirements

The following Python packages are required for this project:

- Tkinter
- PyMuPDF
- PyPDF2
- watchdog
- extract_msg
- python-docx
- Pillow
- pywin32

Make sure to install these packages before running the application.

## License

This project is licensed under the MIT License - see the LICENSE file for details.

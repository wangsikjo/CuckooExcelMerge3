# ExcelPDFPortable

- Drag&drop Excel files, pick sheets, merge, and export to PDF.
- Portable EXE build via GitHub Actions.

## Local build
```
py -m pip install --upgrade pip
py -m pip install pyinstaller pyqt5 openpyxl xlrd pywin32
py -m PyInstaller --onefile --windowed --name ExcelPDFPortable ExcelPDFPortable.py
```

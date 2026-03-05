# BTG/XP Processor GUI Application

This application processes BTG and XP bank statements (PDFs) and reconciles them with an Excel template.

## Files

- `processor_gui.py` - Main GUI application
- `BTG_processor_core.py` - BTG processing logic
- `XP_processor_core.py` - XP processing logic
- `requirements.txt` - Python dependencies
- `build_exe.bat` - Script to build Windows executable

## How to Use

### Option 1: Run as Python Script

1. Install dependencies:
   ```
   pip install -r requirements.txt
   ```

2. Run the application:
   ```
   python processor_gui.py
   ```

### Option 2: Build and Run as .exe

1. Run the build script:
   ```
   build_exe.bat
   ```

2. The executable will be created in the `dist` folder:
   ```
   dist\BTG-XP-Processor.exe
   ```

3. Double-click the .exe to run the application

## Using the Application

1. **Add PDF Files**: Click "Add PDF(s)" to select one or more PDF statements
   - BTG PDFs must have "BTG" in the filename
   - XP PDFs must have "XP" in the filename

2. **Select Excel File**: Click "Browse" to select your Excel template file

3. **Choose Output Directory** (Optional): By default, output will be saved in the same location as the Excel file

4. **Run Processing**:
   - Click "RUN BTG" to process only BTG statements
   - Click "RUN XP" to process only XP statements
   - Click "RUN BOTH" to process both sequentially

5. The application will create/update `output.xlsx` with the reconciled data

## Notes

- If `output.xlsx` already exists, the application will update it in place
- This allows you to run BTG and XP processors sequentially on the same file
- Progress and status messages are shown in the application window

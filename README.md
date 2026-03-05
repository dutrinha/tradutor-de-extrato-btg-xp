# BTG/XP Processor GUI Application

This application processes BTG and XP bank statements (PDFs) and reconciles them with an Excel template.

## Files

- `BTG_processor.py` - BTG processing logic
- `XP_processor.py` - XP processing logic
- `requirements.txt` - Python dependencies

## How to Use

### Option 1: Run as Python Script

1. Install dependencies:
   ```
   pip install -r requirements.txt
   ```

2. Run the specific script for the bank:
   ```
   py BTG-processor.py
   or
   py XP-processor.py

5. The application will create/update `output.xlsx` with the reconciled data

## Notes

- If `output.xlsx` already exists, the application will update it in place
- This allows you to run BTG and XP processors sequentially on the same file
- Progress and status messages are shown in the application window


# Air Canada Aeroplan Registration Bot - GUI Version

A PyQt6-based GUI application to automate Air Canada Aeroplan account registration from CSV data.

## Features

✅ **Modern GUI Interface**
- Easy-to-use graphical interface
- Real-time log viewer with color-coded messages
- Progress bar showing current status

✅ **CSV Management**
- Browse and select CSV files
- Automatically removes processed rows from original CSV
- Resume capability - skips already processed entries

✅ **Logging System**
- Real-time logs with timestamps
- Color-coded messages (Info, Success, Error, Warning)
- Clear logs functionality

✅ **Proxy Support**
- Rotating proxy with random ports (10000-20000)
- Each account gets a fresh IP
- Configurable proxy settings

✅ **Processing Features**
- Processes entries one by one
- Saves results to `processed_data.csv`
- Extracts and saves Aeroplan numbers
- Error handling and resume capability

## Installation

### 1. Install Python Dependencies

```bash
pip install -r requirements.txt
```

Or install manually:

```bash
pip install PyQt6
pip install seleniumbase
```

### 2. Run the GUI Application

```bash
python gui.py
```

## Usage

### Step 1: Configure Proxy
- The default rotating proxy is pre-configured
- You can modify it in the "Proxy Base" field if needed
- Format: `username:password@host`

### Step 2: Select CSV File
1. Click the "📁 Browse CSV File" button
2. Select your CSV file with registration data
3. The application will show the number of entries found

### Step 3: Start Processing
1. Click "▶️ Start Processing" button
2. The bot will open browsers and fill forms automatically
3. Watch the real-time logs for progress
4. Each processed entry is removed from the original CSV

### Step 4: Monitor Progress
- Progress bar shows current status
- Logs display detailed information
- Stop processing anytime with "⏸️ Stop Processing"

## CSV Format

Your CSV file should have the following columns:

```csv
Email,Password,First name,Last name,Gender,Language,Day,Month,Year,Address,Postal Code,City,Province,Phone number,Country,Proxy,Aeroplan number
```

Example:
```csv
john@example.com,Pass123,John,Doe,Male,English,15,Jun,1990,123 Main St,V5K 0A1,Vancouver,British Columbia,6041234567,Canada,,
```

## Output Files

### processed_data.csv
Contains all successfully processed entries with:
- All original CSV data
- `Processed_Date` - Timestamp of processing
- `Aeroplan number` - Extracted membership number

## Features Explained

### 🔄 Auto-Resume
- The bot tracks processed emails
- If you stop and restart, it automatically skips completed entries
- Original CSV is cleaned up as entries are processed

### 🌐 Rotating Proxy
- Each account uses a random port (10000-20000)
- This gives each account a fresh IP address
- Reduces detection and blocking

### 📝 Real-Time Logs
- **White**: General information
- **Green**: Success messages
- **Red**: Error messages
- **Yellow**: Warning messages

### 🛑 Stop & Resume
- Click "Stop Processing" to pause
- Processed entries are already saved
- Restart the application and click "Start" to continue

## Troubleshooting

### CAPTCHA Handling
The bot will automatically wait for CAPTCHA to be solved:
- It detects reCAPTCHA on the page
- Waits up to 5 minutes for you to solve it manually
- Once solved, continues automatically

### Browser Issues
If you see Chrome-related errors:
- The bot uses your system Chrome by default
- Or uses the custom Chrome binary if available
- Make sure Chrome is installed and updated

### Proxy Issues
If proxy connection fails:
- Check your proxy credentials
- Verify the proxy server is working
- Try changing the port range if needed

## Tips

1. **Run in Safe Mode**: Process a few entries first to test
2. **Monitor Logs**: Watch for errors or warnings
3. **Backup CSV**: Keep a copy of your original CSV file
4. **Stable Internet**: Ensure good internet connection
5. **Don't Close Browser**: Let the automation complete each entry

## File Structure

```
airport/
├── gui.py                    # Main GUI application
├── main.py                   # Original CLI version (still works)
├── requirements.txt          # Python dependencies
├── Test data.csv            # Your input CSV file
├── processed_data.csv       # Output with results
└── extension/               # Browser extension (optional)
```

## Command Line Version

The original `main.py` still works for command-line usage:

```bash
python main.py
```

## Support

For issues or questions:
1. Check the logs for detailed error messages
2. Verify your CSV format matches the template
3. Ensure all required fields are filled
4. Test with a single entry first

---

**Note**: This tool automates form filling. Always ensure you comply with Air Canada's terms of service and use responsibly.

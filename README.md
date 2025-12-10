# Excel IP Ping Scanner

A Python GUI tool that:
- Reads an Excel file
- Detects colored cells in column "LAN IP"
- Pings each IP
- Also increments last octet and pings again
- Writes a detailed ping report in `ping_report.txt`

## Features
- GUI file selection (Tkinter)
- Reads Excel with openpyxl
- Detects specific cell background color (yellow)
- Pings using Windows ping command
- Writes detailed results to a text file

## How to Run
pip install -r requirements.txt
python main.py
# S7 Logging and Graphing

This project provides a simple Python GUI tool for connecting to a Siemens S7 PLC using the `snap7` library.  It polls data block values, plots the readings live and can export the collected data to an Excel file. Logged values are also written to a temporary file so logging can continue after a crash.

## Features
- Connect to a PLC by entering IP, rack and slot
- Log multiple addresses simultaneously; use **+** and **-** to add or remove variables
- Read from any data block and address with selectable data type (BOOL, INT, DINT or REAL)
- Poll at a custom interval and display live values on a matplotlib graph with a legend showing each address and its last value
- Choose between line and scatter plots
- Start, stop and resume logging via dedicated buttons
- Export all collected values to an Excel (`.xlsx`) file
- Clear logged data via the **Clear** button (with confirmation) which also removes the temporary log file

## Requirements
Install dependencies:

```bash
pip install python-snap7 pandas matplotlib openpyxl
```

## Usage

```bash
python plc_logger_gui.py
```

1. Enter the PLC connection parameters and data block details.
2. Add or remove addresses using **+** and **-** then click **Connect** to establish the connection.
3. Choose the polling interval and graph type then press **Start** to begin logging. Use **Stop** and **Resume** to control polling.
4. Use **Clear** to reset all data or **Export** to save the collected data to an Excel file.


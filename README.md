# S7 Logging and Graphing

This project provides a simple Python GUI tool for connecting to a Siemens S7 PLC using the `snap7` library.  It polls a data block value, plots the readings live and can export the collected data to an Excel file.

## Features
- Connect to a PLC by entering IP, rack and slot
- Read from any data block and address with selectable data type (BOOL, INT, DINT or REAL)
- Poll at a custom interval and display live values on a matplotlib graph
- Choose between line and scatter plots
- Export all collected values to an Excel (`.xlsx`) file

## Requirements
Install dependencies:

```bash
pip install python-snap7 pandas matplotlib
```

## Usage

```bash
python plc_logger_gui.py
```

1. Enter the PLC connection parameters and data block details.
2. Click **Connect** to establish the connection.
3. Choose the polling interval and graph type then press **Start** to begin logging.
4. Click **Export** to save the collected data to an Excel file.


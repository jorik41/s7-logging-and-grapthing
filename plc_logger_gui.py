import threading
import time
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

import matplotlib
matplotlib.use('TkAgg')
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

try:
    import snap7
except ImportError:  # pragma: no cover - module may not be installed in test env
    snap7 = None

import pandas as pd


def _parse_data(data: bytes, data_type: str):
    """Parse raw bytes from PLC into the chosen data type."""
    if data_type == "BOOL":
        return bool(data[0])
    if data_type == "INT":
        return int.from_bytes(data[0:2], byteorder="big", signed=True)
    if data_type == "DINT":
        return int.from_bytes(data[0:4], byteorder="big", signed=True)
    if data_type == "REAL":
        import struct
        return struct.unpack(">f", data[0:4])[0]
    raise ValueError(f"Unsupported data type: {data_type}")


class PLCLoggerGUI:
    def __init__(self, master: tk.Tk) -> None:
        self.master = master
        master.title("Snap7 PLC Logger")

        # Connection parameters
        self.ip_var = tk.StringVar(value="192.168.0.1")
        self.rack_var = tk.StringVar(value="0")
        self.slot_var = tk.StringVar(value="1")

        # DB read parameters
        self.db_var = tk.StringVar(value="1")
        self.start_var = tk.StringVar(value="0")
        self.data_type_var = tk.StringVar(value="REAL")

        # Polling interval in seconds
        self.interval_var = tk.StringVar(value="1.0")

        # Graph type
        self.graph_type_var = tk.StringVar(value="Line")

        self.client = None
        self.polling = False
        self.data = []

        self._build_ui()

    def _build_ui(self) -> None:
        frame = ttk.Frame(self.master)
        frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        # Connection row
        conn_frame = ttk.LabelFrame(frame, text="Connection")
        conn_frame.pack(fill=tk.X)
        ttk.Label(conn_frame, text="IP").grid(column=0, row=0, sticky=tk.W)
        ttk.Entry(conn_frame, textvariable=self.ip_var, width=15).grid(column=1, row=0)
        ttk.Label(conn_frame, text="Rack").grid(column=2, row=0)
        ttk.Entry(conn_frame, textvariable=self.rack_var, width=5).grid(column=3, row=0)
        ttk.Label(conn_frame, text="Slot").grid(column=4, row=0)
        ttk.Entry(conn_frame, textvariable=self.slot_var, width=5).grid(column=5, row=0)
        ttk.Button(conn_frame, text="Connect", command=self.connect).grid(column=6, row=0, padx=5)

        # Data parameters
        data_frame = ttk.LabelFrame(frame, text="Data")
        data_frame.pack(fill=tk.X, pady=5)
        ttk.Label(data_frame, text="DB").grid(column=0, row=0)
        ttk.Entry(data_frame, textvariable=self.db_var, width=5).grid(column=1, row=0)
        ttk.Label(data_frame, text="Start").grid(column=2, row=0)
        ttk.Entry(data_frame, textvariable=self.start_var, width=5).grid(column=3, row=0)
        ttk.Label(data_frame, text="Type").grid(column=4, row=0)
        ttk.Combobox(
            data_frame,
            values=["BOOL", "INT", "DINT", "REAL"],
            textvariable=self.data_type_var,
            width=6,
            state="readonly",
        ).grid(column=5, row=0)
        ttk.Label(data_frame, text="Interval (s)").grid(column=6, row=0)
        ttk.Entry(data_frame, textvariable=self.interval_var, width=5).grid(column=7, row=0)
        ttk.Label(data_frame, text="Graph").grid(column=8, row=0)
        ttk.Combobox(
            data_frame,
            values=["Line", "Scatter"],
            textvariable=self.graph_type_var,
            width=8,
            state="readonly",
        ).grid(column=9, row=0)
        ttk.Button(data_frame, text="Start", command=self.toggle_polling).grid(column=10, row=0, padx=5)
        ttk.Button(data_frame, text="Export", command=self.export_excel).grid(column=11, row=0)

        # Matplotlib figure
        self.fig = Figure(figsize=(6, 4))
        self.ax = self.fig.add_subplot(111)
        self.ax.set_xlabel("Time")
        self.ax.set_ylabel("Value")
        self.canvas = FigureCanvasTkAgg(self.fig, master=frame)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    def connect(self) -> None:
        if snap7 is None:
            messagebox.showerror(
                "snap7 missing", "The python-snap7 package is required to connect to the PLC."
            )
            return
        if self.client:
            self.client.disconnect()
            self.client = None
        try:
            client = snap7.client.Client()
            client.connect(self.ip_var.get(), int(self.rack_var.get()), int(self.slot_var.get()))
            if not client.get_connected():
                raise RuntimeError("Could not connect to PLC")
            self.client = client
            messagebox.showinfo("Connected", "Connected to PLC successfully")
        except Exception as exc:  # pragma: no cover - runtime path
            messagebox.showerror("Connection error", str(exc))

    def toggle_polling(self) -> None:
        if not self.client:
            messagebox.showwarning("Not connected", "Connect to the PLC first")
            return
        self.polling = not self.polling
        if self.polling:
            threading.Thread(target=self._poll, daemon=True).start()

    def _poll(self) -> None:
        start = int(self.start_var.get())
        db = int(self.db_var.get())
        interval = float(self.interval_var.get())
        data_type = self.data_type_var.get()
        read_len = {"BOOL": 1, "INT": 2, "DINT": 4, "REAL": 4}[data_type]
        while self.polling:
            try:
                raw = self.client.db_read(db, start, read_len)
                value = _parse_data(raw, data_type)
                self.data.append((datetime.now(), value))
                self._update_plot()
            except Exception as exc:  # pragma: no cover - runtime path
                self.polling = False
                messagebox.showerror("Read error", str(exc))
                break
            time.sleep(interval)

    def _update_plot(self) -> None:
        times = [t for t, _ in self.data]
        values = [v for _, v in self.data]
        self.ax.cla()
        self.ax.set_xlabel("Time")
        self.ax.set_ylabel("Value")
        if self.graph_type_var.get() == "Scatter":
            self.ax.scatter(times, values, s=10)
        else:
            self.ax.plot(times, values, linestyle="-", marker="")
        self.fig.autofmt_xdate()
        self.canvas.draw_idle()

    def export_excel(self) -> None:
        if not self.data:
            messagebox.showwarning("No data", "Nothing to export yet")
            return
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")]
        )
        if file_path:
            df = pd.DataFrame(self.data, columns=["timestamp", "value"])
            df.to_excel(file_path, index=False)
            messagebox.showinfo("Exported", f"Data exported to {file_path}")


def main() -> None:
    root = tk.Tk()
    app = PLCLoggerGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()

import os
import threading
import time
from datetime import datetime
import tempfile
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import matplotlib

matplotlib.use("TkAgg")
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

        # Polling interval in seconds
        self.interval_var = tk.StringVar(value="1.0")

        # Graph type
        self.graph_type_var = tk.StringVar(value="Line")

        self.client = None
        self.polling = False

        # Dynamic variable rows
        self.variable_rows = []
        self.active_vars = []

        # Data storage
        self.temp_file = os.path.join(tempfile.gettempdir(), "plc_logger_temp.csv")
        self.data_df = pd.DataFrame()

        self._build_ui()
        self._load_temp_data()
        if not self.variable_rows:
            self._add_variable_row()

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

        # Variable rows container
        self.vars_frame = ttk.Frame(data_frame)
        self.vars_frame.grid(column=0, row=0, columnspan=12, sticky=tk.W)

        # Controls row
        ttk.Button(data_frame, text="+", command=self._add_variable_row).grid(column=0, row=1)
        ttk.Button(data_frame, text="-", command=self._remove_variable_row).grid(column=1, row=1)
        ttk.Label(data_frame, text="Interval (s)").grid(column=2, row=1)
        ttk.Entry(data_frame, textvariable=self.interval_var, width=5).grid(column=3, row=1)
        ttk.Label(data_frame, text="Graph").grid(column=4, row=1)
        ttk.Combobox(
            data_frame,
            values=["Line", "Scatter"],
            textvariable=self.graph_type_var,
            width=8,
            state="readonly",
        ).grid(column=5, row=1)
        ttk.Button(data_frame, text="Start", command=self.start_polling).grid(column=6, row=1, padx=5)
        ttk.Button(data_frame, text="Stop", command=self.stop_polling).grid(column=7, row=1)
        ttk.Button(data_frame, text="Resume", command=self.resume_polling).grid(column=8, row=1)
        ttk.Button(data_frame, text="Export", command=self.export_excel).grid(column=9, row=1)
        ttk.Button(data_frame, text="Clear", command=self.clear_data).grid(column=10, row=1)

        # Matplotlib figure
        self.fig = Figure(figsize=(6, 4))
        self.ax = self.fig.add_subplot(111)
        self.ax.set_xlabel("Time")
        self.ax.set_ylabel("Value")
        self.canvas = FigureCanvasTkAgg(self.fig, master=frame)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    def _add_variable_row(self, db: str = "1", start: str = "0", data_type: str = "REAL") -> None:
        row = len(self.variable_rows)
        db_var = tk.StringVar(value=db)
        start_var = tk.StringVar(value=start)
        type_var = tk.StringVar(value=data_type)

        widgets = []
        widgets.append(ttk.Label(self.vars_frame, text="DB"))
        widgets[-1].grid(column=0, row=row)
        widgets.append(ttk.Entry(self.vars_frame, textvariable=db_var, width=5))
        widgets[-1].grid(column=1, row=row)
        widgets.append(ttk.Label(self.vars_frame, text="Start"))
        widgets[-1].grid(column=2, row=row)
        widgets.append(ttk.Entry(self.vars_frame, textvariable=start_var, width=5))
        widgets[-1].grid(column=3, row=row)
        widgets.append(ttk.Label(self.vars_frame, text="Type"))
        widgets[-1].grid(column=4, row=row)
        widgets.append(
            ttk.Combobox(
                self.vars_frame,
                values=["BOOL", "INT", "DINT", "REAL"],
                textvariable=type_var,
                width=6,
                state="readonly",
            )
        )
        widgets[-1].grid(column=5, row=row)

        self.variable_rows.append(
            {"db": db_var, "start": start_var, "type": type_var, "widgets": widgets}
        )

    def _remove_variable_row(self) -> None:
        if not self.variable_rows:
            return
        row = self.variable_rows.pop()
        for widget in row["widgets"]:
            widget.destroy()

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


    def start_polling(self) -> None:
        if not self.client:
            messagebox.showwarning("Not connected", "Connect to the PLC first")
            return
        if self.polling:
            return
        self.active_vars = []
        for row in self.variable_rows:
            self.active_vars.append(
                (
                    int(row["db"].get()),
                    int(row["start"].get()),
                    row["type"].get(),
                )
            )
        self.polling = True
        threading.Thread(target=self._poll, daemon=True).start()

    def stop_polling(self) -> None:
        self.polling = False

    def resume_polling(self) -> None:
        if not self.client:
            messagebox.showwarning("Not connected", "Connect to the PLC first")
            return
        if self.polling:
            return
        if not self.active_vars:
            self.start_polling()
            return
        self.polling = True
        threading.Thread(target=self._poll, daemon=True).start()

    def _poll(self) -> None:
        interval = float(self.interval_var.get())
        read_settings = []
        for db, start, data_type in self.active_vars:
            read_len = {"BOOL": 1, "INT": 2, "DINT": 4, "REAL": 4}[data_type]
            read_settings.append((db, start, data_type, read_len))
        while self.polling:
            row = {"timestamp": datetime.now()}
            try:
                for db, start, data_type, read_len in read_settings:
                    raw = self.client.db_read(db, start, read_len)
                    value = _parse_data(raw, data_type)
                    col = f"DB{db}_{start}_{data_type}"
                    row[col] = value
                self.data_df = pd.concat([self.data_df, pd.DataFrame([row])], ignore_index=True)
                self._append_temp_file(pd.DataFrame([row]))
                self._update_plot()
            except Exception as exc:  # pragma: no cover - runtime path
                self.polling = False
                messagebox.showerror("Read error", str(exc))
                break
            time.sleep(interval)

    def _update_plot(self) -> None:
        self.ax.cla()
        self.ax.set_xlabel("Time")
        self.ax.set_ylabel("Value")
        if not self.data_df.empty:
            times = self.data_df["timestamp"]
            for col in self.data_df.columns:
                if col == "timestamp":
                    continue
                parts = col.split("_")
                values = self.data_df[col]
                last_val = values.iloc[-1]
                label = f"{parts[0]}.{parts[1]} ({last_val})" if len(parts) >= 2 else f"{col} ({last_val})"
                if self.graph_type_var.get() == "Scatter":
                    self.ax.scatter(times, values, s=10, label=label)
                else:
                    self.ax.plot(times, values, linestyle="-", marker="", label=label)
            self.ax.legend()
        self.fig.autofmt_xdate()
        self.canvas.draw_idle()

    def export_excel(self) -> None:
        if self.data_df.empty:
            messagebox.showwarning("No data", "Nothing to export yet")
            return
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")]
        )
        if file_path:
            try:
                with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
                    self.data_df.to_excel(writer, index=False)
            except ImportError:
                messagebox.showerror("Export error", "openpyxl is required to export Excel files")
                return
            messagebox.showinfo("Exported", f"Data exported to {file_path}")

    def clear_data(self) -> None:
        if not messagebox.askyesno("Clear data", "Are you sure you want to clear the data?"):
            return
        self.data_df = pd.DataFrame()
        if os.path.exists(self.temp_file):
            os.remove(self.temp_file)
        self._update_plot()

    def _append_temp_file(self, df_row: pd.DataFrame) -> None:
        header = not os.path.exists(self.temp_file)
        df_row.to_csv(self.temp_file, mode="a", index=False, header=header)

    def _load_temp_data(self) -> None:
        if os.path.exists(self.temp_file):
            try:
                self.data_df = pd.read_csv(self.temp_file, parse_dates=["timestamp"])
                for col in self.data_df.columns:
                    if col == "timestamp":
                        continue
                    parts = col.split("_")
                    if len(parts) == 3:
                        db = parts[0].replace("DB", "")
                        start, data_type = parts[1], parts[2]
                        self._add_variable_row(db, start, data_type)
                if not self.data_df.empty:
                    self._update_plot()
            except Exception:  # pragma: no cover - defensive
                self.data_df = pd.DataFrame()


def main() -> None:
    root = tk.Tk()
    app = PLCLoggerGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()

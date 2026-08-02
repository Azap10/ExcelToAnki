"""Desktop UI for importing Excel vocabulary into Anki's Add window."""

from __future__ import annotations

import queue
import threading
import time
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

REQUIRED_COLUMNS = ("Chinese", "Pinyin", "English", "Lesson", "Character")
DEFAULT_COORDINATES = {
    "Front": "550,200",
    "Back": "550,300",
    "Pinyin": "550,400",
    "Add": "550,1110",
}


class ExcelToAnkiApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Excel to Anki")
        self.root.minsize(850, 620)
        self.dataframe = None
        self.filtered_rows = None
        self.lesson_values: dict[str, object] = {}
        self.card_type_values: dict[str, object] = {}
        self.stop_requested = threading.Event()
        self.events: queue.Queue[tuple[str, str]] = queue.Queue()

        self.file_path = tk.StringVar()
        self.lesson = tk.StringVar()
        self.card_type = tk.StringVar()
        self.delay = tk.StringVar(value="1.0")
        self.status = tk.StringVar(value="Choose an Excel workbook to begin.")
        self.coordinates = {name: tk.StringVar(value=value) for name, value in DEFAULT_COORDINATES.items()}

        self._build_interface()
        self.root.after(100, self._process_events)

    def _build_interface(self) -> None:
        container = ttk.Frame(self.root, padding=16)
        container.grid(sticky="nsew")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        container.columnconfigure(1, weight=1)
        container.rowconfigure(3, weight=1)

        ttk.Label(container, text="Excel to Anki", font=("Segoe UI", 18, "bold")).grid(column=0, row=0, columnspan=3, sticky="w")
        ttk.Label(container, text="Preview vocabulary from an Excel file, then send selected cards to Anki's Add window.").grid(column=0, row=1, columnspan=3, sticky="w", pady=(2, 14))

        ttk.Label(container, text="Workbook").grid(column=0, row=2, sticky="w")
        ttk.Entry(container, textvariable=self.file_path, state="readonly").grid(column=1, row=2, sticky="ew", padx=8)
        ttk.Button(container, text="Choose Excel file", command=self.choose_file).grid(column=2, row=2)

        filters = ttk.LabelFrame(container, text="Filter cards", padding=10)
        filters.grid(column=0, row=3, columnspan=3, sticky="ew", pady=(16, 8))
        filters.columnconfigure(1, weight=1)
        filters.columnconfigure(3, weight=1)
        ttk.Label(filters, text="Lesson").grid(column=0, row=0, sticky="w")
        self.lesson_box = ttk.Combobox(filters, textvariable=self.lesson, state="readonly")
        self.lesson_box.grid(column=1, row=0, sticky="ew", padx=(8, 18))
        ttk.Label(filters, text="Card type").grid(column=2, row=0, sticky="w")
        self.type_box = ttk.Combobox(filters, textvariable=self.card_type, state="readonly")
        self.type_box.grid(column=3, row=0, sticky="ew", padx=8)
        ttk.Button(filters, text="Preview matching cards", command=self.prepare_cards).grid(column=4, row=0, padx=(10, 0))

        preview = ttk.LabelFrame(container, text="Preview", padding=8)
        preview.grid(column=0, row=4, columnspan=3, sticky="nsew", pady=8)
        container.rowconfigure(4, weight=1)
        preview.columnconfigure(0, weight=1)
        preview.rowconfigure(0, weight=1)
        self.cards = ttk.Treeview(preview, columns=("Chinese", "Pinyin", "English"), show="headings", height=10)
        for column, width in (("Chinese", 200), ("Pinyin", 220), ("English", 340)):
            self.cards.heading(column, text=column)
            self.cards.column(column, width=width, anchor="w")
        scrollbar = ttk.Scrollbar(preview, orient="vertical", command=self.cards.yview)
        self.cards.configure(yscrollcommand=scrollbar.set)
        self.cards.grid(column=0, row=0, sticky="nsew")
        scrollbar.grid(column=1, row=0, sticky="ns")

        automation = ttk.LabelFrame(container, text="Anki automation", padding=10)
        automation.grid(column=0, row=5, columnspan=3, sticky="ew", pady=8)
        for index, name in enumerate(DEFAULT_COORDINATES):
            ttk.Label(automation, text=f"{name} position").grid(column=index * 2, row=0, sticky="w")
            ttk.Entry(automation, textvariable=self.coordinates[name], width=12).grid(column=index * 2 + 1, row=0, padx=(5, 12))
        ttk.Label(automation, text="Delay (seconds)").grid(column=0, row=1, sticky="w", pady=(10, 0))
        ttk.Entry(automation, textvariable=self.delay, width=12).grid(column=1, row=1, sticky="w", padx=(5, 12), pady=(10, 0))
        self.start_button = ttk.Button(automation, text="Start sending cards", command=self.start_automation)
        self.start_button.grid(column=6, row=1, columnspan=2, sticky="e", pady=(10, 0))
        ttk.Button(automation, text="Stop", command=self.stop_automation).grid(column=8, row=1, sticky="e", padx=(8, 0), pady=(10, 0))

        ttk.Label(container, textvariable=self.status, foreground="#245a36").grid(column=0, row=6, columnspan=3, sticky="w", pady=(6, 0))

    def choose_file(self) -> None:
        selected = filedialog.askopenfilename(title="Choose vocabulary workbook", filetypes=[("Excel workbooks", "*.xlsx *.xls"), ("All files", "*.*")])
        if not selected:
            return
        try:
            import pandas as pd

            dataframe = pd.read_excel(selected)
            missing = [column for column in REQUIRED_COLUMNS if column not in dataframe.columns]
            if missing:
                raise ValueError("Missing required columns: " + ", ".join(missing))
        except Exception as error:
            messagebox.showerror("Could not load workbook", str(error))
            return

        self.dataframe = dataframe
        self.file_path.set(selected)
        lessons = sorted(dataframe["Lesson"].dropna().unique())
        card_types = sorted(dataframe["Character"].dropna().unique(), key=str)
        self.lesson_values = {self._lesson_label(value): value for value in lessons}
        self.card_type_values = {self._card_type_label(value): value for value in card_types}
        self.lesson_box["values"] = list(self.lesson_values)
        self.type_box["values"] = list(self.card_type_values)
        if lessons:
            self.lesson.set(self._lesson_label(lessons[0]))
        if card_types:
            self.card_type.set(self._card_type_label(card_types[0]))
        self.filtered_rows = None
        self._clear_preview()
        self.status.set(f"Loaded {Path(selected).name}: {len(dataframe)} rows.")

    def prepare_cards(self) -> None:
        if self.dataframe is None:
            messagebox.showinfo("Choose a workbook", "Choose an Excel workbook first.")
            return
        lesson_label = self.lesson.get()
        card_type_label = self.card_type.get()
        if not lesson_label or not card_type_label:
            messagebox.showinfo("Select filters", "Select both a lesson and a card type.")
            return
        lesson = self.lesson_values[lesson_label]
        card_type = self.card_type_values[card_type_label]
        rows = self.dataframe[
            (self.dataframe["Lesson"] == lesson) & (self.dataframe["Character"] == card_type)
        ][["Chinese", "Pinyin", "English"]].dropna()
        self.filtered_rows = rows
        self._clear_preview()
        for _, row in rows.iterrows():
            self.cards.insert("", "end", values=(row["Chinese"], row["Pinyin"], row["English"]))
        self.status.set(f"{len(rows)} matching card(s) ready to send.")

    def _clear_preview(self) -> None:
        for item in self.cards.get_children():
            self.cards.delete(item)

    @staticmethod
    def _lesson_label(value: object) -> str:
        numeric_value = float(value)
        return str(int(numeric_value)) if numeric_value.is_integer() else str(value)

    @staticmethod
    def _card_type_label(value: object) -> str:
        if str(value).lower() == "true":
            return "Characters"
        if str(value).lower() == "false":
            return "Words"
        return str(value)

    def start_automation(self) -> None:
        if self.filtered_rows is None or self.filtered_rows.empty:
            messagebox.showinfo("Prepare cards", "Preview matching cards before starting automation.")
            return
        try:
            positions = {name: self._parse_position(value.get()) for name, value in self.coordinates.items()}
            delay = float(self.delay.get())
            if delay < 0:
                raise ValueError("Delay cannot be negative.")
        except ValueError as error:
            messagebox.showerror("Invalid automation settings", str(error))
            return
        self.stop_requested.clear()
        self.start_button.configure(state="disabled")
        threading.Thread(target=self._send_cards, args=(positions, delay), daemon=True).start()
        self.status.set("Starting in 3 seconds—focus Anki's Add window now.")

    @staticmethod
    def _parse_position(value: str) -> tuple[int, int]:
        try:
            x, y = (part.strip() for part in value.split(","))
            return int(x), int(y)
        except (TypeError, ValueError) as error:
            raise ValueError("Positions must use the format x,y (for example, 550,200).") from error

    def _send_cards(self, positions: dict[str, tuple[int, int]], delay: float) -> None:
        try:
            import pyautogui
            import pyperclip
        except ImportError:
            self.events.put(("error", "Install requirements first: pip install -r requirements.txt"))
            return
        try:
            time.sleep(3)
            total = len(self.filtered_rows)
            for number, (_, row) in enumerate(self.filtered_rows.iterrows(), start=1):
                if self.stop_requested.is_set():
                    self.events.put(("status", f"Stopped after {number - 1} of {total} card(s)."))
                    return
                for field, text in (("Front", row["Chinese"]), ("Back", row["English"]), ("Pinyin", row["Pinyin"])):
                    pyperclip.copy(str(text))
                    pyautogui.click(*positions[field])
                    pyautogui.hotkey("ctrl", "v")
                pyautogui.click(*positions["Add"])
                self.events.put(("status", f"Sent {number} of {total} card(s)."))
                time.sleep(delay)
            self.events.put(("complete", f"Finished sending {total} card(s)."))
        except Exception as error:
            self.events.put(("error", f"Automation stopped: {error}"))
        finally:
            self.events.put(("done", ""))

    def stop_automation(self) -> None:
        self.stop_requested.set()

    def _process_events(self) -> None:
        try:
            while True:
                event, message = self.events.get_nowait()
                self.status.set(message)
                if event == "error":
                    messagebox.showerror("Automation could not start", message)
                    self.start_button.configure(state="normal")
                elif event == "complete":
                    self.start_button.configure(state="normal")
                elif event == "done":
                    self.start_button.configure(state="normal")
        except queue.Empty:
            pass
        self.root.after(100, self._process_events)


if __name__ == "__main__":
    window = tk.Tk()
    ExcelToAnkiApp(window)
    window.mainloop()

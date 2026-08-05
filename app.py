"""Desktop UI for previewing Excel vocabulary before adding it to Anki."""

from __future__ import annotations

import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from ankiHelpers import add_three_sided_card, list_decks

REQUIRED_COLUMNS = ("Chinese", "Pinyin", "English", "Lesson", "Character")


class ExcelToAnkiApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Excel to Anki")
        self.root.minsize(850, 500)
        self.dataframe = None
        self.filtered_rows = None
        self.lesson_values: dict[str, object] = {}
        self.card_type_values: dict[str, object] = {}
        self.events: queue.Queue[tuple[str, str]] = queue.Queue()

        self.file_path = tk.StringVar()
        self.lesson = tk.StringVar()
        self.card_type = tk.StringVar()
        self.deck_name = tk.StringVar()
        self.status = tk.StringVar(value="Choose an Excel workbook to begin.")

        self._build_interface()
        self.refresh_decks(show_error=False)
        self.root.after(100, self._process_events)

    def _build_interface(self) -> None:
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        style = ttk.Style(self.root)
        style.configure("TNotebook.Tab", padding=(12, 5))

        notebook = ttk.Notebook(self.root)
        notebook.grid(column=0, row=0, sticky="nsew", padx=10, pady=10)

        cards_tab = ttk.Frame(notebook, padding=16)
        pdf_tab = ttk.Frame(notebook, padding=16)
        settings_tab = ttk.Frame(notebook, padding=16)
        notebook.add(cards_tab, text="Cards")
        notebook.add(pdf_tab, text="PDF & OCR")
        notebook.add(settings_tab, text="Settings")

        self._build_cards_tab(cards_tab)
        self._build_pdf_ocr_tab(pdf_tab)
        self._build_settings_tab(settings_tab)

    def _build_cards_tab(self, container: ttk.Frame) -> None:
        container.columnconfigure(1, weight=1)
        container.rowconfigure(5, weight=1)

        ttk.Label(container, text="Excel to Anki", font=("Segoe UI", 18, "bold")).grid(column=0, row=0, columnspan=3, sticky="w")
        ttk.Label(container, text="Preview vocabulary from an Excel file before adding selected cards to Anki.").grid(column=0, row=1, columnspan=3, sticky="w", pady=(2, 14))

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

        deck_selector = ttk.LabelFrame(container, text="Target Anki deck", padding=10)
        deck_selector.grid(column=0, row=4, columnspan=3, sticky="ew", pady=8)
        deck_selector.columnconfigure(1, weight=1)
        ttk.Label(deck_selector, text="Deck").grid(column=0, row=0, sticky="w")
        self.deck_box = ttk.Combobox(deck_selector, textvariable=self.deck_name, state="readonly")
        self.deck_box.grid(column=1, row=0, sticky="ew", padx=8)
        ttk.Button(deck_selector, text="Refresh decks", command=self.refresh_decks).grid(column=2, row=0)
        ttk.Label(deck_selector, text="Close Anki before refreshing decks or adding cards.").grid(column=0, row=1, columnspan=3, sticky="w", pady=(6, 0))

        preview = ttk.LabelFrame(container, text="Preview", padding=8)
        preview.grid(column=0, row=5, columnspan=3, sticky="nsew", pady=8)
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

        actions = ttk.Frame(container)
        actions.grid(column=0, row=6, columnspan=3, sticky="ew", pady=(2, 0))
        actions.columnconfigure(0, weight=1)
        self.add_button = ttk.Button(actions, text="Add previewed cards to Anki", command=self.add_previewed_cards)
        self.add_button.grid(column=1, row=0, sticky="e")

        ttk.Label(container, textvariable=self.status, foreground="#245a36").grid(column=0, row=7, columnspan=3, sticky="w", pady=(6, 0))

    def _build_pdf_ocr_tab(self, container: ttk.Frame) -> None:
        container.columnconfigure(0, weight=1)
        ttk.Label(container, text="PDF & OCR", font=("Segoe UI", 18, "bold")).grid(column=0, row=0, sticky="w")
        ttk.Label(
            container,
            text="Open bilingual PDFs, inspect pages, and recognize printed Chinese and English text.",
        ).grid(column=0, row=1, sticky="w", pady=(2, 16))

        recognition = ttk.LabelFrame(container, text="Recognition workflow", padding=12)
        recognition.grid(column=0, row=2, sticky="ew")
        ttk.Label(
            recognition,
            text="PDF viewing and page-by-page PaddleOCR recognition will be added here next. "
            "The original PDF will remain unchanged during recognition.",
            wraplength=720,
            justify="left",
        ).grid(column=0, row=0, sticky="w")

    def _build_settings_tab(self, container: ttk.Frame) -> None:
        container.columnconfigure(0, weight=1)
        ttk.Label(container, text="Settings", font=("Segoe UI", 18, "bold")).grid(column=0, row=0, sticky="w")
        ttk.Label(container, text="Application and integration settings.").grid(column=0, row=1, sticky="w", pady=(2, 16))

        anki_settings = ttk.LabelFrame(container, text="Anki integration", padding=12)
        anki_settings.grid(column=0, row=2, sticky="ew")
        anki_settings.columnconfigure(0, weight=1)
        ttk.Label(
            anki_settings,
            text="Refresh the available Anki decks after fully closing Anki and waiting for any media sync to finish.",
            wraplength=720,
            justify="left",
        ).grid(column=0, row=0, sticky="w")
        ttk.Button(anki_settings, text="Refresh Anki decks", command=self.refresh_decks).grid(column=1, row=0, sticky="e", padx=(12, 0))

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
        self.status.set(f"{len(rows)} matching card(s) ready to add.")

    def refresh_decks(self, show_error: bool = True) -> None:
        previous_deck = self.deck_name.get()
        try:
            decks = list_decks()
        except Exception as error:
            self.deck_box["values"] = []
            self.deck_name.set("")
            self.status.set("Unable to access your Anki collection. Close Anki, then refresh decks.")
            if show_error:
                messagebox.showerror("Could not load Anki decks", str(error))
            return

        self.deck_box["values"] = decks
        if previous_deck in decks:
            self.deck_name.set(previous_deck)
        elif decks:
            self.deck_name.set(decks[0])
        else:
            self.deck_name.set("")
        self.status.set(f"Loaded {len(decks)} Anki deck(s).")

    def add_previewed_cards(self) -> None:
        if self.filtered_rows is None or self.filtered_rows.empty:
            messagebox.showinfo("Prepare cards", "Preview matching cards before adding them to Anki.")
            return
        deck_name = self.deck_name.get()
        if not deck_name:
            messagebox.showinfo("Choose a deck", "Choose an Anki deck before adding cards.")
            return
        card_count = len(self.filtered_rows)
        confirmed = messagebox.askyesno(
            "Add cards to Anki",
            f"Add {card_count} card(s) to the {deck_name!r} deck?\n\nAnki must remain closed until this finishes.",
        )
        if not confirmed:
            return

        self.add_button.configure(state="disabled")
        rows = self.filtered_rows.copy()
        threading.Thread(target=self._add_cards_worker, args=(deck_name, rows), daemon=True).start()
        self.status.set(f"Adding {card_count} card(s) to {deck_name!r}...")

    def _add_cards_worker(self, deck_name: str, rows) -> None:
        try:
            total = len(rows)
            for number, (_, row) in enumerate(rows.iterrows(), start=1):
                add_three_sided_card(
                    deck_name=deck_name,
                    english=str(row["English"]),
                    chinese=str(row["Chinese"]),
                    pinyin=str(row["Pinyin"]),
                )
                self.events.put(("progress", f"Added {number} of {total} card(s) to {deck_name!r}."))
            self.events.put(("complete", f"Finished adding {total} card(s) to {deck_name!r}."))
        except Exception as error:
            self.events.put(("error", f"No more cards were added: {error}"))
        finally:
            self.events.put(("done", ""))

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

    def _process_events(self) -> None:
        try:
            while True:
                event, message = self.events.get_nowait()
                if message:
                    self.status.set(message)
                if event == "error":
                    messagebox.showerror("Could not add cards", message)
                elif event == "done":
                    self.add_button.configure(state="normal")
        except queue.Empty:
            pass
        self.root.after(100, self._process_events)

    # Legacy cursor-position automation is intentionally disabled now that the
    # app calls ankiHelpers.add_three_sided_card() directly.
    #
    # def start_automation(self) -> None:
    #     """Previously started the cursor-position card-entry workflow."""
    #
    # def _parse_position(self, value: str) -> tuple[int, int]:
    #     """Previously parsed x,y coordinates for the Anki Add window."""
    #
    # def _send_cards(self, positions: dict[str, tuple[int, int]], delay: float) -> None:
    #     """Previously pasted each card field with pyautogui and pyperclip."""
    #
    # def stop_automation(self) -> None:
    #     """Previously stopped the cursor-position card-entry workflow."""


if __name__ == "__main__":
    window = tk.Tk()
    ExcelToAnkiApp(window)
    window.mainloop()

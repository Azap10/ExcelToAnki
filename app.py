"""Desktop UI for previewing Excel vocabulary before adding it to Anki."""

from __future__ import annotations

import base64
from io import BytesIO
import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from ankiHelpers import add_three_sided_card, list_decks
from dictionaryHelpers import DictionaryEntry, definition_parts, lookup, pinyin_with_tone_marks
from ocrHelpers import RENDER_DPI, create_ocr_engine, open_pdf, process_pdf_region

REQUIRED_COLUMNS = ("Chinese", "Pinyin", "English", "Lesson", "Character")
PDF_RENDER_OVERSAMPLE = 2
PDF_VIEWER_PADDING = 16
PDF_VIEWER_MAX_WIDTH = 760
PDF_VIEWER_MAX_HEIGHT = 860
PDF_VIEWER_MIN_WIDTH = 360
PDF_VIEWER_MIN_HEIGHT = 480


class ExcelToAnkiApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Excel to Anki")
        self.root.minsize(850, 500)
        self.root.state("zoomed")
        self.dataframe = None
        self.filtered_rows = None
        self.pdf_document = None
        self.pdf_path: Path | None = None
        self.pdf_page_index = 0
        self.pdf_zoom = 1.0
        self.pdf_image = None
        self.pdf_display_bounds: tuple[int, int, int, int] | None = None
        self.pdf_render_job = None
        self.ocr_engine = None
        self.ocr_engine_error: Exception | None = None
        self.ocr_engine_ready = threading.Event()
        self.entry_selection_active = False
        self.selection_start: tuple[int, int] | None = None
        self.selection_rectangle = None
        self.pending_entry_text = ""
        self.lesson_values: dict[str, object] = {}
        self.card_type_values: dict[str, object] = {}
        self.dictionary_entries: tuple[DictionaryEntry, ...] = ()
        self.manual_card_number = 0
        self.events: queue.Queue[tuple[str, object]] = queue.Queue()

        self.file_path = tk.StringVar()
        self.lesson = tk.StringVar()
        self.card_type = tk.StringVar()
        self.deck_name = tk.StringVar()
        self.pdf_file_name = tk.StringVar(value="No PDF selected")
        self.pdf_page_label = tk.StringVar(value="Page - of -")
        self.ocr_summary = tk.StringVar(value="Open a PDF, then recognize the current page.")
        self.ocr_engine_status = tk.StringVar(value="Preparing OCR engine...")
        self.entry_summary = tk.StringVar(value="Select Create Entry, then draw a box around text on the page.")
        self.dictionary_query = tk.StringVar()
        self.word_list: list[str] = []
        self.word_check_vars: dict[str, tk.BooleanVar] = {}
        self.dictionary_lookup_in_progress = False
        self.dictionary_summary = tk.StringVar(value="Enter a Chinese word or character to search CC-CEDICT.")
        self.dictionary_character_set = tk.StringVar(value="simplified")
        self.draft_chinese = tk.StringVar()
        self.draft_pinyin = tk.StringVar()
        self.card_source = tk.StringVar(value="Excel workbook")
        self.preview_source = None
        self.definition_options: list[tuple[str, tk.BooleanVar]] = []
        self._definition_option_updates = False
        self.status = tk.StringVar(value="Choose an Excel workbook to begin.")

        self._build_interface()
        self.refresh_decks(show_error=False)
        self.root.after(100, self._process_events)
        self.root.protocol("WM_DELETE_WINDOW", self.close_application)

    def _build_interface(self) -> None:
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        style = ttk.Style(self.root)
        style.configure("TNotebook.Tab", padding=(12, 5))
        style.configure("LargeWord.TCheckbutton", font=("Segoe UI", 14))

        notebook = ttk.Notebook(self.root)
        notebook.grid(column=0, row=0, sticky="nsew", padx=10, pady=10)
        self.notebook = notebook

        cards_tab = ttk.Frame(notebook, padding=16)
        pdf_tab = ttk.Frame(notebook, padding=16)
        dictionary_tab = ttk.Frame(notebook, padding=16)
        settings_tab = ttk.Frame(notebook, padding=16)
        notebook.add(pdf_tab, text="Choose Words")
        notebook.add(dictionary_tab, text="Make Cards")
        notebook.add(cards_tab, text="Upload")
        notebook.add(settings_tab, text="Settings")

        self._build_cards_tab(cards_tab)
        self._build_pdf_ocr_tab(pdf_tab)
        self._build_dictionary_tab(dictionary_tab)
        self._build_settings_tab(settings_tab)
        self.root.bind_all("<KeyPress>", self._handle_make_cards_key)

    def _build_cards_tab(self, container: ttk.Frame) -> None:
        container.columnconfigure(0, weight=1)
        container.rowconfigure(3, weight=1)

        ttk.Label(container, text="Upload cards to Anki", font=("Segoe UI", 18, "bold")).grid(column=0, row=0, sticky="w")
        ttk.Label(container, text="Preview cards created in Make Cards, choose a deck, and upload them to Anki.").grid(column=0, row=1, sticky="w", pady=(2, 14))

        deck_selector = ttk.LabelFrame(container, text="Target Anki deck", padding=10)
        deck_selector.grid(column=0, row=2, sticky="ew", pady=8)
        deck_selector.columnconfigure(1, weight=1)
        ttk.Label(deck_selector, text="Deck").grid(column=0, row=0, sticky="w")
        self.deck_box = ttk.Combobox(deck_selector, textvariable=self.deck_name, state="readonly")
        self.deck_box.grid(column=1, row=0, sticky="ew", padx=8)
        ttk.Button(deck_selector, text="Refresh decks", command=self.refresh_decks).grid(column=2, row=0)
        ttk.Label(deck_selector, text="Close Anki before refreshing decks or adding cards.").grid(column=0, row=1, columnspan=3, sticky="w", pady=(6, 0))

        preview = ttk.LabelFrame(container, text="Preview", padding=8)
        preview.grid(column=0, row=3, sticky="nsew", pady=8)
        preview.columnconfigure(0, weight=1)
        preview.rowconfigure(0, weight=1)
        self.cards = ttk.Treeview(preview, columns=("#", "Chinese", "Pinyin", "English"), show="headings", height=10)
        self.cards.tag_configure("even", background="#ffffff")
        self.cards.tag_configure("odd", background="#eef4fb")
        for column, width in (("#", 55), ("Chinese", 200), ("Pinyin", 220), ("English", 340)):
            self.cards.heading(column, text=column)
            self.cards.column(column, width=width, anchor="w")
        scrollbar = ttk.Scrollbar(preview, orient="vertical", command=self.cards.yview)
        self.cards.configure(yscrollcommand=scrollbar.set)
        self.cards.grid(column=0, row=0, sticky="nsew")
        scrollbar.grid(column=1, row=0, sticky="ns")

        actions = ttk.Frame(container)
        actions.grid(column=0, row=4, sticky="ew", pady=(2, 0))
        actions.columnconfigure(0, weight=1)
        ttk.Button(actions, text="Refresh preview", command=self.prepare_cards).grid(column=1, row=0, padx=(0, 8))
        self.add_button = ttk.Button(actions, text="Add previewed cards to Anki", command=self.add_previewed_cards)
        self.add_button.grid(column=2, row=0, sticky="e")

        ttk.Label(container, textvariable=self.status, foreground="#245a36").grid(column=0, row=5, sticky="w", pady=(6, 0))

    def _build_dictionary_tab(self, container: ttk.Frame) -> None:
        """Build the manual dictionary-lookup and card-queue workflow."""
        container.columnconfigure(0, weight=1)
        container.rowconfigure(2, weight=0)
        container.rowconfigure(3, weight=1)
        container.rowconfigure(4, weight=1)
        container.rowconfigure(5, weight=1)

        ttk.Label(container, text="CC-CEDICT Dictionary", font=("Segoe UI", 18, "bold")).grid(column=0, row=0, sticky="w")
        ttk.Label(
            container,
            text="Choose a word, review its dictionary matches, and build a flashcard.",
        ).grid(column=0, row=1, sticky="w", pady=(2, 14))

        results = ttk.Frame(container)
        results.grid(column=0, row=2, sticky="ew", pady=(14, 0))
        results.columnconfigure(0, weight=1)
        results.columnconfigure(1, weight=2)
        results.columnconfigure(2, weight=2)
        results.rowconfigure(0, weight=1)

        word_list_frame = ttk.LabelFrame(results, text="Chosen words", padding=8)
        word_list_frame.grid(column=0, row=0, sticky="nsew", padx=(0, 8))
        word_list_frame.columnconfigure(0, weight=1)
        word_list_frame.rowconfigure(0, weight=1)
        self.chosen_words = tk.Listbox(word_list_frame, exportselection=False, height=4)
        self.chosen_words.grid(column=0, row=0, sticky="nsew")
        self.chosen_words.bind("<<ListboxSelect>>", self._on_chosen_word_selected)
        chosen_scrollbar = ttk.Scrollbar(word_list_frame, orient="vertical", command=self.chosen_words.yview)
        chosen_scrollbar.grid(column=1, row=0, sticky="ns")
        self.chosen_words.configure(yscrollcommand=chosen_scrollbar.set)

        matches_frame = ttk.LabelFrame(results, text="Matches", padding=8)
        matches_frame.grid(column=1, row=0, sticky="nsew", padx=(0, 8))
        matches_frame.columnconfigure(0, weight=1)
        matches_frame.rowconfigure(0, weight=1)
        self.dictionary_matches = tk.Listbox(matches_frame, exportselection=False, height=4)
        self.dictionary_matches.grid(column=0, row=0, sticky="nsew")
        self.dictionary_matches.bind("<<ListboxSelect>>", self._on_dictionary_match_selected)
        matches_scrollbar = ttk.Scrollbar(matches_frame, orient="vertical", command=self.dictionary_matches.yview)
        matches_scrollbar.grid(column=1, row=0, sticky="ns")
        self.dictionary_matches.configure(yscrollcommand=matches_scrollbar.set)

        identity = ttk.LabelFrame(results, text="Flashcard Chinese / Pinyin", padding=8)
        identity.grid(column=2, row=0, sticky="ew")
        identity.columnconfigure(1, weight=1)
        ttk.Label(identity, text="Chinese").grid(column=0, row=0, sticky="w")
        ttk.Entry(identity, textvariable=self.draft_chinese).grid(column=1, row=0, sticky="ew", padx=8)
        ttk.Label(identity, text="Pinyin").grid(column=0, row=1, sticky="w", pady=(8, 0))
        ttk.Entry(identity, textvariable=self.draft_pinyin).grid(column=1, row=1, sticky="ew", padx=8, pady=(8, 0))

        middle = ttk.Frame(container)
        middle.grid(column=0, row=3, sticky="nsew", pady=(14, 0))
        middle.columnconfigure(0, weight=1)
        middle.columnconfigure(1, weight=1)
        middle.rowconfigure(0, weight=1)

        definition_frame = ttk.LabelFrame(middle, text="Definition selection", padding=8)
        definition_frame.grid(column=0, row=0, sticky="nsew", padx=(0, 8))
        definition_frame.columnconfigure(0, weight=1)
        definition_frame.rowconfigure(0, weight=1)
        self.dictionary_definition = ttk.Frame(definition_frame)
        self.dictionary_definition.grid(column=0, row=0, sticky="nsew")

        meaning_frame = ttk.LabelFrame(middle, text="Meaning", padding=8)
        meaning_frame.grid(column=1, row=0, sticky="nsew")
        meaning_frame.columnconfigure(0, weight=1)
        meaning_frame.rowconfigure(0, weight=1)
        self.draft_meaning = tk.Text(meaning_frame, height=8, wrap="word", font=("Segoe UI", 10))
        self.draft_meaning.grid(column=0, row=0, sticky="nsew")
        tag_actions = ttk.Frame(meaning_frame)
        tag_actions.grid(column=0, row=1, sticky="w", pady=(8, 0))
        for tag in ("(adj.)", "(v.)", "(n.)", "(adv.)", "(m.)", "(pron.)", "(idiom)"):
            ttk.Button(tag_actions, text=tag, command=lambda value=tag: self.add_definition_tag(value)).pack(side="left", padx=(0, 5))
        ttk.Button(meaning_frame, text="Add draft to queue", command=self.add_draft_to_queue).grid(column=0, row=2, sticky="e", pady=(8, 0))

        queue_frame = ttk.LabelFrame(container, text="Dictionary card queue", padding=8)
        queue_frame.grid(column=0, row=4, sticky="nsew", pady=(14, 0))
        queue_frame.columnconfigure(0, weight=1)
        queue_frame.rowconfigure(0, weight=1)
        self.dictionary_card_queue = ttk.Treeview(queue_frame, columns=("#", "Chinese", "Pinyin", "English"), show="headings", height=6)
        self.dictionary_card_queue.tag_configure("even", background="#ffffff")
        self.dictionary_card_queue.tag_configure("odd", background="#eef4fb")
        for column, width in (("#", 55), ("Chinese", 180), ("Pinyin", 200), ("English", 480)):
            self.dictionary_card_queue.heading(column, text=column)
            self.dictionary_card_queue.column(column, width=width, anchor="w")
        self.dictionary_card_queue.grid(column=0, row=0, sticky="nsew")
        queue_scrollbar = ttk.Scrollbar(queue_frame, orient="vertical", command=self.dictionary_card_queue.yview)
        queue_scrollbar.grid(column=1, row=0, sticky="ns")
        self.dictionary_card_queue.configure(yscrollcommand=queue_scrollbar.set)
        queue_actions = ttk.Frame(queue_frame)
        queue_actions.grid(column=0, row=1, columnspan=2, sticky="ew", pady=(8, 0))
        queue_actions.columnconfigure(0, weight=1)
        ttk.Button(queue_actions, text="Remove selected", command=self.remove_selected_dictionary_cards).grid(column=1, row=0)
        ttk.Label(queue_actions, text="Use the Upload tab to preview and add this queue to Anki.").grid(column=2, row=0, padx=(12, 0))

    def add_word_to_list(self, _event=None) -> str:
        word = self.word_entry.get().strip()
        if not word:
            return "break"
        if word not in self.word_list:
            self.word_list.append(word)
            self._refresh_word_lists()
        self.word_entry.delete(0, tk.END)
        self.chosen_words.selection_clear(0, tk.END)
        self.chosen_words.selection_set(self.word_list.index(word))
        self.chosen_words.see(self.word_list.index(word))
        self.dictionary_query.set(word)
        self.search_dictionary()
        return "break"

    def _refresh_word_lists(self) -> None:
        """Render the shared word list with stable one-based numbering."""
        for child in self.word_choices_inner.winfo_children():
            child.destroy()
        self.word_check_vars = {}
        self.chosen_words.delete(0, tk.END)
        for position, word in enumerate(self.word_list, start=1):
            label = f"{position}. {word}"
            selected = tk.BooleanVar(value=False)
            self.word_check_vars[word] = selected
            ttk.Checkbutton(
                self.word_choices_inner,
                text=label,
                variable=selected,
                style="LargeWord.TCheckbutton",
            ).grid(column=0, row=position - 1, sticky="w", pady=(0 if position == 1 else 4, 0))
            self.chosen_words.insert(tk.END, label)

    def remove_checked_words(self) -> None:
        checked = {word for word, variable in self.word_check_vars.items() if variable.get()}
        if checked:
            self.word_list = [word for word in self.word_list if word not in checked]
            self._refresh_word_lists()

    def remove_all_words(self) -> None:
        if self.word_list:
            self.word_list.clear()
            self._refresh_word_lists()

    def _on_chosen_word_selected(self, _event=None) -> None:
        selected = self.chosen_words.curselection()
        if not selected or selected[0] >= len(self.word_list):
            return
        self.dictionary_query.set(self.word_list[selected[0]])
        self.search_dictionary()

    def search_dictionary(self, _event=None) -> None:
        query = self.dictionary_query.get().strip()
        if not query:
            messagebox.showinfo("Enter a word", "Enter a Chinese word or character to look it up.")
            return
        if self.dictionary_lookup_in_progress:
            return
        self.dictionary_lookup_in_progress = True
        self.dictionary_summary.set(f"Looking up {query}…")
        threading.Thread(target=self._dictionary_lookup_worker, args=(query,), daemon=True).start()

    def _dictionary_lookup_worker(self, query: str) -> None:
        try:
            self.events.put(("dictionary_result", (query, lookup(query))))
        except Exception as error:
            self.events.put(("dictionary_error", str(error)))

    def _display_dictionary_results(self, query: str, entries: tuple[DictionaryEntry, ...]) -> None:
        self.dictionary_entries = entries
        self.dictionary_matches.delete(0, tk.END)
        self._set_definition_options(())
        if not entries:
            self.dictionary_summary.set(f"No CC-CEDICT entries found for {query!r}.")
            return
        self._refresh_dictionary_match_labels()
        self.dictionary_matches.selection_set(0)
        self._show_selected_dictionary_entry()
        self.dictionary_summary.set(f"Found {len(entries)} CC-CEDICT entr{'y' if len(entries) == 1 else 'ies'} for {query!r}.")

    def _refresh_dictionary_match_labels(self) -> None:
        selected = self.dictionary_matches.curselection()
        self.dictionary_matches.delete(0, tk.END)
        for entry in self.dictionary_entries:
            headword = self._dictionary_headword(entry)
            preview = entry.definitions[0] if entry.definitions else "No definition"
            self.dictionary_matches.insert(tk.END, f"{headword} [{entry.pinyin}] — {preview}")
        if selected and selected[0] < len(self.dictionary_entries):
            self.dictionary_matches.selection_set(selected[0])

    def _dictionary_headword(self, entry: DictionaryEntry) -> str:
        return entry.traditional if self.dictionary_character_set.get() == "traditional" else entry.simplified

    def _on_dictionary_match_selected(self, _event=None) -> None:
        self._show_selected_dictionary_entry()

    def _show_selected_dictionary_entry(self) -> None:
        selected = self.dictionary_matches.curselection()
        if not selected or selected[0] >= len(self.dictionary_entries):
            return
        entry = self.dictionary_entries[selected[0]]
        self._set_definition_options(definition_parts(entry))
        self.draft_chinese.set(self._dictionary_headword(entry))
        self.draft_pinyin.set(pinyin_with_tone_marks(entry.pinyin))

    def _set_definition_options(self, definitions: tuple[str, ...]) -> None:
        """Rebuild definition checkboxes and clear the meaning draft."""
        self._definition_option_updates = True
        try:
            for child in self.dictionary_definition.winfo_children():
                child.destroy()
            self.definition_options = []
            for row, definition in enumerate(definitions):
                selected = tk.BooleanVar(value=False)
                selected.set(False)
                shortcut = ("1234567890"[row] + ". ") if row < 10 else "    "
                ttk.Checkbutton(
                    self.dictionary_definition,
                    text=f"{shortcut}{definition}",
                    variable=selected,
                    command=lambda meaning=definition, value=selected: self._definition_option_toggled(meaning, value),
                ).grid(column=0, row=row, sticky="w", pady=(0 if row == 0 else 4, 0))
                self.definition_options.append((definition, selected))
        finally:
            self._definition_option_updates = False
        self.draft_meaning.delete("1.0", tk.END)

    def _definition_option_toggled(self, meaning: str, selected: tk.BooleanVar) -> None:
        """Add/remove one meaning while preserving unrelated draft edits."""
        if self._definition_option_updates:
            return
        current = self.draft_meaning.get("1.0", "end-1c").strip()
        if selected.get():
            if current == meaning or current.endswith(f"; {meaning}") or current.endswith(f" {meaning}"):
                updated = current
            elif current and current.endswith(("(adj.)", "(v.)", "(n.)", "(adv.)", "(m.)", "(pron.)", "(idiom)")):
                updated = f"{current} {meaning}"
            elif current:
                updated = f"{current}; {meaning}"
            else:
                updated = meaning
        else:
            updated = current
            for delimiter in (f"; {meaning}", f" {meaning}"):
                if updated.endswith(delimiter):
                    updated = updated[: -len(delimiter)].rstrip()
                    break
            if updated == meaning:
                updated = ""
        self.draft_meaning.delete("1.0", tk.END)
        self.draft_meaning.insert("1.0", updated)

    def add_selected_definition_to_draft(self, _event=None) -> str | None:
        selected_entry = self.dictionary_matches.curselection()
        if not selected_entry or selected_entry[0] >= len(self.dictionary_entries):
            messagebox.showinfo("Choose an entry", "Choose a dictionary entry before adding a definition.")
            return "break"
        entry = self.dictionary_entries[selected_entry[0]]
        self.draft_chinese.set(self._dictionary_headword(entry))
        self.draft_pinyin.set(pinyin_with_tone_marks(entry.pinyin))
        for meaning, variable in self.definition_options:
            variable.set(True)
            self._definition_option_toggled(meaning, variable)
        self.status.set("Added selected dictionary text to the flashcard draft.")
        return "break"

    def add_draft_to_queue(self) -> None:
        chinese = self.draft_chinese.get().strip()
        pinyin = self.draft_pinyin.get().strip()
        meaning = self.draft_meaning.get("1.0", "end-1c").strip()
        if not all((chinese, pinyin, meaning)):
            messagebox.showinfo("Complete the draft", "Chinese, Pinyin, and Meaning are all required.")
            return
        self.manual_card_number += 1
        position = len(self.dictionary_card_queue.get_children()) + 1
        self.dictionary_card_queue.insert(
            "", "end", iid=f"dictionary-{self.manual_card_number}",
            values=(position, chinese, pinyin, meaning),
            tags=("odd" if position % 2 else "even",),
        )
        self.draft_meaning.delete("1.0", tk.END)
        self.status.set("Added dictionary card to the queue.")

    def add_definition_tag(self, tag: str) -> None:
        """Append a part-of-speech tag to the meaning without adding punctuation."""
        current = self.draft_meaning.get("1.0", "end-1c").strip()
        if tag in current.split():
            return
        updated = f"{current} {tag}" if current else tag
        self.draft_meaning.delete("1.0", tk.END)
        self.draft_meaning.insert("1.0", updated)

    def _handle_make_cards_key(self, event) -> str | None:
        """Handle tab navigation and Make Cards keyboard shortcuts."""
        if event.keysym in {"Left", "Right"}:
            tabs = self.notebook.tabs()
            current_tab = self.notebook.select()
            if current_tab in tabs:
                current_index = tabs.index(current_tab)
                step = -1 if event.keysym == "Left" else 1
                next_index = max(0, min(len(tabs) - 1, current_index + step))
                self.notebook.select(tabs[next_index])
                return "break"

        if self.notebook.tab(self.notebook.select(), "text") != "Make Cards":
            return None
        focused = self.root.focus_get()
        if focused is not None and focused.winfo_class() in {"Entry", "TEntry", "Text"} and event.keysym not in {"Return", "Up", "Down"}:
            return None

        number_keys = "1234567890"
        if event.char in number_keys:
            index = number_keys.index(event.char)
            if index < len(self.definition_options):
                meaning, variable = self.definition_options[index]
                variable.set(not variable.get())
                self._definition_option_toggled(meaning, variable)
                return "break"

        tag_keys = {
            "n": "(n.)", "v": "(v.)", "m": "(m.)", "a": "(adj.)",
            "d": "(adv.)", "p": "(pron.)", "i": "(idiom)",
        }
        if event.char.lower() in tag_keys and not (focused and focused.winfo_class() in {"Entry", "TEntry", "Text"}):
            self.add_definition_tag(tag_keys[event.char.lower()])
            return "break"

        if event.keysym in {"Up", "Down"}:
            count = len(self.dictionary_entries)
            if count:
                selected = self.dictionary_matches.curselection()
                if selected:
                    current = selected[0]
                else:
                    current = 0 if event.keysym == "Up" else -1
                next_index = max(0, current - 1) if event.keysym == "Up" else min(count - 1, current + 1)
                self.dictionary_matches.selection_clear(0, tk.END)
                self.dictionary_matches.selection_set(next_index)
                self.dictionary_matches.see(next_index)
                self._show_selected_dictionary_entry()
                return "break"

        if event.keysym == "Return" and not (focused and focused.winfo_class() == "Entry"):
            before = len(self.dictionary_card_queue.get_children())
            self.add_draft_to_queue()
            if len(self.dictionary_card_queue.get_children()) > before:
                self._advance_to_next_word()
            return "break"
        return None

    def _advance_to_next_word(self) -> None:
        selected = self.chosen_words.curselection()
        current = selected[0] if selected else -1
        if current < 0 or current >= len(self.word_list) - 1:
            return
        next_index = current + 1
        self.chosen_words.selection_clear(0, tk.END)
        self.chosen_words.selection_set(next_index)
        self.chosen_words.see(next_index)
        self.dictionary_query.set(self.word_list[next_index])
        self.search_dictionary()

    def remove_selected_dictionary_cards(self) -> None:
        selected = self.dictionary_card_queue.selection()
        for item in selected:
            self.dictionary_card_queue.delete(item)
        if selected:
            self._renumber_treeview(self.dictionary_card_queue)
            self.status.set(f"Removed {len(selected)} dictionary card(s) from the queue.")

    def add_dictionary_cards_to_anki(self) -> None:
        items = self.dictionary_card_queue.get_children()
        if not items:
            messagebox.showinfo("Queue is empty", "Add a dictionary card to the queue before sending cards to Anki.")
            return
        deck_name = self.deck_name.get()
        if not deck_name:
            messagebox.showinfo("Choose a deck", "Choose a target Anki deck on the Cards tab first.")
            return
        if not messagebox.askyesno("Add cards to Anki", f"Add {len(items)} dictionary card(s) to {deck_name!r}?"):
            return
        cards = [self.dictionary_card_queue.item(item, "values")[1:4] for item in items]
        self.add_button.configure(state="disabled")
        threading.Thread(target=self._add_dictionary_cards_worker, args=(deck_name, cards), daemon=True).start()

    def _add_dictionary_cards_worker(self, deck_name: str, cards: list[tuple[str, str, str]]) -> None:
        try:
            for number, (chinese, pinyin, english) in enumerate(cards, start=1):
                add_three_sided_card(deck_name=deck_name, english=english, chinese=chinese, pinyin=pinyin)
                self.events.put(("dictionary_progress", f"Added {number} of {len(cards)} dictionary card(s) to {deck_name!r}."))
            self.events.put(("dictionary_complete", f"Finished adding {len(cards)} dictionary card(s) to {deck_name!r}."))
        except Exception as error:
            self.events.put(("dictionary_add_error", f"No more dictionary cards were added: {error}"))
        finally:
            self.events.put(("dictionary_done", ""))

    def _build_pdf_ocr_tab(self, container: ttk.Frame) -> None:
        """Build the Choose Words workspace with a PDF viewer and word list."""
        container.columnconfigure(0, weight=3)
        container.columnconfigure(1, weight=1)
        container.rowconfigure(0, weight=1)
        viewer_frame = ttk.Frame(container)
        viewer_frame.grid(column=0, row=0, sticky="nsew", padx=(0, 16))
        viewer_frame.columnconfigure(0, weight=1)
        viewer_frame.rowconfigure(0, weight=1)
        self.pdf_viewer = ttk.Frame(viewer_frame, width=620, height=820)
        self.pdf_viewer.grid(column=0, row=0, sticky="nsew")
        self.pdf_viewer.grid_propagate(False)
        self.pdf_viewer.columnconfigure(0, weight=1)
        self.pdf_viewer.rowconfigure(0, weight=1)
        self.pdf_canvas = tk.Canvas(self.pdf_viewer, background="#2d2d2d", highlightthickness=0)
        self.pdf_canvas.grid(column=0, row=0, sticky="nsew")
        pdf_vertical_scrollbar = ttk.Scrollbar(self.pdf_viewer, orient="vertical", command=self.pdf_canvas.yview)
        pdf_vertical_scrollbar.grid(column=1, row=0, sticky="ns")
        pdf_horizontal_scrollbar = ttk.Scrollbar(self.pdf_viewer, orient="horizontal", command=self.pdf_canvas.xview)
        pdf_horizontal_scrollbar.grid(column=0, row=1, sticky="ew")
        self.pdf_canvas.configure(xscrollcommand=pdf_horizontal_scrollbar.set, yscrollcommand=pdf_vertical_scrollbar.set)
        self.pdf_canvas.create_text(300, 300, text="Open a PDF to preview it here", fill="#d7d7d7", font=("Segoe UI", 14))
        self.pdf_canvas.bind("<Configure>", self._queue_pdf_render)
        self.pdf_canvas.bind("<ButtonPress-1>", self._begin_pdf_text_selection)
        self.pdf_canvas.bind("<B1-Motion>", self._draw_pdf_text_selection)
        self.pdf_canvas.bind("<ButtonRelease-1>", self._finish_pdf_text_selection)

        controls = ttk.Frame(container)
        controls.grid(column=1, row=0, sticky="nsew")
        controls.columnconfigure(0, weight=1)
        ttk.Label(controls, text="Choose words", font=("Segoe UI", 18, "bold")).grid(column=0, row=0, sticky="w")
        ttk.Label(controls, text="Open a PDF, then drag over selectable text to add it to the word list.", wraplength=320).grid(column=0, row=1, sticky="w", pady=(4, 14))
        pdf_controls = ttk.LabelFrame(controls, text="PDF", padding=10)
        pdf_controls.grid(column=0, row=2, sticky="ew")
        pdf_controls.columnconfigure(0, weight=1)
        ttk.Button(pdf_controls, text="Open PDF", command=self.choose_pdf).grid(column=0, row=0, sticky="ew")
        ttk.Label(pdf_controls, textvariable=self.pdf_file_name, wraplength=300).grid(column=0, row=1, sticky="w", pady=(8, 0))
        ttk.Button(pdf_controls, text="Previous page", command=lambda: self.change_pdf_page(-1)).grid(column=0, row=2, sticky="w", pady=(8, 0))
        ttk.Button(pdf_controls, text="Next page", command=lambda: self.change_pdf_page(1)).grid(column=0, row=3, sticky="w", pady=(4, 0))
        ttk.Label(pdf_controls, textvariable=self.pdf_page_label).grid(column=0, row=4, sticky="w", pady=(6, 0))
        zoom_actions = ttk.Frame(pdf_controls)
        zoom_actions.grid(column=0, row=5, sticky="w", pady=(8, 0))
        ttk.Button(zoom_actions, text="Zoom out", command=lambda: self.adjust_pdf_zoom(-0.25)).pack(side="left")
        ttk.Button(zoom_actions, text="Reset", command=lambda: self.set_pdf_zoom(1.0)).pack(side="left", padx=5)
        ttk.Button(zoom_actions, text="Zoom in", command=lambda: self.adjust_pdf_zoom(0.25)).pack(side="left")
        ttk.Label(pdf_controls, textvariable=self.ocr_engine_status, wraplength=300, foreground="#555555").grid(column=0, row=6, sticky="w", pady=(8, 0))

        word_input = ttk.LabelFrame(controls, text="Add a word", padding=10)
        word_input.grid(column=0, row=3, sticky="ew", pady=(14, 0))
        word_input.columnconfigure(0, weight=1)
        self.word_entry = ttk.Entry(word_input)
        self.word_entry.grid(column=0, row=0, sticky="ew")
        self.word_entry.bind("<Return>", self.add_word_to_list)
        ttk.Button(word_input, text="Add word", command=self.add_word_to_list).grid(column=1, row=0, padx=(8, 0))
        self.word_choices_frame = ttk.LabelFrame(controls, text="Word list", padding=8)
        self.word_choices_frame.grid(column=0, row=4, sticky="nsew", pady=(14, 0))
        controls.rowconfigure(4, weight=1)
        self.word_choices_frame.grid_propagate(False)
        self.word_choices_frame.configure(height=280)
        self.word_choices_frame.columnconfigure(0, weight=1)
        self.word_choices_frame.rowconfigure(0, weight=1)
        self.word_choices_canvas = tk.Canvas(self.word_choices_frame, height=250, highlightthickness=0)
        self.word_choices_canvas.grid(column=0, row=0, sticky="nsew")
        word_choices_scrollbar = ttk.Scrollbar(self.word_choices_frame, orient="vertical", command=self.word_choices_canvas.yview)
        word_choices_scrollbar.grid(column=1, row=0, sticky="ns")
        self.word_choices_canvas.configure(yscrollcommand=word_choices_scrollbar.set)
        self.word_choices_inner = ttk.Frame(self.word_choices_canvas)
        self.word_choices_window = self.word_choices_canvas.create_window((0, 0), window=self.word_choices_inner, anchor="nw")
        self.word_choices_inner.bind("<Configure>", lambda _event: self.word_choices_canvas.configure(scrollregion=self.word_choices_canvas.bbox("all")))
        self.word_choices_canvas.bind("<Configure>", lambda event: self.word_choices_canvas.itemconfigure(self.word_choices_window, width=event.width))
        word_actions = ttk.Frame(controls)
        word_actions.grid(column=0, row=5, sticky="ew", pady=(8, 0))
        ttk.Button(word_actions, text="Remove checked words", command=self.remove_checked_words).pack(side="left")
        ttk.Button(word_actions, text="Remove all words", command=self.remove_all_words).pack(side="left", padx=(8, 0))
        threading.Thread(target=self._warm_up_ocr_engine, daemon=True).start()
        return

        # Legacy OCR controls remain below for later reactivation.
        ttk.Label(container, text="Choose words", font=("Segoe UI", 18, "bold")).grid(column=0, row=0, sticky="w")
        ttk.Label(
            container,
            text="Add Chinese words manually now. Future OCR selection will add words to this same list.",
        ).grid(column=0, row=1, sticky="w", pady=(2, 14))
        word_input = ttk.LabelFrame(container, text="Add a word", padding=10)
        word_input.grid(column=0, row=2, sticky="new")
        word_input.columnconfigure(0, weight=1)
        self.word_entry = ttk.Entry(word_input)
        self.word_entry.grid(column=0, row=0, sticky="ew")
        self.word_entry.bind("<Return>", self.add_word_to_list)
        ttk.Button(word_input, text="Add word", command=self.add_word_to_list).grid(column=1, row=0, padx=(8, 0))
        self.word_choices_frame = ttk.LabelFrame(container, text="Word list", padding=8)
        self.word_choices_frame.grid(column=0, row=3, sticky="nsew", pady=(14, 0))
        self.word_choices_frame.grid_propagate(False)
        self.word_choices_frame.configure(height=280)
        self.word_choices_frame.columnconfigure(0, weight=1)
        self.word_choices_frame.rowconfigure(0, weight=1)
        self.word_choices_canvas = tk.Canvas(self.word_choices_frame, height=250, highlightthickness=0)
        self.word_choices_canvas.grid(column=0, row=0, sticky="nsew")
        word_choices_scrollbar = ttk.Scrollbar(self.word_choices_frame, orient="vertical", command=self.word_choices_canvas.yview)
        word_choices_scrollbar.grid(column=1, row=0, sticky="ns")
        self.word_choices_canvas.configure(yscrollcommand=word_choices_scrollbar.set)
        self.word_choices_inner = ttk.Frame(self.word_choices_canvas)
        self.word_choices_window = self.word_choices_canvas.create_window((0, 0), window=self.word_choices_inner, anchor="nw")
        self.word_choices_inner.bind("<Configure>", lambda _event: self.word_choices_canvas.configure(scrollregion=self.word_choices_canvas.bbox("all")))
        self.word_choices_canvas.bind("<Configure>", lambda event: self.word_choices_canvas.itemconfigure(self.word_choices_window, width=event.width))
        word_actions = ttk.Frame(container)
        word_actions.grid(column=0, row=4, sticky="ew", pady=(8, 0))
        ttk.Button(word_actions, text="Remove checked words", command=self.remove_checked_words).pack(side="left")
        ttk.Button(word_actions, text="Remove all words", command=self.remove_all_words).pack(side="left", padx=(8, 0))
        return

        # The previous PDF viewer and OCR controls remain below for reference
        # while this branch is used for faster builds.
        container.columnconfigure(1, weight=1)
        container.rowconfigure(0, weight=1)

        self.pdf_viewer = ttk.Frame(
            container,
            width=PDF_VIEWER_MIN_WIDTH + PDF_VIEWER_PADDING * 2,
            height=PDF_VIEWER_MIN_HEIGHT + PDF_VIEWER_PADDING * 2,
        )
        self.pdf_viewer.grid(column=0, row=0, sticky="nw")
        self.pdf_viewer.grid_propagate(False)
        viewer = self.pdf_viewer
        viewer.columnconfigure(0, weight=1)
        viewer.rowconfigure(0, weight=1)
        self.pdf_canvas = tk.Canvas(viewer, background="#2d2d2d", highlightthickness=0)
        self.pdf_canvas.grid(column=0, row=0, sticky="nsew")
        self.pdf_canvas.create_text(
            300,
            250,
            text="Open a PDF to preview it here",
            fill="#d7d7d7",
            font=("Segoe UI", 14),
        )
        self.pdf_canvas.bind("<Configure>", self._queue_pdf_render)
        self.pdf_canvas.bind("<ButtonPress-1>", self._begin_pdf_selection)
        self.pdf_canvas.bind("<B1-Motion>", self._draw_pdf_selection)
        self.pdf_canvas.bind("<ButtonRelease-1>", self._finish_pdf_selection)

        processing = ttk.Frame(container)
        processing.grid(column=1, row=0, sticky="nsew", padx=(18, 0))
        processing.columnconfigure(0, weight=1)
        processing.rowconfigure(2, weight=1)

        controls = ttk.LabelFrame(processing, text="PDF controls", padding=12)
        controls.grid(column=0, row=0, sticky="new")
        controls.columnconfigure(0, weight=1)
        ttk.Button(controls, text="Open PDF", command=self.choose_pdf).grid(column=0, row=0, sticky="ew")
        ttk.Label(controls, textvariable=self.pdf_file_name, wraplength=520, justify="left").grid(column=0, row=1, sticky="w", pady=(10, 16))

        navigation = ttk.LabelFrame(controls, text="Page", padding=8)
        navigation.grid(column=0, row=2, sticky="ew")
        navigation.columnconfigure(1, weight=1)
        self.previous_pdf_button = ttk.Button(navigation, text="Previous", command=lambda: self.change_pdf_page(-1), state="disabled")
        self.previous_pdf_button.grid(column=0, row=0, sticky="w")
        self.next_pdf_button = ttk.Button(navigation, text="Next", command=lambda: self.change_pdf_page(1), state="disabled")
        self.next_pdf_button.grid(column=2, row=0, sticky="e")
        ttk.Label(navigation, textvariable=self.pdf_page_label).grid(column=0, row=1, columnspan=3, pady=(8, 0))

        entry = ttk.LabelFrame(controls, text="Create Entry", padding=8)
        entry.grid(column=0, row=3, sticky="ew", pady=(14, 0))
        entry.columnconfigure(0, weight=1)
        self.create_entry_button = ttk.Button(
            entry,
            text="Create Entry",
            command=self.start_entry_selection,
            state="disabled",
        )
        self.create_entry_button.grid(column=0, row=0, sticky="w")
        ttk.Label(entry, textvariable=self.entry_summary, wraplength=500, justify="left").grid(
            column=0, row=1, sticky="w", pady=(8, 0)
        )
        entry_actions = ttk.Frame(entry)
        entry_actions.grid(column=0, row=2, sticky="w", pady=(8, 0))
        self.confirm_entry_button = ttk.Button(
            entry_actions,
            text="Confirm entry",
            command=self.confirm_pending_entry,
            state="disabled",
        )
        self.confirm_entry_button.grid(column=0, row=0)
        self.retry_entry_button = ttk.Button(
            entry_actions,
            text="Retry selection",
            command=self.start_entry_selection,
            state="disabled",
        )
        self.retry_entry_button.grid(column=1, row=0, padx=(8, 0))

        ocr = ttk.LabelFrame(processing, text="OCR", padding=12)
        ocr.grid(column=0, row=1, sticky="ew", pady=(14, 0))
        ocr.columnconfigure(0, weight=1)
        self.recognize_page_button = ttk.Button(
            ocr,
            text="Recognize current page",
            command=self.recognize_current_pdf_page,
            state="disabled",
        )
        self.recognize_page_button.grid(column=0, row=0, sticky="w")
        ttk.Label(
            ocr,
            textvariable=self.ocr_summary,
            wraplength=520,
            justify="left",
        ).grid(column=0, row=1, sticky="w", pady=(8, 0))

        created_entries = ttk.LabelFrame(processing, text="Created entries", padding=8)
        created_entries.grid(column=0, row=2, sticky="nsew", pady=(14, 0))
        created_entries.columnconfigure(0, weight=1)
        created_entries.rowconfigure(0, weight=1)
        self.created_entries = ttk.Treeview(
            created_entries,
            columns=("#", "Page", "Text"),
            show="headings",
            height=16,
        )
        self.created_entries.tag_configure("even", background="#ffffff")
        self.created_entries.tag_configure("odd", background="#eef4fb")
        for column, width in (("#", 55), ("Page", 70), ("Text", 560)):
            self.created_entries.heading(column, text=column)
            self.created_entries.column(column, width=width, anchor="w")
        entries_scrollbar = ttk.Scrollbar(created_entries, orient="vertical", command=self.created_entries.yview)
        self.created_entries.configure(yscrollcommand=entries_scrollbar.set)
        self.created_entries.grid(column=0, row=0, sticky="nsew")
        entries_scrollbar.grid(column=1, row=0, sticky="ns")

    def choose_pdf(self) -> None:
        selected = filedialog.askopenfilename(title="Open PDF", filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")])
        if not selected:
            return
        try:
            import pymupdf

            document = pymupdf.open(selected)
        except Exception as error:
            messagebox.showerror("Could not open PDF", str(error))
            return

        self._close_pdf_document()
        self.pdf_document = document
        self.pdf_path = Path(selected)
        self.pdf_page_index = 0
        self.pdf_file_name.set(Path(selected).name)
        self.ocr_summary.set("Ready to recognize the current page.")
        self._show_pdf_page()
        self.status.set(f"Opened {Path(selected).name}: {document.page_count} page(s).")

    def change_pdf_page(self, offset: int) -> None:
        if self.pdf_document is None:
            return
        target_page = self.pdf_page_index + offset
        if 0 <= target_page < self.pdf_document.page_count:
            self.pdf_page_index = target_page
            self.ocr_summary.set("Ready to recognize the current page.")
            self._show_pdf_page()

    def set_pdf_zoom(self, zoom: float) -> None:
        """Set PDF preview magnification and rerender the current page."""
        self.pdf_zoom = max(0.75, min(2.5, zoom))
        if self.pdf_document is not None:
            self._show_pdf_page()
        self.status.set(f"PDF zoom: {round(self.pdf_zoom * 100)}%.")

    def adjust_pdf_zoom(self, delta: float) -> None:
        self.set_pdf_zoom(self.pdf_zoom + delta)

    def _queue_pdf_render(self, _event=None) -> None:
        if self.pdf_document is None:
            return
        if self.pdf_render_job is not None:
            self.root.after_cancel(self.pdf_render_job)
        self.pdf_render_job = self.root.after(120, self._show_pdf_page)

    def _show_pdf_page(self) -> None:
        self.pdf_render_job = None
        if self.pdf_document is None or self.pdf_canvas.winfo_width() <= 1:
            return
        try:
            import pymupdf
            from PIL import Image

            page = self.pdf_document.load_page(self.pdf_page_index)
            self._size_pdf_viewer(page)
            self.root.update_idletasks()
            available_width = max(self.pdf_canvas.winfo_width() - PDF_VIEWER_PADDING * 2, 100)
            available_height = max(self.pdf_canvas.winfo_height() - PDF_VIEWER_PADDING * 2, 100)
            scale = min(available_width / page.rect.width, available_height / page.rect.height) * self.pdf_zoom
            render_scale = scale * PDF_RENDER_OVERSAMPLE
            pixmap = page.get_pixmap(matrix=pymupdf.Matrix(render_scale, render_scale), alpha=False)
            rendered_image = Image.open(BytesIO(pixmap.tobytes("png")))
            display_size = (
                max(1, round(pixmap.width / PDF_RENDER_OVERSAMPLE)),
                max(1, round(pixmap.height / PDF_RENDER_OVERSAMPLE)),
            )
            rendered_image = rendered_image.resize(display_size, Image.Resampling.LANCZOS)
            image_buffer = BytesIO()
            rendered_image.save(image_buffer, format="PNG")
            image_data = base64.b64encode(image_buffer.getvalue()).decode("ascii")
            self.pdf_image = tk.PhotoImage(data=image_data)
        except Exception as error:
            messagebox.showerror("Could not render PDF page", str(error))
            return

        self.pdf_canvas.delete("all")
        x_offset = max((self.pdf_canvas.winfo_width() - self.pdf_image.width()) // 2, 0)
        y_offset = max((self.pdf_canvas.winfo_height() - self.pdf_image.height()) // 2, 0)
        self.pdf_canvas.create_image(x_offset, y_offset, anchor="nw", image=self.pdf_image)
        self.pdf_canvas.configure(scrollregion=(0, 0, max(self.pdf_canvas.winfo_width(), x_offset + self.pdf_image.width()), max(self.pdf_canvas.winfo_height(), y_offset + self.pdf_image.height())))
        self.pdf_display_bounds = (x_offset, y_offset, self.pdf_image.width(), self.pdf_image.height())
        self.pdf_page_label.set(f"Page {self.pdf_page_index + 1} of {self.pdf_document.page_count}")
        if hasattr(self, "previous_pdf_button"):
            self.previous_pdf_button.configure(state="normal" if self.pdf_page_index > 0 else "disabled")
        is_last_page = self.pdf_page_index >= self.pdf_document.page_count - 1
        if hasattr(self, "next_pdf_button"):
            self.next_pdf_button.configure(state="disabled" if is_last_page else "normal")

    def _begin_pdf_text_selection(self, event) -> None:
        self.pdf_text_selection_start = (event.x, event.y)
        self.pdf_canvas.delete("pdf_text_selection")

    def _draw_pdf_text_selection(self, event) -> None:
        start = getattr(self, "pdf_text_selection_start", None)
        if start is None:
            return
        self.pdf_canvas.delete("pdf_text_selection")
        self.pdf_canvas.create_rectangle(*start, event.x, event.y, outline="#35a7ff", width=2, tags="pdf_text_selection")

    def _finish_pdf_text_selection(self, event) -> None:
        start = getattr(self, "pdf_text_selection_start", None)
        self.pdf_text_selection_start = None
        if start is None or self.pdf_document is None or not self.pdf_display_bounds:
            return
        x0, y0, width, height = self.pdf_display_bounds
        left, right = sorted((max(x0, min(event.x, x0 + width)), max(x0, min(start[0], x0 + width))))
        top, bottom = sorted((max(y0, min(event.y, y0 + height)), max(y0, min(start[1], y0 + height))))
        if right - left < 4 or bottom - top < 4:
            return
        try:
            import pymupdf

            page = self.pdf_document.load_page(self.pdf_page_index)
            scale_x = page.rect.width / width
            scale_y = page.rect.height / height
            clip = pymupdf.Rect((left - x0) * scale_x, (top - y0) * scale_y, (right - x0) * scale_x, (bottom - y0) * scale_y)
            text = " ".join(page.get_text("text", clip=clip).split())
            self._populate_word_from_selection(text, clip)
        except Exception as error:
            messagebox.showerror("Could not select PDF text", str(error))
        finally:
            self.pdf_canvas.delete("pdf_text_selection")

    def _warm_up_ocr_engine(self) -> None:
        """Prepare the compact CPU OCR engine while the user works manually."""
        try:
            self.events.put(("ocr_engine_status", "Loading compact PaddleOCR models..."))
            self.ocr_engine = create_ocr_engine()
            self.events.put(("ocr_engine_status", "OCR ready (CPU, PP-OCRv5 mobile)."))
        except Exception as error:
            self.ocr_engine_error = error
            self.events.put(("ocr_engine_error", str(error)))
        finally:
            self.ocr_engine_ready.set()

    def _populate_word_from_selection(self, text: str, pdf_rect) -> None:
        """Use embedded PDF text immediately, or OCR the region when needed."""
        if text:
            self.word_entry.delete(0, tk.END)
            self.word_entry.insert(0, text)
            self.add_word_to_list()
            return
        if self.pdf_path is None:
            self.status.set("No PDF is open and no text was found in that region.")
            return
        self.status.set("Recognizing the selected region with PaddleOCR...")
        threading.Thread(
            target=self._ocr_selected_region_worker,
            args=(self.pdf_path, self.pdf_page_index, tuple(pdf_rect)),
            daemon=True,
        ).start()

    def _ocr_selected_region_worker(self, pdf_path: Path, page_index: int, pdf_rect: tuple[float, float, float, float]) -> None:
        document = None
        try:
            self.ocr_engine_ready.wait()
            if self.ocr_engine_error is not None:
                raise self.ocr_engine_error
            if self.ocr_engine is None:
                raise RuntimeError("The OCR engine did not finish initializing.")
            document = open_pdf(pdf_path)
            result = process_pdf_region(document, page_index, self.ocr_engine, pdf_rect, dpi=RENDER_DPI)
            text = " ".join(span.text for span in result.spans).strip()
            self.events.put(("ocr_selection_result", text))
        except Exception as error:
            self.events.put(("ocr_selection_error", str(error)))
        finally:
            if document is not None:
                document.close()

    def _size_pdf_viewer(self, page) -> None:
        """Keep the preview sized to the page, leaving workspace for OCR tools."""
        page_scale = min(
            1.0,
            PDF_VIEWER_MAX_WIDTH / page.rect.width,
            PDF_VIEWER_MAX_HEIGHT / page.rect.height,
        )
        width = max(PDF_VIEWER_MIN_WIDTH, round(page.rect.width * page_scale))
        height = max(PDF_VIEWER_MIN_HEIGHT, round(page.rect.height * page_scale))
        self.pdf_viewer.configure(
            width=width + PDF_VIEWER_PADDING * 2,
            height=height + PDF_VIEWER_PADDING * 2,
        )

    def recognize_current_pdf_page(self) -> None:
        if self.pdf_path is None or self.pdf_document is None:
            messagebox.showinfo("Open a PDF", "Open a PDF before running OCR.")
            return

        pdf_path = self.pdf_path
        page_index = self.pdf_page_index
        self.recognize_page_button.configure(state="disabled")
        self.ocr_summary.set(f"Recognizing page {page_index + 1}…")
        threading.Thread(
            target=self._recognize_pdf_page_worker,
            args=(pdf_path, page_index),
            daemon=True,
        ).start()

    def start_entry_selection(self) -> None:
        if self.pdf_document is None or self.pdf_display_bounds is None:
            messagebox.showinfo("Open a PDF", "Open a PDF before creating an entry.")
            return
        self.entry_selection_active = True
        self.selection_start = None
        self.pending_entry_text = ""
        self.pdf_canvas.delete("entry_selection")
        self.pdf_canvas.configure(cursor="crosshair")
        self.entry_summary.set("Draw a box around the text you want to recognize.")
        self.confirm_entry_button.configure(state="disabled")
        self.retry_entry_button.configure(state="disabled")

    def _begin_pdf_selection(self, event) -> None:
        if not self.entry_selection_active or not self._point_in_pdf(event.x, event.y):
            return
        self.selection_start = (event.x, event.y)
        self.pdf_canvas.delete("entry_selection")
        self.selection_rectangle = self.pdf_canvas.create_rectangle(
            event.x,
            event.y,
            event.x,
            event.y,
            outline="#2d8cff",
            width=2,
            tags="entry_selection",
        )

    def _draw_pdf_selection(self, event) -> None:
        if self.selection_start is None or self.selection_rectangle is None:
            return
        x, y = self._clamp_to_pdf(event.x, event.y)
        self.pdf_canvas.coords(self.selection_rectangle, *self.selection_start, x, y)

    def _finish_pdf_selection(self, event) -> None:
        if self.selection_start is None or self.selection_rectangle is None:
            return
        end_x, end_y = self._clamp_to_pdf(event.x, event.y)
        start_x, start_y = self.selection_start
        self.selection_start = None
        if abs(end_x - start_x) < 8 or abs(end_y - start_y) < 8:
            self.pdf_canvas.delete("entry_selection")
            self.selection_rectangle = None
            self.entry_summary.set("Select a larger text region, then release the mouse button.")
            return
        self.pdf_canvas.coords(self.selection_rectangle, start_x, start_y, end_x, end_y)
        self.entry_selection_active = False
        self.pdf_canvas.configure(cursor="")
        self.create_entry_button.configure(state="disabled")
        self.entry_summary.set("Recognizing selected region…")
        pdf_rect = self._canvas_selection_to_pdf_rect(start_x, start_y, end_x, end_y)
        threading.Thread(
            target=self._recognize_pdf_region_worker,
            args=(self.pdf_path, self.pdf_page_index, pdf_rect),
            daemon=True,
        ).start()

    def _point_in_pdf(self, x: int, y: int) -> bool:
        if self.pdf_display_bounds is None:
            return False
        left, top, width, height = self.pdf_display_bounds
        return left <= x <= left + width and top <= y <= top + height

    def _clamp_to_pdf(self, x: int, y: int) -> tuple[int, int]:
        left, top, width, height = self.pdf_display_bounds
        return min(max(x, left), left + width), min(max(y, top), top + height)

    def _canvas_selection_to_pdf_rect(self, x0: int, y0: int, x1: int, y1: int) -> tuple[float, float, float, float]:
        page = self.pdf_document.load_page(self.pdf_page_index)
        left, top, width, height = self.pdf_display_bounds
        return (
            page.rect.x0 + (min(x0, x1) - left) / width * page.rect.width,
            page.rect.y0 + (min(y0, y1) - top) / height * page.rect.height,
            page.rect.x0 + (max(x0, x1) - left) / width * page.rect.width,
            page.rect.y0 + (max(y0, y1) - top) / height * page.rect.height,
        )

    def _recognize_pdf_page_worker(self, pdf_path: Path, page_index: int) -> None:
        document = None
        try:
            if self.ocr_engine is None:
                self.ocr_engine = create_ocr_engine()
            document = open_pdf(pdf_path)
            result = process_pdf_page(document, page_index, self.ocr_engine)
            self.events.put(("ocr_result", result))
        except Exception as error:
            self.events.put(("ocr_error", str(error)))
        finally:
            if document is not None:
                document.close()
            self.events.put(("ocr_done", ""))

    def _recognize_pdf_region_worker(self, pdf_path: Path, page_index: int, pdf_rect) -> None:
        document = None
        try:
            if self.ocr_engine is None:
                self.ocr_engine = create_ocr_engine()
            document = open_pdf(pdf_path)
            result = process_pdf_region(document, page_index, self.ocr_engine, pdf_rect)
            text = "\n".join(span.text for span in result.spans).strip()
            if not text:
                raise ValueError("No text was recognized in the selected region. Try a larger selection.")
            self.events.put(("entry_result", (page_index, text)))
        except Exception as error:
            self.events.put(("entry_error", str(error)))
        finally:
            if document is not None:
                document.close()
            self.events.put(("entry_done", ""))

    def _reset_pending_entry(self) -> None:
        self.entry_selection_active = False
        self.selection_start = None
        self.selection_rectangle = None
        self.pending_entry_text = ""
        self.pdf_canvas.delete("entry_selection")
        self.pdf_canvas.configure(cursor="")
        self.entry_summary.set("Select Create Entry, then draw a box around text on the page.")
        self.confirm_entry_button.configure(state="disabled")
        self.retry_entry_button.configure(state="disabled")

    def _display_ocr_result(self, result) -> None:
        self.ocr_summary.set(f"Recognized {len(result.spans)} text region(s) on page {result.page_index + 1}.")

    def confirm_pending_entry(self) -> None:
        if not self.pending_entry_text:
            return
        position = len(self.created_entries.get_children()) + 1
        self.created_entries.insert(
            "", "end",
            values=(position, self.pdf_page_index + 1, self.pending_entry_text.replace("\n", "  ")),
            tags=("odd" if position % 2 else "even",),
        )
        self._reset_pending_entry()
        self.entry_summary.set("Entry added. Select Create Entry to add another.")

    def _close_pdf_document(self) -> None:
        if self.pdf_document is not None:
            self.pdf_document.close()
            self.pdf_document = None
        self.pdf_path = None
        self.pdf_image = None

    def close_application(self) -> None:
        self._close_pdf_document()
        self.root.destroy()

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

        dictionary_settings = ttk.LabelFrame(container, text="Dictionary character set", padding=12)
        dictionary_settings.grid(column=0, row=3, sticky="ew", pady=(14, 0))
        ttk.Label(
            dictionary_settings,
            text="Choose which character form is shown in dictionary matches and copied into flashcard drafts.",
            wraplength=720,
            justify="left",
        ).grid(column=0, row=0, columnspan=2, sticky="w")
        ttk.Radiobutton(
            dictionary_settings,
            text="Simplified",
            value="simplified",
            variable=self.dictionary_character_set,
            command=self._dictionary_character_set_changed,
        ).grid(column=0, row=1, sticky="w", pady=(8, 0))
        ttk.Radiobutton(
            dictionary_settings,
            text="Traditional",
            value="traditional",
            variable=self.dictionary_character_set,
            command=self._dictionary_character_set_changed,
        ).grid(column=1, row=1, sticky="w", padx=(20, 0), pady=(8, 0))

    def _dictionary_character_set_changed(self) -> None:
        """Refresh dictionary labels without requiring a new lookup."""
        if self.dictionary_entries:
            self._refresh_dictionary_match_labels()
            self._show_selected_dictionary_entry()

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
        self._prepare_dictionary_cards()

    def _prepare_dictionary_cards(self) -> None:
        """Copy the Dictionary-tab queue into the shared Cards-tab preview."""
        import pandas as pd

        queued = [self.dictionary_card_queue.item(item, "values")[1:4] for item in self.dictionary_card_queue.get_children()]
        if not queued:
            self.filtered_rows = None
            self.preview_source = None
            self._clear_preview()
            messagebox.showinfo("Dictionary queue is empty", "Add cards to the Dictionary tab queue first.")
            return
        rows = pd.DataFrame(queued, columns=("Chinese", "Pinyin", "English"))
        self.filtered_rows = rows
        self.preview_source = "Dictionary queue"
        self._clear_preview()
        for _, row in rows.iterrows():
            position = len(self.cards.get_children()) + 1
            self.cards.insert(
                "", "end",
                values=(position, row["Chinese"], row["Pinyin"], row["English"]),
                tags=("odd" if position % 2 else "even",),
            )
        self.status.set(f"{len(rows)} dictionary card(s) ready to add.")

    def _card_source_changed(self, _event=None) -> None:
        """Clear the shared preview when switching between card sources."""
        self.filtered_rows = None
        self.preview_source = None
        self._clear_preview()
        dictionary_source = self.card_source.get() == "Dictionary queue"
        self.lesson_box.configure(state="disabled" if dictionary_source else "readonly")
        self.type_box.configure(state="disabled" if dictionary_source else "readonly")
        if dictionary_source:
            self.status.set("Dictionary queue selected. Click Preview matching cards to load it.")
        else:
            self.status.set("Excel workbook selected. Choose filters, then preview matching cards.")

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
    def _renumber_treeview(tree: ttk.Treeview) -> None:
        """Keep table numbering and alternating row colors consistent after edits."""
        for position, item in enumerate(tree.get_children(), start=1):
            values = list(tree.item(item, "values"))
            if values:
                values[0] = position
                tree.item(item, values=values, tags=("odd" if position % 2 else "even",))

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
                if isinstance(message, str) and message:
                    self.status.set(message)
                if event == "error":
                    messagebox.showerror("Could not add cards", message)
                elif event == "complete":
                    if self.preview_source == "Dictionary queue":
                        for item in self.dictionary_card_queue.get_children():
                            self.dictionary_card_queue.delete(item)
                        self.preview_source = None
                elif event == "done":
                    self.add_button.configure(state="normal")
                elif event == "dictionary_result":
                    query, entries = message
                    self.dictionary_lookup_in_progress = False
                    self._display_dictionary_results(query, entries)
                elif event == "dictionary_error":
                    self.dictionary_lookup_in_progress = False
                    self.dictionary_summary.set("The dictionary could not be loaded.")
                    messagebox.showerror("Could not look up dictionary entry", message)
                elif event == "dictionary_add_error":
                    messagebox.showerror("Could not add dictionary cards", message)
                elif event == "dictionary_complete":
                    for item in self.dictionary_card_queue.get_children():
                        self.dictionary_card_queue.delete(item)
                elif event == "dictionary_done":
                    self.add_button.configure(state="normal")
                elif event == "ocr_selection_result":
                    text = str(message).strip()
                    if text:
                        self.word_entry.delete(0, tk.END)
                        self.word_entry.insert(0, text)
                        self.add_word_to_list()
                    else:
                        self.status.set("PaddleOCR found no text in that region.")
                elif event == "ocr_selection_error":
                    self.status.set("OCR could not recognize the selected region.")
                    messagebox.showerror("Could not recognize selected text", message)
                elif event == "ocr_engine_status":
                    self.ocr_engine_status.set(message)
                elif event == "ocr_engine_error":
                    self.ocr_engine_status.set("OCR unavailable: engine failed to load.")
                    self.status.set("OCR could not be initialized; embedded PDF text remains available.")
                    messagebox.showerror("Could not load OCR", message)
                elif event == "ocr_result":
                    self._display_ocr_result(message)
                elif event == "ocr_error":
                    self.ocr_summary.set("OCR could not be completed.")
                    messagebox.showerror("Could not recognize PDF page", message)
                elif event == "ocr_done":
                    if self.pdf_document is not None:
                        self.recognize_page_button.configure(state="normal")
                elif event == "entry_result":
                    _page_index, text = message
                    self.pending_entry_text = text
                    self.entry_summary.set(f"Review the recognized text:\n{text}")
                    self.confirm_entry_button.configure(state="normal")
                    self.retry_entry_button.configure(state="normal")
                elif event == "entry_error":
                    self.entry_summary.set("No entry was created. Choose Create Entry and select another region.")
                    messagebox.showerror("Could not recognize selected region", message)
                elif event == "entry_done":
                    if self.pdf_document is not None:
                        self.create_entry_button.configure(state="normal")
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

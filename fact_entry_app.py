"""
Fact Entry Companion App.

A lightweight desktop GUI app (tkinter) that runs in the system tray area.
Users copy text from a document they're reviewing, then use this app to
submit that text as a fact linked to the document's Bates number.

Features:
- Clipboard text auto-populated into the fact text field
- Bates number entry (with recent history for quick selection)
- Tag selector with existing tags + ability to add new ones
- Submit writes the fact to the Fact Sheet on SharePoint
- Shows current user name (from Azure AD delegated auth)

Usage:
    python fact_entry_app.py
"""

import logging
import sys
import tkinter as tk
from tkinter import messagebox, ttk
from typing import Optional

import pyperclip

from config.settings import COMPANION_APP_WINDOW_TITLE
from fact_sheet_manager import FactSheetManager
from exhibit_list_manager import ExhibitListManager
from all_docs_indexer import AllDocsIndexer
from graph_api_client import GraphAPIClient

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)


class FactEntryApp:
    """Tkinter-based companion app for creating facts."""

    def __init__(self):
        self.root = tk.Tk()
        self.root.title(COMPANION_APP_WINDOW_TITLE)
        self.root.geometry("650x580")
        self.root.resizable(True, True)

        # Initialize API client with delegated (user) auth
        self.client: Optional[GraphAPIClient] = None
        self.fact_manager: Optional[FactSheetManager] = None
        self.exhibit_manager: Optional[ExhibitListManager] = None
        self.indexer: Optional[AllDocsIndexer] = None
        self.user_name = "Unknown User"
        self.all_tags: list[str] = []
        self.recent_bates: list[str] = []

        self._build_ui()

    def _build_ui(self):
        """Build the GUI layout."""
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # ── Status Bar ──
        status_frame = ttk.Frame(main_frame)
        status_frame.pack(fill=tk.X, pady=(0, 10))

        self.status_label = ttk.Label(status_frame, text="Not connected", foreground="red")
        self.status_label.pack(side=tk.LEFT)

        connect_btn = ttk.Button(status_frame, text="Connect to SharePoint",
                                  command=self._connect)
        connect_btn.pack(side=tk.RIGHT)

        # ── Bates Number ──
        ttk.Label(main_frame, text="Bates Number:", font=("", 10, "bold")).pack(anchor=tk.W)

        bates_frame = ttk.Frame(main_frame)
        bates_frame.pack(fill=tk.X, pady=(2, 8))

        self.bates_var = tk.StringVar()
        self.bates_entry = ttk.Entry(bates_frame, textvariable=self.bates_var, font=("", 11))
        self.bates_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.bates_combo = ttk.Combobox(bates_frame, values=[], width=20, state="readonly")
        self.bates_combo.pack(side=tk.RIGHT, padx=(5, 0))
        self.bates_combo.set("Recent...")
        self.bates_combo.bind("<<ComboboxSelected>>", self._on_recent_bates_selected)

        # ── Fact Text ──
        ttk.Label(main_frame, text="Fact Text:", font=("", 10, "bold")).pack(anchor=tk.W)

        text_frame = ttk.Frame(main_frame)
        text_frame.pack(fill=tk.BOTH, expand=True, pady=(2, 8))

        self.fact_text = tk.Text(text_frame, wrap=tk.WORD, font=("", 10), height=10)
        scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=self.fact_text.yview)
        self.fact_text.configure(yscrollcommand=scrollbar.set)
        self.fact_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        paste_btn = ttk.Button(main_frame, text="Paste from Clipboard",
                                command=self._paste_clipboard)
        paste_btn.pack(anchor=tk.W, pady=(0, 8))

        # ── Tags ──
        ttk.Label(main_frame, text="Tags:", font=("", 10, "bold")).pack(anchor=tk.W)

        tag_frame = ttk.Frame(main_frame)
        tag_frame.pack(fill=tk.X, pady=(2, 8))

        self.tag_combo = ttk.Combobox(tag_frame, values=[], width=30)
        self.tag_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)

        add_tag_btn = ttk.Button(tag_frame, text="Add Tag", command=self._add_tag)
        add_tag_btn.pack(side=tk.RIGHT, padx=(5, 0))

        self.selected_tags_var = tk.StringVar(value="")
        ttk.Label(main_frame, textvariable=self.selected_tags_var,
                  foreground="blue").pack(anchor=tk.W, pady=(0, 8))

        # ── Source (auto-populated) ──
        ttk.Label(main_frame, text="Source / Created By:", font=("", 10, "bold")).pack(anchor=tk.W)
        self.source_var = tk.StringVar(value=self.user_name)
        ttk.Entry(main_frame, textvariable=self.source_var, font=("", 10),
                  state="readonly").pack(fill=tk.X, pady=(2, 8))

        # ── Submit ──
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(5, 0))

        submit_btn = ttk.Button(btn_frame, text="Submit Fact", command=self._submit_fact)
        submit_btn.pack(side=tk.RIGHT)

        clear_btn = ttk.Button(btn_frame, text="Clear", command=self._clear_form)
        clear_btn.pack(side=tk.RIGHT, padx=(0, 5))

        open_doc_btn = ttk.Button(btn_frame, text="Open Document", command=self._open_document)
        open_doc_btn.pack(side=tk.LEFT)

    def _connect(self):
        """Connect to SharePoint and load data."""
        self.status_label.config(text="Connecting...", foreground="orange")
        self.root.update()

        try:
            self.client = GraphAPIClient(use_delegated_auth=True)
            self.user_name = self.client.get_current_user_name()
            self.source_var.set(self.user_name)

            # Load fact sheet
            self.fact_manager = FactSheetManager(self.client)
            self.fact_manager.load()

            # Load tags
            self.all_tags = self.fact_manager.get_all_tags()
            self.tag_combo["values"] = self.all_tags

            # Load indexer (for metadata lookup)
            self.indexer = AllDocsIndexer(self.client)
            # Don't load full index here — too slow. We'll just use exhibit list.

            self.status_label.config(
                text=f"Connected as {self.user_name}", foreground="green"
            )
            logger.info(f"Connected as {self.user_name}")

        except Exception as e:
            self.status_label.config(text=f"Connection failed: {e}", foreground="red")
            logger.error(f"Connection failed: {e}")
            messagebox.showerror("Connection Error", str(e))

    def _paste_clipboard(self):
        """Paste clipboard content into the fact text field."""
        try:
            text = pyperclip.paste()
            if text:
                self.fact_text.delete("1.0", tk.END)
                self.fact_text.insert("1.0", text)
        except Exception as e:
            messagebox.showwarning("Clipboard", f"Could not read clipboard: {e}")

    def _on_recent_bates_selected(self, event):
        """Handle selection from the recent Bates dropdown."""
        selected = self.bates_combo.get()
        if selected and selected != "Recent...":
            self.bates_var.set(selected)

    def _add_tag(self):
        """Add the selected/typed tag to the tag list."""
        tag = self.tag_combo.get().strip()
        if not tag:
            return

        current = self.selected_tags_var.get()
        current_tags = [t.strip() for t in current.split(",") if t.strip()]

        if tag not in current_tags:
            current_tags.append(tag)
            self.selected_tags_var.set(", ".join(current_tags))

        # Add to known tags if new
        if tag not in self.all_tags:
            self.all_tags.append(tag)
            self.all_tags.sort()
            self.tag_combo["values"] = self.all_tags

        self.tag_combo.set("")

    def _open_document(self):
        """Open the document for the current Bates number in the native app."""
        bates = self.bates_var.get().strip()
        if not bates:
            messagebox.showwarning("Missing Info", "Enter a Bates number first.")
            return

        if not self.client:
            messagebox.showwarning("Not Connected", "Connect to SharePoint first.")
            return

        try:
            from config.settings import DOCUMENTS_FOLDER
            folder_files = self.client.list_files_in_folder(DOCUMENTS_FOLDER)
            if bates in folder_files:
                url = folder_files[bates]["webUrl"]
                import webbrowser
                webbrowser.open(url)
            else:
                messagebox.showinfo("Not Found", f"No document found for Bates '{bates}'.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def _submit_fact(self):
        """Submit the fact to the Fact Sheet."""
        bates = self.bates_var.get().strip()
        fact_text = self.fact_text.get("1.0", tk.END).strip()
        tags = self.selected_tags_var.get().strip()

        if not bates:
            messagebox.showwarning("Missing Info", "Bates number is required.")
            return
        if not fact_text:
            messagebox.showwarning("Missing Info", "Fact text is required.")
            return
        if not self.fact_manager:
            messagebox.showwarning("Not Connected", "Connect to SharePoint first.")
            return

        try:
            fact_id = self.fact_manager.add_fact(
                bates_number=bates,
                fact_text=fact_text,
                source=self.user_name,
                tags=tags if tags else "Untagged",
                created_by=self.user_name,
            )

            # Save to SharePoint
            self.fact_manager.save()

            # Update recent Bates list
            if bates not in self.recent_bates:
                self.recent_bates.insert(0, bates)
                self.recent_bates = self.recent_bates[:20]
                self.bates_combo["values"] = self.recent_bates

            messagebox.showinfo("Success", f"Fact {fact_id} created successfully!")
            self._clear_form()

        except Exception as e:
            logger.error(f"Failed to submit fact: {e}")
            messagebox.showerror("Error", f"Failed to submit fact:\n{e}")

    def _clear_form(self):
        """Clear all form fields."""
        self.fact_text.delete("1.0", tk.END)
        self.selected_tags_var.set("")
        # Keep Bates number for consecutive facts from the same document

    def run(self):
        """Start the application."""
        self.root.mainloop()


def main():
    app = FactEntryApp()
    app.run()


if __name__ == "__main__":
    main()

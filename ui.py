"""Tkinter desktop interface. Contains no business logic."""

from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, ttk

import config
from agent import WELCOME_TEXT, CaseAgent

HEADER_BG = "#1F4E78"
USER_BG = "#DCEBFA"
AGENT_BG = "#F1F1F1"
ERROR_FG = "#B00020"
READY_FG = "#1B7F3B"
IDLE_FG = "#6B6B6B"
FONT = ("Segoe UI", 10)


class CaseAgentApp(tk.Tk):
    """Main window: folder picker, status, chat history and command entry."""

    def __init__(self, agent: CaseAgent | None = None) -> None:
        super().__init__()
        self.agent = agent or CaseAgent()
        self.title(config.APP_TITLE)
        self.geometry("900x680")
        self.minsize(720, 520)
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        self.folder_var = tk.StringVar(value="No folder selected")
        self.entry_var = tk.StringVar()

        self._build_header()
        self._build_folder_bar()
        self._build_chat()
        self._build_input()

        self._set_status("No folder selected", IDLE_FG)
        self._append_agent(f"{WELCOME_TEXT}\n\nMode: {self.agent.mode_label}")
        self.entry.focus_set()

    # ------------------------------------------------------------------ layout
    def _build_header(self) -> None:
        header = tk.Frame(self, bg=HEADER_BG, padx=16, pady=12)
        header.grid(row=0, column=0, sticky="ew")
        tk.Label(
            header, text=config.APP_TITLE, bg=HEADER_BG, fg="white",
            font=("Segoe UI", 18, "bold"),
        ).pack(anchor="w")
        tk.Label(
            header, text=config.APP_DESCRIPTION, bg=HEADER_BG, fg="#D6E4F0", font=FONT
        ).pack(anchor="w")
        tk.Label(
            header, text=f"Mode: {self.agent.mode_label}", bg=HEADER_BG, fg="#D6E4F0",
            font=("Segoe UI", 9, "italic"),
        ).pack(anchor="w")

    def _build_folder_bar(self) -> None:
        bar = ttk.Frame(self, padding=(16, 10))
        bar.grid(row=1, column=0, sticky="ew")
        bar.columnconfigure(1, weight=1)
        ttk.Button(bar, text="Select Folder", command=self._on_select_folder).grid(
            row=0, column=0, padx=(0, 12)
        )
        ttk.Label(bar, textvariable=self.folder_var).grid(row=0, column=1, sticky="w")
        self.status_label = ttk.Label(bar, font=("Segoe UI", 10, "bold"))
        self.status_label.grid(row=0, column=2, sticky="e", padx=(12, 0))

    def _build_chat(self) -> None:
        frame = ttk.Frame(self, padding=(16, 0))
        frame.grid(row=2, column=0, sticky="nsew")
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)

        self.chat = tk.Text(
            frame, wrap="word", state="disabled", font=FONT, padx=10, pady=10,
            relief="solid", borderwidth=1, background="white",
        )
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=self.chat.yview)
        self.chat.configure(yscrollcommand=scrollbar.set)
        self.chat.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        small = ("Segoe UI", 8, "bold")
        self.chat.tag_configure("user_label", foreground="#35618F", font=small, justify="right")
        self.chat.tag_configure(
            "user_msg", background=USER_BG, lmargin1=160, lmargin2=160, rmargin=8,
            spacing1=2, spacing3=2,
        )
        self.chat.tag_configure("agent_label", foreground="#555555", font=small)
        self.chat.tag_configure(
            "agent_msg", background=AGENT_BG, lmargin1=8, lmargin2=8, rmargin=160,
            spacing1=2, spacing3=2,
        )
        self.chat.tag_configure(
            "error_msg", background="#FBE9EC", foreground=ERROR_FG, lmargin1=8,
            lmargin2=8, rmargin=160, spacing1=2, spacing3=2,
        )

    def _build_input(self) -> None:
        bar = ttk.Frame(self, padding=16)
        bar.grid(row=3, column=0, sticky="ew")
        bar.columnconfigure(0, weight=1)
        self.entry = ttk.Entry(bar, textvariable=self.entry_var, font=FONT)
        self.entry.grid(row=0, column=0, sticky="ew", ipady=4)
        self.entry.bind("<Return>", self._on_send)
        ttk.Button(bar, text="Send", command=self._on_send).grid(row=0, column=1, padx=(8, 0))
        ttk.Button(bar, text="Clear Chat", command=self._on_clear).grid(
            row=0, column=2, padx=(8, 0)
        )

    # ---------------------------------------------------------------- handlers
    def _on_select_folder(self) -> None:
        chosen = filedialog.askdirectory(
            title=f"Select a folder for {config.WORKBOOK_NAME}", mustexist=True
        )
        if not chosen:
            return
        folder = Path(chosen)
        self.folder_var.set(str(folder))
        response = self.agent.select_folder(folder)
        if self.agent.is_ready:
            self._set_status("Workbook ready", READY_FG)
        else:
            self._set_status("Workbook not ready", ERROR_FG)
        self._append_agent(response.text, error=response.is_error)

    def _on_send(self, _event: object = None) -> str:
        text = self.entry_var.get().strip()
        if not text:
            return "break"
        self.entry_var.set("")
        self._append_message("You", text, "user_label", "user_msg")
        response = self.agent.handle(text)
        self._append_agent(response.text, error=response.is_error)
        return "break"

    def _on_clear(self) -> None:
        """Clear the visible conversation only; Excel records are untouched."""
        self.chat.configure(state="normal")
        self.chat.delete("1.0", "end")
        self.chat.configure(state="disabled")
        self._append_agent("Chat cleared. Your cases in Excel were not changed.")

    # ----------------------------------------------------------------- helpers
    def _set_status(self, text: str, color: str) -> None:
        self.status_label.configure(text=f"\u25CF {text}", foreground=color)

    def _append_agent(self, text: str, error: bool = False) -> None:
        self._append_message("Agent", text, "agent_label", "error_msg" if error else "agent_msg")

    def _append_message(self, sender: str, text: str, label_tag: str, msg_tag: str) -> None:
        self.chat.configure(state="normal")
        self.chat.insert("end", f"{sender}\n", label_tag)
        self.chat.insert("end", f"{text}\n", msg_tag)
        self.chat.insert("end", "\n")
        self.chat.configure(state="disabled")
        self.chat.see("end")

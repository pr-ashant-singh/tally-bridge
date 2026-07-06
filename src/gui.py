"""
TallyBridge GUI module.

A modern CustomTkinter-based desktop interface for converting
Zerodha Tax P&L Excel files to Tally-compatible format.
"""

import os
import sys
import threading
import traceback
from pathlib import Path

import customtkinter as ctk
from tkinter import filedialog, messagebox

from src.parser import parse_file
from src.generator import generate_all


# App appearance
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


class TallyBridgeApp(ctk.CTk):
    """Main application window."""

    def __init__(self):
        super().__init__()

        self.title("TallyBridge - Zerodha to Tally Converter")
        self.geometry("700x580")
        self.minsize(600, 500)
        self.resizable(True, True)

        # State
        self.input_file = None
        self.output_dir = None

        self._build_ui()

    def _build_ui(self):
        """Build the complete UI layout."""
        # Main container with padding
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)

        # --- Header ---
        header_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        header_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=(20, 10))

        ctk.CTkLabel(
            header_frame,
            text="TallyBridge",
            font=ctk.CTkFont(size=28, weight="bold"),
        ).pack(anchor="w")

        ctk.CTkLabel(
            header_frame,
            text="Convert Zerodha Tax P&L reports to Tally-compatible Excel files",
            font=ctk.CTkFont(size=13),
            text_color="gray",
        ).pack(anchor="w", pady=(2, 0))

        # --- Input Section ---
        input_frame = ctk.CTkFrame(self)
        input_frame.grid(row=1, column=0, sticky="ew", padx=20, pady=10)
        input_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            input_frame,
            text="📁 Input File",
            font=ctk.CTkFont(size=14, weight="bold"),
        ).grid(row=0, column=0, columnspan=3, sticky="w", padx=15, pady=(15, 5))

        ctk.CTkLabel(
            input_frame,
            text="Zerodha Tax P&L Excel:",
            font=ctk.CTkFont(size=12),
        ).grid(row=1, column=0, sticky="w", padx=15, pady=5)

        self.file_entry = ctk.CTkEntry(
            input_frame,
            placeholder_text="Select your taxpnl-*.xlsx file...",
            height=36,
        )
        self.file_entry.grid(row=1, column=1, sticky="ew", padx=5, pady=5)

        ctk.CTkButton(
            input_frame,
            text="Browse",
            width=90,
            height=36,
            command=self._browse_file,
        ).grid(row=1, column=2, padx=(5, 15), pady=5)

        # --- Output Section ---
        output_frame = ctk.CTkFrame(self)
        output_frame.grid(row=2, column=0, sticky="ew", padx=20, pady=10)
        output_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            output_frame,
            text="📂 Output Directory",
            font=ctk.CTkFont(size=14, weight="bold"),
        ).grid(row=0, column=0, columnspan=3, sticky="w", padx=15, pady=(15, 5))

        ctk.CTkLabel(
            output_frame,
            text="Save files to:",
            font=ctk.CTkFont(size=12),
        ).grid(row=1, column=0, sticky="w", padx=15, pady=5)

        self.output_entry = ctk.CTkEntry(
            output_frame,
            placeholder_text="Same folder as input file (default)",
            height=36,
        )
        self.output_entry.grid(row=1, column=1, sticky="ew", padx=5, pady=5)

        ctk.CTkButton(
            output_frame,
            text="Browse",
            width=90,
            height=36,
            command=self._browse_output,
        ).grid(row=1, column=2, padx=(5, 15), pady=5)

        # --- Results / Log Section ---
        results_frame = ctk.CTkFrame(self)
        results_frame.grid(row=3, column=0, sticky="nsew", padx=20, pady=10)
        results_frame.grid_columnconfigure(0, weight=1)
        results_frame.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(
            results_frame,
            text="📋 Status",
            font=ctk.CTkFont(size=14, weight="bold"),
        ).grid(row=0, column=0, sticky="w", padx=15, pady=(15, 5))

        self.log_textbox = ctk.CTkTextbox(
            results_frame,
            height=120,
            font=ctk.CTkFont(family="Courier", size=12),
        )
        self.log_textbox.grid(row=1, column=0, sticky="nsew", padx=15, pady=(5, 15))
        self.log_textbox.configure(state="disabled")

        # --- Action Buttons ---
        button_frame = ctk.CTkFrame(self, fg_color="transparent")
        button_frame.grid(row=4, column=0, sticky="ew", padx=20, pady=(5, 20))

        self.generate_btn = ctk.CTkButton(
            button_frame,
            text="🚀 Generate Tally Files",
            font=ctk.CTkFont(size=15, weight="bold"),
            height=45,
            command=self._generate,
        )
        self.generate_btn.pack(fill="x", pady=5)

        self.progress_bar = ctk.CTkProgressBar(button_frame)
        self.progress_bar.pack(fill="x", pady=(5, 0))
        self.progress_bar.set(0)

    def _browse_file(self):
        """Open file browser for input Excel file."""
        filepath = filedialog.askopenfilename(
            title="Select Zerodha Tax P&L Excel",
            filetypes=[
                ("Excel files", "*.xlsx"),
                ("All files", "*.*"),
            ],
        )
        if filepath:
            self.input_file = filepath
            self.file_entry.delete(0, "end")
            self.file_entry.insert(0, filepath)

            # Default output to same directory
            if not self.output_dir:
                default_output = os.path.join(
                    os.path.dirname(filepath), "TallyBridge_Output"
                )
                self.output_entry.delete(0, "end")
                self.output_entry.insert(0, default_output)
                self.output_dir = default_output

    def _browse_output(self):
        """Open directory browser for output folder."""
        dirpath = filedialog.askdirectory(title="Select Output Directory")
        if dirpath:
            self.output_dir = dirpath
            self.output_entry.delete(0, "end")
            self.output_entry.insert(0, dirpath)

    def _log(self, message: str):
        """Append a message to the log textbox."""
        self.log_textbox.configure(state="normal")
        self.log_textbox.insert("end", message + "\n")
        self.log_textbox.see("end")
        self.log_textbox.configure(state="disabled")

    def _clear_log(self):
        """Clear the log textbox."""
        self.log_textbox.configure(state="normal")
        self.log_textbox.delete("1.0", "end")
        self.log_textbox.configure(state="disabled")

    def _generate(self):
        """Start the generation process in a background thread."""
        # Validate input
        input_path = self.file_entry.get().strip()
        if not input_path:
            messagebox.showerror("Error", "Please select an input file.")
            return

        if not os.path.isfile(input_path):
            messagebox.showerror("Error", f"File not found:\n{input_path}")
            return

        if not input_path.endswith(".xlsx"):
            messagebox.showerror("Error", "Please select a valid .xlsx file.")
            return

        output_path = self.output_entry.get().strip()
        if not output_path:
            output_path = os.path.join(
                os.path.dirname(input_path), "TallyBridge_Output"
            )

        self.input_file = input_path
        self.output_dir = output_path

        # Disable button and start processing
        self.generate_btn.configure(state="disabled", text="Processing...")
        self.progress_bar.set(0)
        self._clear_log()
        self._log("Starting conversion...")

        # Run in background thread to keep UI responsive
        thread = threading.Thread(target=self._run_generation, daemon=True)
        thread.start()

    def _run_generation(self):
        """Background thread for file processing."""
        try:
            # Step 1: Parse
            self.after(0, lambda: self._log("📖 Parsing input file..."))
            self.after(0, lambda: self.progress_bar.set(0.2))

            data = parse_file(self.input_file)

            self.after(0, lambda: self._log(
                f"   Client: {data.client_name} ({data.client_id})"
            ))
            self.after(0, lambda: self._log(
                f"   Trades: {len(data.trades)} "
                f"(Profit: {len(data.profit_trades)}, Loss: {len(data.loss_trades)})"
            ))
            self.after(0, lambda: self._log(
                f"   Charges: {len(data.charges)} | Dividends: {len(data.dividends)}"
            ))
            self.after(0, lambda: self.progress_bar.set(0.5))

            # Step 2: Generate output
            self.after(0, lambda: self._log("\n📝 Generating Tally files..."))

            files = generate_all(data, self.output_dir)

            self.after(0, lambda: self.progress_bar.set(0.9))

            # Step 3: Report results
            self.after(0, lambda: self._log(f"\n✅ Done! {len(files)} files created:"))
            for f in files:
                size = os.path.getsize(f) / 1024
                name = os.path.basename(f)
                self.after(0, lambda n=name, s=size: self._log(f"   📄 {n} ({s:.1f} KB)"))

            self.after(0, lambda: self._log(f"\n📁 Output: {self.output_dir}"))
            self.after(0, lambda: self.progress_bar.set(1.0))

            # Show success dialog
            self.after(0, lambda: messagebox.showinfo(
                "Success",
                f"Generated {len(files)} Tally files!\n\n"
                f"Output folder:\n{self.output_dir}",
            ))

        except Exception as e:
            error_msg = f"❌ Error: {str(e)}"
            self.after(0, lambda: self._log(error_msg))
            self.after(0, lambda: self._log(traceback.format_exc()))
            self.after(0, lambda: messagebox.showerror(
                "Error", f"Failed to process file:\n\n{str(e)}"
            ))

        finally:
            self.after(0, lambda: self.generate_btn.configure(
                state="normal", text="🚀 Generate Tally Files"
            ))


def run_app():
    """Launch the TallyBridge application."""
    app = TallyBridgeApp()
    app.mainloop()

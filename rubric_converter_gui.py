"""Tkinter interface for converting Turnitin rubric files."""

from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from rubric_converter import RubricConversionError, build_table, convert_rubric, load_rubric


class RubricConverterApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Turnitin Rubric Converter")
        self.root.minsize(700, 500)
        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.output_format = tk.StringVar(value="CSV")
        self.use_name_and_value = tk.BooleanVar()
        self.status = tk.StringVar(value="Select an .rbc file to begin.")
        self._build_ui()

    def _build_ui(self) -> None:
        frame = ttk.Frame(self.root, padding=16)
        frame.pack(fill=tk.BOTH, expand=True)
        frame.columnconfigure(1, weight=1)
        ttk.Label(frame, text="Input .rbc file:").grid(row=0, column=0, sticky="w", pady=5)
        ttk.Entry(frame, textvariable=self.input_path).grid(row=0, column=1, sticky="ew", pady=5)
        ttk.Button(frame, text="Browse...", command=self.select_input).grid(row=0, column=2, padx=(8, 0))
        ttk.Label(frame, text="Output file:").grid(row=1, column=0, sticky="w", pady=5)
        ttk.Entry(frame, textvariable=self.output_path).grid(row=1, column=1, sticky="ew", pady=5)
        ttk.Button(frame, text="Browse...", command=self.select_output).grid(row=1, column=2, padx=(8, 0))
        ttk.Label(frame, text="Format:").grid(row=2, column=0, sticky="w", pady=5)
        format_box = ttk.Combobox(
            frame, textvariable=self.output_format, values=("CSV", "Excel"), state="readonly", width=12
        )
        format_box.grid(row=2, column=1, sticky="w", pady=5)
        format_box.bind("<<ComboboxSelected>>", lambda _event: self._update_extension())
        ttk.Checkbutton(
            frame, text="Use criterion name and value in the Criteria column",
            variable=self.use_name_and_value, command=self.preview,
        ).grid(row=3, column=0, columnspan=3, sticky="w", pady=5)
        buttons = ttk.Frame(frame)
        buttons.grid(row=4, column=0, columnspan=3, sticky="w", pady=(8, 12))
        ttk.Button(buttons, text="Preview", command=self.preview).pack(side=tk.LEFT)
        ttk.Button(buttons, text="Convert", command=self.convert).pack(side=tk.LEFT, padx=8)
        ttk.Label(frame, textvariable=self.status).grid(row=5, column=0, columnspan=3, sticky="w")
        self.preview_tree = ttk.Treeview(frame, show="headings", height=12)
        self.preview_tree.grid(row=6, column=0, columnspan=3, sticky="nsew", pady=(10, 0))
        frame.rowconfigure(6, weight=1)

    def select_input(self) -> None:
        selected = filedialog.askopenfilename(
            title="Select Turnitin rubric", filetypes=[("Rubric files", "*.rbc"), ("All files", "*.*")]
        )
        if selected:
            self.input_path.set(selected)
            if not self.output_path.get():
                self.output_path.set(str(Path(selected).with_suffix(".csv")))
            self.preview()

    def select_output(self) -> None:
        extension = ".xlsx" if self.output_format.get() == "Excel" else ".csv"
        selected = filedialog.asksaveasfilename(
            title="Save converted rubric", defaultextension=extension,
            filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if selected:
            self.output_path.set(selected)

    def _update_extension(self) -> None:
        path = self.output_path.get()
        if path:
            self.output_path.set(str(Path(path).with_suffix(".xlsx" if self.output_format.get() == "Excel" else ".csv")))

    def preview(self) -> None:
        try:
            table = build_table(load_rubric(self.input_path.get()), self.use_name_and_value.get())
        except RubricConversionError as exc:
            self.status.set(str(exc))
            return
        self.preview_tree.delete(*self.preview_tree.get_children())
        self.preview_tree["columns"] = table.headers
        for header in table.headers:
            self.preview_tree.heading(header, text=header)
            self.preview_tree.column(header, width=160, anchor="w")
        for row in table.rows[:100]:
            self.preview_tree.insert("", tk.END, values=row)
        self.status.set(f"Previewing {len(table.rows)} criteria.")

    def convert(self) -> None:
        if not self.input_path.get() or not self.output_path.get():
            messagebox.showerror("Missing file", "Select both an input rubric and an output file.")
            return
        output = Path(self.output_path.get())
        expected = ".xlsx" if self.output_format.get() == "Excel" else ".csv"
        if output.suffix.lower() != expected:
            output = output.with_suffix(expected)
            self.output_path.set(str(output))
        if output.exists() and not messagebox.askyesno("Overwrite file?", f"{output} already exists. Replace it?"):
            return
        try:
            convert_rubric(self.input_path.get(), output, self.output_format.get(), self.use_name_and_value.get())
        except RubricConversionError as exc:
            messagebox.showerror("Conversion failed", str(exc))
            return
        self.status.set(f"Saved {output}")
        messagebox.showinfo("Conversion complete", f"File saved to:\n{output}")


def main() -> None:
    root = tk.Tk()
    RubricConverterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

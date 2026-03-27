import csv
import socket
import time
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

import openpyxl
import pandas as pd


SCRIPT_VERSION = "4.0"
DEFAULT_KEY_COLUMN = "LegacyId_m__c"
SAMMENLIGN_EXCEL = Path("sammenlign.xlsx")

SAMMENLIGN_COLUMNS = [
    "tidsstempel",
    "klynge",
    "legacy_fil",
    "target_fil",
    "primaernoegle_kolonne",
    "raekker_legacy",
    "raekker_target",
    "kolonner_sammenlignet",
    "celler_sammenlignet",
    "afvigelser",
    "status",
    "eksekveringstid_ms",
    "maskine",
    "script_version",
]


def detect_delimiter(file_path, encoding="utf-8-sig"):
    with open(file_path, "r", encoding=encoding, newline="") as f:
        sample = f.read(4096)
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
            return dialect.delimiter
        except csv.Error:
            return ","


def load_csv_normalized(file_path):
    delimiter = detect_delimiter(file_path)

    df = pd.read_csv(
        file_path,
        sep=delimiter,
        dtype=str,
        keep_default_na=False,
        encoding="utf-8-sig"
    )

    df.columns = [str(col).strip() for col in df.columns]

    for col in df.columns:
        df[col] = df[col].astype(str).map(lambda x: x.strip())

    return df


def try_parse_decimal(value):
    if value is None:
        return None

    text = str(value).strip()
    if text == "":
        return None

    text = text.replace(" ", "")

    if "," in text and "." in text:
        candidate = text.replace(",", "")
        try:
            return Decimal(candidate)
        except InvalidOperation:
            pass

    if "," in text and "." not in text:
        text = text.replace(",", ".")

    try:
        return Decimal(text)
    except InvalidOperation:
        return None


def format_decimal(value):
    if value is None:
        return ""
    normalized = value.normalize()
    text = format(normalized, "f")
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    return text


def identify_numeric_columns(df):
    numeric_columns = []

    for col in df.columns:
        non_empty_values = [str(v).strip() for v in df[col].tolist() if str(v).strip() != ""]
        if not non_empty_values:
            continue

        parsed = [try_parse_decimal(v) for v in non_empty_values]
        valid_count = sum(1 for p in parsed if p is not None)

        if valid_count == len(non_empty_values):
            numeric_columns.append(col)

    return numeric_columns


def calculate_numeric_sums(df, numeric_columns):
    result = {}
    for col in numeric_columns:
        total = Decimal("0")
        for value in df[col].tolist():
            parsed = try_parse_decimal(value)
            if parsed is not None:
                total += parsed
        result[col] = total
    return result


def build_row_map(df, key_column=None):
    row_map = {}

    if key_column and key_column in df.columns:
        for idx, row in df.iterrows():
            key_value = str(row.get(key_column, "")).strip()
            if key_value == "":
                key_value = f"ROW_{idx + 1}"

            original_key = key_value
            suffix = 2
            while key_value in row_map:
                key_value = f"{original_key}__{suffix}"
                suffix += 1

            row_map[key_value] = row.to_dict()
    else:
        for idx, row in df.iterrows():
            row_map[f"ROW_{idx + 1}"] = row.to_dict()

    return row_map


def compare_dataframes(df_legacy, df_target, key_column=None):
    legacy_cols = list(df_legacy.columns)
    target_cols = list(df_target.columns)

    common_cols = sorted(set(legacy_cols) & set(target_cols))
    only_legacy_cols = sorted(set(legacy_cols) - set(target_cols))
    only_target_cols = sorted(set(target_cols) - set(legacy_cols))

    legacy_map = build_row_map(df_legacy, key_column)
    target_map = build_row_map(df_target, key_column)

    all_row_ids = sorted(set(legacy_map.keys()) | set(target_map.keys()))

    differences = []
    compared_cell_count = 0

    for row_id in all_row_ids:
        legacy_row = legacy_map.get(row_id)
        target_row = target_map.get(row_id)

        if legacy_row is None:
            differences.append({
                "column": "<hele rækken>",
                "row_id": row_id,
                "legacy_value": "<mangler>",
                "target_value": "Række findes kun i target-fil"
            })
            continue

        if target_row is None:
            differences.append({
                "column": "<hele rækken>",
                "row_id": row_id,
                "legacy_value": "Række findes kun i legacy-fil",
                "target_value": "<mangler>"
            })
            continue

        for col in common_cols:
            compared_cell_count += 1
            legacy_value = str(legacy_row.get(col, ""))
            target_value = str(target_row.get(col, ""))

            if legacy_value != target_value:
                differences.append({
                    "column": col,
                    "row_id": row_id,
                    "legacy_value": legacy_value,
                    "target_value": target_value
                })

    for col in only_legacy_cols:
        differences.append({
            "column": col,
            "row_id": "<ikke relevant>",
            "legacy_value": "Kolonnen findes kun i legacy-fil",
            "target_value": "<mangler>"
        })

    for col in only_target_cols:
        differences.append({
            "column": col,
            "row_id": "<ikke relevant>",
            "legacy_value": "<mangler>",
            "target_value": "Kolonnen findes kun i target-fil"
        })

    return {
        "common_cols": common_cols,
        "only_legacy_cols": only_legacy_cols,
        "only_target_cols": only_target_cols,
        "differences": differences,
        "compared_cell_count": compared_cell_count
    }


def build_report(
    legacy_file,
    target_file,
    cluster_id="",
    key_column="",
    script_version=SCRIPT_VERSION
):
    start = time.perf_counter()

    df_legacy = load_csv_normalized(legacy_file)
    df_target = load_csv_normalized(target_file)

    numeric_cols_legacy = set(identify_numeric_columns(df_legacy))
    numeric_cols_target = set(identify_numeric_columns(df_target))
    numeric_common_cols = sorted(numeric_cols_legacy & numeric_cols_target)

    sums_legacy = calculate_numeric_sums(df_legacy, numeric_common_cols)
    sums_target = calculate_numeric_sums(df_target, numeric_common_cols)

    compare_result = compare_dataframes(df_legacy, df_target, key_column=key_column.strip() or None)
    differences = compare_result["differences"]

    status = "IDENTISKE" if len(differences) == 0 else "FORSKELLE FUNDET"
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    runtime_ms = int((time.perf_counter() - start) * 1000)
    hostname = socket.gethostname()

    lines = []
    lines.append("========================================")
    lines.append("         DATA SAMMENLIGNINGSRAPPORT")
    lines.append("========================================")
    lines.append("")
    lines.append(f"Klynge: {cluster_id or '<ikke angivet>'}")
    lines.append(f"Legacy-fil: {legacy_file}")
    lines.append(f"Target-fil: {target_file}")
    lines.append(f"Dato for sammenligning: {timestamp}")
    lines.append(f"Diff-script version: {script_version}")
    lines.append("")

    lines.append("----------------------------------------")
    lines.append("1. Rækketælling")
    lines.append("----------------------------------------")
    lines.append(f"Antal rækker i legacy-fil:      {len(df_legacy)}")
    lines.append(f"Antal rækker i target-fil:      {len(df_target)}")
    lines.append("")

    lines.append("----------------------------------------")
    lines.append("2. Numeriske kolonne-summer")
    lines.append("----------------------------------------")

    if numeric_common_cols:
        for col in numeric_common_cols:
            lines.append(f"Kolonne: {col}")
            lines.append(f"  SUM legacy: {format_decimal(sums_legacy.get(col))}")
            lines.append(f"  SUM target: {format_decimal(sums_target.get(col))}")
            lines.append("")
    else:
        lines.append("Ingen fælles numeriske kolonner fundet.")
        lines.append("")

    lines.append("----------------------------------------")
    lines.append("3. Kolonne- og cellematch")
    lines.append("----------------------------------------")
    lines.append(f"Antal kolonner sammenlignet:    {len(compare_result['common_cols'])}")
    lines.append(f"Antal celler sammenlignet:      {compare_result['compared_cell_count']}")
    lines.append(f"Antal afvigelser fundet:        {len(differences)}")
    lines.append("")

    lines.append("----------------------------------------")
    lines.append("4. Resultat")
    lines.append("----------------------------------------")
    lines.append(f"STATUS: {status}")
    lines.append("")

    if compare_result["only_legacy_cols"]:
        lines.append("Kolonner kun i legacy-fil:")
        for col in compare_result["only_legacy_cols"]:
            lines.append(f"- {col}")
        lines.append("")

    if compare_result["only_target_cols"]:
        lines.append("Kolonner kun i target-fil:")
        for col in compare_result["only_target_cols"]:
            lines.append(f"- {col}")
        lines.append("")

    for diff in differences:
        lines.append(f"- Kolonne: {diff['column']}")
        lines.append(f"- Række-ID / primærnøgle: {diff['row_id']}")
        lines.append(f"- Legacy værdi: {diff['legacy_value']}")
        lines.append(f"- Target værdi: {diff['target_value']}")
        lines.append("")

    lines.append("----------------------------------------")
    lines.append("5. Runtime-information")
    lines.append("----------------------------------------")
    lines.append(f"Script eksekveringstid: {runtime_ms} ms")
    lines.append(f"Miljø/Maskine: {hostname}")
    lines.append("")
    lines.append("========================================")
    lines.append("                SLUT PÅ RAPPORT")
    lines.append("========================================")

    report_text = "\n".join(lines)

    summary_row = {
        "tidsstempel": timestamp,
        "klynge": cluster_id or "",
        "legacy_fil": str(legacy_file),
        "target_fil": str(target_file),
        "primaernoegle_kolonne": key_column.strip(),
        "raekker_legacy": len(df_legacy),
        "raekker_target": len(df_target),
        "kolonner_sammenlignet": len(compare_result["common_cols"]),
        "celler_sammenlignet": compare_result["compared_cell_count"],
        "afvigelser": len(differences),
        "status": status,
        "eksekveringstid_ms": runtime_ms,
        "maskine": hostname,
        "script_version": script_version,
    }

    return report_text, status == "IDENTISKE", summary_row


def append_to_sammenlign_excel(summary_row, excel_path=SAMMENLIGN_EXCEL):
    if excel_path.exists():
        wb = openpyxl.load_workbook(excel_path)
        ws = wb.active
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sammenligninger"
        ws.append(SAMMENLIGN_COLUMNS)

    ws.append([summary_row[col] for col in SAMMENLIGN_COLUMNS])
    wb.save(excel_path)


class CsvCompareApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CSV-sammenligning med rapport")
        self.root.geometry("1150x780")

        self.file1_path = tk.StringVar()
        self.file2_path = tk.StringVar()
        self.cluster_id_var = tk.StringVar()
        self.key_column_var = tk.StringVar()

        self.last_report_text = ""
        self.last_summary_row = None

        self.build_ui()

    def build_ui(self):
        frame_top = tk.Frame(self.root, padx=10, pady=10)
        frame_top.pack(fill="x")

        tk.Label(frame_top, text="Legacy-fil:").grid(row=0, column=0, sticky="w", pady=5)
        tk.Entry(frame_top, textvariable=self.file1_path, width=95).grid(row=0, column=1, padx=5, pady=5, sticky="we")
        tk.Button(frame_top, text="Vælg fil", command=self.choose_file1).grid(row=0, column=2, padx=5, pady=5)

        tk.Label(frame_top, text="Target-fil:").grid(row=1, column=0, sticky="w", pady=5)
        tk.Entry(frame_top, textvariable=self.file2_path, width=95).grid(row=1, column=1, padx=5, pady=5, sticky="we")
        tk.Button(frame_top, text="Vælg fil", command=self.choose_file2).grid(row=1, column=2, padx=5, pady=5)

        tk.Label(frame_top, text="Klynge:").grid(row=2, column=0, sticky="w", pady=5)
        tk.Entry(frame_top, textvariable=self.cluster_id_var, width=30).grid(row=2, column=1, padx=5, pady=5, sticky="w")

        tk.Label(frame_top, text="Primærnøgle-kolonne:").grid(row=3, column=0, sticky="w", pady=5)
        self.key_entry = tk.Entry(frame_top, textvariable=self.key_column_var, width=30)
        self.key_entry.grid(row=3, column=1, padx=5, pady=5, sticky="w")

        tk.Label(frame_top, text="Scriptversion:").grid(row=4, column=0, sticky="w", pady=5)
        self.version_label = tk.Label(frame_top, text=SCRIPT_VERSION, anchor="w", width=20, relief="sunken", bg="#f0f0f0")
        self.version_label.grid(row=4, column=1, padx=5, pady=5, sticky="w")

        frame_buttons = tk.Frame(self.root, padx=10, pady=5)
        frame_buttons.pack(fill="x")

        tk.Button(frame_buttons, text="Sammenlign filer", command=self.run_comparison, height=2).pack(side="left")
        tk.Button(frame_buttons, text="Ryd", command=self.clear_output, height=2).pack(side="left", padx=10)
        tk.Button(frame_buttons, text="Download rapport", command=self.download_report, height=2).pack(side="left", padx=10)
        tk.Button(frame_buttons, text="Eksporter til sammenlign.xlsx", command=self.export_to_csv, height=2).pack(side="left", padx=10)

        self.status_label = tk.Label(
            frame_buttons,
            text="Afventer",
            bg="light gray",
            fg="black",
            font=("Arial", 12, "bold"),
            width=12,
            padx=10,
            pady=5
        )
        self.status_label.pack(side="left", padx=15)

        frame_output = tk.Frame(self.root, padx=10, pady=10)
        frame_output.pack(fill="both", expand=True)

        tk.Label(frame_output, text="Rapport:").pack(anchor="w")

        self.output_box = scrolledtext.ScrolledText(
            frame_output,
            wrap=tk.WORD,
            font=("Consolas", 10)
        )
        self.output_box.pack(fill="both", expand=True, pady=5)

    def choose_file1(self):
        file_path = filedialog.askopenfilename(
            title="Vælg legacy-fil",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if file_path:
            self.file1_path.set(file_path)
            self.auto_set_default_key_column()

    def choose_file2(self):
        file_path = filedialog.askopenfilename(
            title="Vælg target-fil",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if file_path:
            self.file2_path.set(file_path)
            self.auto_set_default_key_column()

    def auto_set_default_key_column(self):
        current_value = self.key_column_var.get().strip()
        if current_value:
            return

        file_candidates = [self.file1_path.get().strip(), self.file2_path.get().strip()]

        for file_path in file_candidates:
            if not file_path:
                continue
            try:
                df = load_csv_normalized(file_path)
                if DEFAULT_KEY_COLUMN in df.columns:
                    self.key_column_var.set(DEFAULT_KEY_COLUMN)
                    return
            except Exception:
                pass

    def run_comparison(self):
        file1 = self.file1_path.get().strip()
        file2 = self.file2_path.get().strip()

        if not file1 or not file2:
            messagebox.showwarning("Manglende filer", "Vælg venligst både legacy-fil og target-fil.")
            return

        try:
            report_text, is_identical, summary_row = build_report(
                legacy_file=file1,
                target_file=file2,
                cluster_id=self.cluster_id_var.get().strip(),
                key_column=self.key_column_var.get().strip(),
                script_version=SCRIPT_VERSION
            )

            self.last_report_text = report_text
            self.last_summary_row = summary_row

            self.output_box.delete("1.0", tk.END)
            self.output_box.insert(tk.END, report_text)

            if is_identical:
                self.status_label.config(text="Ens", bg="green", fg="white")
            else:
                self.status_label.config(text="Forskel", bg="red", fg="white")

        except Exception as e:
            self.status_label.config(text="Fejl", bg="orange", fg="black")
            messagebox.showerror("Fejl", f"Der opstod en fejl under sammenligningen:\n\n{e}")

    def export_to_csv(self):
        if self.last_summary_row is None:
            messagebox.showwarning("Ingen data", "Kør en sammenligning først, før du eksporterer.")
            return

        try:
            file_existed = SAMMENLIGN_EXCEL.exists()
            append_to_sammenlign_excel(self.last_summary_row)

            if file_existed:
                msg = f"Række tilføjet til eksisterende fil:\n\n{SAMMENLIGN_EXCEL.resolve()}"
            else:
                msg = f"Ny fil oprettet med første række:\n\n{SAMMENLIGN_EXCEL.resolve()}"

            messagebox.showinfo("Eksport gennemført", msg)

        except Exception as e:
            messagebox.showerror("Eksportfejl", f"Kunne ikke skrive til sammenlign.csv:\n\n{e}")

    def download_report(self):
        if not self.last_report_text.strip():
            messagebox.showwarning("Ingen rapport", "Der er ingen rapport at gemme endnu.")
            return

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_name = f"sammenligningsrapport_{timestamp}.txt"

        file_path = filedialog.asksaveasfilename(
            title="Gem rapport som",
            defaultextension=".txt",
            initialfile=default_name,
            filetypes=[("Tekstfiler", "*.txt"), ("Alle filer", "*.*")]
        )

        if not file_path:
            return

        try:
            Path(file_path).write_text(self.last_report_text, encoding="utf-8")
            messagebox.showinfo("Rapport gemt", f"Rapporten er gemt her:\n\n{file_path}")
        except Exception as e:
            messagebox.showerror("Gemmefejl", f"Kunne ikke gemme rapporten:\n\n{e}")

    def clear_output(self):
        self.output_box.delete("1.0", tk.END)
        self.last_report_text = ""
        self.last_summary_row = None
        self.status_label.config(text="Afventer", bg="light gray", fg="black")


if __name__ == "__main__":
    root = tk.Tk()
    app = CsvCompareApp(root)
    root.mainloop()

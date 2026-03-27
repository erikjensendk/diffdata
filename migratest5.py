import json
import socket
import time
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

import oracledb
import openpyxl
from simple_salesforce import Salesforce

SCRIPT_VERSION = "5.0"
JOIN_KEY = "LegacyId_m__c"
CONFIG_FILE = Path("config.json")
SAMMENLIGN_EXCEL = Path("sammenlign.xlsx")

SAMMENLIGN_COLUMNS = [
    "tidsstempel",
    "klynge",
    "oracle_view",
    "salesforce_objekt",
    "primaernoegle_kolonne",
    "raekker_oracle",
    "raekker_salesforce",
    "kolonner_sammenlignet",
    "celler_sammenlignet",
    "afvigelser",
    "status",
    "eksekveringstid_ms",
    "maskine",
    "script_version",
]

DEFAULT_CONFIG = {
    "oracle": {
        "host": "",
        "port": 1521,
        "service_name": "",
        "username": "",
        "password": ""
    },
    "salesforce": {
        "username": "",
        "password": "",
        "security_token": "",
        "domain": "test"
    }
}


# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

def load_config():
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return json.loads(json.dumps(DEFAULT_CONFIG))


def save_config(cfg):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2, ensure_ascii=False)


# ---------------------------------------------------------------------------
# Decimal helpers (same logic as v4)
# ---------------------------------------------------------------------------

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


def identify_numeric_columns(rows, columns):
    numeric_columns = []
    for col in columns:
        non_empty = [str(row.get(col, "")).strip() for row in rows if str(row.get(col, "")).strip() != ""]
        if not non_empty:
            continue
        parsed = [try_parse_decimal(v) for v in non_empty]
        if all(p is not None for p in parsed):
            numeric_columns.append(col)
    return numeric_columns


def calculate_numeric_sums(rows, numeric_columns):
    result = {}
    for col in numeric_columns:
        total = Decimal("0")
        for row in rows:
            parsed = try_parse_decimal(row.get(col))
            if parsed is not None:
                total += parsed
        result[col] = total
    return result


# ---------------------------------------------------------------------------
# Oracle
# ---------------------------------------------------------------------------

def oracle_connect(cfg):
    o = cfg["oracle"]
    dsn = f"{o['host']}:{o['port']}/{o['service_name']}"
    return oracledb.connect(user=o["username"], password=o["password"], dsn=dsn)


def oracle_get_views(conn):
    cursor = conn.cursor()
    cursor.execute("SELECT view_name FROM user_views ORDER BY view_name")
    return [row[0] for row in cursor.fetchall()]


def oracle_fetch_rows(conn, view_name):
    cursor = conn.cursor()
    cursor.execute(f"SELECT * FROM {view_name}")  # noqa: S608 — view_name from dropdown, not user text input
    columns = [col[0] for col in cursor.description]
    rows = [dict(zip(columns, row)) for row in cursor.fetchall()]
    # Normalize all values to string
    for row in rows:
        for k in row:
            row[k] = "" if row[k] is None else str(row[k]).strip()
    return columns, rows


# ---------------------------------------------------------------------------
# Salesforce
# ---------------------------------------------------------------------------

def sf_connect(cfg):
    s = cfg["salesforce"]
    kwargs = dict(
        username=s["username"],
        password=s["password"],
        security_token=s["security_token"],
    )
    if s.get("domain"):
        kwargs["domain"] = s["domain"]
    return Salesforce(**kwargs)


def sf_fetch_rows(sf, sf_object, oracle_fields):
    fields = ", ".join(oracle_fields)
    query = f"SELECT {fields} FROM {sf_object}"
    result = sf.query_all(query)
    rows = []
    for r in result["records"]:
        cleaned = {k: ("" if v is None else str(v).strip()) for k, v in r.items() if k != "attributes"}
        rows.append(cleaned)
    return rows


# ---------------------------------------------------------------------------
# Comparison
# ---------------------------------------------------------------------------

def build_row_map(rows, key_column):
    row_map = {}
    for idx, row in enumerate(rows):
        key_value = str(row.get(key_column, "")).strip()
        if key_value == "":
            key_value = f"ROW_{idx + 1}"
        original_key = key_value
        suffix = 2
        while key_value in row_map:
            key_value = f"{original_key}__{suffix}"
            suffix += 1
        row_map[key_value] = row
    return row_map


def compare_rows(oracle_rows, sf_rows, oracle_columns):
    oracle_map = build_row_map(oracle_rows, JOIN_KEY)
    sf_map = build_row_map(sf_rows, JOIN_KEY)

    all_row_ids = sorted(set(oracle_map.keys()) | set(sf_map.keys()))
    differences = []
    compared_cell_count = 0

    for row_id in all_row_ids:
        o_row = oracle_map.get(row_id)
        s_row = sf_map.get(row_id)

        if o_row is None:
            differences.append({
                "column": "<hele rækken>",
                "row_id": row_id,
                "oracle_value": "<mangler>",
                "sf_value": "Række findes kun i Salesforce"
            })
            continue

        if s_row is None:
            differences.append({
                "column": "<hele rækken>",
                "row_id": row_id,
                "oracle_value": "Række findes kun i Oracle",
                "sf_value": "<mangler>"
            })
            continue

        for col in oracle_columns:
            compared_cell_count += 1
            o_val = str(o_row.get(col, ""))
            s_val = str(s_row.get(col, ""))
            if o_val != s_val:
                differences.append({
                    "column": col,
                    "row_id": row_id,
                    "oracle_value": o_val,
                    "sf_value": s_val
                })

    return differences, compared_cell_count


# ---------------------------------------------------------------------------
# Report
# ---------------------------------------------------------------------------

def build_report(oracle_view, sf_object, cluster_id, cfg):
    start = time.perf_counter()

    conn = oracle_connect(cfg)
    oracle_columns, oracle_rows = oracle_fetch_rows(conn, oracle_view)
    conn.close()

    sf = sf_connect(cfg)
    sf_rows = sf_fetch_rows(sf, sf_object, oracle_columns)

    numeric_cols = identify_numeric_columns(oracle_rows, oracle_columns)
    sums_oracle = calculate_numeric_sums(oracle_rows, numeric_cols)
    sums_sf = calculate_numeric_sums(sf_rows, numeric_cols)

    differences, compared_cell_count = compare_rows(oracle_rows, sf_rows, oracle_columns)

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
    lines.append(f"Oracle view: {oracle_view}")
    lines.append(f"Salesforce objekt: {sf_object}")
    lines.append(f"Dato for sammenligning: {timestamp}")
    lines.append(f"Diff-script version: {SCRIPT_VERSION}")
    lines.append("")

    lines.append("----------------------------------------")
    lines.append("1. Rækketælling")
    lines.append("----------------------------------------")
    lines.append(f"Antal rækker i Oracle:      {len(oracle_rows)}")
    lines.append(f"Antal rækker i Salesforce:  {len(sf_rows)}")
    lines.append("")

    lines.append("----------------------------------------")
    lines.append("2. Numeriske kolonne-summer")
    lines.append("----------------------------------------")
    if numeric_cols:
        for col in numeric_cols:
            lines.append(f"Kolonne: {col}")
            lines.append(f"  SUM Oracle:      {format_decimal(sums_oracle.get(col))}")
            lines.append(f"  SUM Salesforce:  {format_decimal(sums_sf.get(col))}")
            lines.append("")
    else:
        lines.append("Ingen numeriske kolonner fundet.")
        lines.append("")

    lines.append("----------------------------------------")
    lines.append("3. Kolonne- og cellematch")
    lines.append("----------------------------------------")
    lines.append(f"Antal kolonner sammenlignet:    {len(oracle_columns)}")
    lines.append(f"Antal celler sammenlignet:      {compared_cell_count}")
    lines.append(f"Antal afvigelser fundet:        {len(differences)}")
    lines.append("")

    lines.append("----------------------------------------")
    lines.append("4. Resultat")
    lines.append("----------------------------------------")
    lines.append(f"STATUS: {status}")
    lines.append("")

    for diff in differences:
        lines.append(f"- Kolonne: {diff['column']}")
        lines.append(f"- Række-ID / primærnøgle: {diff['row_id']}")
        lines.append(f"- Oracle værdi:      {diff['oracle_value']}")
        lines.append(f"- Salesforce værdi:  {diff['sf_value']}")
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
        "oracle_view": oracle_view,
        "salesforce_objekt": sf_object,
        "primaernoegle_kolonne": JOIN_KEY,
        "raekker_oracle": len(oracle_rows),
        "raekker_salesforce": len(sf_rows),
        "kolonner_sammenlignet": len(oracle_columns),
        "celler_sammenlignet": compared_cell_count,
        "afvigelser": len(differences),
        "status": status,
        "eksekveringstid_ms": runtime_ms,
        "maskine": hostname,
        "script_version": SCRIPT_VERSION,
    }

    return report_text, status == "IDENTISKE", summary_row


# ---------------------------------------------------------------------------
# Excel export
# ---------------------------------------------------------------------------

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


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Oracle–Salesforce Datasammenligning v5.0")
        self.root.geometry("1150x820")

        self.cfg = load_config()
        self.last_report_text = ""
        self.last_summary_row = None

        self._build_ui()
        self._load_config_into_ui()

    def _build_ui(self):
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill="both", expand=True, padx=10, pady=10)

        self.tab_compare = tk.Frame(notebook, padx=10, pady=10)
        self.tab_conn = tk.Frame(notebook, padx=10, pady=10)

        notebook.add(self.tab_compare, text="Sammenligning")
        notebook.add(self.tab_conn, text="Forbindelser")

        self._build_compare_tab()
        self._build_conn_tab()

    # --- Sammenligning-fane ---

    def _build_compare_tab(self):
        frame = self.tab_compare

        tk.Label(frame, text="Oracle view:").grid(row=0, column=0, sticky="w", pady=5)
        self.view_var = tk.StringVar()
        self.view_dropdown = ttk.Combobox(frame, textvariable=self.view_var, width=40, state="readonly")
        self.view_dropdown.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        tk.Button(frame, text="Hent views", command=self.fetch_views).grid(row=0, column=2, padx=5, pady=5)

        tk.Label(frame, text="Salesforce objekt:").grid(row=1, column=0, sticky="w", pady=5)
        self.sf_object_var = tk.StringVar()
        tk.Entry(frame, textvariable=self.sf_object_var, width=42).grid(row=1, column=1, padx=5, pady=5, sticky="w")

        tk.Label(frame, text="Klynge:").grid(row=2, column=0, sticky="w", pady=5)
        self.cluster_var = tk.StringVar()
        tk.Entry(frame, textvariable=self.cluster_var, width=30).grid(row=2, column=1, padx=5, pady=5, sticky="w")

        frame_buttons = tk.Frame(frame)
        frame_buttons.grid(row=3, column=0, columnspan=3, sticky="w", pady=10)

        tk.Button(frame_buttons, text="Sammenlign", command=self.run_comparison, height=2).pack(side="left")
        tk.Button(frame_buttons, text="Ryd", command=self.clear_output, height=2).pack(side="left", padx=10)
        tk.Button(frame_buttons, text="Download rapport", command=self.download_report, height=2).pack(side="left", padx=10)
        tk.Button(frame_buttons, text="Eksporter til sammenlign.xlsx", command=self.export_to_excel, height=2).pack(side="left", padx=10)

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

        tk.Label(frame, text="Rapport:").grid(row=4, column=0, sticky="w")
        self.output_box = scrolledtext.ScrolledText(frame, wrap=tk.WORD, font=("Consolas", 10))
        self.output_box.grid(row=5, column=0, columnspan=3, sticky="nsew", pady=5)

        frame.rowconfigure(5, weight=1)
        frame.columnconfigure(1, weight=1)

    # --- Forbindelser-fane ---

    def _build_conn_tab(self):
        frame = self.tab_conn
        row = 0

        # Oracle
        tk.Label(frame, text="Oracle", font=("Arial", 11, "bold")).grid(row=row, column=0, sticky="w", pady=(5, 2))
        row += 1

        fields_oracle = [("Host:", "o_host"), ("Port:", "o_port"), ("Service name:", "o_service"),
                         ("Username:", "o_user"), ("Password:", "o_pass")]
        self.o_host = tk.StringVar()
        self.o_port = tk.StringVar()
        self.o_service = tk.StringVar()
        self.o_user = tk.StringVar()
        self.o_pass = tk.StringVar()

        for label, attr in fields_oracle:
            tk.Label(frame, text=label).grid(row=row, column=0, sticky="w", pady=3)
            show = "*" if attr == "o_pass" else ""
            tk.Entry(frame, textvariable=getattr(self, attr), width=40, show=show).grid(row=row, column=1, padx=5, pady=3, sticky="w")
            row += 1

        tk.Button(frame, text="Test Oracle-forbindelse", command=self.test_oracle).grid(row=row, column=1, sticky="w", pady=5)
        row += 1

        ttk.Separator(frame, orient="horizontal").grid(row=row, column=0, columnspan=3, sticky="ew", pady=10)
        row += 1

        # Salesforce
        tk.Label(frame, text="Salesforce", font=("Arial", 11, "bold")).grid(row=row, column=0, sticky="w", pady=(5, 2))
        row += 1

        self.sf_user = tk.StringVar()
        self.sf_pass = tk.StringVar()
        self.sf_token = tk.StringVar()
        self.sf_domain = tk.StringVar()

        fields_sf = [("Username:", "sf_user"), ("Password:", "sf_pass"),
                     ("Security token:", "sf_token"), ("Domain:", "sf_domain")]

        for label, attr in fields_sf:
            tk.Label(frame, text=label).grid(row=row, column=0, sticky="w", pady=3)
            show = "*" if attr == "sf_pass" else ""
            tk.Entry(frame, textvariable=getattr(self, attr), width=40, show=show).grid(row=row, column=1, padx=5, pady=3, sticky="w")
            row += 1

        tk.Button(frame, text="Test Salesforce-forbindelse", command=self.test_salesforce).grid(row=row, column=1, sticky="w", pady=5)
        row += 1

        ttk.Separator(frame, orient="horizontal").grid(row=row, column=0, columnspan=3, sticky="ew", pady=10)
        row += 1

        tk.Button(frame, text="Gem forbindelser", command=self.save_connections, height=2).grid(row=row, column=1, sticky="w", pady=5)

    # --- Config helpers ---

    def _load_config_into_ui(self):
        o = self.cfg.get("oracle", {})
        self.o_host.set(o.get("host", ""))
        self.o_port.set(str(o.get("port", 1521)))
        self.o_service.set(o.get("service_name", ""))
        self.o_user.set(o.get("username", ""))
        self.o_pass.set(o.get("password", ""))

        s = self.cfg.get("salesforce", {})
        self.sf_user.set(s.get("username", ""))
        self.sf_pass.set(s.get("password", ""))
        self.sf_token.set(s.get("security_token", ""))
        self.sf_domain.set(s.get("domain", "test"))

    def _build_cfg_from_ui(self):
        return {
            "oracle": {
                "host": self.o_host.get().strip(),
                "port": int(self.o_port.get().strip() or 1521),
                "service_name": self.o_service.get().strip(),
                "username": self.o_user.get().strip(),
                "password": self.o_pass.get(),
            },
            "salesforce": {
                "username": self.sf_user.get().strip(),
                "password": self.sf_pass.get(),
                "security_token": self.sf_token.get().strip(),
                "domain": self.sf_domain.get().strip(),
            }
        }

    # --- Actions ---

    def save_connections(self):
        self.cfg = self._build_cfg_from_ui()
        save_config(self.cfg)
        messagebox.showinfo("Gemt", f"Forbindelsesoplysninger gemt i {CONFIG_FILE.resolve()}")

    def test_oracle(self):
        cfg = self._build_cfg_from_ui()
        try:
            conn = oracle_connect(cfg)
            conn.close()
            messagebox.showinfo("Oracle", "Forbindelsen lykkedes.")
        except Exception as e:
            messagebox.showerror("Oracle fejl", str(e))

    def test_salesforce(self):
        cfg = self._build_cfg_from_ui()
        try:
            sf_connect(cfg)
            messagebox.showinfo("Salesforce", "Forbindelsen lykkedes.")
        except Exception as e:
            messagebox.showerror("Salesforce fejl", str(e))

    def fetch_views(self):
        cfg = self._build_cfg_from_ui()
        try:
            conn = oracle_connect(cfg)
            views = oracle_get_views(conn)
            conn.close()
            self.view_dropdown["values"] = views
            if views:
                self.view_dropdown.current(0)
            messagebox.showinfo("Views hentet", f"{len(views)} view(s) fundet.")
        except Exception as e:
            messagebox.showerror("Fejl", f"Kunne ikke hente views:\n\n{e}")

    def run_comparison(self):
        oracle_view = self.view_var.get().strip()
        sf_object = self.sf_object_var.get().strip()

        if not oracle_view or not sf_object:
            messagebox.showwarning("Manglende input", "Vælg et Oracle view og angiv et Salesforce objekt.")
            return

        self.cfg = self._build_cfg_from_ui()

        try:
            report_text, is_identical, summary_row = build_report(
                oracle_view=oracle_view,
                sf_object=sf_object,
                cluster_id=self.cluster_var.get().strip(),
                cfg=self.cfg,
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

    def export_to_excel(self):
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
            messagebox.showerror("Eksportfejl", f"Kunne ikke skrive til sammenlign.xlsx:\n\n{e}")

    def download_report(self):
        if not self.last_report_text.strip():
            messagebox.showwarning("Ingen rapport", "Der er ingen rapport at gemme endnu.")
            return
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        file_path = filedialog.asksaveasfilename(
            title="Gem rapport som",
            defaultextension=".txt",
            initialfile=f"sammenligningsrapport_{timestamp}.txt",
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
    app = App(root)
    root.mainloop()

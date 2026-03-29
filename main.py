# -*- coding: utf-8 -*-
"""
DataCopilot - Offline Desktop Excel AI Assistant / 离线桌面 Excel AI 助手
"""
import atexit
import os
import subprocess
import sys
import threading
import time
from datetime import datetime
from tkinter import filedialog, messagebox, scrolledtext
import tkinter as tk

import duckdb
import pandas as pd
import requests


# ─── Path Resolution ──────────────────────────────────────────────────────────
# Works both in dev (plain Python) and after PyInstaller packaging

def _base_dir():
    if hasattr(sys, "_MEIPASS"):
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))


BASE = _base_dir()
LLAMA_EXE = os.path.join(BASE, "engine", "llama-server.exe")
MODEL_FILE = os.path.join(BASE, "model", "qwen2.5-coder-1.5b-instruct-q4_k_m.gguf")
SERVER_BASE = "http://127.0.0.1:8080"
CHAT_URL = f"{SERVER_BASE}/v1/chat/completions"
HEALTH_URL = f"{SERVER_BASE}/health"


# ─── Process Guardian ─────────────────────────────────────────────────────────

_server_proc = None


def start_server():
    global _server_proc
    cmd = [
        LLAMA_EXE,
        "-m", MODEL_FILE,
        "--port", "8080",
        "--host", "127.0.0.1",
        "--threads", "4",
        "-c", "1024",
        "--batch-size", "128",
        "-np", "1",
    ]
    _server_proc = subprocess.Popen(
        cmd,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        creationflags=subprocess.CREATE_NO_WINDOW,
    )
    atexit.register(stop_server)


def stop_server():
    global _server_proc
    if _server_proc and _server_proc.poll() is None:
        _server_proc.terminate()
        try:
            _server_proc.wait(timeout=5)
        except subprocess.TimeoutExpired:
            _server_proc.kill()


def wait_for_server(timeout=90):
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            r = requests.get(HEALTH_URL, timeout=2)
            if r.status_code == 200:
                return True
        except requests.exceptions.RequestException:
            pass
        time.sleep(1)
    return False


# ─── Data Pipeline ────────────────────────────────────────────────────────────

def load_file(path: str) -> pd.DataFrame:
    if path.lower().endswith(".csv"):
        return pd.read_csv(path)
    return pd.read_excel(path, engine="openpyxl")


def schema_description(df: pd.DataFrame) -> str:
    parts = [f'"{col}" ({dtype})' for col, dtype in zip(df.columns, df.dtypes)]
    return ", ".join(parts)


def run_sql(sql: str, df: pd.DataFrame) -> pd.DataFrame:
    conn = duckdb.connect()
    conn.register("df", df)
    result = conn.execute(sql).df()
    conn.close()
    return result


def save_result(result_df: pd.DataFrame, source_path: str) -> str:
    directory = os.path.dirname(source_path)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out = os.path.join(directory, f"Result_{timestamp}.xlsx")
    result_df.to_excel(out, index=False)
    return out


# ─── LLM Router & Self-Correction ────────────────────────────────────────────

def _system_prompt(df: pd.DataFrame) -> str:
    col_list = ", ".join(f'"{c}"' for c in df.columns)
    schema = schema_description(df)
    sample = df.head(3).to_string(index=False)
    return (
        "You are a data analysis assistant. Convert the user's natural language instructions "
        "into DuckDB SQL queries. The user may write in Chinese or English.\n\n"
        "[STRICT RULES - never break these]\n"
        f"1. The table name is always 'df'. Available columns: {col_list}\n"
        f"2. Column data types: {schema}\n"
        f"3. Sample data (first 3 rows) - study this to understand what each column actually contains:\n{sample}\n"
        "4. IMPORTANT: Match the user's search value to the correct column based on its language "
        "and content. E.g. if the user searches a Chinese word, look for it in the column that "
        "contains Chinese values, not the English column.\n"
        "5. Output ONLY a single valid DuckDB SQL statement. "
        "No explanations, no Markdown code fences, no backticks.\n"
        "6. If the user's instruction is vague (e.g. 'analyze this', 'clean the data'), "
        "reply starting with CLARIFY: and ask a specific question.\n"
        "7. Charts, colors, fonts, cell formatting, and cross-file operations are NOT supported. "
        "If asked, reply: CLARIFY: Sorry, I only support data filtering, calculation and aggregation.\n"
        "8. Only one table 'df' is available. No cross-file joins."
    )


def _call_llm(messages: list) -> str:
    payload = {
        "model": "local",
        "messages": messages,
        "temperature": 0.1,
        "max_tokens": 300,
        "stop": ["\n\n", "```"],
    }
    resp = requests.post(CHAT_URL, json=payload, timeout=90)
    resp.raise_for_status()
    return resp.json()["choices"][0]["message"]["content"].strip()


def _clean_sql(text: str) -> str:
    text = text.strip()
    if text.startswith("```"):
        lines = [l for l in text.splitlines() if not l.strip().startswith("```")]
        text = "\n".join(lines).strip()
    return text


SQL_STARTERS = ("SELECT", "WITH", "INSERT", "UPDATE", "DELETE", "CREATE", "DROP", "ALTER")


def process_query(user_query: str, df: pd.DataFrame):
    """
    Returns one of:
      ("result",  result_df, sql_string)
      ("clarify", message_string)
      ("error",   error_string)
    """
    messages = [
        {"role": "system", "content": _system_prompt(df)},
        {"role": "user",   "content": user_query},
    ]

    for attempt in range(3):
        raw = _call_llm(messages)

        if raw.upper().startswith("CLARIFY:"):
            return ("clarify", raw[len("CLARIFY:"):].strip())

        sql = _clean_sql(raw)

        if not sql.upper().startswith(SQL_STARTERS):
            return ("clarify", raw)

        try:
            result_df = run_sql(sql, df)
            return ("result", result_df, sql)
        except Exception as exc:
            if attempt < 2:
                messages.append({"role": "assistant", "content": raw})
                messages.append({
                    "role": "user",
                    "content": (
                        f"The SQL above failed with error: {exc}\n"
                        "Please fix it and output only the corrected SQL, no explanation."
                    ),
                })
            else:
                return ("error", f"SQL execution failed after 2 retries.\nError: {exc}\nLast SQL: {sql}")

    return ("error", "Could not generate a valid SQL query. Please rephrase your request.")


# ─── UI ───────────────────────────────────────────────────────────────────────

BG        = "#1e1e2e"
BG_PANEL  = "#313244"
BG_INPUT  = "#45475a"
BG_CHAT   = "#181825"
FG_MAIN   = "#cdd6f4"
FG_DIM    = "#6c7086"
C_BLUE    = "#89b4fa"
C_GREEN   = "#a6e3a1"
C_RED     = "#f38ba8"
C_ORANGE  = "#fab387"


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("DataCopilot - Offline Excel AI Assistant")
        self.root.geometry("860x640")
        self.root.minsize(640, 480)
        self.root.configure(bg=BG)

        self.df = None
        self.file_path = None

        self._build_ui()
        self._launch_server()

    # ── UI Construction ───────────────────────────────────────────────────────

    def _build_ui(self):
        # Top bar
        top = tk.Frame(self.root, bg=BG_PANEL, pady=8, padx=14)
        top.pack(fill="x")

        tk.Label(top, text="DataCopilot", font=("Segoe UI", 13, "bold"),
                 bg=BG_PANEL, fg=FG_MAIN).pack(side="left")

        self.status_lbl = tk.Label(top, text="Loading AI model...",
                                   font=("Segoe UI", 9), bg=BG_PANEL, fg=C_RED)
        self.status_lbl.pack(side="right")

        # File bar
        file_bar = tk.Frame(self.root, bg=BG, pady=8, padx=14)
        file_bar.pack(fill="x")

        self.open_btn = tk.Button(
            file_bar, text="Open File / \u6253\u5f00\u6587\u4ef6", command=self._open_file,
            bg=C_BLUE, fg=BG, font=("Segoe UI", 10, "bold"),
            relief="flat", padx=12, pady=4, cursor="hand2", state="disabled",
        )
        self.open_btn.pack(side="left")

        self.file_lbl = tk.Label(file_bar, text="No file selected / \u672a\u9009\u62e9\u6587\u4ef6",
                                 font=("Segoe UI", 9), bg=BG, fg=FG_DIM)
        self.file_lbl.pack(side="left", padx=10)

        # Chat area
        chat_wrap = tk.Frame(self.root, bg=BG, padx=14, pady=0)
        chat_wrap.pack(fill="both", expand=True)

        self.chat = scrolledtext.ScrolledText(
            chat_wrap, state="disabled", wrap=tk.WORD,
            bg=BG_CHAT, fg=FG_MAIN, font=("Consolas", 10),
            relief="flat", padx=10, pady=10, insertbackground=FG_MAIN,
        )
        self.chat.pack(fill="both", expand=True)

        self.chat.tag_config("user",   foreground=C_BLUE,   font=("Segoe UI", 10, "bold"))
        self.chat.tag_config("ai",     foreground=C_GREEN,  font=("Segoe UI", 10))
        self.chat.tag_config("sql",    foreground=C_ORANGE, font=("Consolas", 9))
        self.chat.tag_config("system", foreground=FG_DIM,   font=("Segoe UI", 9, "italic"))
        self.chat.tag_config("error",  foreground=C_RED,    font=("Segoe UI", 10))

        # Input bar
        input_bar = tk.Frame(self.root, bg=BG_PANEL, pady=10, padx=14)
        input_bar.pack(fill="x", side="bottom")

        self.input_box = tk.Text(
            input_bar, height=3, bg=BG_INPUT, fg=FG_MAIN,
            font=("Segoe UI", 10), relief="flat", padx=8, pady=6,
            insertbackground=FG_MAIN, wrap=tk.WORD,
        )
        self.input_box.pack(side="left", fill="x", expand=True)
        self.input_box.bind("<Return>", self._on_enter)

        self.send_btn = tk.Button(
            input_bar, text="Send / \u53d1\u9001", command=self._send,
            bg=C_GREEN, fg=BG, font=("Segoe UI", 10, "bold"),
            relief="flat", padx=16, pady=6, cursor="hand2", state="disabled",
        )
        self.send_btn.pack(side="right", padx=(10, 0))

    # ── Server Lifecycle ──────────────────────────────────────────────────────

    def _launch_server(self):
        def _run():
            try:
                start_server()
                ok = wait_for_server(timeout=90)
            except Exception as exc:
                self.root.after(0, lambda: self._on_server_fail(str(exc)))
                return
            if ok:
                self.root.after(0, self._on_server_ready)
            else:
                self.root.after(0, lambda: self._on_server_fail("Startup timed out"))

        threading.Thread(target=_run, daemon=True).start()

    def _on_server_ready(self):
        self.status_lbl.config(text="AI Ready", fg=C_GREEN)
        self.open_btn.config(state="normal")
        self._log("System", "AI model loaded. Click 'Open File' to load an Excel or CSV file.", "system")

    def _on_server_fail(self, reason: str):
        self.status_lbl.config(text="Model failed to start", fg=C_RED)
        self._log("System", f"AI startup failed: {reason}\nCheck that engine/llama-server.exe and model/ exist.", "error")

    # ── File Handling ─────────────────────────────────────────────────────────

    def _open_file(self):
        path = filedialog.askopenfilename(
            title="Select data file",
            filetypes=[("Excel / CSV", "*.xlsx *.xls *.csv"), ("All files", "*.*")],
        )
        if not path:
            return
        try:
            df = load_file(path)
        except Exception as exc:
            messagebox.showerror("File load error", str(exc))
            return

        self.df = df
        self.file_path = path
        fname = os.path.basename(path)
        rows, cols = df.shape
        self.file_lbl.config(text=fname, fg=FG_MAIN)
        self._log(
            "System",
            f"Loaded: {fname}\n{rows} rows x {cols} columns\nColumns: {schema_description(df)}",
            "system",
        )
        self.send_btn.config(state="normal")
        self.input_box.focus()

    # ── Messaging ─────────────────────────────────────────────────────────────

    def _on_enter(self, event):
        if not (event.state & 0x1):  # Shift not held -> send
            self._send()
            return "break"

    def _send(self):
        if self.df is None:
            return
        query = self.input_box.get("1.0", "end-1c").strip()
        if not query:
            return

        self.input_box.delete("1.0", tk.END)
        self._log("You", query, "user")
        self._set_input_state(False)

        threading.Thread(target=self._query_worker, args=(query,), daemon=True).start()

    def _query_worker(self, query: str):
        try:
            result = process_query(query, self.df)
        except Exception as exc:
            result = ("error", str(exc))
        self.root.after(0, lambda: self._handle_result(result))

    def _handle_result(self, result):
        kind = result[0]

        if kind == "clarify":
            self._log("AI", result[1], "ai")

        elif kind == "result":
            _, result_df, sql = result
            self._log("SQL", sql, "sql")
            rows, cols = result_df.shape

            if rows <= 20:
                # Small result — show directly in chat, no file needed
                self._log("AI", result_df.to_string(index=False), "ai")
            else:
                # Large result — save to Excel, show preview
                preview = result_df.head(5).to_string(index=False)
                self._log("AI", f"{rows} rows x {cols} columns. Preview (first 5):\n{preview}", "ai")
                try:
                    out = save_result(result_df, self.file_path)
                    self._log("System", f"Result saved to: {out}", "system")
                except Exception as exc:
                    self._log("System", f"Save failed: {exc}", "error")

        elif kind == "error":
            self._log("Error", result[1], "error")

        self._set_input_state(True)

    # ── Helpers ───────────────────────────────────────────────────────────────

    def _log(self, sender: str, text: str, tag: str):
        self.chat.config(state="normal")
        self.chat.insert(tk.END, f"\n[{sender}]\n", tag)
        self.chat.insert(tk.END, text + "\n", tag)
        self.chat.config(state="disabled")
        self.chat.see(tk.END)

    def _set_input_state(self, enabled: bool):
        state = "normal" if enabled else "disabled"
        self.send_btn.config(state=state)
        self.input_box.config(state=state)
        if enabled:
            self.input_box.focus()


# ─── Entry Point ──────────────────────────────────────────────────────────────

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.protocol("WM_DELETE_WINDOW", lambda: (stop_server(), root.destroy()))
    root.mainloop()

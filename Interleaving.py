"""
Stratified Question Sampler
----------------------------
Downloads a python "database" file from GitHub consisting of several
top-level variables, each a list of dicts (e.g. ChECalc, Sepa, Thermo...).

The user is prompted (via Tkinter) to:
  1. Pick which topics (variables) to sample from.
  2. Enter how many total questions they want.

The program samples as evenly as possible across the selected topics.
If the request can't be satisfied evenly, the user is asked whether to
go back and resample (change topics/number) or accept the largest
even allocation that IS achievable (an "undersampled" result).

Results are shown in two columns:
  - Left:  all sampled questions merged into a single copy-pasteable
           python variable.
  - Right: the same sampled questions, but broken out per original
           topic (for reference / verification), also copy-pasteable.
"""

import math
import random
import urllib.request
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext

DATABASE_URL = (
    "https://raw.githubusercontent.com/bananabonds/streamlit_proj/"
    "refs/heads/main/ARCLectureBank.py"
)

MERGED_VAR_NAME = "SampledQuestions"


# --------------------------------------------------------------------------
# Data fetching
# --------------------------------------------------------------------------
def fetch_database(url: str) -> dict:
    """
    Downloads the remote .py file and executes it in an isolated namespace,
    returning only the top-level variables that are lists of dicts
    (i.e. the actual question banks).
    """
    with urllib.request.urlopen(url) as response:
        raw_code = response.read().decode("utf-8")

    namespace = {}
    exec(raw_code, namespace)

    topics = {}
    for name, value in namespace.items():
        if name.startswith("__"):
            continue
        if isinstance(value, list) and all(isinstance(item, dict) for item in value):
            topics[name] = value

    if not topics:
        raise ValueError("No list-of-dict variables were found in the remote file.")

    return topics


# --------------------------------------------------------------------------
# Sampling logic
# --------------------------------------------------------------------------
class UnevenSampleError(Exception):
    """Raised when the requested number of questions can't be evenly sampled."""

    def __init__(self, requested: int, n_topics: int, min_size: int, min_topic: str):
        self.requested = requested
        self.n_topics = n_topics
        self.min_size = min_size
        self.min_topic = min_topic
        self.max_even_total = n_topics * min_size
        super().__init__(
            f"Can't sample {requested} questions evenly across {n_topics} topic(s).\n"
            f"The smallest selected topic ('{min_topic}') only has {min_size} "
            f"question(s) available, so the largest evenly-sampled total is "
            f"{self.max_even_total} ({n_topics} topics x {min_size})."
        )


def compute_even_allocation(selected_sizes: dict, num_questions: int) -> dict:
    """
    Given {topic: size} for the selected topics and a target total number
    of questions, work out how many questions to pull from each topic so
    that the split is as even as possible.

    Raises UnevenSampleError if the request can't be satisfied evenly.
    """
    n = len(selected_sizes)
    if n == 0:
        raise ValueError("No topics selected.")

    if num_questions <= 0:
        raise ValueError("Number of questions must be a positive integer.")

    min_topic = min(selected_sizes, key=selected_sizes.get)
    min_size = selected_sizes[min_topic]

    required_per_topic = math.ceil(num_questions / n)
    if required_per_topic > min_size:
        raise UnevenSampleError(num_questions, n, min_size, min_topic)

    base = num_questions // n
    remainder = num_questions % n

    topic_names = list(selected_sizes.keys())
    bonus_topics = set(random.sample(topic_names, remainder)) if remainder else set()

    allocation = {name: base + (1 if name in bonus_topics else 0) for name in topic_names}
    return allocation


def max_even_allocation(selected_sizes: dict) -> dict:
    """Largest possible even allocation: every topic contributes its min size."""
    min_size = min(selected_sizes.values())
    return {name: min_size for name in selected_sizes}


def stratified_sample(full_database: dict, allocation: dict) -> dict:
    """Draws the actual random samples for each topic per the allocation."""
    sampled = {}
    for topic, count in allocation.items():
        sampled[topic] = random.sample(full_database[topic], count)
    return sampled


# --------------------------------------------------------------------------
# Formatting
# --------------------------------------------------------------------------
def format_merged(sampled: dict, var_name: str = MERGED_VAR_NAME) -> str:
    """
    Merges every sampled question (regardless of original topic) into a
    single copy-pasteable python variable, e.g.:

        SampledQuestions = [
            {...},
            {...},
        ]
    """
    lines = [f"{var_name} = ["]
    for topic, items in sampled.items():
        for item in items:
            lines.append(f"    {item!r},")
    lines.append("]")
    return "\n".join(lines) + "\n"


def format_by_topic(sampled: dict) -> str:
    """
    Formats the sampled data grouped by original topic, kept in the same
    variable-per-topic style as the source database, for reference.
    """
    lines = []
    for topic, items in sampled.items():
        lines.append(f"# {topic} ({len(items)} sampled)")
        lines.append(f"{topic} = [")
        for item in items:
            lines.append(f"    {item!r},")
        lines.append("]")
        lines.append("")
    return "\n".join(lines).rstrip() + "\n"


# --------------------------------------------------------------------------
# Small reusable dialog: Resample vs Keep Undersampled
# --------------------------------------------------------------------------
def ask_resample_or_undersample(root, error: UnevenSampleError) -> bool:
    """
    Shows a modal dialog explaining the shortfall and asks the user to
    choose between going back to resample, or accepting the largest
    achievable even allocation (undersampled).

    Returns True if the user chose to keep the undersampled result,
    False if they want to resample instead.
    """
    result = {"keep_undersampled": False}

    dialog = tk.Toplevel(root)
    dialog.title("Requested amount not achievable")
    dialog.resizable(False, False)
    dialog.grab_set()
    dialog.transient(root)

    msg = (
        f"{error}\n\n"
        f"Would you like to go back and resample (change topics or the "
        f"number requested), or keep going with an undersampled result "
        f"of {error.max_even_total} question(s) total "
        f"({error.n_topics} topics x {error.min_size} each)?"
    )
    tk.Label(dialog, text=msg, wraplength=420, justify="left", padx=16, pady=16).pack()

    button_frame = ttk.Frame(dialog)
    button_frame.pack(pady=(0, 16))

    def choose_resample():
        result["keep_undersampled"] = False
        dialog.destroy()

    def choose_keep():
        result["keep_undersampled"] = True
        dialog.destroy()

    ttk.Button(button_frame, text="Resample", command=choose_resample).pack(
        side="left", padx=8
    )
    ttk.Button(
        button_frame, text="Keep Undersampled", command=choose_keep
    ).pack(side="left", padx=8)

    dialog.wait_window()
    return result["keep_undersampled"]


# --------------------------------------------------------------------------
# Tkinter UI
# --------------------------------------------------------------------------
class SamplerApp:
    def __init__(self, root, database: dict):
        self.root = root
        self.database = database
        self.topic_vars = {}

        root.title("Question Bank Sampler")
        root.geometry("420x480")

        tk.Label(
            root, text="Select the topics you want to sample from:",
            font=("Segoe UI", 11, "bold")
        ).pack(pady=(12, 4), padx=12, anchor="w")

        container = ttk.Frame(root)
        container.pack(fill="both", expand=True, padx=12)

        canvas = tk.Canvas(container, borderwidth=0, highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        checklist_frame = ttk.Frame(canvas)

        checklist_frame.bind(
            "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=checklist_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        for topic, items in sorted(self.database.items()):
            var = tk.BooleanVar(value=False)
            self.topic_vars[topic] = var
            ttk.Checkbutton(
                checklist_frame,
                text=f"{topic}  ({len(items)} questions available)",
                variable=var,
            ).pack(anchor="w", pady=2)

        entry_frame = ttk.Frame(root)
        entry_frame.pack(fill="x", padx=12, pady=(10, 4))

        tk.Label(entry_frame, text="Total number of questions to sample:").pack(anchor="w")
        self.num_questions_var = tk.StringVar()
        self.entry = ttk.Entry(entry_frame, textvariable=self.num_questions_var)
        self.entry.pack(fill="x", pady=4)

        ttk.Button(root, text="Sample Questions", command=self.on_submit).pack(pady=10)

    def get_selected_topics(self) -> dict:
        return {t: self.database[t] for t, v in self.topic_vars.items() if v.get()}

    def on_submit(self):
        selected = self.get_selected_topics()

        if not selected:
            messagebox.showerror("No topics selected", "Please select at least one topic.")
            return

        raw_value = self.num_questions_var.get().strip()
        try:
            num_questions = int(raw_value)
        except ValueError:
            messagebox.showerror("Invalid input", "Please enter a whole number of questions.")
            return

        sizes = {t: len(v) for t, v in selected.items()}

        try:
            allocation = compute_even_allocation(sizes, num_questions)
        except UnevenSampleError as e:
            keep_undersampled = ask_resample_or_undersample(self.root, e)
            if not keep_undersampled:
                # User wants to resample: leave them on the main window to adjust.
                return
            allocation = max_even_allocation(sizes)
        except ValueError as e:
            messagebox.showerror("Invalid request", str(e))
            return

        sampled = stratified_sample(self.database, allocation)
        self.show_results(sampled, allocation)

    def show_results(self, sampled: dict, allocation: dict):
        total = sum(allocation.values())

        # Console summary
        print("\nSampled question counts per topic:")
        for topic, count in allocation.items():
            print(f"  {topic}: {count}")
        print(f"Total sampled: {total}\n")

        merged_text = format_merged(sampled)
        by_topic_text = format_by_topic(sampled)

        result_win = tk.Toplevel(self.root)
        result_win.title("Sampled Questions (copy-pasteable)")
        result_win.geometry("1100x650")

        summary_text = "  |  ".join(f"{t}: {c}" for t, c in allocation.items())
        tk.Label(
            result_win,
            text=f"Sampled {total} question(s) total  —  {summary_text}",
            font=("Segoe UI", 10, "bold"),
            wraplength=1080,
            justify="left",
        ).pack(padx=10, pady=(10, 4), anchor="w")

        columns = ttk.Frame(result_win)
        columns.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        columns.columnconfigure(0, weight=1)
        columns.columnconfigure(1, weight=1)
        columns.rowconfigure(1, weight=1)

        # --- Left column: merged single variable ---
        tk.Label(
            columns, text=f"Merged ({MERGED_VAR_NAME}) — copy this",
            font=("Segoe UI", 10, "bold")
        ).grid(row=0, column=0, sticky="w", pady=(0, 4))

        merged_box = scrolledtext.ScrolledText(columns, wrap="none", font=("Consolas", 10))
        merged_box.grid(row=1, column=0, sticky="nsew", padx=(0, 5))
        merged_box.insert("1.0", merged_text)

        ttk.Button(
            columns, text="Copy Merged",
            command=lambda: self.copy_to_clipboard(merged_text, "Merged variable copied.")
        ).grid(row=2, column=0, pady=6, sticky="w")

        # --- Right column: per-topic breakdown ---
        tk.Label(
            columns, text="By Original Topic — reference",
            font=("Segoe UI", 10, "bold")
        ).grid(row=0, column=1, sticky="w", pady=(0, 4))

        by_topic_box = scrolledtext.ScrolledText(columns, wrap="none", font=("Consolas", 10))
        by_topic_box.grid(row=1, column=1, sticky="nsew", padx=(5, 0))
        by_topic_box.insert("1.0", by_topic_text)

        ttk.Button(
            columns, text="Copy By-Topic",
            command=lambda: self.copy_to_clipboard(by_topic_text, "By-topic breakdown copied.")
        ).grid(row=2, column=1, pady=6, sticky="w")

    def copy_to_clipboard(self, text: str, confirmation_message: str):
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        messagebox.showinfo("Copied", confirmation_message)


# --------------------------------------------------------------------------
# Entry point
# --------------------------------------------------------------------------
def main():
    root = tk.Tk()
    root.withdraw()  # hide until data is loaded

    try:
        database = fetch_database(DATABASE_URL)
    except Exception as e:
        messagebox.showerror("Error loading database", f"Could not load the question bank:\n{e}")
        root.destroy()
        return

    root.deiconify()
    SamplerApp(root, database)
    root.mainloop()


if __name__ == "__main__":
    main()
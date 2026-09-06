import ssl
import sys
import os
import json
import random
import re
import io
import time
import datetime
import tkinter as tk
from tkinter import messagebox, ttk, filedialog
import urllib.request
from html import escape as html_escape
from collections import Counter

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

# Pillow powers the in-app ink/sketch canvas (mirroring hand-drawn
# strokes into an off-screen raster image that can later be embedded in a
# PDF). Sketching is simply disabled if Pillow isn't installed.
try:
    from PIL import Image, ImageDraw, ImageTk
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

# reportlab compiles each question's typed explanation + sketch into
# a single downloadable "Notes & Explanations" PDF from the report screen.
try:
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.units import inch
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image as RLImage, PageBreak
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

# Fixed size (in on-screen pixels, 1:1 with the raster image used for PDF
# export) of the sketch/ink panel inside the Notes section.
NOTES_CANVAS_W = 760
NOTES_CANVAS_H = 220


def _text_to_paragraph_html(text):
    """Escapes plain text and converts newlines to <br/> for reportlab Paragraphs."""
    return html_escape(text).replace("\n", "<br/>")


database_url = "https://raw.githubusercontent.com/bananabonds/streamlit_proj/refs/heads/main/ARCLectureBank.py"

# ----------------------------------------------------------------------
# Local, writable storage (works both as a plain .py script and once
# frozen into a single .exe with PyInstaller). PyInstaller's onefile
# bundle unpacks itself into a temporary, read-only-ish directory at
# runtime, so we must NOT try to write next to sys.executable's internal
# extraction path -- instead we always resolve to the folder the actual
# .exe (or .py) lives in on disk, which stays put between runs and is
# writable on a normal Windows user account.
# ----------------------------------------------------------------------
def get_app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


DATA_DIR = os.path.join(get_app_dir(), "cble_data")
os.makedirs(DATA_DIR, exist_ok=True)

LOCAL_BANKS_FILE = os.path.join(DATA_DIR, "local_banks.json")
SAMPLE_SETS_FILE = os.path.join(DATA_DIR, "sample_sets.json")
TEST_HISTORY_FILE = os.path.join(DATA_DIR, "test_history.json")


def load_json_file(path, default):
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return default
    return default


def save_json_file(path, data):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    os.replace(tmp, path)


def tag_source(questions, list_name):
    """
    Returns a copy of `questions` where every dict is annotated with
    _source_list (which bank it came from) and _source_index (its
    position in THAT bank, 0-based -- i.e. list[0] -> reference no. 1).
    This survives shuffling, sampling, and cleaning, all the way through
    to the Excel report, so a question can always be traced back to its
    exact spot in the original reference material.
    """
    tagged = []
    for idx, q in enumerate(questions):
        q2 = dict(q)
        q2["_source_list"] = list_name
        q2["_source_index"] = idx
        tagged.append(q2)
    return tagged


# ----------------------------------------------------------------------
# Data loading
# ----------------------------------------------------------------------
# Cache so we only hit the network once per run, even though we now need
# the full topic list (for the setup dropdown) AND a specific topic's
# questions (for the exam itself).
_topic_cache = None


def fetch_all_topics():
    """
    Downloads the remote question-bank source once, execs it, and returns
    a dict of {topic_name: list_of_question_dicts} for every top-level
    variable in that file that actually looks like a question bank
    (a non-empty list of dicts each containing 'text', 'options',
    and 'correct'). Every question is tagged with its source list name
    and its original index (see tag_source).
    """
    global _topic_cache
    if _topic_cache is not None:
        return _topic_cache

    context = ssl._create_unverified_context()

    topics = {}
    try:
        with urllib.request.urlopen(database_url, context=context) as response:
            raw_code = response.read().decode("utf-8")

        namespace = {}
        exec(raw_code, namespace)

        for key, value in namespace.items():
            if key.startswith("__"):
                continue
            if isinstance(value, list) and value and all(
                isinstance(item, dict) and {"text", "options", "correct"} <= item.keys()
                for item in value
            ):
                topics[key] = tag_source(value, key)

    except Exception as e:
        print(f"An error occurred while fetching topics: {e}")
        messagebox.showerror(
            "Data Load Failed",
            f"Could not load the question bank from the server:\n{e}\n\n"
            f"Locally saved samples and local question banks (if any) are still available."
        )

    _topic_cache = topics
    return _topic_cache


def load_local_banks():
    """Returns {bank_name: tagged_question_list} for every user-created
    local bank saved in local_banks.json."""
    raw = load_json_file(LOCAL_BANKS_FILE, {})
    tagged = {}
    for name, qs in raw.items():
        tagged[name] = tag_source(qs, f"[Local] {name}")
    return tagged


def save_local_bank(bank_name, questions):
    """Persists (creates or overwrites) a local bank. `questions` should be
    plain dicts with text/options/correct (no internal _source_* tags)."""
    raw = load_json_file(LOCAL_BANKS_FILE, {})
    clean = [{"text": q["text"], "options": list(q["options"]), "correct": q["correct"]} for q in questions]
    raw[bank_name] = clean
    save_json_file(LOCAL_BANKS_FILE, raw)


def delete_local_bank(bank_name):
    raw = load_json_file(LOCAL_BANKS_FILE, {})
    if bank_name in raw:
        del raw[bank_name]
        save_json_file(LOCAL_BANKS_FILE, raw)


def get_all_lists():
    """Merges remote topics + local banks into one registry of
    {display_name: tagged_question_list} for the workbench."""
    merged = {}
    for name, qs in fetch_all_topics().items():
        merged[name] = qs
    for name, qs in load_local_banks().items():
        merged[f"[Local] {name}"] = qs
    return merged


def get_used_indices_by_list():
    """Scans every sample set ever saved and returns
    {list_name: set(of _source_index values that have appeared in a
    saved sample at least once)}. This is the basis for the 'percent of
    this bank sampled so far' tracker."""
    sample_sets = load_json_file(SAMPLE_SETS_FILE, [])
    used = {}
    for s in sample_sets:
        for q in s.get("questions", []):
            ln = q.get("_source_list")
            idx = q.get("_source_index")
            if ln is None or idx is None:
                continue
            used.setdefault(ln, set()).add(idx)
    return used


def compute_coverage(all_lists):
    """Returns {list_name: (used_count, total_count, pct)}"""
    used_by_list = get_used_indices_by_list()
    coverage = {}
    for name, qs in all_lists.items():
        total = len(qs)
        used = len(used_by_list.get(name, set()))
        pct = (used / total * 100.0) if total else 0.0
        coverage[name] = (used, total, pct)
    return coverage


def even_sample(all_lists, selected_names, total_count, prefer_unused=True):
    """
    Draws `total_count` questions spread as evenly as possible across
    every list in `selected_names`. When prefer_unused is True, questions
    that have never appeared in any previously-saved sample set are drawn
    first (per list) before falling back to already-seen ones, so repeated
    use of the workbench naturally works its way through the whole bank.
    Returns a flat, shuffled list of (still-tagged) question dicts.
    """
    selected_names = [n for n in selected_names if n in all_lists and all_lists[n]]
    n_lists = len(selected_names)
    if n_lists == 0 or total_count <= 0:
        return []

    used_by_list = get_used_indices_by_list()

    base = total_count // n_lists
    rem = total_count % n_lists
    counts = {}
    for i, name in enumerate(selected_names):
        counts[name] = base + (1 if i < rem else 0)

    result = []
    for name in selected_names:
        pool = list(all_lists.get(name, []))
        want = min(counts.get(name, 0), len(pool))
        if want <= 0:
            continue
        if prefer_unused:
            used_idx = used_by_list.get(name, set())
            unused = [q for q in pool if q["_source_index"] not in used_idx]
            seen = [q for q in pool if q["_source_index"] in used_idx]
            random.shuffle(unused)
            random.shuffle(seen)
            chosen = (unused + seen)[:want]
        else:
            chosen = random.sample(pool, want)
        result.extend(chosen)

    random.shuffle(result)
    return result


def save_sample_set(name, composition, questions):
    sample_sets = load_json_file(SAMPLE_SETS_FILE, [])
    record = {
        "id": f"{int(time.time() * 1000)}",
        "name": name,
        "created_at": datetime.datetime.now().isoformat(timespec="seconds"),
        "composition": composition,
        "questions": questions,
    }
    sample_sets.append(record)
    save_json_file(SAMPLE_SETS_FILE, sample_sets)
    return record


def delete_sample_set(sample_id):
    sample_sets = load_json_file(SAMPLE_SETS_FILE, [])
    sample_sets = [s for s in sample_sets if s.get("id") != sample_id]
    save_json_file(SAMPLE_SETS_FILE, sample_sets)


def save_test_history(source_name, score, total, time_used_seconds):
    history = load_json_file(TEST_HISTORY_FILE, [])
    pct = round((score / total * 100.0) if total else 0.0, 1)
    history.append({
        "id": f"{int(time.time() * 1000)}",
        "source": source_name,
        "score": score,
        "total": total,
        "pct": pct,
        "date": datetime.datetime.now().isoformat(timespec="seconds"),
        "time_used_seconds": round(time_used_seconds, 1),
    })
    save_json_file(TEST_HISTORY_FILE, history)


def webDataExtract(chosenTopic):
    topics = fetch_all_topics()
    return topics.get(chosenTopic)


def clean_ocr_text(text):
    """
    Cleans up common OCR / scraping artifacts in question bank text so the
    database doesn't need to be hand-fixed entry by entry:

      1. Rejoins words that were split across a line break with a hyphen
         (e.g. "equilib-\\nrium" -> "equilibrium").
      2. Collapses stray newlines/tabs/runs of whitespace into single
         spaces, while still respecting an intentional blank-line paragraph
         break if one is present in the source.
      3. Inserts a missing space after . , ; : ! ? when the next character
         is a letter jammed directly against the punctuation.
      4. Inserts a missing space where a lowercase word runs directly into
         a following capitalized word with no space between them (a common
         OCR line-wrap artifact), e.g. "the reactionOccurs quickly".
      5. Inserts a missing space where a number runs directly into a
         following capitalized word, e.g. "at 350KThe pressure...".

    Steps 4 and 5 are deliberately conservative (requiring a run of 2+
    lowercase letters on each side) so they don't mangle chemical formulas
    like NaOH, CaCO3, or H2SO4, which have short, alternating-case runs.
    """
    if not isinstance(text, str) or not text:
        return text

    text = re.sub(r'(\w)-\s*\n\s*(\w)', r'\1\2', text)

    PARAGRAPH_MARKER = "\x00"
    text = re.sub(r'\n\s*\n', PARAGRAPH_MARKER, text)
    text = re.sub(r'\s+', ' ', text)
    text = text.replace(PARAGRAPH_MARKER, '\n\n')

    text = re.sub(r'([.,;:!?])(?=[A-Za-z])', r'\1 ', text)

    text = re.sub(r'([a-z]{2,})([A-Z][a-z]{2,})', r'\1 \2', text)

    text = re.sub(r'(\d)([A-Z][a-z]{2,})', r'\1 \2', text)

    text = re.sub(r' {2,}', ' ', text)

    return text.strip()


def clean_question_bank(questions):
    """Applies clean_ocr_text to every question's text and answer options.
    Preserves any extra keys already on the dict (e.g. _source_list /
    _source_index) since it copies via dict(q)."""
    cleaned = []
    for q in questions:
        q = dict(q)
        if "text" in q:
            q["text"] = clean_ocr_text(q["text"])
        if "options" in q:
            q["options"] = [clean_ocr_text(opt) for opt in q["options"]]
        cleaned.append(q)
    return cleaned


# ----------------------------------------------------------------------
# Small reusable popup: prompts for one question's text, its answer
# options, and which option is correct. Used by the Local Bank Editor
# tab of the workbench.
# ----------------------------------------------------------------------
class QuestionEntryPopup:
    def __init__(self, master, on_submit, existing=None):
        self.on_submit = on_submit

        self.top = tk.Toplevel(master)
        self.top.title("Add Question" if existing is None else "Edit Question")
        self.top.configure(bg="#F3F4F6")
        self.top.grab_set()
        self.top.resizable(False, False)

        tk.Label(self.top, text="Question text:", font=("Arial", 9, "bold"),
                 bg="#F3F4F6", anchor="w").pack(fill=tk.X, padx=15, pady=(15, 2))
        self.text_box = tk.Text(self.top, font=("Arial", 10), wrap=tk.WORD, height=4,
                                 width=60, bd=1, relief=tk.SOLID)
        self.text_box.pack(padx=15, pady=(0, 10))

        self.option_vars = []
        options_frame = tk.LabelFrame(self.top, text=" Answer Options (pick the correct one) ",
                                       font=("Arial", 9, "bold"), bg="white")
        options_frame.pack(fill=tk.X, padx=15, pady=5)

        self.correct_var = tk.IntVar(value=0)
        self.entries = []
        labels = ["A", "B", "C", "D"]
        for i, letter in enumerate(labels):
            row = tk.Frame(options_frame, bg="white")
            row.pack(fill=tk.X, padx=8, pady=4)
            tk.Radiobutton(row, variable=self.correct_var, value=i, bg="white").pack(side=tk.LEFT)
            tk.Label(row, text=f"{letter}.", bg="white", font=("Arial", 9, "bold"), width=2).pack(side=tk.LEFT)
            e = tk.Entry(row, font=("Arial", 10), width=55)
            e.pack(side=tk.LEFT, fill=tk.X, expand=True)
            self.entries.append(e)

        if existing is not None:
            self.text_box.insert(tk.END, existing.get("text", ""))
            for i, opt in enumerate(existing.get("options", [])[:4]):
                self.entries[i].insert(0, opt)
            self.correct_var.set(existing.get("correct", 0))

        btn_frame = tk.Frame(self.top, bg="#F3F4F6")
        btn_frame.pack(fill=tk.X, padx=15, pady=15)
        tk.Button(btn_frame, text="Cancel", width=10, command=self.top.destroy).pack(side=tk.LEFT)
        tk.Button(btn_frame, text="Save Question", width=14, bg="#1E3A8A", fg="white",
                  font=("Arial", 9, "bold"), command=self._submit).pack(side=tk.RIGHT)

    def _submit(self):
        text = self.text_box.get("1.0", tk.END).strip()
        options = [e.get().strip() for e in self.entries]
        if not text:
            messagebox.showwarning("Missing Text", "Please enter the question text.", parent=self.top)
            return
        if any(not o for o in options):
            messagebox.showwarning("Missing Options", "Please fill in all four answer options.", parent=self.top)
            return
        question = {"text": text, "options": options, "correct": self.correct_var.get()}
        self.on_submit(question)
        self.top.destroy()


# ----------------------------------------------------------------------
# Workbench: the new setup window. Replaces the old single-topic /
# single-duration SetupDialog with a tabbed workbench that can:
#   - evenly sample questions across one or several reference lists,
#   - save those samples locally and re-use them later,
#   - show per-list "percent already sampled" progress,
#   - show a history of tests already taken,
#   - let the user build their own local question banks by hand.
# ----------------------------------------------------------------------
class WorkbenchDialog:
    DURATION_PRESETS = {
        "1 Hour": 1 * 3600,
        "1.5 Hours": int(1.5 * 3600),
        "2 Hours": 2 * 3600,
        "3 Hours (Default)": 3 * 3600,
        "4 Hours": 4 * 3600,
        "Custom": None,
    }

    def __init__(self, master):
        self.master = master
        self.selected_topic = None
        self.selected_questions = None
        self.duration_seconds = 3 * 3600

        self.all_lists = get_all_lists()
        self._generated_questions = []
        self._generated_composition = {}
        self._editor_questions = []

        self.top = tk.Toplevel(master)
        self.top.title("CBLE Simulator - Workbench")
        self.top.geometry("980x720")
        self.top.minsize(900, 650)
        self.top.configure(bg="#F3F4F6")
        self.top.grab_set()
        self.top.protocol("WM_DELETE_WINDOW", self._on_cancel)

        header = tk.Label(self.top, text="Exam Workbench", font=("Arial", 15, "bold"),
                           bg="#F3F4F6", fg="#1E3A8A")
        header.pack(pady=(12, 6))

        self.notebook = ttk.Notebook(self.top)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 5))

        self.tab_builder = tk.Frame(self.notebook, bg="#F3F4F6")
        self.tab_saved = tk.Frame(self.notebook, bg="#F3F4F6")
        self.tab_history = tk.Frame(self.notebook, bg="#F3F4F6")
        self.tab_editor = tk.Frame(self.notebook, bg="#F3F4F6")

        self.notebook.add(self.tab_builder, text=" Sample Builder ")
        self.notebook.add(self.tab_saved, text=" Saved Samples ")
        self.notebook.add(self.tab_history, text=" Test History ")
        self.notebook.add(self.tab_editor, text=" Local Bank Editor ")

        self._build_sample_builder_tab(self.tab_builder)
        self._build_saved_samples_tab(self.tab_saved)
        self._build_history_tab(self.tab_history)
        self._build_editor_tab(self.tab_editor)

        # --- Shared duration + cancel bar at the bottom of the workbench ---
        bottom = tk.LabelFrame(self.top, text=" Exam Duration (applies to whichever exam you start) ",
                                font=("Arial", 9, "bold"), bg="white")
        bottom.pack(fill=tk.X, padx=15, pady=(0, 10))

        dur_row = tk.Frame(bottom, bg="white")
        dur_row.pack(fill=tk.X, padx=10, pady=8)

        self.duration_var = tk.StringVar(value="3 Hours (Default)")
        duration_menu = ttk.Combobox(dur_row, textvariable=self.duration_var,
                                      values=list(self.DURATION_PRESETS.keys()),
                                      state="readonly", font=("Arial", 10), width=20)
        duration_menu.pack(side=tk.LEFT)
        duration_menu.bind("<<ComboboxSelected>>", self._on_duration_change)

        self.custom_minutes_var = tk.StringVar(value="180")
        self._custom_label = tk.Label(dur_row, text="Custom minutes:", bg="white", font=("Arial", 9))
        self.custom_entry = tk.Entry(dur_row, textvariable=self.custom_minutes_var, width=8, font=("Arial", 9))
        # Both start hidden; _on_duration_change() shows them only when
        # "Custom" is picked from the duration dropdown.

        tk.Button(bottom, text="Close Workbench", font=("Arial", 9, "bold"), width=16,
                  command=self._on_cancel).pack(side=tk.RIGHT, padx=10, pady=8)

    # ------------------------------------------------------------
    # Shared helpers
    # ------------------------------------------------------------
    def _on_duration_change(self, event=None):
        if self.duration_var.get() == "Custom":
            self._custom_label.pack(side=tk.LEFT, padx=(15, 5))
            self.custom_entry.pack(side=tk.LEFT)
        else:
            self.custom_entry.pack_forget()
            self._custom_label.pack_forget()

    def _get_duration_seconds(self):
        choice = self.duration_var.get()
        if choice == "Custom":
            try:
                minutes = int(self.custom_minutes_var.get())
                if minutes <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showwarning("Invalid Duration", "Please enter a valid positive number of minutes.",
                                        parent=self.top)
                return None
            return minutes * 60
        return self.DURATION_PRESETS[choice]

    def _finalize_and_close(self, label, questions):
        if not questions:
            messagebox.showwarning("No Questions", "There are no questions to start an exam with.",
                                    parent=self.top)
            return
        duration = self._get_duration_seconds()
        if duration is None:
            return
        self.selected_topic = label
        self.selected_questions = questions
        self.duration_seconds = duration
        self.top.destroy()

    def _on_cancel(self):
        self.selected_topic = None
        self.selected_questions = None
        self.top.destroy()

    # ------------------------------------------------------------
    # Tab 1: Sample Builder (evenly sample across 1+ lists, save it,
    # and see per-list "already sampled" progress)
    # ------------------------------------------------------------
    def _build_sample_builder_tab(self, parent):
        top_frame = tk.Frame(parent, bg="#F3F4F6")
        top_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        left = tk.LabelFrame(top_frame, text=" Available Lists (select 1 or more) ",
                              font=("Arial", 9, "bold"), bg="white")
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 8))

        self.lists_box = tk.Listbox(left, selectmode=tk.EXTENDED, font=("Consolas", 9), height=16,
                                     activestyle="none")
        lists_scroll = ttk.Scrollbar(left, orient="vertical", command=self.lists_box.yview)
        self.lists_box.configure(yscrollcommand=lists_scroll.set)
        self.lists_box.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=8, pady=8)
        lists_scroll.pack(side=tk.RIGHT, fill=tk.Y, pady=8)

        btn_row = tk.Frame(left, bg="white")
        btn_row.pack(fill=tk.X, padx=8, pady=(0, 8))
        tk.Button(btn_row, text="Select All", command=self._select_all_lists).pack(side=tk.LEFT)
        tk.Button(btn_row, text="Clear", command=lambda: self.lists_box.selection_clear(0, tk.END)).pack(
            side=tk.LEFT, padx=6)
        tk.Button(btn_row, text="Refresh", command=self._refresh_lists_display).pack(side=tk.LEFT)

        right = tk.Frame(top_frame, bg="#F3F4F6", width=300)
        right.pack(side=tk.LEFT, fill=tk.Y)
        right.pack_propagate(False)

        settings = tk.LabelFrame(right, text=" Even Sampling Settings ", font=("Arial", 9, "bold"), bg="white")
        settings.pack(fill=tk.X, pady=(0, 8))

        tk.Label(settings, text="Total questions to sample:", bg="white", font=("Arial", 9)).pack(
            anchor="w", padx=10, pady=(10, 2))
        self.total_count_var = tk.StringVar(value="50")
        tk.Entry(settings, textvariable=self.total_count_var, width=10, font=("Arial", 10)).pack(
            anchor="w", padx=10, pady=(0, 8))

        self.prefer_unused_var = tk.BooleanVar(value=True)
        tk.Checkbutton(settings, text="Prioritize questions not yet\nincluded in any saved sample",
                        variable=self.prefer_unused_var, bg="white", font=("Arial", 8), justify=tk.LEFT).pack(
            anchor="w", padx=10, pady=(0, 10))

        tk.Button(settings, text="Generate Sample \u25b6", font=("Arial", 9, "bold"), bg="#1E3A8A", fg="white",
                  command=self._on_generate_sample).pack(fill=tk.X, padx=10, pady=(0, 10))

        preview_frame = tk.LabelFrame(right, text=" Preview ", font=("Arial", 9, "bold"), bg="white")
        preview_frame.pack(fill=tk.BOTH, expand=True)

        self.preview_text = tk.Text(preview_frame, font=("Consolas", 9), wrap=tk.WORD, height=8,
                                     bg="#FAFAFA", bd=0)
        self.preview_text.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        self.preview_text.config(state=tk.DISABLED)

        save_row = tk.Frame(right, bg="#F3F4F6")
        save_row.pack(fill=tk.X, pady=8)
        self.sample_name_var = tk.StringVar(value="")
        tk.Entry(save_row, textvariable=self.sample_name_var, font=("Arial", 9)).pack(
            side=tk.LEFT, fill=tk.X, expand=True)
        tk.Button(save_row, text="Save Sample", command=self._on_save_sample).pack(side=tk.LEFT, padx=(6, 0))

        tk.Button(right, text="Start Exam With This Sample \u25b6", font=("Arial", 10, "bold"),
                  bg="#166534", fg="white", command=self._on_start_from_builder).pack(fill=tk.X, pady=(0, 4))

        self._list_names_order = []
        self._refresh_lists_display()

    def _select_all_lists(self):
        self.lists_box.selection_set(0, tk.END)

    def _refresh_lists_display(self):
        self.all_lists = get_all_lists()
        coverage = compute_coverage(self.all_lists)
        self._list_names_order = sorted(self.all_lists.keys())

        self.lists_box.delete(0, tk.END)
        for name in self._list_names_order:
            used, total, pct = coverage.get(name, (0, 0, 0.0))
            line = f"{name:<28} | Total: {total:>4} | Sampled so far: {used:>4}/{total:<4} ({pct:5.1f}%)"
            self.lists_box.insert(tk.END, line)

        # Keep the other tabs' dropdowns/lists in sync too -- but only once
        # they actually exist. During the very first call (made while the
        # Sample Builder tab is still being constructed in __init__), the
        # Saved Samples and Local Bank Editor tabs haven't been built yet,
        # so checking for their *widgets* (not just these method names,
        # which always exist on the class) avoids an AttributeError.
        if hasattr(self, "saved_box"):
            self._refresh_saved_samples_list()
        if hasattr(self, "editor_bank_combo"):
            self._refresh_editor_bank_choices()

    def _on_generate_sample(self):
        selected_indices = self.lists_box.curselection()
        if not selected_indices:
            messagebox.showwarning("No Lists Selected", "Please select at least one list to sample from.",
                                    parent=self.top)
            return
        selected_names = [self._list_names_order[i] for i in selected_indices]

        try:
            total_count = int(self.total_count_var.get())
            if total_count <= 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("Invalid Count", "Please enter a positive whole number of questions.",
                                    parent=self.top)
            return

        questions = even_sample(self.all_lists, selected_names, total_count,
                                 prefer_unused=self.prefer_unused_var.get())
        self._generated_questions = questions
        composition = dict(Counter(q.get("_source_list", "?") for q in questions))
        self._generated_composition = composition

        self.preview_text.config(state=tk.NORMAL)
        self.preview_text.delete("1.0", tk.END)
        self.preview_text.insert(tk.END, f"Generated {len(questions)} question(s):\n\n")
        for name, count in sorted(composition.items()):
            self.preview_text.insert(tk.END, f"  {name}: {count}\n")
        self.preview_text.config(state=tk.DISABLED)

        if not self.sample_name_var.get().strip():
            self.sample_name_var.set(f"Sample_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}")

    def _on_save_sample(self):
        if not self._generated_questions:
            messagebox.showwarning("Nothing to Save", "Generate a sample first.", parent=self.top)
            return
        name = self.sample_name_var.get().strip()
        if not name:
            messagebox.showwarning("Missing Name", "Please give this sample a name.", parent=self.top)
            return
        save_sample_set(name, self._generated_composition, self._generated_questions)
        messagebox.showinfo("Saved", f"Sample '{name}' saved locally.", parent=self.top)
        self._refresh_lists_display()

    def _on_start_from_builder(self):
        if not self._generated_questions:
            messagebox.showwarning("Nothing Generated", "Generate a sample first.", parent=self.top)
            return
        label = self.sample_name_var.get().strip() or "Custom Sample"
        self._finalize_and_close(label, list(self._generated_questions))

    # ------------------------------------------------------------
    # Tab 2: Saved Samples
    # ------------------------------------------------------------
    def _build_saved_samples_tab(self, parent):
        left = tk.LabelFrame(parent, text=" Previously Saved Samples ", font=("Arial", 9, "bold"), bg="white")
        left.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.saved_box = tk.Listbox(left, font=("Consolas", 9), height=16, activestyle="none")
        saved_scroll = ttk.Scrollbar(left, orient="vertical", command=self.saved_box.yview)
        self.saved_box.configure(yscrollcommand=saved_scroll.set)
        self.saved_box.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=8, pady=8)
        saved_scroll.pack(side=tk.RIGHT, fill=tk.Y, pady=8)

        btns = tk.Frame(parent, bg="#F3F4F6")
        btns.pack(fill=tk.X, padx=5, pady=(0, 5))
        tk.Button(btns, text="Start Exam With Selected \u25b6", font=("Arial", 9, "bold"), bg="#166534",
                  fg="white", command=self._on_start_from_saved).pack(side=tk.LEFT)
        tk.Button(btns, text="View Composition", command=self._on_view_saved_composition).pack(side=tk.LEFT, padx=8)
        tk.Button(btns, text="Delete Selected", fg="#B91C1C", command=self._on_delete_saved).pack(side=tk.LEFT)
        tk.Button(btns, text="Refresh", command=self._refresh_saved_samples_list).pack(side=tk.LEFT, padx=8)

        self._refresh_saved_samples_list()

    def _refresh_saved_samples_list(self):
        self._saved_sets = load_json_file(SAMPLE_SETS_FILE, [])
        self.saved_box.delete(0, tk.END)
        for s in sorted(self._saved_sets, key=lambda x: x.get("created_at", ""), reverse=True):
            n = len(s.get("questions", []))
            self.saved_box.insert(
                tk.END, f"{s.get('name', '(unnamed)'):<30} | {n:>3} questions | saved {s.get('created_at', '')}"
            )

    def _get_selected_saved_set(self):
        sel = self.saved_box.curselection()
        if not sel:
            return None
        ordered = sorted(self._saved_sets, key=lambda x: x.get("created_at", ""), reverse=True)
        return ordered[sel[0]]

    def _on_view_saved_composition(self):
        s = self._get_selected_saved_set()
        if not s:
            messagebox.showwarning("No Selection", "Select a saved sample first.", parent=self.top)
            return
        comp = s.get("composition", {})
        lines = [f"{k}: {v}" for k, v in sorted(comp.items())]
        messagebox.showinfo(f"Composition of '{s.get('name')}'", "\n".join(lines) or "(no data)", parent=self.top)

    def _on_start_from_saved(self):
        s = self._get_selected_saved_set()
        if not s:
            messagebox.showwarning("No Selection", "Select a saved sample first.", parent=self.top)
            return
        self._finalize_and_close(s.get("name", "Saved Sample"), list(s.get("questions", [])))

    def _on_delete_saved(self):
        s = self._get_selected_saved_set()
        if not s:
            messagebox.showwarning("No Selection", "Select a saved sample first.", parent=self.top)
            return
        if messagebox.askyesno("Confirm Delete", f"Delete saved sample '{s.get('name')}'? This cannot be undone.",
                                parent=self.top):
            delete_sample_set(s.get("id"))
            self._refresh_saved_samples_list()
            self._refresh_lists_display()

    # ------------------------------------------------------------
    # Tab 3: Test History
    # ------------------------------------------------------------
    def _build_history_tab(self, parent):
        wrap = tk.LabelFrame(parent, text=" Tests Taken ", font=("Arial", 9, "bold"), bg="white")
        wrap.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        columns = ("date", "source", "score", "pct", "time_used")
        self.history_tree = ttk.Treeview(wrap, columns=columns, show="headings", height=16)
        headings = {"date": "Date", "source": "Source", "score": "Score",
                    "pct": "Percent", "time_used": "Time Used"}
        widths = {"date": 160, "source": 260, "score": 90, "pct": 90, "time_used": 110}
        for c in columns:
            self.history_tree.heading(c, text=headings[c])
            self.history_tree.column(c, width=widths[c], anchor="center")
        hist_scroll = ttk.Scrollbar(wrap, orient="vertical", command=self.history_tree.yview)
        self.history_tree.configure(yscrollcommand=hist_scroll.set)
        self.history_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=8, pady=8)
        hist_scroll.pack(side=tk.RIGHT, fill=tk.Y, pady=8)

        btns = tk.Frame(parent, bg="#F3F4F6")
        btns.pack(fill=tk.X, padx=5, pady=(0, 5))
        tk.Button(btns, text="Refresh", command=self._refresh_history_list).pack(side=tk.LEFT)
        tk.Button(btns, text="Clear History", fg="#B91C1C", command=self._on_clear_history).pack(side=tk.LEFT, padx=8)

        self._refresh_history_list()

    def _refresh_history_list(self):
        for row in self.history_tree.get_children():
            self.history_tree.delete(row)
        history = load_json_file(TEST_HISTORY_FILE, [])
        for h in sorted(history, key=lambda x: x.get("date", ""), reverse=True):
            time_used = CBLESimulator._format_duration(h.get("time_used_seconds", 0))
            self.history_tree.insert("", tk.END, values=(
                h.get("date", ""), h.get("source", ""),
                f"{h.get('score', 0)}/{h.get('total', 0)}", f"{h.get('pct', 0)}%", time_used
            ))

    def _on_clear_history(self):
        if messagebox.askyesno("Confirm", "Clear all locally saved test history? This cannot be undone.",
                                parent=self.top):
            save_json_file(TEST_HISTORY_FILE, [])
            self._refresh_history_list()

    # ------------------------------------------------------------
    # Tab 4: Local Bank Editor (build your own list of dicts)
    # ------------------------------------------------------------
    def _build_editor_tab(self, parent):
        top = tk.Frame(parent, bg="#F3F4F6")
        top.pack(fill=tk.X, padx=5, pady=5)

        tk.Label(top, text="Bank name:", bg="#F3F4F6", font=("Arial", 9, "bold")).pack(side=tk.LEFT)
        self.editor_bank_var = tk.StringVar(value="")
        self.editor_bank_combo = ttk.Combobox(top, textvariable=self.editor_bank_var, font=("Arial", 9), width=30)
        self.editor_bank_combo.pack(side=tk.LEFT, padx=8)
        tk.Button(top, text="Load", command=self._on_load_editor_bank).pack(side=tk.LEFT)
        tk.Button(top, text="New / Clear", command=self._on_new_editor_bank).pack(side=tk.LEFT, padx=8)

        mid = tk.LabelFrame(parent, text=" Questions in This Bank ", font=("Arial", 9, "bold"), bg="white")
        mid.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.editor_box = tk.Listbox(mid, font=("Consolas", 9), height=14, activestyle="none")
        editor_scroll = ttk.Scrollbar(mid, orient="vertical", command=self.editor_box.yview)
        self.editor_box.configure(yscrollcommand=editor_scroll.set)
        self.editor_box.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=8, pady=8)
        editor_scroll.pack(side=tk.RIGHT, fill=tk.Y, pady=8)

        btns = tk.Frame(parent, bg="#F3F4F6")
        btns.pack(fill=tk.X, padx=5, pady=(0, 5))
        tk.Button(btns, text="Add Question", font=("Arial", 9, "bold"), bg="#EDE9FE", fg="#5B21B6",
                  command=self._open_add_question_popup).pack(side=tk.LEFT)
        tk.Button(btns, text="Delete Selected", command=self._delete_selected_editor_question).pack(
            side=tk.LEFT, padx=8)
        tk.Button(btns, text="Save Bank", font=("Arial", 9, "bold"), bg="#1E3A8A", fg="white",
                  command=self._save_editor_bank).pack(side=tk.LEFT, padx=8)
        tk.Button(btns, text="Delete Bank", fg="#B91C1C", command=self._delete_editor_bank).pack(side=tk.LEFT)

        self._refresh_editor_bank_choices()

    def _refresh_editor_bank_choices(self):
        raw = load_json_file(LOCAL_BANKS_FILE, {})
        self.editor_bank_combo["values"] = sorted(raw.keys())

    def _on_new_editor_bank(self):
        self.editor_bank_var.set("")
        self._editor_questions = []
        self._refresh_editor_listbox()

    def _on_load_editor_bank(self):
        name = self.editor_bank_var.get().strip()
        if not name:
            messagebox.showwarning("No Bank Name", "Type or pick a bank name to load.", parent=self.top)
            return
        raw = load_json_file(LOCAL_BANKS_FILE, {})
        self._editor_questions = list(raw.get(name, []))
        self._refresh_editor_listbox()

    def _refresh_editor_listbox(self):
        self.editor_box.delete(0, tk.END)
        for i, q in enumerate(self._editor_questions):
            preview = q["text"][:70].replace("\n", " ")
            self.editor_box.insert(tk.END, f"{i + 1}. {preview}")

    def _open_add_question_popup(self):
        def on_submit(question):
            self._editor_questions.append(question)
            self._refresh_editor_listbox()
        QuestionEntryPopup(self.top, on_submit)

    def _delete_selected_editor_question(self):
        sel = self.editor_box.curselection()
        if not sel:
            return
        idx = sel[0]
        del self._editor_questions[idx]
        self._refresh_editor_listbox()

    def _save_editor_bank(self):
        name = self.editor_bank_var.get().strip()
        if not name:
            messagebox.showwarning("Missing Name", "Please give this local bank a name.", parent=self.top)
            return
        if not self._editor_questions:
            messagebox.showwarning("Empty Bank", "Add at least one question first.", parent=self.top)
            return
        save_local_bank(name, self._editor_questions)
        messagebox.showinfo("Saved", f"Local bank '{name}' saved with {len(self._editor_questions)} question(s).",
                             parent=self.top)
        self._refresh_editor_bank_choices()
        self._refresh_lists_display()

    def _delete_editor_bank(self):
        name = self.editor_bank_var.get().strip()
        if not name:
            return
        if messagebox.askyesno("Confirm Delete", f"Delete local bank '{name}'? This cannot be undone.",
                                parent=self.top):
            delete_local_bank(name)
            self._on_new_editor_bank()
            self._refresh_editor_bank_choices()
            self._refresh_lists_display()


class CBLESimulator:
    def __init__(self, root, topic_label, questions, duration_seconds=3 * 3600):
        self.root = root
        self.topic = topic_label
        self.root.title(f"PRC CBLE Simulation - Chemical Engineering Licensure Exam ({topic_label})")
        # A fixed 750px-tall window wasn't always enough room for the
        # question box + all 4 answer options + action buttons + status bar,
        # especially with longer question/option text -- options C and D
        # would get pushed below the visible window with no way to see them.
        # Start maximized (falls back to a taller fixed size if the window
        # manager doesn't support "zoomed"), and enforce a minimum size so
        # it can't be shrunk back down to the point where content clips.
        self.root.minsize(1100, 780)
        try:
            self.root.state("zoomed")
        except tk.TclError:
            self.root.geometry("1200x900")
        self.root.configure(bg="#F3F4F6")

        raw_questions = questions
        if not raw_questions:
            messagebox.showerror(
                "Startup Error",
                f"No question data found for '{topic_label}'. The simulator will now close."
            )
            self.root.destroy()
            return

        # Auto-clean OCR/scrape artifacts (missing spaces, broken hyphenation,
        # stray line breaks) so the source database doesn't need manual fixing.
        cleaned_questions = clean_question_bank(list(raw_questions))
        # CRITICAL: Shuffle data randomly at startup execution
        self.questions = cleaned_questions
        random.shuffle(self.questions)

        # State Tracking
        self.current_index = 0
        self.total_questions = len(self.questions)
        self.user_answers = {i: None for i in range(self.total_questions)}
        self.flagged_questions = {i: False for i in range(self.total_questions)}
        self.time_left = duration_seconds  # configurable via the Workbench (default 3 hours)

        # Per-question time tracking (seconds spent on each question)
        self.question_times = {i: 0.0 for i in range(self.total_questions)}
        self.question_start_time = None  # set whenever a question is loaded/viewed

        # Per-question "Notes" -- a typed explanation and/or a freehand
        # sketch, both edited together in a single self-contained popup (see
        # open_notes_popup) rather than an inline panel, so visibility never
        # depends on how much vertical space is left in the exam window.
        # note_images holds a PIL.Image per question (created lazily, only
        # once the user actually draws something) so we can tell drawn
        # questions apart from untouched ones when compiling the notes PDF.
        self.note_texts = {i: "" for i in range(self.total_questions)}
        self.note_images = {i: None for i in range(self.total_questions)}
        self._notes_last_point = None
        self._notes_popup_index = None    # which question the open popup (if any) belongs to
        self._notes_popup_text = None     # the popup's Text widget
        self._notes_popup_canvas = None   # the popup's live drawing canvas
        self._notes_popup_photo = None    # keeps the popup's background PhotoImage alive

        self.selected_option = tk.IntVar(value=-1)

        self.create_widgets()
        self.update_timer()
        self.load_question()

    def create_widgets(self):
        # 1. Top Bar (Header & Timer)
        top_bar = tk.Frame(self.root, bg="#1E3A8A", height=60)
        top_bar.pack(fill=tk.X, side=tk.TOP)

        title_lbl = tk.Label(top_bar, text="BOARD OF CHEMICAL ENGINEERING - CBLE RANDOMIZED SIMULATION (DAY 1)", fg="white", bg="#1E3A8A", font=("Arial", 12, "bold"))
        title_lbl.pack(side=tk.LEFT, padx=15, pady=15)

        self.timer_lbl = tk.Label(top_bar, text="Time Remaining: 03:00:00", fg="#EF4444", bg="#1E3A8A", font=("Arial", 13, "bold"))
        self.timer_lbl.pack(side=tk.RIGHT, padx=15, pady=15)

        # 2. Main Body Container
        main_body = tk.Frame(self.root, bg="#F3F4F6")
        main_body.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Left Column: Exam Presentation
        left_col = tk.Frame(main_body, bg="#F3F4F6")
        left_col.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))

        # Question Panel
        q_frame = tk.LabelFrame(left_col, text=" Question Box ", font=("Arial", 10, "bold"), bg="white", bd=2, relief=tk.GROOVE)
        q_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        # Text widgets default to a height of 24 rows when none is
        # given, which alone could eat most of the window's vertical space
        # and squeeze the answer options off-screen. Give it a modest fixed
        # height plus its own scrollbar, so long questions stay fully
        # readable (via scrolling) without shrinking the space reserved for
        # the answer options below.
        q_text_frame = tk.Frame(q_frame, bg="white")
        q_text_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)

        self.q_text = tk.Text(q_text_frame, font=("Arial", 11), wrap=tk.WORD, height=8, bg="white", bd=0, highlightthickness=0)
        q_text_scrollbar = ttk.Scrollbar(q_text_frame, orient="vertical", command=self.q_text.yview)
        self.q_text.configure(yscrollcommand=q_text_scrollbar.set)

        self.q_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        q_text_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.q_text.config(state=tk.DISABLED)

        # Multiple Choice Options Frame
        self.options_frame = tk.LabelFrame(left_col, text=" Select Your Answer ", font=("Arial", 10, "bold"), bg="white", bd=2, relief=tk.GROOVE)
        self.options_frame.pack(fill=tk.X, pady=5)

        self.radio_buttons = []
        for i in range(4):
            # wraplength/justify so long answer options wrap onto a new
            # line within the window instead of extending past its edge
            # (which was making the frame request more width than was
            # visible, and the buttons below get clipped as a result).
            rb = tk.Radiobutton(self.options_frame, text="", variable=self.selected_option, value=i,
                                 font=("Arial", 11), bg="white", anchor="w", justify=tk.LEFT,
                                 wraplength=1000, command=self.save_answer, padx=20, pady=8)
            rb.pack(fill=tk.X)
            self.radio_buttons.append(rb)

        # Keep the wraplength in sync with the actual available width so
        # options wrap correctly whether the window is maximized on a
        # small laptop screen or a large monitor.
        def _update_option_wraplength(event):
            wrap = max(200, event.width - 80)
            for rb in self.radio_buttons:
                rb.config(wraplength=wrap)
        self.options_frame.bind("<Configure>", _update_option_wraplength)

        # Bottom Actions Panel
        action_frame = tk.Frame(left_col, bg="#F3F4F6")
        action_frame.pack(fill=tk.X, pady=10)

        self.btn_prev = tk.Button(action_frame, text="\u25c0 Previous", font=("Arial", 10, "bold"), width=12, command=self.prev_question)
        self.btn_prev.pack(side=tk.LEFT)

        self.btn_flag = tk.Button(action_frame, text="\U0001F3F3 Flag Question", font=("Arial", 10, "bold"), fg="#B45309", bg="#FEF3C7", width=15, command=self.toggle_flag)
        self.btn_flag.pack(side=tk.LEFT, padx=10)

        # Notes -- opens a self-contained popup (typed explanation +
        # sketch) rather than an inline panel, so it's never at the mercy of
        # how much vertical space is left in the exam window.
        self.btn_notes = tk.Button(action_frame, text="\U0001F4DD Notes", font=("Arial", 10, "bold"),
                                    fg="#5B21B6", bg="#EDE9FE", width=12, command=self.open_notes_popup)
        self.btn_notes.pack(side=tk.LEFT, padx=(0, 10))

        self.btn_next = tk.Button(action_frame, text="Next \u25b6", font=("Arial", 10, "bold"), width=12, command=self.next_question)
        self.btn_next.pack(side=tk.LEFT)

        self.btn_submit = tk.Button(action_frame, text="Submit Exam", font=("Arial", 10, "bold"), fg="white", bg="#DC2626", width=12, command=self.confirm_submission)
        self.btn_submit.pack(side=tk.RIGHT)

        # Live Summary Notice Status Bar
        self.status_lbl = tk.Label(left_col, text="Answered: 0 | Flagged: 0 | Remaining: 100", font=("Arial", 10, "italic"), bg="#F3F4F6", fg="#4B5563")
        self.status_lbl.pack(side=tk.BOTTOM, fill=tk.X, pady=5)

        # Right Column: Question Navigation Grid Matrix
        right_col = tk.LabelFrame(main_body, text=" Question Matrix ", font=("Arial", 10, "bold"), bg="white", width=280, bd=2, relief=tk.GROOVE)
        right_col.pack(side=tk.RIGHT, fill=tk.Y, padx=(10, 0))
        right_col.pack_propagate(False)

        self.matrix_canvas = tk.Canvas(right_col, bg="white", borderwidth=0, highlightthickness=0)
        matrix_scrollbar = ttk.Scrollbar(right_col, orient="vertical", command=self.matrix_canvas.yview)
        self.matrix_container = tk.Frame(self.matrix_canvas, bg="white")

        self._matrix_window = self.matrix_canvas.create_window((0, 0), window=self.matrix_container, anchor="nw")

        self.matrix_container.bind(
            "<Configure>",
            lambda e: self.matrix_canvas.configure(scrollregion=self.matrix_canvas.bbox("all"))
        )
        # Keep the inner frame's width in sync with the canvas so the button
        # grid re-wraps correctly if the panel is ever resized.
        self.matrix_canvas.bind(
            "<Configure>",
            lambda e: self.matrix_canvas.itemconfigure(self._matrix_window, width=e.width)
        )
        self.matrix_canvas.configure(yscrollcommand=matrix_scrollbar.set)

        self.matrix_canvas.pack(side="left", fill="both", expand=True, padx=5, pady=5)
        matrix_scrollbar.pack(side="right", fill="y")

        # Mouse-wheel scrolling for the matrix panel. Bound only while the
        # cursor is over the panel so it doesn't hijack scrolling elsewhere
        # in the window. Handles Windows/Mac (<MouseWheel>) and Linux
        # (<Button-4>/<Button-5>) separately.
        def _on_matrix_mousewheel(event):
            if event.num == 4:          # Linux scroll up
                delta = -1
            elif event.num == 5:        # Linux scroll down
                delta = 1
            else:                       # Windows / Mac: event.delta is +/-120 (or +/-1 on Mac)
                delta = -1 if event.delta > 0 else 1
            self.matrix_canvas.yview_scroll(delta, "units")

        def _bind_matrix_mousewheel(event):
            self.matrix_canvas.bind_all("<MouseWheel>", _on_matrix_mousewheel)
            self.matrix_canvas.bind_all("<Button-4>", _on_matrix_mousewheel)
            self.matrix_canvas.bind_all("<Button-5>", _on_matrix_mousewheel)

        def _unbind_matrix_mousewheel(event):
            self.matrix_canvas.unbind_all("<MouseWheel>")
            self.matrix_canvas.unbind_all("<Button-4>")
            self.matrix_canvas.unbind_all("<Button-5>")

        self.matrix_canvas.bind("<Enter>", _bind_matrix_mousewheel)
        self.matrix_canvas.bind("<Leave>", _unbind_matrix_mousewheel)

        self.matrix_buttons = {}
        for idx in range(self.total_questions):
            row = idx // 5
            col = idx % 5
            btn = tk.Button(self.matrix_container, text=str(idx + 1), font=("Arial", 8, "bold"), width=4, height=2, bg="#E5E7EB", relief=tk.FLAT, command=lambda i=idx: self.jump_to_question(i))
            btn.grid(row=row, column=col, padx=3, pady=3)
            self.matrix_buttons[idx] = btn

    # ------------------------------------------------------------------
    # Notes -- a single self-contained popup holding both the typed
    # explanation and the ink sketch for whichever question is on screen.
    # Putting everything in its own Toplevel (rather than an inline panel)
    # means it's never at the mercy of how much vertical space is left in
    # the exam window -- it always gets its own reasonably-sized window.
    # ------------------------------------------------------------------
    def _get_or_create_note_image(self, index):
        """Returns (image, draw) for the given question, creating a blank
        white canvas-sized image the first time the user actually draws on it."""
        if self.note_images[index] is None:
            self.note_images[index] = Image.new("RGB", (NOTES_CANVAS_W, NOTES_CANVAS_H), "white")
        return self.note_images[index], ImageDraw.Draw(self.note_images[index])

    def open_notes_popup(self):
        """Opens a modal popup with a text box (explanation) and, if Pillow
        is available, an ink canvas (sketch) for the current question."""
        index = self.current_index
        self._notes_popup_index = index

        popup = tk.Toplevel(self.root)
        popup.title(f"Notes \u2014 Question {index + 1}")
        popup.configure(bg="#F3F4F6")
        popup.resizable(False, False)
        popup.grab_set()
        popup.protocol("WM_DELETE_WINDOW", lambda: self._close_notes_popup(popup))

        tk.Label(popup, text=f"Question {index + 1} \u2014 explain your answer and/or sketch below",
                 font=("Arial", 10, "bold"), bg="#F3F4F6", fg="#1E3A8A").pack(padx=15, pady=(15, 8))

        text_frame = tk.Frame(popup, bg="#F3F4F6")
        text_frame.pack(fill=tk.X, padx=15)
        tk.Label(text_frame, text="Explain your reasoning:", font=("Arial", 9, "bold"),
                 bg="#F3F4F6", anchor="w").pack(fill=tk.X)

        text_box_frame = tk.Frame(text_frame, bg="#F3F4F6")
        text_box_frame.pack(fill=tk.X, pady=(4, 10))
        self._notes_popup_text = tk.Text(text_box_frame, font=("Arial", 10), wrap=tk.WORD, height=6,
                                          bg="white", bd=1, relief=tk.SOLID, highlightthickness=0)
        notes_scrollbar = ttk.Scrollbar(text_box_frame, orient="vertical", command=self._notes_popup_text.yview)
        self._notes_popup_text.configure(yscrollcommand=notes_scrollbar.set)
        self._notes_popup_text.pack(side=tk.LEFT, fill=tk.X, expand=True)
        notes_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self._notes_popup_text.insert(tk.END, self.note_texts.get(index, ""))

        if PIL_AVAILABLE:
            tk.Label(popup, text="Sketch / diagram:", font=("Arial", 9, "bold"),
                     bg="#F3F4F6", anchor="w").pack(fill=tk.X, padx=15)

            self._notes_popup_canvas = tk.Canvas(popup, width=NOTES_CANVAS_W, height=NOTES_CANVAS_H,
                                                  bg="white", bd=1, relief=tk.SOLID, highlightthickness=0,
                                                  cursor="pencil")
            self._notes_popup_canvas.pack(padx=15, pady=(4, 5))

            img = self.note_images.get(index)
            if img is not None:
                self._notes_popup_photo = ImageTk.PhotoImage(img)
                self._notes_popup_canvas.create_image(0, 0, anchor="nw", image=self._notes_popup_photo)
            else:
                self._notes_popup_photo = None

            self._notes_popup_canvas.bind("<Button-1>", self._notes_canvas_press)
            self._notes_popup_canvas.bind("<B1-Motion>", self._notes_canvas_drag)
            self._notes_popup_canvas.bind("<ButtonRelease-1>", self._notes_canvas_release)

            tk.Button(popup, text="Clear Drawing", font=("Arial", 8, "bold"), bg="#FEE2E2", fg="#B91C1C",
                      relief=tk.FLAT, command=self._clear_notes_popup_drawing).pack(padx=15, pady=(0, 5), anchor="w")
        else:
            self._notes_popup_canvas = None
            tk.Label(popup, text="\u26a0 Install 'Pillow' (pip install Pillow) to enable sketching.",
                     font=("Arial", 9, "italic"), bg="#F3F4F6", fg="#DC2626").pack(padx=15, pady=(0, 10))

        tk.Button(popup, text="Save & Close", font=("Arial", 10, "bold"), bg="#1E3A8A", fg="white", width=14,
                  command=lambda: self._close_notes_popup(popup)).pack(pady=(5, 15))

    def _clear_notes_popup_drawing(self):
        if self._notes_popup_canvas is not None:
            self._notes_popup_canvas.delete("all")
        self.note_images[self._notes_popup_index] = None
        self._notes_popup_photo = None

    def _close_notes_popup(self, popup):
        if self._notes_popup_text is not None:
            self.note_texts[self._notes_popup_index] = self._notes_popup_text.get("1.0", tk.END).rstrip("\n")
        popup.grab_release()
        popup.destroy()
        self._notes_popup_text = None
        self._notes_popup_canvas = None
        self._notes_popup_photo = None
        self._notes_popup_index = None
        self._update_notes_button_indicator()

    def _notes_canvas_press(self, event):
        self._notes_last_point = (event.x, event.y)

    def _notes_canvas_drag(self, event):
        if not PIL_AVAILABLE or self._notes_last_point is None or self._notes_popup_canvas is None:
            return
        x0, y0 = self._notes_last_point
        x1, y1 = event.x, event.y

        # Live feedback on the popup's visible canvas...
        self._notes_popup_canvas.create_line(x0, y0, x1, y1, fill="black", width=2,
                                              capstyle=tk.ROUND, smooth=True)
        # ...mirrored into the off-screen raster image used later for the PDF.
        _, draw = self._get_or_create_note_image(self._notes_popup_index)
        draw.line([x0, y0, x1, y1], fill="black", width=2)

        self._notes_last_point = (x1, y1)

    def _notes_canvas_release(self, event):
        self._notes_last_point = None

    def _update_notes_button_indicator(self):
        """Marks the Notes button when the current question already has a
        saved explanation and/or sketch, so it's obvious at a glance."""
        if not hasattr(self, "btn_notes"):
            return
        has_notes = bool(self.note_texts.get(self.current_index, "").strip()) or \
            (self.note_images.get(self.current_index) is not None)
        if has_notes:
            self.btn_notes.config(text="\U0001F4DD Notes \u2713", bg="#DDD6FE")
        else:
            self.btn_notes.config(text="\U0001F4DD Notes", bg="#EDE9FE")

    # ------------------------------------------------------------------
    # Centralized time tracking helper.
    # Call this BEFORE moving away from the currently displayed question
    # so the elapsed seconds get credited to the right question index.
    # ------------------------------------------------------------------
    def _bank_elapsed_time(self):
        if self.question_start_time is not None:
            elapsed = time.time() - self.question_start_time
            self.question_times[self.current_index] += elapsed
        self.question_start_time = None

    def load_question(self):
        q = self.questions[self.current_index]

        self.q_text.config(state=tk.NORMAL)
        self.q_text.delete("1.0", tk.END)
        # Formats the current sequence string smoothly dynamically
        self.q_text.insert(tk.END, f"Question {self.current_index + 1}:\n" + q["text"])
        self.q_text.config(state=tk.DISABLED)

        for i, opt in enumerate(q["options"]):
            self.radio_buttons[i].config(text=opt)

        saved = self.user_answers[self.current_index]
        self.selected_option.set(saved if saved is not None else -1)

        if self.flagged_questions[self.current_index]:
            self.btn_flag.config(text="\U0001F3F3 Unflag", bg="#F59E0B", fg="white")
        else:
            self.btn_flag.config(text="\U0001F3F3 Flag Question", bg="#FEF3C7", fg="#B45309")

        self._update_notes_button_indicator()

        self.refresh_matrix()

        # Start (or restart) the clock for whichever question is now on screen
        self.question_start_time = time.time()

    def save_answer(self):
        self.user_answers[self.current_index] = self.selected_option.get()
        self.refresh_matrix()

    def toggle_flag(self):
        self.flagged_questions[self.current_index] = not self.flagged_questions[self.current_index]
        if self.flagged_questions[self.current_index]:
            self.btn_flag.config(text="\U0001F3F3 Unflag", bg="#F59E0B", fg="white")
        else:
            self.btn_flag.config(text="\U0001F3F3 Flag Question", bg="#FEF3C7", fg="#B45309")
        self.refresh_matrix()

    def refresh_matrix(self):
        answered_count = 0
        flagged_count = 0

        for idx, btn in self.matrix_buttons.items():
            is_answered = self.user_answers[idx] is not None
            is_flagged = self.flagged_questions[idx]

            if is_answered: answered_count += 1
            if is_flagged: flagged_count += 1

            if idx == self.current_index:
                btn.config(bd=1, relief=tk.SOLID, highlightbackground="black")
            else:
                btn.config(bd=0, relief=tk.FLAT)

            if is_flagged:
                btn.config(bg="#F59E0B", fg="white")
            elif is_answered:
                btn.config(bg="#10B981", fg="white")
            else:
                btn.config(bg="#E5E7EB", fg="black")

        remaining = self.total_questions - answered_count
        self.status_lbl.config(text=f"Answered: {answered_count} | Flagged: {flagged_count} | Unanswered Remaining: {remaining}")

        self.btn_prev.config(state=tk.DISABLED if self.current_index == 0 else tk.NORMAL)
        self.btn_next.config(state=tk.DISABLED if self.current_index == self.total_questions - 1 else tk.NORMAL)

    def jump_to_question(self, index):
        self._bank_elapsed_time()
        self.current_index = index
        self.load_question()

    def next_question(self):
        if self.current_index < self.total_questions - 1:
            self._bank_elapsed_time()
            self.current_index += 1
            self.load_question()

    def prev_question(self):
        if self.current_index > 0:
            self._bank_elapsed_time()
            self.current_index -= 1
            self.load_question()

    def update_timer(self):
        if self.time_left > 0:
            self.time_left -= 1
            hrs = self.time_left // 3600
            mins = (self.time_left % 3600) // 60
            secs = self.time_left % 60
            self.timer_lbl.config(text=f"Time Remaining: {hrs:02d}:{mins:02d}:{secs:02d}")
            self.root.after(1000, self.update_timer)
        else:
            self.show_report()

    def confirm_submission(self):
        unanswered = sum(1 for ans in self.user_answers.values() if ans is None)
        msg = "Are you sure you want to submit and end the exam?"
        if unanswered > 0:
            msg = f"You still have {unanswered} unanswered question(s). \n\n" + msg

        if messagebox.askyesno("Confirm Submit", msg):
            self.show_report()

    @staticmethod
    def _format_duration(seconds):
        seconds = int(round(seconds))
        hrs = seconds // 3600
        mins = (seconds % 3600) // 60
        secs = seconds % 60
        if hrs > 0:
            return f"{hrs:02d}:{mins:02d}:{secs:02d}"
        return f"{mins:02d}:{secs:02d}"

    def show_report(self):
        # Make sure whatever question was on-screen at submit time gets its
        # elapsed time banked before we build the report. (Typed notes are
        # already saved the moment their popup is closed, so nothing extra
        # to flush here.)
        self._bank_elapsed_time()

        self.score = 0
        self.all_items_report = []
        self.sort_by_wrong_first = False

        for i, q in enumerate(self.questions):
            user_ans = self.user_answers[i]
            correct_ans = q["correct"]
            is_correct = (user_ans == correct_ans)
            is_answered = user_ans is not None  # explicit flag used for the flagged x answered combo report

            if is_correct:
                self.score += 1

            if user_ans is not None and 0 <= user_ans < len(q["options"]):
                chosen_text = q["options"][user_ans]
            else:
                chosen_text = "No Answer Provided (Skipped)"

            # Trace this question back to its exact spot in the original
            # reference list (e.g. list[0] -> Reference No. 1) so it can be
            # found in the source material without having to search for it.
            source_list = q.get("_source_list", self.topic)
            source_index = q.get("_source_index")
            reference_no = (source_index + 1) if source_index is not None else ""

            self.all_items_report.append({
                "number": i + 1,
                "is_correct": is_correct,
                "is_flagged": self.flagged_questions[i],
                "is_answered": is_answered,
                "time_spent_seconds": self.question_times[i],
                "question": q["text"],
                "your_answer": chosen_text,
                "correct_answer": q["options"][correct_ans],
                "source_list": source_list,
                "reference_no": reference_no,
            })

        # Record this attempt in the local test-history log so it shows up
        # in the Workbench's "Test History" tab next time it's opened.
        try:
            save_test_history(self.topic, self.score, self.total_questions, sum(self.question_times.values()))
        except Exception as e:
            print(f"Could not save test history: {e}")

        self.report_win = tk.Toplevel(self.root)
        self.report_win.title("ChELE Diagnostic Correction Report")
        self.report_win.geometry("900x700")
        self.report_win.configure(bg="white")
        self.report_win.grab_set()

        summary_frame = tk.Frame(self.report_win, bg="#1E3A8A")
        summary_frame.pack(fill=tk.X, padx=15, pady=10)

        status_text = "PASSED" if (self.score / self.total_questions) >= 0.70 else "FAILED"
        header_lbl = tk.Label(summary_frame, text=f"EXAM RESULTS: {self.score}/{self.total_questions} ({status_text})", fg="white", bg="#1E3A8A", font=("Arial", 14, "bold"))
        header_lbl.pack(pady=10)

        control_bar = tk.Frame(self.report_win, bg="#E5E7EB")
        control_bar.pack(fill=tk.X, padx=15, pady=(5, 0))

        self.btn_sort = tk.Button(control_bar, text="Sort Order: Wrong Answers First \u2195", font=("Arial", 10, "bold"), bg="#FFFFFF", fg="#1E3A8A", command=self.toggle_report_sort)
        self.btn_sort.pack(side=tk.LEFT, padx=10, pady=5)

        self.btn_copy = tk.Button(control_bar, text="\U0001F4CB Copy Report to Clipboard", font=("Arial", 10, "bold"), bg="#FFFFFF", fg="#1E3A8A", command=self.copy_report_to_clipboard)
        self.btn_copy.pack(side=tk.LEFT, padx=10, pady=5)

        self.btn_export = tk.Button(control_bar, text="\U0001F4CA Export to Excel", font=("Arial", 10, "bold"), bg="#FFFFFF", fg="#166534", command=self.export_report_to_excel)
        self.btn_export.pack(side=tk.LEFT, padx=10, pady=5)

        self.btn_export_notes = tk.Button(control_bar, text="\U0001F4DD Export Notes to PDF", font=("Arial", 10, "bold"), bg="#FFFFFF", fg="#7C3AED", command=self.export_notes_to_pdf)
        self.btn_export_notes.pack(side=tk.LEFT, padx=10, pady=5)

        # Single scrollable, selectable, copy-pasteable Text widget replaces
        # the old per-question card layout.
        text_frame = tk.Frame(self.report_win, bg="white")
        text_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=10)

        self.report_text = tk.Text(text_frame, font=("Consolas", 10), wrap=tk.WORD, bg="#FAFAFA", fg="#111827", bd=1, relief=tk.SOLID)
        report_scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=self.report_text.yview)
        self.report_text.configure(yscrollcommand=report_scrollbar.set)

        self.report_text.pack(side="left", fill="both", expand=True)
        report_scrollbar.pack(side="right", fill="y")

        self.render_report_text()

        def exit_all():
            self.report_win.destroy()
            self.root.destroy()

        close_btn = tk.Button(self.report_win, text="Close and Finish Simulator", font=("Arial", 11, "bold"), bg="#4B5563", fg="white", command=exit_all)
        close_btn.pack(side=tk.BOTTOM, fill=tk.X, padx=15, pady=10)

    def _build_report_string(self):
        """Builds the full plain-text, copy-pasteable exam report."""
        total_time_used = sum(self.question_times.values())
        pct = (self.score / self.total_questions) * 100 if self.total_questions else 0
        status_text = "PASSED" if (self.score / self.total_questions) >= 0.70 else "FAILED"

        if self.sort_by_wrong_first:
            display_list = sorted(self.all_items_report, key=lambda x: x["is_correct"])
            sort_label = "Wrong Answers First"
        else:
            display_list = sorted(self.all_items_report, key=lambda x: x["number"])
            sort_label = "Sequential (1 - N)"

        lines = []
        sep = "=" * 78
        subsep = "-" * 78

        lines.append(sep)
        lines.append("PRC CBLE SIMULATION - EXAM DIAGNOSTIC REPORT")
        lines.append(sep)
        lines.append(f"Topic:          {self.topic}")
        lines.append(f"Generated:      {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        lines.append(f"Score:          {self.score}/{self.total_questions} ({pct:.1f}%) - {status_text}")
        lines.append(f"Total Time Used:{' ':1}{self._format_duration(total_time_used)}")
        lines.append(f"Sort Order:     {sort_label}")
        lines.append(sep)
        lines.append("")

        for item in display_list:
            status_tag = "CORRECT" if item["is_correct"] else "INCORRECT/SKIPPED"
            flag_tag = "FLAGGED" if item["is_flagged"] else "Not Flagged"
            time_tag = self._format_duration(item["time_spent_seconds"])
            ref_tag = f"{item['source_list']} #{item['reference_no']}" if item["reference_no"] != "" else "N/A"

            lines.append(subsep)
            lines.append(f"Q{item['number']}  |  {status_tag}  |  {flag_tag}  |  Ref: {ref_tag}  |  Time Spent: {time_tag}")
            lines.append(subsep)
            lines.append(f"Question: {item['question']}")
            lines.append(f"Your Answer:    {item['your_answer']}")
            lines.append(f"Correct Answer: {item['correct_answer']}")
            lines.append("")

        return "\n".join(lines)

    def render_report_text(self):
        report_str = self._build_report_string()
        self.report_text.config(state=tk.NORMAL)
        self.report_text.delete("1.0", tk.END)
        self.report_text.insert(tk.END, report_str)
        # Leave state NORMAL so the user can select/copy freely; text
        # is regenerated wholesale on sort-toggle so accidental edits
        # by the user don't persist across a re-render.

    def copy_report_to_clipboard(self):
        report_str = self._build_report_string()
        self.root.clipboard_clear()
        self.root.clipboard_append(report_str)
        self.root.update()  # ensures clipboard content persists after app focus changes
        messagebox.showinfo("Copied", "The full report has been copied to your clipboard.")

    def export_report_to_excel(self):
        if not OPENPYXL_AVAILABLE:
            messagebox.showerror(
                "Missing Dependency",
                "Excel export requires the 'openpyxl' package.\n\nInstall it with:\n    pip install openpyxl"
            )
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")],
            initialfile=f"CBLE_Report_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            title="Save Exam Report As"
        )
        if not file_path:
            return

        # Fill colors keyed by (is_flagged, is_answered) so every row
        # in the Question Detail sheet is highlighted by which of the four
        # combinations it falls into. Reused for the legend in Summary too.
        COMBO_FILLS = {
            (True, True):   PatternFill(start_color="FDE68A", end_color="FDE68A", fill_type="solid"),  # Flagged & Answered
            (True, False):  PatternFill(start_color="FCA5A5", end_color="FCA5A5", fill_type="solid"),  # Flagged & Not Answered
            (False, True):  PatternFill(start_color="D1FAE5", end_color="D1FAE5", fill_type="solid"),  # Not Flagged & Answered
            (False, False): PatternFill(start_color="E5E7EB", end_color="E5E7EB", fill_type="solid"),  # Not Flagged & Not Answered
        }
        COMBO_LABELS = [
            ("Flagged & Answered", (True, True)),
            ("Flagged & Not Answered", (True, False)),
            ("Not Flagged & Answered", (False, True)),
            ("Not Flagged & Not Answered", (False, False)),
        ]

        wb = openpyxl.Workbook()

        # --- Summary sheet ---
        ws_summary = wb.active
        ws_summary.title = "Summary"
        total_time_used = sum(self.question_times.values())
        pct = (self.score / self.total_questions) * 100 if self.total_questions else 0
        status_text = "PASSED" if (self.score / self.total_questions) >= 0.70 else "FAILED"

        summary_rows = [
            ("Topic / Sample", self.topic),
            ("Generated", datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')),
            ("Score", f"{self.score}/{self.total_questions}"),
            ("Percentage", f"{pct:.1f}%"),
            ("Result", status_text),
            ("Total Time Used", self._format_duration(total_time_used)),
        ]
        for r_idx, (label, value) in enumerate(summary_rows, start=1):
            ws_summary.cell(row=r_idx, column=1, value=label).font = Font(bold=True)
            ws_summary.cell(row=r_idx, column=2, value=value)
        ws_summary.column_dimensions["A"].width = 24
        ws_summary.column_dimensions["B"].width = 30

        # Color legend so the Question Detail highlighting is self-explanatory
        legend_header_row = len(summary_rows) + 2
        ws_summary.cell(row=legend_header_row, column=1,
                         value="Color Legend (Question Detail sheet):").font = Font(bold=True)
        for i, (label, key) in enumerate(COMBO_LABELS):
            r = legend_header_row + 1 + i
            label_cell = ws_summary.cell(row=r, column=1, value=label)
            swatch_cell = ws_summary.cell(row=r, column=2, value="")
            label_cell.fill = COMBO_FILLS[key]
            swatch_cell.fill = COMBO_FILLS[key]

        # --- Detail sheet ---
        ws = wb.create_sheet("Question Detail")
        # "Source List" + "Reference #" let you go straight back to the
        # exact entry in the original reference material -- list[0] in
        # "Source List" corresponds to Reference # 1, etc. -- without
        # having to search for the question text in that source.
        headers = ["Q#", "Source List", "Reference #", "Status", "Answered", "Flagged",
                   "Time Spent (mm:ss)", "Time Spent (sec)", "Question", "Your Answer", "Correct Answer"]
        header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        for col_idx, header in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=col_idx, value=header)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")

        sorted_for_export = sorted(self.all_items_report, key=lambda x: x["number"])
        correct_font = Font(color="065F46", bold=True)   # dark green
        incorrect_font = Font(color="B91C1C", bold=True)  # dark red

        for row_idx, item in enumerate(sorted_for_export, start=2):
            combo_key = (item["is_flagged"], item["is_answered"])
            row_fill = COMBO_FILLS[combo_key]

            values = [
                item["number"],
                item["source_list"],
                item["reference_no"],
                "Correct" if item["is_correct"] else "Incorrect/Skipped",
                "Yes" if item["is_answered"] else "No",
                "Yes" if item["is_flagged"] else "No",
                self._format_duration(item["time_spent_seconds"]),
                round(item["time_spent_seconds"], 1),
                item["question"],
                item["your_answer"],
                item["correct_answer"],
            ]
            for col_idx, value in enumerate(values, start=1):
                cell = ws.cell(row=row_idx, column=col_idx, value=value)
                cell.fill = row_fill
                cell.alignment = Alignment(vertical="top", wrap_text=(col_idx >= 9))
                if col_idx == 4:
                    cell.font = correct_font if item["is_correct"] else incorrect_font

        col_widths = [6, 22, 12, 18, 10, 10, 18, 16, 60, 40, 40]
        for i, width in enumerate(col_widths, start=1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = width

        # Autofilter + frozen header row so the user can filter by
        # Status / Answered / Flagged / Source List directly in Excel,
        # e.g. Answered = No AND Flagged = Yes AND Source List = Thermo.
        last_row = len(sorted_for_export) + 1
        last_col_letter = openpyxl.utils.get_column_letter(len(headers))
        ws.auto_filter.ref = f"A1:{last_col_letter}{last_row}"
        ws.freeze_panes = "A2"

        try:
            wb.save(file_path)
            messagebox.showinfo("Export Successful", f"Report exported to:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Export Failed", f"Could not save the file:\n{e}")

    def export_notes_to_pdf(self):
        if not REPORTLAB_AVAILABLE:
            messagebox.showerror(
                "Missing Dependency",
                "Exporting notes requires the 'reportlab' package.\n\nInstall it with:\n    pip install reportlab"
            )
            return

        entries = []
        for i in range(self.total_questions):
            text = self.note_texts.get(i, "").strip()
            img = self.note_images.get(i)
            if text or img is not None:
                entries.append((i, text, img))

        if not entries:
            messagebox.showinfo(
                "No Notes",
                "You haven't written any notes or drawn anything yet, so there's nothing to export."
            )
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF Document", "*.pdf")],
            initialfile=f"CBLE_Notes_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
            title="Save Notes & Explanations PDF As"
        )
        if not file_path:
            return

        try:
            doc = SimpleDocTemplate(
                file_path, pagesize=letter,
                topMargin=0.75 * inch, bottomMargin=0.75 * inch,
                leftMargin=0.75 * inch, rightMargin=0.75 * inch
            )
            styles = getSampleStyleSheet()
            story = []

            story.append(Paragraph(f"CBLE Simulation \u2014 Notes &amp; Explanations ({html_escape(self.topic)})",
                                    styles["Title"]))
            story.append(Paragraph(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), styles["Normal"]))
            story.append(Spacer(1, 0.3 * inch))

            for pos, (idx, text, img) in enumerate(entries):
                q = self.questions[idx]
                story.append(Paragraph(f"Question {idx + 1}", styles["Heading2"]))
                story.append(Paragraph(_text_to_paragraph_html(q["text"]), styles["Normal"]))
                story.append(Spacer(1, 0.1 * inch))

                if text:
                    story.append(Paragraph("<b>Explanation:</b>", styles["Normal"]))
                    story.append(Paragraph(_text_to_paragraph_html(text), styles["Normal"]))
                    story.append(Spacer(1, 0.1 * inch))

                if img is not None:
                    story.append(Paragraph("<b>Sketch:</b>", styles["Normal"]))
                    buf = io.BytesIO()
                    img.save(buf, format="PNG")
                    buf.seek(0)
                    max_w = 6.0 * inch
                    w, h = img.size
                    scale = min(1.0, max_w / float(w))
                    story.append(RLImage(buf, width=w * scale, height=h * scale))

                if pos < len(entries) - 1:
                    story.append(PageBreak())

            doc.build(story)
            messagebox.showinfo("Export Successful", f"Notes PDF saved to:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Export Failed", f"Could not save the PDF:\n{e}")

    def toggle_report_sort(self):
        self.sort_by_wrong_first = not self.sort_by_wrong_first
        if self.sort_by_wrong_first:
            self.btn_sort.config(text="Sort Order: Sequential Chronological (1-100) \u2195")
        else:
            self.btn_sort.config(text="Sort Order: Wrong Answers First \u2195")
        self.render_report_text()


if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()  # hide the main window until setup is complete

    workbench = WorkbenchDialog(root)
    root.wait_window(workbench.top)

    if not workbench.selected_questions:
        # user cancelled, or never generated/selected anything to run
        root.destroy()
    else:
        root.deiconify()
        app = CBLESimulator(root, topic_label=workbench.selected_topic,
                             questions=workbench.selected_questions,
                             duration_seconds=workbench.duration_seconds)
        root.mainloop()
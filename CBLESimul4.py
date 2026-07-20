import random
import tkinter as tk
from tkinter import messagebox, ttk
import CBLEReviewer_Database
import ARCLectureBank

class CBLESimulator:
    def __init__(self, root):
        self.root = root
        self.root.title("PRC CBLE Simulation - Chemical Engineering Licensure Exam")
        self.root.geometry("1200x750")
        self.root.configure(bg="#F3F4F6")

        # 100 Balanced Board Exam Questions (Day 1: Physical and Chemical Principles)
        raw_questions = ARCLectureBank.ChECalc
        
        # CRITICAL: Shuffle data randomly at startup execution
        self.questions = list(raw_questions)
        random.shuffle(self.questions)
        
        # State Tracking
        self.current_index = 0
        self.total_questions = len(self.questions)
        self.user_answers = {i: None for i in range(self.total_questions)}
        self.flagged_questions = {i: False for i in range(self.total_questions)}
        self.time_left = 3 * 3600  # 3 Hours

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
        
        self.q_text = tk.Text(q_frame, font=("Arial", 11), wrap=tk.WORD, bg="white", bd=0, highlightthickness=0)
        self.q_text.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)
        self.q_text.config(state=tk.DISABLED)

        # Multiple Choice Options Frame
        self.options_frame = tk.LabelFrame(left_col, text=" Select Your Answer ", font=("Arial", 10, "bold"), bg="white", bd=2, relief=tk.GROOVE)
        self.options_frame.pack(fill=tk.X, pady=5)

        self.radio_buttons = []
        for i in range(4):
            rb = tk.Radiobutton(self.options_frame, text="", variable=self.selected_option, value=i, font=("Arial", 11), bg="white", anchor="w", command=self.save_answer, padx=20, pady=8)
            rb.pack(fill=tk.X)
            self.radio_buttons.append(rb)

        # Bottom Actions Panel
        action_frame = tk.Frame(left_col, bg="#F3F4F6")
        action_frame.pack(fill=tk.X, pady=10)

        self.btn_prev = tk.Button(action_frame, text="◀ Previous", font=("Arial", 10, "bold"), width=12, command=self.prev_question)
        self.btn_prev.pack(side=tk.LEFT)

        self.btn_flag = tk.Button(action_frame, text="🏳 Flag Question", font=("Arial", 10, "bold"), fg="#B45309", bg="#FEF3C7", width=15, command=self.toggle_flag)
        self.btn_flag.pack(side=tk.LEFT, padx=10)

        self.btn_next = tk.Button(action_frame, text="Next ▶", font=("Arial", 10, "bold"), width=12, command=self.next_question)
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

        matrix_canvas = tk.Canvas(right_col, bg="white", borderwidth=0, highlightthickness=0)
        matrix_scrollbar = ttk.Scrollbar(right_col, orient="vertical", command=matrix_canvas.yview)
        self.matrix_container = tk.Frame(matrix_canvas, bg="white")

        self.matrix_container.bind(
            "<Configure>",
            lambda e: matrix_canvas.configure(scrollregion=matrix_canvas.bbox("all"))
        )
        matrix_canvas.create_window((0, 0), window=self.matrix_container, anchor="nw")
        matrix_canvas.configure(yscrollcommand=matrix_scrollbar.set)

        matrix_canvas.pack(side="left", fill="both", expand=True, padx=5, pady=5)
        matrix_scrollbar.pack(side="right", fill="y")
        
        self.matrix_buttons = {}
        for idx in range(self.total_questions):
            row = idx // 5   
            col = idx % 5
            btn = tk.Button(self.matrix_container, text=str(idx + 1), font=("Arial", 8, "bold"), width=4, height=2, bg="#E5E7EB", relief=tk.FLAT, command=lambda i=idx: self.jump_to_question(i))
            btn.grid(row=row, column=col, padx=3, pady=3)
            self.matrix_buttons[idx] = btn

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
            self.btn_flag.config(text="🏳 Unflag", bg="#F59E0B", fg="white")
        else:
            self.btn_flag.config(text="🏳 Flag Question", bg="#FEF3C7", fg="#B45309")

        self.refresh_matrix()

    def save_answer(self):
        self.user_answers[self.current_index] = self.selected_option.get()
        self.refresh_matrix()

    def toggle_flag(self):
        self.flagged_questions[self.current_index] = not self.flagged_questions[self.current_index]
        if self.flagged_questions[self.current_index]:
            self.btn_flag.config(text="🏳 Unflag", bg="#F59E0B", fg="white")
        else:
            self.btn_flag.config(text="🏳 Flag Question", bg="#FEF3C7", fg="#B45309")
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
        self.current_index = index
        self.load_question()

    def next_question(self):
        if self.current_index < self.total_questions - 1:
            self.current_index += 1
            self.load_question()

    def prev_question(self):
        if self.current_index > 0:
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

    def show_report(self):
        self.score = 0
        self.all_items_report = []
        self.sort_by_wrong_first = False  

        for i, q in enumerate(self.questions):
            user_ans = self.user_answers[i]
            correct_ans = q["correct"]
            is_correct = (user_ans == correct_ans)
            
            if is_correct:
                self.score += 1

            if user_ans is not None and 0 <= user_ans < len(q["options"]):
                chosen_text = q["options"][user_ans]
            else:
                chosen_text = "No Answer Provided (Skipped)"

            self.all_items_report.append({
                "number": i + 1,
                "is_correct": is_correct,
                "question": q["text"],
                "your_answer": chosen_text,
                "correct_answer": q["options"][correct_ans]
            })

        self.report_win = tk.Toplevel(self.root)
        self.report_win.title("ChELE Diagnostic Correction Report")
        self.report_win.geometry("800x650")
        self.report_win.configure(bg="white")
        self.report_win.grab_set()

        summary_frame = tk.Frame(self.report_win, bg="#1E3A8A")
        summary_frame.pack(fill=tk.X, padx=15, pady=10)
        
        status_text = "PASSED" if (self.score / self.total_questions) >= 0.70 else "FAILED"
        header_lbl = tk.Label(summary_frame, text=f"EXAM RESULTS: {self.score}/{self.total_questions} ({status_text})", fg="white", bg="#1E3A8A", font=("Arial", 14, "bold"))
        header_lbl.pack(pady=10)

        control_bar = tk.Frame(self.report_win, bg="#E5E7EB")
        control_bar.pack(fill=tk.X, padx=15, pady=(5, 0))
        
        self.btn_sort = tk.Button(control_bar, text="Sort Order: Wrong Answers First ↕", font=("Arial", 10, "bold"), bg="#FFFFFF", fg="#1E3A8A", command=self.toggle_report_sort)
        self.btn_sort.pack(side=tk.LEFT, padx=10, pady=5)

        self.scroll_canvas = tk.Canvas(self.report_win, bg="white", borderwidth=0, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.report_win, orient="vertical", command=self.scroll_canvas.yview)
        self.cards_inner_frame = tk.Frame(self.scroll_canvas, bg="white")

        self.cards_inner_frame.bind(
            "<Configure>",
            lambda e: self.scroll_canvas.configure(scrollregion=self.scroll_canvas.bbox("all"))
        )
        self.scroll_canvas.create_window((0, 0), window=self.cards_inner_frame, anchor="nw")
        self.scroll_canvas.configure(yscrollcommand=scrollbar.set)

        self.scroll_canvas.pack(side="left", fill="both", expand=True, padx=15, pady=15)
        scrollbar.pack(side="right", fill="y")

        self.render_report_cards()

        def exit_all():
            self.report_win.destroy()
            self.root.destroy()

        close_btn = tk.Button(self.report_win, text="Close and Finish Simulator", font=("Arial", 11, "bold"), bg="#4B5563", fg="white", command=exit_all)
        close_btn.pack(side=tk.BOTTOM, fill=tk.X, padx=15, pady=10)

    def render_report_cards(self):
        for widget in self.cards_inner_frame.winfo_children():
            widget.destroy()

        if self.sort_by_wrong_first:
            sorted_display_list = sorted(self.all_items_report, key=lambda x: x["is_correct"])
        else:
            sorted_display_list = sorted(self.all_items_report, key=lambda x: x["number"])

        for item in sorted_display_list:
            if item["is_correct"]:
                bg_color = "#F0FDF4"   
                fg_title = "#166534"   
                badge_lbl = "✅ CORRECT"
            else:
                bg_color = "#FEF2F2"   
                fg_title = "#991B1B"   
                badge_lbl = "❌ INCORRECT / SKIPPED"

            card_box = tk.LabelFrame(self.cards_inner_frame, text=f" Position #{item['number']} — {badge_lbl} ", font=("Arial", 10, "bold"), bg=bg_color, fg=fg_title, bd=1, relief=tk.SOLID)
            card_box.pack(fill=tk.X, pady=6, ipady=4, ipadx=5, expand=True)

            q_lbl = tk.Label(card_box, text=item["question"], font=("Arial", 10), bg=bg_color, justify=tk.LEFT, anchor="w", wraplength=680)
            q_lbl.pack(fill=tk.X, padx=5, pady=3)

            ans_lbl = tk.Label(card_box, text=f"Your Registered Choice: {item['your_answer']}", font=("Arial", 10, "bold" if not item['is_correct'] else "normal"), fg="#DC2626" if not item['is_correct'] else "#4B5563", bg=bg_color, anchor="w")
            ans_lbl.pack(fill=tk.X, padx=5)

            cor_lbl = tk.Label(card_box, text=f"Correct Answer Key: {item['correct_answer']}", font=("Arial", 10, "bold"), fg="#10B981", bg=bg_color, anchor="w")
            cor_lbl.pack(fill=tk.X, padx=5, pady=(0, 2))

    def toggle_report_sort(self):
        self.sort_by_wrong_first = not self.sort_by_wrong_first
        if self.sort_by_wrong_first:
            self.btn_sort.config(text="Sort Order: Sequential Chronological (1-100) ↕")
        else:
            self.btn_sort.config(text="Sort Order: Wrong Answers First ↕")
        self.render_report_cards()


if __name__ == "__main__":
    root = tk.Tk()
    app = CBLESimulator(root)
    root.mainloop()
#!/usr/bin/env python3
"""
JSON-to-DOCX Workflow Launcher - GUI Interface
Simple graphical interface with buttons for the entire workflow.
No terminal commands needed!
"""

import sys
import os
import json
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from pathlib import Path
import threading

# Ensure UTF-8 encoding for output
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except:
        pass


class WorkflowGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("JSON-to-DOCX Workflow Launcher")
        self.root.geometry("600x700")
        self.root.resizable(False, False)
        
        # Get project root
        self.project_root = Path(__file__).resolve().parent
        self.venv_path = self.project_root / ".venv"
        self.python_exe = self._get_python_exe()
        
        # Setup GUI
        self._setup_ui()
        self._check_venv()
    
    def _get_python_exe(self):
        """Get Python executable from .venv"""
        if sys.platform == "win32":
            exe = self.venv_path / "Scripts" / "python.exe"
        else:
            exe = self.venv_path / "bin" / "python"
        return exe
    
    def _check_venv(self):
        """Check if .venv exists and update status"""
        if self.venv_path.exists() and self.python_exe.exists():
            self.status_label.config(text="✓ Environment Ready", fg="green")
            self._enable_generation()
        else:
            self.status_label.config(text="✗ Setup Required - Click 'Run Setup'", fg="red")
            self._disable_generation()
    
    def _setup_ui(self):
        """Build the GUI"""
        
        # Header
        header = ttk.Label(
            self.root,
            text="JSON-to-DOCX Workflow",
            font=("Arial", 16, "bold")
        )
        header.pack(pady=10)
        
        # Status
        self.status_label = ttk.Label(
            self.root,
            text="Checking setup...",
            font=("Arial", 10)
        )
        self.status_label.pack(pady=5)
        
        # Section 1: Setup
        setup_frame = ttk.LabelFrame(self.root, text="1. SETUP (Run Once)", padding=10)
        setup_frame.pack(fill="x", padx=10, pady=10)
        
        ttk.Button(
            setup_frame,
            text="Run Setup",
            command=self.run_setup,
            width=30
        ).pack(pady=5)
        
        ttk.Label(
            setup_frame,
            text="Creates virtual environment and installs packages",
            font=("Arial", 9),
            foreground="gray"
        ).pack()
        
        # Section 2: Data Source
        data_frame = ttk.LabelFrame(self.root, text="2. SELECT STUDENT DATA", padding=10)
        data_frame.pack(fill="x", padx=10, pady=10)
        
        ttk.Label(data_frame, text="JSON File:").pack(side="left", padx=5)
        self.json_var = tk.StringVar(value="e_dnevnik_unified_state_v7.json")
        self.json_combo = ttk.Combobox(
            data_frame,
            textvariable=self.json_var,
            width=40,
            state="readonly"
        )
        self.json_combo.pack(side="left", padx=5)
        self._load_json_files()
        
        # Section 3: Template
        template_frame = ttk.LabelFrame(self.root, text="3. SELECT TEMPLATE", padding=10)
        template_frame.pack(fill="x", padx=10, pady=10)
        
        ttk.Label(template_frame, text="Template:").pack(side="left", padx=5)
        self.template_var = tk.StringVar(value="template.docx")
        template_combo = ttk.Combobox(
            template_frame,
            textvariable=self.template_var,
            width=40,
            state="readonly"
        )
        template_combo.pack(side="left", padx=5)
        self._load_templates()
        
        # Section 4: Actions
        action_frame = ttk.LabelFrame(self.root, text="4. GENERATE DOCUMENTS", padding=10)
        action_frame.pack(fill="x", padx=10, pady=10)
        
        button_frame = ttk.Frame(action_frame)
        button_frame.pack(fill="x", pady=5)
        
        ttk.Button(
            button_frame,
            text="List Students",
            command=self.list_students,
            width=18
        ).pack(side="left", padx=5)
        
        ttk.Button(
            button_frame,
            text="Generate All",
            command=self.generate_all,
            width=18
        ).pack(side="left", padx=5)
        
        ttk.Button(
            button_frame,
            text="Generate One...",
            command=self.generate_one,
            width=18
        ).pack(side="left", padx=5)
        
        # Section 5: Results
        results_frame = ttk.LabelFrame(self.root, text="5. OUTPUT", padding=10)
        results_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        ttk.Label(results_frame, text="Documents created in: OUT/").pack(anchor="w", pady=5)
        
        # Output text
        self.output_text = tk.Text(
            results_frame,
            height=10,
            width=70,
            font=("Courier", 9),
            state="disabled"
        )
        self.output_text.pack(fill="both", expand=True)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(self.output_text)
        scrollbar.pack(side="right", fill="y")
        self.output_text.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.output_text.yview)
        
        # Section 6: Footer buttons
        footer_frame = ttk.Frame(self.root)
        footer_frame.pack(fill="x", padx=10, pady=10)
        
        ttk.Button(
            footer_frame,
            text="Open Output Folder (OUT/)",
            command=self.open_output,
            width=28
        ).pack(side="left", padx=5)
        
        ttk.Button(
            footer_frame,
            text="Exit",
            command=self.root.quit,
            width=10
        ).pack(side="right", padx=5)
    
    def _load_json_files(self):
        """Load JSON files from current directory"""
        json_files = list(self.project_root.glob("e_dnevnik_unified_state*.json"))
        json_names = [f.name for f in json_files]
        if json_names:
            self.json_combo["values"] = json_names
    
    def _load_templates(self):
        """Load DOCX templates from current directory"""
        docx_files = list(self.project_root.glob("*.docx"))
        docx_names = [f.name for f in docx_files]
        if docx_names:
            self.json_combo.master.nametowidget(
                self.json_combo.master.winfo_children()[-1]
            )
    
    def _enable_generation(self):
        """Enable generation buttons"""
        for widget in self.root.winfo_children():
            if isinstance(widget, ttk.Frame):
                for child in widget.winfo_children():
                    if isinstance(child, ttk.Button):
                        child.config(state="normal")
    
    def _disable_generation(self):
        """Disable generation buttons"""
        for widget in self.root.winfo_children():
            if isinstance(widget, ttk.LabelFrame):
                for child in widget.winfo_children():
                    if isinstance(child, ttk.Button) and "Setup" not in child.cget("text"):
                        child.config(state="disabled")
    
    def _print_output(self, message):
        """Print message to output text"""
        self.output_text.config(state="normal")
        self.output_text.insert("end", message + "\n")
        self.output_text.see("end")
        self.output_text.config(state="disabled")
        self.root.update()
    
    def run_setup(self):
        """Run setup.py"""
        self._print_output("=" * 60)
        self._print_output("Running setup... (this may take 1-2 minutes)")
        self._print_output("=" * 60)
        
        def _run():
            try:
                cmd = [sys.executable, str(self.project_root / "setup.py"), "--no-prompt"]
                proc = subprocess.run(
                    cmd,
                    capture_output=False,
                    text=True,
                    cwd=str(self.project_root)
                )
                if proc.returncode == 0:
                    self._print_output("\n✓ Setup completed successfully!")
                    self._check_venv()
                else:
                    self._print_output(f"\n✗ Setup failed with code {proc.returncode}")
            except Exception as e:
                self._print_output(f"\n✗ Error: {e}")
        
        thread = threading.Thread(target=_run, daemon=True)
        thread.start()
    
    def list_students(self):
        """List all students"""
        json_file = self.project_root / self.json_var.get()
        
        if not json_file.exists():
            messagebox.showerror("Error", f"JSON file not found: {json_file}")
            return
        
        self._print_output("\n" + "=" * 60)
        self._print_output("STUDENT LIST")
        self._print_output("=" * 60)
        
        def _run():
            try:
                cmd = [
                    str(self.python_exe),
                    str(self.project_root / "compact_factory.py"),
                    "list-students",
                    "--unified", str(json_file)
                ]
                proc = subprocess.run(cmd, capture_output=True, text=True)
                if proc.returncode == 0:
                    self._print_output(proc.stdout)
                else:
                    self._print_output(f"✗ Error: {proc.stderr}")
            except Exception as e:
                self._print_output(f"✗ Exception: {e}")
        
        thread = threading.Thread(target=_run, daemon=True)
        thread.start()
    
    def generate_all(self):
        """Generate documents for all students"""
        json_file = self.project_root / self.json_var.get()
        template = self.project_root / self.template_var.get()
        
        if not json_file.exists():
            messagebox.showerror("Error", f"JSON not found: {json_file}")
            return
        if not template.exists():
            messagebox.showerror("Error", f"Template not found: {template}")
            return
        
        self._print_output("\n" + "=" * 60)
        self._print_output("GENERATING DOCUMENTS FOR ALL STUDENTS...")
        self._print_output("=" * 60)
        self._print_output(f"JSON: {json_file.name}")
        self._print_output(f"Template: {template.name}")
        self._print_output("Please wait...")
        
        def _run():
            try:
                out_dir = self.project_root / "OUT"
                cmd = [
                    str(self.python_exe),
                    str(self.project_root / "compact_factory.py"),
                    "gen-template",
                    "--unified", str(json_file),
                    "--template", str(template),
                    "--out", str(out_dir)
                ]
                proc = subprocess.run(cmd, capture_output=True, text=True)
                if proc.returncode == 0:
                    self._print_output("\n✓ Generation complete!")
                    self._print_output(f"\nDocuments created in: OUT/")
                    # List created files
                    for doc in out_dir.glob("*.docx"):
                        self._print_output(f"  ✓ {doc.name}")
                else:
                    self._print_output(f"\n✗ Error: {proc.stderr}")
            except Exception as e:
                self._print_output(f"\n✗ Exception: {e}")
        
        thread = threading.Thread(target=_run, daemon=True)
        thread.start()
    
    def generate_one(self):
        """Generate document for one student"""
        # Create dialog to select student
        def list_and_select():
            json_file = self.project_root / self.json_var.get()
            try:
                with open(json_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                
                # Extract students
                students = []
                if isinstance(data, dict):
                    for key in ("students", "student_records", "studentRecords"):
                        if key in data and isinstance(data[key], list):
                            students = data[key]
                            break
                
                # Get names
                names = []
                for s in students:
                    if isinstance(s, dict):
                        fn = s.get("firstName", "").strip()
                        ln = s.get("lastName", "").strip()
                        name = f"{fn} {ln}".strip()
                        if name:
                            names.append(name)
                
                return names
            except:
                return []
        
        names = list_and_select()
        if not names:
            messagebox.showerror("Error", "Could not read students from JSON")
            return
        
        # Create selection dialog
        dialog = tk.Toplevel(self.root)
        dialog.title("Select Student")
        dialog.geometry("400x300")
        
        ttk.Label(dialog, text="Select a student:", font=("Arial", 10)).pack(padx=10, pady=10)
        
        listbox = tk.Listbox(dialog, font=("Arial", 10))
        listbox.pack(fill="both", expand=True, padx=10, pady=5)
        
        for name in sorted(names):
            listbox.insert("end", name)
        
        def proceed():
            selection = listbox.curselection()
            if not selection:
                messagebox.showwarning("Warning", "Please select a student")
                return
            
            student_name = listbox.get(selection[0])
            dialog.destroy()
            self._do_generate_one(student_name)
        
        ttk.Button(dialog, text="Generate", command=proceed).pack(pady=10)
    
    def _do_generate_one(self, student_name):
        """Actually generate for one student"""
        json_file = self.project_root / self.json_var.get()
        template = self.project_root / self.template_var.get()
        
        self._print_output("\n" + "=" * 60)
        self._print_output(f"GENERATING DOCUMENT FOR: {student_name}")
        self._print_output("=" * 60)
        
        def _run():
            try:
                out_dir = self.project_root / "OUT"
                cmd = [
                    str(self.python_exe),
                    str(self.project_root / "compact_factory.py"),
                    "gen-template",
                    "--unified", str(json_file),
                    "--template", str(template),
                    "--out", str(out_dir),
                    "--student", student_name
                ]
                proc = subprocess.run(cmd, capture_output=True, text=True)
                if proc.returncode == 0:
                    self._print_output("\n✓ Generation complete!")
                else:
                    self._print_output(f"\n✗ Error: {proc.stderr}")
            except Exception as e:
                self._print_output(f"\n✗ Exception: {e}")
        
        thread = threading.Thread(target=_run, daemon=True)
        thread.start()
    
    def open_output(self):
        """Open OUT folder in file explorer"""
        out_dir = self.project_root / "OUT"
        
        if not out_dir.exists():
            messagebox.showinfo("Info", "OUT folder doesn't exist yet.\nGenerate some documents first!")
            return
        
        try:
            if sys.platform == "win32":
                os.startfile(out_dir)
            elif sys.platform == "darwin":
                subprocess.run(["open", str(out_dir)])
            else:
                subprocess.run(["xdg-open", str(out_dir)])
        except Exception as e:
            messagebox.showerror("Error", f"Could not open folder: {e}")


def main():
    root = tk.Tk()
    app = WorkflowGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()

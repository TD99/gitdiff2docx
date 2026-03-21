#!/usr/bin/env python3
"""
GitDiff2Docx GUI Application

A graphical user interface for GitDiff2Docx that provides an easy-to-use
interface for converting git diffs to Word documents.
"""

import sys
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import json

# Add the parent directory to the path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from gitdiff2docx.config.constants import SCRIPT_DIR
from gitdiff2docx.git.operations import get_first_commit, get_head_commit
from gitdiff2docx.theme.manager import list_available_themes
from gitdiff2docx.main import main as cli_main


class GitDiff2DocxGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("GitDiff2Docx - Git Diff to Word Document Converter")
        self.root.geometry("800x700")
        self.root.minsize(700, 600)
        
        # Load configuration
        self.config_path = os.path.join(SCRIPT_DIR, "config.json")
        self.load_config()
        
        # Create UI
        self.create_widgets()
        
    def load_config(self):
        """Load configuration from config.json"""
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                self.config = json.load(f)
        except:
            self.config = {}
    
    def create_widgets(self):
        """Create all GUI widgets"""
        # Main container with padding
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Title
        title = ttk.Label(main_frame, text="GitDiff2Docx", font=("Arial", 18, "bold"))
        title.grid(row=0, column=0, columnspan=3, pady=(0, 10))
        
        subtitle = ttk.Label(main_frame, text="Convert Git Diffs to Word Documents", font=("Arial", 10))
        subtitle.grid(row=1, column=0, columnspan=3, pady=(0, 20))
        
        # Target Directory
        ttk.Label(main_frame, text="Target Git Directory:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.target_dir_var = tk.StringVar(value=os.getcwd())
        ttk.Entry(main_frame, textvariable=self.target_dir_var).grid(row=2, column=1, sticky=(tk.W, tk.E), pady=5)
        ttk.Button(main_frame, text="Browse...", command=self.browse_target_dir).grid(row=2, column=2, padx=(5, 0), pady=5)
        
        # First Commit
        ttk.Label(main_frame, text="First Commit:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.commit1_var = tk.StringVar()
        ttk.Entry(main_frame, textvariable=self.commit1_var).grid(row=3, column=1, sticky=(tk.W, tk.E), pady=5)
        ttk.Button(main_frame, text="Auto", command=self.auto_commit1).grid(row=3, column=2, padx=(5, 0), pady=5)
        
        # Last Commit
        ttk.Label(main_frame, text="Last Commit:").grid(row=4, column=0, sticky=tk.W, pady=5)
        self.commit2_var = tk.StringVar()
        ttk.Entry(main_frame, textvariable=self.commit2_var).grid(row=4, column=1, sticky=(tk.W, tk.E), pady=5)
        ttk.Button(main_frame, text="Auto (HEAD)", command=self.auto_commit2).grid(row=4, column=2, padx=(5, 0), pady=5)
        
        # Output File
        ttk.Label(main_frame, text="Output File:").grid(row=5, column=0, sticky=tk.W, pady=5)
        self.output_var = tk.StringVar(value=os.path.join(SCRIPT_DIR, "output.docx"))
        ttk.Entry(main_frame, textvariable=self.output_var).grid(row=5, column=1, sticky=(tk.W, tk.E), pady=5)
        ttk.Button(main_frame, text="Save As...", command=self.browse_output).grid(row=5, column=2, padx=(5, 0), pady=5)
        
        # Options Frame
        options_frame = ttk.LabelFrame(main_frame, text="Options", padding="10")
        options_frame.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        options_frame.columnconfigure(1, weight=1)
        
        # Theme Selection
        ttk.Label(options_frame, text="Theme:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.theme_var = tk.StringVar(value=self.config.get("theme", "classic"))
        themes = self.get_available_themes()
        theme_combo = ttk.Combobox(options_frame, textvariable=self.theme_var, values=themes, state="readonly")
        theme_combo.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=5, padx=(5, 0))
        
        # Language Selection
        ttk.Label(options_frame, text="Language:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.language_var = tk.StringVar(value=self.config.get("language", "en"))
        lang_combo = ttk.Combobox(options_frame, textvariable=self.language_var, values=["en", "de"], state="readonly")
        lang_combo.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5, padx=(5, 0))
        
        # Verbose Mode
        self.verbose_var = tk.BooleanVar(value=self.config.get("verbose", False))
        ttk.Checkbutton(options_frame, text="Verbose output", variable=self.verbose_var).grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=5)
        
        # Include unchanged lines
        self.include_unchanged_var = tk.BooleanVar(value=self.config.get("include_unchanged_lines", True))
        ttk.Checkbutton(options_frame, text="Include unchanged lines", variable=self.include_unchanged_var).grid(row=3, column=0, columnspan=2, sticky=tk.W, pady=5)
        
        # Open after creation
        self.open_after_var = tk.BooleanVar(value=self.config.get("open_after_creation", False))
        ttk.Checkbutton(options_frame, text="Open document after creation", variable=self.open_after_var).grid(row=4, column=0, columnspan=2, sticky=tk.W, pady=5)
        
        # Generate Button
        generate_btn = ttk.Button(main_frame, text="Generate Document", command=self.generate_document)
        generate_btn.grid(row=7, column=0, columnspan=3, pady=20)
        
        # Progress/Log Frame
        log_frame = ttk.LabelFrame(main_frame, text="Output Log", padding="10")
        log_frame.grid(row=8, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        main_frame.rowconfigure(8, weight=1)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=10, state='disabled')
        self.log_text.pack(fill=tk.BOTH, expand=True)
    
    def get_available_themes(self):
        """Get list of available themes"""
        try:
            themes_dir = os.path.join(SCRIPT_DIR, "themes")
            return list_available_themes(themes_dir, excluded_filenames=["_overrides.json"])
        except:
            return ["classic"]
    
    def browse_target_dir(self):
        """Browse for target directory"""
        directory = filedialog.askdirectory(initialdir=self.target_dir_var.get())
        if directory:
            self.target_dir_var.set(directory)
    
    def browse_output(self):
        """Browse for output file"""
        filename = filedialog.asksaveasfilename(
            initialdir=os.path.dirname(self.output_var.get()),
            defaultextension=".docx",
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")]
        )
        if filename:
            self.output_var.set(filename)
    
    def auto_commit1(self):
        """Auto-fill first commit"""
        try:
            target = self.target_dir_var.get()
            if os.path.isdir(target):
                os.chdir(target)
                commit = get_first_commit()
                self.commit1_var.set(commit)
                self.log_message(f"Auto-filled first commit: {commit}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to get first commit: {e}")
    
    def auto_commit2(self):
        """Auto-fill last commit (HEAD)"""
        try:
            target = self.target_dir_var.get()
            if os.path.isdir(target):
                os.chdir(target)
                commit = get_head_commit()
                self.commit2_var.set(commit)
                self.log_message(f"Auto-filled last commit (HEAD): {commit}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to get HEAD commit: {e}")
    
    def log_message(self, message):
        """Add a message to the log"""
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, message + '\n')
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')
        self.root.update()
    
    def clear_log(self):
        """Clear the log"""
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state='disabled')
    
    def generate_document(self):
        """Generate the document in a separate thread"""
        # Validate inputs
        target_dir = self.target_dir_var.get()
        if not os.path.isdir(target_dir):
            messagebox.showerror("Error", "Please select a valid target directory")
            return
        
        output_path = self.output_var.get()
        if not output_path:
            messagebox.showerror("Error", "Please specify an output file path")
            return
        
        # Clear log
        self.clear_log()
        
        # Update config with current values
        temp_config_path = os.path.join(SCRIPT_DIR, "temp_config.json")
        config = self.config.copy()
        config["theme"] = self.theme_var.get()
        config["language"] = self.language_var.get()
        config["verbose"] = self.verbose_var.get()
        config["include_unchanged_lines"] = self.include_unchanged_var.get()
        config["open_after_creation"] = self.open_after_var.get()
        
        with open(temp_config_path, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=4)
        
        # Prepare CLI arguments
        args = [
            '--target-dir', target_dir,
            '--output', output_path,
            '--config-file', temp_config_path,
        ]
        
        if self.commit1_var.get():
            args.extend(['--commit1', self.commit1_var.get()])
        
        if self.commit2_var.get():
            args.extend(['--commit2', self.commit2_var.get()])
        
        # Run in a separate thread to keep GUI responsive
        thread = threading.Thread(target=self.run_cli, args=(args,))
        thread.daemon = True
        thread.start()
    
    def run_cli(self, args):
        """Run the CLI in a separate thread"""
        try:
            self.log_message("Starting document generation...")
            self.log_message(f"Target directory: {self.target_dir_var.get()}")
            self.log_message(f"Output file: {self.output_var.get()}")
            
            # Redirect stdout to capture CLI output
            import io
            from contextlib import redirect_stdout
            
            f = io.StringIO()
            with redirect_stdout(f):
                # Import click's standalone mode context manager
                import click
                try:
                    with click.Context(cli_main) as ctx:
                        ctx.params = {}
                        # Parse args manually
                        i = 0
                        while i < len(args):
                            if args[i] == '--target-dir':
                                ctx.params['target_dir'] = args[i+1]
                                i += 2
                            elif args[i] == '--output':
                                ctx.params['output'] = args[i+1]
                                i += 2
                            elif args[i] == '--config-file':
                                ctx.params['config_file'] = args[i+1]
                                i += 2
                            elif args[i] == '--commit1':
                                ctx.params['commit1'] = args[i+1]
                                i += 2
                            elif args[i] == '--commit2':
                                ctx.params['commit2'] = args[i+1]
                                i += 2
                            else:
                                i += 1
                        
                        # Fill in defaults
                        ctx.params.setdefault('create_theme', False)
                        ctx.params.setdefault('verbose', None)
                        ctx.params.setdefault('theme', None)
                        ctx.params.setdefault('language', None)
                        
                        result = cli_main.invoke(ctx)
                except SystemExit:
                    pass
            
            output = f.getvalue()
            if output:
                self.log_message(output)
            
            self.root.after(0, lambda: messagebox.showinfo("Success", "Document generated successfully!"))
            
        except Exception as e:
            error_msg = str(e)
            self.log_message(f"Error: {error_msg}")
            self.root.after(0, lambda: messagebox.showerror("Error", f"Failed to generate document: {error_msg}"))


def main():
    """Main entry point for the GUI"""
    root = tk.Tk()
    app = GitDiff2DocxGUI(root)
    root.mainloop()


if __name__ == '__main__':
    main()

import os
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox
from ttkthemes import ThemedTk
import sv_ttk
from threading import Thread

class UnusedAppRemover:
    def __init__(self, root):
        self.root = root
        self.root.title("Unused App Remover")
        self.root.geometry("950x700")
        
        sv_ttk.set_theme("dark")

        self.create_ui()

    def create_ui(self):
        # Title bar
        title_frame = tk.Frame(self.root, bg="#2b2b2b")
        title_frame.pack(fill=tk.X, pady=10)
        title_label = tk.Label(
            title_frame,
            text="Unused App Remover",
            font=("Segoe UI", 20, "bold"),
            bg="#2b2b2b",
            fg="#56b6c2",
        )
        title_label.pack(pady=10)

        # Description
        desc_label = tk.Label(
            self.root,
            text="Analyze and remove unused software to free up space and optimize your system.",
            font=("Segoe UI", 12),
            bg="#2b2b2b",
            fg="#abb2bf",
        )
        desc_label.pack(pady=5)

        # Analyze button
        self.analyze_button = ttk.Button(
            self.root, text="Analyze Installed Applications", command=self.analyze_apps, style="Accent.TButton"
        )
        self.analyze_button.pack(pady=20)

        # Table for apps
        self.tree_frame = ttk.Frame(self.root)
        self.tree_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        self.tree = ttk.Treeview(
            self.tree_frame, columns=("App Name", "Install Date", "Size"), show="headings", height=20
        )
        self.tree.heading("App Name", text="App Name", anchor=tk.W)
        self.tree.heading("Install Date", text="Install Date", anchor=tk.CENTER)
        self.tree.heading("Size", text="Size (MB)", anchor=tk.CENTER)

        self.tree.column("App Name", width=500, anchor=tk.W)
        self.tree.column("Install Date", width=150, anchor=tk.CENTER)
        self.tree.column("Size", width=120, anchor=tk.CENTER)

        self.tree.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        tree_scroll = ttk.Scrollbar(self.tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=tree_scroll.set)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # Button frame
        button_frame = tk.Frame(self.root, bg="#2b2b2b")
        button_frame.pack(fill=tk.X, pady=10)

        self.uninstall_button = ttk.Button(
            button_frame, text="Uninstall Selected App", command=self.uninstall_app, style="Accent.TButton"
        )
        self.uninstall_button.pack(side=tk.RIGHT, padx=20, pady=10)

    def analyze_apps(self):
        self.tree.delete(*self.tree.get_children())
        self.analyze_button.config(state=tk.DISABLED, text="Analyzing...")

        def fetch_installed_apps():
            try:
                apps = subprocess.check_output(
                    "wmic product get name,installlocation,installdate", shell=True, universal_newlines=True
                )
                for line in apps.splitlines()[2:]:
                    if line.strip():
                        parts = line.split()
                        app_name = " ".join(parts[:-2])
                        install_date = parts[-2] if parts[-2].isdigit() else "Unknown"
                        size = "N/A"

                        self.tree.insert("", tk.END, values=(app_name, install_date, size))

            except Exception as e:
                messagebox.showerror("Error", f"Failed to analyze apps: {e}")
            finally:
                self.analyze_button.config(state=tk.NORMAL, text="Analyze Installed Applications")

        Thread(target=fetch_installed_apps).start()

    def uninstall_app(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("No Selection", "Please select an app to uninstall.")
            return

        app_name = self.tree.item(selected_item, "values")[0]

        if messagebox.askyesno("Confirm Uninstallation", f"Are you sure you want to uninstall {app_name}?"):
            try:
                uninstall_cmd = f"wmic product where name=\"{app_name}\" call uninstall"
                subprocess.run(uninstall_cmd, shell=True, check=True)
                self.tree.delete(selected_item)
                messagebox.showinfo("Success", f"{app_name} has been uninstalled.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to uninstall app: {e}")

if __name__ == "__main__":
    root = ThemedTk(theme="equilux")
    app = UnusedAppRemover(root)
    root.mainloop()

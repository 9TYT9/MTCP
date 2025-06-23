import tkinter as tk
from datetime import datetime
import threading
from tkinter import ttk, messagebox, filedialog
from pymodbus.client import ModbusTcpClient
import json
import os
import csv
import re
import time

# File paths for storing configurations
PLC_CONFIG_FILE = "plc_configs.json"
TRACEABILITY_CONFIG_FILE = "traceability_configs.json"
ERROR_CODE_CONFIG_FILE = "error_code_configs.json"
DOWN_TIME_CONFIG_FILE = "down_time_configs.json"


class PLCManagerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PLC Data Manager")
        self.root.geometry("600x700")  # Window size

        # Make the window fixed in size
        self.root.resizable(False, False)

        # Create a ttk.Style to handle tab colors
        self.style = ttk.Style()

        # Global Widget Styles
        self.style.configure("TNotebook", background="white")  # Background of Notebook
        self.style.configure("TNotebook.Tab", padding=[10, 5])  # Default tab padding
        self.style.configure("TFrame", background="white")  # Default tab background color

        self.style.map("TNotebook.Tab", background=[("selected", "lightblue"), ("!selected", "white")])

        # Create Notebook (tab container)
        self.tab_control = ttk.Notebook(root, style="TNotebook")

        # PLC Configuration Tab
        self.plc_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(self.plc_tab, text="PLC Configuration")
        self.initialize_plc_tab()

        # TRACEABILITY Tab
        self.traceability_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(self.traceability_tab, text="TRACEABILITY")
        self.initialize_traceability_tab()

        # ERROR_CODE Tab
        self.error_code_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(self.error_code_tab, text="ERROR_CODE")
        self.initialize_error_code_tab()

        # DOWN_TIME Tab
        self.down_time_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(self.down_time_tab, text="DOWN_TIME")
        self.initialize_down_time_tab()

        self.tab_control.pack(expand=1, fill="both")

        self.plc_configs = []
        self.traceability_configs = []
        self.error_code_configs = []
        self.down_time_configs = []

        self.monitoring = False
        self.selected_plc_index = None
        self.selected_traceability_index = None
        self.selected_error_code_index = None
        self.selected_down_time_index = None

        self.load_plc_configs()
        self.load_traceability_configs()
        self.load_error_code_configs()
        self.load_down_time_configs()

        # Monitoring control buttons
        self.running = False

        # Monitoring control buttons
        self.run_button = tk.Button(self.root, text="Run", command=self.start_monitoring, bg="white", fg="black")
        self.run_button.pack(side="left", padx=10, pady=10)

        self.stop_button = tk.Button(self.root, text="Stop", command=self.stop_monitoring, bg="white", fg="black")
        self.stop_button.pack(side="left", padx=10, pady=10)

        # Textbox output for monitoring logs
        self.output_text = tk.Text(self.root, height=5, width=80)
        self.output_text.pack(side="bottom", padx=5, pady=5)

        # Save data and clean up on close
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def on_tab_change(self, event):
        """
        Change the background color dynamically for the selected and non-selected tabs.
        """
        self.apply_tab_colors()

    def apply_tab_colors(self):
        """
        Apply background colors for all tabs.
        Highlight the currently selected tab and reset others.
        """
        # Colors for tabs
        selected_tab_color = "lightblue"
        non_selected_tab_color = "lightgray"

        # Loop through all tabs
        for index in range(len(self.tab_control.tabs())):
            tab_id = self.tab_control.tabs()[index]  # Get tab widget ID
            tab_widget = self.tab_control.nametowidget(tab_id)  # Convert ID to a widget

            # Change color based on whether the tab is selected
            if index == self.tab_control.index("current"):
                tab_widget.configure(background=selected_tab_color)  # Selected tab color
            else:
                tab_widget.configure(background=non_selected_tab_color)  # Non-selected tab color

    def initialize_plc_tab(self):
        """Initialize PLC Configuration Tab."""
        ttk.Label(self.plc_tab, text="Line Name:").grid(row=0, column=0, padx=5, pady=5)
        self.line_entry = ttk.Entry(self.plc_tab)
        self.line_entry.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(self.plc_tab, text="Equipment Name:").grid(row=1, column=0, padx=5, pady=5)
        self.equipment_entry = ttk.Entry(self.plc_tab)
        self.equipment_entry.grid(row=1, column=1, padx=5, pady=5)

        ttk.Label(self.plc_tab, text="PLC IP Address:").grid(row=2, column=0, padx=5, pady=5)
        self.ip_entry = ttk.Entry(self.plc_tab)
        self.ip_entry.grid(row=2, column=1, padx=5, pady=5)

        ttk.Label(self.plc_tab, text="PLC Port:").grid(row=3, column=0, padx=5, pady=5)
        self.port_entry = ttk.Entry(self.plc_tab)
        self.port_entry.grid(row=3, column=1, padx=5, pady=5)

        # Buttons for Add, Edit, Save, and Delete
        ttk.Button(self.plc_tab, text="Add PLC", command=self.add_plc_config).grid(row=4, column=0, padx=5, pady=5)
        ttk.Button(self.plc_tab, text="Edit PLC", command=self.edit_plc_config).grid(row=4, column=1, padx=5, pady=5)
        ttk.Button(self.plc_tab, text="Save PLC", command=self.save_plc_config).grid(row=4, column=2, padx=5, pady=5)
        ttk.Button(self.plc_tab, text="Delete PLC", command=self.delete_plc_config).grid(row=4, column=3, padx=5, pady=5)

        # Listbox to display PLC configurations
        self.plc_list = tk.Listbox(self.plc_tab, height=23, width=100)
        self.plc_list.grid(row=5, column=0, columnspan=4, padx=10, pady=10)

    def initialize_traceability_tab(self):
        """Initialize TRACEABILITY Tab."""
        ttk.Label(self.traceability_tab, text="File Name:").grid(row=0, column=0, padx=5, pady=5)
        self.file_name_entry1 = ttk.Entry(self.traceability_tab)
        self.file_name_entry1.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(self.traceability_tab, text="Register Type:").grid(row=1, column=0, padx=5, pady=5)
        self.reg_type_combobox1 = ttk.Combobox(self.traceability_tab, values=["Coil", "Discrete", "Holding", "Input"])
        self.reg_type_combobox1.grid(row=1, column=1, padx=5, pady=5)

        ttk.Label(self.traceability_tab, text="Start Register:").grid(row=2, column=0, padx=5, pady=5)
        self.start_reg_entry1 = ttk.Entry(self.traceability_tab)
        self.start_reg_entry1.grid(row=2, column=1, padx=5, pady=5)

        ttk.Label(self.traceability_tab, text="Register Range:").grid(row=3, column=0, padx=5, pady=5)
        self.range_entry1 = ttk.Entry(self.traceability_tab)
        self.range_entry1.grid(row=3, column=1, padx=5, pady=5)

        ttk.Label(self.traceability_tab, text="Trigger Register Type:").grid(row=4, column=0, padx=5, pady=5)
        self.trigger_type_combobox1 = ttk.Combobox(self.traceability_tab, values=["Coil", "Discrete", "Holding", "Input"])
        self.trigger_type_combobox1.grid(row=4, column=1, padx=5, pady=5)

        ttk.Label(self.traceability_tab, text="Trigger Register:").grid(row=5, column=0, padx=5, pady=5)
        self.trigger_entry1 = ttk.Entry(self.traceability_tab)
        self.trigger_entry1.grid(row=5, column=1, padx=5, pady=5)

        # Label for folder path in TRACEABILITY tab
        ttk.Label(self.traceability_tab, text="Folder Path:").grid(row=6, column=0, padx=5, pady=5)

        # Single Entry widget for the folder path
        self.traceability_folder_path_entry = ttk.Entry(self.traceability_tab, width=20)
        self.traceability_folder_path_entry.grid(row=6, column=1, padx=5, pady=5)

        # Browse Button
        ttk.Button(self.traceability_tab, text="Browse", command=lambda: self.browse_folder_path(self.traceability_folder_path_entry)).grid(row=6, column=2, padx=5, pady=5)

        # Buttons for Add, Edit, Save, and Delete
        ttk.Button(self.traceability_tab, text="Add Output", command=self.add_traceability_config).grid(row=7, column=0, padx=5, pady=5)
        ttk.Button(self.traceability_tab, text="Edit Output", command=self.edit_traceability_config).grid(row=7, column=1, padx=5, pady=5)
        ttk.Button(self.traceability_tab, text="Save Output", command=self.save_traceability_config).grid(row=7, column=2, padx=5, pady=5)
        ttk.Button(self.traceability_tab, text="Delete Output", command=self.delete_traceability_config).grid(row=7, column=3, padx=5, pady=5)

        # Listbox to display TRACEABILITY
        self.traceability_list = tk.Listbox(self.traceability_tab, height=16, width=100)
        self.traceability_list.grid(row=8, column=0, columnspan=4, padx=10, pady=10)

    def initialize_error_code_tab(self):
        """Initialize ERROR_CODE Tab."""
        ttk.Label(self.error_code_tab, text="File Name:").grid(row=0, column=0, padx=5, pady=5)
        self.file_name_entry2 = ttk.Entry(self.error_code_tab)
        self.file_name_entry2.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(self.error_code_tab, text="Register Type:").grid(row=1, column=0, padx=5, pady=5)
        self.reg_type_combobox2 = ttk.Combobox(self.error_code_tab, values=["Coil", "Discrete", "Holding", "Input"])
        self.reg_type_combobox2.grid(row=1, column=1, padx=5, pady=5)

        ttk.Label(self.error_code_tab, text="Start Register:").grid(row=2, column=0, padx=5, pady=5)
        self.start_reg_entry2 = ttk.Entry(self.error_code_tab)
        self.start_reg_entry2.grid(row=2, column=1, padx=5, pady=5)

        ttk.Label(self.error_code_tab, text="Register Range:").grid(row=3, column=0, padx=5, pady=5)
        self.range_entry2 = ttk.Entry(self.error_code_tab)
        self.range_entry2.grid(row=3, column=1, padx=5, pady=5)

        ttk.Label(self.error_code_tab, text="Trigger Register Type:").grid(row=4, column=0, padx=5, pady=5)
        self.trigger_type_combobox2 = ttk.Combobox(self.error_code_tab, values=["Coil", "Discrete", "Holding", "Input"])
        self.trigger_type_combobox2.grid(row=4, column=1, padx=5, pady=5)

        ttk.Label(self.error_code_tab, text="Trigger Register:").grid(row=5, column=0, padx=5, pady=5)
        self.trigger_entry2 = ttk.Entry(self.error_code_tab)
        self.trigger_entry2.grid(row=5, column=1, padx=5, pady=5)

        # Label for folder path in ERROR_CODE tab
        ttk.Label(self.error_code_tab, text="Folder Path:").grid(row=6, column=0, padx=5, pady=5)

        # Single Entry widget for the folder path
        self.error_code_folder_path_entry = ttk.Entry(self.error_code_tab, width=20)
        self.error_code_folder_path_entry.grid(row=6, column=1, padx=5, pady=5)

        # Browse Button
        ttk.Button(self.error_code_tab, text="Browse", command=lambda: self.browse_folder_path(self.error_code_folder_path_entry)).grid(row=6, column=2, padx=5, pady=5)

        # Buttons for Add, Edit, Save, and Delete
        ttk.Button(self.error_code_tab, text="Add Output", command=self.add_error_code_config).grid(row=7, column=0, padx=5, pady=5)
        ttk.Button(self.error_code_tab, text="Edit Output", command=self.edit_error_code_config).grid(row=7, column=1, padx=5, pady=5)
        ttk.Button(self.error_code_tab, text="Save Output", command=self.save_error_code_config).grid(row=7, column=2, padx=5, pady=5)
        ttk.Button(self.error_code_tab, text="Delete Output", command=self.delete_error_code_config).grid(row=7, column=3, padx=5, pady=5)

        # Listbox to display TRACEABILITY
        self.error_code_list = tk.Listbox(self.error_code_tab, height=16, width=100)
        self.error_code_list.grid(row=8, column=0, columnspan=4, padx=10, pady=10)

    def initialize_down_time_tab(self):
        """Initialize DOWN_TIME Tab."""
        ttk.Label(self.down_time_tab, text="File Name:").grid(row=0, column=0, padx=5, pady=5)
        self.file_name_entry3 = ttk.Entry(self.down_time_tab)
        self.file_name_entry3.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(self.down_time_tab, text="Register Type:").grid(row=1, column=0, padx=5, pady=5)
        self.reg_type_combobox3 = ttk.Combobox(self.down_time_tab, values=["Coil", "Discrete", "Holding", "Input"])
        self.reg_type_combobox3.grid(row=1, column=1, padx=5, pady=5)

        ttk.Label(self.down_time_tab, text="Start Register:").grid(row=2, column=0, padx=5, pady=5)
        self.start_reg_entry3 = ttk.Entry(self.down_time_tab)
        self.start_reg_entry3.grid(row=2, column=1, padx=5, pady=5)

        ttk.Label(self.down_time_tab, text="Register Range:").grid(row=3, column=0, padx=5, pady=5)
        self.range_entry3 = ttk.Entry(self.down_time_tab)
        self.range_entry3.grid(row=3, column=1, padx=5, pady=5)

        ttk.Label(self.down_time_tab, text="Trigger Register Type:").grid(row=4, column=0, padx=5, pady=5)
        self.trigger_type_combobox3 = ttk.Combobox(self.down_time_tab, values=["Coil", "Discrete", "Holding", "Input"])
        self.trigger_type_combobox3.grid(row=4, column=1, padx=5, pady=5)

        ttk.Label(self.down_time_tab, text="Trigger Register:").grid(row=5, column=0, padx=5, pady=5)
        self.trigger_entry3 = ttk.Entry(self.down_time_tab)
        self.trigger_entry3.grid(row=5, column=1, padx=5, pady=5)

        # Label for folder path in DOWN_TIME tab
        ttk.Label(self.down_time_tab, text="Folder Path:").grid(row=6, column=0, padx=5, pady=5)

        # Single Entry widget for the folder path
        self.down_time_folder_path_entry = ttk.Entry(self.down_time_tab, width=20)
        self.down_time_folder_path_entry.grid(row=6, column=1, padx=5, pady=5)

        # Browse Button
        ttk.Button(self.down_time_tab, text="Browse", command=lambda: self.browse_folder_path(self.down_time_folder_path_entry)).grid(row=6, column=2, padx=5, pady=5)

        # Buttons for Add, Edit, Save, and Delete
        ttk.Button(self.down_time_tab, text="Add Output", command=self.add_down_time_config).grid(row=7, column=0, padx=5, pady=5)
        ttk.Button(self.down_time_tab, text="Edit Output", command=self.edit_down_time_config).grid(row=7, column=1, padx=5, pady=5)
        ttk.Button(self.down_time_tab, text="Save Output", command=self.save_down_time_config).grid(row=7, column=2, padx=5, pady=5)
        ttk.Button(self.down_time_tab, text="Delete Output", command=self.delete_down_time_config).grid(row=7, column=3, padx=5, pady=5)

        # Listbox to display TRACEABILITY
        self.down_time_list = tk.Listbox(self.down_time_tab, height=16, width=100)
        self.down_time_list.grid(row=8, column=0, columnspan=4, padx=10, pady=10)

    def load_plc_configs(self):
        """Load PLC configurations from JSON file."""
        if os.path.exists(PLC_CONFIG_FILE):
            with open(PLC_CONFIG_FILE, "r") as file:
                self.plc_configs = json.load(file)
            self.refresh_plc_list()
        else:
            self.plc_configs = []

    def load_traceability_configs(self):
        """Load TRACEABILITY configurations from JSON file."""
        if os.path.exists(TRACEABILITY_CONFIG_FILE):
            with open(TRACEABILITY_CONFIG_FILE, "r") as file:
                self.traceability_configs = json.load(file)
            self.refresh_traceability_list()
        else:
            self.traceability_configs = []

    def load_error_code_configs(self):
        """Load ERROR_CODE configurations from JSON file."""
        if os.path.exists(ERROR_CODE_CONFIG_FILE):
            with open(ERROR_CODE_CONFIG_FILE, "r") as file:
                self.error_code_configs = json.load(file)
            self.refresh_error_code_list()
        else:
            self.error_code_configs = []

    def load_down_time_configs(self):
        """Load DOWN_TIME configurations from JSON file."""
        if os.path.exists(DOWN_TIME_CONFIG_FILE):
            with open(DOWN_TIME_CONFIG_FILE, "r") as file:
                self.down_time_configs = json.load(file)
            self.refresh_down_time_list()
        else:
            self.down_time_configs = []

    def save_all_configs(self):
        """Save all configurations to JSON files."""
        with open(PLC_CONFIG_FILE, "w") as file:
            json.dump(self.plc_configs, file, indent=4)

        with open(TRACEABILITY_CONFIG_FILE, "w") as file:
            json.dump(self.traceability_configs, file, indent=4)

        with open(ERROR_CODE_CONFIG_FILE, "w") as file:
            json.dump(self.error_code_configs, file, indent=4)

        with open(DOWN_TIME_CONFIG_FILE, "w") as file:
            json.dump(self.down_time_configs, file, indent=4)

    def refresh_plc_list(self):
        """Refresh the Listbox with updated PLC configurations."""
        self.plc_list.delete(0, tk.END)
        for i, plc in enumerate(self.plc_configs):
            status = plc.get("status", "Not Connected")
            self.plc_list.insert(tk.END, f"{i + 1}. {plc['line_name']} ({plc['ip_address']}:{plc['port']}) - {status}")

    def refresh_traceability_list(self):
        """Refresh the Listbox with updated TRACEABILITY configurations."""
        self.traceability_list.delete(0, tk.END)
        for i, traceability in enumerate(self.traceability_configs):
            self.traceability_list.insert(tk.END, f"{i + 1}. {traceability['file_name']} Type:{traceability['register_type']} Start:{traceability['start_register']} Range:{traceability.get('range')}")

    def refresh_error_code_list(self):
        """Refresh the Listbox with updated ERROR_CODE configurations."""
        self.error_code_list.delete(0, tk.END)
        for i, error_code in enumerate(self.error_code_configs):
            self.error_code_list.insert(tk.END, f"{i + 1}. {error_code['file_name']} Type:{error_code['register_type']} Start:{error_code['start_register']} Range:{error_code.get('range')}")

    def refresh_down_time_list(self):
        """Refresh the Listbox with updated DOWN_TIME configurations."""
        self.down_time_list.delete(0, tk.END)
        for i, down_time in enumerate(self.down_time_configs):
            self.down_time_list.insert(tk.END, f"{i + 1}. {down_time['file_name']} Type:{down_time['register_type']} Start:{down_time['start_register']} Range:{down_time.get('range')}")

    def browse_folder_path(self, target_entry):
        """Browse for folder path."""
        folder_path = filedialog.askdirectory(title="Select Folder")
        if folder_path:
            target_entry.delete(0, tk.END)  # Clear existing text
            target_entry.insert(0, folder_path)  # Insert the selected folder path

    def add_plc_config(self):
        """Add a new PLC configuration."""
        line_name = self.line_entry.get()
        equipment_name = self.equipment_entry.get()
        ip_address = self.ip_entry.get()
        port = self.port_entry.get()

        # Validate inputs
        if not line_name or not equipment_name or not ip_address or not port:
            messagebox.showwarning("Warning", "All fields must be filled.")
            return

        # Validate IP address format
        if not re.match(r'^\d{1,3}(\.\d{1,3}){3}$', ip_address):
            messagebox.showerror("Error", "Invalid IP address format.")
            return

        # Validate port range
        try:
            port = int(port)
            if not 1 <= port <= 65535:
                raise ValueError
        except ValueError:
            messagebox.showerror("Error", "Port must be a valid integer between 1 and 65535.")
            return

        plc_config = {
            "line_name": line_name,
            "equipment_name": equipment_name,
            "ip_address": ip_address,
            "port": port,
            "status": "Not Connected"
        }

        try:
            client = ModbusTcpClient(ip_address, port=port)
            connection = client.connect()
            plc_config["status"] = "Connected" if connection else "Failed"
        except Exception as e:
            plc_config["status"] = "Failed"
            print(f"Connection Error: {e}")
        finally:
            client.close()

        self.plc_configs.append(plc_config)
        self.refresh_plc_list()
        self.save_all_configs()
        messagebox.showinfo("Info", f"PLC Configuration added successfully!\nConnection Status: {plc_config['status']}")

    def add_traceability_config(self):
        """Add a new TRACEABILITY configuration."""
        file_name = self.file_name_entry1.get()
        register_type = self.reg_type_combobox1.get()
        start_register = self.start_reg_entry1.get()
        reg_range = self.range_entry1.get()
        trigger_register_type = self.trigger_type_combobox1.get()
        trigger_register = self.trigger_entry1.get()
        folder_path = self.traceability_folder_path_entry.get()

        # Validate that start_register and reg_range are integers
        try:
            start_register = int(start_register)
            reg_range = int(reg_range)
        except ValueError:
            messagebox.showerror("Error", "Start register and range must be valid integers.")
            return

        # Validate trigger_register is an integer
        try:
            trigger_register = int(trigger_register)
        except ValueError:
            messagebox.showerror("Error", "Trigger register must be a valid integer.")
            return

        # Validate folder_path is a valid path (optional, can add more checks)
        if not os.path.isdir(folder_path):
            messagebox.showerror("Error", f"Invalid folder path: {folder_path}")
            return

        # Create configuration dictionary
        traceability_config = {
            "file_name": file_name,
            "register_type": register_type,
            "start_register": start_register,
            "range": reg_range,
            "trigger_register_type": trigger_register_type,
            "trigger_register": trigger_register,
            "folder_path": folder_path,
        }

        # Append to traceability_configs list and refresh UI
        self.traceability_configs.append(traceability_config)
        self.refresh_traceability_list()
        self.save_all_configs()

        # Notify user
        messagebox.showinfo("Info", "TRACEABILITY Configuration added successfully!")

    def add_error_code_config(self):
        """Add a new ERROR_CODE configuration."""
        file_name = self.file_name_entry2.get()
        register_type = self.reg_type_combobox2.get()
        start_register = self.start_reg_entry2.get()
        reg_range = self.range_entry2.get()
        trigger_register_type = self.trigger_type_combobox2.get()
        trigger_register = self.trigger_entry2.get()
        folder_path = self.error_code_folder_path_entry.get()

        # Validate that start_register and reg_range are integers
        try:
            start_register = int(start_register)
            reg_range = int(reg_range)
        except ValueError:
            messagebox.showerror("Error", "Start register and range must be valid integers.")
            return

        # Validate trigger_register is an integer
        try:
            trigger_register = int(trigger_register)
        except ValueError:
            messagebox.showerror("Error", "Trigger register must be a valid integer.")
            return

        # Validate folder_path is a valid path (optional, can add more checks)
        if not os.path.isdir(folder_path):
            messagebox.showerror("Error", f"Invalid folder path: {folder_path}")
            return

        # Create configuration dictionary
        error_code_config = {
            "file_name": file_name,
            "register_type": register_type,
            "start_register": start_register,
            "range": reg_range,
            "trigger_register_type": trigger_register_type,
            "trigger_register": trigger_register,
            "folder_path": folder_path,
        }

        # Append to error_code_configs list and refresh UI
        self.error_code_configs.append(error_code_config)
        self.refresh_error_code_list()
        self.save_all_configs()

        # Notify user
        messagebox.showinfo("Info", "ERROR_CODE Configuration added successfully!")

    def add_down_time_config(self):
        """Add a new DOWN_TIME configuration."""
        file_name = self.file_name_entry3.get()
        register_type = self.reg_type_combobox3.get()
        start_register = self.start_reg_entry3.get()
        reg_range = self.range_entry3.get()
        trigger_register_type = self.trigger_type_combobox3.get()
        trigger_register = self.trigger_entry3.get()
        folder_path = self.down_time_folder_path_entry.get()

        # Validate that start_register and reg_range are integers
        try:
            start_register = int(start_register)
            reg_range = int(reg_range)
        except ValueError:
            messagebox.showerror("Error", "Start register and range must be valid integers.")
            return

        # Validate trigger_register is an integer
        try:
            trigger_register = int(trigger_register)
        except ValueError:
            messagebox.showerror("Error", "Trigger register must be a valid integer.")
            return

        # Validate folder_path is a valid path (optional, can add more checks)
        if not os.path.isdir(folder_path):
            messagebox.showerror("Error", f"Invalid folder path: {folder_path}")
            return

        # Create configuration dictionary
        down_time_config = {
            "file_name": file_name,
            "register_type": register_type,
            "start_register": start_register,
            "range": reg_range,
            "trigger_register_type": trigger_register_type,
            "trigger_register": trigger_register,
            "folder_path": folder_path,
        }

        # Append to traceability_configs list and refresh UI
        self.down_time_configs.append(down_time_config)
        self.refresh_down_time_list()
        self.save_all_configs()

        # Notify user
        messagebox.showinfo("Info", "DOWN_TIME Configuration added successfully!")

    def edit_plc_config(self):
        """Edit an existing PLC configuration."""
        selected_index = self.plc_list.curselection()
        if not selected_index:
            messagebox.showwarning("Warning", "Please select PLC configuration to edit.")
            return

        self.selected_plc_index = selected_index[0]
        plc_config = self.plc_configs[self.selected_plc_index]

        self.line_entry.delete(0, tk.END)
        self.line_entry.insert(0, plc_config["line_name"])
        self.equipment_entry.delete(0, tk.END)
        self.equipment_entry.insert(0, plc_config["equipment_name"])
        self.ip_entry.delete(0, tk.END)
        self.ip_entry.insert(0, plc_config["ip_address"])
        self.port_entry.delete(0, tk.END)
        self.port_entry.insert(0, plc_config["port"])

    def edit_traceability_config(self):
        """Edit an existing TRACEABILITY configuration."""
        selected_index = self.traceability_list.curselection()
        if not selected_index:
            messagebox.showwarning("Warning", "Please select TRACEABILITY configuration to edit.")
            return

        self.selected_traceability_index = selected_index[0]
        traceability_config = self.traceability_configs[self.selected_traceability_index]

        self.file_name_entry1.delete(0, tk.END)
        self.file_name_entry1.insert(0, traceability_config["file_name"])
        self.reg_type_combobox1.set(traceability_config["register_type"])
        self.start_reg_entry1.delete(0, tk.END)
        self.start_reg_entry1.insert(0, traceability_config["start_register"])
        self.range_entry1.delete(0, tk.END)
        self.range_entry1.insert(0, traceability_config["range"])
        self.trigger_type_combobox1.set(traceability_config["trigger_register_type"])
        self.trigger_entry1.delete(0, tk.END)
        self.trigger_entry1.insert(0, traceability_config["trigger_register"])
        self.traceability_folder_path_entry.delete(0, tk.END)
        self.traceability_folder_path_entry.insert(0, traceability_config["folder_path"])

    def edit_error_code_config(self):
        """Edit an existing ERROR_CODE configuration."""
        selected_index = self.error_code_list.curselection()
        if not selected_index:
            messagebox.showwarning("Warning", "Please select ERROR_CODE configuration to edit.")
            return

        self.selected_error_code_index = selected_index[0]
        error_code_config = self.error_code_configs[self.selected_error_code_index]

        self.file_name_entry2.delete(0, tk.END)
        self.file_name_entry2.insert(0, error_code_config["file_name"])
        self.reg_type_combobox2.set(error_code_config["register_type"])
        self.start_reg_entry2.delete(0, tk.END)
        self.start_reg_entry2.insert(0, error_code_config["start_register"])
        self.range_entry2.delete(0, tk.END)
        self.range_entry2.insert(0, error_code_config["range"])
        self.trigger_type_combobox2.set(error_code_config["trigger_register_type"])
        self.trigger_entry2.delete(0, tk.END)
        self.trigger_entry2.insert(0, error_code_config["trigger_register"])
        self.error_code_folder_path_entry.delete(0, tk.END)
        self.error_code_folder_path_entry.insert(0, error_code_config["folder_path"])

    def edit_down_time_config(self):
        """Edit an existing DOWN_TIME configuration."""
        selected_index = self.down_time_list.curselection()
        if not selected_index:
            messagebox.showwarning("Warning", "Please select DOWN_TIME configuration to edit.")
            return

        self.selected_down_time_index = selected_index[0]
        down_time_config = self.down_time_configs[self.selected_down_time_index]

        self.file_name_entry3.delete(0, tk.END)
        self.file_name_entry3.insert(0, down_time_config["file_name"])
        self.reg_type_combobox3.set(down_time_config["register_type"])
        self.start_reg_entry3.delete(0, tk.END)
        self.start_reg_entry3.insert(0, down_time_config["start_register"])
        self.range_entry3.delete(0, tk.END)
        self.range_entry3.insert(0, down_time_config["range"])
        self.trigger_type_combobox3.set(down_time_config["trigger_register_type"])
        self.trigger_entry3.delete(0, tk.END)
        self.trigger_entry3.insert(0, down_time_config["trigger_register"])
        self.down_time_folder_path_entry.delete(0, tk.END)
        self.down_time_folder_path_entry.insert(0, down_time_config["folder_path"])

    def save_plc_config(self):
        """Save changes made to an existing PLC configuration."""
        if self.selected_plc_index is None:
            messagebox.showwarning("Warning", "No configuration selected for saving.")
            return

        self.plc_configs[self.selected_plc_index] = {
            "line_name": self.line_entry.get(),
            "equipment_name": self.equipment_entry.get(),
            "ip_address": self.ip_entry.get(),
            "port": self.port_entry.get(),
            "status": self.plc_configs[self.selected_plc_index].get("status", "Not Connected")
        }

        self.selected_plc_index = None
        self.refresh_plc_list()
        self.save_all_configs()
        messagebox.showinfo("Info", "PLC Configuration saved successfully!")

    def save_traceability_config(self):
        """Save changes made to an existing TRACEABILITY configuration."""
        if self.selected_traceability_index is None:
            messagebox.showwarning("Warning", "No configuration selected for saving.")
            return

        self.traceability_configs[self.selected_traceability_index] = {
            "file_name": self.file_name_entry1.get(),
            "register_type": self.reg_type_combobox1.get(),
            "start_register": int(self.start_reg_entry1.get()),
            "range": int(self.range_entry1.get()),
            "trigger_register_type": self.trigger_type_combobox1.get(),
            "trigger_register": int(self.trigger_entry1.get()),
            "folder_path": self.traceability_folder_path_entry.get()
        }

        self.selected_traceability_index = None
        self.refresh_traceability_list()
        self.save_all_configs()
        messagebox.showinfo("Info", "TRACEABILITY Configuration saved successfully!")

    def save_error_code_config(self):
        """Save changes made to an existing ERROR_CODE configuration."""
        if self.selected_error_code_index is None:
            messagebox.showwarning("Warning", "No configuration selected for saving.")
            return

        self.error_code_configs[self.selected_error_code_index] = {
            "file_name": self.file_name_entry2.get(),
            "register_type": self.reg_type_combobox2.get(),
            "start_register": int(self.start_reg_entry2.get()),
            "range": int(self.range_entry2.get()),
            "trigger_register_type": self.trigger_type_combobox2.get(),
            "trigger_register": int(self.trigger_entry2.get()),
            "folder_path": self.error_code_folder_path_entry.get()
        }

        self.selected_error_code_index = None
        self.refresh_error_code_list()
        self.save_all_configs()
        messagebox.showinfo("Info", "ERROR_CODE Configuration saved successfully!")

    def save_down_time_config(self):
        """Save changes made to an existing DOWN_TIME configuration."""
        if self.selected_down_time_index is None:
            messagebox.showwarning("Warning", "No configuration selected for saving.")
            return

        self.down_time_configs[self.selected_down_time_index] = {
            "file_name": self.file_name_entry3.get(),
            "register_type": self.reg_type_combobox3.get(),
            "start_register": int(self.start_reg_entry3.get()),
            "range": int(self.range_entry3.get()),
            "trigger_register_type": self.trigger_type_combobox3.get(),
            "trigger_register": int(self.trigger_entry3.get()),
            "folder_path": self.down_time_folder_path_entry.get()
        }

        self.selected_down_time_index = None
        self.refresh_down_time_list()
        self.save_all_configs()
        messagebox.showinfo("Info", "DOWN_TIME Configuration saved successfully!")

    def delete_plc_config(self):
        """Delete an existing PLC configuration."""
        selected_index = self.plc_list.curselection()
        if not selected_index:
            messagebox.showwarning("Warning", "Please select a PLC configuration to delete.")
            return

        confirmed = messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this PLC configuration?")
        if confirmed:
            del self.plc_configs[selected_index[0]]
            self.refresh_plc_list()
            self.save_all_configs()
            messagebox.showinfo("Info", "PLC Configuration deleted successfully!")

    def delete_traceability_config(self):
        """Delete an existing TRACEABILITY configuration."""
        selected_index = self.traceability_list.curselection()
        if not selected_index:
            messagebox.showwarning("Warning", "Please select a TRACEABILITY configuration to delete.")
            return

        confirmed = messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this TRACEABILITY configuration?")
        if confirmed:
            del self.traceability_configs[selected_index[0]]
            self.refresh_traceability_list()
            self.save_all_configs()
            messagebox.showinfo("Info", "TRACEABILITY Configuration deleted successfully!")

    def delete_error_code_config(self):
        """Delete an existing ERROR_CODE configuration."""
        selected_index = self.error_code_list.curselection()
        if not selected_index:
            messagebox.showwarning("Warning", "Please select a ERROR_CODE configuration to delete.")
            return

        confirmed = messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this ERROR_CODE configuration?")
        if confirmed:
            del self.error_code_configs[selected_index[0]]
            self.refresh_error_code_list()
            self.save_all_configs()
            messagebox.showinfo("Info", "ERROR_CODE Configuration deleted successfully!")

    def delete_down_time_config(self):
        """Delete an existing DOWN_TIME configuration."""
        selected_index = self.down_time_list.curselection()
        if not selected_index:
            messagebox.showwarning("Warning", "Please select a DOWN_TIME configuration to delete.")
            return

        confirmed = messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this DOWN_TIME configuration?")
        if confirmed:
            del self.down_time_configs[selected_index[0]]
            self.refresh_down_time_list()
            self.save_all_configs()
            messagebox.showinfo("Info", "DOWN_TIME Configuration deleted successfully!")

    def start_monitoring(self):
        """Start monitoring with clear log updates."""
        self.monitoring = True
        self.run_button.config(bg="green", fg="white")  # Change the background of Run button to green
        self.stop_button.config(bg="white", fg="black")  # Reset the Stop button appearance
        self.update_output_text("Monitoring started...")
        threading.Thread(target=self.real_time_read_registers, daemon=True).start()

    def stop_monitoring(self):
        """Stop monitoring with log update."""
        self.monitoring = False
        self.stop_button.config(bg="red", fg="white")  # Change the background of Stop button to red
        self.run_button.config(bg="white", fg="black")  # Reset the Run button appearance
        self.update_output_text("Monitoring stopped.")

    def update_output_text(self, message):
        """
        Displays only the latest log message in the text widget.
        Clears the widget before displaying the new message.
        """
        # Clear any previous content
        self.output_text.delete(1.0, tk.END)

        # Insert the new log message
        self.output_text.insert(tk.END, message)

        # Optional: Force the UI to refresh immediately (helps if multiple updates occur quickly)
        self.output_text.update_idletasks()

    def real_time_read_registers(self):
        """
        Continuously monitor all PLC registers in real-time using threading.
        Write each PLC's register data to a unified CSV file only when:
            - The trigger register transitions ON (value = 1).
            - Register data has changed since the last recorded state.
            - Data is written only once per trigger event.
        """
        # Dictionary to track previous states for each PLC
        previous_register_data = {plc["line_name"]: None for plc in
                                  self.plc_configs}  # Tracks previous register data per PLC
        previous_trigger_values = {plc["line_name"]: False for plc in self.plc_configs}  # Tracks trigger status per PLC
        lock = threading.Lock()  # Lock for thread safety when writing to the CSV file

        def process_registers(plc, output, category):
            """
            Continuously monitor registers for a single PLC and write new data to the CSV file only when:
            - The trigger transitions from OFF to ON.
            - The register data has changed since the last logged state.
            """
            try:
                # Validate required keys in Output configuration
                required_keys = ["start_register", "range", "trigger_register", "trigger_register_type"]
                if not all(key in output for key in required_keys):
                    self.output_text.insert(
                        tk.END, f"Invalid Output Configuration for PLC '{plc['line_name']}' in category '{category}'.\n"
                    )
                    return

                # Set up Modbus client connection
                client = ModbusTcpClient(plc["ip_address"], port=plc["port"])
                if not client.connect():
                    self.output_text.insert(
                        tk.END, f"Failed to connect to PLC '{plc['line_name']}' for category '{category}'.\n"
                    )
                    return

                # Extract values from Output configuration
                trigger_register = int(output["trigger_register"])
                trigger_type = output["trigger_register_type"]
                start_register = int(output["start_register"])
                register_range = int(output["range"])

                # Initialize tracking variables specific to this PLC
                logged_register_data = None  # Stores the last written register data
                previous_trigger_status = False  # Tracks the trigger's ON/OFF state

                while True:  # Continuous monitoring loop
                    # Read the trigger register
                    if trigger_type == "Coil":
                        trigger_response = client.read_coils(address=trigger_register, count=1)
                    elif trigger_type == "Discrete":
                        trigger_response = client.read_discrete_inputs(address=trigger_register, count=1)
                    elif trigger_type == "Holding":
                        trigger_response = client.read_holding_registers(address=trigger_register, count=1)
                    elif trigger_type == "Input":
                        trigger_response = client.read_input_registers(address=trigger_register, count=1)
                    else:
                        self.output_text.insert(
                            tk.END,
                            f"Unsupported trigger register type: '{trigger_type}' for PLC '{plc['line_name']}' in category '{category}'.\n"
                        )
                        break

                    if trigger_response.isError():
                        self.output_text.insert(
                            tk.END,
                            f"Error reading trigger register for PLC '{plc['line_name']}' in category '{category}'.\n"
                        )
                        break

                    # Get the current trigger value
                    current_trigger_value = trigger_response.bits[0] if trigger_type in ["Coil", "Discrete"] else \
                        trigger_response.registers[0]

                    # Check if the trigger transitions to ON
                    if current_trigger_value == 1:
                        if not previous_trigger_status:  # Transition from OFF to ON detected
                            # Read register data only if the trigger register is ON
                            register_response = client.read_holding_registers(address=start_register,
                                                                              count=register_range)
                            if register_response.isError() or not register_response.registers:
                                self.output_text.insert(
                                    tk.END,
                                    f"Error reading registers for PLC '{plc['line_name']}' in category '{category}'.\n"
                                )
                                break

                            current_registers = register_response.registers

                            # Write to CSV **only if data has changed** since the last logged state
                            with lock:  # Ensure thread-safe access to shared resources
                                if current_registers != logged_register_data:
                                    self.log_to_csv(output, plc, current_registers, category)  # Write data to CSV
                                    logged_register_data = current_registers  # Update the last logged data

                            previous_trigger_status = True  # Update trigger status to ON

                    elif previous_trigger_status:  # Transition back to OFF
                        previous_trigger_status = False  # Reset trigger status

                    time.sleep(0.1)  # Adjust polling interval

            except Exception as e:
                self.output_text.insert(
                    tk.END, f"Error in monitoring for PLC '{plc['line_name']}' in category '{category}': {e}\n"
                )
            finally:
                if client.is_socket_open():
                    client.close()

        # Spawn threads for each PLC and Configuration
        threads = []
        for plc in self.plc_configs:
            for config_type, configs in [("Traceability", self.traceability_configs),
                                         ("ErrorCodes", self.error_code_configs),
                                         ("DownTime", self.down_time_configs)]:
                if configs:
                    for config in configs:
                        thread = threading.Thread(target=process_registers, args=(plc, config, config_type))
                        thread.daemon = True
                        thread.start()
                        threads.append(thread)

    def log_to_csv(self, output, plc, register_data, category):
        """
        Write register data row by row into a unified CSV file for all PLCs.
        """
        output_folder = output.get("folder_path", "")
        file_name = output.get("file_name", "")
        if not output_folder or not file_name:
            self.output_text.insert(
                tk.END,
                f"No valid folder path or file name configured for PLC '{plc['line_name']}' in category '{category}'.\n"
            )
            return

        # Unified CSV file name for all PLCs
        file_date = datetime.now().strftime("%Y-%m-%d")
        csv_file_name = os.path.join(output_folder, f"{file_name}_{file_date}.csv")

        # Ensure the output folder exists
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        # Prepare data row for this PLC
        timestamp_data = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        row = {
            "Line Name": plc["line_name"],
            "Equipment Name": plc["equipment_name"],
            "IP Address": plc["ip_address"],
            "Timestamp": timestamp_data,
        }

        # Append register data as additional columns
        row.update({
            f"Register_{i + 1}": value for i, value in enumerate(register_data)
        })

        # Prepare headers dynamically for registers
        headers = ["Line Name", "Equipment Name", "IP Address", "Timestamp"] + [f"Register_{i + 1}" for i in range(len(register_data))]

        # Write the data to the CSV file
        try:
            with open(csv_file_name, "a", newline="") as csv_file:
                csv_writer = csv.DictWriter(csv_file, fieldnames=headers)
                if csv_file.tell() == 0:  # Write headers only if file is empty/new
                    csv_writer.writeheader()
                csv_writer.writerow(row)
                self.output_text.insert(
                    tk.END, f"Written data for PLC '{plc['line_name']}' to {csv_file_name}.\n"
                )
        except Exception as e:
            self.output_text.insert(
                tk.END, f"Error writing to {csv_file_name} for PLC '{plc['line_name']}' in category '{category}': {e}\n"
            )
    def on_close(self):
        """
        Handle the app window close event.
        Stop monitoring and save all configurations.
        """
        self.monitoring = False  # Ensure monitoring stops
        self.save_all_configs()  # Save configurations to JSON files
        self.root.destroy()  # Close the application window

if __name__ == "__main__":
    root = tk.Tk()
    app = PLCManagerApp(root)
    root.mainloop()








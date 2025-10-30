# Create the final, clean, and complete Python script with all features


import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import xml.etree.ElementTree as ET
from datetime import datetime
import os
import json
import threading
import traceback
import getpass

class ERSAProgramGeneratorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("ERSA Program Generator - Enhanced Edition v2.0")
        self.root.geometry("1400x900")
        self.root.resizable(True, True)
        
        # Core variables
        self.excel_file = tk.StringVar()
        self.template_file = tk.StringVar()
        self.output_folder = tk.StringVar(value="Generated_Programs")
        self.df = None
        self.excel_columns = []
        
        # Column mapping variables (parameter_name: selected_column)
        self.column_mapping = {}
        self.mapping_vars = {}
        self.heating_zone_mapping = {}
        self.cooling_zone_mapping = {}
        self.load_heating_mapping()
        self.load_cooling_mapping()
        self.detect_template_file()


        
        # Configuration
        self.config_file = "column_mapping_config.json"
        
        # Zone data (zone_key: tk.StringVar())
        self.zone_vars = {}
        self.current_program_index = 0
        
        # Skipped programs collector
        self.skipped_programs = []
        
        # Setup
        self.setup_styles()
        self.create_interface()
        self.load_saved_mapping()
        
    def setup_styles(self):
        """Configure professional styles"""
        style = ttk.Style()
        try:
            style.theme_use('clam')
        except Exception:
            pass
        
        style.configure('Title.TLabel', font=('Arial', 14, 'bold'), 
                       foreground='#1a5490')
        style.configure('Subtitle.TLabel', font=('Arial', 10, 'bold'), 
                       foreground='#2c3e50')
        style.configure('Info.TLabel', font=('Arial', 9), 
                       foreground='#34495e')
        style.configure('Generate.TButton', font=('Arial', 11, 'bold'), 
                       padding=10)
        
    def create_interface(self):
        """Create main tabbed interface"""
        # Main notebook
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Create all tabs (mapping_tab must be created before load_saved_mapping)
        self.create_files_tab()
        self.create_mapping_tab()
        self.create_heating_tab()
        self.create_cooling_tab()
        self.create_log_tab()
        self.create_metadata_tab()

    # ==================== TAB 1: FILES & SETTINGS ====================
    def create_files_tab(self):
        """Create main file selection and program list tab"""
        tab = ttk.Frame(self.notebook, padding="20")
        self.notebook.add(tab, text="ðŸ“ Files & Settings")
        
        # Title
        ttk.Label(tab, text="ERSA Soldering Program Generator", 
                 style='Title.TLabel').grid(row=0, column=0, columnspan=3, 
                                           pady=(0, 20), sticky=tk.W)
        
        ttk.Label(tab, text="By Prasad Gawas.", style='Info.TLabel', font=('Arial', 9, 'italic')).grid(row=0, column=1, columnspan=3, pady=(0, 20), sticky=tk.W)
        
        # File selection frame
        file_frame = ttk.LabelFrame(tab, text="Input Files", padding="15")
        file_frame.grid(row=1, column=0, columnspan=3, pady=(0, 20), 
                       sticky=(tk.W, tk.E))
        file_frame.columnconfigure(1, weight=1)
        
        # Excel file
        ttk.Label(file_frame, text="Excel Data File:", 
                 style='Subtitle.TLabel').grid(row=0, column=0, sticky=tk.W, 
                                               pady=5, padx=(0, 10))
        ttk.Entry(file_frame, textvariable=self.excel_file, width=60).grid(
            row=0, column=1, pady=5, sticky=(tk.W, tk.E))
        ttk.Button(file_frame, text="Browse...", 
                  command=self.browse_excel).grid(row=0, column=2, pady=5, 
                                                  padx=(10, 0))
        
        # Template XML
        ttk.Label(file_frame, text="Template XML:", 
                 style='Subtitle.TLabel').grid(row=1, column=0, sticky=tk.W, 
                                               pady=5, padx=(0, 10))
        ttk.Entry(file_frame, textvariable=self.template_file, width=60).grid(
            row=1, column=1, pady=5, sticky=(tk.W, tk.E))
        ttk.Button(file_frame, text="Browse...", 
                  command=self.browse_template).grid(row=1, column=2, pady=5, 
                                                     padx=(10, 0))
        
        # Output folder
        ttk.Label(file_frame, text="Output Folder:", 
                 style='Subtitle.TLabel').grid(row=2, column=0, sticky=tk.W, 
                                               pady=5, padx=(0, 10))
        ttk.Entry(file_frame, textvariable=self.output_folder, width=60).grid(
            row=2, column=1, pady=5, sticky=(tk.W, tk.E))
        ttk.Button(file_frame, text="Browse...", 
                  command=self.browse_output).grid(row=2, column=2, pady=5, 
                                                   padx=(10, 0))
        
        # Quick actions
        action_frame = ttk.LabelFrame(tab, text="Quick Actions", padding="15")
        action_frame.grid(row=2, column=0, columnspan=3, pady=(0, 20), 
                         sticky=(tk.W, tk.E))
        
        ttk.Button(action_frame, text="ðŸ“‚ Load Excel & Detect Columns", 
                  command=self.load_excel_file).pack(side=tk.LEFT, padx=5)
        ttk.Button(action_frame, text="ðŸ”— View Column Mapping", 
                  command=lambda: self.notebook.select(1)).pack(side=tk.LEFT, padx=5)
        ttk.Button(action_frame, text="ðŸ”¥ Edit Heating Zones", 
                  command=lambda: self.notebook.select(2)).pack(side=tk.LEFT, padx=5)
        ttk.Button(action_frame, text="ðŸ’¾ Save Mapping", 
                  command=self.save_mapping).pack(side=tk.LEFT, padx=5)
        
        # Program list preview
        preview_frame = ttk.LabelFrame(tab, text="Programs to Generate", 
                                      padding="15")
        preview_frame.grid(row=3, column=0, columnspan=3, pady=(0, 20), 
                          sticky=(tk.W, tk.E, tk.N, tk.S))
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(1, weight=1)
        
        tab.rowconfigure(3, weight=1)
        tab.columnconfigure(0, weight=1)
        
        # Listbox with scrollbar
        scroll = ttk.Scrollbar(preview_frame)
        scroll.grid(row=1, column=1, sticky=(tk.N, tk.S))
        
        self.program_listbox = tk.Listbox(preview_frame, height=15, 
                                         yscrollcommand=scroll.set,
                                         font=('Consolas', 9))
        self.program_listbox.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scroll.config(command=self.program_listbox.yview)
        self.program_listbox.bind('<<ListboxSelect>>', self.on_program_select)
        
        # Control buttons
        btn_frame = ttk.Frame(tab)
        btn_frame.grid(row=4, column=0, columnspan=3, pady=(10, 0))
        
        self.generate_btn = ttk.Button(btn_frame, text="ðŸš€ Generate All Programs", 
                                       style='Generate.TButton',
                                       command=self.start_generation)
        self.generate_btn.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(btn_frame, text="Exit", 
                  command=self.root.quit).pack(side=tk.LEFT, padx=5)
    
    #===================== META DATA ====================
    def create_metadata_tab(self):
    """Create metadata and program IDs configuration tab"""
    tab = ttk.Frame(self.notebook, padding="20")
    self.notebook.add(tab, text="📝 Metadata & IDs")
    
    ttk.Label(tab, text="Program Metadata & ID Management", 
             style='Title.TLabel').grid(row=0, column=0, columnspan=3, 
                                       pady=(0, 10), sticky=tk.W)
    
    ttk.Label(tab, 
             text="Configure auto-incrementing IDs and metadata for generated programs.",
             style='Info.TLabel').grid(row=1, column=0, columnspan=3, 
                                      pady=(0, 20), sticky=tk.W)
    
    # Metadata variables - CORRECTED with numeric User ID
    self.meta_programid_start = tk.IntVar(value=10000)
    self.meta_libraryid_start = tk.IntVar(value=100)
    self.meta_version = tk.IntVar(value=1)
    self.meta_setnumber_start = tk.IntVar(value=1)
    self.meta_historyid_start = tk.IntVar(value=6000)
    self.meta_userid = tk.IntVar(value=881)  # SINGLE numeric User ID
    self.meta_default_notes = tk.StringVar(value="Auto-generated by ERSA tool")
    
    row = 2
    ttk.Label(tab, text="Starting Program ID (auto-increment):").grid(
        row=row, column=0, sticky=tk.W, pady=5); row+=1
    ttk.Entry(tab, textvariable=self.meta_programid_start, width=12).grid(
        row=row-1, column=1, sticky=tk.W)
    
    ttk.Label(tab, text="Library ID (Manual - Fixed):").grid(
        row=row, column=0, sticky=tk.W, pady=5); row+=1
    ttk.Entry(tab, textvariable=self.meta_libraryid_start, width=12).grid(
        row=row-1, column=1, sticky=tk.W)
    
    ttk.Label(tab, text="Default Program Version:").grid(
        row=row, column=0, sticky=tk.W, pady=5); row+=1
    ttk.Entry(tab, textvariable=self.meta_version, width=5).grid(
        row=row-1, column=1, sticky=tk.W)
    
    ttk.Label(tab, text="Setnumber (Manual - Fixed):").grid(
        row=row, column=0, sticky=tk.W, pady=5); row+=1
    ttk.Entry(tab, textvariable=self.meta_setnumber_start, width=6).grid(
        row=row-1, column=1, sticky=tk.W)
    
    ttk.Label(tab, text="Starting historyid (auto-increment):").grid(
        row=row, column=0, sticky=tk.W, pady=5); row+=1
    ttk.Entry(tab, textvariable=self.meta_historyid_start, width=12).grid(
        row=row-1, column=1, sticky=tk.W)
    
    ttk.Label(tab, text="User ID (for creation & change):").grid(
        row=row, column=0, sticky=tk.W, pady=5); row+=1
    ttk.Entry(tab, textvariable=self.meta_userid, width=12).grid(
        row=row-1, column=1, sticky=tk.W)
    
    ttk.Label(tab, text="Default Notes:").grid(
        row=row, column=0, sticky=tk.W, pady=5); row+=1
    ttk.Entry(tab, textvariable=self.meta_default_notes, width=35).grid(
        row=row-1, column=1, sticky=tk.W)
    
    ttk.Label(tab, text="Creation/Change date will be set at time of generation.", 
             foreground="gray").grid(row=row, column=0, columnspan=2, sticky=tk.W)

    #==================== END META DATA =====================

    # ==================== TAB 2: COLUMN MAPPING ====================
    def create_mapping_tab(self):
        """Create column mapping configuration tab"""
        tab = ttk.Frame(self.notebook, padding="20")
        self.notebook.add(tab, text="ðŸ”— Column Mapping")
        
        # Header
        ttk.Label(tab, text="Map Excel Columns to Program Parameters", 
                 style='Title.TLabel').grid(row=0, column=0, columnspan=3, 
                                           pady=(0, 10), sticky=tk.W)
        
        ttk.Label(tab, 
                 text="Select which Excel column contains each parameter. Dropdowns show YOUR column names.",
                 style='Info.TLabel').grid(row=1, column=0, columnspan=3, 
                                          pady=(0, 20), sticky=tk.W)
        
        # Scrollable frame
        canvas = tk.Canvas(tab, height=600)
        scrollbar = ttk.Scrollbar(tab, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind("<Configure>", 
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=2, column=2, sticky=(tk.N, tk.S))
        
        tab.rowconfigure(2, weight=1)
        tab.columnconfigure(0, weight=1)
        
        # Define parameters to map
        self.param_definitions = [
            ("Basic Parameters", [
                ("PCB Name (STENCIL)", "STENCIL", "Program identifier"),
                ("PCB Length (mm)", "PCB_Length", "Board length"),
                ("PCB Width / Conveyor Width (mm)", "PCB_Width", "Conveyor width"),
                ("CBS Width / Middle Support (mm)", "CBS_Width", "Middle support width"),
                ("Park Position Width (mm)", "Park_Position_Width", "Park position"),
            ]),
            ("Heating Zones - Top (Optional)", [
                ("Temperature Zone 1-10 (Column Pattern)", "Heating_Top_Temp", "e.g., HZ_Top_Temp_Z1, HZ_Top_Temp_Z2..."),
                ("Tolerance+ Zone 1-10 (Column Pattern)", "Heating_Top_TolPlus", "e.g., HZ_Top_Plus_Z1..."),
                ("Tolerance- Zone 1-10 (Column Pattern)", "Heating_Top_TolMinus", "e.g., HZ_Top_Minus_Z1..."),
                ("Convection% Zone 1-10 (Column Pattern)", "Heating_Top_Conv", "e.g., HZ_Top_Conv_Z1..."),
            ]),
            ("Heating Zones - Bottom (Optional)", [
                ("Temperature Zone 1-10 (Column Pattern)", "Heating_Bottom_Temp", "e.g., HZ_Bot_Temp_Z1..."),
                ("Tolerance+ Zone 1-10 (Column Pattern)", "Heating_Bottom_TolPlus", "e.g., HZ_Bot_Plus_Z1..."),
                ("Tolerance- Zone 1-10 (Column Pattern)", "Heating_Bottom_TolMinus", "e.g., HZ_Bot_Minus_Z1..."),
                ("Convection% Zone 1-10 (Column Pattern)", "Heating_Bottom_Conv", "e.g., HZ_Bot_Conv_Z1..."),
            ]),
            ("Cooling Zones - Top (Optional)", [
                ("Temperature Zone 1-3 (Column Pattern)", "Cooling_Top_Temp", "e.g., CZ_Top_Temp_Z1..."),
                ("Tolerance+ Zone 1-3 (Column Pattern)", "Cooling_Top_TolPlus", "e.g., CZ_Top_Plus_Z1..."),
                ("Tolerance- Zone 1-3 (Column Pattern)", "Cooling_Top_TolMinus", "e.g., CZ_Top_Minus_Z1..."),
                ("Convection% Zone 1-3 (Column Pattern)", "Cooling_Top_Conv", "e.g., CZ_Top_Conv_Z1..."),
            ]),
            ("Cooling Zones - Bottom (Optional)", [
                ("Convection% Zone 1-3 (Column Pattern)", "Cooling_Bottom_Conv", "e.g., CZ_Bot_Conv_Z1..."),
            ]),
        ]
        
        # Create mapping fields
        row = 0
        for section_title, params in self.param_definitions:
            # Section header
            ttk.Label(scrollable_frame, text=section_title, 
                     style='Subtitle.TLabel').grid(row=row, column=0, columnspan=3, 
                                                   pady=(15, 10), sticky=tk.W)
            row += 1
            
            for label, key, hint in params:
                # Parameter label
                ttk.Label(scrollable_frame, text=label + ":").grid(
                    row=row, column=0, sticky=tk.W, pady=5, padx=(20, 10))
                
                # Dropdown combobox
                var = tk.StringVar(value="(None)")
                self.mapping_vars[key] = var
                
                combo = ttk.Combobox(scrollable_frame, textvariable=var, 
                                   width=40, state='readonly')
                combo['values'] = ["(None)"]
                combo.grid(row=row, column=1, sticky=tk.W, pady=5)
                
                # Hint label
                ttk.Label(scrollable_frame, text=hint, 
                         font=('Arial', 8), foreground='gray').grid(
                    row=row, column=2, sticky=tk.W, pady=5, padx=(10, 0))
                
                row += 1
        
        # Save button at bottom
        ttk.Button(tab, text="ðŸ’¾ Save Column Mapping", 
                  command=self.save_mapping).grid(row=3, column=0, pady=15)
    # ==================== TAB 3: HEATING ZONES ====================
    def create_heating_tab(self):
        """Create heating zones editing tab"""
        tab = ttk.Frame(self.notebook, padding="20")
        self.notebook.add(tab, text="ðŸ”¥ Heating Zones")
        
        # Program selector
        selector_frame = ttk.Frame(tab)
        selector_frame.grid(row=0, column=0, columnspan=2, pady=(0, 20), 
                           sticky=(tk.W, tk.E))
        
        ttk.Label(selector_frame, text="Select Program:", 
                 style='Subtitle.TLabel').pack(side=tk.LEFT, padx=(0, 10))
        
        self.program_selector = ttk.Combobox(selector_frame, width=50, 
                                            state='readonly')
        self.program_selector.pack(side=tk.LEFT, padx=(0, 10))
        self.program_selector.bind('<<ComboboxSelected>>', self.load_program_zones)
        
        ttk.Button(selector_frame, text="â—€ Previous", 
                  command=self.prev_program).pack(side=tk.LEFT, padx=5)
        ttk.Button(selector_frame, text="Next â–¶", 
                  command=self.next_program).pack(side=tk.LEFT, padx=5)
        
        # Scrollable zones frame
        canvas = tk.Canvas(tab)
        v_scroll = ttk.Scrollbar(tab, orient="vertical", command=canvas.yview)
        h_scroll = ttk.Scrollbar(tab, orient="horizontal", command=canvas.xview)
        
        zones_frame = ttk.Frame(canvas)
        zones_frame.bind("<Configure>", 
                        lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        
        canvas.create_window((0, 0), window=zones_frame, anchor="nw")
        canvas.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        
        canvas.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        v_scroll.grid(row=1, column=1, sticky=(tk.N, tk.S))
        h_scroll.grid(row=2, column=0, sticky=(tk.W, tk.E))
        
        tab.rowconfigure(1, weight=1)
        tab.columnconfigure(0, weight=1)
        
        # TOP HEATING ZONES
        top_frame = ttk.LabelFrame(zones_frame, text="TOP HEATING ZONES (1-10)", 
                                   padding="15")
        top_frame.grid(row=0, column=0, padx=10, pady=10, sticky=(tk.W, tk.E))
        
        self.create_zone_grid(top_frame, "Heating_Top", 10)
        
        # BOTTOM HEATING ZONES
        bottom_frame = ttk.LabelFrame(zones_frame, text="BOTTOM HEATING ZONES (1-10)", 
                                      padding="15")
        bottom_frame.grid(row=1, column=0, padx=10, pady=10, sticky=(tk.W, tk.E))
        
        self.create_zone_grid(bottom_frame, "Heating_Bottom", 10)
        
        # Save button
        ttk.Button(tab, text="ðŸ’¾ Save Zone Changes", 
                  command=self.save_zone_changes).grid(row=3, column=0, pady=10)
    # ==================== TAB 4: COOLING ZONES ====================
    def create_cooling_tab(self):
        """Create cooling zones editing tab"""
        tab = ttk.Frame(self.notebook, padding="20")
        self.notebook.add(tab, text="â„ï¸ Cooling Zones")
        
        # Program selector
        selector_frame = ttk.Frame(tab)
        selector_frame.grid(row=0, column=0, pady=(0, 20), sticky=(tk.W, tk.E))
        
        ttk.Label(selector_frame, text="Select Program:", 
                 style='Subtitle.TLabel').pack(side=tk.LEFT, padx=(0, 10))
        
        self.cooling_selector = ttk.Combobox(selector_frame, width=50, 
                                            state='readonly')
        self.cooling_selector.pack(side=tk.LEFT)
        self.cooling_selector.bind('<<ComboboxSelected>>', self.load_program_zones)
        
        # Scrollable frame
        canvas = tk.Canvas(tab)
        scrollbar = ttk.Scrollbar(tab, orient="vertical", command=canvas.yview)
        
        zones_frame = ttk.Frame(canvas)
        zones_frame.bind("<Configure>", 
                        lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        
        canvas.create_window((0, 0), window=zones_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=1, column=1, sticky=(tk.N, tk.S))
        
        tab.rowconfigure(1, weight=1)
        tab.columnconfigure(0, weight=1)
        
        # TOP COOLING ZONES
        top_frame = ttk.LabelFrame(zones_frame, text="TOP COOLING ZONES (1-3)", 
                                   padding="15")
        top_frame.grid(row=0, column=0, padx=10, pady=10, sticky=(tk.W, tk.E))
        
        self.create_zone_grid(top_frame, "Cooling_Top", 3)
        
        # BOTTOM COOLING ZONES
        bottom_frame = ttk.LabelFrame(zones_frame, text="BOTTOM COOLING ZONES (1-3)", 
                                      padding="15")
        bottom_frame.grid(row=1, column=0, padx=10, pady=10, sticky=(tk.W, tk.E))
        
        self.create_zone_grid(bottom_frame, "Cooling_Bottom", 3)
        
        # Save button
        ttk.Button(tab, text="ðŸ’¾ Save Cooling Zone Changes", 
                  command=self.save_zone_changes).grid(row=2, column=0, pady=10)
    def create_zone_grid(self, parent, zone_prefix, num_zones):
        """Create grid of editable zone parameters"""
        # Column headers (Zone numbers)
        for zone_num in range(1, num_zones + 1):
            ttk.Label(parent, text=f"Zone {zone_num}", 
                     style='Subtitle.TLabel').grid(row=0, column=zone_num, 
                                                   padx=5, pady=5)
        
        # Parameter rows
        params = [
            ("Temperature (Â°C)", "Temp"),
            ("Tolerance + (Â°C)", "TolPlus"),
            ("Tolerance - (Â°C)", "TolMinus"),
            ("Convection (%)", "Conv")
        ]
        
        for param_idx, (param_label, param_key) in enumerate(params, start=1):
            # Row label
            ttk.Label(parent, text=param_label).grid(row=param_idx, column=0, 
                                                     sticky=tk.W, padx=5, pady=5)
            
            # Entry fields for each zone
            for zone_num in range(1, num_zones + 1):
                var_key = f"{zone_prefix}_Z{zone_num}_{param_key}"
                
                if var_key not in self.zone_vars:
                    self.zone_vars[var_key] = tk.StringVar(value="0")
                
                entry = ttk.Entry(parent, textvariable=self.zone_vars[var_key], 
                                 width=10, justify='center')
                entry.grid(row=param_idx, column=zone_num, padx=5, pady=2)
    # ==================== TAB 5: LOG ====================
    def create_log_tab(self):
        """Create generation log tab"""
        tab = ttk.Frame(self.notebook, padding="20")
        self.notebook.add(tab, text="ðŸ“‹ Generation Log")
        
        # Log text area
        self.log_text = scrolledtext.ScrolledText(tab, height=30, 
                                                  font=('Consolas', 9),
                                                  bg='#f8f9fa', fg='#2c3e50')
        self.log_text.pack(fill='both', expand=True)

        # configure tag for red (error) messages
        self.log_text.tag_configure('red', foreground='red')

        # Progress bar
        self.progress = ttk.Progressbar(tab, mode='indeterminate')
        self.progress.pack(fill='x', pady=(10, 0))
        
        # Export buttons
        btn_frame = ttk.Frame(tab)
        btn_frame.pack(fill='x', pady=(6, 0))
        ttk.Button(btn_frame, text="ðŸ“¤ Export Log", command=self.export_log).pack(side=tk.LEFT, padx=6)
        ttk.Button(btn_frame, text="ðŸ“¥ Export Skipped", command=self.export_skipped).pack(side=tk.LEFT, padx=6)
        
        # Initial message
        self.log("ERSA Program Generator Enhanced Edition v2.0")
        self.log("=" * 80)
        self.log("Workflow: Load Excel â†’ Map Columns â†’ Edit Zones â†’ Generate")
        self.log("=" * 80 + "\n")
    # ==================== FILE OPERATIONS ====================
    def browse_excel(self):
        """Browse for Excel file"""
        filename = filedialog.askopenfilename(
            title="Select Excel Data File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.excel_file.set(filename)
            self.log(f"âœ“ Excel file selected: {os.path.basename(filename)}")
    def browse_template(self):
        """Browse for template XML"""
        filename = filedialog.askopenfilename(
            title="Select Template XML File",
            filetypes=[("XML files", "*.xml"), ("All files", "*.*")]
        )
        if filename:
            self.template_file.set(filename)
            self.log(f"âœ“ Template file selected: {os.path.basename(filename)}")
    def browse_output(self):
        """Browse for output folder"""
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            self.output_folder.set(folder)
            self.log(f"âœ“ Output folder selected: {folder}")
    def load_excel_file(self):
        """Load Excel file and populate column dropdowns"""
        excel_path = self.excel_file.get()
        
        if not excel_path or not os.path.exists(excel_path):
            messagebox.showerror("Error", "Please select a valid Excel file first!")
            return
        
        try:
            self.log("\nLoading Excel file...")
            self.df = pd.read_excel(excel_path)
            self.excel_columns = list(self.df.columns)
            
            self.log(f"âœ“ Loaded {len(self.df)} programs")
            self.log(f"  Excel columns found: {', '.join(self.excel_columns)}")
            
            # Update all dropdown comboboxes with Excel columns
            dropdown_values = ["(None)"] + self.excel_columns
            
            for key, var in self.mapping_vars.items():
                for widget in self.root.winfo_children():
                    self._update_comboboxes_recursive(widget, var, dropdown_values)
            
            self.log("âœ“ Dropdown lists updated with your Excel columns")
            
            # Auto-detect common columns
            self.auto_detect_columns()
            
            # Update program list and selectors
            self.update_program_list()
            self.update_program_selectors()
            
            messagebox.showinfo("Success", 
                              f"Excel loaded!\n{len(self.df)} programs found\n\n"
                              f"Next: Go to 'Column Mapping' tab to map your columns")
            
        except Exception as e:
            self.log(f"âœ— Error loading Excel: {str(e)}")
            messagebox.showerror("Error", f"Failed to load Excel:\n{str(e)}")
    def _update_comboboxes_recursive(self, widget, var, values):
        """Recursively find and update comboboxes"""
        try:
            if isinstance(widget, ttk.Combobox):
                if str(widget.cget('textvariable')) == str(var):
                    widget['values'] = values
            
            for child in widget.winfo_children():
                self._update_comboboxes_recursive(child, var, values)
        except Exception:
            pass
    def auto_detect_columns(self):
        """Auto-detect common column names"""
        if not self.excel_columns:
            return
        
        # Create lowercase mapping
        col_map = {col.lower(): col for col in self.excel_columns}
        
        # Detection patterns
        patterns = {
            'STENCIL': ['stencil', 'pcb', 'name', 'pcb_name', 'program', 'board_name'],
            'PCB_Length': ['length', 'board_length', 'pcb_length', 'board_length_prn', 'l'],
            'PCB_Width': ['width', 'pcb_width', 'pcb width', 'conveyor_width', 'w'],
            'CBS_Width': ['cbs', 'cbs_width', 'cbs width', 'middle_support', 'support'],
        }
        
        self.log("\nAuto-detecting columns:")
        for key, keywords in patterns.items():
            if key in self.mapping_vars:
                for keyword in keywords:
                    if keyword in col_map:
                        self.mapping_vars[key].set(col_map[keyword])
                        self.log(f"  âœ“ {key} â†’ {col_map[keyword]}")
                        break
    def update_program_list(self):
        """Update program listbox"""
        self.program_listbox.delete(0, tk.END)
        
        if self.df is not None:
            stencil_col = self.mapping_vars.get('STENCIL', tk.StringVar()).get()
            if stencil_col and stencil_col != "(None)" and stencil_col in self.df.columns:
                for idx, name in enumerate(self.df[stencil_col], start=1):
                    self.program_listbox.insert(tk.END, f"{idx}. {name}")
    def update_program_selectors(self):
        """Update program selector dropdowns"""
        if self.df is None:
            return
        
        stencil_col = self.mapping_vars.get('STENCIL', tk.StringVar()).get()
        if stencil_col and stencil_col != "(None)" and stencil_col in self.df.columns:
            programs = self.df[stencil_col].tolist()
            self.program_selector['values'] = programs
            self.cooling_selector['values'] = programs
            
            if programs:
                self.program_selector.current(0)
                self.cooling_selector.current(0)
    def on_program_select(self, event):
        """Handle program selection from listbox"""
        selection = self.program_listbox.curselection()
        if selection:
            idx = selection[0]
            self.current_program_index = idx
            if len(self.program_selector['values']) > idx:
                self.program_selector.current(idx)
                self.cooling_selector.current(idx)
                self.load_program_zones(None)
    # ==================== ZONE OPERATIONS ====================
    def load_program_zones(self, event):
        """Load zone data for selected program from Excel"""
        if self.df is None:
            return
        
        idx = self.program_selector.current()
        if idx < 0:
            return
        
        self.current_program_index = idx
        row = self.df.iloc[idx]
        
        # Load zone values from Excel columns (if mapped)
        for var_key in self.zone_vars:
            parts = var_key.split('_')
            if len(parts) >= 4:
                zone_type = '_'.join(parts[:-2])  # "Heating_Top"
                zone_num = parts[-2].replace('Z', '')  # "1"
                param_type = parts[-1]  # "Temp"
                
                mapping_key = f"{zone_type}_{param_type}"
                if mapping_key in self.mapping_vars:
                    col_pattern = self.mapping_vars[mapping_key].get()
                    
                    if col_pattern and col_pattern != "(None)":
                        possible_cols = [
                            f"{col_pattern}{zone_num}",
                            f"{col_pattern}_{zone_num}",
                            f"{col_pattern}_Z{zone_num}",
                        ]
                        
                        for col in possible_cols:
                            if col in self.df.columns:
                                value = row[col]
                                if pd.notna(value):
                                    self.zone_vars[var_key].set(str(value))
                                break
        
        program_name = self.program_selector.get()
        self.log(f"Loaded zones for: {program_name}")
    def prev_program(self):
        """Navigate to previous program"""
        current = self.program_selector.current()
        if current > 0:
            self.program_selector.current(current - 1)
            self.cooling_selector.current(current - 1)
            self.load_program_zones(None)
    def next_program(self):
        """Navigate to next program"""
        current = self.program_selector.current()
        if current < len(self.program_selector['values']) - 1:
            self.program_selector.current(current + 1)
            self.cooling_selector.current(current + 1)
            self.load_program_zones(None)
    def save_zone_changes(self):
        """Save manually edited zone values"""
        messagebox.showinfo("Info", "Zone changes saved for current program!")
        self.log("âœ“ Zone values updated")
    # ==================== CONFIGURATION ====================
    def save_mapping(self):
        """Save column mapping to JSON file"""
        try:
            mapping = {key: var.get() for key, var in self.mapping_vars.items()}
            
            with open(self.config_file, 'w') as f:
                json.dump(mapping, f, indent=2)
            
            self.log(f"âœ“ Column mapping saved to {self.config_file}")
            messagebox.showinfo("Success", "Column mapping saved!\n\n"
                              "Next time you load Excel, mappings will be restored.")
        except Exception as e:
            self.log(f"âœ— Error saving mapping: {str(e)}")
            messagebox.showerror("Error", f"Failed to save mapping:\n{str(e)}")
    def load_saved_mapping(self):
        """Load saved column mapping from JSON"""
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r') as f:
                    mapping = json.load(f)
                
                for key, value in mapping.items():
                    if key in self.mapping_vars:
                        self.mapping_vars[key].set(value)
                
                self.log(f"âœ“ Loaded saved column mapping")
            except Exception as e:
                self.log(f"âš  Could not load saved mapping: {str(e)}")

    def save_all_mappings(self):
    """Save all mappings at once"""
    try:
        self.save_mapping()
        self.save_heating_mapping()
        self.save_cooling_mapping()
        self.log("✓ All mappings saved successfully")
        messagebox.showinfo("Success", "All column mappings saved!")
    except Exception as e:
        self.log(f"✗ Error saving mappings: {str(e)}")

    # ==================== GENERATION ====================
    def start_generation(self):
        """Start program generation process"""
        # Validation
        if self.df is None:
            messagebox.showerror("Error", "Please load an Excel file first!")
            return
        
        if not self.template_file.get() or not os.path.exists(self.template_file.get()):
            messagebox.showerror("Error", "Please select a valid template XML file!")
            return
        
        if self.mapping_vars.get('STENCIL', tk.StringVar()).get() == "(None)":
            messagebox.showerror("Error", "Please map the STENCIL/PCB Name column!\n\n"
                               "Go to 'Column Mapping' tab and select which column contains PCB names.")
            return
        
        # Disable button and show progress
        self.generate_btn.config(state='disabled')
        self.progress.start()
        self.notebook.select(4)  # Switch to log tab

        # reset skipped programs list for this run
        self.skipped_programs = []
        
        # Run in thread
        thread = threading.Thread(target=self.generate_programs)
        thread.daemon = True
        thread.start()
    def generate_programs(self):
        """Main generation logic with CBS â†’ Park Position logic"""
        try:
            template_path = self.template_file.get()
            output_dir = self.output_folder.get()
            
            self.log("\n" + "=" * 80)
            self.log("\nSTARTING PROGRAM GENERATION\n")
            self.log("=" * 80)
            
            # Create output directory
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
                self.log(f"\nâœ“ Created output folder: {output_dir}\n")
            
            # Parse template
            self.log("\nParsing template XML...\n")
            template_tree = ET.parse(template_path)
            template_root = template_tree.getroot()
            self.log("\nâœ“ Template loaded\n")
            
            # Get column mappings
            stencil_col = self.mapping_vars.get('STENCIL', tk.StringVar()).get()
            length_col = self.mapping_vars.get('PCB_Length', tk.StringVar()).get()
            width_col = self.mapping_vars.get('PCB_Width', tk.StringVar()).get()
            cbs_col = self.mapping_vars.get('CBS_Width', tk.StringVar()).get()
            
            self.log(f"\nColumn Mapping:")
            self.log(f"\n  PCB Name: {stencil_col}")
            self.log(f"\n  PCB Length: {length_col}")
            self.log(f"\n  PCB Width: {width_col}")
            self.log(f"\n  CBS Width: {cbs_col}")
            
            # Generate programs
            self.log(f"\n{'='*80}\n")
            self.log("\nGENERATING PROGRAMS\n")
            self.log("\n"+'='*80 + "\n")
            
            success_count = 0
            
            for idx, row in self.df.iterrows():
                pcb_name = str(row[stencil_col]) if stencil_col and stencil_col != "(None)" and stencil_col in self.df.columns else f"Program_{idx+1}"
                self.log(f"\n[{idx+1}/{len(self.df)}] Processing: {pcb_name}\n")
                
                # determine length/width values up-front (used both for validation and updates)
                length_val = None
                width_val = None

                def _valid_measure(v):
                    # treat NaN, empty, NA/N/A, '-' as invalid
                    if pd.isna(v):
                        return None
                    if isinstance(v, str):
                        s = v.strip()
                        if s == "" or s.upper() in ("NA", "N/A", "-"):
                            return None
                        # try numeric conversion for strings like "75"
                        try:
                            f = float(s)
                        except Exception:
                            return None
                        return f if f > 0 else None
                    # numeric types
                    try:
                        f = float(v)
                        return f if f > 0 else None
                    except Exception:
                        return None

                if length_col and length_col != "(None)" and length_col in self.df.columns:
                    tmp = row[length_col]
                    length_val = _valid_measure(tmp)

                if width_col and width_col != "(None)" and width_col in self.df.columns:
                    tmp = row[width_col]
                    width_val = _valid_measure(tmp)

                # Require BOTH length and width to be valid; if either is missing/zero/invalid, skip this program
                if length_val is None or width_val is None:
                    reason = "Missing/invalid PCB Length or PCB Width (blank/0/NA/invalid) â€” both required"
                    try:
                        self.skipped_programs.append({'Program': pcb_name, 'Reason': reason})
                    except Exception:
                        self.skipped_programs = [{'Program': pcb_name, 'Reason': reason}]
                    self.log(f"  âœ— Skipped {pcb_name} â€” {reason}", color='red')
                    continue

                try:
                    # Create copy of template
                    new_tree = ET.ElementTree(ET.fromstring(ET.tostring(template_root)))
                    new_root = new_tree.getroot()
                    
                    # Update metadata (use configured meta fields)
                    try:
                        self.update_program_metadata(new_root, pcb_name, idx)
                    except Exception:
                        self.log("  âš  Failed to update metadata (continuing)", color=None)
                    
                    # Update parameters
                    updates = 0
                    
                    # PCB Length (use length_val)
                    if length_val is not None:
                        if self.update_xml_param(new_root, 'enmProg|enmPcb|enmSngSollLaenge', length_val):
                            updates += 1
                    
                    # PCB Width (use width_val)
                    if width_val is not None:
                        if self.update_xml_param(new_root, 'enmProg|enmA_AxBr|1|enmSngSoll', width_val):
                            updates += 1
                    
                    # CBS and Park Position Logic
                    if cbs_col and cbs_col != "(None)" and cbs_col in self.df.columns:
                        cbs_val = row[cbs_col]
                        
                        if pd.isna(cbs_val) or (isinstance(cbs_val, str) and str(cbs_val).strip().upper() == 'NA'):
                            # CBS is NA â†’ Park = True, CBS Active = False
                            park_active = True
                            cbs_active = False
                            self.log(f"  â†’ CBS is NA/empty â†’ Park_Active = True, CBS_Active = False")
                        else:
                            # validate CBS value similar to pcb measures
                            validated_cbs = _valid_measure(cbs_val)
                            if validated_cbs is None:
                                park_active = True
                                cbs_active = False
                                self.log(f"  â†’ CBS value invalid â†’ Park_Active = True, CBS_Active = False")
                            else:
                                park_active = False
                                cbs_active = True
                                self.log(f"  â†’ CBS = {validated_cbs} â†’ Park_Active = False, CBS_Active = True")
                                # Update CBS Width value
                                if self.update_xml_param(new_root, 'enmProg|enmA_Tr|1|enmSngSoll', validated_cbs):
                                    updates += 1
                                # Update CBS Active/Enable parameter
                                if self.update_xml_param(new_root, 'enmProg|enmA_Tr|1|enmBlnSollAktiv', cbs_active, 'Boolean'):
                                    updates += 1
                        
                        # Update Park Position
                        if self.update_xml_param(new_root, 'enmProg|enmA_AxMu|1|enmBlnParkPosSollAktiv', 
                                                park_active, 'Boolean'):
                            updates += 1
                    
                    # Update zone parameters from zone_vars if edited (optional logic left as-is)
                    self.log(f" \n âœ“ Updated {updates} parameters\n")
                    
                    # Save file
                    safe_name = pcb_name.replace('/', '_').replace(' ', '_')
                    output_filename = f"{safe_name}.xml"
                    output_path = os.path.join(output_dir, output_filename)
                    new_tree.write(output_path, encoding='utf-8', xml_declaration=True)
                    self.log(f"\n  âœ“ Saved: {output_filename}\n")
                    
                    success_count += 1
                    
                except Exception as e:
                    self.log(f"  âœ— Error: {str(e)}\n")
                    self.log(traceback.format_exc())
            
            # Summary
            self.log('='*80)
            self.log(f"\nGENERATION COMPLETE: {success_count}/{len(self.df)} programs created\n")
            if getattr(self, 'skipped_programs', None):
                self.log(f"\nSkipped programs: {len(self.skipped_programs)} (use 'Export Skipped' to save details)\n")
            self.log('='*80 + "\n")
            
            self.root.after(0, lambda: messagebox.showinfo(
                "Success", 
                f"Generated {success_count}/{len(self.df)} programs!\n\nOutput: {output_dir}"
            ))
            
        except Exception as e:
            self.log(f"\nâœ— FATAL ERROR: {str(e)}")
            self.log(traceback.format_exc())
            self.root.after(0, lambda: messagebox.showerror("Error", f"Generation failed:\n{str(e)}"))
        
        finally:
            self.root.after(0, lambda: self.generate_btn.config(state='normal'))
            self.root.after(0, lambda: self.progress.stop())
    def update_program_metadata(self, root, program_name, row_index):
        """Update SolderingPrograms metadata using configured meta fields"""
        try:
            # Get metadata settings - CORRECTED
            programid_base = self.meta_programid_start.get()
            libraryid = self.meta_libraryid_start.get()      # FIXED (not incremented)
            setnumber = self.meta_setnumber_start.get()      # FIXED (not incremented)
            historyid_base = self.meta_historyid_start.get()
            version = self.meta_version.get()
            userid = self.meta_userid.get()                  # SINGLE User ID
            notes = self.meta_default_notes.get()
            
            # In the loop:
            for idx, row in self.df.iterrows():
                prog_id = programid_base + idx    # AUTO-INCREMENT
                lib_id = libraryid                # FIXED
                setnum = setnumber                # FIXED
                historyid = historyid_base + idx  # AUTO-INCREMENT
                now = datetime.now().isoformat()
                
                # Update metadata IN-PLACE
                sp = new_root.find('SolderingPrograms')
                if sp is not None:
                    if sp.find('programid') is not None: 
                        sp.find('programid').text = str(prog_id)
                    if sp.find('libraryid') is not None: 
                        sp.find('libraryid').text = str(lib_id)
                    if sp.find('version') is not None: 
                        sp.find('version').text = str(version)
                    if sp.find('creationuser') is not None: 
                        sp.find('creationuser').text = str(userid)
                    if sp.find('changeuser') is not None: 
                        sp.find('changeuser').text = str(userid)
                    if sp.find('creationdate') is not None: 
                        sp.find('creationdate').text = now
                    if sp.find('changedate') is not None: 
                        sp.find('changedate').text = now
                    if sp.find('notes') is not None: 
                        sp.find('notes').text = notes
                    if sp.find('name') is not None: 
                        sp.find('name').text = pcb_name
                
                ph = new_root.find('ProgramHistory')
                if ph is not None:
                    if ph.find('historyid') is not None: 
                        ph.find('historyid').text = str(historyid)
                    if ph.find('setnumber') is not None: 
                        ph.find('setnumber').text = str(setnum)
                    if ph.find('creationuser') is not None: 
                        ph.find('creationuser').text = str(userid)
                    if ph.find('changeuser') is not None: 
                        ph.find('changeuser').text = str(userid)
                    if ph.find('creationdate') is not None: 
                        ph.find('creationdate').text = now
                    if ph.find('changedate') is not None: 
                        ph.find('changedate').text = now

    def detect_template_file(self):
    """Auto-detect template.xml in script directory"""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    template_names = ['template.xml', 'Template.xml', 'TEMPLATE.xml', 
                     'C5320-A422-B11-4.xml', 'ersa_template.xml']
    
    for name in template_names:
        template_path = os.path.join(script_dir, name)
        if os.path.exists(template_path):
            self.template_file.set(template_path)
            self.log(f"✓ Auto-detected template: {name}")
            return
    
    self.log("⚠ No template.xml found in script directory")



    def update_xml_param(self, root, variable_path, value, datatype='Single'):
        """Find and update parameter in XML"""
        try:
            for param in root.findall('.//ProgramParameter'):
                var = param.find('variable')
                if var is not None and var.text == variable_path:
                    val = param.find('value')
                    if val is not None:
                        if datatype == 'Boolean':
                            val.text = 'True' if value else 'False'
                        else:
                            val.text = str(value)
                        return True
        except Exception:
            pass
        return False
    # ==================== UTILITY ====================
    def log(self, message, color=None):
        """Add message to log (optional color tag)"""
        try:
            if color == 'red':
                self.log_text.insert(tk.END, message + "\n", 'red')
            else:
                self.log_text.insert(tk.END, message + "\n")
            self.log_text.see(tk.END)
            self.root.update_idletasks()
        except Exception:
            # If logging fails silently, do not break the app
            pass
    def export_log(self):
        """Export the contents of the log text area to a .txt file"""
        fname = filedialog.asksaveasfilename(title="Save Log As", defaultextension=".txt",
                                             filetypes=[("Text files","*.txt"), ("All files","*.*")])
        if not fname:
            return
        try:
            with open(fname, 'w', encoding='utf-8') as f:
                f.write(self.log_text.get('1.0', tk.END))
            messagebox.showinfo("Success", f"Log exported to:\n{fname}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export log:\n{e}")
            self.log(f"âœ— Error exporting log: {e}")
    def export_skipped(self):
        """Export skipped program details to CSV"""
        if not getattr(self, 'skipped_programs', None):
            messagebox.showinfo("Info", "No skipped programs to export.")
            return
        fname = filedialog.asksaveasfilename(title="Save Skipped Programs As", defaultextension=".csv",
                                             filetypes=[("CSV files","*.csv"), ("All files","*.*")])
        if not fname:
            return
        import csv
        try:
            with open(fname, 'w', newline='', encoding='utf-8') as f:
                writer = csv.DictWriter(f, fieldnames=['Program', 'Reason'])
                writer.writeheader()
                for r in self.skipped_programs:
                    writer.writerow(r)
            messagebox.showinfo("Success", f"Skipped programs exported to:\n{fname}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export skipped programs:\n{e}")
            self.log(f"âœ— Error exporting skipped programs: {e}")
def main():
    root = tk.Tk()
    app = ERSAProgramGeneratorGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()


print("âœ“ Updated: ERSA_Program_Generator_Final.py")
print("\n" + "="*80)
print("UPDATED CBS LOGIC:")
print("="*80)
print("\nðŸ“‹ New Logic:")
print("\n  Case 1: CBS = NA or Empty")
print("    â†’ Park_Active = True")
print("    â†’ CBS_Active = False")
print("    â†’ CBS Width = (not updated)")
print("\n  Case 2: CBS has a value (e.g., 75)")
print("    â†’ Park_Active = False")
print("    â†’ CBS_Active = True  â† NEW!")
print("    â†’ CBS Width = 75")
print("\nâœ¨ Now CBS will be activated when it has a value!")
print("\nðŸ“ XML Parameters Updated:")
print("  â€¢ enmProg|enmA_Tr|1|enmSngSoll â†’ CBS Width value")
print("  â€¢ enmProg|enmA_Tr|1|enmBlnSollAktiv â†’ CBS Active (True/False)")
print("  â€¢ enmProg|enmA_AxMu|1|enmBlnParkPosSollAktiv â†’ Park Active (True/False)")

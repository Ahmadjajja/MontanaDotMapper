import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import geopandas as gpd
from shapely.geometry import Polygon, Point, box
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import numpy as np
import os
from typing import Dict, List, Tuple, Optional
import re
import sys
from matplotlib.colors import to_rgb

# Montana Dot Map Generator
# This application generates dot maps for Montana using lat/long data
# showing individual specimen locations as dots

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    import sys, os
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

def get_icon_path():
    """Get the path to the application icon"""
    try:
        # First try the .ico file
        if getattr(sys, 'frozen', False):
            # If running as compiled executable
            base_path = sys._MEIPASS
        else:
            # If running as script
            base_path = os.path.dirname(os.path.abspath(__file__))
        
        ico_path = os.path.join(base_path, 'app_icon.ico')
        png_path = os.path.join(base_path, 'app_icon.png')
        
        # Prefer ICO on Windows, PNG on other platforms
        if os.name == 'nt' and os.path.exists(ico_path):
            return ico_path
        elif os.path.exists(png_path):
            return png_path
        else:
            print("Warning: Icon files not found")
            return None
    except Exception as e:
        print(f"Warning: Error getting icon path: {str(e)}")
        return None

class SplashScreen:
    def __init__(self, parent):
        self.parent = parent
        self.splash = tk.Toplevel(parent)
        self.splash.title("Montana Dot Map Generator")
        
        # Set icon
        icon_path = get_icon_path()
        if icon_path and os.path.exists(icon_path):
            try:
                self.splash.iconbitmap(icon_path)
            except Exception as e:
                print(f"Warning: Could not set icon for splash screen: {str(e)}")
        
        # Get screen dimensions
        screen_width = self.splash.winfo_screenwidth()
        screen_height = self.splash.winfo_screenheight()
        
        # Calculate position
        width = 400
        height = 200
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        
        self.splash.geometry(f"{width}x{height}+{x}+{y}")
        self.splash.overrideredirect(True)
        self.splash.configure(bg='white')
        
        # Add loading text
        self.status_label = tk.Label(
            self.splash,
            text="Initializing...",
            bg='white',
            font=('Arial', 12)
        )
        self.status_label.pack(pady=20)
        
        # Add progress bar
        self.progress = ttk.Progressbar(
            self.splash,
            length=300,
            mode='determinate'
        )
        self.progress.pack(pady=20)
        
        self.splash.update()

    def update_status(self, message: str, progress: int = None):
        self.status_label.config(text=message)
        if progress is not None:
            self.progress['value'] = progress
        self.splash.update()

    def destroy(self):
        self.splash.destroy()

class ToastNotification:
    def __init__(self, parent):
        self.parent = parent
        
    def show_toast(self, message: str, duration: int = 3000, error: bool = False):
        toast = tk.Toplevel(self.parent)
        toast.overrideredirect(True)
        
        # Set icon
        icon_path = get_icon_path()
        if icon_path and os.path.exists(icon_path):
            try:
                toast.iconbitmap(icon_path)
            except Exception as e:
                print(f"Warning: Could not set icon for toast: {str(e)}")
        
        # Position toast at bottom right
        toast.geometry(f"+{self.parent.winfo_screenwidth() - 310}+{self.parent.winfo_screenheight() - 100}")
        
        # Configure toast appearance
        bg_color = '#ff4444' if error else '#44aa44'
        frame = tk.Frame(toast, bg=bg_color, padx=10, pady=5)
        frame.pack(fill='both', expand=True)
        
        tk.Label(
            frame,
            text=message,
            bg=bg_color,
            fg='white',
            wraplength=250,
            font=('Arial', 10)
        ).pack()
        
        toast.after(duration, toast.destroy)

class LoadingIndicator:
    def __init__(self, parent, message="Loading..."):
        self.parent = parent
        self.loading_window = tk.Toplevel(parent)
        self.loading_window.title("Loading")
        
        # Set icon
        icon_path = get_icon_path()
        if icon_path:
            try:
                self.loading_window.iconbitmap(icon_path) if icon_path.endswith('.ico') else \
                self.loading_window.iconphoto(True, tk.PhotoImage(file=icon_path))
            except Exception as e:
                print(f"Warning: Could not set icon for loading window: {str(e)}")
        
        # Get screen dimensions
        screen_width = self.loading_window.winfo_screenwidth()
        screen_height = self.loading_window.winfo_screenheight()
        
        # Calculate position
        width = 300
        height = 100
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        
        self.loading_window.geometry(f"{width}x{height}+{x}+{y}")
        self.loading_window.overrideredirect(True)
        self.loading_window.configure(bg='white')
        
        # Add loading text
        self.status_label = tk.Label(
            self.loading_window,
            text=message,
            bg='white',
            font=('Arial', 12)
        )
        self.status_label.pack(pady=(20, 10))
        
        # Add progress bar
        self.progress = ttk.Progressbar(
            self.loading_window,
            length=250,
            mode='indeterminate'
        )
        self.progress.pack(pady=(0, 20))
        
        # Start the progress bar
        self.progress.start(10)
        
        # Make sure the window is on top
        self.loading_window.lift()
        self.loading_window.attributes('-topmost', True)
        
        # Update the window
        self.loading_window.update()

    def update_message(self, message):
        self.status_label.config(text=message)
        self.loading_window.update()
    
    def destroy(self):
        self.progress.stop()
        self.loading_window.destroy()

class SummaryDialog:
    def __init__(self, parent, file_path, data):
        self.parent = parent
        self.window = tk.Toplevel(parent)
        self.window.title("File Upload Success")
        
        # Set icon
        icon_path = get_icon_path()
        if icon_path:
            try:
                self.window.iconbitmap(icon_path) if icon_path.endswith('.ico') else \
                self.window.iconphoto(True, tk.PhotoImage(file=icon_path))
            except Exception as e:
                print(f"Warning: Could not set icon for summary dialog: {str(e)}")
        
        # Get screen dimensions and calculate center position
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        width = 600
        height = 500
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        
        # Set window geometry and properties
        self.window.geometry(f"{width}x{height}+{x}+{y}")
        self.window.configure(bg='#ffffff')
        self.window.resizable(False, False)
        
        # Make window modal
        self.window.transient(parent)
        self.window.grab_set()
        
        # Create canvas with scrollbar
        self.canvas = tk.Canvas(self.window, bg='#ffffff', highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.window, orient="vertical", command=self.canvas.yview)
        
        # Create main container frame that will be scrolled
        main_frame = tk.Frame(self.canvas, bg='#ffffff')
        
        # Configure canvas
        self.canvas.configure(yscrollcommand=scrollbar.set)
        
        # Pack scrollbar and canvas
        scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        
        # Add main frame to canvas
        canvas_frame = self.canvas.create_window((0, 0), window=main_frame, anchor="nw")
        
        # Configure canvas scrolling
        def configure_scroll_region(event):
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        
        def configure_canvas_width(event):
            self.canvas.itemconfig(canvas_frame, width=event.width)
        
        main_frame.bind('<Configure>', configure_scroll_region)
        self.canvas.bind('<Configure>', configure_canvas_width)
        
        # Enable mousewheel scrolling
        def on_mousewheel(event):
            if self.canvas.winfo_exists():
                self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        
        self.mousewheel_binding = self.canvas.bind_all("<MouseWheel>", on_mousewheel)
        
        # Bind window close event
        self.window.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # Content container with padding
        content_frame = tk.Frame(main_frame, bg='#ffffff', padx=30, pady=20)
        content_frame.pack(fill='both', expand=True)
        
        # Success icon and title
        title_frame = tk.Frame(content_frame, bg='#ffffff')
        title_frame.pack(fill='x', pady=(0, 20))
        
        # Success message
        tk.Label(
            title_frame,
            text="✓ File Loaded Successfully",
            font=('Arial', 18, 'bold'),
            bg='#ffffff',
            fg='#2ecc71'  # Green color
        ).pack()
        
        # File info section
        file_frame = tk.Frame(content_frame, bg='#ffffff')
        file_frame.pack(fill='x', pady=(0, 20))
        
        tk.Label(
            file_frame,
            text="FILE INFORMATION",
            font=('Arial', 12, 'bold'),
            bg='#ffffff',
            fg='#2c3e50'
        ).pack(anchor='w')
        
        ttk.Separator(file_frame, orient='horizontal').pack(fill='x', pady=(5, 10))
        
        # File statistics in a grid
        stats_frame = tk.Frame(file_frame, bg='#ffffff')
        stats_frame.pack(fill='x')
        
        stats = [
            ("File Name:", os.path.basename(file_path)),
            ("Total Records:", f"{len(data):,}"),
            ("Year Range:", f"{int(data['year'].min())} - {int(data['year'].max())}"),
            ("Unique Families:", f"{len(data['family'].unique()):,}"),
            ("Unique Genera:", f"{len(data['genus'].unique()):,}"),
            ("Unique Species:", f"{len(data['species'].unique()):,}")
        ]
        
        for i, (label, value) in enumerate(stats):
            row = i // 2
            col = i % 2
            
            stat_container = tk.Frame(stats_frame, bg='#f8f9fa', padx=10, pady=5)
            stat_container.grid(row=row, column=col, padx=5, pady=5, sticky='ew')
            
            tk.Label(
                stat_container,
                text=label,
                font=('Arial', 10),
                bg='#f8f9fa',
                fg='#7f8c8d'
            ).pack(anchor='w')
            
            tk.Label(
                stat_container,
                text=str(value),
                font=('Arial', 11, 'bold'),
                bg='#f8f9fa',
                fg='#2c3e50'
            ).pack(anchor='w')
        
        stats_frame.grid_columnconfigure(0, weight=1)
        stats_frame.grid_columnconfigure(1, weight=1)
        
        # Instructions section
        instructions_frame = tk.Frame(content_frame, bg='#ffffff')
        instructions_frame.pack(fill='x', pady=(20, 0))
        
        tk.Label(
            instructions_frame,
            text="NEXT STEPS",
            font=('Arial', 12, 'bold'),
            bg='#ffffff',
            fg='#2c3e50'
        ).pack(anchor='w')
        
        ttk.Separator(instructions_frame, orient='horizontal').pack(fill='x', pady=(5, 10))
        
        instructions = [
            "1. Select Family, Genus, and Species from the dropdowns",
            "2. Click 'Generate Dot Map' to create a map showing specimen locations",
            "3. Each red dot represents a specimen found at that location",
            "4. The map includes topographic background showing terrain features",
            "5. Use 'Download Dot Map' to save as high-resolution TIFF"
        ]
        
        for instruction in instructions:
            tk.Label(
                instructions_frame,
                text=instruction,
                font=('Arial', 10),
                bg='#ffffff',
                fg='#34495e',
                justify='left',
                anchor='w'
            ).pack(anchor='w', pady=2)
        
        # Close button at bottom
        button_frame = tk.Frame(content_frame, bg='#ffffff')
        button_frame.pack(fill='x', pady=(20, 0))
        
        close_button = tk.Button(
            button_frame,
            text="Got it!",
            font=('Arial', 10, 'bold'),
            bg='#3498db',
            fg='white',
            relief='flat',
            padx=20,
            pady=5,
            command=self.on_closing,
            cursor='hand2'
        )
        close_button.pack()
        
        # Bind hover effects for the button
        close_button.bind('<Enter>', lambda e: close_button.configure(bg='#2980b9'))
        close_button.bind('<Leave>', lambda e: close_button.configure(bg='#3498db'))
        
        # Center the window and make it modal
        self.window.focus_set()
        
        # Ensure the window is on top
        self.window.lift()
        self.window.attributes('-topmost', True)
        self.window.attributes('-topmost', False)
        
        self.window.wait_window()
    
    def on_closing(self):
        """Handle window closing and cleanup"""
        # Unbind the mousewheel event
        self.canvas.unbind_all("<MouseWheel>")
        self.window.destroy()

class MainApplication:
    def __init__(self):
        self.root = tk.Tk()
        self.root.withdraw()  # Hide main window initially
        
        # Set icon
        icon_path = get_icon_path()
        if icon_path:
            try:
                self.root.iconbitmap(icon_path) if icon_path.endswith('.ico') else \
                self.root.iconphoto(True, tk.PhotoImage(file=icon_path))
            except Exception as e:
                print(f"Warning: Could not set icon for main window: {str(e)}")
        
        # Show splash screen
        self.splash = SplashScreen(self.root)
        self.splash.update_status("Loading application...", 0)
        
        # Initialize variables
        self.excel_data = None
        self.montana_counties = None
        self.current_dots = None  # Will store the dot data
        
        # Add variables for species selection
        self.selected_family = tk.StringVar()
        self.selected_genus = tk.StringVar()
        self.selected_species = tk.StringVar()
        
        # Configure main window
        self.root.title("Montana Geographic Distribution Mapper")
        self.root.state('zoomed')  # Start maximized
        
        # Initialize notification system
        self.toast = ToastNotification(self.root)
        
        # Set up the GUI
        self.initialize_gui()
        
        # Destroy splash screen and show main window
        self.splash.destroy()
        self.root.deiconify()

    def initialize_gui(self):
        # Configure style
        style = ttk.Style()
        style.configure('TFrame', background='white')
        style.configure('TLabel', background='white')
        style.configure('TButton', padding=5)
        
        # Main container
        self.main_container = ttk.Frame(self.root)
        self.main_container.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Left panel (inputs)
        self.left_panel = ttk.Frame(self.main_container, style='TFrame')
        self.left_panel.pack(side='left', fill='y', padx=(0, 10))
        
        # Right panel (map display)
        self.right_panel = ttk.Frame(self.main_container, style='TFrame')
        self.right_panel.pack(side='right', fill='both', expand=True)
        
        self._setup_input_fields()
        self._setup_map_display()
        
        # Bind resize event
        self.root.bind('<Configure>', self.on_window_resize)

    def _setup_input_fields(self):
        # File selection
        ttk.Label(self.left_panel, text="Excel File:").pack(anchor='w', pady=(0, 5))
        self.file_frame = ttk.Frame(self.left_panel)
        self.file_frame.pack(fill='x', pady=(0, 20))
        
        self.file_path_var = tk.StringVar()
        ttk.Entry(self.file_frame, textvariable=self.file_path_var, state='readonly').pack(side='left', fill='x', expand=True)
        ttk.Button(self.file_frame, text="Browse", command=self.load_excel).pack(side='right', padx=(5, 0))
        
        # Species Selection Section
        species_frame = ttk.LabelFrame(self.left_panel, text="Species Selection", padding="10")
        species_frame.pack(fill='x', pady=(0, 20))
        
        # Family
        ttk.Label(species_frame, text="Family:", style='TLabel').pack(fill='x')
        self.family_dropdown = ttk.Combobox(species_frame, textvariable=self.selected_family, state="readonly")
        self.family_dropdown.pack(fill='x', pady=(0, 10))
        
        # Genus
        ttk.Label(species_frame, text="Genus:", style='TLabel').pack(fill='x')
        self.genus_dropdown = ttk.Combobox(species_frame, textvariable=self.selected_genus, state="readonly")
        self.genus_dropdown.pack(fill='x', pady=(0, 10))
        
        # Species
        ttk.Label(species_frame, text="Species:", style='TLabel').pack(fill='x')
        self.species_dropdown = ttk.Combobox(species_frame, textvariable=self.selected_species, state="readonly")
        self.species_dropdown.pack(fill='x', pady=(0, 10))
        
        # Dot Color Selection
        ttk.Label(species_frame, text="Dot Color:", style='TLabel').pack(fill='x')
        self.dot_color_var = tk.StringVar(value="#ff0000")  # Default red
        color_frame = ttk.Frame(species_frame)
        color_frame.pack(fill='x', pady=(0, 10))
        
        ttk.Entry(color_frame, textvariable=self.dot_color_var, width=10).pack(side='left')
        ttk.Button(color_frame, text="Choose Color", command=self.choose_color).pack(side='right', padx=(5, 0))
        
        # Map options
        options_frame = ttk.LabelFrame(self.left_panel, text="Map Options", padding="10")
        options_frame.pack(fill='x', pady=(20, 10))
        
        # County lines checkbox
        self.show_county_lines = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Show County Lines", variable=self.show_county_lines).pack(anchor='w')
        
        # Action buttons
        ttk.Button(self.left_panel, text="Generate Dot Map", command=self.generate_dot_map).pack(fill='x', pady=(10, 5))
        ttk.Button(self.left_panel, text="Download Dot Map", command=self.download_map).pack(fill='x', pady=(5, 0))
        
        # Bind dropdowns
        self.family_dropdown.bind("<<ComboboxSelected>>", self.update_genus_dropdown)
        self.genus_dropdown.bind("<<ComboboxSelected>>", self.update_species_dropdown)

    def choose_color(self):
        """Open color picker dialog and update the color variable"""
        from tkinter import colorchooser
        color = colorchooser.askcolor(title="Choose Dot Color", color=self.dot_color_var.get())
        if color[1]:  # If a color was selected
            self.dot_color_var.set(color[1])

    def _setup_map_display(self):
        self.figure = Figure(figsize=(10, 8))
        self.ax = self.figure.add_subplot(111)
        # Remove the box from initial display
        self.ax.set_frame_on(False)
        self.ax.set_xticks([])
        self.ax.set_yticks([])
        self.canvas = FigureCanvasTkAgg(self.figure, master=self.right_panel)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(fill='both', expand=True)

    def load_excel(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if not file_path:
            return
            
        try:
            # Show loading indicator
            loading = LoadingIndicator(self.root, "Loading Excel file...")
            
            self.excel_data = pd.read_excel(file_path)
            required_columns = ['lat', 'lat_dir', 'long', 'long_dir', 'family', 'genus', 'species', 'year']
            if not all(col in self.excel_data.columns for col in required_columns):
                loading.destroy()
                raise ValueError("Excel file must contain 'lat', 'lat_dir', 'long', 'long_dir', 'family', 'genus', 'species', and 'year' columns")
                
            self.file_path_var.set(file_path)
            
            # Process the data
            loading.update_message("Processing data...")
            for col in ["family", "genus", "species"]:
                self.excel_data[col] = self.excel_data[col].astype(str).str.strip().str.lower()
            
            # Convert year to numeric, handling any non-numeric values
            self.excel_data['year'] = pd.to_numeric(self.excel_data['year'], errors='coerce')
            
            # Get valid families (non-empty/non-null values)
            loading.update_message("Updating dropdowns...")
            valid_families = sorted(self.excel_data["family"].dropna().unique())
            valid_families = [f for f in valid_families if str(f).strip() and str(f).lower() != 'nan']  # Remove empty strings and 'nan'
            
            # Capitalize family names
            family_values = ["All"] + [f.title() for f in valid_families]
            
            # Update Family dropdown
            self.family_dropdown["values"] = family_values
            self.family_dropdown.set("Select Family")
            
            # Reset other dropdowns
            self.genus_dropdown.set("Select Genus")
            self.genus_dropdown["values"] = []
            self.species_dropdown.set("Select Species")
            self.species_dropdown["values"] = []
            
            # Load Montana counties
            loading.update_message("Loading Montana counties...")
            all_counties = gpd.read_file(resource_path("shapefiles/cb_2021_us_county_5m.shp"))
            self.montana_counties = all_counties[all_counties['STATEFP'] == '30']
            self.montana_counties = self.montana_counties.to_crs("EPSG:32100")
            
            # Bind dropdowns
            self.family_dropdown.bind("<<ComboboxSelected>>", self.update_genus_dropdown)
            self.genus_dropdown.bind("<<ComboboxSelected>>", self.update_species_dropdown)
            
            # Destroy loading indicator
            loading.destroy()
            
            # Show summary dialog
            SummaryDialog(self.root, file_path, self.excel_data)
            
            self.toast.show_toast("Excel file loaded successfully")
            
        except Exception as e:
            if 'loading' in locals():
                loading.destroy()
            self.toast.show_toast(f"Error loading file: {str(e)}", error=True)

    def dms_to_decimal(self, coord):
        """
        Convert a coordinate in DMS format (e.g., '44°41.576'') to decimal degrees.
        Handles both unicode and ascii degree/minute/second symbols.
        """
        if isinstance(coord, float) or isinstance(coord, int):
            return float(coord)
        if not isinstance(coord, str):
            return float('nan')
        # Remove unwanted characters and normalize
        coord = coord.replace("'", "'").replace("″", '"').replace("""", '"').replace(""", '"')
        dms_pattern = r"(\d+)[°\s]+(\d+(?:\.\d+)?)[\'′]?\s*(\d*(?:\.\d+)?)[\"″]?"
        match = re.match(dms_pattern, coord.strip())
        if match:
            deg = float(match.group(1))
            min_ = float(match.group(2))
            sec = float(match.group(3)) if match.group(3) else 0.0
            return deg + min_ / 60 + sec / 3600
        try:
            return float(coord)
        except Exception:
            return float('nan')

    def convert_coordinates(self, row):
        """Convert coordinates taking into account direction (N/S, E/W) and DMS/decimal formats"""
        try:
            # Handle potential data type issues
            lat_val = row['lat']
            long_val = row['long']
            lat_dir_val = row['lat_dir']
            long_dir_val = row['long_dir']
            
            # Skip if coordinates are missing or invalid
            if pd.isna(lat_val) or pd.isna(long_val):
                return Point(0, 0)  # Will be filtered out
            
            # Convert coordinates
            lat = self.dms_to_decimal(lat_val)
            long = self.dms_to_decimal(long_val)
            
            # Skip if conversion failed
            if pd.isna(lat) or pd.isna(long):
                return Point(0, 0)  # Will be filtered out
            
            # Handle direction values more robustly
            lat_dir = 'N'  # Default for Montana
            long_dir = 'W'  # Default for Montana
            
            # Only process direction if it's a valid string
            if pd.notna(lat_dir_val) and isinstance(lat_dir_val, str):
                lat_dir = lat_dir_val.strip().upper()
                if lat_dir not in ['N', 'S']:
                    lat_dir = 'N'  # Default to North for Montana
            
            if pd.notna(long_dir_val) and isinstance(long_dir_val, str):
                long_dir = long_dir_val.strip().upper()
                if long_dir not in ['E', 'W']:
                    long_dir = 'W'  # Default to West for Montana
            
            # Adjust for direction
            if lat_dir == 'S':  # If Southern hemisphere
                lat = -lat
            if long_dir == 'W':  # If Western hemisphere
                long = -long
            
            # Montana is roughly between 44°N to 49°N and 104°W to 116°W
            # Validate the coordinates are somewhat reasonable
            if not (44 <= abs(lat) <= 49 and 104 <= abs(long) <= 116):
                # Don't print warning for every coordinate, just return invalid point
                return Point(0, 0)  # Will be filtered out
            
            return Point(long, lat)
        except Exception as e:
            # Don't print error for every coordinate, just return invalid point
            return Point(0, 0)  # Will be filtered out

    def generate_dot_map(self):
        if self.excel_data is None:
            self.toast.show_toast("Please load an Excel file first", error=True)
            return
            
        if self.montana_counties is None:
            self.toast.show_toast("Please load an Excel file first to initialize county data", error=True)
            return
            
        try:
            loading = LoadingIndicator(self.root, "Generating dot map...")
            
            # Validate required columns
            required_columns = ['lat', 'lat_dir', 'long', 'long_dir', 'family', 'genus', 'species', 'year']
            if not all(col in self.excel_data.columns for col in required_columns):
                loading.destroy()
                raise ValueError("Excel file must contain 'lat', 'lat_dir', 'long', 'long_dir', 'family', 'genus', 'species', and 'year' columns")
            
            # Get species selection
            fam = self.selected_family.get().strip()
            gen = self.selected_genus.get().strip()
            spec = self.selected_species.get().strip()
            
            if not fam or fam == "Select Family" or not gen or gen == "Select Genus" or not spec or spec == "Select Species":
                loading.destroy()
                messagebox.showerror("Missing Input", "Please select Family, Genus, and Species.")
                return
            
            loading.update_message("Filtering data...")
            
            # Filter data based on species selection
            filtered = self.excel_data.copy()
            
            if fam == "All":
                filtered = filtered[filtered["family"].notna() & (filtered["family"].str.strip() != "")]
            else:
                filtered = filtered[filtered["family"].str.lower() == fam.lower()]
                
            if gen == "All":
                filtered = filtered[filtered["genus"].notna() & (filtered["genus"].str.strip() != "")]
            else:
                filtered = filtered[filtered["genus"].str.lower() == gen.lower()]
                
            if spec == "all":
                filtered = filtered[filtered["species"].notna() & (filtered["species"].str.strip() != "")]
            else:
                filtered = filtered[filtered["species"].str.lower() == spec.lower()]
            
            if len(filtered) == 0:
                loading.destroy()
                self.toast.show_toast("No data found for selected species", error=True)
                return
            
            loading.update_message("Converting coordinates...")
            
            # Convert coordinates to points
            geometries = filtered.apply(self.convert_coordinates, axis=1)
            points = gpd.GeoDataFrame(
                filtered,
                geometry=geometries,
                crs="EPSG:4326"
            )
            
            # Filter points within Montana
            montana_boundary = self.montana_counties.dissolve().to_crs("EPSG:4326").geometry.iloc[0]
            points = points[points.geometry.within(montana_boundary)]
            
            if len(points) == 0:
                loading.destroy()
                self.toast.show_toast("No points found within Montana's boundaries", error=True)
                return
            
            # Store the filtered points data
            self.current_dots = {
                'points': points,
                'species_info': f"{fam} > {gen} > {spec}",
                'count': len(points)
            }
            
            # Display the dot map
            loading.update_message("Rendering dot map...")
            self.display_dot_map()
            
            loading.destroy()
            self.toast.show_toast(f"Dot map generated with {len(points)} specimens")
            
        except Exception as e:
            if 'loading' in locals():
                loading.destroy()
            self.toast.show_toast(f"Error generating dot map: {str(e)}", error=True)

    def display_dot_map(self):
        """Display simple dot map with Montana counties and specimen locations"""
        if self.current_dots is None:
            return
        
        # Clear the figure
        self.figure.clf()
        
        # Create single subplot
        self.ax = self.figure.add_subplot(111)
        
        # Configure axis
        self.ax.set_frame_on(False)
        self.ax.set_xticks([])
        self.ax.set_yticks([])
        self.ax.set_aspect('equal')
        
        # Get points data and convert to Web Mercator projection
        points = self.current_dots['points']
        points_web_mercator = points.to_crs(epsg=3857)
        
        # Get Montana counties in Web Mercator projection
        counties_web_mercator = self.montana_counties.to_crs(epsg=3857)
        
        # Get bounds for the map
        bounds = counties_web_mercator.total_bounds
        padding = (bounds[2] - bounds[0]) * 0.05  # Small padding
        
        # Set the plot bounds
        self.ax.set_xlim([bounds[0] - padding, bounds[2] + padding])
        self.ax.set_ylim([bounds[1] - padding, bounds[3] + padding])
        
        # Simple color scheme
        colors = {
            'county_border': '#000000',      # Black county borders
            'county_fill': '#f8f9fa',        # Light gray fill for counties
            'dots': self.dot_color_var.get(), # User-selected dot color
            'text': '#000000'                # Black text
        }
        
        # Plot county boundaries with simple styling (only if checkbox is checked)
        if self.show_county_lines.get():
            for idx, county in counties_web_mercator.iterrows():
                self.ax.fill(county.geometry.exterior.xy[0], 
                            county.geometry.exterior.xy[1],
                            facecolor=colors['county_fill'],     # Light gray fill
                            edgecolor=colors['county_border'],   # Black borders
                            linewidth=0.8,
                            alpha=1.0,
                            zorder=5)
        else:
            # Show Montana's outer boundary when county lines are hidden
            montana_boundary = counties_web_mercator.dissolve().geometry.iloc[0]
            self.ax.fill(montana_boundary.exterior.xy[0], 
                        montana_boundary.exterior.xy[1],
                        facecolor='white',                      # White fill
                        edgecolor=colors['county_border'],      # Black outer border
                        linewidth=1.5,                         # Slightly thicker border
                        alpha=1.0,
                        zorder=5)
        
        # Extract coordinates for plotting dots
        x_coords = [point.x for point in points_web_mercator.geometry]
        y_coords = [point.y for point in points_web_mercator.geometry]
        
        # Plot dots with simple styling
        self.ax.scatter(x_coords, y_coords, 
                       c=colors['dots'],     # Red dots
                       s=25,                 # Appropriate size
                       alpha=1.0,           # Full opacity
                       edgecolors='none',   # No border
                       linewidth=0,         # No line
                       zorder=15)           # Ensure dots are on top
        
        # Add title
        species_info = self.current_dots['species_info']
        if species_info:
            family, genus, species = species_info.split(' > ')
            title = f"Known geographic distribution of {genus} {species} in Montana"
            self.figure.suptitle(title, x=0.5, y=0.98, 
                               ha='center', va='top',
                               fontsize=12, fontweight='normal',
                               color=colors['text'],
                               style='italic')
        
        # Add legend
        import matplotlib.patches as mpatches
        legend_elements = [
            mpatches.Patch(facecolor=colors['dots'], 
                          edgecolor='none',
                          label=f'Specimen Location ({self.current_dots["count"]} total)')
        ]
        
        legend = self.ax.legend(handles=legend_elements,
                              loc='lower right',
                              frameon=False,
                              fontsize=10,
                              title_fontsize=10)
        
        # Add north arrow
        self.ax.annotate('N', xy=(0.05, 0.95), xycoords='axes fraction',
                        fontsize=14, fontweight='bold',
                        color=colors['text'],
                        ha='center', va='center')
        
        # Add scale bar (accurately calculated for 100 km)
        # Calculate the actual distance in meters for 100 km at Montana's latitude
        # Montana is roughly at 47°N latitude
        import math
        lat_rad = math.radians(47)  # Montana's approximate latitude
        # Web Mercator projection scale factor at this latitude
        scale_factor = 1 / math.cos(lat_rad)
        
        # Calculate 100 km in Web Mercator units
        # 1 degree longitude ≈ 111,320 meters at equator
        # At 47°N, 1 degree longitude ≈ 111,320 * cos(47°) meters
        meters_per_degree = 111320 * math.cos(lat_rad)
        km_100_in_degrees = 100000 / meters_per_degree
        
        # Convert to Web Mercator projection units
        scale_length_meters = km_100_in_degrees * 111320 * scale_factor
        
        # Position scale bar in bottom-left corner
        scale_x = bounds[0] + (bounds[2] - bounds[0]) * 0.05  # 5% from left edge
        scale_y = bounds[1] + (bounds[3] - bounds[1]) * 0.05  # 5% from bottom edge
        
        # Draw the scale bar
        self.ax.plot([scale_x, scale_x + scale_length_meters], [scale_y, scale_y], 
                    color=colors['text'], linewidth=2)
        self.ax.text(scale_x + scale_length_meters/2, scale_y - (bounds[3] - bounds[1]) * 0.02,
                    '100 km', ha='center', va='top',
                    fontsize=8, color=colors['text'])
        
        # Adjust layout
        self.figure.subplots_adjust(left=0.05, right=0.95, 
                                  bottom=0.05, top=0.92)
        
        # Draw the canvas
        self.canvas.draw()



    def download_map(self):
        if self.current_dots is None:
            self.toast.show_toast("Please generate dot map first", error=True)
            return
            
        try:
            import datetime
            from pathlib import Path
            import os
            
            # Get Downloads folder path
            downloads_path = str(Path.home() / "Downloads")
            
            # Get current date and time in the desired format
            now = datetime.datetime.now()
            timestamp = now.strftime("%I_%M_%p_%m_%d_%Y")  # e.g., 12_49_PM_6_12_2025
            
            # Create a meaningful filename
            filename = f"MontanaDotMap_{timestamp}.tiff"
            file_path = os.path.join(downloads_path, filename)
            
            # Save the figure
            self.figure.savefig(file_path, format="tiff", dpi=300, bbox_inches='tight')
            
            # Show toast notification
            self.toast.show_toast(f"Dot map saved as {filename}")
            
            print(f"✅ TIFF dot map saved as '{file_path}'")
            
        except Exception as e:
            messagebox.showerror("Error", 
                f"Error saving file:\n{str(e)}\n\n"
                "Please try again."
            )

    def on_window_resize(self, event=None):
        # Update the figure size to match the panel size
        w = self.right_panel.winfo_width() / 100
        h = self.right_panel.winfo_height() / 100
        self.figure.set_size_inches(w, h)
        
        # If we have maps displayed, redraw them to maintain proper layout
        if self.current_dots is not None:
            self.display_dot_map()
        else:
            self.canvas.draw()

    def update_genus_dropdown(self, event=None):
        family = self.selected_family.get().strip()
        
        if family == "Select Family":
            self.genus_dropdown["values"] = []
            self.genus_dropdown.set("Select Genus")
            return
        
        # Filter based on family selection
        if family == "All":
            # Get all non-empty genus values
            filtered = self.excel_data[self.excel_data["genus"].notna() & (self.excel_data["genus"].str.strip() != "")]
        else:
            # Get genus for specific family (case-insensitive)
            filtered = self.excel_data[self.excel_data["family"].str.lower() == family.lower()]
        
        # Get valid genera (non-empty/non-null values)
        valid_genera = sorted(filtered["genus"].dropna().unique())
        valid_genera = [g for g in valid_genera if str(g).strip() and str(g).lower() != 'nan']  # Remove empty strings and 'nan'
        
        # Create genus list with special options
        genus_values = ["All"] + [g.title() for g in valid_genera]
        
        # Update Genus dropdown
        self.genus_dropdown["values"] = genus_values
        self.genus_dropdown.set("Select Genus")
        
        # Reset species dropdown
        self.species_dropdown.set("Select Species")
        self.species_dropdown["values"] = []
    
    def update_species_dropdown(self, event=None):
        family = self.selected_family.get().strip()
        genus = self.selected_genus.get().strip()
        
        if family == "Select Family" or genus == "Select Genus":
            self.species_dropdown["values"] = []
            self.species_dropdown.set("Select Species")
            return
        
        # Start with base DataFrame
        filtered = self.excel_data
        
        # Apply family filter
        if family == "All":
            filtered = filtered[filtered["family"].notna() & (filtered["family"].str.strip() != "")]
        else:
            filtered = filtered[filtered["family"].str.lower() == family.lower()]
        
        # Apply genus filter
        if genus == "All":
            filtered = filtered[filtered["genus"].notna() & (filtered["genus"].str.strip() != "")]
        else:
            filtered = filtered[filtered["genus"].str.lower() == genus.lower()]
        
        # Get valid species (non-empty/non-null values)
        valid_species = sorted(filtered["species"].dropna().unique())
        valid_species = [s for s in valid_species if str(s).strip() and str(s).lower() != 'nan']  # Remove empty strings and 'nan'
        
        # Create species list with special options - note lowercase for species
        species_values = ["all"] + valid_species
        
        # Update Species dropdown
        self.species_dropdown["values"] = species_values
        self.species_dropdown.set("Select Species")

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    if getattr(sys, 'frozen', False):
        base = sys._MEIPASS
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    os.environ['GDAL_DATA'] = os.path.join(base, 'gdal-data')
    os.environ['PROJ_LIB'] = os.path.join(base, 'proj')
    app = MainApplication()
    app.run() 
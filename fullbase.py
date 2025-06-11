import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
import os
import datetime
from collections import defaultdict
import subprocess
import sys
import json

def install_and_import(package, import_name=None):
    import importlib
    try:
        return importlib.import_module(import_name or package)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        return importlib.import_module(import_name or package)

# --------- AGENT/BOM DATA/DEFAULTS ----------
AGENT_TABLES = {
    "FM-200 (HFC-227ea)": {6.25: 0.47, 7.00: 0.58, 8.00: 0.70, 9.00: 0.85, 10.00: 1.00},
    "Novec 1230 (FK-5-1-12)": {4.5: 0.41, 5.0: 0.47, 5.5: 0.52, 6.0: 0.57, 6.5: 0.62, 7.0: 0.67},
    "High Pressure CO2": {34.0: 0.612}
}
AGENT_DEFAULTS = {
    "FM-200 (HFC-227ea)": {"design_concentration": 7.0, "temperature": 20.0, "altitude": 0.0},
    "Novec 1230 (FK-5-1-12)": {"design_concentration": 5.0, "temperature": 20.0, "altitude": 0.0},
    "High Pressure CO2": {"design_concentration": 34.0, "temperature": 20.0, "altitude": 0.0}
}
REFERENCE_LINKS = [
    ("NFPA 2001 Standard", "https://www.nfpa.org/codes-and-standards/all-codes-and-standards/list-of-codes-and-standards/detail?code=2001"),
    ("3M Novec 1230 Guide", "https://multimedia.3m.com/mws/media/753982O/3m-novec-1230-fire-protection-fluid-engineering-guide.pdf"),
    ("Kidde CO2 Manual", "https://kidde-fenwal.com/Lists/TechnicalManuals/CO2_Total_Flooding_System_Design_Manual.pdf"),
]
def altitude_correction(altitude_m):
    pressure = 101.325 * (1 - 2.25577e-5 * altitude_m)**5.25588
    return 101.325 / pressure
def temperature_correction(temperature_C):
    return 1.0 if temperature_C <= 20 else 1 + 0.013 * (temperature_C - 20)

def get_viking_bom(agent_kg, actuation_type):
    if agent_kg <= 52: cylinder = "889099"
    elif agent_kg <= 106: cylinder = "889101"
    elif agent_kg <= 147: cylinder = "889104"
    elif agent_kg <= 180: cylinder = "910509"
    else: cylinder = "910510"
    actuator_dict = {"Electrical": ("07-235070-001", "Electrically Operated Actuator"),
                     "Pneumatic": ("07-235070-002", "Pneumatic Actuator"),
                     "Manual": ("07-235070-003", "Manual Actuator")}
    actuator = actuator_dict.get(actuation_type, actuator_dict["Electrical"])
    return [
        {"part_number": cylinder, "description": "Viking FM-200 Cylinder", "qty": 1, "unit": "pcs"},
        {"part_number": "07-235068-001", "description": "Cylinder Valve", "qty": 1, "unit": "pcs"},
        {"part_number": actuator[0], "description": actuator[1], "qty": 1, "unit": "pcs"},
        {"part_number": "07-235098-001", "description": "Discharge Nozzle", "qty": 1, "unit": "pcs"},
    ]
def get_kidde_bom(agent_kg, actuation_type):
    if agent_kg <= 40: cylinder = ("06-236204-001", "Kidde 40L FM-200 Cylinder")
    elif agent_kg <= 106: cylinder = ("06-236212-001", "Kidde 106L FM-200 Cylinder")
    else: cylinder = ("06-236214-001", "Kidde 180L FM-200 Cylinder")
    actuator_dict = {"Electrical": ("06-236240-001", "Electrically Operated Actuator"),
                     "Pneumatic": ("06-236241-001", "Pneumatic Actuator"),
                     "Manual": ("06-236242-001", "Manual Actuator")}
    actuator = actuator_dict.get(actuation_type, actuator_dict["Electrical"])
    return [
        {"part_number": cylinder[0], "description": cylinder[1], "qty": 1, "unit": "pcs"},
        {"part_number": "06-236230-001", "description": "Kidde FM-200 Valve", "qty": 1, "unit": "pcs"},
        {"part_number": actuator[0], "description": actuator[1], "qty": 1, "unit": "pcs"},
        {"part_number": "06-236250-001", "description": "Discharge Nozzle", "qty": 1, "unit": "pcs"},
    ]
def get_hygood_bom(agent_kg, actuation_type):
    if agent_kg <= 8: part_number, desc = "303.205.015", "Hygood FM-200 8L Cylinder"
    elif agent_kg <= 16: part_number, desc = "303.205.016", "Hygood FM-200 16L Cylinder"
    elif agent_kg <= 32: part_number, desc = "303.205.017", "Hygood FM-200 32L Cylinder"
    elif agent_kg <= 52: part_number, desc = "303.205.018", "Hygood FM-200 52L Cylinder"
    elif agent_kg <= 106: part_number, desc = "303.205.019", "Hygood FM-200 106L Cylinder"
    elif agent_kg <= 147: part_number, desc = "303.205.020", "Hygood FM-200 147L Cylinder"
    elif agent_kg <= 180: part_number, desc = "303.205.021", "Hygood FM-200 180L Cylinder"
    else: part_number, desc = "303.205.022", "Hygood FM-200 343L Cylinder"
    actuator_dict = {
        "Electrical": ("304.205.010", "Electrical Actuator (Suppression Diode)"),
        "Pneumatic": ("304.209.004", "Pneumatic Actuator"),
        "Manual": ("304.209.002", "Manual Actuator")
    }
    actuator = actuator_dict.get(actuation_type, actuator_dict["Electrical"])
    return [
        {"part_number": part_number, "description": desc, "qty": 1, "unit": "pcs"},
        {"part_number": "302.209.002", "description": "Cylinder Valve (2” for 8-180L)", "qty": 1, "unit": "pcs"},
        {"part_number": actuator[0], "description": actuator[1], "qty": 1, "unit": "pcs"},
        {"part_number": "306.207.003", "description": "Flexible Discharge Hose (2”)", "qty": 1, "unit": "pcs"},
        {"part_number": "306.205.005", "description": "Discharge Nozzle (typical)", "qty": 1, "unit": "pcs"},
    ]
OEM_BOM_DATABASES = {"Viking": {"func": get_viking_bom},
                     "Kidde": {"func": get_kidde_bom},
                     "Tyco Hygood": {"func": get_hygood_bom}}
def aggregate_bom(bom_lists):
    agg = defaultdict(lambda: {"description": "", "qty": 0, "unit": ""})
    for bom in bom_lists:
        for item in bom:
            k = item["part_number"]
            agg[k]["description"] = item["description"]
            agg[k]["qty"] += item["qty"]
            agg[k]["unit"] = item["unit"]
    return [{"part_number": k, "description": v["description"], "qty": v["qty"], "unit": v["unit"]}
            for k, v in agg.items()]

# --------- CALCULATION ENGINE ----------
class Room:
    def __init__(self, name, length, width, height, design_concentration, altitude, temperature,
                 units="metric", agent="FM-200 (HFC-227ea)", actuation_type="Electrical", oem="Viking"):
        self.name = name
        self.length = length
        self.width = width
        self.height = height
        self.design_concentration = design_concentration
        self.altitude = altitude
        self.temperature = temperature
        self.units = units
        self.agent = agent
        self.actuation_type = actuation_type
        self.oem = oem
    def volume(self):
        if self.units == "imperial":
            return self.length * 0.3048 * self.width * 0.3048 * self.height * 0.3048
        return self.length * self.width * self.height
    def to_dict(self):
        return self.__dict__.copy()
    @staticmethod
    def from_dict(data):
        return Room(**data)
    def calculate_required_agent(self):
        vol = self.volume()
        agent_table = AGENT_TABLES[self.agent]
        if self.agent == "High Pressure CO2":
            factor = agent_table.get(self.design_concentration, 0.612)
            required = vol * factor
            alt_corr = temp_corr = 1.0
        else:
            factor = self.get_agent_factor(self.design_concentration, agent_table)
            alt_corr = altitude_correction(self.altitude)
            temp_corr = temperature_correction(self.temperature)
            required = vol * factor * alt_corr * temp_corr
        return required, factor, alt_corr, temp_corr, vol
    @staticmethod
    def get_agent_factor(design_concentration, agent_table):
        keys = sorted(agent_table.keys())
        if design_concentration <= keys[0]: return agent_table[keys[0]]
        if design_concentration >= keys[-1]: return agent_table[keys[-1]]
        for i in range(len(keys) - 1):
            if keys[i] <= design_concentration <= keys[i + 1]:
                x0, x1 = keys[i], keys[i + 1]
                y0, y1 = agent_table[x0], agent_table[x1]
                return y0 + (design_concentration - x0) * (y1 - y0) / (x1 - x0)
        raise KeyError("Design concentration out of range for interpolation.")

class AgentCalculator:
    def __init__(self): self.rooms = []
    def add_room(self, room):
        if any(r.name == room.name for r in self.rooms): raise ValueError("Room name must be unique.")
        self.rooms.append(room)
    def remove_room(self, room_name): self.rooms = [room for room in self.rooms if room.name != room_name]
    def clear(self): self.rooms.clear()

class ToolTip(object):
    def __init__(self, widget, text):
        self.widget = widget; self.text = text; self.tipwindow = None
        widget.bind("<Enter>", self.showtip); widget.bind("<Leave>", self.hidetip)
    def showtip(self, event=None):
        if self.tipwindow: return
        x, y, _, _ = self.widget.bbox("insert") or (0, 0, 0, 0)
        x = x + self.widget.winfo_rootx() + 40
        y = y + self.widget.winfo_rooty() + 20
        self.tipwindow = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True); tw.geometry(f"+{x}+{y}")
        label = tk.Label(tw, text=self.text, background="#ffffe0", relief=tk.SOLID, borderwidth=1, font=("tahoma", "9", "normal")); label.pack(ipadx=1)
    def hidetip(self, event=None):
        tw = self.tipwindow; self.tipwindow = None;  tw and tw.destroy()

# --------- GUI ----------
class ModernFM200App:
    def __init__(self, master):
        self.master = master
        self.calculator = AgentCalculator()
        self.bom_per_room = []
        self.project_bom = []
        self.latest_results = ""
        self.setup_style(); self.setup_ui(); self.load_rooms_to_treeview(); self.status("Ready.")
    def setup_style(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TNotebook.Tab", font=("Segoe UI", 12, "bold"), padding=[15, 8])
        style.configure("TLabel", font=("Segoe UI", 11))
        style.configure("TEntry", font=("Segoe UI", 11))
        style.configure("TButton", font=("Segoe UI", 11, "bold"), padding=6)
        style.configure("Treeview.Heading", font=("Segoe UI", 11, "bold"))
        style.configure("Treeview", font=("Segoe UI", 10), rowheight=28)
        style.map("TButton", foreground=[('pressed', 'white'), ('active', '#2E8B57')], background=[('pressed', '#1E3D59'), ('active', '#70A1FF')])
    def setup_ui(self):
        self.master.title("Clean Agent Calculator")
        self.master.geometry("1300x900")
        self.tabs = ttk.Notebook(self.master); self.tabs.pack(fill=tk.BOTH, expand=True, padx=10, pady=6)
        # TAB 1: INPUT
        input_tab = ttk.Frame(self.tabs); self.tabs.add(input_tab, text="Add Room")
        input_frame = ttk.Frame(input_tab, padding=18); input_frame.pack(fill=tk.BOTH, expand=True)
        col1 = ttk.Frame(input_frame); col2 = ttk.Frame(input_frame)
        col1.pack(side=tk.LEFT, fill=tk.Y, expand=True, padx=15); col2.pack(side=tk.RIGHT, fill=tk.Y, expand=True, padx=15)
        ttk.Label(col1, text="Room Name:").pack(anchor=tk.W, pady=3); self.room_name_entry = ttk.Entry(col1, width=24); self.room_name_entry.pack(pady=3, anchor=tk.W)
        ToolTip(self.room_name_entry, "Room name must be unique.")
        for lbl, ent in [("Length (m):", "length_entry"), ("Width (m):", "width_entry"), ("Height (m):", "height_entry")]:
            ttk.Label(col1, text=lbl).pack(anchor=tk.W, pady=3); setattr(self, ent, ttk.Entry(col1, width=24)); getattr(self, ent).pack(pady=3, anchor=tk.W)
        ttk.Label(col1, text="Units:").pack(anchor=tk.W, pady=3); self.unit_var = tk.StringVar(value="metric")
        unit_select = ttk.Combobox(col1, textvariable=self.unit_var, values=["metric", "imperial"], width=22, state="readonly"); unit_select.pack(pady=3, anchor=tk.W)
        ToolTip(unit_select, "Metric = meters, Imperial = feet.")
        ttk.Label(col1, text="Design Concentration (%):").pack(anchor=tk.W, pady=3)
        self.design_conc_entry = ttk.Entry(col1, width=24); self.design_conc_entry.pack(pady=3, anchor=tk.W)
        ttk.Label(col1, text="Altitude (m):").pack(anchor=tk.W, pady=3); self.altitude_entry = ttk.Entry(col1, width=24); self.altitude_entry.pack(pady=3, anchor=tk.W)
        ttk.Label(col1, text="Temperature (°C):").pack(anchor=tk.W, pady=3); self.temperature_entry = ttk.Entry(col1, width=24); self.temperature_entry.pack(pady=3, anchor=tk.W)
        # AGENT TYPE + OEM
        ttk.Label(col2, text="Agent Type:").pack(anchor=tk.W, pady=3); self.room_agent_var = tk.StringVar(value="FM-200 (HFC-227ea)")
        self.room_agent_select = ttk.Combobox(col2, textvariable=self.room_agent_var, values=list(AGENT_TABLES.keys()), width=24, state="readonly")
        self.room_agent_select.pack(pady=3, anchor=tk.W); self.room_agent_select.bind("<<ComboboxSelected>>", self.set_defaults_for_agent)
        ttk.Label(col2, text="Actuation Type:").pack(anchor=tk.W, pady=3)
        self.actuation_type_var = tk.StringVar(value="Electrical")
        self.actuation_select = ttk.Combobox(col2, textvariable=self.actuation_type_var, values=["Electrical", "Pneumatic", "Manual"], width=24, state="readonly")
        self.actuation_select.pack(pady=3, anchor=tk.W)
        ttk.Label(col2, text="System OEM (FM-200):").pack(anchor=tk.W, pady=3)
        self.oem_var = tk.StringVar(value="Viking")
        self.oem_select = ttk.Combobox(col2, textvariable=self.oem_var, values=list(OEM_BOM_DATABASES.keys()), width=24, state="readonly")
        self.oem_select.pack(pady=3, anchor=tk.W)
        ttk.Label(col2, text="Project Name:").pack(anchor=tk.W, pady=3); self.project_name_entry = ttk.Entry(col2, width=26); self.project_name_entry.pack(pady=3, anchor=tk.W)
        ttk.Label(col2, text="Customer Name:").pack(anchor=tk.W, pady=3); self.customer_name_entry = ttk.Entry(col2, width=26); self.customer_name_entry.pack(pady=3, anchor=tk.W)
        ttk.Button(col2, text="Add Room", command=self.add_room).pack(pady=8, anchor=tk.W)
        ttk.Button(col2, text="Clear All Fields", command=self.clear_inputs).pack(pady=4, anchor=tk.W)
        ttk.Button(col2, text="Import Rooms from Excel", command=self.import_from_excel).pack(pady=4, anchor=tk.W)
        # TAB 2: ROOM LIST
        rooms_tab = ttk.Frame(self.tabs); self.tabs.add(rooms_tab, text="Rooms List")
        list_frame = ttk.Frame(rooms_tab, padding=10); list_frame.pack(fill=tk.BOTH, expand=True)
        columns = ("Name", "Length", "Width", "Height", "Units", "Design %", "Altitude", "Temp", "Agent", "Actuation", "OEM")
        self.treeview = ttk.Treeview(list_frame, columns=columns, show="headings", height=18); self.treeview.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        for col in columns:
            self.treeview.heading(col, text=col); self.treeview.column(col, width=90 if col == "Name" else 70, anchor=tk.CENTER)
        ttk.Button(list_frame, text="Remove Selected", command=self.remove_room).pack(pady=3)
        ttk.Button(list_frame, text="Save Project", command=self.save_project).pack(pady=3)
        ttk.Button(list_frame, text="Load Project", command=self.open_project).pack(pady=3)
        # TAB 3: CALC & BOM
        bom_tab = ttk.Frame(self.tabs); self.tabs.add(bom_tab, text="Calculation & BOM")
        # -------- Horizontal Button Bar --------
        button_row = ttk.Frame(bom_tab)
        button_row.pack(anchor=tk.W, padx=10, pady=8)
        ttk.Button(button_row, text="Calculate Agent", command=self.calculate_agent).pack(side=tk.LEFT, padx=4)
        ttk.Button(button_row, text="Generate BOM", command=self.generate_bom).pack(side=tk.LEFT, padx=4)
        ttk.Button(button_row, text="Export Calculation to PDF", command=self.export_to_pdf).pack(side=tk.LEFT, padx=4)
        ttk.Button(button_row, text="Export Calculation to Word", command=self.export_to_word).pack(side=tk.LEFT, padx=4)
        ttk.Button(button_row, text="Export BOM to Excel", command=self.export_bom_to_excel).pack(side=tk.LEFT, padx=4)
        # ---------------------------------------
        self.bom_viewer = scrolledtext.ScrolledText(bom_tab, width=115, height=20, font=("Consolas", 11)); self.bom_viewer.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        # TAB 4: CALC REPORT
        result_tab = ttk.Frame(self.tabs); self.tabs.add(result_tab, text="Calculation Report")
        self.result_box = scrolledtext.ScrolledText(result_tab, width=115, height=28, font=("Consolas", 11)); self.result_box.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.status_bar = ttk.Label(self.master, text="Ready", anchor=tk.W, font=("Segoe UI", 10), background="#f2f2f2")
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        self.set_defaults_for_agent()
    def status(self, message): self.status_bar.config(text=message); self.status_bar.update_idletasks()
    def set_defaults_for_agent(self, event=None):
        agent = self.room_agent_var.get(); defaults = AGENT_DEFAULTS.get(agent)
        if defaults:
            self.design_conc_entry.delete(0, tk.END); self.design_conc_entry.insert(0, str(defaults["design_concentration"]))
            self.temperature_entry.delete(0, tk.END); self.temperature_entry.insert(0, str(defaults["temperature"]))
            self.altitude_entry.delete(0, tk.END); self.altitude_entry.insert(0, str(defaults["altitude"]))
            self.status(f"Defaults set for {agent}.")
    def clear_inputs(self):
        for e in [self.room_name_entry, self.length_entry, self.width_entry, self.height_entry,
                  self.design_conc_entry, self.altitude_entry, self.temperature_entry]:
            e.delete(0, tk.END)
        self.set_defaults_for_agent()
    def load_rooms_to_treeview(self):
        for item in self.treeview.get_children(): self.treeview.delete(item)
        for room in self.calculator.rooms:
            self.treeview.insert("", tk.END, values=(room.name, room.length, room.width, room.height, room.units, room.design_concentration, room.altitude, room.temperature, room.agent, room.actuation_type, room.oem))
    def add_room(self):
        try:
            name = self.room_name_entry.get().strip()
            if not name: raise ValueError("Room name required.")
            length = float(self.length_entry.get()); width = float(self.width_entry.get()); height = float(self.height_entry.get())
            design_conc = float(self.design_conc_entry.get()); altitude = float(self.altitude_entry.get()); temperature = float(self.temperature_entry.get())
            units = self.unit_var.get(); agent = self.room_agent_var.get(); actuation_type = self.actuation_type_var.get(); oem = self.oem_var.get()
            if agent != "FM-200 (HFC-227ea)": actuation_type = ""; oem = ""
            room = Room(name, length, width, height, design_conc, altitude, temperature, units, agent, actuation_type, oem)
            self.calculator.add_room(room); self.treeview.insert("", tk.END, values=(name, length, width, height, units, design_conc, altitude, temperature, agent, actuation_type, oem))
            self.status(f"Room '{name}' added."); self.clear_inputs()
        except ValueError as e:
            messagebox.showerror("Input Error", str(e)); self.status(f"Failed to add room: {e}")
    def remove_room(self):
        selected = self.treeview.selection()
        if not selected: messagebox.showerror("Error", "Please select a room to remove."); return
        for item in selected:
            values = self.treeview.item(item, "values"); self.calculator.remove_room(values[0]); self.treeview.delete(item)
        self.status("Room(s) removed.")
    def import_from_excel(self):
        load_workbook = install_and_import('openpyxl').load_workbook
        file_path = filedialog.askopenfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if not file_path: return
        try:
            wb = load_workbook(file_path, data_only=True); ws = wb.active
            headers = [str(cell.value).strip().lower() for cell in ws[1]]
            required_fields = ["name", "length", "width", "height", "design_concentration", "altitude", "temperature", "units", "agent"]
            missing = [f for f in required_fields if f not in headers]
            if missing: raise ValueError(f"Excel file is missing columns: {', '.join(missing)}")
            col_map = {h: i for i, h in enumerate(headers)}
            for row in ws.iter_rows(min_row=2, values_only=True):
                if all(cell is None or str(cell).strip() == "" for cell in row): continue
                data = {field: row[col_map[field]] for field in required_fields}
                data["length"] = float(data["length"]); data["width"] = float(data["width"]); data["height"] = float(data["height"])
                data["design_concentration"] = float(data["design_concentration"]); data["altitude"] = float(data["altitude"]); data["temperature"] = float(data["temperature"])
                room = Room(**data)
                try: self.calculator.add_room(room)
                except ValueError: continue
            self.load_rooms_to_treeview(); self.status("Rooms imported from Excel.")
        except Exception as e:
            messagebox.showerror("Excel Import Error", str(e)); self.status(f"Failed to import: {e}")
    def save_project(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("Project Files", "*.json")])
        if not file_path: return
        data = {
            "rooms": [room.to_dict() for room in self.calculator.rooms],
        }
        with open(file_path, "w") as file: json.dump(data, file, indent=2)
        self.status(f"Project saved as {os.path.basename(file_path)}.")
    def open_project(self):
        file_path = filedialog.askopenfilename(defaultextension=".json", filetypes=[("Project Files", "*.json")])
        if not file_path: return
        with open(file_path, "r") as file: data = json.load(file)
        self.calculator.rooms = [Room.from_dict(r) for r in data.get("rooms", [])]
        self.load_rooms_to_treeview(); self.status(f"Project loaded from {os.path.basename(file_path)}.")
    def calculate_agent(self):
        agent_groups = defaultdict(list)
        for room in self.calculator.rooms: agent_groups[room.agent].append(room)
        out = []; timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        project = self.project_name_entry.get().strip(); customer = self.customer_name_entry.get().strip()
        grand_total = 0.0
        out.append("Clean Agent Fire Suppression Calculation Report")
        out.append(f"Date: {timestamp}")
        if project: out.append(f"Project: {project}"); 
        if customer: out.append(f"Customer: {customer}")
        out.append("-" * 110)
        for agent, rooms in agent_groups.items():
            total = 0.0; out.append(f"\nAGENT: {agent}")
            out.append(f"{'Room':<14}{'Vol(m³)':>8}  {'Design%':>7}  {'Factor':>7}  {'AltC':>5}  {'TmpC':>5}  {'Req (kg)':>12}"); out.append("-" * 85)
            for room in rooms:
                required, factor, alt_corr, temp_corr, vol = room.calculate_required_agent(); total += required
                out.append(f"{room.name:<14}{vol:>8.2f}  {room.design_concentration:>7.2f}  {factor:>7.3f}  {alt_corr:>5.2f}  {temp_corr:>5.2f}  {required:>12.2f}")
            out.append("-" * 85); out.append(f"{'Total for ' + agent:<62}{total:>12.2f} kg\n"); grand_total += total
        out.append("-" * 110); out.append(f"{'Grand Total (all agents)':<62}{grand_total:>12.2f} kg\n"); out.append(""); out.append("References:")
        for title, url in REFERENCE_LINKS: out.append(f"- {title}: {url}")
        self.result_box.delete(1.0, tk.END); self.result_box.insert(tk.END, "\n".join(out)); self.latest_results = "\n".join(out)
        self.status("Calculation complete."); self.tabs.select(3)
    def generate_bom(self):
        self.bom_per_room = []
        for room in self.calculator.rooms:
            agent_kg, *_ = room.calculate_required_agent()
            if room.agent == "FM-200 (HFC-227ea)":
                actuation_type = getattr(room, "actuation_type", "Electrical"); oem = getattr(room, "oem", "Viking")
                oem_entry = OEM_BOM_DATABASES.get(oem, OEM_BOM_DATABASES["Viking"]); bom_func = oem_entry["func"]
                room_bom = bom_func(agent_kg, actuation_type)
                self.bom_per_room.append({"room": room.name, "oem": oem, "bom": room_bom, "agent": room.agent})
            else:
                self.bom_per_room.append({"room": room.name, "oem": "", "bom": [{"part_number":"-","description":f"No BOM defined for {room.agent}.","qty":"","unit":""}], "agent": room.agent})
        self.project_bom = aggregate_bom([x["bom"] for x in self.bom_per_room if x["agent"] == "FM-200 (HFC-227ea)"])
        self.display_bom_viewer(); self.status("BOM generated and displayed.")
    def display_bom_viewer(self):
        self.bom_viewer.delete(1.0, tk.END); self.bom_viewer.insert(tk.END, "========= Room-by-Room BOM =========\n")
        for entry in self.bom_per_room:
            self.bom_viewer.insert(tk.END, f"\nRoom: {entry['room']} (Agent: {entry['agent']}" + (f", OEM: {entry['oem']}" if entry['oem'] else "") + ")\n")
            self.bom_viewer.insert(tk.END, f"{'Part Number':<15} | {'Description':<50} | {'Qty':>5} | {'Unit':<4}\n")
            self.bom_viewer.insert(tk.END, "-"*90+"\n")
            for item in entry["bom"]:
                self.bom_viewer.insert(tk.END, f"{item['part_number']:<15} | {item['description']:<50} | {item['qty']:>5} | {item['unit']:<4}\n")
        self.bom_viewer.insert(tk.END, "\n========= Project BOM (FM-200 Only, Total) =========\n")
        self.bom_viewer.insert(tk.END, f"{'Part Number':<15} | {'Description':<50} | {'Qty':>5} | {'Unit':<4}\n")
        self.bom_viewer.insert(tk.END, "-"*90+"\n")
        for item in self.project_bom:
            self.bom_viewer.insert(tk.END, f"{item['part_number']:<15} | {item['description']:<50} | {item['qty']:>5} | {item['unit']:<4}\n")
        self.tabs.select(2)
    def export_bom_to_excel(self):
        openpyxl = install_and_import('openpyxl'); Workbook = openpyxl.Workbook; get_column_letter = openpyxl.utils.get_column_letter
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")]);  wb = Workbook(); ws = wb.active; ws.title = "BOM"
        ws.append(["Room", "Agent", "OEM", "Part Number", "Description", "Qty", "Unit"])
        for entry in self.bom_per_room:
            for item in entry["bom"]:
                ws.append([entry["room"], entry["agent"], entry["oem"], item["part_number"], item["description"], item["qty"], item["unit"]])
        ws.append([]); ws.append(["PROJECT TOTAL (FM-200 Only)"]); ws.append(["Part Number", "Description", "Qty", "Unit"])
        for item in self.project_bom:
            ws.append([item["part_number"], item["description"], item["qty"], item["unit"]])
        for col in range(1, 8): ws.column_dimensions[get_column_letter(col)].width = 22
        if file_path: wb.save(file_path); self.status(f"BOM exported as {os.path.basename(file_path)}.")
    def export_to_pdf(self):
        reportlab = install_and_import('reportlab'); A4 = reportlab.lib.pagesizes.A4; getSampleStyleSheet = reportlab.lib.styles.getSampleStyleSheet
        SimpleDocTemplate = reportlab.platypus.SimpleDocTemplate; Paragraph = reportlab.platypus.Paragraph; Spacer = reportlab.platypus.Spacer
        Table = reportlab.platypus.Table; TableStyle = reportlab.platypus.TableStyle; colors = reportlab.lib.colors
        file_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
        if not file_path: return
        doc = SimpleDocTemplate(file_path, pagesize=A4); styles = getSampleStyleSheet(); elements = []
        def para(txt): return Paragraph(txt, styles["Normal"])
        lines = self.latest_results.split("\n")
        for line in lines:
            if line.strip().startswith("AGENT:"): elements.append(Spacer(1, 8)); elements.append(Paragraph(f"<b>{line.strip()}</b>", styles["Heading4"]))
            elif line.strip().startswith("Clean Agent Fire Suppression"): elements.append(Paragraph(f"<b>{line.strip()}</b>", styles["Title"]))
            elif line.strip().startswith("Date:") or line.strip().startswith("Project:") or line.strip().startswith("Customer:"): elements.append(para(line.strip()))
            elif line.strip().startswith("References:"): elements.append(Spacer(1, 12)); elements.append(Paragraph("<b>References:</b>", styles["Normal"]))
            elif line.strip().startswith("- "): elements.append(para(line.strip()))
            elif line.strip() == "": elements.append(Spacer(1, 8))
            elif "Room" in line and "Vol" in line: tbl_data = []; header = [h for h in line.split() if h.strip()]; tbl_data.append(header)
            elif line.strip().startswith("-" * 5): continue
            elif "Grand Total" in line or "Total for" in line: elements.append(Spacer(1, 4)); elements.append(para(f"<b>{line}</b>"))
            elif ":" in line and not line.startswith(" "): elements.append(para(line.strip()))
            else:
                parts = [p for p in line.split() if p];  # Table rows
                if len(parts) >= 7:
                    tbl_data.append(parts)
                    if len(tbl_data) == 2 or (len(tbl_data) > 2 and tbl_data[-2][0] == "Room"):
                        table = Table(tbl_data)
                        table.setStyle(TableStyle([
                            ("BACKGROUND", (0, 0), (-1, 0), colors.grey),
                            ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
                            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                            ("FONTSIZE", (0, 0), (-1, 0), 10),
                            ("BOTTOMPADDING", (0, 0), (-1, 0), 8),
                            ("BACKGROUND", (0, 1), (-1, -1), colors.beige),
                            ("GRID", (0, 0), (-1, -1), 1, colors.black),
                        ])); elements.append(table); tbl_data = []
        doc.build(elements); self.status(f"PDF report saved as {os.path.basename(file_path)}.")
    def export_to_word(self):
        Document = install_and_import('python-docx', 'docx').Document
        file_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Files", "*.docx")])
        if not file_path: return
        doc = Document(); doc.add_heading('Clean Agent Fire Suppression Calculation Report', 0)
        lines = self.latest_results.split("\n")
        for line in lines:
            if line.strip().startswith("AGENT:"): doc.add_heading(line.strip(), level=2)
            elif line.strip().startswith("Date:") or line.strip().startswith("Project:") or line.strip().startswith("Customer:"): doc.add_paragraph(line.strip())
            elif "Room" in line and "Vol" in line: tbl_data = []; header = [h for h in line.split() if h.strip()]; tbl_data.append(header)
            elif line.strip().startswith("-" * 5): continue
            elif "Grand Total" in line or "Total for" in line: doc.add_paragraph(line.strip(), style='Intense Quote')
            elif "References:" in line: doc.add_heading("References:", level=3)
            elif line.strip().startswith("- "): doc.add_paragraph(line.strip())
            elif line.strip() == "": doc.add_paragraph("")
            else:
                parts = [p for p in line.split() if p]
                if len(parts) >= 7:
                    if not tbl_data: tbl_data = []
                    tbl_data.append(parts)
                    if len(tbl_data) == 2 or (len(tbl_data) > 2 and tbl_data[-2][0] == "Room"):
                        table = doc.add_table(rows=1, cols=len(tbl_data[0]))
                        hdr_cells = table.rows[0].cells
                        for i, h in enumerate(tbl_data[0]): hdr_cells[i].text = h
                        for row_data in tbl_data[1:]:
                            row_cells = table.add_row().cells
                            for i, val in enumerate(row_data): row_cells[i].text = val
                        tbl_data = []
                else: doc.add_paragraph(line.strip())
        doc.save(file_path); self.status(f"Word report saved as {os.path.basename(file_path)}.")

if __name__ == "__main__":
    root = tk.Tk(); app = ModernFM200App(root); root.mainloop()

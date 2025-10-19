import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import csv

# openpyxl is optional: try to import, otherwise we'll show a message if user attempts Excel export
try:
    from openpyxl import Workbook
    _HAS_OPENPYXL = True
except ImportError:
    _HAS_OPENPYXL = False

# -------------------- Calculation Functions -------------------- #
def calculate_values():
    try:
        width = float(width_entry.get())
        length = float(length_entry.get())
        grille_type = grille_type_var.get()
        airflow = airflow_entry.get().strip()
        avg_velocity = avg_velocity_entry.get().strip()

        free_area = grille_types[grille_type]

        effective_area = width * length * free_area

        # Determine which value to calculate
        if airflow and not avg_velocity:
            airflow = float(airflow)
            avg_velocity_calc = airflow / (effective_area * 3600)
            avg_velocity_entry.delete(0, tk.END)
            avg_velocity_entry.insert(0, f"{avg_velocity_calc:.2f}")
        elif avg_velocity and not airflow:
            avg_velocity = float(avg_velocity)
            airflow_calc = effective_area * avg_velocity * 3600
            airflow_entry.delete(0, tk.END)
            airflow_entry.insert(0, f"{airflow_calc:.2f}")
        else:
            messagebox.showwarning("Input Error", "Please fill either Air Flow OR Average Velocity, not both.")
            return

        messagebox.showinfo("Calculation Complete", f"Effective Area = {effective_area:.4f} m²")

    except ValueError:
        messagebox.showerror("Invalid Input", "Please ensure all numerical inputs are valid.")

def add_to_table():
    try:
        width = float(width_entry.get())
        length = float(length_entry.get())
        grille_type = grille_type_var.get()
        airflow = airflow_entry.get()
        avg_velocity = avg_velocity_entry.get()

        free_area = grille_types[grille_type]
        effective_area = width * length * free_area

        table.insert("", "end", values=(
            width, length, grille_type, round(free_area, 2),
            round(effective_area, 4), airflow, avg_velocity
        ))

    except ValueError:
        messagebox.showerror("Error", "Please check your inputs.")

def delete_selected():
    selected = table.selection()
    if not selected:
        messagebox.showwarning("Warning", "Please select a row to delete.")
        return
    for sel in selected:
        table.delete(sel)

def clear_table():
    for row in table.get_children():
        table.delete(row)

def export_to_excel():
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                             filetypes=[("Excel Files", "*.xlsx")])
    if not file_path:
        return
    if not _HAS_OPENPYXL:
        messagebox.showerror("Missing Dependency", "openpyxl is not installed. Install it to export Excel files.")
        return

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Velocity Data"

    headers = ["Width (m)", "Length (m)", "Type", "Free Area", "Effective Area (m²)", "Air Flow (CMH)", "Avg Velocity (m/s)"]
    sheet.append(headers)

    for row in table.get_children():
        sheet.append(table.item(row)['values'])

    workbook.save(file_path)
    messagebox.showinfo("Export Complete", f"Data exported to {file_path}")

# -------------------- GUI Setup -------------------- #
root = tk.Tk()
root.title("Average Velocity Calculator by RSHAN Build 1.0.0-2025")
root.geometry("900x600")

# Dropdown options
grille_types = {
    "Double Deflection": 0.70,
    "Eggcrate": 0.88,
    "4-Way Sag": 0.50
}

# Inputs
frame_inputs = ttk.LabelFrame(root, text="Input Parameters")
frame_inputs.pack(fill="x", padx=10, pady=10)

ttk.Label(frame_inputs, text="Width (m):").grid(row=0, column=0, padx=5, pady=5)
width_entry = ttk.Entry(frame_inputs)
width_entry.grid(row=0, column=1, padx=5, pady=5)

ttk.Label(frame_inputs, text="Length (m):").grid(row=0, column=2, padx=5, pady=5)
length_entry = ttk.Entry(frame_inputs)
length_entry.grid(row=0, column=3, padx=5, pady=5)

ttk.Label(frame_inputs, text="Grille Type:").grid(row=0, column=4, padx=5, pady=5)
grille_type_var = tk.StringVar(value="Double Deflection")
grille_type_menu = ttk.Combobox(frame_inputs, textvariable=grille_type_var, values=list(grille_types.keys()), state="readonly")
grille_type_menu.grid(row=0, column=5, padx=5, pady=5)

ttk.Label(frame_inputs, text="Air Flow (CMH):").grid(row=1, column=0, padx=5, pady=5)
airflow_entry = ttk.Entry(frame_inputs)
airflow_entry.grid(row=1, column=1, padx=5, pady=5)

ttk.Label(frame_inputs, text="Avg Velocity (m/s):").grid(row=1, column=2, padx=5, pady=5)
avg_velocity_entry = ttk.Entry(frame_inputs)
avg_velocity_entry.grid(row=1, column=3, padx=5, pady=5)

ttk.Button(frame_inputs, text="Calculate", command=calculate_values).grid(row=1, column=4, padx=5, pady=5)
ttk.Button(frame_inputs, text="Add to List", command=add_to_table).grid(row=1, column=5, padx=5, pady=5)

# Table
frame_table = ttk.LabelFrame(root, text="System List")
frame_table.pack(fill="both", expand=True, padx=10, pady=10)

columns = ("Width", "Length", "Type", "FreeArea", "EffectiveArea", "AirFlow", "AvgVelocity")
table = ttk.Treeview(frame_table, columns=columns, show="headings")

for col in columns:
    table.heading(col, text=col)
    table.column(col, anchor="center", width=120)

table.pack(fill="both", expand=True)

# Buttons
frame_buttons = ttk.Frame(root)
frame_buttons.pack(fill="x", padx=10, pady=10)

ttk.Button(frame_buttons, text="Delete Selected", command=delete_selected).pack(side="left", padx=5)
ttk.Button(frame_buttons, text="Clear Table", command=clear_table).pack(side="left", padx=5)
ttk.Button(frame_buttons, text="Export to Excel", command=export_to_excel).pack(side="right", padx=5)

root.mainloop()


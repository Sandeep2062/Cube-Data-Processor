import openpyxl
from openpyxl.drawing.image import Image as XLImage
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import os
import winsound
from copy import deepcopy
import io

"""
╔══════════════════════════════════════════════════════════════════╗
║                    CUBE DATA PROCESSOR v3.0                       ║
║                                                                   ║
║  Developer: Sandeep (https://github.com/Sandeep2062)            ║
║  Repository: https://github.com/Sandeep2062/Cube-Merge          ║
║  Description: Excel data processor with logo preservation        ║
║                                                                   ║
║  © 2025 Sandeep - All Rights Reserved                           ║
╚══════════════════════════════════════════════════════════════════╝
"""

# Detect grade from filename
def extract_grade(filename):
    name = os.path.basename(filename).split('.')[0].upper()
    name = name.replace("_", ":").replace("-", ":")
    return name.strip()


# FILLED ROW CHECKER
def get_last_row(ws):
    row = 2
    while True:
        if ws.cell(row=row, column=2).value in (None, ""):
            return row - 1
        row += 1


# PRESERVE ALL IMAGES/LOGOS FROM SOURCE WORKSHEET
def preserve_images(source_ws, target_ws):
    """Copy all images (logos) from source to target worksheet"""
    try:
        if hasattr(source_ws, '_images') and source_ws._images:
            for img in source_ws._images:
                # Create a new image from the existing one
                new_img = XLImage(img.ref)
                # Copy all anchor properties
                new_img.anchor = deepcopy(img.anchor)
                # Add to target worksheet
                target_ws.add_image(new_img)
    except Exception as e:
        print(f"Warning: Could not copy images - {e}")


# MAIN PROCESSING - SEPARATE MODE
def process_grade_separate(grade_file, office_file, output_folder, log):
    try:
        grade_wb = openpyxl.load_workbook(grade_file)
        grade_ws = grade_wb.active

        grade_name = extract_grade(grade_file)
        log(f"\n=== Processing {grade_file}")
        log(f"Detected Grade: {grade_name}")

        # Load office file - keep images
        office_wb = openpyxl.load_workbook(office_file, keep_vba=True)

        last_row = get_last_row(grade_ws)
        log(f"Total data rows: {last_row - 1}")

        # Get all sheets that match this grade
        matching_sheets = []
        for sheet_name in office_wb.sheetnames:
            ws = office_wb[sheet_name]
            b12 = str(ws["B12"].value).replace(" ", "").upper()
            if b12 == grade_name:
                matching_sheets.append(sheet_name)
        
        log(f"Found {len(matching_sheets)} sheets matching grade '{grade_name}'")
        
        if len(matching_sheets) == 0:
            log(f"⚠ WARNING: No sheets found with '{grade_name}' in cell B12!")
            return 0

        copy_count = 0
        sheet_index = 0

        # Loop through each data row
        for r in range(2, last_row + 1):
            if sheet_index >= len(matching_sheets):
                log(f"⚠ Warning: More data rows than matching sheets. Stopping at row {r}")
                break

            current_sheet_name = matching_sheets[sheet_index]
            ws = office_wb[current_sheet_name]
            
            # Read weight and strength values
            weight_values = [grade_ws.cell(row=r, column=c).value for c in range(2, 8)]
            strength_values = [grade_ws.cell(row=r, column=c).value for c in range(9, 15)]

            # Write values (ONLY modify data cells, logos remain untouched)
            for i, v in enumerate(weight_values):
                ws.cell(row=25, column=3 + i, value=v)
            for i, v in enumerate(strength_values):
                ws.cell(row=27, column=3 + i, value=v)

            copy_count += 1
            log(f"✓ Row {r} → Sheet: {current_sheet_name}")
            sheet_index += 1

        # Save with logos preserved
        base = os.path.basename(office_file).split(".")[0]
        outname = f"{base}_{grade_name}_Processed.xlsx"
        outpath = os.path.join(output_folder, outname)

        office_wb.save(outpath)
        log(f"✓ Saved → {outpath} (Logos preserved)")

        return copy_count

    except Exception as e:
        log(f"✖ ERROR: {e}")
        return 0


# MAIN PROCESSING - COMBINE MODE
def process_all_grades_combined(grade_files, office_file, output_folder, log):
    try:
        log(f"\n=== COMBINE MODE: Processing all grades into one file ===")
        
        # Load office file ONCE - keep images
        office_wb = openpyxl.load_workbook(office_file, keep_vba=True)
        
        total_copy_count = 0
        
        # Process each grade file
        for grade_file in grade_files:
            grade_wb = openpyxl.load_workbook(grade_file)
            grade_ws = grade_wb.active
            grade_name = extract_grade(grade_file)
            
            log(f"\n--- Processing Grade: {grade_name} ---")
            
            last_row = get_last_row(grade_ws)
            log(f"Data rows: {last_row - 1}")
            
            # Find matching sheets for this grade
            matching_sheets = []
            for sheet_name in office_wb.sheetnames:
                ws = office_wb[sheet_name]
                b12 = str(ws["B12"].value).replace(" ", "").upper()
                if b12 == grade_name:
                    matching_sheets.append(sheet_name)
            
            log(f"Found {len(matching_sheets)} sheets for '{grade_name}'")
            
            if len(matching_sheets) == 0:
                log(f"⚠ No sheets found for '{grade_name}'")
                continue
            
            sheet_index = 0
            
            # Copy data for this grade (logos stay intact)
            for r in range(2, last_row + 1):
                if sheet_index >= len(matching_sheets):
                    log(f"⚠ More rows than sheets for {grade_name}")
                    break
                
                current_sheet_name = matching_sheets[sheet_index]
                ws = office_wb[current_sheet_name]
                
                weight_values = [grade_ws.cell(row=r, column=c).value for c in range(2, 8)]
                strength_values = [grade_ws.cell(row=r, column=c).value for c in range(9, 15)]
                
                for i, v in enumerate(weight_values):
                    ws.cell(row=25, column=3 + i, value=v)
                for i, v in enumerate(strength_values):
                    ws.cell(row=27, column=3 + i, value=v)
                
                total_copy_count += 1
                log(f"✓ {grade_name} Row {r} → {current_sheet_name}")
                sheet_index += 1
        
        # Save ONE combined file with all logos
        base = os.path.basename(office_file).split(".")[0]
        outname = f"{base}_ALL_GRADES_Combined.xlsx"
        outpath = os.path.join(output_folder, outname)
        
        office_wb.save(outpath)
        log(f"\n✓✓✓ COMBINED FILE SAVED → {outpath} (Logos preserved)")
        
        return total_copy_count
        
    except Exception as e:
        log(f"✖ ERROR: {e}")
        return 0


# ------------- GUI LOGIC -------------

def run_processing():
    if not grade_files:
        messagebox.showerror("Error", "Please select grade files.")
        return

    if not office_path.get():
        messagebox.showerror("Error", "Select office format file.")
        return

    if not output_path.get():
        messagebox.showerror("Error", "Select output folder.")
        return

    log_box.delete("1.0", "end")
    total = 0

    if mode_var.get() == 2:  # COMBINE MODE
        progress["value"] = 50
        root.update_idletasks()
        
        total = process_all_grades_combined(
            grade_files,
            office_path.get(),
            output_path.get(),
            log=lambda m: log_box.insert(tk.END, m + "\n")
        )
        
        progress["value"] = 100
        
    else:  # SEPARATE MODE
        for i, file in enumerate(grade_files):
            progress["value"] = (i + 1) / len(grade_files) * 100
            root.update_idletasks()

            total += process_grade_separate(
                file,
                office_path.get(),
                output_path.get(),
                log=lambda m: log_box.insert(tk.END, m + "\n")
            )

    winsound.MessageBeep()
    messagebox.showinfo("✓ Completed", f"Processing Complete!\n\nTotal Rows Copied: {total}")


def add_grades():
    files = filedialog.askopenfilenames(filetypes=[("Excel Files", "*.xlsx")])
    for f in files:
        if f not in grade_files:
            grade_files.append(f)
            grade_listbox.insert(tk.END, os.path.basename(f))


def clear_grades():
    grade_files.clear()
    grade_listbox.delete(0, tk.END)


def pick_office():
    path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if path:
        office_path.set(path)


def pick_output_folder():
    folder = filedialog.askdirectory()
    if folder:
        output_path.set(folder)


# ------------------- ENHANCED MODERN GUI -------------------

root = tk.Tk()
root.title("Cube Data Processor v3.0")
root.geometry("900x800")
root.configure(bg="#f5f5f5")

# Try to set window icon (if icon.ico exists)
try:
    root.iconbitmap("icon.ico")
except:
    pass

grade_files = []
office_path = tk.StringVar()
output_path = tk.StringVar()
mode_var = tk.IntVar(value=1)

# Enhanced Style
style = ttk.Style()
style.theme_use('clam')
style.configure('TButton', padding=8, relief="flat", background="#0078d4", 
                foreground="white", font=("Segoe UI", 9))
style.map('TButton', background=[('active', '#005a9e')])
style.configure('Action.TButton', padding=10, background="#28a745", font=("Segoe UI", 10, "bold"))
style.map('Action.TButton', background=[('active', '#218838')])

# GRADIENT HEADER with Logo Space
header_frame = tk.Frame(root, bg="#0066cc", height=90)
header_frame.pack(fill=tk.X)
header_frame.pack_propagate(False)

# Logo placeholder (LEFT SIDE) - Add your logo here
logo_container = tk.Frame(header_frame, bg="#0066cc")
logo_container.place(x=20, y=15)

# Try to load logo if exists
try:
    from PIL import Image, ImageTk
    logo_img = Image.open("logo.png")  # Your logo here
    logo_img = logo_img.resize((60, 60), Image.Resampling.LANCZOS)
    logo_photo = ImageTk.PhotoImage(logo_img)
    logo_label = tk.Label(logo_container, image=logo_photo, bg="#0066cc")
    logo_label.image = logo_photo
    logo_label.pack()
except:
    # Fallback: Text logo
    logo_label = tk.Label(logo_container, text="🔷", font=("Segoe UI", 40), bg="#0066cc", fg="white")
    logo_label.pack()

# Title (CENTER)
title_container = tk.Frame(header_frame, bg="#0066cc")
title_container.pack(expand=True)

title_label = tk.Label(title_container, text="Cube Data Processor", 
                       font=("Segoe UI", 22, "bold"), bg="#0066cc", fg="white")
title_label.pack()

version_label = tk.Label(title_container, text="v3.0 - Professional Edition", 
                        font=("Segoe UI", 9), bg="#0066cc", fg="#e0e0e0")
version_label.pack()

# Developer Credit (RIGHT SIDE)
credit_frame = tk.Frame(header_frame, bg="#0066cc")
credit_frame.place(relx=1.0, y=15, anchor="ne", x=-20)

credit_label = tk.Label(credit_frame, text="Developed by Sandeep", 
                       font=("Segoe UI", 8), bg="#0066cc", fg="white")
credit_label.pack()

github_label = tk.Label(credit_frame, text="github.com/Sandeep2062", 
                       font=("Segoe UI", 7), bg="#0066cc", fg="#b3d9ff", cursor="hand2")
github_label.pack()
github_label.bind("<Button-1>", lambda e: os.system("start https://github.com/Sandeep2062/Cube-Merge"))

# Main container with shadow effect
main_container = tk.Frame(root, bg="#f5f5f5")
main_container.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)

# Grade Files Section
grade_frame = tk.LabelFrame(main_container, text="  📁 Grade Files  ", 
                            font=("Segoe UI", 11, "bold"), bg="white", 
                            fg="#333", relief=tk.FLAT, bd=2, padx=15, pady=15)
grade_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

btn_frame = tk.Frame(grade_frame, bg="white")
btn_frame.pack(fill=tk.X, pady=(0, 10))

ttk.Button(btn_frame, text="➕ Add Files", command=add_grades).pack(side=tk.LEFT, padx=5)
ttk.Button(btn_frame, text="🗑️ Clear All", command=clear_grades).pack(side=tk.LEFT, padx=5)

grade_listbox = tk.Listbox(grade_frame, height=5, font=("Consolas", 9), 
                           bg="#fafafa", relief=tk.FLAT, bd=0, 
                           highlightthickness=1, highlightbackground="#ddd")
grade_listbox.pack(fill=tk.BOTH, expand=True)

# Office File Section
office_frame = tk.LabelFrame(main_container, text="  📄 Office Format File  ", 
                            font=("Segoe UI", 11, "bold"), bg="white", 
                            fg="#333", relief=tk.FLAT, bd=2, padx=15, pady=15)
office_frame.pack(fill=tk.X, pady=(0, 10))

ttk.Button(office_frame, text="📂 Select File", command=pick_office).pack(anchor=tk.W, pady=(0, 8))
office_entry = tk.Entry(office_frame, textvariable=office_path, font=("Segoe UI", 9),
                       bg="#fafafa", relief=tk.FLAT, bd=0)
office_entry.pack(fill=tk.X, ipady=8, padx=2)

# Output Folder Section
output_frame = tk.LabelFrame(main_container, text="  💾 Output Folder  ", 
                            font=("Segoe UI", 11, "bold"), bg="white", 
                            fg="#333", relief=tk.FLAT, bd=2, padx=15, pady=15)
output_frame.pack(fill=tk.X, pady=(0, 10))

ttk.Button(output_frame, text="📂 Select Folder", command=pick_output_folder).pack(anchor=tk.W, pady=(0, 8))
output_entry = tk.Entry(output_frame, textvariable=output_path, font=("Segoe UI", 9),
                       bg="#fafafa", relief=tk.FLAT, bd=0)
output_entry.pack(fill=tk.X, ipady=8, padx=2)

# Processing Mode
mode_frame = tk.LabelFrame(main_container, text="  ⚙️ Processing Mode  ", 
                          font=("Segoe UI", 11, "bold"), bg="white", 
                          fg="#333", relief=tk.FLAT, bd=2, padx=15, pady=15)
mode_frame.pack(fill=tk.X, pady=(0, 10))

tk.Radiobutton(mode_frame, text="📑 Separate Files (One file per grade)", 
               variable=mode_var, value=1, font=("Segoe UI", 10),
               bg="white", activebackground="white", fg="#333").pack(anchor=tk.W, pady=3)
tk.Radiobutton(mode_frame, text="📦 Combined File (All grades merged into one)", 
               variable=mode_var, value=2, font=("Segoe UI", 10),
               bg="white", activebackground="white", fg="#333").pack(anchor=tk.W, pady=3)

# Start Button - More prominent
start_btn = tk.Button(main_container, text="▶️  START PROCESSING", command=run_processing,
                     font=("Segoe UI", 12, "bold"), bg="#28a745", fg="white",
                     activebackground="#218838", relief=tk.FLAT, cursor="hand2",
                     padx=30, pady=15, borderwidth=0)
start_btn.pack(pady=15)

# Progress Bar - Enhanced
progress_frame = tk.Frame(main_container, bg="#f5f5f5")
progress_frame.pack(fill=tk.X, pady=(0, 10))

progress = ttk.Progressbar(progress_frame, length=500, mode="determinate")
progress.pack()

# Log Section - Enhanced
log_frame = tk.LabelFrame(main_container, text="  📋 Processing Log  ", 
                         font=("Segoe UI", 11, "bold"), bg="white", 
                         fg="#333", relief=tk.FLAT, bd=2, padx=15, pady=15)
log_frame.pack(fill=tk.BOTH, expand=True)

log_scrollbar = tk.Scrollbar(log_frame)
log_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

log_box = tk.Text(log_frame, height=10, font=("Consolas", 8), bg="#f8f8f8",
                 relief=tk.FLAT, bd=0, wrap=tk.WORD, yscrollcommand=log_scrollbar.set)
log_box.pack(fill=tk.BOTH, expand=True)
log_scrollbar.config(command=log_box.yview)

# Footer
footer = tk.Label(root, text="© 2025 Sandeep | github.com/Sandeep2062/Cube-Merge | Preserves Logos ✓", 
                 font=("Segoe UI", 8), bg="#f5f5f5", fg="#666")
footer.pack(pady=5)

root.mainloop()
import os
import shutil
import threading
import re
import stat
import subprocess
import sys
import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, messagebox

# ─── Permissions & File Access Management ────────────────────────────────────

def remove_readonly(path):
    """
    Remove read-only attribute from a file or folder so any app can modify or overwrite it.
    Works on Windows (clearing win32 FILE_ATTRIBUTE_READONLY) and POSIX.
    """
    path_str = str(path)
    try:
        if os.path.isdir(path_str):
            os.chmod(path_str, stat.S_IRWXU | stat.S_IRWXG | stat.S_IRWXO)
        else:
            os.chmod(path_str, stat.S_IREAD | stat.S_IWRITE | stat.S_IRGRP | stat.S_IWGRP | stat.S_IROTH | stat.S_IWOTH)
    except Exception:
        pass

    if sys.platform.startswith("win"):
        try:
            import ctypes
            attrs = ctypes.windll.kernel32.GetFileAttributesW(path_str)
            FILE_ATTRIBUTE_READONLY = 0x01
            FILE_ATTRIBUTE_NORMAL = 0x80
            if attrs != -1 and (attrs & FILE_ATTRIBUTE_READONLY):
                new_attrs = attrs & ~FILE_ATTRIBUTE_READONLY
                if new_attrs == 0:
                    new_attrs = FILE_ATTRIBUTE_NORMAL
                ctypes.windll.kernel32.SetFileAttributesW(path_str, new_attrs)
        except Exception:
            pass


def reset_to_default_permissions(path, is_dir=None, recursive=False):
    """
    Ensure the path has default Windows inherited permissions and is not read-only.
    Does NOT add any special or custom ACLs (keeps pure default Windows inheritance).
    """
    path_str = str(path)
    if not os.path.exists(path_str):
        return

    remove_readonly(path_str)

    if sys.platform.startswith("win"):
        try:
            icacls_bin = os.path.join(os.environ.get("SystemRoot", r"C:\Windows"), "System32", "icacls.exe")
            if not os.path.isfile(icacls_bin):
                icacls_bin = "icacls"

            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = 0  # SW_HIDE
            creationflags = getattr(subprocess, 'CREATE_NO_WINDOW', 0x08000000)

            args = [icacls_bin, path_str, "/reset"]
            if recursive:
                args.extend(["/T", "/C", "/Q"])
            else:
                args.extend(["/C", "/Q"])

            subprocess.run(
                args,
                capture_output=True,
                startupinfo=startupinfo,
                creationflags=creationflags,
                timeout=30 if recursive else 5
            )
        except Exception:
            pass

grant_full_permissions = reset_to_default_permissions


def prepare_destination(dst):
    """
    Ensure destination file or folder path is writeable before writing, copying, or moving to it.
    If it exists, removes read-only attribute.
    """
    try:
        path_str = str(dst)
        if os.path.exists(path_str):
            remove_readonly(path_str)
    except Exception:
        pass


def make_dir_with_permissions(dir_path):
    """
    Create directory and ensure default Windows inheritance and no read-only attribute.
    """
    path_str = str(dir_path)
    os.makedirs(path_str, exist_ok=True)
    remove_readonly(path_str)
    return path_str


def ensure_full_permissions_tree(root_dir):
    """
    Recursively remove read-only attribute and ensure default Windows inheritance
    without any special ACLs.
    """
    path_str = str(root_dir)
    if not os.path.exists(path_str):
        return

    try:
        remove_readonly(path_str)
        for root, dirs, files in os.walk(path_str):
            for d in dirs:
                remove_readonly(os.path.join(root, d))
            for f in files:
                remove_readonly(os.path.join(root, f))
    except Exception:
        pass

    reset_to_default_permissions(path_str, is_dir=True, recursive=True)


def safe_copy(src, dst):
    """
    Copy file ensuring destination is writeable, binary identical, and inherits default Windows permissions.
    """
    src_str = str(src)
    dst_str = str(dst)
    prepare_destination(dst_str)
    shutil.copyfile(src_str, dst_str)
    remove_readonly(dst_str)


def safe_move(src, dst):
    """
    Move file ensuring destination is writeable, falling back to copy+delete if locked or across drives,
    and ensuring destination has full read/write permissions for all apps.
    """
    src_str = str(src)
    dst_str = str(dst)
    prepare_destination(dst_str)
    remove_readonly(src_str)
    try:
        shutil.move(src_str, dst_str)
    except Exception:
        safe_copy(src_str, dst_str)
        try:
            remove_readonly(src_str)
            os.remove(src_str)
        except Exception:
            pass

    remove_readonly(dst_str)


def safe_df_to_excel(df, filepath, index=False, log_callback=None):
    """
    Save pandas DataFrame to Excel, handling file-in-use locks gracefully.
    """
    path_str = str(filepath)
    prepare_destination(path_str)
    try:
        df.to_excel(path_str, index=index)
        remove_readonly(path_str)
        return path_str
    except PermissionError:
        import time
        base, ext = os.path.splitext(path_str)
        timestamp = time.strftime("_%Y%m%d_%H%M%S")
        alt_path = f"{base}{timestamp}{ext}"
        if log_callback:
            log_callback(f"⚠ Fișierul {os.path.basename(path_str)} este deschis în altă aplicație (Excel). Salvat ca: {os.path.basename(alt_path)}")
        prepare_destination(alt_path)
        df.to_excel(alt_path, index=index)
        remove_readonly(alt_path)
        return alt_path

class NCSorterApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("NC/DSTV File Sorter")
        self.geometry("750x650")
        self.minsize(600, 500)

        # Variables
        self.source_folder = ctk.StringVar()
        self.action_var = ctk.StringVar(value="Copy")
        self.opt_hot_cold = ctk.BooleanVar(value=False)
        self.opt_separate_tubes = ctk.BooleanVar(value=False)
        self.opt_verify_geometry = ctk.BooleanVar(value=True)

        self.create_widgets()

    def create_widgets(self):
        # Header
        header_label = ctk.CTkLabel(self, text="NC/DSTV File Sorter", font=ctk.CTkFont(size=24, weight="bold"))
        header_label.pack(pady=20)

        # Source Folder Selection
        src_frame = ctk.CTkFrame(self)
        src_frame.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkLabel(src_frame, text="Source Folder:").pack(side="left", padx=10)
        ctk.CTkEntry(src_frame, textvariable=self.source_folder, width=400).pack(side="left", padx=10, fill="x", expand=True)
        ctk.CTkButton(src_frame, text="Browse", command=self.browse_source).pack(side="left", padx=10)

        # Action Toggle
        action_frame = ctk.CTkFrame(self)
        action_frame.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkLabel(action_frame, text="Action:").pack(side="left", padx=10)
        ctk.CTkRadioButton(action_frame, text="Copy files", variable=self.action_var, value="Copy").pack(side="left", padx=10)
        ctk.CTkRadioButton(action_frame, text="Move files", variable=self.action_var, value="Move").pack(side="left", padx=10)

        # Options Toggle
        opt_frame = ctk.CTkFrame(self)
        opt_frame.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkLabel(opt_frame, text="Options:").pack(side="left", padx=10)
        ctk.CTkCheckBox(opt_frame, text="Sortare profile Laminat la Cald / Rece", variable=self.opt_hot_cold).pack(side="left", padx=10)
        ctk.CTkCheckBox(opt_frame, text="Separare Profile de Țevi", variable=self.opt_separate_tubes).pack(side="left", padx=10)
        ctk.CTkCheckBox(opt_frame, text="Validare Geometrică (Erori 3D)", variable=self.opt_verify_geometry).pack(side="left", padx=10)

        # Start Button
        self.start_btn = ctk.CTkButton(self, text="Start Sorting", command=self.start_sorting, font=ctk.CTkFont(size=16, weight="bold"), height=40)
        self.start_btn.pack(pady=20)

        # Progress Bar
        self.progress = ctk.CTkProgressBar(self)
        self.progress.pack(fill="x", padx=20, pady=10)
        self.progress.set(0)

        # Terminal / Log Output
        self.log_box = ctk.CTkTextbox(self, state="disabled")
        self.log_box.pack(fill="both", padx=20, pady=10, expand=True)

    def browse_source(self):
        folder = filedialog.askdirectory(title="Select Source Folder")
        if folder:
            self.source_folder.set(folder)

    def log(self, message):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", message + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def start_sorting(self):
        src = self.source_folder.get()

        if not src:
            messagebox.showerror("Error", "Please select a Source folder.")
            return

        if not os.path.exists(src):
            messagebox.showerror("Error", "Source folder does not exist.")
            return

        self.start_btn.configure(state="disabled")
        self.progress.set(0)
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")

        # Start sorting in a background thread to keep UI responsive
        threading.Thread(target=self.process_files, args=(src, self.action_var.get(), self.opt_hot_cold.get(), self.opt_separate_tubes.get(), self.opt_verify_geometry.get()), daemon=True).start()

    def determine_category(self, profile, material, opt_hot_cold=False, opt_separate_tubes=False):
        prof_upper = profile.strip().upper()
        mat_upper = material.strip().upper()

        inox_materials = ["INOX", "1.4301", "304", "316", "1.4404", "1.4307", "1.4571", "1.4401", "1.4016", "A2", "A4"]
        alu_materials = ["ALUMINIU", "AW", "ALU", "5754", "6082", "6060", "7075"]

        # Step 1: T-URI (Everything starting with T: T, TB, TABLA, T=...)
        if prof_upper.startswith("T"):
            return "T-URI", "T-URI"

        # Step 2: Check for Bars
        is_bar = False
        bar_prefixes = ["RO", "GST", "VKT", "SQ", "PT", "RD", "FI", "ARM"]
        for p in bar_prefixes:
            if prof_upper.startswith(p):
                is_bar = True
                break
        
        if not is_bar and (prof_upper.startswith("D") or prof_upper.startswith("M")):
            # Ensure it's D or M followed by a number/space to avoid catching drawing names
            rest = prof_upper[1:].strip()
            if rest and (rest[0].isdigit() or rest[0] in ['_', '-']):
                is_bar = True
            elif not rest: # just exactly "D" or "M"
                is_bar = True
                
        if is_bar:
            if any(x in mat_upper for x in inox_materials):
                return os.path.join("BARE", "INOX"), "BARA INOX"
            elif any(x in mat_upper for x in alu_materials):
                return os.path.join("BARE", "ALUMINIU"), "BARA ALUMINIU"
            else:
                return "BARE", "BARA"
        
        # Step 3: Check for Box Girders
        if "HSQ" in prof_upper or "CHESON" in prof_upper:
            return "CHESON", "CHESON"
        
        # Step 4: Check for Plates
        is_plate = False
        plate_prefixes = ["PL", "BL", "FB", "FL", "SH", "SHEET"]
        for p in plate_prefixes:
            if prof_upper.startswith(p):
                is_plate = True
                break
                
        if not is_plate and prof_upper.startswith("B") and len(prof_upper) > 1 and prof_upper[1].isdigit():
            is_plate = True

        if not is_plate and len(prof_upper) > 0 and prof_upper[0].isdigit():
            is_plate = True

        # Special check to not miss plates like PL-304
        if prof_upper.startswith("PL") and len(prof_upper) > 2 and (prof_upper[2].isdigit() or prof_upper[2] in ['-', '_']):
            is_plate = True

        if is_plate:
            if "Z35" in mat_upper:
                return os.path.join("PLACI", "Z35"), "PLACA"
            elif "S355" in mat_upper:
                return os.path.join("PLACI", "S355"), "PLACA"
            elif "S235" in mat_upper:
                return os.path.join("PLACI", "S235"), "PLACA"
            elif any(x in mat_upper for x in inox_materials):
                return os.path.join("PLACI", "INOX"), "PLACA INOX"
            elif any(x in mat_upper for x in alu_materials):
                return os.path.join("PLACI", "ALUMINIU"), "PLACA ALUMINIU"
            else:
                return os.path.join("PLACI", "Other"), "PLACA"

        # Step 5: Profiles (Catch-all)
        # All profiles go directly to PROFILE (or Laminat Cald/Rece), ignoring material as requested.
        tube_hot_prefixes = ["HF", "VK"]
        tube_cold_prefixes = ["KKR", "RHS", "SHS", "CHS", "CFRHS", "KRR", "ROHR", "TUBE", "PIPE", "O"]

        is_tube = False
        is_tube_hot = False

        for p in tube_hot_prefixes:
            if prof_upper.startswith(p):
                is_tube = True
                is_tube_hot = True
                break
                
        if not is_tube:
            for p in tube_cold_prefixes:
                if prof_upper.startswith(p):
                    is_tube = True
                    break

        if is_tube:
            base_folder = "TEVI" if opt_separate_tubes else "PROFILE"
            if opt_hot_cold:
                if is_tube_hot:
                    return os.path.join(base_folder, "LAMINAT LA CALD"), "TEAVA CALD"
                else:
                    return os.path.join(base_folder, "LAMINAT LA RECE"), "TEAVA RECE"
            else:
                return base_folder, "TEAVA"

        # All other standard profiles (HEA, IPE, L, UPN, Z, C, etc.) go here
        if opt_hot_cold:
            return os.path.join("PROFILE", "LAMINAT LA RECE"), "PROFIL LAMINAT LA RECE"
        else:
            return "PROFILE", "PROFIL"

    def parse_nc_file(self, filepath):
        profile = ""
        material = ""
        
        try:
            with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
                lines = [line.strip() for line in f.readlines()[:15]]
                
                # Heuristic: Find material line by looking for common steel grades
                mat_idx = -1
                safe_grades = ["S355", "S235", "Z35", "S275", "S460", "STEEL", "B500", "K500", "INOX", "1.4301", "1.4404", "ALUMINIU", "ALU", "EN AW", "AW-", "S420", "S690", "CORTEN", "PC52", "OB37", "1.4307", "1.4571", "1.4401", "1.4016", "5754", "6082", "6060", "7075", "C45", "42CRMO4"]
                exact_grades = ["304", "316", "304L", "316L", "A2", "A4", "8.8", "10.9"]
                
                for i, line in enumerate(lines):
                    upper_line = line.upper()
                    if i < 4: continue # Skip first 4 lines (ST, ** filename, Job No, Part No)
                    if "**" in upper_line or ".NC" in upper_line: continue
                    
                    if any(grade in upper_line for grade in safe_grades) or upper_line in exact_grades:
                        mat_idx = i
                        break
                
                profile = ""
                if mat_idx != -1:
                    material = lines[mat_idx]
                    
                    # Gather candidates for Profile. 
                    candidates = []
                    if mat_idx + 2 < len(lines):
                        candidates.append(lines[mat_idx + 2])
                    if mat_idx > 0:
                        candidates.append(lines[mat_idx - 1])
                        
                    exclude_marks = []
                    if len(lines) > 3: exclude_marks.append(lines[3].upper())
                    if len(lines) > 1: exclude_marks.append(lines[1].upper().replace("**", "").strip())
                    if len(lines) > 2: exclude_marks.append(lines[2].upper())
                    
                    # Ensure we don't accidentally exclude purely numeric values (like a plate thickness of "15")
                    exclude_marks = [m for m in exclude_marks if not m.isdigit()]
                    
                    prefixes = ["PL", "B", "BL", "RO", "D", "O", "HEA", "HEB", "IPE", "UPN", "RHS", "CHS", "SHS", "CFRHS", "HFRHS", "HSQ", "CHESON", "L", "VKR", "GST", "ARM", "M", "UNP", "UPE", "INP", "FL", "HEM", "HL", "HD", "UB", "UC", "W", "UAP", "PFC", "C", "EA", "UA", "KKR", "KRR", "ROHR", "VKT", "SQ", "PT", "RD", "FI", "T", "TB", "Z", "FB", "FLT", "TABLA", "SHEET", "SH"]
                    # Sort by length descending to match longest prefixes first
                    prefixes.sort(key=len, reverse=True)
                    
                    for cand in candidates:
                        upper_cand = cand.upper()
                        if upper_cand in exclude_marks:
                            continue
                            
                        # Must either start with a digit (e.g. "15"), or start with a prefix followed by a digit
                        is_valid = False
                        if upper_cand and upper_cand[0].isdigit():
                            is_valid = True
                        else:
                            for p in prefixes:
                                if upper_cand.startswith(p):
                                    rest = upper_cand[len(p):].strip()
                                    if not rest or rest[0].isdigit() or rest[0] in ['_', '-', '*', 'M', 'X', '=']:
                                        is_valid = True
                                        break
                                        
                        if is_valid:
                            profile = cand
                            break
                            
                    # Absolute fallback if neither matched valid patterns
                    if not profile:
                        profile = lines[mat_idx - 1] if mat_idx > 0 else ""
                
                # Absolute fallback if material isn't matched at all
                if mat_idx == -1:
                    safe_fallback_prefixes = ["PL", "BL", "RO", "HEA", "HEB", "IPE", "UPN", "RHS", "CHS", "SHS", "CFRHS", "HFRHS", "HSQ", "CHESON", "VKR", "GST", "HEM", "HL", "HD", "UB", "UC", "UAP", "PFC", "KKR", "KRR", "ROHR", "VKT", "SQ", "PT", "RD", "FI", "FB", "FLT", "TABLA", "SHEET"]
                    for i, line in enumerate(lines):
                        upper_line = line.upper()
                        if any(upper_line.startswith(p) for p in safe_fallback_prefixes):
                            profile = line
                            if i + 1 < len(lines):
                                material = lines[i + 1]
                            break
                    
                    if not profile and len(lines) >= 6:
                        # Absolute fallback based on standard lines
                        profile = lines[3] if len(lines) > 3 else ""
                        material = lines[5] if len(lines) > 5 else ""
                        
        except Exception as e:
            self.log(f"Error reading {os.path.basename(filepath)}: {e}")
            
        return profile, material

    def verify_geometry(self, filepath, subfolder):
        # Negative Filters for safe exclusion
        try:
            with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
                lines = [l.strip() for l in f.readlines()]
                
            is_plate = "PLACI" in subfolder.upper()
            
            # 1. Check for plates having 3D operations
            if is_plate:
                in_op_block = False
                for line in lines:
                    if line in ["AK", "IK", "BO", "PU", "KO"]:
                        in_op_block = True
                        continue
                    if line == "EN":
                        in_op_block = False
                        continue
                        
                    if in_op_block and len(line) > 0:
                        face = line[0].lower()
                        # A flat plate should not have operations on top (o) or bottom (u) faces
                        if face in ['o', 'u']:
                            return False, "Placă cu operațiuni 3D (fețe o/u) - posibil profil exportat greșit"

            # 2. Check for Tubes/Profiles that are actually unwrapped plates
            if not is_plate and len(lines) > 9:
                # In standard DSTV, line 9 (0-indexed) is ALWAYS the Shape Code (Profile Type)
                shape_code = lines[9].upper()
                if shape_code == "B":
                    return False, "Profil/Țeavă exportată ca Placă 2D desfășurată (Shape Code 'B')"

        except Exception:
            pass # If read fails, don't fail the verification (assume true)
            
        return True, ""

    def process_files(self, src, action, opt_hot_cold, opt_separate_tubes, opt_verify_geometry):
        out = os.path.join(src, "Sorted")
        make_dir_with_permissions(out)
        self.log(f"Starting parsing and sorting from {src} to {out}")
        self.log(f"Action: {action}\n")

        nc_files = []
        for root, _, files in os.walk(src):
            for file in files:
                if file.lower().endswith(('.nc', '.nc1')):
                    nc_files.append(os.path.join(root, file))

        total_files = len(nc_files)
        if total_files == 0:
            self.log("No .nc files found in the source directory.")
            self.after(0, lambda: self.start_btn.configure(state="normal"))
            return

        report_data = []

        for i, filepath in enumerate(nc_files):
            filename = os.path.basename(filepath)
            profile, material = self.parse_nc_file(filepath)
            
            # Log the extracted fields to help debug false positives
            self.log(f"[{filename}] Parsed Profile: '{profile}', Material: '{material}'")

            subfolder, elem_type = self.determine_category(profile, material, opt_hot_cold, opt_separate_tubes)
            
            if opt_verify_geometry:
                is_valid, reason = self.verify_geometry(filepath, subfolder)
                if not is_valid:
                    self.log(f"  [!] GEOMETRY ERROR: {reason} -> Mutat la UNSORTED")
                    subfolder = "UNSORTED"
                    elem_type = "EROARE"

            dest_dir = os.path.join(out, subfolder)
            make_dir_with_permissions(dest_dir)
            dest_path = os.path.join(dest_dir, filename)
            
            try:
                if action == "Copy":
                    safe_copy(filepath, dest_path)
                    act_text = "Copied"
                else:
                    safe_move(filepath, dest_path)
                    act_text = "Moved"
                
                self.log(f"{act_text} {filename} -> {subfolder}")
                
                # --- Find and Move Associated Files (PDFs, DWGs, etc) ---
                nc_base = os.path.splitext(filename)[0]
                nc_id = re.sub(r'[^a-zA-Z0-9]', '', nc_base).upper()
                if nc_id:
                    # Build regex: e.g. "DSB2083" -> "^D[^a-zA-Z0-9]*S...3(?![a-zA-Z0-9])"
                    regex_str = "^" + "[^a-zA-Z0-9]*".join(list(nc_id)) + r"(?![a-zA-Z0-9])"
                    pattern = re.compile(regex_str, re.IGNORECASE)
                    
                    parent_dir = os.path.dirname(filepath)
                    for sib in os.listdir(parent_dir):
                        # Skip if it's the current file or another NC file
                        if sib == filename or sib.lower().endswith(('.nc', '.nc1')):
                            continue
                            
                        sib_base = os.path.splitext(sib)[0]
                        if pattern.match(sib_base):
                            sib_path = os.path.join(parent_dir, sib)
                            if not os.path.isfile(sib_path):
                                continue
                                
                            sib_dest_path = os.path.join(dest_dir, sib)
                            try:
                                if action == "Copy":
                                    safe_copy(sib_path, sib_dest_path)
                                else:
                                    safe_move(sib_path, sib_dest_path)
                                self.log(f"  + {act_text} associated file: {sib}")
                            except Exception as e:
                                self.log(f"  Failed to {action.lower()} associated file {sib}: {e}")
                # --------------------------------------------------------
                
                report_data.append({
                    "Nume Fisier": filename,
                    "Tip Element": elem_type,
                    "Sectiune": profile,
                    "Material Extras": material,
                    "Cale Destinatie": subfolder
                })
            except Exception as e:
                self.log(f"Failed to {action.lower()} {filename}: {e}")

            # Update progress
            progress_val = (i + 1) / total_files
            self.after(0, self.progress.set, progress_val)

        # Generate Excel Report
        self.log("\nGenerating Excel report...")
        report_path = os.path.join(out, "raport_piese_nc.xlsx")
        try:
            df = pd.DataFrame(report_data)
            safe_df_to_excel(df, report_path, index=False, log_callback=self.log)
            self.log(f"Report generated successfully: {report_path}")
        except Exception as e:
            self.log(f"Failed to generate Excel report: {e}")

        # Ensure full read/write permissions for all apps
        self.log("\nApplying full read/write permissions for all apps...")
        ensure_full_permissions_tree(out)
        self.log("Permissions configured successfully!")

        self.log("\nSorting complete!")
        self.after(0, lambda: self.start_btn.configure(state="normal"))

if __name__ == "__main__":
    app = NCSorterApp()
    app.mainloop()

"""


"""

__author__ = "iRNHO"
__contact__ = "Message 'iRNHO' on XBOX or 'irnho' on discord regarding any questions, feedback, bug-reporting etc."
__discord__ = "https://discord.gg/--------"


#################### SECTION BREAK ####################

##### CONSTANTS #####



#################### SECTION BREAK ####################

##### IMPORT STATEMENTS #####

import pandas as pd
import customtkinter as ctk

from pathlib import Path
from openpyxl import load_workbook
from platformdirs import user_data_dir


#################### SECTION BREAK ####################

##### DATA LOADING #####

root_directory = Path(user_data_dir("Damage Calculator", "iRNHO"))
assets_directory = root_directory / "assets"
data_directory = root_directory / "data"

#dev
assets_directory = Path(r"C:\Users\smorg\Documents\damage-calculator\assets")
data_directory = Path(r"C:\Users\smorg\Documents\damage-calculator\data")
#dev

data = {name: pd.read_excel(data_directory / f"{name}.xlsx", sheet_name=None) for name in ["Armor", "Skill", "Talent", "Weapon"]}


#################### SECTION BREAK ####################

##### GUI CLASS #####

class DamageCalculatorApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # Configure window
        self.title("iRNHO's Damage Calculator")
        self.geometry("1000x700")
        
        # Set window icon
        self.iconbitmap(assets_directory / "sub_machine_gun.ico")
        
        # Set appearance mode and load custom theme
        ctk.set_appearance_mode("dark")  # Modes: "system", "light", "dark"
        ctk.set_default_color_theme(str(assets_directory / "theme.json"))
        
        # Configure grid layout
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        
        # Create navigation frame at the top
        self.nav_frame = ctk.CTkFrame(self, height=60, corner_radius=0)
        self.nav_frame.grid(row=0, column=0, sticky="ew", padx=0, pady=0)
        self.nav_frame.grid_columnconfigure((0, 1, 2), weight=1)
        
        # Navigation buttons
        self.build_creator_btn = ctk.CTkButton(
            self.nav_frame, 
            text="Build Creator",
            command=self.show_build_creator,
            height=50,
            font=ctk.CTkFont(size=16, weight="bold")
        )
        self.build_creator_btn.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        
        self.build_tuning_btn = ctk.CTkButton(
            self.nav_frame,
            text="Build Tuning",
            command=self.show_build_tuning,
            height=50,
            font=ctk.CTkFont(size=16, weight="bold")
        )
        self.build_tuning_btn.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        
        self.damage_output_btn = ctk.CTkButton(
            self.nav_frame,
            text="Damage Output",
            command=self.show_damage_output,
            height=50,
            font=ctk.CTkFont(size=16, weight="bold")
        )
        self.damage_output_btn.grid(row=0, column=2, padx=5, pady=5, sticky="ew")
        
        # Create container frame for main content
        self.content_frame = ctk.CTkFrame(self, corner_radius=0)
        self.content_frame.grid(row=1, column=0, sticky="nsew", padx=0, pady=0)
        self.content_frame.grid_columnconfigure(0, weight=1)
        self.content_frame.grid_rowconfigure(0, weight=1)
        
        # Create the three main windows as frames
        self.build_creator_frame = ctk.CTkFrame(self.content_frame)
        self.build_tuning_frame = ctk.CTkFrame(self.content_frame)
        self.damage_output_frame = ctk.CTkFrame(self.content_frame)
        
        # Setup each frame
        self.setup_build_creator()
        self.setup_build_tuning()
        self.setup_damage_output()
        
        # Show Build Creator by default
        self.current_frame = None
        self.show_build_creator()
    
    def setup_build_creator(self):
        """Setup the Build Creator window"""
        # Make the frame scrollable
        self.build_creator_frame.grid_columnconfigure(0, weight=1)
        self.build_creator_frame.grid_rowconfigure(0, weight=1)
        
        # Create scrollable frame
        scrollable_frame = ctk.CTkScrollableFrame(self.build_creator_frame)
        scrollable_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        scrollable_frame.grid_columnconfigure((0, 1, 2), weight=1)
        
        # Define gear pieces with their icons
        # Left column: Mask, Body Armor, Holster
        left_column_gear = [
            ("Mask", "mask.png"),
            ("Body Armor", "body_armor.png"),
            ("Holster", "holster.png")
        ]
        
        # Right column: Backpack, Gloves, Kneepads
        right_column_gear = [
            ("Backpack", "backpack.png"),
            ("Gloves", "gloves.png"),
            ("Kneepads", "kneepads.png")
        ]
        
        # Create left column gear sections
        for idx, (gear_name, icon_file) in enumerate(left_column_gear):
            self.create_gear_section(scrollable_frame, gear_name, icon_file, idx, 0)
        
        # Create right column gear sections
        for idx, (gear_name, icon_file) in enumerate(right_column_gear):
            self.create_gear_section(scrollable_frame, gear_name, icon_file, idx, 1)
        
        # Add skills at the bottom of each column
        # Left column skill (row 3 after 3 gear pieces)
        self.create_skill_section(scrollable_frame, "Skill Left", "seeker_mine.png", 3, 0)
        
        # Right column skill (row 3 after 3 gear pieces)
        self.create_skill_section(scrollable_frame, "Skill Right", "seeker_mine.png", 3, 1)
        
        # Create a container for specialization and weapon in column 3 starting at row 0
        # This prevents them from affecting the row heights of the gear
        weapon_container = ctk.CTkFrame(scrollable_frame, fg_color="transparent")
        weapon_container.grid(row=0, column=2, rowspan=4, sticky="new", padx=0, pady=0)
        weapon_container.grid_columnconfigure(0, weight=1)
        
        # Add specialization above weapon in the container
        self.create_specialization_section(weapon_container, "Specialization", "demolitionist.png")
        
        # Add weapon in the container
        self.create_weapon_section_in_container(weapon_container, "Weapon", "sub_machine_gun.png")
    
    def create_specialization_section(self, parent, spec_name, icon_file):
        """Create a specialization section with dropdown"""
        # Main frame for specialization with border
        spec_frame = ctk.CTkFrame(parent, border_width=2)
        spec_frame.pack(padx=10, pady=10, fill="x")
        spec_frame.grid_columnconfigure(1, weight=1)
        spec_frame.grid_rowconfigure(0, weight=1)
        
        # Try to load and display icon - centered vertically
        try:
            from PIL import Image
            icon_path = assets_directory / icon_file
            if icon_path.exists():
                icon_image = ctk.CTkImage(
                    light_image=Image.open(icon_path),
                    dark_image=Image.open(icon_path),
                    size=(40, 40)
                )
                icon_label = ctk.CTkLabel(spec_frame, image=icon_image, text="")
                # Empty sticky means center both horizontally and vertically
                icon_label.grid(row=0, column=0, padx=10, pady=10, sticky="")
        except:
            pass
        
        # Create inner grid frame for dropdown
        dropdown_frame = ctk.CTkFrame(spec_frame)
        dropdown_frame.grid(row=0, column=1, sticky="ew", padx=10, pady=10)
        dropdown_frame.grid_columnconfigure(0, weight=1)
        
        # Specialization dropdown - defaults to "Select Specialization"
        spec_var = ctk.StringVar(value="Select Specialization")
        spec_values = ["Demolitionist", "Firewall", "Gunner", "Sharpshooter", "Survivalist", "Technician"]
        spec_dropdown = ctk.CTkComboBox(
            dropdown_frame,
            values=spec_values,
            variable=spec_var,
            command=lambda choice: self.update_specialization(spec_var, spec_dropdown),
            state="readonly"
        )
        spec_dropdown.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        
        # Store reference (optional, for future use)
        if not hasattr(self, 'specialization_var'):
            self.specialization_var = spec_var
        if not hasattr(self, 'specialization_dropdown'):
            self.specialization_dropdown = spec_dropdown
    
    def update_specialization(self, spec_var, spec_dropdown):
        """Update specialization dropdown to exclude current selection"""
        current_spec = spec_var.get()
        
        if current_spec == "Select Specialization":
            return
        
        # Get all specializations and exclude the current one
        all_specs = ["Demolitionist", "Firewall", "Gunner", "Sharpshooter", "Survivalist", "Technician"]
        available_specs = [s for s in all_specs if s != current_spec]
        
        # Update the dropdown values
        spec_dropdown.configure(values=available_specs)
        
        # Update weapon class options (enable/disable Signature Weapons)
        if hasattr(self, 'weapon_data'):
            self.update_weapon_class_options()
            
            # If Signature Weapons is currently selected, update the name and talent
            class_val = self.weapon_data['variables']['class'].get()
            if class_val == "Signature Weapons":
                # Re-run the weapon name update to get the new signature weapon
                self.update_weapon_name("Weapon", self.weapon_data)
            
            # Check if weapon name requires a different specialization
            if class_val != "Select Class" and class_val != "Signature Weapons":
                if 'name' in self.weapon_data['variables']:
                    name_val = self.weapon_data['variables']['name'].get()
                    
                    if name_val != "Select Name":
                        # Check if this weapon requires a different specialization
                        original_weapon_names = self.weapon_data.get('original_weapon_names', [])
                        needs_weapon_reset = False
                        
                        for orig_name in original_weapon_names:
                            if '(' in orig_name and ')' in orig_name:
                                paren_start = orig_name.rindex('(')
                                paren_end = orig_name.rindex(')')
                                base_name = orig_name[:paren_start].strip()
                                paren_content = orig_name[paren_start+1:paren_end].strip()
                                
                                # If this weapon matches and requires different spec, reset it
                                if self.is_specialization_name(paren_content):
                                    if base_name == name_val and paren_content != current_spec:
                                        needs_weapon_reset = True
                                        break
                        
                        if needs_weapon_reset:
                            # Reset weapon name and everything below it
                            self.weapon_data['variables']['name'].set("Select Name")
                            self.weapon_data['prev_name'] = "Select Name"
                            self.weapon_data['matched_weapon_name'] = "Select Name"
                            
                            # Clear all weapon attributes, mods, talents
                            for key in list(self.weapon_data['dropdowns'].keys()):
                                if key not in ['class', 'type', 'name']:
                                    self.weapon_data['dropdowns'][key].grid_forget()
                                    del self.weapon_data['dropdowns'][key]
                            for key in list(self.weapon_data['variables'].keys()):
                                if key not in ['class', 'type', 'name']:
                                    del self.weapon_data['variables'][key]
                            
                            # Update weapon name dropdown options
                            type_val = self.weapon_data['variables'].get('type', ctk.StringVar(value="Select Type")).get()
                            if type_val != "Select Type":
                                weapon_sheet = data["Weapon"][class_val]
                                all_names = [str(name) for name in weapon_sheet[weapon_sheet["Type"] == type_val]["Name"].tolist()]
                                
                                # Process names for display
                                processed_names = []
                                for name in all_names:
                                    if '(' in name and ')' in name:
                                        paren_start = name.rindex('(')
                                        paren_end = name.rindex(')')
                                        base_name = name[:paren_start].strip()
                                        paren_content = name[paren_start+1:paren_end].strip()
                                        
                                        if self.is_specialization_name(paren_content):
                                            if current_spec == paren_content:
                                                processed_names.append(base_name)
                                            else:
                                                processed_names.append(f"{base_name} (Select {paren_content})")
                                        else:
                                            processed_names.append(name)
                                    else:
                                        processed_names.append(name)
                                
                                if 'name' in self.weapon_data['dropdowns']:
                                    self.weapon_data['dropdowns']['name'].configure(values=processed_names)
                        else:
                            # Weapon doesn't need reset, but update name dropdown and check mods
                            type_val = self.weapon_data['variables'].get('type', ctk.StringVar(value="Select Type")).get()
                            if type_val != "Select Type":
                                weapon_sheet = data["Weapon"][class_val]
                                all_names = [str(name) for name in weapon_sheet[weapon_sheet["Type"] == type_val]["Name"].tolist()]
                                
                                # Process names for display
                                processed_names = []
                                for name in all_names:
                                    if '(' in name and ')' in name:
                                        paren_start = name.rindex('(')
                                        paren_end = name.rindex(')')
                                        base_name = name[:paren_start].strip()
                                        paren_content = name[paren_start+1:paren_end].strip()
                                        
                                        if self.is_specialization_name(paren_content):
                                            if current_spec == paren_content:
                                                processed_names.append(base_name)
                                            else:
                                                processed_names.append(f"{base_name} (Select {paren_content})")
                                        else:
                                            processed_names.append(name)
                                    else:
                                        processed_names.append(name)
                                
                                # Exclude current selection
                                if name_val != "Select Name":
                                    available_names = [n for n in processed_names if n != name_val]
                                else:
                                    available_names = processed_names
                                
                                if 'name' in self.weapon_data['dropdowns']:
                                    self.weapon_data['dropdowns']['name'].configure(values=available_names)
                            
                            # Check and reset weapon mods that require a different specialization
                            original_weapon_mods_dict = self.weapon_data.get('original_weapon_mods', {})
                    
                    for key, mod_var in list(self.weapon_data['variables'].items()):
                        if key.startswith('mod_'):
                            mod_val = mod_var.get()
                            original_mods = original_weapon_mods_dict.get(key, [])
                            
                            # Check if current mod requires a different specialization
                            for orig_mod in original_mods:
                                if '(' in orig_mod and ')' in orig_mod:
                                    paren_start = orig_mod.rindex('(')
                                    paren_end = orig_mod.rindex(')')
                                    base_mod = orig_mod[:paren_start].strip()
                                    paren_content = orig_mod[paren_start+1:paren_end].strip()
                                    
                                    # If this mod matches and requires different spec, reset it
                                    if self.is_specialization_name(paren_content):
                                        if base_mod == mod_val and paren_content != current_spec:
                                            # Reset this specific mod
                                            mod_index = int(key.split('_')[1])
                                            mod_types = ["Optics Mod", "Magazine Mod", "Underbarrel Mod", "Muzzle Mod"]
                                            if mod_index <= len(mod_types):
                                                column_name = mod_types[mod_index - 1]
                                                mod_var.set(f"Select {column_name}")
                                                self.weapon_data[f'prev_{key}'] = f"Select {column_name}"
                                            break
                    
                    # Update weapon mod dropdown options to show correct "(Select X)" indicators
                    # Only update if weapon hasn't been reset
                    if name_val != "Select Name" and 'matched_weapon_name' in self.weapon_data:
                        weapon_sheet = data["Weapon"][class_val]
                        type_val = self.weapon_data['variables'].get('type', ctk.StringVar(value="Select Type")).get()
                        if type_val != "Select Type":
                            # Use the matched weapon name for database lookup
                            lookup_name = self.weapon_data.get('matched_weapon_name', name_val)
                            
                            # Make sure the lookup name is valid (not "Select Name")
                            if lookup_name != "Select Name":
                                weapon_rows = weapon_sheet[(weapon_sheet["Name"] == lookup_name) & (weapon_sheet["Type"] == type_val)]
                                if not weapon_rows.empty:
                                    weapon_row = weapon_rows.iloc[0]
                                    
                                    mod_types = [
                                        ("Optics Mod", "Optics Mods"),
                                        ("Magazine Mod", "Magazine Mods"),
                                        ("Underbarrel Mod", "Underbarrel Mods"),
                                        ("Muzzle Mod", "Muzzle Mods")
                                    ]
                                    
                                    for i, (column_name, sheet_name) in enumerate(mod_types, start=1):
                                        key = f'mod_{i}'
                                        if key in self.weapon_data['dropdowns'] and key in self.weapon_data['variables']:
                                            mod_cell = weapon_row.get(column_name, pd.NA)
                                            if not pd.isna(mod_cell) and isinstance(mod_cell, str) and mod_cell.startswith("*"):
                                                mod_rail = mod_cell[1:]
                                                
                                                if sheet_name in data["Weapon"]:
                                                    mods_df = data["Weapon"][sheet_name]
                                                    if mod_rail in mods_df.columns:
                                                        all_mods = mods_df[mods_df[mod_rail] == '✓']['Stats'].tolist()
                                                        
                                                        # Process mods to handle specialization requirements
                                                        processed_mods = []
                                                        for mod in all_mods:
                                                            if '(' in mod and ')' in mod:
                                                                paren_start = mod.rindex('(')
                                                                paren_end = mod.rindex(')')
                                                                base_mod = mod[:paren_start].strip()
                                                                paren_content = mod[paren_start+1:paren_end].strip()
                                                                
                                                                if self.is_specialization_name(paren_content):
                                                                    if current_spec == paren_content:
                                                                        processed_mods.append(base_mod)
                                                                    else:
                                                                        processed_mods.append(f"{base_mod} (Select {paren_content})")
                                                                else:
                                                                    processed_mods.append(mod)
                                                            else:
                                                                processed_mods.append(mod)
                                                        
                                                        # Exclude current selection
                                                        current_mod = self.weapon_data['variables'][key].get()
                                                        if current_mod not in [f"Select {column_name}"]:
                                                            available_mods = [m for m in processed_mods if m != current_mod]
                                                        else:
                                                            available_mods = processed_mods
                                                        
                                                        self.weapon_data['dropdowns'][key].configure(values=available_mods)
        
        # Update all skill dropdowns to reflect specialization changes
        if hasattr(self, 'skill_sections'):
            for skill_name, skill_data in self.skill_sections.items():
                class_val = skill_data['variables']['class'].get()
                
                # Skip if no class selected or name variable doesn't exist yet
                if class_val == "Select Class" or 'name' not in skill_data['variables']:
                    continue
                
                name_val = skill_data['variables']['name'].get()
                
                if name_val == "Select Name":
                    # Still update dropdown options even if no name selected
                    if 'name' in skill_data['dropdowns']:
                        variant_row = data["Skill"]["Variants"][data["Skill"]["Variants"]["Class"] == class_val].iloc[0]
                        names_str = variant_row["Name"]
                        all_names = [n.strip() for n in names_str.split(';')]
                        
                        # Process names for display based on current specialization
                        processed_names = []
                        for name in all_names:
                            if '(' in name and ')' in name:
                                base_name = name[:name.index('(')].strip()
                                required_spec = name[name.index('(')+1:name.index(')')].strip()
                                if current_spec == required_spec:
                                    processed_names.append(base_name)
                                else:
                                    processed_names.append(f"{base_name} (Select {required_spec})")
                            else:
                                processed_names.append(name)
                        
                        skill_data['dropdowns']['name'].configure(values=processed_names)
                    continue
                
                # Check if current skill requires a specialization that doesn't match
                original_names = skill_data.get('original_names', [])
                
                # Find which original name corresponds to the current selection
                needs_reset = False
                for orig_name in original_names:
                    if '(' in orig_name and ')' in orig_name:
                        base_name = orig_name[:orig_name.index('(')].strip()
                        required_spec = orig_name[orig_name.index('(')+1:orig_name.index(')')].strip()
                        
                        # If current skill matches this base name and it requires a DIFFERENT specialization, reset it
                        if base_name == name_val and required_spec != current_spec:
                            needs_reset = True
                            break
                
                if needs_reset:
                    # Reset the skill name selection
                    skill_data['variables']['name'].set("Select Name")
                    skill_data['prev_name'] = "Select Name"
                    # Clear mods
                    for key in list(skill_data['dropdowns'].keys()):
                        if key not in ['class', 'name']:
                            skill_data['dropdowns'][key].grid_forget()
                            del skill_data['dropdowns'][key]
                    for key in list(skill_data['variables'].keys()):
                        if key not in ['class', 'name']:
                            del skill_data['variables'][key]
                else:
                    # Skill name doesn't need reset, but check individual mods
                    # Reset any mods that require a different specialization
                    original_mods_dict = skill_data.get('original_mods', {})
                    for key, mod_var in list(skill_data['variables'].items()):
                        if key.startswith('mod_'):
                            mod_val = mod_var.get()
                            original_mods = original_mods_dict.get(key, [])
                            
                            # Check if current mod requires a different specialization
                            for orig_mod in original_mods:
                                if '(' in orig_mod and ')' in orig_mod:
                                    paren_start = orig_mod.rindex('(')
                                    paren_end = orig_mod.rindex(')')
                                    base_mod = orig_mod[:paren_start].strip()
                                    paren_content = orig_mod[paren_start+1:paren_end].strip()
                                    
                                    # If this mod matches and requires different spec, reset it
                                    if self.is_specialization_name(paren_content):
                                        if base_mod == mod_val and paren_content != current_spec:
                                            # Reset this specific mod
                                            col_index = int(key.split('_')[1]) - 1
                                            mods_sheet_name = f"{class_val} Mods"
                                            if mods_sheet_name in data["Skill"]:
                                                mods_df = data["Skill"][mods_sheet_name]
                                                all_columns = mods_df.columns.tolist()
                                                mod_columns = [col for col in all_columns if col not in ['Stats', 'Last Checked']]
                                                if col_index < len(mod_columns):
                                                    col_name = mod_columns[col_index]
                                                    mod_var.set(f"Select {col_name} Mod")
                                                    skill_data[f'prev_{key}'] = f"Select {col_name} Mod"
                                            break
                
                # Update the name dropdown options (without recreating everything)
                if 'name' in skill_data['dropdowns']:
                    variant_row = data["Skill"]["Variants"][data["Skill"]["Variants"]["Class"] == class_val].iloc[0]
                    names_str = variant_row["Name"]
                    all_names = [n.strip() for n in names_str.split(';')]
                    
                    # Process names for display based on current specialization
                    processed_names = []
                    for name in all_names:
                        if '(' in name and ')' in name:
                            base_name = name[:name.index('(')].strip()
                            required_spec = name[name.index('(')+1:name.index(')')].strip()
                            if current_spec == required_spec:
                                processed_names.append(base_name)
                            else:
                                processed_names.append(f"{base_name} (Select {required_spec})")
                        else:
                            processed_names.append(name)
                    
                    # Exclude current selection
                    current_name = skill_data['variables']['name'].get()
                    if current_name != "Select Name":
                        available_names = [n for n in processed_names if n != current_name]
                    else:
                        available_names = processed_names
                    
                    skill_data['dropdowns']['name'].configure(values=available_names)
                
                # Update mod dropdown options to show correct "(Select X)" indicators
                if not needs_reset:
                    mods_sheet_name = f"{class_val} Mods"
                    if mods_sheet_name in data["Skill"]:
                        mods_df = data["Skill"][mods_sheet_name]
                        all_columns = mods_df.columns.tolist()
                        mod_columns = [col for col in all_columns if col not in ['Stats', 'Last Checked']]
                        
                        for i, col_name in enumerate(mod_columns):
                            key = f'mod_{i+1}'
                            if key in skill_data['dropdowns'] and key in skill_data['variables']:
                                # Get all mods for this column
                                all_mods = mods_df[mods_df[col_name] == '✓']['Stats'].tolist()
                                
                                # Process mods to handle specialization requirements
                                processed_mods = []
                                for mod in all_mods:
                                    if '(' in mod and ')' in mod:
                                        paren_start = mod.rindex('(')
                                        paren_end = mod.rindex(')')
                                        base_mod = mod[:paren_start].strip()
                                        paren_content = mod[paren_start+1:paren_end].strip()
                                        
                                        if self.is_specialization_name(paren_content):
                                            if current_spec == paren_content:
                                                processed_mods.append(base_mod)
                                            else:
                                                processed_mods.append(f"{base_mod} (Select {paren_content})")
                                        else:
                                            processed_mods.append(mod)
                                    else:
                                        processed_mods.append(mod)
                                
                                # Exclude current selection
                                current_mod = skill_data['variables'][key].get()
                                if current_mod not in [f"Select {col_name} Mod"]:
                                    available_mods = [m for m in processed_mods if m != current_mod]
                                else:
                                    available_mods = processed_mods
                                
                                skill_data['dropdowns'][key].configure(values=available_mods)

    def create_gear_section(self, parent, gear_name, icon_file, row_num, column_num):
        """Create a gear section with all dropdown fields"""
        # Main frame for this gear piece with border
        gear_frame = ctk.CTkFrame(parent, border_width=2)
        gear_frame.grid(row=row_num, column=column_num, sticky="nsew", padx=10, pady=10)
        gear_frame.grid_columnconfigure(1, weight=1)
        gear_frame.grid_rowconfigure(0, weight=1)
        
        # Try to load and display icon - centered vertically
        try:
            from PIL import Image
            icon_path = assets_directory / icon_file
            if icon_path.exists():
                icon_image = ctk.CTkImage(
                    light_image=Image.open(icon_path),
                    dark_image=Image.open(icon_path),
                    size=(40, 40)
                )
                icon_label = ctk.CTkLabel(gear_frame, image=icon_image, text="")
                # Empty sticky means center both horizontally and vertically
                icon_label.grid(row=0, column=0, padx=10, pady=10, sticky="")
        except:
            pass
        
        # Create inner grid frame for dropdowns
        dropdown_frame = ctk.CTkFrame(gear_frame)
        dropdown_frame.grid(row=0, column=1, sticky="ew", padx=10, pady=10)
        dropdown_frame.grid_columnconfigure(0, weight=1)
        
        # Store references to all dropdowns for this gear piece
        gear_data = {
            'slot': gear_name,
            'frame': dropdown_frame,
            'dropdowns': {},
            'variables': {}
        }
        
        # Type dropdown - always shown
        type_var = ctk.StringVar(value="Select Type")
        type_dropdown = ctk.CTkComboBox(
            dropdown_frame, 
            values=["Improvised", "Brand Set", "Gear Set", "Named", "Exotic"],
            variable=type_var,
            command=lambda choice: self.update_gear(gear_name, gear_data),
            state="readonly"
        )
        type_dropdown.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        
        gear_data['variables']['type'] = type_var
        gear_data['dropdowns']['type'] = type_dropdown
        
        # Store gear_data for later access
        if not hasattr(self, 'gear_sections'):
            self.gear_sections = {}
        self.gear_sections[gear_name] = gear_data
    
    def update_all_type_dropdowns(self):
        """Update all gear type dropdowns to handle Exotic exclusivity"""
        if not hasattr(self, 'gear_sections'):
            return
        
        # Check if any gear piece has Exotic selected
        has_exotic = any(
            gear_data['variables']['type'].get() == "Exotic" 
            for gear_data in self.gear_sections.values()
        )
        
        # Determine available types
        if has_exotic:
            base_types = ["Improvised", "Brand Set", "Gear Set", "Named"]
        else:
            base_types = ["Improvised", "Brand Set", "Gear Set", "Named", "Exotic"]
        
        # Update each gear's type dropdown
        for gear_name, gear_data in self.gear_sections.items():
            current_type = gear_data['variables']['type'].get()
            type_dropdown = gear_data['dropdowns']['type']
            
            # Start with base types (already accounts for Exotic exclusivity)
            # Then exclude the current selection from the available options
            available_types = [t for t in base_types if t != current_type]
            
            # Update the dropdown values
            type_dropdown.configure(values=available_types)
    
    def update_gear(self, gear_name, gear_data):
        """Dynamically update gear dropdowns based on selections"""
        type_val = gear_data['variables']['type'].get()
        frame = gear_data['frame']
        
        # Update ALL type dropdowns to handle Exotic exclusivity
        self.update_all_type_dropdowns()
        
        # Clear existing dropdowns except type
        self.update_all_type_dropdowns()
        
        # Clear existing dropdowns except type
        for key in list(gear_data['dropdowns'].keys()):
            if key != 'type':
                gear_data['dropdowns'][key].grid_forget()
                del gear_data['dropdowns'][key]
        for key in list(gear_data['variables'].keys()):
            if key != 'type':
                del gear_data['variables'][key]
        
        row_idx = 1
        
        # Name dropdown
        name_var = ctk.StringVar()
        
        if type_val == "Improvised":
            # Improvised has fixed name
            name_var.set(f"Improvised {gear_name}")
            name_dropdown = ctk.CTkComboBox(frame, values=[f"Improvised {gear_name}"], variable=name_var, state="disabled")
        else:
            # Get names based on type
            if type_val == "Brand Set":
                names = data["Armor"]["Brand Sets"]["Name"].tolist()
            elif type_val == "Gear Set":
                names = data["Armor"]["Gear Sets"]["Name"].tolist()
            elif type_val == "Named":
                names = data["Armor"]["Named"][data["Armor"]["Named"]["Slot"] == gear_name]["Name"].tolist()
            elif type_val == "Exotic":
                names = data["Armor"]["Exotic"][data["Armor"]["Exotic"]["Slot"] == gear_name]["Name"].tolist()
            else:
                names = []
            
            name_var.set("Select Name")
            name_dropdown = ctk.CTkComboBox(
                frame,
                values=names,
                variable=name_var,
                command=lambda choice: self.update_gear_attributes(gear_name, gear_data),
                state="readonly"
            )
        
        name_dropdown.grid(row=row_idx, column=0, padx=5, pady=5, sticky="ew")
        gear_data['variables']['name'] = name_var
        gear_data['dropdowns']['name'] = name_dropdown
        row_idx += 1
        
        # If Improvised, automatically update attributes
        if type_val == "Improvised":
            self.update_gear_attributes(gear_name, gear_data, preserve_selections=False)
    
    def update_gear_attributes(self, gear_name, gear_data, preserve_selections=True):
        """Update core attributes, attributes, mods, and talents based on gear selection"""
        type_val = gear_data['variables']['type'].get()
        name_val = gear_data['variables']['name'].get()
        frame = gear_data['frame']
        
        if name_val in ["Select Name", "Select Type"]:
            return
        
        # Preserve existing selections if requested
        saved_selections = {}
        if preserve_selections:
            for key, var in gear_data['variables'].items():
                if key not in ['type', 'name']:
                    saved_selections[key] = var.get()
        
        # Clear existing attribute dropdowns
        for key in list(gear_data['dropdowns'].keys()):
            if key not in ['type', 'name']:
                gear_data['dropdowns'][key].grid_forget()
                del gear_data['dropdowns'][key]
        for key in list(gear_data['variables'].keys()):
            if key not in ['type', 'name']:
                del gear_data['variables'][key]
        
        row_idx = 2
        
        # Determine what attributes this gear piece has
        if type_val in ["Improvised", "Brand Set"]:
            core_cells = ["*", pd.NA, pd.NA]
            attr_cells = ["*", "*", pd.NA]
            mod_cells = ["*" if type_val == "Improvised" or gear_name in ["Mask", "Body Armor", "Backpack"] else pd.NA, pd.NA]
            talent_cells = ["*" if gear_name in ["Body Armor", "Backpack"] else pd.NA, pd.NA]
        elif type_val == "Gear Set":
            core_cells = ["*", pd.NA, pd.NA]
            attr_cells = ["*", pd.NA, pd.NA]
            mod_cells = ["*" if gear_name in ["Mask", "Body Armor", "Backpack"] else pd.NA, pd.NA]
            # For Gear Sets, get the talent from the Gear Sets sheet
            if gear_name in ["Body Armor", "Backpack"]:
                gear_set_row = data["Armor"]["Gear Sets"][data["Armor"]["Gear Sets"]["Name"] == name_val].iloc[0]
                talent_value = gear_set_row.get(f"{gear_name} Talent", pd.NA)
                talent_cells = [talent_value, pd.NA]
            else:
                talent_cells = [pd.NA, pd.NA]
        else:
            # Named or Exotic
            sheet_name = "Named" if type_val == "Named" else "Exotic"
            gear_row = data["Armor"][sheet_name][data["Armor"][sheet_name]["Name"] == name_val].iloc[0]
            core_cells = [gear_row.get(f"Core Attribute {i}", pd.NA) for i in [1, 2, 3]]
            attr_cells = [gear_row.get(f"Attribute {i}", pd.NA) for i in [1, 2, 3]]
            mod_cells = [gear_row.get(f"Mod Slot {i}", pd.NA) for i in [1, 2]]
            talent_cells = [gear_row.get(f"Talent {i}", pd.NA) for i in [1, 2]]
        
        # Validate saved selections against new gear's allowed values
        # This ensures switching from Named/Exotic gear doesn't leave invalid values
        
        # Validate core attributes
        all_cores = ["15.0% Weapon Damage", "170,000 Armor", "1 Skill Tier"]
        for i, cell_val in enumerate(core_cells):
            key = f'core_{i+1}'
            if key in saved_selections and saved_selections[key] not in ["Select Core Attribute"]:
                if pd.isna(cell_val):
                    saved_selections[key] = "Select Core Attribute"
                elif cell_val == "*":
                    if saved_selections[key] not in all_cores:
                        saved_selections[key] = "Select Core Attribute"
                elif isinstance(cell_val, str) and cell_val.startswith("!"):
                    excluded = cell_val[1:]
                    if saved_selections[key] == excluded or saved_selections[key] not in all_cores:
                        saved_selections[key] = "Select Core Attribute"
                # For fixed values, they'll be overwritten anyway
        
        # Validate attributes
        all_attrs = data["Armor"]["Attributes"]["Stats"].tolist()
        for i, cell_val in enumerate(attr_cells):
            key = f'attr_{i+1}'
            if key in saved_selections and saved_selections[key] not in ["Select Attribute"]:
                if pd.isna(cell_val):
                    saved_selections[key] = "Select Attribute"
                elif cell_val == "*":
                    if saved_selections[key] not in all_attrs:
                        saved_selections[key] = "Select Attribute"
                elif isinstance(cell_val, str) and cell_val.startswith("*") and len(cell_val) > 1:
                    # Type-specific attribute (e.g., "*Offensive")
                    attr_type = cell_val[1:]
                    type_attrs = data["Armor"]["Attributes"][data["Armor"]["Attributes"]["Type"] == attr_type]["Stats"].tolist()
                    if saved_selections[key] not in type_attrs:
                        saved_selections[key] = "Select Attribute"
                elif isinstance(cell_val, str) and cell_val.startswith("!"):
                    excluded = cell_val[1:]
                    if saved_selections[key] == excluded or saved_selections[key] not in all_attrs:
                        saved_selections[key] = "Select Attribute"
                # For fixed values, they'll be overwritten anyway
        
        # Validate mods
        all_mods = data["Armor"]["Mods"]["Stats"].tolist()
        for i, cell_val in enumerate(mod_cells):
            key = f'mod_{i+1}'
            if key in saved_selections and saved_selections[key] not in ["Select Mod"]:
                if pd.isna(cell_val):
                    saved_selections[key] = "Select Mod"
                elif saved_selections[key] not in all_mods:
                    saved_selections[key] = "Select Mod"
        
        # Validate talents
        for i, cell_val in enumerate(talent_cells):
            key = f'talent_{i+1}'
            if key in saved_selections and saved_selections[key] not in ["Select Talent"]:
                if pd.isna(cell_val):
                    saved_selections[key] = "Select Talent"
                elif cell_val == "*":
                    # Check against universal talents for this slot
                    all_talents = data["Armor"]["Talents"][data["Armor"]["Talents"]["Slot"] == gear_name]["Name"].tolist()
                    if saved_selections[key] not in all_talents:
                        saved_selections[key] = "Select Talent"
                # For fixed values (Gear Set talents or Named/Exotic talents), they'll be overwritten
        
        # Determine FINAL values for each slot (accounting for fixed values that will override saved selections)
        # This is critical for proper exclusion lists - we need to know what will ACTUALLY be selected
        final_core_values = {}
        for i, cell_val in enumerate(core_cells):
            if pd.isna(cell_val):
                continue
            key = f'core_{i+1}'
            # Fixed values always override saved selections
            if isinstance(cell_val, str) and cell_val not in ["*"] and not cell_val.startswith("!"):
                final_core_values[key] = cell_val
            else:
                # User-selectable slots keep their saved value (or default)
                final_core_values[key] = saved_selections.get(key, "Select Core Attribute")
        
        final_attr_values = {}
        for i, cell_val in enumerate(attr_cells):
            if pd.isna(cell_val):
                continue
            key = f'attr_{i+1}'
            # Fixed values always override saved selections
            if isinstance(cell_val, str) and cell_val not in ["*"] and not cell_val.startswith("!") and not cell_val.startswith("*"):
                final_attr_values[key] = cell_val
            else:
                # User-selectable slots keep their saved value (or default)
                final_attr_values[key] = saved_selections.get(key, "Select Attribute")
        
        # Core Attributes
        for i, cell_val in enumerate(core_cells):
            if pd.isna(cell_val):
                continue
            
            # Use final value for this slot
            key = f'core_{i+1}'
            saved_value = final_core_values[key]
            core_var = ctk.StringVar(value=saved_value)
            
            if cell_val == "*":
                # User can select any core attribute
                # Exclude final values from OTHER core slots (not this one)
                current_selections = [final_core_values[f'core_{j+1}'] 
                                     for j in range(len(core_cells)) 
                                     if j != i and f'core_{j+1}' in final_core_values 
                                     and final_core_values[f'core_{j+1}'] not in ["Select Core Attribute"]]
                # Also exclude this dropdown's own current selection
                if saved_value not in ["Select Core Attribute"]:
                    current_selections.append(saved_value)
                available_cores = [c for c in ["15.0% Weapon Damage", "170,000 Armor", "1 Skill Tier"] 
                                  if c not in current_selections]
                core_dropdown = ctk.CTkComboBox(
                    frame, 
                    values=available_cores, 
                    variable=core_var,
                    command=lambda choice, gear=gear_name, gdata=gear_data: self.update_gear_attributes(gear, gdata, preserve_selections=True),
                    state="readonly"
                )
            elif isinstance(cell_val, str) and cell_val.startswith("!"):
                # Cannot select this specific one
                excluded = cell_val[1:]
                current_selections = [final_core_values[f'core_{j+1}'] 
                                     for j in range(len(core_cells)) 
                                     if j != i and f'core_{j+1}' in final_core_values 
                                     and final_core_values[f'core_{j+1}'] not in ["Select Core Attribute"]]
                # Also exclude this dropdown's own current selection
                if saved_value not in ["Select Core Attribute"]:
                    current_selections.append(saved_value)
                available_cores = [c for c in ["15.0% Weapon Damage", "170,000 Armor", "1 Skill Tier"] 
                                  if c != excluded and c not in current_selections]
                core_dropdown = ctk.CTkComboBox(
                    frame, 
                    values=available_cores, 
                    variable=core_var,
                    command=lambda choice, gear=gear_name, gdata=gear_data: self.update_gear_attributes(gear, gdata, preserve_selections=True),
                    state="readonly"
                )
            else:
                # Fixed core attribute
                core_var.set(cell_val)
                core_dropdown = ctk.CTkComboBox(frame, values=[cell_val], variable=core_var, state="disabled")
            
            core_dropdown.grid(row=row_idx, column=0, padx=5, pady=5, sticky="ew")
            gear_data['variables'][f'core_{i+1}'] = core_var
            gear_data['dropdowns'][f'core_{i+1}'] = core_dropdown
            row_idx += 1
        
        # Attributes
        for i, cell_val in enumerate(attr_cells):
            if pd.isna(cell_val):
                continue
            
            # Use final value for this slot
            key = f'attr_{i+1}'
            saved_value = final_attr_values[key]
            attr_var = ctk.StringVar(value=saved_value)
            
            if cell_val == "*":
                # User can select any attribute
                # Exclude final values from OTHER attribute slots (not this one)
                current_selections = [final_attr_values[f'attr_{j+1}'] 
                                     for j in range(len(attr_cells)) 
                                     if j != i and f'attr_{j+1}' in final_attr_values 
                                     and final_attr_values[f'attr_{j+1}'] not in ["Select Attribute"]]
                # Also exclude this dropdown's own current selection
                if saved_value not in ["Select Attribute"]:
                    current_selections.append(saved_value)
                available_attrs = [a for a in data["Armor"]["Attributes"]["Stats"].tolist() 
                                  if a not in current_selections]
                attr_dropdown = ctk.CTkComboBox(
                    frame, 
                    values=available_attrs, 
                    variable=attr_var,
                    command=lambda choice, gear=gear_name, gdata=gear_data: self.update_gear_attributes(gear, gdata, preserve_selections=True),
                    state="readonly"
                )
            elif isinstance(cell_val, str) and cell_val.startswith("*") and len(cell_val) > 1:
                # Specific type of attribute (e.g., "*Offensive")
                attr_type = cell_val[1:]
                current_selections = [final_attr_values[f'attr_{j+1}'] 
                                     for j in range(len(attr_cells)) 
                                     if j != i and f'attr_{j+1}' in final_attr_values 
                                     and final_attr_values[f'attr_{j+1}'] not in ["Select Attribute"]]
                # Also exclude this dropdown's own current selection
                if saved_value not in ["Select Attribute"]:
                    current_selections.append(saved_value)
                available_attrs = data["Armor"]["Attributes"][data["Armor"]["Attributes"]["Type"] == attr_type]["Stats"].tolist()
                available_attrs = [a for a in available_attrs if a not in current_selections]
                attr_dropdown = ctk.CTkComboBox(
                    frame, 
                    values=available_attrs, 
                    variable=attr_var,
                    command=lambda choice, gear=gear_name, gdata=gear_data: self.update_gear_attributes(gear, gdata, preserve_selections=True),
                    state="readonly"
                )
            elif isinstance(cell_val, str) and cell_val.startswith("!"):
                # Cannot select this specific one
                excluded = cell_val[1:]
                current_selections = [final_attr_values[f'attr_{j+1}'] 
                                     for j in range(len(attr_cells)) 
                                     if j != i and f'attr_{j+1}' in final_attr_values 
                                     and final_attr_values[f'attr_{j+1}'] not in ["Select Attribute"]]
                # Also exclude this dropdown's own current selection
                if saved_value not in ["Select Attribute"]:
                    current_selections.append(saved_value)
                available_attrs = [a for a in data["Armor"]["Attributes"]["Stats"].tolist() 
                                  if a != excluded and a not in current_selections]
                attr_dropdown = ctk.CTkComboBox(
                    frame, 
                    values=available_attrs, 
                    variable=attr_var,
                    command=lambda choice, gear=gear_name, gdata=gear_data: self.update_gear_attributes(gear, gdata, preserve_selections=True),
                    state="readonly"
                )
            else:
                # Fixed attribute
                attr_var.set(cell_val)
                attr_dropdown = ctk.CTkComboBox(frame, values=[cell_val], variable=attr_var, state="disabled")
            
            attr_dropdown.grid(row=row_idx, column=0, padx=5, pady=5, sticky="ew")
            gear_data['variables'][f'attr_{i+1}'] = attr_var
            gear_data['dropdowns'][f'attr_{i+1}'] = attr_dropdown
            row_idx += 1
        
        # Mods
        for i, cell_val in enumerate(mod_cells):
            if pd.isna(cell_val):
                continue
            
            saved_value = saved_selections.get(f'mod_{i+1}', "Select Mod")
            mod_var = ctk.StringVar(value=saved_value)
            
            # Get all available mods and exclude this dropdown's current selection
            all_mods = data["Armor"]["Mods"]["Stats"].tolist()
            if saved_value not in ["Select Mod"]:
                available_mods = [m for m in all_mods if m != saved_value]
            else:
                available_mods = all_mods
            
            mod_dropdown = ctk.CTkComboBox(
                frame, 
                values=available_mods, 
                variable=mod_var, 
                command=lambda choice, gear=gear_name, gdata=gear_data: self.update_gear_attributes(gear, gdata, preserve_selections=True),
                state="readonly"
            )
            mod_dropdown.grid(row=row_idx, column=0, padx=5, pady=5, sticky="ew")
            gear_data['variables'][f'mod_{i+1}'] = mod_var
            gear_data['dropdowns'][f'mod_{i+1}'] = mod_dropdown
            row_idx += 1
        
        # Talents
        for i, cell_val in enumerate(talent_cells):
            if pd.isna(cell_val):
                continue
            
            saved_value = saved_selections.get(f'talent_{i+1}', "Select Talent")
            talent_var = ctk.StringVar(value=saved_value)
            
            if cell_val == "*":
                # User can select any universal talent for this slot
                all_talents = data["Armor"]["Talents"][data["Armor"]["Talents"]["Slot"] == gear_name]["Name"].tolist()
                # Exclude this dropdown's current selection
                if saved_value not in ["Select Talent"]:
                    available_talents = [t for t in all_talents if t != saved_value]
                else:
                    available_talents = all_talents
                talent_dropdown = ctk.CTkComboBox(
                    frame, 
                    values=available_talents, 
                    variable=talent_var,
                    command=lambda choice, gear=gear_name, gdata=gear_data: self.update_gear_attributes(gear, gdata, preserve_selections=True),
                    state="readonly"
                )
            elif type_val == "Gear Set":
                # Fixed gear set talent
                talent_var.set(cell_val)
                talent_dropdown = ctk.CTkComboBox(frame, values=[cell_val], variable=talent_var, state="disabled")
            else:
                # Fixed talent
                talent_var.set(cell_val)
                talent_dropdown = ctk.CTkComboBox(frame, values=[cell_val], variable=talent_var, state="disabled")
            
            talent_dropdown.grid(row=row_idx, column=0, padx=5, pady=5, sticky="ew")
            gear_data['variables'][f'talent_{i+1}'] = talent_var
            gear_data['dropdowns'][f'talent_{i+1}'] = talent_dropdown
            row_idx += 1
    
    def create_skill_section(self, parent, skill_name, icon_file, row_num, column_num):
        """Create a skill section with dropdown fields"""
        # Main frame for this skill with slight visual separation
        skill_frame = ctk.CTkFrame(parent, border_width=2)
        skill_frame.grid(row=row_num, column=column_num, sticky="nsew", padx=10, pady=(20, 10))
        skill_frame.grid_columnconfigure(1, weight=1)
        skill_frame.grid_rowconfigure(0, weight=1)
        
        # Try to load and display icon - centered vertically
        try:
            from PIL import Image
            icon_path = assets_directory / icon_file
            if icon_path.exists():
                icon_image = ctk.CTkImage(
                    light_image=Image.open(icon_path),
                    dark_image=Image.open(icon_path),
                    size=(40, 40)
                )
                icon_label = ctk.CTkLabel(skill_frame, image=icon_image, text="")
                # Empty sticky means center both horizontally and vertically
                icon_label.grid(row=0, column=0, padx=10, pady=10, sticky="")
        except:
            pass
        
        # Create inner grid frame for dropdowns
        dropdown_frame = ctk.CTkFrame(skill_frame)
        dropdown_frame.grid(row=0, column=1, sticky="ew", padx=10, pady=10)
        dropdown_frame.grid_columnconfigure(0, weight=1)
        
        # Store references to all dropdowns for this skill
        skill_data = {
            'name': skill_name,
            'frame': dropdown_frame,
            'dropdowns': {},
            'variables': {}
        }
        
        # Class dropdown - always shown
        class_var = ctk.StringVar(value="Select Class")
        class_values = data["Skill"]["Variants"]["Class"].tolist()
        class_dropdown = ctk.CTkComboBox(
            dropdown_frame,
            values=class_values,
            variable=class_var,
            command=lambda choice: self.update_skill(skill_name, skill_data),
            state="readonly"
        )
        class_dropdown.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        
        skill_data['variables']['class'] = class_var
        skill_data['dropdowns']['class'] = class_dropdown
        
        # Store skill_data for later access
        if not hasattr(self, 'skill_sections'):
            self.skill_sections = {}
        self.skill_sections[skill_name] = skill_data
    
    def is_specialization_name(self, text):
        """Check if the given text is a valid specialization name"""
        specializations = ["Demolitionist", "Firewall", "Gunner", "Sharpshooter", "Survivalist", "Technician"]
        return text in specializations
    
    def update_all_skill_class_dropdowns(self):
        """Update all skill class dropdowns to handle class exclusivity"""
        if not hasattr(self, 'skill_sections'):
            return
        
        # Get all base classes
        all_classes = data["Skill"]["Variants"]["Class"].tolist()
        
        # Update each skill's class dropdown
        for skill_name, skill_data in self.skill_sections.items():
            current_class = skill_data['variables']['class'].get()
            class_dropdown = skill_data['dropdowns']['class']
            
            # Start with all classes
            available_classes = all_classes.copy()
            
            # Exclude classes selected in OTHER skills
            for other_skill_name, other_skill_data in self.skill_sections.items():
                if other_skill_name != skill_name:
                    other_class = other_skill_data['variables']['class'].get()
                    if other_class != "Select Class" and other_class in available_classes:
                        available_classes.remove(other_class)
            
            # Also exclude this skill's current selection
            if current_class != "Select Class" and current_class in available_classes:
                available_classes.remove(current_class)
            
            # Update the dropdown values
            class_dropdown.configure(values=available_classes)
    
    def update_skill(self, skill_name, skill_data):
        """Dynamically update skill dropdowns based on selections"""
        class_val = skill_data['variables']['class'].get()
        frame = skill_data['frame']
        
        # Update ALL skill class dropdowns to handle exclusivity
        self.update_all_skill_class_dropdowns()
        
        if class_val == "Select Class":
            return
        
        # Clear existing dropdowns except class
        for key in list(skill_data['dropdowns'].keys()):
            if key != 'class':
                skill_data['dropdowns'][key].grid_forget()
                del skill_data['dropdowns'][key]
        for key in list(skill_data['variables'].keys()):
            if key != 'class':
                del skill_data['variables'][key]
        
        row_idx = 1
        
        # Name dropdown - get names from Variants sheet
        variant_row = data["Skill"]["Variants"][data["Skill"]["Variants"]["Class"] == class_val].iloc[0]
        names_str = variant_row["Name"]
        names = [n.strip() for n in names_str.split(';')]
        
        # Process names to handle specialization requirements
        processed_names = []
        spec_val = self.specialization_var.get() if hasattr(self, 'specialization_var') else "Select Specialization"
        
        for name in names:
            # Check if name has specialization requirement (contains parentheses)
            if '(' in name and ')' in name:
                # Extract base name and required specialization
                base_name = name[:name.index('(')].strip()
                required_spec = name[name.index('(')+1:name.index(')')].strip()
                
                # Check if user has the required specialization
                if spec_val == required_spec:
                    # User has correct spec, show just the base name
                    processed_names.append(base_name)
                else:
                    # User doesn't have the spec, show with "Select X" indicator
                    processed_names.append(f"{base_name} (Select {required_spec})")
            else:
                # No specialization requirement
                processed_names.append(name)
        
        name_var = ctk.StringVar(value="Select Name")
        name_dropdown = ctk.CTkComboBox(
            frame,
            values=processed_names,
            variable=name_var,
            command=lambda choice: self.update_skill_name_selection(skill_name, skill_data),
            state="readonly"
        )
        name_dropdown.grid(row=row_idx, column=0, padx=5, pady=5, sticky="ew")
        skill_data['variables']['name'] = name_var
        skill_data['dropdowns']['name'] = name_dropdown
        
        # Store original names for reference
        skill_data['original_names'] = names
    
    def update_skill_name_selection(self, skill_name, skill_data):
        """Handle skill name selection and validate specialization requirements"""
        name_val = skill_data['variables']['name'].get()
        class_val = skill_data['variables']['class'].get()
        
        if name_val == "Select Name":
            return
        
        # Check if this is a disabled specialization-locked skill
        if "(Select " in name_val:
            # Revert to previous selection
            prev_name = skill_data.get('prev_name', "Select Name")
            skill_data['variables']['name'].set(prev_name)
            return
        
        # Store the current selection for potential revert
        skill_data['prev_name'] = name_val
        
        # Continue to update mods
        self.update_skill_mods(skill_name, skill_data)
    
    def update_skill_mods(self, skill_name, skill_data, preserve_selections=True):
        """Update skill mod dropdowns based on class selection"""
        class_val = skill_data['variables']['class'].get()
        name_val = skill_data['variables']['name'].get()
        frame = skill_data['frame']
        
        if name_val == "Select Name":
            return
        
        # Match the selected name against original names to find the correct variant
        # The displayed name might be just the base name (e.g., "Striker") while original is "Striker (Firewall)"
        original_names = skill_data.get('original_names', [])
        matched_name = name_val
        
        # Try to find matching original name
        for orig_name in original_names:
            # Check if this original name matches (either exact or base name match)
            if orig_name == name_val:
                matched_name = orig_name
                break
            elif '(' in orig_name:
                base_name = orig_name[:orig_name.index('(')].strip()
                if base_name == name_val:
                    matched_name = orig_name
                    break
        
        # Update name dropdown to exclude current selection
        variant_row = data["Skill"]["Variants"][data["Skill"]["Variants"]["Class"] == class_val].iloc[0]
        names_str = variant_row["Name"]
        all_names = [n.strip() for n in names_str.split(';')]
        
        # Process names for display (same logic as in update_skill)
        spec_val = self.specialization_var.get() if hasattr(self, 'specialization_var') else "Select Specialization"
        processed_names = []
        for name in all_names:
            if '(' in name and ')' in name:
                base_name = name[:name.index('(')].strip()
                required_spec = name[name.index('(')+1:name.index(')')].strip()
                if spec_val == required_spec:
                    processed_names.append(base_name)
                else:
                    processed_names.append(f"{base_name} (Select {required_spec})")
            else:
                processed_names.append(name)
        
        # Exclude current selection from processed names
        available_names = [n for n in processed_names if n != name_val]
        skill_data['dropdowns']['name'].configure(values=available_names)
        
        # Preserve existing selections if requested
        saved_selections = {}
        if preserve_selections:
            for key, var in skill_data['variables'].items():
                if key not in ['class', 'name']:
                    saved_selections[key] = var.get()
        
        # Clear existing mod dropdowns
        for key in list(skill_data['dropdowns'].keys()):
            if key not in ['class', 'name']:
                skill_data['dropdowns'][key].grid_forget()
                del skill_data['dropdowns'][key]
        for key in list(skill_data['variables'].keys()):
            if key not in ['class', 'name']:
                del skill_data['variables'][key]
        
        # Get the mods dataframe for this class
        mods_sheet_name = f"{class_val} Mods"
        if mods_sheet_name not in data["Skill"]:
            return
        
        mods_df = data["Skill"][mods_sheet_name]
        
        # Get column names (excluding 'Stats' and 'Last Checked')
        all_columns = mods_df.columns.tolist()
        mod_columns = [col for col in all_columns if col not in ['Stats', 'Last Checked']]
        
        # Determine final values for each mod slot (for proper exclusion lists)
        final_mod_values = {}
        for i, col_name in enumerate(mod_columns):
            key = f'mod_{i+1}'
            final_mod_values[key] = saved_selections.get(key, f"Select {col_name} Mod")
        
        row_idx = 2
        
        # Create mod dropdowns
        for i, col_name in enumerate(mod_columns):
            key = f'mod_{i+1}'
            saved_value = final_mod_values[key]
            
            mod_var = ctk.StringVar(value=saved_value)
            
            # Get available mods for this column (where column value is '✓')
            all_mods = mods_df[mods_df[col_name] == '✓']['Stats'].tolist()
            
            # Process mods to handle specialization requirements
            spec_val = self.specialization_var.get() if hasattr(self, 'specialization_var') else "Select Specialization"
            processed_mods = []
            original_mods = []  # Store original mod names for matching
            
            for mod in all_mods:
                original_mods.append(mod)
                # Check if mod has specialization requirement
                if '(' in mod and ')' in mod:
                    paren_start = mod.rindex('(')  # Use rindex to get the last occurrence
                    paren_end = mod.rindex(')')
                    base_mod = mod[:paren_start].strip()
                    paren_content = mod[paren_start+1:paren_end].strip()
                    
                    # Check if parentheses contain a specialization
                    if self.is_specialization_name(paren_content):
                        # This is a specialization-specific mod
                        if spec_val == paren_content:
                            # User has correct spec, show just the base name
                            processed_mods.append(base_mod)
                        else:
                            # User doesn't have the spec, show with "Select X" indicator
                            processed_mods.append(f"{base_mod} (Select {paren_content})")
                    else:
                        # Not a specialization, keep the full mod name
                        processed_mods.append(mod)
                else:
                    # No parentheses, keep as is
                    processed_mods.append(mod)
            
            # Match saved value against processed mods
            # If saved value is a base name, it should match the processed version
            display_value = saved_value
            if saved_value not in [f"Select {col_name} Mod"]:
                # Try to find the processed version of the saved value
                for j, orig_mod in enumerate(original_mods):
                    if '(' in orig_mod and ')' in orig_mod:
                        paren_start = orig_mod.rindex('(')
                        base_mod = orig_mod[:paren_start].strip()
                        if base_mod == saved_value or orig_mod == saved_value:
                            display_value = processed_mods[j]
                            break
            
            mod_var.set(display_value)
            
            # Exclude this dropdown's current selection
            if display_value not in [f"Select {col_name} Mod"]:
                available_mods = [m for m in processed_mods if m != display_value]
            else:
                available_mods = processed_mods
            
            mod_dropdown = ctk.CTkComboBox(
                frame,
                values=available_mods,
                variable=mod_var,
                command=lambda choice, sname=skill_name, sdata=skill_data: self.update_skill_mod_selection(sname, sdata),
                state="readonly"
            )
            mod_dropdown.grid(row=row_idx, column=0, padx=5, pady=5, sticky="ew")
            skill_data['variables'][key] = mod_var
            skill_data['dropdowns'][key] = mod_dropdown
            
            # Store original mods for reference
            if 'original_mods' not in skill_data:
                skill_data['original_mods'] = {}
            skill_data['original_mods'][key] = original_mods
            
            row_idx += 1
    
    def update_skill_mod_selection(self, skill_name, skill_data):
        """Handle skill mod selection and validate specialization requirements"""
        # Check if any selected mod is locked due to specialization
        for key, mod_var in skill_data['variables'].items():
            if key.startswith('mod_'):
                mod_val = mod_var.get()
                
                # Check if this is a disabled specialization-locked mod
                if "(Select " in mod_val:
                    # Revert to previous selection
                    prev_mod = skill_data.get(f'prev_{key}', f"Select Mod")
                    mod_var.set(prev_mod)
                    return
                
                # Store the current selection for potential revert
                skill_data[f'prev_{key}'] = mod_val
        
        # Refresh mods to update exclusion lists
        self.update_skill_mods(skill_name, skill_data, preserve_selections=True)

    
    def create_weapon_section_in_container(self, parent, weapon_name, icon_file):
        """Create a weapon section with all dropdown fields in a container"""
        # Main frame for weapon with border
        weapon_frame = ctk.CTkFrame(parent, border_width=2)
        weapon_frame.pack(padx=10, pady=10, fill="both", expand=True)
        weapon_frame.grid_columnconfigure(1, weight=1)
        weapon_frame.grid_rowconfigure(0, weight=1)
        
        # Try to load and display icon - centered vertically
        try:
            from PIL import Image
            icon_path = assets_directory / icon_file
            if icon_path.exists():
                icon_image = ctk.CTkImage(
                    light_image=Image.open(icon_path),
                    dark_image=Image.open(icon_path),
                    size=(40, 40)
                )
                icon_label = ctk.CTkLabel(weapon_frame, image=icon_image, text="")
                # Empty sticky means center both horizontally and vertically
                icon_label.grid(row=0, column=0, padx=10, pady=10, sticky="")
        except:
            pass
        
        # Create inner grid frame for dropdowns
        dropdown_frame = ctk.CTkFrame(weapon_frame)
        dropdown_frame.grid(row=0, column=1, sticky="ew", padx=10, pady=10)
        dropdown_frame.grid_columnconfigure(0, weight=1)
        
        # Store references to all dropdowns for this weapon
        weapon_data = {
            'name': weapon_name,
            'frame': dropdown_frame,
            'dropdowns': {},
            'variables': {}
        }
        
        # Class dropdown - always shown
        class_var = ctk.StringVar(value="Select Class")
        class_values = ["Assault Rifles", "Light Machineguns", "Marksman Rifles", "Pistols", "Rifles", "Shotguns", "Sub Machine Guns", "Signature Weapons"]
        class_dropdown = ctk.CTkComboBox(
            dropdown_frame,
            values=class_values,
            variable=class_var,
            command=lambda choice: self.update_weapon(weapon_name, weapon_data),
            state="readonly"
        )
        class_dropdown.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        
        weapon_data['variables']['class'] = class_var
        weapon_data['dropdowns']['class'] = class_dropdown
        
        # Store weapon_data for later access
        if not hasattr(self, 'weapon_data'):
            self.weapon_data = weapon_data
        
        # Initial check for Signature Weapons greying
        self.update_weapon_class_options()
    
    def update_weapon_class_options(self):
        """Update weapon class dropdown based on specialization selection"""
        if not hasattr(self, 'weapon_data') or not hasattr(self, 'specialization_var'):
            return
        
        spec_val = self.specialization_var.get()
        current_class = self.weapon_data['variables']['class'].get()
        class_dropdown = self.weapon_data['dropdowns']['class']
        
        all_classes = ["Assault Rifles", "Light Machineguns", "Marksman Rifles", "Pistols", "Rifles", "Shotguns", "Sub Machine Guns"]
        
        # Add Signature Weapons with strikethrough if no specialization selected
        if spec_val == "Select Specialization":
            # Use strikethrough to indicate disabled state
            signature_display = "Signature Weapons (Select Specialization)"
        else:
            signature_display = "Signature Weapons"
        
        all_classes.append(signature_display)
        
        # Exclude current selection (but check both versions of Signature Weapons)
        if current_class == "Signature Weapons" or current_class.startswith("Signature Weapons"):
            available_classes = [c for c in all_classes if not c.startswith("Signature Weapons")]
        else:
            available_classes = [c for c in all_classes if c != current_class]
        
        class_dropdown.configure(values=available_classes)
    
    def update_weapon(self, weapon_name, weapon_data):
        """Dynamically update weapon dropdowns based on selections"""
        class_val = weapon_data['variables']['class'].get()
        frame = weapon_data['frame']
        
        if class_val == "Select Class":
            return
        
        # Validate Signature Weapons selection - prevent if no specialization
        if class_val.startswith("Signature Weapons"):
            spec_val = self.specialization_var.get()
            if spec_val == "Select Specialization":
                # Revert to previous class
                prev_class = weapon_data.get('prev_class', "Select Class")
                weapon_data['variables']['class'].set(prev_class)
                return
            else:
                # Normalize the class value to just "Signature Weapons"
                class_val = "Signature Weapons"
                weapon_data['variables']['class'].set(class_val)
        
        # Update class dropdown options
        self.update_weapon_class_options()
        
        # Check if class changed - if so, clear name and subsequent dropdowns but preserve type
        # Store previous class to detect changes
        prev_class = weapon_data.get('prev_class', None)
        if prev_class is not None and prev_class != class_val:
            # Class changed - clear name and everything below it
            keys_to_clear = ['name', 'core_1', 'core_2', 'attribute', 'mod_1', 'mod_2', 'mod_3', 'mod_4', 'talent_1', 'talent_2']
            for key in keys_to_clear:
                if key in weapon_data['dropdowns']:
                    weapon_data['dropdowns'][key].grid_forget()
                    del weapon_data['dropdowns'][key]
                if key in weapon_data['variables']:
                    del weapon_data['variables'][key]
        weapon_data['prev_class'] = class_val
        
        # Clear type dropdown if switching to/from Signature Weapons
        if class_val == "Signature Weapons":
            if 'type' in weapon_data['dropdowns']:
                weapon_data['dropdowns']['type'].grid_forget()
                del weapon_data['dropdowns']['type']
                del weapon_data['variables']['type']
        elif 'type' not in weapon_data['dropdowns']:
            # Not Signature Weapons and no type dropdown exists yet - create it
            row_idx = 1
            type_var = ctk.StringVar(value="Select Type")
            type_values = ["High-End", "Named", "Exotic"]
            type_dropdown = ctk.CTkComboBox(
                frame,
                values=type_values,
                variable=type_var,
                command=lambda choice: self.update_weapon_name(weapon_name, weapon_data),
                state="readonly"
            )
            type_dropdown.grid(row=row_idx, column=0, padx=5, pady=5, sticky="ew")
            weapon_data['variables']['type'] = type_var
            weapon_data['dropdowns']['type'] = type_dropdown
            return  # Wait for type selection
        
        # If we have type dropdown, wait for selection before proceeding
        if 'type' in weapon_data['variables']:
            type_val = weapon_data['variables']['type'].get()
            if type_val == "Select Type":
                return
        
        # Create name dropdown
        if class_val == "Signature Weapons":
            self.update_weapon_name(weapon_name, weapon_data)
        elif 'type' in weapon_data['variables']:
            # Only create name if type is selected
            self.update_weapon_name(weapon_name, weapon_data)
    
    def update_weapon_name(self, weapon_name, weapon_data):
        """Update weapon name dropdown based on class and type"""
        class_val = weapon_data['variables']['class'].get()
        frame = weapon_data['frame']
        
        # Clear existing name dropdown and everything below it (core attributes, attribute, mods, talents)
        keys_to_clear = ['name', 'core_1', 'core_2', 'attribute', 'mod_1', 'mod_2', 'mod_3', 'mod_4', 'talent_1', 'talent_2']
        for key in keys_to_clear:
            if key in weapon_data['dropdowns']:
                weapon_data['dropdowns'][key].grid_forget()
                del weapon_data['dropdowns'][key]
            if key in weapon_data['variables']:
                del weapon_data['variables'][key]
        
        row_idx = 2 if 'type' in weapon_data['dropdowns'] else 1
        
        name_var = ctk.StringVar()
        
        if class_val == "Signature Weapons":
            # Get name from Specialization Weapons sheet based on specialization
            spec_val = self.specialization_var.get()
            
            sig_weapons = data["Weapon"]["Signature Weapons"]
            weapon_row = sig_weapons[sig_weapons["Specialization"] == spec_val]
            if not weapon_row.empty:
                weapon_name_val = str(weapon_row.iloc[0]["Name"])
                name_var.set(weapon_name_val)
                name_dropdown = ctk.CTkComboBox(frame, values=[weapon_name_val], variable=name_var, state="disabled")
                
                # Store the name dropdown first
                name_dropdown.grid(row=row_idx, column=0, padx=5, pady=5, sticky="ew")
                weapon_data['variables']['name'] = name_var
                weapon_data['dropdowns']['name'] = name_dropdown
                
                # Now create the talent dropdown
                self.update_weapon_attributes(weapon_name, weapon_data)
                return
            else:
                return
        else:
            # Get names from the class sheet filtered by type
            type_val = weapon_data['variables']['type'].get()
            if type_val == "Select Type":
                return
            
            # Update type dropdown to exclude current selection
            all_types = ["High-End", "Named", "Exotic"]
            available_types = [t for t in all_types if t != type_val]
            weapon_data['dropdowns']['type'].configure(values=available_types)
            
            # Get weapon names and convert to strings
            weapon_sheet = data["Weapon"][class_val]
            all_names = [str(name) for name in weapon_sheet[weapon_sheet["Type"] == type_val]["Name"].tolist()]
            
            # Process names to handle specialization requirements
            spec_val = self.specialization_var.get() if hasattr(self, 'specialization_var') else "Select Specialization"
            processed_names = []
            original_weapon_names = []
            
            for name in all_names:
                original_weapon_names.append(name)
                # Check if name has specialization requirement (rightmost parentheses)
                if '(' in name and ')' in name:
                    # Get the last occurrence of parentheses
                    paren_start = name.rindex('(')
                    paren_end = name.rindex(')')
                    base_name = name[:paren_start].strip()
                    paren_content = name[paren_start+1:paren_end].strip()
                    
                    # Check if rightmost parentheses contain a specialization
                    if self.is_specialization_name(paren_content):
                        # This is a specialization-specific weapon
                        if spec_val == paren_content:
                            # User has correct spec, show without the specialization parentheses
                            processed_names.append(base_name)
                        else:
                            # User doesn't have the spec, show with "Select X" indicator
                            processed_names.append(f"{base_name} (Select {paren_content})")
                    else:
                        # Not a specialization, keep the full weapon name
                        processed_names.append(name)
                else:
                    # No parentheses, keep as is
                    processed_names.append(name)
            
            # Store original names for matching later
            weapon_data['original_weapon_names'] = original_weapon_names
            
            name_var.set("Select Name")
            name_dropdown = ctk.CTkComboBox(
                frame,
                values=processed_names,
                variable=name_var,
                command=lambda choice: self.update_weapon_name_selection(weapon_name, weapon_data),
                state="readonly"
            )
        
        name_dropdown.grid(row=row_idx, column=0, padx=5, pady=5, sticky="ew")
        weapon_data['variables']['name'] = name_var
        weapon_data['dropdowns']['name'] = name_dropdown
    
    def update_weapon_name_selection(self, weapon_name, weapon_data):
        """Handle weapon name selection to exclude current selection from list"""
        name_val = weapon_data['variables']['name'].get()
        class_val = weapon_data['variables']['class'].get()
        type_val = weapon_data['variables']['type'].get()
        
        if name_val == "Select Name":
            return
        
        # Check if this is a disabled specialization-locked weapon
        if "(Select " in name_val:
            # Revert to previous selection
            prev_name = weapon_data.get('prev_name', "Select Name")
            weapon_data['variables']['name'].set(prev_name)
            return
        
        # Store the current selection for potential revert
        weapon_data['prev_name'] = name_val
        
        # For Signature Weapons, just create the talent dropdown
        if class_val == "Signature Weapons":
            self.update_weapon_attributes(weapon_name, weapon_data)
            return
        
        # Match the selected name against original names
        original_weapon_names = weapon_data.get('original_weapon_names', [])
        matched_name = name_val
        
        # Try to find matching original name
        for orig_name in original_weapon_names:
            if orig_name == name_val:
                matched_name = orig_name
                break
            elif '(' in orig_name and ')' in orig_name:
                # Check if this is a specialization-specific weapon
                paren_start = orig_name.rindex('(')
                paren_end = orig_name.rindex(')')
                base_name = orig_name[:paren_start].strip()
                paren_content = orig_name[paren_start+1:paren_end].strip()
                
                if self.is_specialization_name(paren_content) and base_name == name_val:
                    matched_name = orig_name
                    break
        
        # Update weapon_data to use the matched name for lookups
        weapon_data['matched_weapon_name'] = matched_name
        
        # Get all names for this class and type, converted to strings
        weapon_sheet = data["Weapon"][class_val]
        all_names = [str(name) for name in weapon_sheet[weapon_sheet["Type"] == type_val]["Name"].tolist()]
        
        # Process names for display
        spec_val = self.specialization_var.get() if hasattr(self, 'specialization_var') else "Select Specialization"
        processed_names = []
        
        for name in all_names:
            if '(' in name and ')' in name:
                paren_start = name.rindex('(')
                paren_end = name.rindex(')')
                base_name = name[:paren_start].strip()
                paren_content = name[paren_start+1:paren_end].strip()
                
                if self.is_specialization_name(paren_content):
                    if spec_val == paren_content:
                        processed_names.append(base_name)
                    else:
                        processed_names.append(f"{base_name} (Select {paren_content})")
                else:
                    processed_names.append(name)
            else:
                processed_names.append(name)
        
        # Exclude current selection from processed names
        available_names = [n for n in processed_names if n != name_val]
        
        # Update dropdown
        weapon_data['dropdowns']['name'].configure(values=available_names)
        
        # Create core attributes and attribute dropdowns
        self.update_weapon_attributes(weapon_name, weapon_data)
    
    def update_weapon_mod_selection(self, weapon_name, weapon_data):
        """Handle weapon mod selection and validate specialization requirements"""
        # Check if any selected mod is locked due to specialization
        for key, mod_var in weapon_data['variables'].items():
            if key.startswith('mod_'):
                mod_val = mod_var.get()
                
                # Check if this is a disabled specialization-locked mod
                if "(Select " in mod_val:
                    # Determine the correct placeholder based on mod type
                    mod_index = int(key.split('_')[1])
                    mod_types = ["Optics Mod", "Magazine Mod", "Underbarrel Mod", "Muzzle Mod"]
                    if mod_index <= len(mod_types):
                        column_name = mod_types[mod_index - 1]
                        prev_mod = weapon_data.get(f'prev_{key}', f"Select {column_name}")
                    else:
                        prev_mod = weapon_data.get(f'prev_{key}', "Select Mod")
                    
                    # Revert to previous selection
                    mod_var.set(prev_mod)
                    return
                
                # Store the current selection for potential revert
                weapon_data[f'prev_{key}'] = mod_val
        
        # Refresh weapon attributes to update exclusion lists
        self.update_weapon_attributes(weapon_name, weapon_data, preserve_selections=True)
    
    def update_weapon_attributes(self, weapon_name, weapon_data, preserve_selections=True):
        """Update weapon core attributes and attribute based on name selection"""
        class_val = weapon_data['variables']['class'].get()
        name_val = weapon_data['variables']['name'].get()
        frame = weapon_data['frame']
        
        if name_val == "Select Name":
            return
        
        # Preserve existing selections if requested
        saved_selections = {}
        if preserve_selections:
            for key, var in weapon_data['variables'].items():
                if key not in ['class', 'type', 'name']:
                    saved_selections[key] = var.get()
        
        # Clear existing attribute dropdowns
        for key in list(weapon_data['dropdowns'].keys()):
            if key not in ['class', 'type', 'name']:
                weapon_data['dropdowns'][key].grid_forget()
                del weapon_data['dropdowns'][key]
        for key in list(weapon_data['variables'].keys()):
            if key not in ['class', 'type', 'name']:
                del weapon_data['variables'][key]
        
        row_idx = 3  # After class, type, name
        
        # Handle Signature Weapons separately (they only have a talent)
        if class_val == "Signature Weapons":
            sig_weapons = data["Weapon"]["Signature Weapons"]
            sig_row = sig_weapons[sig_weapons["Name"] == name_val]
            if not sig_row.empty:
                talent_val = sig_row.iloc[0].get("Talent", pd.NA)
                if not pd.isna(talent_val):
                    talent_var = ctk.StringVar(value=talent_val)
                    talent_dropdown = ctk.CTkComboBox(frame, values=[talent_val], variable=talent_var, state="disabled")
                    talent_dropdown.grid(row=row_idx, column=0, padx=5, pady=5, sticky="ew")
                    weapon_data['variables']['talent_1'] = talent_var
                    weapon_data['dropdowns']['talent_1'] = talent_dropdown
            return
        
        # Get the actual weapon name to use for lookups (may differ from displayed name)
        lookup_name = weapon_data.get('matched_weapon_name', name_val)
        
        # Get weapon row for normal weapons
        weapon_sheet = data["Weapon"][class_val]
        type_val = weapon_data['variables']['type'].get()
        weapon_row = weapon_sheet[(weapon_sheet["Name"] == lookup_name) & (weapon_sheet["Type"] == type_val)].iloc[0]
        
        # Core Attributes (2 slots) - always predetermined or NA
        for i in range(1, 3):
            core_val = weapon_row.get(f"Core Attribute {i}", pd.NA)
            if pd.isna(core_val):
                continue
            
            core_var = ctk.StringVar(value=core_val)
            core_dropdown = ctk.CTkComboBox(frame, values=[core_val], variable=core_var, state="disabled")
            core_dropdown.grid(row=row_idx, column=0, padx=5, pady=5, sticky="ew")
            weapon_data['variables'][f'core_{i}'] = core_var
            weapon_data['dropdowns'][f'core_{i}'] = core_dropdown
            row_idx += 1
        
        # Attribute (1 slot) - more complex logic
        attr_cell = weapon_row.get("Attribute", pd.NA)
        if not pd.isna(attr_cell):
            saved_value = saved_selections.get('attribute', "Select Attribute")
            attr_var = ctk.StringVar(value=saved_value)
            
            if attr_cell == "*":
                # User can select any attribute
                all_attrs = data["Weapon"]["Attributes"]["Stats"].tolist()
                # Exclude current selection
                if saved_value not in ["Select Attribute"]:
                    available_attrs = [a for a in all_attrs if a != saved_value]
                else:
                    available_attrs = all_attrs
                
                attr_dropdown = ctk.CTkComboBox(
                    frame,
                    values=available_attrs,
                    variable=attr_var,
                    command=lambda choice, wname=weapon_name, wdata=weapon_data: self.update_weapon_attributes(wname, wdata, preserve_selections=True),
                    state="readonly"
                )
            elif isinstance(attr_cell, str) and attr_cell.startswith("!"):
                # Cannot select this specific one
                excluded = attr_cell[1:]
                all_attrs = data["Weapon"]["Attributes"]["Stats"].tolist()
                # Exclude the specified value and current selection
                if saved_value not in ["Select Attribute"]:
                    available_attrs = [a for a in all_attrs if a != excluded and a != saved_value]
                else:
                    available_attrs = [a for a in all_attrs if a != excluded]
                
                attr_dropdown = ctk.CTkComboBox(
                    frame,
                    values=available_attrs,
                    variable=attr_var,
                    command=lambda choice, wname=weapon_name, wdata=weapon_data: self.update_weapon_attributes(wname, wdata, preserve_selections=True),
                    state="readonly"
                )
            else:
                # Fixed attribute
                attr_var.set(attr_cell)
                attr_dropdown = ctk.CTkComboBox(frame, values=[attr_cell], variable=attr_var, state="disabled")
            
            attr_dropdown.grid(row=row_idx, column=0, padx=5, pady=5, sticky="ew")
            weapon_data['variables']['attribute'] = attr_var
            weapon_data['dropdowns']['attribute'] = attr_dropdown
            row_idx += 1
        
        # Mods (4 slots) - Optics, Magazine, Underbarrel, Muzzle
        mod_types = [
            ("Optics Mod", "Optics Mods"),
            ("Magazine Mod", "Magazine Mods"),
            ("Underbarrel Mod", "Underbarrel Mods"),
            ("Muzzle Mod", "Muzzle Mods")
        ]
        
        for i, (column_name, sheet_name) in enumerate(mod_types, start=1):
            mod_cell = weapon_row.get(column_name, pd.NA)
            if pd.isna(mod_cell):
                continue
            
            key = f'mod_{i}'
            saved_value = saved_selections.get(key, f"Select {column_name}")
            mod_var = ctk.StringVar(value=saved_value)
            
            if isinstance(mod_cell, str) and mod_cell.startswith("*"):
                # User-selectable mod - get mod rail type
                mod_rail = mod_cell[1:]  # Remove '*' prefix (e.g., "*Long Optics Rail" -> "Long Optics Rail")
                
                # Get mods sheet (not weapon-class specific)
                if sheet_name in data["Weapon"]:
                    mods_df = data["Weapon"][sheet_name]
                    
                    # Get available mods for this rail type
                    if mod_rail in mods_df.columns:
                        all_mods = mods_df[mods_df[mod_rail] == '✓']['Stats'].tolist()
                        
                        # Process mods to handle specialization requirements
                        spec_val = self.specialization_var.get() if hasattr(self, 'specialization_var') else "Select Specialization"
                        processed_mods = []
                        original_mods = []  # Store original mod names for matching
                        
                        for mod in all_mods:
                            original_mods.append(mod)
                            # Check if mod has specialization requirement
                            if '(' in mod and ')' in mod:
                                paren_start = mod.rindex('(')  # Use rindex to get the last occurrence
                                paren_end = mod.rindex(')')
                                base_mod = mod[:paren_start].strip()
                                paren_content = mod[paren_start+1:paren_end].strip()
                                
                                # Check if parentheses contain a specialization
                                if self.is_specialization_name(paren_content):
                                    # This is a specialization-specific mod
                                    if spec_val == paren_content:
                                        # User has correct spec, show just the base name
                                        processed_mods.append(base_mod)
                                    else:
                                        # User doesn't have the spec, show with "Select X" indicator
                                        processed_mods.append(f"{base_mod} (Select {paren_content})")
                                else:
                                    # Not a specialization, keep the full mod name
                                    processed_mods.append(mod)
                            else:
                                # No parentheses, keep as is
                                processed_mods.append(mod)
                        
                        # Match saved value against processed mods
                        display_value = saved_value
                        if saved_value not in [f"Select {column_name}"]:
                            # Try to find the processed version of the saved value
                            for j, orig_mod in enumerate(original_mods):
                                if '(' in orig_mod and ')' in orig_mod:
                                    paren_start = orig_mod.rindex('(')
                                    base_mod = orig_mod[:paren_start].strip()
                                    if base_mod == saved_value or orig_mod == saved_value:
                                        display_value = processed_mods[j]
                                        break
                        
                        mod_var.set(display_value)
                        
                        # Exclude current selection
                        if display_value not in [f"Select {column_name}"]:
                            available_mods = [m for m in processed_mods if m != display_value]
                        else:
                            available_mods = processed_mods
                        
                        mod_dropdown = ctk.CTkComboBox(
                            frame,
                            values=available_mods,
                            variable=mod_var,
                            command=lambda choice, wname=weapon_name, wdata=weapon_data: self.update_weapon_mod_selection(wname, wdata),
                            state="readonly"
                        )
                        
                        # Store original mods for reference
                        if 'original_weapon_mods' not in weapon_data:
                            weapon_data['original_weapon_mods'] = {}
                        weapon_data['original_weapon_mods'][key] = original_mods
                    else:
                        continue
                else:
                    continue
            else:
                # Fixed mod
                mod_var.set(mod_cell)
                mod_dropdown = ctk.CTkComboBox(frame, values=[mod_cell], variable=mod_var, state="disabled")
            
            mod_dropdown.grid(row=row_idx, column=0, padx=5, pady=5, sticky="ew")
            weapon_data['variables'][key] = mod_var
            weapon_data['dropdowns'][key] = mod_dropdown
            row_idx += 1
        
        # Talents (up to 2 slots for normal weapons)
        for i in range(1, 3):
            talent_cell = weapon_row.get(f"Talent {i}", pd.NA)
            if pd.isna(talent_cell):
                continue
            
            key = f'talent_{i}'
            saved_value = saved_selections.get(key, "Select Talent")
            talent_var = ctk.StringVar(value=saved_value)
            
            if talent_cell == "*":
                # User-selectable talent
                talents_df = data["Weapon"]["Talents"]
                
                # Get talents available for this weapon class
                if class_val in talents_df.columns:
                    all_talents = talents_df[talents_df[class_val] == '✓']['Name'].tolist()
                    
                    # Exclude current selection
                    if saved_value not in ["Select Talent"]:
                        available_talents = [t for t in all_talents if t != saved_value]
                    else:
                        available_talents = all_talents
                    
                    talent_dropdown = ctk.CTkComboBox(
                        frame,
                        values=available_talents,
                        variable=talent_var,
                        command=lambda choice, wname=weapon_name, wdata=weapon_data: self.update_weapon_attributes(wname, wdata, preserve_selections=True),
                        state="readonly"
                    )
                else:
                    continue
            else:
                # Fixed talent
                talent_var.set(talent_cell)
                talent_dropdown = ctk.CTkComboBox(frame, values=[talent_cell], variable=talent_var, state="disabled")
            
            talent_dropdown.grid(row=row_idx, column=0, padx=5, pady=5, sticky="ew")
            weapon_data['variables'][key] = talent_var
            weapon_data['dropdowns'][key] = talent_dropdown
            row_idx += 1

    
    def setup_build_tuning(self):
        """Setup the Build Tuning window"""
        label = ctk.CTkLabel(
            self.build_tuning_frame,
            text="Build Tuning",
            font=ctk.CTkFont(size=32, weight="bold")
        )
        label.pack(pady=50)
    
    def setup_damage_output(self):
        """Setup the Damage Output window"""
        label = ctk.CTkLabel(
            self.damage_output_frame,
            text="Damage Output",
            font=ctk.CTkFont(size=32, weight="bold")
        )
        label.pack(pady=50)
    
    def show_build_creator(self):
        """Show the Build Creator window"""
        if self.current_frame:
            self.current_frame.grid_forget()
        self.build_creator_frame.grid(row=0, column=0, sticky="nsew")
        self.current_frame = self.build_creator_frame
        
        # Update button states
        self.build_creator_btn.configure(state="disabled")
        self.build_tuning_btn.configure(state="normal")
        self.damage_output_btn.configure(state="normal")
    
    def show_build_tuning(self):
        """Show the Build Tuning window"""
        if self.current_frame:
            self.current_frame.grid_forget()
        self.build_tuning_frame.grid(row=0, column=0, sticky="nsew")
        self.current_frame = self.build_tuning_frame
        
        # Update button states
        self.build_creator_btn.configure(state="normal")
        self.build_tuning_btn.configure(state="disabled")
        self.damage_output_btn.configure(state="normal")
    
    def show_damage_output(self):
        """Show the Damage Output window"""
        if self.current_frame:
            self.current_frame.grid_forget()
        self.damage_output_frame.grid(row=0, column=0, sticky="nsew")
        self.current_frame = self.damage_output_frame
        
        # Update button states
        self.build_creator_btn.configure(state="normal")
        self.build_tuning_btn.configure(state="normal")
        self.damage_output_btn.configure(state="disabled")


#################### SECTION BREAK ####################

##### MAIN EXECUTION #####

if __name__ == "__main__":
    app = DamageCalculatorApp()
    app.mainloop()

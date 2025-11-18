"""


"""

__author__ = "iRNHO"
__contact__ = "Message 'iRNHO' on XBOX or 'irnho' on discord regarding any questions, feedback, bug-reporting etc."
__discord__ = "https://discord.gg/--------"
__title__ = "iRNHO's Damage Calculator"


#################### SECTION BREAK ####################

##### IMPORT STATEMENTS #####

import json

import customtkinter as ctk
import pandas as pd

from pathlib import Path
from PIL import Image
from platformdirs import user_data_dir


#################### SECTION BREAK ####################

##### DATA LOADING #####

root_directory = Path(user_data_dir(__title__, __author__))
root_directory = Path.cwd() # Remove this line for production use
assets_directory = root_directory / "assets"
builds_directory = root_directory / "builds"

builds_directory.mkdir(parents=True, exist_ok=True)

spreadsheet = pd.read_excel(root_directory / "iRNHO'S Spreadsheet.xlsx", sheet_name=None)
spreadsheet["Rifles"]["Name"] = [str(name) for name in spreadsheet["Rifles"]["Name"]]


#################### SECTION BREAK ####################

##### CONSTANTS #####

SPECIALIZATIONS = spreadsheet["Specializations"]["Name"].tolist()

WEAPON_VARIABLES = ["Class", "Type", "Name", "Core Attribute 1", "Core Attribute 2", "Attribute", "Optics Mod", "Magazine Mod", "Underbarrel Mod", "Muzzle Mod", "Talent 1", "Talent 2", "Expertise"]
WEAPON_VARIABLES_NO_EXPERTISE = [variable for variable in WEAPON_VARIABLES if variable != "Expertise"]
WEAPON_CLASSES = ["Assault Rifles", "Light Machineguns", "Marksman Rifles", "Pistols", "Rifles", "Shotguns", "Sub Machine Guns", "Signature Weapons"]
WEAPON_CLASSES_NO_SIGNATURE = [weapon_class for weapon_class in WEAPON_CLASSES if weapon_class != "Signature Weapons"]
WEAPON_TYPES = ["High-End", "Named", "Exotic"]
WEAPON_MODS = ["Optics Mod", "Magazine Mod", "Underbarrel Mod", "Muzzle Mod"]

GEAR_VARIABLES = ["Type", "Name", "Core Attribute 1", "Core Attribute 2", "Core Attribute 3", "Attribute 1", "Attribute 2", "Attribute 3", "Mod 1", "Mod 2", "Talent 1", "Talent 2"]
GEAR_SLOTS = ["Mask", "Body Armor", "Holster", "Backpack", "Gloves", "Kneepads"]
GEAR_TYPES = ["Improvised", "Brand Set", "Gear Set", "Named", "Exotic"]
GEAR_TYPES_NO_EXOTIC = [gear_type for gear_type in GEAR_TYPES if gear_type != "Exotic"]

SKILL_VARIABLES = ["Class", "Name", "Mod 1", "Mod 2", "Mod 3", "Expertise"]
SKILL_VARIABLES_NO_EXPERTISE = [variable for variable in SKILL_VARIABLES if variable != "Expertise"]
SKILL_SLOTS = ["Left", "Right"]
SKILL_CLASSES = spreadsheet["Skills"]["Class"].tolist()


#################### SECTION BREAK ####################

##### GUI CLASS #####

class DamageCalculatorApp(ctk.CTk):

    def __init__(self):
        """
        Initializes the application window and its components.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.

        """

        super().__init__()
        
        self.title(__title__)
        self.iconbitmap(assets_directory / "sub_machine_gun.ico")
        self.geometry("1000x700")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        
        ctk.set_default_color_theme(assets_directory / "custom_theme.json")
        ctk.set_appearance_mode("dark")        
        
        self.navigation_bar = ctk.CTkFrame(self, height=60, corner_radius=0)
        self.navigation_bar.grid(row=0, column=0, sticky="ew", padx=0, pady=0)
        self.navigation_bar.grid_columnconfigure(1, weight=1)
        
        self.burger_button = ctk.CTkButton(self.navigation_bar, text="☰", command=self.toggle_sidebar, width=40, height=40, font=ctk.CTkFont(size=20), fg_color="transparent")
        self.burger_button.grid(row=0, column=0, padx=10, pady=10, sticky="w")
        
        button_container = ctk.CTkFrame(self.navigation_bar, fg_color="transparent")
        button_container.grid(row=0, column=1, sticky="")

        self.build_creator_button = ctk.CTkButton(button_container, text="Build Creator", command=self.show_build_creator, width=200, height=50, font=ctk.CTkFont(size=16, weight="bold"))
        self.build_creator_button.pack(side="left", padx=5)

        self.build_tuning_button = ctk.CTkButton(button_container, text="Build Tuning", command=self.show_build_tuning, width=200, height=50, font=ctk.CTkFont(size=16, weight="bold"))
        self.build_tuning_button.pack(side="left", padx=5)

        self.damage_output_button = ctk.CTkButton(button_container, text="Damage Output", command=self.show_damage_output, width=200, height=50, font=ctk.CTkFont(size=16, weight="bold"))
        self.damage_output_button.pack(side="left", padx=5)

        self.main_container = ctk.CTkFrame(self)
        self.main_container.grid(row=1, column=0, sticky="nsew")
        self.main_container.grid_rowconfigure(0, weight=1)
        self.main_container.grid_columnconfigure(0, weight=1)
        
        self.sidebar_container = None
        
        self.build_creator_container = ctk.CTkFrame(self.main_container)
        self.build_tuning_container = ctk.CTkFrame(self.main_container)
        self.damage_output_container = ctk.CTkFrame(self.main_container)
        
        self.setup_build_creator()
        self.setup_build_tuning()
        self.setup_damage_output()
        
        self.build_creator_container.grid(row=0, column=0, sticky="nsew")
        self.active_tab_container = self.build_creator_container

    
    ####################

    def toggle_sidebar(self):
        """
        Toggles the sidebar.
        
        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
        
        """
        
        if self.sidebar_container:
            self.sidebar_container.destroy()
            self.sidebar_container = None
            self.unbind("<Escape>")

        else:
            self.sidebar_container = ctk.CTkFrame(self, width=550, corner_radius=0)
            self.sidebar_container.place(x=0, y=0, relheight=1.0)
            self.sidebar_container.pack_propagate(False)
            
            close_button = ctk.CTkButton(self.sidebar_container, text="✕", command=self.toggle_sidebar, width=30, height=30, font=ctk.CTkFont(size=16), fg_color="transparent")
            close_button.place(x=510, y=10)
            
            scrollable_container = ctk.CTkScrollableFrame(self.sidebar_container)
            scrollable_container.pack(fill="both", expand=True, padx=10, pady=(50, 10))

            build_dictionaries = []
        
            for build_file in builds_directory.iterdir():
                with open(build_file, "r") as file:
                    build_dictionaries.append(json.load(file))

            for build_dictionary in sorted(build_dictionaries, key=lambda build_dictionary: build_dictionary["Build Name"].lower()):
                build_container = ctk.CTkFrame(scrollable_container)
                build_container.pack(fill="x", padx=5, pady=5)
                
                build_label = ctk.CTkLabel(build_container, text=build_dictionary["Build Name"], font=ctk.CTkFont(size=14, weight="bold"), anchor="w")
                build_label.pack(side="left", padx=10, pady=10, fill="x", expand=True)
                
                button_container = ctk.CTkFrame(build_container, fg_color="transparent")
                button_container.pack(side="right", padx=5, pady=5)

                load_icon_path = assets_directory / "load.png"
                load_button = ctk.CTkButton(button_container, text="", image=ctk.CTkImage(light_image=Image.open(load_icon_path), dark_image=Image.open(load_icon_path), size=(30, 30)), command=lambda command=self.load_build: command(build_dictionary), width=40, height=40, fg_color="transparent", border_width=2, border_color="gray30")
                load_button.pack(side="left", padx=2)

                rename_icon_path = assets_directory / "rename.png"
                rename_button = ctk.CTkButton(button_container, text="", image=ctk.CTkImage(light_image=Image.open(rename_icon_path), dark_image=Image.open(rename_icon_path), size=(30, 30)), command=lambda command=self.rename_build: command(build_dictionary), width=40, height=40, fg_color="transparent", border_width=2, border_color="gray30")
                rename_button.pack(side="left", padx=2)

                overwrite_icon_path = assets_directory / "overwrite.png"
                overwrite_button = ctk.CTkButton(button_container, text="", image=ctk.CTkImage(light_image=Image.open(overwrite_icon_path), dark_image=Image.open(overwrite_icon_path), size=(30, 30)), command=lambda command=self.overwrite_build: command(build_dictionary), width=40, height=40, fg_color="transparent", border_width=2, border_color="gray30")
                overwrite_button.pack(side="left", padx=2)

                delete_icon_path = assets_directory / "delete.png"
                delete_button = ctk.CTkButton(button_container, text="", image=ctk.CTkImage(light_image=Image.open(delete_icon_path), dark_image=Image.open(delete_icon_path), size=(30, 30)), command=lambda command=self.delete_build: command(build_dictionary), width=40, height=40, fg_color="transparent", border_width=2, border_color="gray30")
                delete_button.pack(side="left", padx=2)
            
            new_build_container = ctk.CTkFrame(scrollable_container, height=80)
            new_build_container.pack(fill="x", padx=5, pady=10)
            
            new_build_button = ctk.CTkButton(new_build_container, text="+ Save Build", command=self.create_new_build, height=60, font=ctk.CTkFont(size=16, weight="bold"))
            new_build_button.pack(fill="both", expand=True, padx=10, pady=10)
            
            self.bind("<Escape>", lambda event: self.toggle_sidebar())


    def show_build_creator(self):
        """
        Displays the 'Build Creator' tab.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.

        """

        self.active_tab_container.grid_forget()

        self.build_creator_container.grid(row=0, column=0, sticky="nsew")
        self.active_tab_container = self.build_creator_container
        
        self.build_creator_button.configure(state="disabled")
        self.build_tuning_button.configure(state="normal")
        self.damage_output_button.configure(state="normal")
    
    
    def show_build_tuning(self):
        """
        Displays the 'Build Tuning' tab.
        
        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            
        """

        self.active_tab_container.grid_forget()

        self.build_tuning_container.grid(row=0, column=0, sticky="nsew")
        self.active_tab_container = self.build_tuning_container
        
        self.build_creator_button.configure(state="normal")
        self.build_tuning_button.configure(state="disabled")
        self.damage_output_button.configure(state="normal")
    

    def show_damage_output(self):
        """
        Displays the 'Damage Output' tab.
        
        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.

        """

        self.active_tab_container.grid_forget()

        self.damage_output_container.grid(row=0, column=0, sticky="nsew")
        self.active_tab_container = self.damage_output_container
        
        self.build_creator_button.configure(state="normal")
        self.build_tuning_button.configure(state="normal")
        self.damage_output_button.configure(state="disabled")


    def setup_build_creator(self):
        """
        Populates the 'Build Creator' tab with its UI components.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
        
        """

        self.build_creator_container.grid_columnconfigure(0, weight=1)
        self.build_creator_container.grid_rowconfigure(0, weight=1)
        
        scrollable_container = ctk.CTkScrollableFrame(self.build_creator_container)
        scrollable_container.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        scrollable_container.grid_columnconfigure((0, 1, 2), weight=1)

        isolated_container = ctk.CTkFrame(scrollable_container, fg_color="transparent")
        isolated_container.grid(row=0, column=2, rowspan=4, sticky="new", padx=0, pady=0)
        isolated_container.grid_columnconfigure(0, weight=1)
        
        self.create_specialization_section(isolated_container, "demolitionist.png")
        self.create_weapon_section(isolated_container, "sub_machine_gun.png")

        self.create_gear_section(scrollable_container, "Mask", "mask.png", 0, 0)
        self.create_gear_section(scrollable_container, "Body Armor", "body_armor.png", 1, 0)
        self.create_gear_section(scrollable_container, "Holster", "holster.png", 2, 0)
        self.create_gear_section(scrollable_container, "Backpack", "backpack.png", 0, 1)
        self.create_gear_section(scrollable_container, "Gloves", "gloves.png", 1, 1)
        self.create_gear_section(scrollable_container, "Kneepads", "kneepads.png", 2, 1)

        self.create_skill_section(scrollable_container, "Left", "seeker_mine.png", 3, 0)
        self.create_skill_section(scrollable_container, "Right", "seeker_mine.png", 3, 1)


    def setup_build_tuning(self):
        """
        Populates the 'Build Tuning' tab with its UI components.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.

        """

        pass


    def setup_damage_output(self):
        """
        
        Populates the 'Damage Output' tab with its UI components.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.

        """

        pass


    ####################

    def load_build(self, build_dictionary):
        """
        Loads a build into the current environment.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            build_dictionary: A dictionary representing the build to load.
        
        """


        self.specialization_variable.set(build_dictionary["Specialization"])
        self.update_specialization_section()

        for variable_name, value in build_dictionary["Weapon"].items():
            if variable_name in WEAPON_VARIABLES_NO_EXPERTISE:
                if value:
                    self.weapon_data["Variable-Dropdown Pairs"][variable_name][0].set(value)
                    self.update_weapon_section(variable_name)

            elif variable_name == "Expertise":
                self.weapon_data["Variable-Label Pairs"]["Expertise"][0].set(value)
                self.weapon_data["Variable-Label Pairs"]["Expertise"][1].configure(text=f"Expertise: {value}")

            else:
                raise ValueError(f"Method 'load_build' encountered an invalid weapon variable name: '{variable_name}'")

        for skill_slot in SKILL_SLOTS:
            for variable_name, value in build_dictionary[skill_slot].items():
                if variable_name in SKILL_VARIABLES_NO_EXPERTISE:
                    if value:
                        self.skill_sections[skill_slot]["Variable-Dropdown Pairs"][variable_name][0].set(value)
                        self.update_skill_section(skill_slot, variable_name)

                elif variable_name == "Expertise":
                    self.skill_sections[skill_slot]["Variable-Label Pairs"]["Expertise"][0].set(value)
                    self.skill_sections[skill_slot]["Variable-Label Pairs"]["Expertise"][1].configure(text=f"Expertise: {value}")

                else:
                    raise ValueError(f"Method 'load_build' encountered an invalid skill variable name: '{variable_name}'")
            
        for gear_slot in GEAR_SLOTS:
            for variable_name, value in build_dictionary[gear_slot].items():
                if variable_name in GEAR_VARIABLES:
                    if value:
                        self.gear_sections[gear_slot]["Variable-Dropdown Pairs"][variable_name][0].set(value)
                        self.update_gear_section(gear_slot, variable_name)

                else:
                    raise ValueError(f"Method 'load_build' encountered an invalid gear variable name: '{variable_name}'")

        self.toggle_sidebar()


    def rename_build(self, build_dictionary):
        """
        Renames a build.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            build_dictionary: A dictionary representing the build to rename.
        
        """

        self.unbind("<Escape>")
        
        dialog_container = ctk.CTkFrame(self, width=400, height=135, corner_radius=10, border_width=2, border_color="gray30")
        dialog_container.place(relx=0.5, rely=0.5, anchor="center")
        dialog_container.pack_propagate(False)
        
        def close_dialog():
            dialog_container.destroy()

            if self.sidebar_container:
                self.bind("<Escape>", lambda event: self.toggle_sidebar())

        close_button = ctk.CTkButton(dialog_container, text="✕", command=close_dialog, width=30, height=30, font=ctk.CTkFont(size=16), fg_color="transparent")
        close_button.place(x=360, y=10)
        
        instruction_label = ctk.CTkLabel(dialog_container, text="Enter new build name:", font=ctk.CTkFont(size=14))
        instruction_label.pack(pady=(20, 5))

        build_name = build_dictionary["Build Name"]
        
        entry_field = ctk.CTkEntry(dialog_container, width=300, height=35)
        entry_field.pack(pady=5)
        entry_field.insert(0, build_name)
        entry_field.select_range(0, "end")
        
        character_counter = ctk.CTkLabel(dialog_container, text=f"{len(build_name)}/20", font=ctk.CTkFont(size=12), text_color="gray60")
        character_counter.pack(pady=(0, 10))
        
        def update_counter(*args):
            length = len(entry_field.get())

            if length > 20:
                entry_field.delete(20, "end")
                length = 20

            character_counter.configure(text=f"{length}/20", text_color="orange" if length == 20 else "gray60")
        
        entry_field.bind("<KeyPress>", update_counter)
        entry_field.bind("<KeyRelease>", update_counter)
        entry_field.focus()
        
        new_name = None
        
        def on_ok():
            nonlocal new_name
            text = entry_field.get().strip()

            if text:
                new_name = entry_field.get()
                close_dialog()
        
        entry_field.bind("<Return>", lambda event: on_ok())
        entry_field.bind("<Escape>", lambda event: close_dialog())
        dialog_container.bind("<Escape>", lambda event: close_dialog())
        
        self.wait_window(dialog_container)
        
        if new_name and new_name != build_name:            
            original_name = new_name
            counter = 1
            
            while True:
                if not (builds_directory / f"{self.name_to_filename(new_name)}.json").exists():
                    break

                new_name = f"{original_name} ({counter})"
                counter += 1

            if new_name != build_name:
            
                (builds_directory / f"{self.name_to_filename(build_name)}.json").unlink()

                build_dictionary["Build Name"] = new_name
                self.save_build_to_file(build_dictionary, builds_directory / f"{self.name_to_filename(new_name)}.json")

                self.toggle_sidebar()
                self.toggle_sidebar()

    
    def overwrite_build(self, build_dictionary):
        """
        Overwrites an existing build with the current selections.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            build_dictionary: A dictionary representing the build to overwrite.
        
        """

        build_name = build_dictionary["Build Name"]
        self.save_build_to_file(self.get_current_build_state(build_name), builds_directory / f"{self.name_to_filename(build_name)}.json")
        
        self.toggle_sidebar()
        self.toggle_sidebar()

    
    def delete_build(self, build_dictionary):
        """
        Deletes a build file.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            build_dictionary: A dictionary representing the build to delete.
        
        """

        (builds_directory / f"{self.name_to_filename(build_dictionary["Build Name"])}.json").unlink()

        self.toggle_sidebar()
        self.toggle_sidebar()


    def create_new_build(self):
        """
        Creates a new build by saving the current selections.
        
        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
        
        """

        self.unbind("<Escape>")
        
        dialog_container = ctk.CTkFrame(self, width=400, height=135, corner_radius=10, border_width=2, border_color="gray30")
        dialog_container.place(relx=0.5, rely=0.5, anchor="center")
        dialog_container.pack_propagate(False)
        
        def close_dialog():
            dialog_container.destroy()

            if self.sidebar_container:
                self.bind("<Escape>", lambda event: self.toggle_sidebar())
        
        close_button = ctk.CTkButton(dialog_container, text="✕", command=close_dialog, width=30, height=30, font=ctk.CTkFont(size=16), fg_color="transparent")
        close_button.place(x=360, y=10)
        
        instruction_label = ctk.CTkLabel(dialog_container, text="Enter build name:", font=ctk.CTkFont(size=14))
        instruction_label.pack(pady=(20, 5))
        
        entry_field = ctk.CTkEntry(dialog_container, width=300, height=35)
        entry_field.pack(pady=5)
        
        character_counter = ctk.CTkLabel(dialog_container, text="0/20", font=ctk.CTkFont(size=12), text_color="gray60")
        character_counter.pack(pady=(0, 10))
        
        def update_counter(*args):
            length = len(entry_field.get())

            if length > 20:
                entry_field.delete(20, "end")
                length = 20

            character_counter.configure(text=f"{length}/20", text_color="orange" if length == 20 else "gray60")
        
        entry_field.bind("<KeyPress>", update_counter)
        entry_field.bind("<KeyRelease>", update_counter)
        entry_field.focus()
        
        build_name = None
        
        def on_ok():
            nonlocal build_name
            text = entry_field.get().strip()

            if text:
                build_name = text
                close_dialog()
        
        entry_field.bind("<Return>", lambda event: on_ok())
        entry_field.bind("<Escape>", lambda event: close_dialog())
        dialog_container.bind("<Escape>", lambda event: close_dialog())
        
        self.wait_window(dialog_container)
        
        if build_name:
            original_name = build_name
            counter = 1

            while True:
                if not (builds_directory / f"{self.name_to_filename(build_name)}.json").exists():
                    break

                build_name = f"{original_name} ({counter})"
                counter += 1

            self.save_build_to_file(self.get_current_build_state(build_name), builds_directory / f"{self.name_to_filename(build_name)}.json")
            
            self.toggle_sidebar()
            self.toggle_sidebar()


    def create_specialization_section(self, parent_container, icon_file):
        """
        Creates the specialization selection section.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            parent_container: The parent container where the specialization section will be added.
            icon_file: The filename of the icon to display.

        """

        specialization_container = ctk.CTkFrame(parent_container, border_width=2)
        specialization_container.pack(padx=10, pady=10, fill="x")
        specialization_container.grid_columnconfigure(1, weight=1)
        specialization_container.grid_rowconfigure(0, weight=1)
        
        icon_path = assets_directory / icon_file
        icon_label = ctk.CTkLabel(specialization_container, image=ctk.CTkImage(light_image=Image.open(icon_path), dark_image=Image.open(icon_path), size=(40, 40)), text="")
        icon_label.grid(row=0, column=0, padx=10, pady=10, sticky="")
        
        dropdown_container = ctk.CTkFrame(specialization_container)
        dropdown_container.grid(row=0, column=1, sticky="ew", padx=10, pady=10)
        dropdown_container.grid_columnconfigure(0, weight=1)
        
        specialization_variable = ctk.StringVar(value="Select Specialization")
        specialization_dropdown = ctk.CTkComboBox(dropdown_container, values=SPECIALIZATIONS, variable=specialization_variable, command=lambda choice: self.update_specialization_section(), state="readonly")
        specialization_dropdown.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        
        self.specialization_variable = specialization_variable
        self.specialization_dropdown = specialization_dropdown


    def create_weapon_section(self, parent_container, icon_file):
        """
        Creates the weapon selection section.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            parent_container: The parent container where the weapon section will be added.
            icon_file: The filename of the icon to display.

        """

        weapon_container = ctk.CTkFrame(parent_container, border_width=2)
        weapon_container.pack(padx=10, pady=10, fill="both", expand=True)
        weapon_container.grid_columnconfigure(1, weight=1)
        weapon_container.grid_rowconfigure(0, weight=1)
        
        icon_path = assets_directory / icon_file
        icon_label = ctk.CTkLabel(weapon_container, image=ctk.CTkImage(light_image=Image.open(icon_path), dark_image=Image.open(icon_path), size=(40, 40)), text="")
        icon_label.grid(row=0, column=0, padx=10, pady=10, sticky="")
        
        dropdown_container = ctk.CTkFrame(weapon_container)
        dropdown_container.grid(row=0, column=1, sticky="ew", padx=10, pady=10)
        dropdown_container.grid_columnconfigure(0, weight=1)
        
        weapon_data = {
            "Dropdown Container": dropdown_container,
            "Variable-Dropdown Pairs": {},
            "Variable-Label Pairs": {},
            "Value History": {
                "Time": 0
            }
        }
        
        class_variable = ctk.StringVar(value="Select Class")
        class_dropdown = ctk.CTkComboBox(dropdown_container, values=WEAPON_CLASSES, variable=class_variable, command=lambda choice: self.update_weapon_section("Class"), state="readonly")
        class_dropdown.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        
        weapon_data["Variable-Dropdown Pairs"]["Class"] = (class_variable, class_dropdown)
        
        expertise_variable = ctk.IntVar(value=30)
        expertise_label = ctk.CTkLabel(dropdown_container, text=f"Expertise: 30")
        expertise_label.grid(row=100, column=0, padx=5, pady=(10, 2), sticky="w")
        
        expertise_slider = ctk.CTkSlider(dropdown_container, from_=0, to=30, number_of_steps=30, variable=expertise_variable, command=lambda value: expertise_label.configure(text=f"Expertise: {int(value)}"))
        expertise_slider.grid(row=101, column=0, padx=5, pady=(0, 5), sticky="ew")
        
        weapon_data["Variable-Label Pairs"]["Expertise"] = (expertise_variable, expertise_label)
        
        self.weapon_data = weapon_data


    def create_gear_section(self, parent_container, gear_slot, icon_file, row_index, column_index):
        """
        Creates a gear piece selection section.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            parent_container: The parent container where the gear section will be added.
            gear_slot: The name of the gear piece slot.
            icon_file: The filename of the icon to display.
            row_index: The row index for grid placement.
            column_index: The column index for grid placement.

        """

        gear_container = ctk.CTkFrame(parent_container, border_width=2)
        gear_container.grid(row=row_index, column=column_index, sticky="nsew", padx=10, pady=10)
        gear_container.grid_columnconfigure(1, weight=1)
        gear_container.grid_rowconfigure(0, weight=1)
        
        icon_path = assets_directory / icon_file
        icon_label = ctk.CTkLabel(gear_container, image=ctk.CTkImage(light_image=Image.open(icon_path), dark_image=Image.open(icon_path), size=(40, 40)), text="")
        icon_label.grid(row=0, column=0, padx=10, pady=10, sticky="")

        dropdown_container = ctk.CTkFrame(gear_container)
        dropdown_container.grid(row=0, column=1, sticky="ew", padx=10, pady=10)
        dropdown_container.grid_columnconfigure(0, weight=1)
        
        gear_data = {
            "Dropdown Container": dropdown_container,
            "Variable-Dropdown Pairs": {},
            "Value History": {
                "Time": 0
            }
        }
        
        type_variable = ctk.StringVar(value="Select Type")
        type_dropdown = ctk.CTkComboBox(dropdown_container, values=GEAR_TYPES, variable=type_variable, command=lambda choice: self.update_gear_section(gear_slot, "Type"), state="readonly")
        type_dropdown.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        
        gear_data["Variable-Dropdown Pairs"]["Type"] = (type_variable, type_dropdown)

        if not hasattr(self, "gear_sections"):
            self.gear_sections = {}

        self.gear_sections[gear_slot] = gear_data


    def create_skill_section(self, parent_container, skill_slot, icon_file, row_index, column_index):
        """
        Creates a skill selection section.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            parent_container: The parent container where the skill section will be added.
            skill_slot: The name of the skill slot.
            icon_file: The filename of the icon to display.
            row_index: The row index for grid placement.
            column_index: The column index for grid placement.

        """

        skill_container = ctk.CTkFrame(parent_container, border_width=2)
        skill_container.grid(row=row_index, column=column_index, sticky="nsew", padx=10, pady=(20, 10))
        skill_container.grid_columnconfigure(1, weight=1)
        skill_container.grid_rowconfigure(0, weight=1)

        icon_path = assets_directory / icon_file
        icon_label = ctk.CTkLabel(skill_container, image=ctk.CTkImage(light_image=Image.open(icon_path), dark_image=Image.open(icon_path), size=(40, 40)), text="")
        icon_label.grid(row=0, column=0, padx=10, pady=10, sticky="")
        
        dropdown_container = ctk.CTkFrame(skill_container)
        dropdown_container.grid(row=0, column=1, sticky="ew", padx=10, pady=10)
        dropdown_container.grid_columnconfigure(0, weight=1)
        
        skill_data = {
            "Dropdown Container": dropdown_container,
            "Variable-Dropdown Pairs": {},
            "Variable-Label Pairs": {}
        }
        
        class_variable = ctk.StringVar(value="Select Class")
        class_dropdown = ctk.CTkComboBox(dropdown_container, values=SKILL_CLASSES, variable=class_variable, command=lambda choice: self.update_skill_section(skill_slot, "Class"), state="readonly")
        class_dropdown.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        
        skill_data["Variable-Dropdown Pairs"]["Class"] = (class_variable, class_dropdown)
        
        expertise_variable = ctk.IntVar(value=30)
        expertise_label = ctk.CTkLabel(dropdown_container, text=f"Expertise: 30")
        expertise_label.grid(row=100, column=0, padx=5, pady=(10, 2), sticky="w")
        
        expertise_slider = ctk.CTkSlider(dropdown_container, from_=0, to=30, number_of_steps=30, variable=expertise_variable, command=lambda value: expertise_label.configure(text=f"Expertise: {int(value)}"))
        expertise_slider.grid(row=101, column=0, padx=5, pady=(0, 5), sticky="ew")
        
        skill_data["Variable-Label Pairs"]["Expertise"] = (expertise_variable, expertise_label)
        
        if not hasattr(self, "skill_sections"):
            self.skill_sections = {}

        if not hasattr(self, "skill_value_history"):
            self.skill_value_history = {
                "Time": 0
            }

        self.skill_sections[skill_slot] = skill_data

    
    ####################

    def name_to_filename(self, name):
        """
        Converts a name to a sanitized filename by removing invalid characters.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            name: The original name string.

        Returns:
            The sanitized filename.
        
        """

        return "".join(character for character in name if character not in "<>:\"/\\|?*").strip()


    def get_current_build_state(self, build_name):
        """
        Retrieves the current state of all selections as a dictionary.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            build_name: The name for the current build.

        Returns:
            A dictionary representing the current build selection.
        
        """

        build_dictionary = {"Build Name": build_name}        
        build_dictionary["Specialization"] = self.specialization_variable.get()
        weapon_dictionary = {variable: None for variable in WEAPON_VARIABLES}

        for variable in WEAPON_VARIABLES_NO_EXPERTISE:
            weapon_dictionary[variable] = self.weapon_data["Variable-Dropdown Pairs"][variable][0].get() if variable in self.weapon_data["Variable-Dropdown Pairs"] else None
    
        weapon_dictionary["Expertise"] = self.weapon_data["Variable-Label Pairs"]["Expertise"][0].get()
        build_dictionary["Weapon"] = weapon_dictionary

        skill_dictionaries = {skill_slot: {variable: None for variable in SKILL_VARIABLES} for skill_slot in SKILL_SLOTS}
        
        for skill_slot in SKILL_SLOTS:
            for variable in SKILL_VARIABLES_NO_EXPERTISE:
                skill_dictionaries[skill_slot][variable] = self.skill_sections[skill_slot]["Variable-Dropdown Pairs"][variable][0].get() if variable in self.skill_sections[skill_slot]["Variable-Dropdown Pairs"] else None

            skill_dictionaries[skill_slot]["Expertise"] = self.skill_sections[skill_slot]["Variable-Label Pairs"]["Expertise"][0].get()

        build_dictionary.update(skill_dictionaries)

        gear_dictionaries = {gear_slot: {variable: None for variable in GEAR_VARIABLES} for gear_slot in GEAR_SLOTS}

        for gear_slot in GEAR_SLOTS:
            for variable in GEAR_VARIABLES:
                gear_dictionaries[gear_slot][variable] = self.gear_sections[gear_slot]["Variable-Dropdown Pairs"][variable][0].get() if variable in self.gear_sections[gear_slot]["Variable-Dropdown Pairs"] else None

        build_dictionary.update(gear_dictionaries)

        return build_dictionary


    def save_build_to_file(self, build_dictionary, file_path):
        """
        Saves a given build dictionary to a specified file path in JSON format.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            build_dictionary: A dictionary representing the build to save.
            file_path: The file path where the build will be saved.
        
        """

        with open(file_path, "w") as file:
            json.dump(build_dictionary, file, indent=4)


    def update_specialization_section(self):
        """
        Updates the specialization section dynamically on user input.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.

        """

        specialization_value = self.specialization_variable.get()

        if specialization_value == "Select Specialization":
            return

        self.specialization_dropdown.configure(values=[specialization for specialization in SPECIALIZATIONS if specialization != specialization_value])

        self.update_weapon_section("Class")
        self.update_skill_section("Left", "Class")
        self.update_skill_section("Right", "Class")


    def update_weapon_section(self, trigger_point):
        """
        Updates the weapon section dynamically on user input.
        
        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            trigger_point: The variable that triggered the update.

        """

        if trigger_point == "Class":
            for variable_name in [variable_name for variable_name in list(self.weapon_data["Variable-Dropdown Pairs"].keys()) if variable_name != "Class"]:
                self.weapon_data["Variable-Dropdown Pairs"][variable_name][1].grid_forget()
                del self.weapon_data["Variable-Dropdown Pairs"][variable_name]

            class_value = self.weapon_data["Variable-Dropdown Pairs"]["Class"][0].get()
            self.weapon_data["Variable-Dropdown Pairs"]["Class"][1].configure(values=[weapon_class for weapon_class in WEAPON_CLASSES if weapon_class != class_value])

            if class_value == "Select Class":
                return

            if class_value in WEAPON_CLASSES_NO_SIGNATURE:
                type_value = self.check_history(self.weapon_data["Value History"], "Type", "Select Type", WEAPON_TYPES)
                type_variable = ctk.StringVar(value=type_value)
                type_dropdown = ctk.CTkComboBox(self.weapon_data["Dropdown Container"], values=[weapon_type for weapon_type in WEAPON_TYPES if weapon_type != type_value], variable=type_variable, command=lambda choice: self.update_weapon_section("Type"), state="readonly")
                type_dropdown.grid(row=1, column=0, padx=5, pady=5, sticky="ew")
                self.weapon_data["Variable-Dropdown Pairs"]["Type"] = (type_variable, type_dropdown)

                if type_value != "Select Type":
                    self.update_weapon_section("Type")

            elif class_value == "Signature Weapons":
                self.update_weapon_section("Type")

            else:
                raise ValueError(f"Method 'update_weapon_section' encountered an invalid class value: {class_value}")
            
        elif trigger_point == "Type":
            for variable_name in [variable_name for variable_name in list(self.weapon_data["Variable-Dropdown Pairs"].keys()) if variable_name not in ["Class", "Type"]]:
                self.weapon_data["Variable-Dropdown Pairs"][variable_name][1].grid_forget()
                del self.weapon_data["Variable-Dropdown Pairs"][variable_name]

            class_value = self.weapon_data["Variable-Dropdown Pairs"]["Class"][0].get()

            if class_value in WEAPON_CLASSES_NO_SIGNATURE:
                type_value = self.weapon_data["Variable-Dropdown Pairs"]["Type"][0].get()
                self.weapon_data["Variable-Dropdown Pairs"]["Type"][1].configure(values=[weapon_type for weapon_type in WEAPON_TYPES if weapon_type != type_value])

                if type_value == "Select Type":
                    return

                if "Type" not in self.weapon_data["Value History"]:
                    self.weapon_data["Value History"]["Type"] = []
                
                self.weapon_data["Value History"]["Type"].append((type_value, self.weapon_data["Value History"]["Time"]))
                self.weapon_data["Value History"]["Time"] += 1
                row_index = 2

            elif class_value == "Signature Weapons":
                row_index = 1
            
            else:
                raise ValueError(f"Method 'update_weapon_section' encountered an invalid class value: {class_value}")

            allowed_names = self.weapon_allowed_names(class_value, type_value if class_value in WEAPON_CLASSES_NO_SIGNATURE else None)
            name_value = self.check_history(self.weapon_data["Value History"], "Name", "Select Name", allowed_names)
            name_variable = ctk.StringVar(value=name_value)
            name_dropdown = ctk.CTkComboBox(self.weapon_data["Dropdown Container"], values=[allowed_name for allowed_name in allowed_names if allowed_name != name_value], variable=name_variable, command=lambda choice: self.update_weapon_section("Name"), state="readonly")
            name_dropdown.grid(row=row_index, column=0, padx=5, pady=5, sticky="ew")
            self.weapon_data["Variable-Dropdown Pairs"]["Name"] = (name_variable, name_dropdown)

            if name_value != "Select Name":
                self.update_weapon_section("Name")

        elif trigger_point == "Name":
            class_value = self.weapon_data["Variable-Dropdown Pairs"]["Class"][0].get()
            type_value = self.weapon_data["Variable-Dropdown Pairs"]["Type"][0].get() if class_value in WEAPON_CLASSES_NO_SIGNATURE else None
            name_value = self.weapon_data["Variable-Dropdown Pairs"]["Name"][0].get()

            if "(Select" in name_value:
                previous_name_value = self.check_history(self.weapon_data["Value History"], "Name", "Select Name", self.weapon_allowed_names(class_value, type_value))
                self.weapon_data["Variable-Dropdown Pairs"]["Name"][0].set(previous_name_value)
                return
            
            for variable_name in [variable_name for variable_name in list(self.weapon_data["Variable-Dropdown Pairs"].keys()) if variable_name not in ["Class", "Type", "Name"]]:
                self.weapon_data["Variable-Dropdown Pairs"][variable_name][1].grid_forget()
                del self.weapon_data["Variable-Dropdown Pairs"][variable_name]

            if name_value == "Select Name":
                return
            
            if "Name" not in self.weapon_data["Value History"]:
                self.weapon_data["Value History"]["Name"] = []

            self.weapon_data["Value History"]["Name"].append((name_value, self.weapon_data["Value History"]["Time"]))
            self.weapon_data["Value History"]["Time"] += 1

            self.weapon_data["Variable-Dropdown Pairs"]["Name"][1].configure(values=[allowed_name for allowed_name in self.weapon_allowed_names(class_value, type_value) if allowed_name != name_value])
            row_index = 3 if class_value in WEAPON_CLASSES_NO_SIGNATURE else 2

            ## Core Attributes ##
            core_attributes = self.weapon_core_attributes(class_value, name_value)

            for n, core_attribute in enumerate(core_attributes, 1):
                if core_attribute:
                    core_attribute_variable = ctk.StringVar(value=core_attribute)
                    core_attribute_dropdown = ctk.CTkComboBox(self.weapon_data["Dropdown Container"], values=[], variable=core_attribute_variable, state="disabled")
                    core_attribute_dropdown.grid(row=row_index, column=0, padx=5, pady=5, sticky="ew")
                    self.weapon_data["Variable-Dropdown Pairs"][f"Core Attribute {n}"] = (core_attribute_variable, core_attribute_dropdown)
                    row_index += 1

            ## Attribute ##
            allowed_attributes = self.weapon_allowed_attributes(class_value, name_value)

            if allowed_attributes:
                attribute_value = allowed_attributes[0] if len(allowed_attributes) == 1 else self.check_history(self.weapon_data["Value History"], "Attribute", "Select Attribute", allowed_attributes)
                attribute_variable = ctk.StringVar(value=attribute_value)
                attribute_dropdown = ctk.CTkComboBox(self.weapon_data["Dropdown Container"], values=[allowed_attribute for allowed_attribute in allowed_attributes if allowed_attribute != attribute_value], variable=attribute_variable, command=lambda choice: self.update_weapon_section("Attribute"), state="disabled" if len(allowed_attributes) == 1 else "readonly")
                attribute_dropdown.grid(row=row_index, column=0, padx=5, pady=5, sticky="ew")
                self.weapon_data["Variable-Dropdown Pairs"]["Attribute"] = (attribute_variable, attribute_dropdown)
                row_index += 1

            ## Mods ##
            allowed_modss= self.weapon_allowed_modss(class_value, name_value)

            for slot_name, allowed_mods in allowed_modss.items():
                if allowed_mods:
                    mod_value = allowed_mods[0] if len(allowed_mods) == 1 else self.check_history(self.weapon_data["Value History"], slot_name, f"Select {slot_name}", allowed_mods)
                    mod_variable = ctk.StringVar(value=mod_value)
                    mod_dropdown = ctk.CTkComboBox(self.weapon_data["Dropdown Container"], values=[allowed_mod for allowed_mod in allowed_mods if allowed_mod != mod_value], variable=mod_variable, command=lambda choice, slot_name=slot_name: self.update_weapon_section(slot_name), state="disabled" if len(allowed_mods) == 1 else "readonly")
                    mod_dropdown.grid(row=row_index, column=0, padx=5, pady=5, sticky="ew")
                    self.weapon_data["Variable-Dropdown Pairs"][slot_name] = (mod_variable, mod_dropdown)
                    row_index += 1

            ## Talents ##
            allowed_talentss = self.weapon_allowed_talentss(class_value, name_value)
            
            for n, allowed_talents in enumerate(allowed_talentss, 1):
                if allowed_talents:
                    talent_value = allowed_talents[0] if len(allowed_talents) == 1 else self.check_history(self.weapon_data["Value History"], f"Talent {n}", "Select Talent", allowed_talents)
                    talent_variable = ctk.StringVar(value=talent_value)
                    talent_dropdown = ctk.CTkComboBox(self.weapon_data["Dropdown Container"], values=[allowed_talent for allowed_talent in allowed_talents if allowed_talent != talent_value], variable=talent_variable, command=lambda choice, n=n: self.update_weapon_section(f"Talent {n}"), state="disabled" if len(allowed_talents) == 1 else "readonly")
                    talent_dropdown.grid(row=row_index, column=0, padx=5, pady=5, sticky="ew")
                    self.weapon_data["Variable-Dropdown Pairs"][f"Talent {n}"] = (talent_variable, talent_dropdown)
                    row_index += 1

        elif trigger_point in ["Core Attribute 1", "Core Attribute 2"]:
            return
        
        elif trigger_point == "Attribute":
            attribute_value = self.weapon_data["Variable-Dropdown Pairs"]["Attribute"][0].get()

            if attribute_value == "Select Attribute":
                return
            
            if "Attribute" not in self.weapon_data["Value History"]:
                self.weapon_data["Value History"]["Attribute"] = []

            self.weapon_data["Value History"]["Attribute"].append((attribute_value, self.weapon_data["Value History"]["Time"]))
            self.weapon_data["Value History"]["Time"] += 1

            class_value = self.weapon_data["Variable-Dropdown Pairs"]["Class"][0].get()
            name_value = self.weapon_data["Variable-Dropdown Pairs"]["Name"][0].get()

            allowed_attributes = self.weapon_allowed_attributes(class_value, name_value)

            self.weapon_data["Variable-Dropdown Pairs"]["Attribute"][1].configure(values=[allowed_attribute for allowed_attribute in allowed_attributes if allowed_attribute != attribute_value])

        elif trigger_point in WEAPON_MODS:
            class_value = self.weapon_data["Variable-Dropdown Pairs"]["Class"][0].get()
            name_value = self.weapon_data["Variable-Dropdown Pairs"]["Name"][0].get()
            mod_value = self.weapon_data["Variable-Dropdown Pairs"][trigger_point][0].get()

            if "(Select" in mod_value:
                previous_mod_value = self.check_history(self.weapon_data["Value History"], trigger_point, f"Select {trigger_point}", self.weapon_allowed_modss(class_value, name_value)[trigger_point])
                self.weapon_data["Variable-Dropdown Pairs"][trigger_point][0].set(previous_mod_value)
                return

            if mod_value == f"Select {trigger_point}":
                return
            
            if trigger_point not in self.weapon_data["Value History"]:
                self.weapon_data["Value History"][trigger_point] = []

            self.weapon_data["Value History"][trigger_point].append((mod_value, self.weapon_data["Value History"]["Time"]))
            self.weapon_data["Value History"]["Time"] += 1

            allowed_mods = self.weapon_allowed_modss(class_value, name_value)[trigger_point]

            self.weapon_data["Variable-Dropdown Pairs"][trigger_point][1].configure(values=[allowed_mod for allowed_mod in allowed_mods if allowed_mod != mod_value])

        elif trigger_point in ["Talent 1", "Talent 2"]:
            talent_value = self.weapon_data["Variable-Dropdown Pairs"][trigger_point][0].get()

            if talent_value == "Select Talent":
                return
            
            if trigger_point not in self.weapon_data["Value History"]:
                self.weapon_data["Value History"][trigger_point] = []

            self.weapon_data["Value History"][trigger_point].append((talent_value, self.weapon_data["Value History"]["Time"]))
            self.weapon_data["Value History"]["Time"] += 1

            class_value = self.weapon_data["Variable-Dropdown Pairs"]["Class"][0].get()
            name_value = self.weapon_data["Variable-Dropdown Pairs"]["Name"][0].get()

            allowed_talents = self.weapon_allowed_talentss(class_value, name_value)[int(trigger_point.split()[-1]) - 1]

            self.weapon_data["Variable-Dropdown Pairs"][trigger_point][1].configure(values=[allowed_talent for allowed_talent in allowed_talents if allowed_talent != talent_value])

        else:
            raise ValueError(f"Method 'update_weapon_section' received invalid 'trigger_point' parameter: {trigger_point}")


    def update_gear_section(self, gear_slot, trigger_point):
        """
        Updates a gear piece section dynamically on user input.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            gear_slot: The name of the gear piece slot.
            trigger_point: The variable that triggered the update.

        """

        gear_data = self.gear_sections[gear_slot]

        if trigger_point == "Type":
            self.enforce_gear_type_exclusivity()

            for variable_name in [variable_name for variable_name in list(gear_data["Variable-Dropdown Pairs"].keys()) if variable_name != "Type"]:
                gear_data["Variable-Dropdown Pairs"][variable_name][1].grid_forget()
                del gear_data["Variable-Dropdown Pairs"][variable_name]

            type_value = gear_data["Variable-Dropdown Pairs"]["Type"][0].get()

            if type_value == "Select Type":
                return
            
            allowed_names = self.gear_allowed_names(gear_slot, type_value)
            name_value = allowed_names[0] if len(allowed_names) == 1 else self.check_history(self.gear_sections[gear_slot]["Value History"], "Name", "Select Name", allowed_names)
            name_variable = ctk.StringVar(value=name_value)
            name_dropdown = ctk.CTkComboBox(gear_data["Dropdown Container"], values=[allowed_name for allowed_name in allowed_names if allowed_name != name_value], variable=name_variable, command=lambda choice: self.update_gear_section(gear_slot, "Name"), state="disabled" if len(allowed_names) == 1 else "readonly")
            name_dropdown.grid(row=1, column=0, padx=5, pady=5, sticky="ew")
            gear_data["Variable-Dropdown Pairs"]["Name"] = (name_variable, name_dropdown)

            if len(allowed_names) == 1 or name_value != "Select Name":
                self.update_gear_section(gear_slot, "Name")

        elif trigger_point == "Name":
            for variable_name in [variable_name for variable_name in list(gear_data["Variable-Dropdown Pairs"].keys()) if variable_name not in ["Type", "Name"]]:
                gear_data["Variable-Dropdown Pairs"][variable_name][1].grid_forget()
                del gear_data["Variable-Dropdown Pairs"][variable_name]

            name_value = gear_data["Variable-Dropdown Pairs"]["Name"][0].get()

            if name_value == "Select Name":
                return

            if "Name" not in gear_data["Value History"]:
                gear_data["Value History"]["Name"] = []
            
            gear_data["Value History"]["Name"].append((name_value, gear_data["Value History"]["Time"]))
            gear_data["Value History"]["Time"] += 1

            type_value = gear_data["Variable-Dropdown Pairs"]["Type"][0].get()
            gear_data["Variable-Dropdown Pairs"]["Name"][1].configure(values=[allowed_name for allowed_name in self.gear_allowed_names(gear_slot, type_value) if allowed_name != name_value])
            row_index = 2
            
            ## Core Attributes ##
            allowed_core_attributess = self.gear_allowed_core_attributess(type_value, name_value)

            for n, allowed_core_attributes in enumerate(allowed_core_attributess, 1):
                if allowed_core_attributes:
                    core_attribute_value = allowed_core_attributes[0] if len(allowed_core_attributes) == 1 else self.check_history(self.gear_sections[gear_slot]["Value History"], f"Core Attribute {n}", "Select Core Attribute", allowed_core_attributes)
                    core_attribute_variable = ctk.StringVar(value=core_attribute_value)
                    core_attribute_dropdown = ctk.CTkComboBox(gear_data["Dropdown Container"], values=[allowed_core_attribute for allowed_core_attribute in allowed_core_attributes if allowed_core_attribute != core_attribute_value], variable=core_attribute_variable, command=lambda choice, n=n: self.update_gear_section(gear_slot, f"Core Attribute {n}"), state="disabled" if len(allowed_core_attributes) == 1 else "readonly")
                    core_attribute_dropdown.grid(row=row_index, column=0, padx=5, pady=5, sticky="ew")
                    gear_data["Variable-Dropdown Pairs"][f"Core Attribute {n}"] = (core_attribute_variable, core_attribute_dropdown)
                    row_index += 1

            ## Attributes ##
            allowed_attributess = self.gear_allowed_attributess(type_value, name_value)

            multiple_value_n = []
            for n, allowed_attributes in enumerate(allowed_attributess, 1):
                if allowed_attributes and len(allowed_attributes) > 1:
                    multiple_value_n.append(n)

            if len(multiple_value_n) == 0:
                historical_values = {}

            elif len(multiple_value_n) == 1:
                n = multiple_value_n[0]
                historical_values = {f"Attribute {n}": self.check_history(self.gear_sections[gear_slot]["Value History"], f"Attribute {n}", "Select Attribute", allowed_attributess[n - 1])}

            elif len(multiple_value_n) == 2:
                n1, n2 = multiple_value_n
                historical_value1, historical_value2 = self.check_history(self.gear_sections[gear_slot]["Value History"], [f"Attribute {n1}", f"Attribute {n2}"], "Select Attribute", [allowed_attributess[n1 - 1], allowed_attributess[n2 - 1]])
                historical_values = {f"Attribute {n1}": historical_value1, f"Attribute {n2}": historical_value2}

            else:
                raise ValueError("Method 'update_gear_section' encountered more than two attributes with multiple possible values.")

            selected_attributes = set(historical_values.values())

            for n, allowed_attributes in enumerate(allowed_attributess, 1):
                if allowed_attributes:
                    attribute_value = allowed_attributes[0] if len(allowed_attributes) == 1 else historical_values[f"Attribute {n}"]
                    attribute_variable = ctk.StringVar(value=attribute_value)
                    attribute_dropdown = ctk.CTkComboBox(gear_data["Dropdown Container"], values=[allowed_attribute for allowed_attribute in allowed_attributes if allowed_attribute not in selected_attributes], variable=attribute_variable, command=lambda choice, n=n: self.update_gear_section(gear_slot, f"Attribute {n}"), state="disabled" if len(allowed_attributes) == 1 else "readonly")
                    attribute_dropdown.grid(row=row_index, column=0, padx=5, pady=5, sticky="ew")
                    gear_data["Variable-Dropdown Pairs"][f"Attribute {n}"] = (attribute_variable, attribute_dropdown)
                    row_index += 1

            ## Mods ##
            allowed_modss = self.gear_allowed_modss(gear_slot, type_value, name_value)

            for n, allowed_mods in enumerate(allowed_modss, 1):
                if allowed_mods:
                    mod_value = self.check_history(self.gear_sections[gear_slot]["Value History"], f"Mod {n}", "Select Mod", allowed_mods)
                    mod_variable = ctk.StringVar(value=mod_value)
                    mod_dropdown = ctk.CTkComboBox(gear_data["Dropdown Container"], values=[allowed_mod for allowed_mod in allowed_mods if allowed_mod != mod_value], variable=mod_variable, command=lambda choice, n=n: self.update_gear_section(gear_slot, f"Mod {n}"), state="readonly")
                    mod_dropdown.grid(row=row_index, column=0, padx=5, pady=5, sticky="ew")
                    gear_data["Variable-Dropdown Pairs"][f"Mod {n}"] = (mod_variable, mod_dropdown)
                    row_index += 1

            ## Talents ##
            allowed_talentss = self.gear_allowed_talentss(gear_slot, type_value, name_value)

            for n, allowed_talents in enumerate(allowed_talentss, 1):
                if allowed_talents:
                    talent_value = allowed_talents[0] if len(allowed_talents) == 1 else self.check_history(self.gear_sections[gear_slot]["Value History"], f"Talent {n}", "Select Talent", allowed_talents)
                    talent_variable = ctk.StringVar(value=talent_value)
                    talent_dropdown = ctk.CTkComboBox(gear_data["Dropdown Container"], values=[allowed_talent for allowed_talent in allowed_talents if allowed_talent != talent_value], variable=talent_variable, command=lambda choice, n=n: self.update_gear_section(gear_slot, f"Talent {n}"), state="disabled" if len(allowed_talents) == 1 else "readonly")
                    talent_dropdown.grid(row=row_index, column=0, padx=5, pady=5, sticky="ew")
                    gear_data["Variable-Dropdown Pairs"][f"Talent {n}"] = (talent_variable, talent_dropdown)
                    row_index += 1

        elif trigger_point in ["Core Attribute 1", "Core Attribute 2", "Core Attribute 3"]:
            core_attribute_value = gear_data["Variable-Dropdown Pairs"][trigger_point][0].get()

            if core_attribute_value == "Select Core Attribute":
                return

            if trigger_point not in gear_data["Value History"]:
                gear_data["Value History"][trigger_point] = []
            
            gear_data["Value History"][trigger_point].append((core_attribute_value, gear_data["Value History"]["Time"]))
            gear_data["Value History"]["Time"] += 1

            type_value = gear_data["Variable-Dropdown Pairs"]["Type"][0].get()
            name_value = gear_data["Variable-Dropdown Pairs"]["Name"][0].get()

            allowed_core_attributes = self.gear_allowed_core_attributess(type_value, name_value)[int(trigger_point[-1]) - 1]

            gear_data["Variable-Dropdown Pairs"][trigger_point][1].configure(values=[allowed_core_attribute for allowed_core_attribute in allowed_core_attributes if allowed_core_attribute != core_attribute_value])

        elif trigger_point in ["Attribute 1", "Attribute 2", "Attribute 3"]:
            attribute_value = gear_data["Variable-Dropdown Pairs"][trigger_point][0].get()

            if attribute_value == "Select Attribute":
                return

            if trigger_point not in gear_data["Value History"]:
                gear_data["Value History"][trigger_point] = []
            
            gear_data["Value History"][trigger_point].append((attribute_value, gear_data["Value History"]["Time"]))
            gear_data["Value History"]["Time"] += 1

            selected_attributes = set()
            for variable_name in ["Attribute 1", "Attribute 2", "Attribute 3"]:
                if variable_name in gear_data["Variable-Dropdown Pairs"]:
                    selected_attributes.add(gear_data["Variable-Dropdown Pairs"][variable_name][0].get())

            type_value = gear_data["Variable-Dropdown Pairs"]["Type"][0].get()
            name_value = gear_data["Variable-Dropdown Pairs"]["Name"][0].get()

            allowed_attributess = self.gear_allowed_attributess(type_value, name_value)

            for n, allowed_attributes in enumerate(allowed_attributess, 1):
                if allowed_attributes and len(allowed_attributes) > 1:
                    gear_data["Variable-Dropdown Pairs"][f"Attribute {n}"][1].configure(values=[allowed_attribute for allowed_attribute in allowed_attributes if allowed_attribute not in selected_attributes])

        elif trigger_point in ["Mod 1", "Mod 2"]:
            mod_value = gear_data["Variable-Dropdown Pairs"][trigger_point][0].get()

            if mod_value == "Select Mod":
                return
            
            if trigger_point not in gear_data["Value History"]:
                gear_data["Value History"][trigger_point] = []
            
            gear_data["Value History"][trigger_point].append((mod_value, gear_data["Value History"]["Time"]))
            gear_data["Value History"]["Time"] += 1

            type_value = gear_data["Variable-Dropdown Pairs"]["Type"][0].get()
            name_value = gear_data["Variable-Dropdown Pairs"]["Name"][0].get()

            allowed_mods = self.gear_allowed_modss(gear_slot, type_value, name_value)[int(trigger_point[-1]) - 1]

            gear_data["Variable-Dropdown Pairs"][trigger_point][1].configure(values=[allowed_mod for allowed_mod in allowed_mods if allowed_mod != mod_value])

        elif trigger_point in ["Talent 1", "Talent 2"]:
            talent_value = gear_data["Variable-Dropdown Pairs"][trigger_point][0].get()

            if talent_value == "Select Talent":
                return
            
            if trigger_point not in gear_data["Value History"]:
                gear_data["Value History"][trigger_point] = []
            
            gear_data["Value History"][trigger_point].append((talent_value, gear_data["Value History"]["Time"]))
            gear_data["Value History"]["Time"] += 1

            type_value = gear_data["Variable-Dropdown Pairs"]["Type"][0].get()
            name_value = gear_data["Variable-Dropdown Pairs"]["Name"][0].get()

            allowed_talents = self.gear_allowed_talentss(gear_slot, type_value, name_value)[int(trigger_point[-1]) - 1]

            gear_data["Variable-Dropdown Pairs"][trigger_point][1].configure(values=[allowed_talent for allowed_talent in allowed_talents if allowed_talent != talent_value])

        else:
            raise ValueError(f"Method 'update_gear_section' received invalid 'trigger_point' parameter: {trigger_point}")


    def update_skill_section(self, skill_slot, trigger_point):
        """
        Updates a skill section dynamically on user input.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            skill_slot: The name of the skill slot.
            trigger_point: The variable that triggered the update.

        """

        skill_data = self.skill_sections[skill_slot]

        if trigger_point == "Class":
            self.enforce_skill_class_exclusivity()

            for variable_name in [variable_name for variable_name in list(skill_data["Variable-Dropdown Pairs"].keys()) if variable_name != "Class"]:
                skill_data["Variable-Dropdown Pairs"][variable_name][1].grid_forget()
                del skill_data["Variable-Dropdown Pairs"][variable_name]

            class_value = skill_data["Variable-Dropdown Pairs"]["Class"][0].get()

            if class_value == "Select Class":
                return
            
            allowed_names = self.skill_allowed_names(class_value)
            name_value = allowed_names[0] if len(allowed_names) == 1 else self.check_history(self.skill_value_history, "Name", "Select Name", allowed_names)
            name_variable = ctk.StringVar(value=name_value)
            name_dropdown = ctk.CTkComboBox(skill_data["Dropdown Container"], values=[allowed_name for allowed_name in allowed_names if allowed_name != name_value], variable=name_variable, command=lambda choice: self.update_skill_section(skill_slot, "Name"), state="disabled" if len(allowed_names) == 1 else "readonly")
            name_dropdown.grid(row=1, column=0, padx=5, pady=5, sticky="ew")
            skill_data["Variable-Dropdown Pairs"]["Name"] = (name_variable, name_dropdown)

            if len(allowed_names) == 1 or name_value != "Select Name":
                self.update_skill_section(skill_slot, "Name")

        elif trigger_point == "Name":
            class_value = skill_data["Variable-Dropdown Pairs"]["Class"][0].get()
            name_value = skill_data["Variable-Dropdown Pairs"]["Name"][0].get()

            if "(Select" in name_value:
                previous_name_value = self.check_history(self.skill_value_history, "Name", "Select Name", self.skill_allowed_names(class_value))
                skill_data["Variable-Dropdown Pairs"]["Name"][0].set(previous_name_value)
                return

            for variable_name in [variable_name for variable_name in list(skill_data["Variable-Dropdown Pairs"].keys()) if variable_name not in ["Class", "Name"]]:
                skill_data["Variable-Dropdown Pairs"][variable_name][1].grid_forget()
                del skill_data["Variable-Dropdown Pairs"][variable_name]

            if name_value == "Select Name":
                return
            
            if "Name" not in self.skill_value_history:
                self.skill_value_history["Name"] = []

            self.skill_value_history["Name"].append((name_value, self.skill_value_history["Time"]))
            self.skill_value_history["Time"] += 1

            skill_data["Variable-Dropdown Pairs"]["Name"][1].configure(values=[allowed_name for allowed_name in self.skill_allowed_names(class_value) if allowed_name != name_value])
            row_index = 2

            ## Mods ##
            allowed_modss = self.skill_allowed_modss(class_value)

            for n, (slot_name, allowed_mods) in enumerate(allowed_modss.items(), 1):
                mod_value = allowed_mods[0] if len(allowed_mods) == 1 else self.check_history(self.skill_value_history, f"Mod {n}", f"Select {slot_name} Mod", allowed_mods, class_value)
                mod_variable = ctk.StringVar(value=mod_value)
                mod_dropdown = ctk.CTkComboBox(skill_data["Dropdown Container"], values=[allowed_mod for allowed_mod in allowed_mods if allowed_mod != mod_value], variable=mod_variable, command=lambda choice, n=n: self.update_skill_section(skill_slot, f"Mod {n}"), state="disabled" if len(allowed_mods) == 1 else "readonly")
                mod_dropdown.grid(row=row_index, column=0, padx=5, pady=5, sticky="ew")
                skill_data["Variable-Dropdown Pairs"][f"Mod {n}"] = (mod_variable, mod_dropdown)
                row_index += 1

        elif trigger_point in ["Mod 1", "Mod 2", "Mod 3"]:
            class_value = skill_data["Variable-Dropdown Pairs"]["Class"][0].get()
            mod_value = skill_data["Variable-Dropdown Pairs"][trigger_point][0].get()
        
            if "(Select" in mod_value:
                slot_name, allowed_mods = list(self.skill_allowed_modss(class_value).items())[int(trigger_point[-1]) - 1]
                previous_mod_value = self.check_history(self.skill_value_history, trigger_point, f"Select {slot_name} Mod", allowed_mods, class_value)
                skill_data["Variable-Dropdown Pairs"][trigger_point][0].set(previous_mod_value)
                return

            if mod_value.startswith("Select"):
                return
            
            if trigger_point not in self.skill_value_history:
                self.skill_value_history[trigger_point] = []

            self.skill_value_history[trigger_point].append((mod_value, self.skill_value_history["Time"], class_value))
            self.skill_value_history["Time"] += 1

            allowed_mods = list(self.skill_allowed_modss(class_value).values())[int(trigger_point[-1]) - 1]

            skill_data["Variable-Dropdown Pairs"][trigger_point][1].configure(values=[allowed_mod for allowed_mod in allowed_mods if allowed_mod != mod_value])

        else:
            raise ValueError(f"Method 'update_skill_section' received invalid 'trigger_point' parameter: {trigger_point}")


    ####################

    def check_history(self, history_dictionary, variable_name_or_names, default_value, allowed_values_or_valuess, constraint=None):
        """
        Checks the value history for one or two variables and returns the most recent valid value(s).

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            history_dictionary: The history dictionary for a particular section.
            variable_name_or_names: A single variable name (string) or a list of two variable names (list of strings).
            default_value: The default value to return if no valid historical value is found.
            allowed_values: A list of allowed values for the variable(s).
            constraint: An optional constraint that must be met for historical entries to be considered valid.

        """
        
        def find_valid_value(history_list, start_index, allowed_values):
            """
            Finds the next valid value in 'history_list' from 'start_index' backwards.

            Parameters:
                history_list: List of (value, timestamp) tuples.
                start_index: Index to start searching from.
                allowed_values: List of allowed values.        
            
            """

            for index in range(start_index, -1, -1):
                historical_entry = history_list[index]

                if len(historical_entry) == 2:
                    value, time = history_list[index]

                elif len(historical_entry) == 3:
                    value, time, entry_constraint = history_list[index]

                    if entry_constraint != constraint:
                        continue

                else:
                    raise ValueError("Method 'find_valid_value' encountered a historical entry with an invalid format.")

                if value in allowed_values:
                    return index, value, time

            return None, None, None
        
        if isinstance(variable_name_or_names, str):

            if any(isinstance(element, list) for element in allowed_values_or_valuess):
                raise ValueError(f"Method 'check_history' expected 'allowed_values_or_valuess' to be a list of strings when 'variable_name_or_names' is a string.")
            
            if variable_name_or_names not in history_dictionary:
                return default_value
            
            history_list = history_dictionary[variable_name_or_names]
            index, value, _ = find_valid_value(history_list, len(history_list) - 1, allowed_values_or_valuess)
            
            if index is not None:
                history_list.pop(index)
                history_list.append((value, history_dictionary["Time"]))
                history_dictionary["Time"] += 1
                return value
                    
            return default_value
        
        elif isinstance(variable_name_or_names, list) and len(variable_name_or_names) == 2:

            if not all(isinstance(element, list) for element in allowed_values_or_valuess) and len(allowed_values_or_valuess) == 2:
                raise ValueError(f"Method 'check_history' expected 'allowed_values_or_valuess' to be a list of two lists of strings when 'variable_name_or_names' is a list of two strings.")

            allowed_values_0, allowed_values_1 = allowed_values_or_valuess

            has_variable_0 = variable_name_or_names[0] in history_dictionary
            has_variable_1 = variable_name_or_names[1] in history_dictionary
            
            if not has_variable_0 and not has_variable_1:
                return default_value, default_value
            
            if has_variable_0 and not has_variable_1:
                return self.check_history(history_dictionary, variable_name_or_names[0], default_value, allowed_values_0), default_value
            
            if not has_variable_0 and has_variable_1:
                return default_value, self.check_history(history_dictionary, variable_name_or_names[1], default_value, allowed_values_1)
            
            history_list_0 = history_dictionary[variable_name_or_names[0]]
            history_list_1 = history_dictionary[variable_name_or_names[1]]
            
            pointer_0 = len(history_list_0) - 1
            pointer_1 = len(history_list_1) - 1
            
            while pointer_0 >= 0 or pointer_1 >= 0:
                index_0, value_0, time_0 = find_valid_value(history_list_0, pointer_0, allowed_values_0) if pointer_0 >= 0 else (None, None, None)
                index_1, value_1, time_1 = find_valid_value(history_list_1, pointer_1, allowed_values_1) if pointer_1 >= 0 else (None, None, None)
                
                if index_0 is None and index_1 is None:
                    return default_value, default_value
                
                if index_0 is None:
                    history_list_1.pop(index_1)
                    history_list_1.append((value_1, history_dictionary["Time"]))
                    history_dictionary["Time"] += 1
                    return default_value, value_1
                
                if index_1 is None:
                    history_list_0.pop(index_0)
                    history_list_0.append((value_0, history_dictionary["Time"]))
                    history_dictionary["Time"] += 1
                    return value_0, default_value
                
                if value_0 != value_1:
                    history_list_0.pop(index_0)
                    history_list_0.append((value_0, history_dictionary["Time"]))
                    history_list_1.pop(index_1)
                    history_list_1.append((value_1, history_dictionary["Time"]))
                    history_dictionary["Time"] += 1
                    return value_0, value_1
                
                if time_0 > time_1:
                    pointer_1 = index_1 - 1

                elif time_1 > time_0:
                    pointer_0 = index_0 - 1

                else:
                    raise Exception("Method 'check_history' encountered identical timestamps for two different values.")

        else:
            raise ValueError(f"Method 'check_history' received invalid 'variable_name_or_names' parameter: {variable_name_or_names}")


    def filter_names_by_specialization(self, names):
            """
            Filters item names based on the selected specialization.

            Parameters:
                self: The instance of the 'DamageCalculatorApp' class.
                names: List of item names to filter.

            Returns:
                List of filtered item names.

            """

            specialization_value = self.specialization_variable.get()
            filtered_names = []

            for name in names:
                if "(" in name:
                    base_name = name[:name.rindex("(") - 1]
                    parentheses_content = name[name.rindex("(") + 1:name.rindex(")")]

                    if parentheses_content in SPECIALIZATIONS:
                        if parentheses_content == specialization_value:
                            filtered_names.append(base_name)

                        else:
                            filtered_names.append(f"{base_name} (Select {parentheses_content})")

                        continue

                filtered_names.append(name)

            return filtered_names
    

    def weapon_allowed_names(self, weapon_class, weapon_type=None):
        """
        Returns the allowed weapon names for a given weapon class and type.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            weapon_class: The class of the weapon.
            weapon_type: The type of the weapon (optional).

        Returns:
            A list of allowed weapon names.
        
        """

        if weapon_class in WEAPON_CLASSES_NO_SIGNATURE:
            return self.filter_names_by_specialization(spreadsheet[weapon_class][spreadsheet[weapon_class]["Type"] == weapon_type]["Name"].tolist())
        
        elif weapon_class == "Signature Weapons":
            return self.filter_names_by_specialization(spreadsheet["Signature Weapons"]["Name"].tolist())
        
        else:
            raise ValueError(f"Method 'weapon_allowed_names' received invalid 'weapon_class' parameter: {weapon_class}")
    

    def unfiltered_weapon_name(self, weapon_class, weapon_name):
            """
            Returns the unfiltered weapon name by reintroducing any omitted specialization suffixes.

            Parameters:
                self: The instance of the 'DamageCalculatorApp' class.
                weapon_class: The class of the weapon.
                weapon_name: The name of the weapon.

            Returns:
                The unfiltered weapon name.
            
            """

            for candidate_name in spreadsheet[weapon_class]["Name"].tolist():
                if candidate_name.startswith(weapon_name):
                    return candidate_name

        
    def weapon_core_attributes(self, weapon_class, weapon_name):
        """
        Returns the core attributes for a given weapon.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            weapon_class: The class of the weapon.
            weapon_name: The name of the weapon.

        Returns:
            A list of core attributes.
        
        """
                
        if weapon_class in WEAPON_CLASSES_NO_SIGNATURE:
            return [(lambda value: None if pd.isna(value) else value)(spreadsheet[weapon_class][spreadsheet[weapon_class]["Name"] == self.unfiltered_weapon_name(weapon_class, weapon_name)].iloc[0][f"Core Attribute {n}"]) for n in [1, 2]]

        elif weapon_class == "Signature Weapons":
            return [None, None]
        
        else:
            raise ValueError(f"Method 'weapon_core_attributes' received invalid 'weapon_class' parameter: {weapon_class}")

    def weapon_allowed_attributes(self, weapon_class, weapon_name):
        """
        Returns the allowed attributes for a given weapon.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            weapon_class: The class of the weapon.
            weapon_name: The name of the weapon.

        Returns:
            A list of allowed attributes.
        
        """

        def cell_value_processor(cell_value):
            """
            Processes a cell value to determine allowed attributes.

            Parameters:
                cell_value: The value from the spreadsheet cell.

            Returns:
                A list of allowed attributes or 'None'.
                
            """

            if pd.isna(cell_value):
                return None

            elif cell_value == "*":
                return base_attributes

            elif cell_value.startswith("!"):
                return [attribute for attribute in base_attributes if attribute != cell_value[1:]]

            else:
                return [cell_value]

        base_attributes = spreadsheet["Weapon Attributes"]["Stats"].tolist()

        if weapon_class in WEAPON_CLASSES_NO_SIGNATURE:
            return cell_value_processor(spreadsheet[weapon_class][spreadsheet[weapon_class]["Name"] == self.unfiltered_weapon_name(weapon_class, weapon_name)].iloc[0][f"Attribute"])
        
        elif weapon_class == "Signature Weapons":
            return None
        
        else:
            raise ValueError(f"Method 'weapon_allowed_attributes' received invalid 'weapon_class' parameter: {weapon_class}")


    def weapon_allowed_modss(self, weapon_class, weapon_name):
        """
        Returns the allowed mods for a given weapon.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            weapon_class: The class of the weapon.
            weapon_name: The name of the weapon.

        Returns:
            A dictionary mapping mod slot names to lists of allowed mods.
        
        """

        def cell_value_processor(cell_value, slot_name):
            """
            Processes a cell value to determine allowed mods.

            Parameters:
                cell_value: The value from the spreadsheet cell.

            Returns:
                A list of allowed mods or 'None'.
                
            """

            if pd.isna(cell_value):
                return None

            elif cell_value.startswith("*"):
                return self.filter_names_by_specialization(spreadsheet[f"{slot_name}s"][spreadsheet[f"{slot_name}s"][cell_value[1:]] == "✓"]["Stats"].tolist())
            
            else:
                return [cell_value]

        if weapon_class in WEAPON_CLASSES_NO_SIGNATURE:
            return {slot_name: cell_value_processor(spreadsheet[weapon_class][spreadsheet[weapon_class]["Name"] == self.unfiltered_weapon_name(weapon_class, weapon_name)].iloc[0][slot_name], slot_name) for slot_name in WEAPON_MODS}
        
        elif weapon_class == "Signature Weapons":
            return {}
        
        else:
            raise ValueError(f"Method 'weapon_allowed_modss' received invalid 'weapon_class' parameter: {weapon_class}")


    def weapon_allowed_talentss(self, weapon_class, weapon_name):
        """
        Returns the allowed talents for a given weapon.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            weapon_class: The class of the weapon.
            weapon_name: The name of the weapon.

        Returns:
            A list of allowed talents.

        """

        def cell_value_processor(cell_value, weapon_class):
            """
            Processes a cell value to determine allowed talents.

            Parameters:
                cell_value: The value from the spreadsheet cell.

            Returns:
                A list of allowed talents or 'None'.
                
            """

            if pd.isna(cell_value):
                return None

            elif cell_value == "*":
                return spreadsheet["Weapon Talents"][spreadsheet["Weapon Talents"][weapon_class] == "✓"]["Name"].tolist()
            
            else:
                return [cell_value]

        if weapon_class in WEAPON_CLASSES_NO_SIGNATURE:
            return [cell_value_processor(spreadsheet[weapon_class][spreadsheet[weapon_class]["Name"] == self.unfiltered_weapon_name(weapon_class, weapon_name)].iloc[0][f"Talent {n}"], weapon_class) for n in [1, 2]]
        
        elif weapon_class == "Signature Weapons":
            return [[spreadsheet["Signature Weapons"][spreadsheet["Signature Weapons"]["Name"] == self.unfiltered_weapon_name(weapon_class, weapon_name)].iloc[0]["Talent"]], None]
        
        else:
            raise ValueError(f"Method 'weapon_allowed_talentss' received invalid 'weapon_class' parameter: {weapon_class}")


    def enforce_gear_type_exclusivity(self):
        """
        Updates all gear type dropdowns to enforce exclusivity.
        
        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.

        """

        base_types = GEAR_TYPES_NO_EXOTIC if any(gear_data["Variable-Dropdown Pairs"]["Type"][0].get() == "Exotic" for gear_data in self.gear_sections.values()) else GEAR_TYPES

        for gear_data in self.gear_sections.values():
            type_variable, type_dropdown = gear_data["Variable-Dropdown Pairs"]["Type"]
            type_dropdown.configure(values=[gear_type for gear_type in base_types if gear_type != type_variable.get()])
        
    
    def gear_allowed_names(self, gear_slot, gear_type):
        """
        Returns the allowed gear piece names for a given gear slot and type.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            gear_slot: The name of the gear piece slot.
            gear_type: The type of the gear piece.

        Returns:
            A list of allowed gear piece names.
        
        """

        if gear_type == "Improvised":
            return [f"Improvised {gear_slot}"]

        elif gear_type in ["Brand Set", "Gear Set"]:
            return spreadsheet[f"{gear_type}s"]["Name"].tolist()

        elif gear_type in ["Named", "Exotic"]:
            return spreadsheet[f"{gear_type} Gear"][spreadsheet[f"{gear_type} Gear"]["Slot"] == gear_slot]["Name"].tolist()

        else:
            raise ValueError(f"Method 'gear_allowed_names' received invalid 'gear_type' parameter: {gear_type}")

    
    def gear_allowed_core_attributess(self, gear_type, gear_name):
        """
        Returns the allowed core attributes for a given gear piece.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            gear_slot: The name of the gear piece slot.
            gear_type: The type of the gear piece.
            gear_name: The name of the gear piece.

        Returns:
            A list of allowed core attributes.
        
        """

        def cell_value_processor(cell_value):
            """
            Processes a cell value to determine allowed core attributes.
            
            Parameters:
                cell_value: The value from the spreadsheet cell.
                
            Returns:
                A list of allowed core attributes or 'None'.
                
            """

            if pd.isna(cell_value):
                return None

            elif cell_value == "*":
                return base_core_attributes

            elif cell_value.startswith("!"):
                return [core_attribute for core_attribute in base_core_attributes if core_attribute != cell_value[1:]]

            else:
                return [cell_value]

        base_core_attributes = spreadsheet["Gear Core Attributes"]["Stats"].tolist()

        if gear_type in ["Improvised", "Brand Set", "Gear Set"]:
            return [base_core_attributes, None, None]

        elif gear_type in ["Named", "Exotic"]:
            return [cell_value_processor(cell_value) for cell_value in [spreadsheet[f"{gear_type} Gear"][spreadsheet[f"{gear_type} Gear"]["Name"] == gear_name].iloc[0][f"Core Attribute {n}"] for n in [1, 2, 3]]]
        
        else:
            raise ValueError(f"Method 'gear_allowed_core_attributes' received invalid 'gear_type' parameter: {gear_type}")


    def gear_allowed_attributess(self, gear_type, gear_name):
        """
        Returns the allowed attributes for a given gear piece.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            gear_slot: The name of the gear piece slot.
            gear_type: The type of the gear piece.
            gear_name: The name of the gear piece.

        Returns:
            A list of allowed attributes.
        
        """

        def cell_value_processor(cell_value):
            """
            Processes a cell value to determine allowed attributes.

            Parameters:
                cell_value: The value from the spreadsheet cell.

            Returns:
                A list of allowed attributes or 'None'.
                
            """

            if pd.isna(cell_value):
                return None

            elif cell_value == "*":
                return base_attributes

            elif cell_value.startswith("*"):
                if cell_value[1:] == "Defensive":
                    return base_defensive_attributes

                elif cell_value[1:] == "Offensive":
                    return base_offensive_attributes

                elif cell_value[1:] == "Utility":
                    return base_utility_attributes

            elif cell_value.startswith("!"):
                return [attribute for attribute in base_attributes if attribute != cell_value[1:]]

            else:
                return [cell_value]

        base_defensive_attributes = spreadsheet["Gear Attributes"][spreadsheet["Gear Attributes"]["Category"] == "Defensive"]["Stats"].tolist()
        base_offensive_attributes = spreadsheet["Gear Attributes"][spreadsheet["Gear Attributes"]["Category"] == "Offensive"]["Stats"].tolist()
        base_utility_attributes = spreadsheet["Gear Attributes"][spreadsheet["Gear Attributes"]["Category"] == "Utility"]["Stats"].tolist()
        base_attributes = base_defensive_attributes + base_offensive_attributes + base_utility_attributes

        if gear_type in ["Improvised", "Brand Set"]:
            return [base_attributes, base_attributes, None]
        
        elif gear_type == "Gear Set":
            return [base_attributes, None, None]
        
        elif gear_type in ["Named", "Exotic"]:
            return [cell_value_processor(cell_value) for cell_value in [spreadsheet[f"{gear_type} Gear"][spreadsheet[f"{gear_type} Gear"]["Name"] == gear_name].iloc[0][f"Attribute {n}"] for n in [1, 2, 3]]]
        
        else:
            raise ValueError(f"Method 'gear_allowed_attributes' received invalid 'gear_type' parameter: {gear_type}")
        

    def gear_allowed_modss(self, gear_slot, gear_type, gear_name):
        """
        Returns the allowed mods for a given gear piece.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            gear_slot: The name of the gear piece slot.
            gear_type: The type of the gear piece.
            gear_name: The name of the gear piece.

        Returns:
            A list of allowed mods.
        
        """

        def cell_value_processor(cell_value):
            """
            Processes a cell value to determine allowed mods.

            Parameters:
                cell_value: The value from the spreadsheet cell.

            Returns:
                A list of allowed mods or 'None'.
                
            """

            if pd.isna(cell_value):
                return None

            elif cell_value == "*":
                return base_mods

            else:
                raise ValueError(f"Method 'gear_allowed_mods' encountered an invalid cell value: {cell_value}")

        base_mods = spreadsheet["Gear Mods"]["Stats"].tolist()

        if gear_type == "Improvised":
            return [base_mods, None]
        
        elif gear_type in ["Brand Set", "Gear Set"]:
            return [base_mods if gear_slot in ["Mask", "Body Armor", "Backpack"] else None, None]
        
        elif gear_type in ["Named", "Exotic"]:
            return [cell_value_processor(cell_value) for cell_value in [spreadsheet[f"{gear_type} Gear"][spreadsheet[f"{gear_type} Gear"]["Name"] == gear_name].iloc[0][f"Mod {n}"] for n in [1, 2]]]
        
        else:
            raise ValueError(f"Method 'gear_allowed_mods' received invalid 'gear_type' parameter: {gear_type}")
        

    def gear_allowed_talentss(self, gear_slot, gear_type, gear_name):
        """
        Returns the allowed talents for a given gear piece.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            gear_slot: The name of the gear piece slot.
            gear_type: The type of the gear piece.
            gear_name: The name of the gear piece.
            
        Returns:
            A list of allowed talents.
        
        """

        def cell_value_processor(cell_value):
            """
            Processes a cell value to determine allowed talents.

            Parameters:
                cell_value: The value from the spreadsheet cell.

            Returns:
                A list of allowed talents or 'None'.
                
            """

            if pd.isna(cell_value):
                return None

            else:
                return [cell_value]

        if gear_type in ["Improvised", "Brand Set"]:
            return [spreadsheet["Gear Talents"][spreadsheet["Gear Talents"]["Slot"] == gear_slot]["Name"].tolist() if gear_slot in ["Body Armor", "Backpack"] else None, None]
        
        elif gear_type == "Gear Set":
            return [[spreadsheet["Gear Sets"][spreadsheet["Gear Sets"]["Name"] == gear_name].iloc[0][f"{gear_slot} Talent"]] if gear_slot in ["Body Armor", "Backpack"] else None, None]
        
        elif gear_type in ["Named", "Exotic"]:
            return [cell_value_processor(cell_value) for cell_value in [spreadsheet[f"{gear_type} Gear"][spreadsheet[f"{gear_type} Gear"]["Name"] == gear_name].iloc[0][f"Talent {n}"] for n in [1, 2]]]
        
        else:
            raise ValueError(f"Method 'gear_allowed_talents' received invalid 'gear_type' parameter: {gear_type}")


    def enforce_skill_class_exclusivity(self):
        """
        Updates all skill class dropdowns to enforce exclusivity.
        
        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.

        """

        selected_classes = {skill_data["Variable-Dropdown Pairs"]["Class"][0].get() for skill_data in self.skill_sections.values()}

        for skill_data in self.skill_sections.values():
            skill_data["Variable-Dropdown Pairs"]["Class"][1].configure(values=[skill_class for skill_class in SKILL_CLASSES if skill_class not in selected_classes])

    
    def skill_allowed_names(self, skill_class):
        """
        Returns the allowed skill names for a given skill class.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            skill_class: The class of the skill.

        Returns:
            A list of allowed skill names.
        
        """

        return self.filter_names_by_specialization(spreadsheet["Skills"][spreadsheet["Skills"]["Class"] == skill_class].iloc[0]["Name"].split("; "))


    def skill_allowed_modss(self, skill_class):
        """
        Returns the allowed mods for a given skill class.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            skill_class: The class of the skill.

        Returns:
            A dictionary mapping mod slot names to lists of allowed mods.

        """

        return {slot_name: self.filter_names_by_specialization(spreadsheet[f"{skill_class} Mods"][spreadsheet[f"{skill_class} Mods"][slot_name] == "✓"]["Stats"].tolist()) for slot_name in spreadsheet[f"{skill_class} Mods"].columns[1:-1]}


#################### SECTION BREAK ####################

##### MAIN EXECUTION #####

if __name__ == "__main__":
    app = DamageCalculatorApp()
    app.mainloop()

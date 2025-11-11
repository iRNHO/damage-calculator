"""


"""

__author__ = "iRNHO"
__contact__ = "Message 'iRNHO' on XBOX or 'irnho' on discord regarding any questions, feedback, bug-reporting etc."
__discord__ = "https://discord.gg/--------"
__title__ = "iRNHO's Damage Calculator"


#################### SECTION BREAK ####################

##### CONSTANTS #####

GEAR_VARIABLES = ["Type", "Name", "Core Attribute 1", "Core Attribute 2", "Core Attribute 3", "Attribute 1", "Attribute 2", "Attribute 3", "Mod 1", "Mod 2", "Talent 1", "Talent 2"]
GEAR_CORE_ATTRIBUTES = ["15.0% Weapon Damage", "170,000 Armor", "1 Skill Tier"]
GEAR_SLOTS = ["Mask", "Body Armor", "Holster", "Backpack", "Gloves", "Kneepads"]
GEAR_TYPES = ["Improvised", "Brand Set", "Gear Set", "Named", "Exotic"]

WEAPON_VARIABLES = ["Class", "Type", "Name", "Core Attribute 1", "Core Attribute 2", "Attribute", "Optics Mod", "Magazine Mod", "Underbarrel Mod", "Muzzle Mod", "Talent 1", "Talent 2", "Expertise"]
WEAPON_CLASSES = ["Assault Rifles", "Light Machineguns", "Marksman Rifles", "Pistols", "Rifles", "Shotguns", "Sub Machine Guns", "Signature Weapons"]

SKILL_VARIABLES = ["Class", "Name", "Mod 1", "Mod 2", "Mod 3", "Expertise"]
SKILL_SLOTS = ["Skill Left", "Skill Right"]
SKILL_CLASSES = ["Ballistic Shield", "Chem Launcher", "Decoy", "Drone", "Firefly", "Hive", "Pulse", "Seeker Mine", "Smart Cover", "Sticky Bomb", "Trap", "Turret"]

SPECIALIZATIONS = ["Demolitionist", "Firewall", "Gunner", "Sharpshooter", "Survivalist", "Technician"]


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
assets_directory = root_directory / "assets"
builds_directory = root_directory / "builds"
data_directory = root_directory / "data"

builds_directory.mkdir(parents=True, exist_ok=True)

#dev
assets_directory = Path(r"C:\Users\smorg\Documents\damage-calculator\assets")
builds_directory = Path(r"C:\Users\smorg\Documents\damage-calculator\builds")
data_directory = Path(r"C:\Users\smorg\Documents\damage-calculator\data")
#dev

data = {name: pd.read_excel(data_directory / f"{name}.xlsx", sheet_name=None) for name in ["Armor", "Skill", "Talent", "Weapon"]}


#################### SECTION BREAK ####################

##### GUI CLASS #####

class DamageCalculatorApp(ctk.CTk):

    def __init__(self):
        """
        Initializes the application window and its components.

        Parameters:
            self: The instance of the DamageCalculatorApp class.

        """

        super().__init__()
        
        self.title("iRNHO's Damage Calculator")
        self.iconbitmap(assets_directory / "sub_machine_gun.ico")
        self.geometry("1000x700")

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        
        ctk.set_default_color_theme(str(assets_directory / "custom_theme.json"))
        ctk.set_appearance_mode("dark")        
        
        self.navigation_bar = ctk.CTkFrame(self, height=60, corner_radius=0)
        self.navigation_bar.grid(row=0, column=0, sticky="ew", padx=0, pady=0)
        self.navigation_bar.grid_columnconfigure(1, weight=1)
        
        self.burger_button = ctk.CTkButton(self.navigation_bar, text="☰", command=self.toggle_sidebar_visibility, width=40, height=40, font=ctk.CTkFont(size=20), fg_color="transparent")
        self.burger_button.grid(row=0, column=0, padx=10, pady=10, sticky="w")
        
        button_container = ctk.CTkFrame(self.navigation_bar, fg_color="transparent")
        button_container.grid(row=0, column=1, sticky="")
        
        navigation_buttons = [
            ("Build Creator", self.show_build_creator, "build_creator_button"),
            ("Build Tuning", self.show_build_tuning, "build_tuning_button"),
            ("Damage Output", self.show_damage_output, "damage_output_button")
        ]

        for tab_name, command, attribute_name in navigation_buttons:
            button = ctk.CTkButton(button_container, text=tab_name, command=command, width=200, height=50, font=ctk.CTkFont(size=16, weight="bold"))
            button.pack(side="left", padx=5)
            setattr(self, attribute_name, button)
        
        self.main_content_container = ctk.CTkFrame(self)
        self.main_content_container.grid(row=1, column=0, sticky="nsew")
        self.main_content_container.grid_rowconfigure(0, weight=1)
        self.main_content_container.grid_columnconfigure(0, weight=1)
        
        self.sidebar_container = None
        self.sidebar_visible = False
        
        self.build_creator_container = ctk.CTkFrame(self.main_content_container)
        self.build_tuning_container = ctk.CTkFrame(self.main_content_container)
        self.damage_output_container = ctk.CTkFrame(self.main_content_container)
        
        self.setup_build_creator()
        self.setup_build_tuning()
        self.setup_damage_output()
        
        self.build_creator_container.grid(row=0, column=0, sticky="nsew")
        self.current_tab = self.build_creator_container

    
    ####################

    def toggle_sidebar_visibility(self):
        """
        Toggles the sidebar visibility.
        
        Parameters:
            self: The instance of the DamageCalculatorApp class.
        
        """
        
        if self.sidebar_visible:
            if self.sidebar_container:
                self.sidebar_container.place_forget()

            self.sidebar_visible = False
            self.unbind("<Escape>")

        else:
            if self.sidebar_container:
                self.sidebar_container.destroy()
                self.sidebar_container = None
            
            self.create_sidebar()
            self.sidebar_container.place(x=0, y=0, relheight=1.0)
            self.sidebar_visible = True
            self.bind("<Escape>", lambda event: self.toggle_sidebar_visibility())


    def show_build_creator(self):
        """
        Displays the 'Build Creator' tab.

        Parameters:
            self: The instance of the DamageCalculatorApp class.

        """

        self.current_tab.grid_forget()

        self.build_creator_container.grid(row=0, column=0, sticky="nsew")
        self.current_tab = self.build_creator_container
        
        self.build_creator_button.configure(state="disabled")
        self.build_tuning_button.configure(state="normal")
        self.damage_output_button.configure(state="normal")
    
    
    def show_build_tuning(self):
        """
        Displays the 'Build Tuning' tab.
        
        Parameters:
            self: The instance of the DamageCalculatorApp class.
            
        """

        self.current_tab.grid_forget()

        self.build_tuning_container.grid(row=0, column=0, sticky="nsew")
        self.current_tab = self.build_tuning_container
        
        self.build_creator_button.configure(state="normal")
        self.build_tuning_button.configure(state="disabled")
        self.damage_output_button.configure(state="normal")
    

    def show_damage_output(self):
        """
        Displays the 'Damage Output' tab.
        
        Parameters:
            self: The instance of the DamageCalculatorApp class.

        """

        self.current_tab.grid_forget()

        self.damage_output_container.grid(row=0, column=0, sticky="nsew")
        self.current_tab = self.damage_output_container
        
        self.build_creator_button.configure(state="normal")
        self.build_tuning_button.configure(state="normal")
        self.damage_output_button.configure(state="disabled")


    def setup_build_creator(self):
        """
        Populates the 'Build Creator' tab with its UI components.

        Parameters:
            self: The instance of the DamageCalculatorApp class.
        
        """

        self.build_creator_container.grid_columnconfigure(0, weight=1)
        self.build_creator_container.grid_rowconfigure(0, weight=1)
        
        scrollable_container = ctk.CTkScrollableFrame(self.build_creator_container)
        scrollable_container.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        scrollable_container.grid_columnconfigure((0, 1, 2), weight=1)

        self.create_gear_section(scrollable_container, "Mask", "mask.png", 0, 0)
        self.create_gear_section(scrollable_container, "Body Armor", "body_armor.png", 1, 0)
        self.create_gear_section(scrollable_container, "Holster", "holster.png", 2, 0)
        self.create_gear_section(scrollable_container, "Backpack", "backpack.png", 0, 1)
        self.create_gear_section(scrollable_container, "Gloves", "gloves.png", 1, 1)
        self.create_gear_section(scrollable_container, "Kneepads", "kneepads.png", 2, 1)

        self.create_skill_section(scrollable_container, "Skill Left", "seeker_mine.png", 3, 0)
        self.create_skill_section(scrollable_container, "Skill Right", "seeker_mine.png", 3, 1)
        
        isolated_container = ctk.CTkFrame(scrollable_container, fg_color="transparent")
        isolated_container.grid(row=0, column=2, rowspan=4, sticky="new", padx=0, pady=0)
        isolated_container.grid_columnconfigure(0, weight=1)
        
        self.create_specialization_section(isolated_container, "demolitionist.png")
        self.create_weapon_section(isolated_container, "sub_machine_gun.png")


    def setup_build_tuning(self):
        """
        Populates the 'Build Tuning' tab with its UI components.

        Parameters:
            self: The instance of the DamageCalculatorApp class.

        """

        pass


    def setup_damage_output(self):
        """
        
        Populates the 'Damage Output' tab with its UI components.

        Parameters:
            self: The instance of the DamageCalculatorApp class.

        """

        pass


    ####################
    
    def create_sidebar(self):
        """
        Creates the sidebar menu.
        
        Parameters:
            self: The instance of the DamageCalculatorApp class.
        
        """

        self.sidebar_container = ctk.CTkFrame(self, width=550, corner_radius=0)
        self.sidebar_container.pack_propagate(False)
        
        close_button = ctk.CTkButton(self.sidebar_container, text="✕", command=self.toggle_sidebar_visibility, width=30, height=30, font=ctk.CTkFont(size=16), fg_color="transparent")
        close_button.place(x=510, y=10)
        
        scrollable_container = ctk.CTkScrollableFrame(self.sidebar_container)
        scrollable_container.pack(fill="both", expand=True, padx=10, pady=(50, 10))
        
        self.refresh_build_list(scrollable_container)

    
    def create_gear_section(self, parent_container, gear_slot, icon_file, row, column):
        """
        Creates a gear piece selection section.

        Parameters:
            self: The instance of the DamageCalculatorApp class.
            parent_container: The parent container where the gear section will be added.
            gear_name: The name of the gear piece.
            icon_file: The filename of the icon to display.
            row: The row position in the grid.
            column: The column position in the grid.

        """

        gear_container = ctk.CTkFrame(parent_container, border_width=2)
        gear_container.grid(row=row, column=column, sticky="nsew", padx=10, pady=10)
        gear_container.grid_columnconfigure(1, weight=1)
        gear_container.grid_rowconfigure(0, weight=1)
        
        icon_path = assets_directory / icon_file
        icon_label = ctk.CTkLabel(gear_container, image=ctk.CTkImage(light_image=Image.open(icon_path), dark_image=Image.open(icon_path), size=(40, 40)), text="")
        icon_label.grid(row=0, column=0, padx=10, pady=10, sticky="")

        dropdown_container = ctk.CTkFrame(gear_container)
        dropdown_container.grid(row=0, column=1, sticky="ew", padx=10, pady=10)
        dropdown_container.grid_columnconfigure(0, weight=1)
        
        gear_data = {
            "Slot": gear_slot,
            "Frame": dropdown_container,
            "Variables": {},
            "Dropdowns": {}
        }
        
        type_variable = ctk.StringVar(value="Select Type")
        type_dropdown = ctk.CTkComboBox(dropdown_container, values=GEAR_TYPES, variable=type_variable, command=lambda choice: self.update_gear(gear_slot, gear_data, trigger_point="Type"), state="readonly")
        type_dropdown.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        
        gear_data["Variables"]["Type"] = type_variable
        gear_data["Dropdowns"]["Type"] = type_dropdown
        
        if not hasattr(self, "gear_sections"):
            self.gear_data = {}

        self.gear_data[gear_slot] = gear_data


    def create_skill_section(self, parent_container, skill_slot, icon_file, row, column):
        """
        Creates a skill selection section.
        
        Parameters:
            self: The instance of the DamageCalculatorApp class.
            parent_container: The parent container where the skill section will be added.
            skill_name: The name of the skill.
            icon_file: The filename of the icon to display.
            row: The row position in the grid.
            column: The column position in the grid.
        
        """

        skill_container = ctk.CTkFrame(parent_container, border_width=2)
        skill_container.grid(row=row, column=column, sticky="nsew", padx=10, pady=(20, 10))
        skill_container.grid_columnconfigure(1, weight=1)
        skill_container.grid_rowconfigure(0, weight=1)

        icon_path = assets_directory / icon_file
        icon_label = ctk.CTkLabel(skill_container, image=ctk.CTkImage(light_image=Image.open(icon_path), dark_image=Image.open(icon_path), size=(40, 40)), text="")
        icon_label.grid(row=0, column=0, padx=10, pady=10, sticky="")
        
        dropdown_container = ctk.CTkFrame(skill_container)
        dropdown_container.grid(row=0, column=1, sticky="ew", padx=10, pady=10)
        dropdown_container.grid_columnconfigure(0, weight=1)
        
        skill_data = {
            "Slot": skill_slot,
            "Frame": dropdown_container,
            "Variables": {},
            "Dropdowns": {},
            "Labels": {}
        }
        
        class_variable = ctk.StringVar(value="Select Class")
        class_dropdown = ctk.CTkComboBox(dropdown_container, values=SKILL_CLASSES, variable=class_variable, command=lambda choice: self.update_skill(skill_slot, skill_data, trigger_point="Class"), state="readonly")
        class_dropdown.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        
        skill_data["Variables"]["Class"] = class_variable
        skill_data["Dropdowns"]["Class"] = class_dropdown
        
        expertise_variable = ctk.IntVar(value=30)
        expertise_label = ctk.CTkLabel(dropdown_container, text=f"Expertise: 30")
        expertise_label.grid(row=100, column=0, padx=5, pady=(10, 2), sticky="w")
        
        expertise_slider = ctk.CTkSlider(dropdown_container, from_=0, to=30, number_of_steps=30, variable=expertise_variable, command=lambda value: expertise_label.configure(text=f"Expertise: {int(value)}"))
        expertise_slider.grid(row=101, column=0, padx=5, pady=(0, 5), sticky="ew")
        
        skill_data["Variables"]["Expertise"] = expertise_variable
        skill_data["Labels"]["Expertise"] = expertise_label
        
        if not hasattr(self, "skill_sections"):
            self.skill_data = {}

        self.skill_data[skill_slot] = skill_data


    def create_specialization_section(self, parent_container, icon_file):
        """
        Creates the specialization selection section.
        
        Parameters:
            self: The instance of the DamageCalculatorApp class.
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
        specialization_dropdown = ctk.CTkComboBox(dropdown_container, values=SPECIALIZATIONS, variable=specialization_variable, command=lambda choice: self.update_specialization(specialization_variable, specialization_dropdown), state="readonly")
        specialization_dropdown.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        
        self.specialization_variable = specialization_variable
        self.specialization_dropdown = specialization_dropdown


    def create_weapon_section(self, parent_container, icon_file):
        """
        Creates the weapon selection section.

        Parameters:
            self: The instance of the DamageCalculatorApp class.
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
            "Frame": dropdown_container,
            "Variables": {},
            "Dropdowns": {},
            "Labels": {}
        }
        
        class_variable = ctk.StringVar(value="Select Class")
        class_dropdown = ctk.CTkComboBox(dropdown_container, values=WEAPON_CLASSES, variable=class_variable, command=lambda choice: self.update_weapon("Weapon", weapon_data, trigger_point="Class"), state="readonly")
        class_dropdown.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        
        weapon_data["Variables"]["Class"] = class_variable
        weapon_data["Dropdowns"]["Class"] = class_dropdown
        
        expertise_variable = ctk.IntVar(value=30)
        expertise_label = ctk.CTkLabel(dropdown_container, text=f"Expertise: 30")
        expertise_label.grid(row=100, column=0, padx=5, pady=(10, 2), sticky="w")
        
        expertise_slider = ctk.CTkSlider(dropdown_container, from_=0, to=30, number_of_steps=30, variable=expertise_variable, command=lambda value: expertise_label.configure(text=f"Expertise: {int(value)}"))
        expertise_slider.grid(row=101, column=0, padx=5, pady=(0, 5), sticky="ew")
        
        weapon_data["Variables"]["Expertise"] = expertise_variable
        weapon_data["Labels"]["Expertise"] = expertise_label
        
        self.weapon_data = weapon_data


    ####################

    def refresh_build_list(self, scrollable_container):
        """
        Refreshes the list of builds within the sidebar.

        Parameters:
            self: The instance of the DamageCalculatorApp class.
            scrollable_container: The scrollable container to populate with build entries.
        
        """

        for build_entry in scrollable_container.winfo_children():
            build_entry.destroy()
                
        for build_dictionary in self.get_all_builds():
            self.create_build_entry(scrollable_container, build_dictionary)
        
        new_build_container = ctk.CTkFrame(scrollable_container, height=80)
        new_build_container.pack(fill="x", padx=5, pady=10)
        
        new_build_button = ctk.CTkButton(new_build_container, text="+ Save Build", command=self.create_new_build, height=60, font=ctk.CTkFont(size=16, weight="bold"))
        new_build_button.pack(fill="both", expand=True, padx=10, pady=10)


    ####################
    
    def get_all_builds(self):
        """
        Retrieves all saved builds from the builds directory.
        
        Parameters:
            self: The instance of the DamageCalculatorApp class.

        Returns:
            List of dictionaries representing each build.        

        """

        build_dictionaries = []
        
        for build_file in builds_directory.iterdir():
            with open(build_file, "r") as file:
                build_dictionaries.append(json.load(file))

        return sorted(build_dictionaries, key=lambda build_dictionary: build_dictionary["Build Name"].lower())
    

    def create_build_entry(self, scrollable_container, build_dictionary):
        """
        Creates a single build entry with load, rename, overwrite, and delete buttons.

        Parameters:
            self: The instance of the DamageCalculatorApp class.
            scrollable_container: The container to which the build entry will be added.
            build_dictionary: A dictionary containing the build's data.
        
        """

        build_container = ctk.CTkFrame(scrollable_container)
        build_container.pack(fill="x", padx=5, pady=5)
        
        name_label = ctk.CTkLabel(build_container, text=build_dictionary["Build Name"], font=ctk.CTkFont(size=14, weight="bold"), anchor="w")
        name_label.pack(side="left", padx=10, pady=10, fill="x", expand=True)
        
        button_container = ctk.CTkFrame(build_container, fg_color="transparent")
        button_container.pack(side="right", padx=5, pady=5)

        buttons = [
            ("load", self.load_build),
            ("rename", self.rename_build),
            ("overwrite", self.overwrite_build),
            ("delete", self.delete_build)
        ]

        for icon_name, command in buttons:
            icon_path = assets_directory / f"{icon_name}.png"
            button = ctk.CTkButton(button_container, text="", image=ctk.CTkImage(light_image=Image.open(icon_path), dark_image=Image.open(icon_path), size=(30, 30)), command=lambda command=command: command(build_dictionary), width=40, height=40, fg_color="transparent", border_width=2, border_color="gray30")
            button.pack(side="left", padx=2)


    def create_new_build(self):
        """
        Creates a new build by saving the current selections.
        
        Parameters:
            self: The instance of the DamageCalculatorApp class.
        
        """

        self.unbind("<Escape>")
        
        dialog_container = ctk.CTkFrame(self, width=400, height=135, corner_radius=10, border_width=2, border_color="gray30")
        dialog_container.place(relx=0.5, rely=0.5, anchor="center")
        dialog_container.pack_propagate(False)
        
        def close_dialog():
            dialog_container.destroy()

            if self.sidebar_visible:
                self.bind("<Escape>", lambda event: self.toggle_sidebar_visibility())
        
        close_button = ctk.CTkButton(dialog_container, text="✕", command=close_dialog, width=30, height=30, font=ctk.CTkFont(size=16), fg_color="transparent")
        close_button.place(x=360, y=10)
        
        label = ctk.CTkLabel(dialog_container, text="Enter build name:", font=ctk.CTkFont(size=14))
        label.pack(pady=(20, 5))
        
        entry = ctk.CTkEntry(dialog_container, width=300, height=35)
        entry.pack(pady=5)
        
        character_counter = ctk.CTkLabel(dialog_container, text="0/20", font=ctk.CTkFont(size=12), text_color="gray60")
        character_counter.pack(pady=(0, 10))
        
        def update_counter(*args):
            text = entry.get()

            if len(text) > 20:
                entry.delete(20, "end")
                text = entry.get()

            count = len(text)
            character_counter.configure(text=f"{count}/20")

            if count == 20:
                character_counter.configure(text_color="#ff7f27")

            else:
                character_counter.configure(text_color="gray60")
        
        entry.bind("<KeyRelease>", update_counter)
        entry.focus()
        
        build_name = None
        
        def on_ok():
            nonlocal build_name
            text = entry.get().strip()

            if text:
                build_name = text
                close_dialog()
        
        entry.bind("<Return>", lambda event: on_ok())
        entry.bind("<Escape>", lambda event: close_dialog())
        dialog_container.bind("<Escape>", lambda event: close_dialog())
        
        self.wait_window(dialog_container)
        
        if build_name:
            original_name = build_name
            counter = 1

            while True:
                if not (builds_directory / f"{self.sanitize_filename(build_name)}.json").exists():
                    break

                build_name = f"{original_name} ({counter})"
                counter += 1
            
            build_data = self.get_current_build_state(build_name)
            
            file_path = builds_directory / f"{self.sanitize_filename(build_name)}.json"
            self.save_build_to_file(build_data, file_path)
            
            self.toggle_sidebar_visibility()
            self.toggle_sidebar_visibility()


    ####################

    def load_build(self, build_dictionary):
        """
        Loads a build into the current environment.

        Parameters:
            self: The instance of the DamageCalculatorApp class.
            build_dictionary: A dictionary containing the build's data.
        
        """

        specialization_value = build_dictionary["Specialization"]
        if specialization_value:
            self.specialization_variable.set(specialization_value)
            self.update_specialization(self.specialization_variable, self.specialization_dropdown)

        weapon_dictionary = build_dictionary["Weapon"]

        for key, value in weapon_dictionary.items():
            if value:
                self.weapon_data["Variables"][key].set(value)
                self.update_weapon("Weapon", self.weapon_data, key)
                
                if key in ["Class", "Type", "Name"]:
                    self.after(50)

                elif key == "Expertise": # MOVE TO WHERE DROPDOWNS ARE UPDATED??????????
                    self.weapon_data["Labels"]["Expertise"].configure(text=f"Expertise: {value}")

        skill_dictionaries = {
            "Skill Left": build_dictionary["Skill Left"],
            "Skill Right": build_dictionary["Skill Right"]
        }

        for skill_name, skill_dictionary in skill_dictionaries.items():
            for key, value in skill_dictionary.items():
                if value:
                    self.skill_data[skill_name]["Variables"][key].set(value)
                    self.update_skill(skill_name, self.skill_data[skill_name], key)
                    
                    if key in ["Class", "Name"]:
                        self.after(50)

                    elif key == "Expertise":
                        self.skill_data[skill_name]["Labels"]["Expertise"].configure(text=f"Expertise: {value}")
            
        for gear_name in ["Mask", "Body Armor", "Holster", "Backpack", "Gloves", "Kneepads"]:
            gear_dictionary = build_dictionary[gear_name]

            for key, value in gear_dictionary.items():
                if value:
                    self.gear_data[gear_name]["Variables"][key].set(value)
                    self.update_gear(gear_name, self.gear_data[gear_name], key)

                    if key in ["Type", "Name"]:
                        self.after(50)

        self.toggle_sidebar_visibility()


    def rename_build(self, build_dictionary):
        """
        Renames a build.

        Parameters:
            self: The instance of the DamageCalculatorApp class.
            build_dictionary: A dictionary containing the build's data.
        
        """

        self.unbind("<Escape>")
        build_name = build_dictionary["Build Name"]
        
        dialog_container = ctk.CTkFrame(self, width=400, height=135, corner_radius=10, border_width=2, border_color="gray30")
        dialog_container.place(relx=0.5, rely=0.5, anchor="center")
        dialog_container.pack_propagate(False)
        
        def close_dialog():
            dialog_container.destroy()

            if self.sidebar_visible:
                self.bind("<Escape>", lambda event: self.toggle_sidebar_visibility())

        close_button = ctk.CTkButton(dialog_container, text="✕", command=close_dialog, width=30, height=30, font=ctk.CTkFont(size=16), fg_color="transparent")
        close_button.place(x=360, y=10)
        
        label = ctk.CTkLabel(dialog_container, text="Enter new build name:", font=ctk.CTkFont(size=14))
        label.pack(pady=(20, 5))
        
        entry = ctk.CTkEntry(dialog_container, width=300, height=35)
        entry.pack(pady=5)
        entry.insert(0, build_name)
        entry.select_range(0, "end")
        
        character_counter = ctk.CTkLabel(dialog_container, text=f"{len(build_name)}/20", font=ctk.CTkFont(size=12), text_color="gray60")
        character_counter.pack(pady=(0, 10))
        
        def update_counter(*args):
            text = entry.get()

            if len(text) > 20:
                entry.delete(20, "end")
                text = entry.get()

            count = len(text)
            character_counter.configure(text=f"{count}/20")

            if count == 20:
                character_counter.configure(text_color="#ff7f27")

            else:
                character_counter.configure(text_color="gray60")
        
        entry.bind("<KeyRelease>", update_counter)
        entry.focus()
        
        new_name = None
        
        def on_ok():
            nonlocal new_name
            text = entry.get().strip()

            if text:
                new_name = entry.get()
                close_dialog()
        
        entry.bind("<Return>", lambda event: on_ok())
        entry.bind("<Escape>", lambda event: close_dialog())
        dialog_container.bind("<Escape>", lambda event: close_dialog())
        
        self.wait_window(dialog_container)
        
        if new_name and new_name != build_name:            
            original_name = new_name
            counter = 1
            
            while True:
                if not (builds_directory / f"{self.sanitize_filename(new_name)}.json").exists():
                    break

                new_name = f"{original_name} ({counter})"
                counter += 1
            
            build_dictionary["Build Name"] = new_name
            
            (builds_directory / (f"{self.sanitize_filename(build_name)}.json")).unlink()
            
            self.save_build_to_file(build_dictionary, builds_directory / f"{self.sanitize_filename(new_name)}.json")

            self.toggle_sidebar_visibility()
            self.toggle_sidebar_visibility()

    
    def overwrite_build(self, build_dictionary):
        """
        Overwrites an existing build with the current selections.

        Parameters:
            self: The instance of the DamageCalculatorApp class.
            build_dictionary: A dictionary containing the build's data.
        
        """

        build_name = build_dictionary["Build Name"]
        
        file_path = builds_directory / f"{self.sanitize_filename(build_name)}.json"
        
        self.save_build_to_file(self.get_current_build_state(build_name), file_path)
        
        self.toggle_sidebar_visibility()
        self.toggle_sidebar_visibility()

    
    def delete_build(self, build_dictionary):
        """
        Deletes a build file.

        Parameters:
            self: The instance of the DamageCalculatorApp class.
            build_dictionary: A dictionary containing the build's data.
        
        """

        (builds_directory / (f"{self.sanitize_filename(build_dictionary["Build Name"])}.json")).unlink()

        self.toggle_sidebar_visibility()
        self.toggle_sidebar_visibility()


    def sanitize_filename(self, build_name):
        """
        Converts a build name to a valid filename.

        Parameters:
            self: The instance of the DamageCalculatorApp class.
            build_name: The original build name.

        Returns:
            The sanitized filename.
        
        """

        return "".join(character for character in build_name if character not in "<>:\"/\\|?*").strip()

    
    def get_current_build_state(self, build_name):
        """
        Retrieves the current state of all selections as a dictionary.

        Parameters:
            self: The instance of the DamageCalculatorApp class.
            build_name: The name of the build.

        Returns:
            A dictionary representing the current build state.
        
        """

        build_dictionary = {"Build Name": build_name}        

        specialization_value = self.specialization_variable.get()
        build_dictionary["Specialization"] = None if specialization_value == "Select Specialization" else specialization_value

        weapon_dictionary = {variable: None for variable in WEAPON_VARIABLES}

        for variable in WEAPON_VARIABLES:
            try:
                value = self.weapon_data["Variables"][variable].get() # THIS WOULD CHANGE IF THE DATA WAS ALL NONE TO START

            except Exception:
                continue

            if isinstance(value, str):
                if value.startswith("Select"):
                    continue
            
            weapon_dictionary[variable] = value
    
        build_dictionary["Weapon"] = weapon_dictionary

        skill_dictionaries = {skill_slot: {variable: None for variable in SKILL_VARIABLES} for skill_slot in SKILL_SLOTS}
        
        for skill_slot in SKILL_SLOTS:
            for variable in SKILL_VARIABLES:
                try:
                    value = self.skill_data[skill_slot]["Variables"][variable].get()
                
                except Exception:
                    continue

                if isinstance(value, str):
                    if value.startswith("Select"):
                        continue

                skill_dictionaries[skill_slot][variable] = value

        build_dictionary.update(skill_dictionaries)

        gear_dictionaries = {gear_slot: {variable: None for variable in GEAR_VARIABLES} for gear_slot in GEAR_SLOTS}

        for gear_slot in GEAR_SLOTS:
            for variable in GEAR_VARIABLES:
                try:
                    value = self.gear_data[gear_slot]["Variables"][variable].get()
                
                except Exception:
                    continue

                if value.startswith("Select"):
                    continue

                gear_dictionaries[gear_slot][variable] = value

        build_dictionary.update(gear_dictionaries)

        return build_dictionary


    def save_build_to_file(self, build_dictionary, file_path):
        """
        Saves given build data to a specified file path in JSON format.

        Parameters:
            self: The instance of the DamageCalculatorApp class.
            build_dictionary: A dictionary containing the build's data.
            file_path: The file path where the build data will be saved.
        
        """

        with open(file_path, "w") as file:
            json.dump(build_dictionary, file, indent=4)










    #################### SECTION BREAK ####################

    ##### TAB SETUP METHODS #####


    def _process_names_for_specialization(self, names, specialization_value):
        """Processing names to indicate specialization requirements:"""
        processed_names = []
        original_names = []
        
        for name in names:
            original_names.append(name)
            
            if '(' in name and ')' in name:
                # Extract base name and specialization from LAST parentheses (for weapons):
                paren_start = name.rindex('(')
                paren_end = name.rindex(')')
                base_name = name[:paren_start].strip()
                paren_content = name[paren_start+1:paren_end].strip()
                
                if self.is_specialization_name(paren_content):
                    if specialization_value == paren_content:
                        processed_names.append(base_name)
                    else:
                        processed_names.append(f"{base_name} (Select {paren_content})")
                else:
                    processed_names.append(name)
            else:
                processed_names.append(name)
        
        return processed_names, original_names


    def _process_mods_for_specialization(self, mods, specialization_value):
        """Processing mod names to indicate specialization requirements:"""
        processed_mods = []
        
        for mod in mods:
            if '(' in mod and ')' in mod:
                # Extract base mod and specialization from LAST parentheses:
                paren_start = mod.rindex('(')
                paren_end = mod.rindex(')')
                base_mod = mod[:paren_start].strip()
                paren_content = mod[paren_start+1:paren_end].strip()
                
                if self.is_specialization_name(paren_content):
                    if specialization_value == paren_content:
                        processed_mods.append(base_mod)
                    else:
                        processed_mods.append(f"{base_mod} (Select {paren_content})")
                else:
                    processed_mods.append(mod)
            else:
                processed_mods.append(mod)
        
        return processed_mods
    

    def _update_skill_name_dropdown(self, skill_data, class_val, specialization_value):
        """Updating skill name dropdown to reflect specialization availability:"""
        if "Name" not in skill_data["Dropdowns"]:
            return
        
        variant_row = data["Skill"]["Variants"][data["Skill"]["Variants"]["Class"] == class_val].iloc[0]
        names_str = variant_row["Name"]
        all_names = [n.strip() for n in names_str.split(';')]
        
        # Processing names for display:
        processed_names, _ = self._process_names_for_specialization(all_names, specialization_value)
        
        # Updating dropdown with processed names:
        current_name = skill_data["Variables"]["Name"].get()
        if current_name != "Select Name":
            available_names = [n for n in processed_names if n != current_name]
        else:
            available_names = processed_names
        
        skill_data["Dropdowns"]["Name"].configure(values=available_names)


    def _reset_skill_if_specialization_mismatch(self, skill_data, specialization_value):
        """Checking and resetting skill if it requires a different specialization:"""
        name_val = skill_data["Variables"]["Name"].get()
        
        if name_val == "Select Name":
            return False
        
        # Checking if skill requires a different specialization:
        original_names = skill_data.get('original_names', [])
        
        for orig_name in original_names:
            if '(' in orig_name and ')' in orig_name:
                # Extract base name and specialization from LAST parentheses:
                base_name = orig_name[:orig_name.rindex('(')].strip()
                required_spec = orig_name[orig_name.rindex('(')+1:orig_name.rindex(')')].strip()
                
                if base_name == name_val and required_spec != specialization_value:
                    # Resetting skill selection:
                    skill_data["Variables"]["Name"].set("Select Name")
                    skill_data['prev_name'] = "Select Name"
                    
                    # Clearing skill mods:
                    for key in list(skill_data["Dropdowns"].keys()):
                        if key not in ["Class", "Name"]:
                            skill_data["Dropdowns"][key].grid_forget()
                            del skill_data["Dropdowns"][key]
                    for key in list(skill_data["Variables"].keys()):
                        if key not in ["Class", "Name", "Expertise"]:
                            del skill_data["Variables"][key]
                    
                    return True
        
        return False


    def _update_skill_mods_for_specialization(self, skill_data, class_val, specialization_value):
        """Updating skill mods to reflect specialization changes:"""
        original_mods_dict = skill_data.get('original_mods', {})
        
        # Resetting individual mods that require different specializations:
        for key, mod_var in list(skill_data["Variables"].items()):
            if key.startswith("Mod "):
                mod_val = mod_var.get()
                original_mods = original_mods_dict.get(key, [])
                
                for orig_mod in original_mods:
                    if '(' in orig_mod and ')' in orig_mod:
                        # Extract base mod and specialization from LAST parentheses:
                        base_mod = orig_mod[:orig_mod.rindex('(')].strip()
                        paren_content = orig_mod[orig_mod.rindex('(')+1:orig_mod.rindex(')')].strip()
                        
                        if self.is_specialization_name(paren_content):
                            if base_mod == mod_val and paren_content != specialization_value:
                                col_index = int(key.split('_')[1]) - 1
                                mods_sheet_name = f"{class_val} Mods"
                                if mods_sheet_name in data["Skill"]:
                                    mods_df = data["Skill"][mods_sheet_name]
                                    all_columns = mods_df.columns.tolist()
                                    mod_columns = [col for col in all_columns if col not in ['Stats', 'Last Checked']]
                                    if col_index < len(mod_columns):
                                        col_name = mod_columns[col_index]
                                        mod_var.set(f"Select {col_name} Mod")
                                        skill_data[f"Previous {key}"] = f"Select {col_name} Mod"
                                break
        
        # Updating skill mod dropdown labels:
        mods_sheet_name = f"{class_val} Mods"
        if mods_sheet_name not in data["Skill"]:
            return
        
        mods_df = data["Skill"][mods_sheet_name]
        all_columns = mods_df.columns.tolist()
        mod_columns = [col for col in all_columns if col not in ['Stats', 'Last Checked']]
        
        for i, col_name in enumerate(mod_columns):
            key = f"Mod {i+1}"
            if key in skill_data["Dropdowns"] and key in skill_data["Variables"]:
                all_mods = mods_df[mods_df[col_name] == '✓']['Stats'].tolist()
                
                processed_mods = self._process_mods_for_specialization(all_mods, specialization_value)
                
                current_mod = skill_data["Variables"][key].get()
                if current_mod not in [f"Select {col_name} Mod"]:
                    available_mods = [m for m in processed_mods if m != current_mod]
                else:
                    available_mods = processed_mods
                
                skill_data["Dropdowns"][key].configure(values=available_mods)


    def _update_skills_for_specialization(self, specialization_value):
        
        for skill_name, skill_data in self.skill_data.items():
            class_val = skill_data["Variables"]["Class"].get()
            
            # Skipping if no class selected or name variable doesn't exist:
            if class_val == "Select Class" or "Name" not in skill_data["Variables"]:
                continue
            
            # Updating skill name dropdown:
            self._update_skill_name_dropdown(skill_data, class_val, specialization_value)
            
            # Checking if skill needs reset:
            skill_was_reset = self._reset_skill_if_specialization_mismatch(skill_data, specialization_value)
            
            # Updating skill mods only if skill wasn't reset:
            if not skill_was_reset:
                self._update_skill_mods_for_specialization(skill_data, class_val, specialization_value)


    def _update_weapon_name_dropdown(self, class_val, specialization_value):
        """Updating weapon name dropdown to reflect specialization availability:"""
        if class_val == "Signature Weapons":
            weapon_sheet = data["Weapon"]["Signature Weapons"]
            all_names = [str(name) for name in weapon_sheet["Name"].tolist()]
        else:
            type_val = self.weapon_data["Variables"].get("Type", ctk.StringVar(value="Select Type")).get()
            if type_val == "Select Type":
                return
            weapon_sheet = data["Weapon"][class_val]
            all_names = [str(name) for name in weapon_sheet[weapon_sheet["Type"] == type_val]["Name"].tolist()]
        
        # Processing names for display:
        processed_names, original_names = self._process_names_for_specialization(all_names, specialization_value)
        
        # Storing original names for reference:
        if class_val == "Signature Weapons":
            self.weapon_data["Original Weapon Names"] = original_names
        
        # Updating dropdown with processed names:
        name_val = self.weapon_data["Variables"]["Name"].get()
        if name_val != "Select Name":
            available_names = [n for n in processed_names if n != name_val]
        else:
            available_names = processed_names
        
        self.weapon_data["Dropdowns"]["Name"].configure(values=available_names)


    def _reset_weapon_if_specialization_mismatch(self, specialization_value):
        """Checking and resetting weapon if it requires a different specialization:"""
        name_val = self.weapon_data["Variables"]["Name"].get()
        
        if name_val == "Select Name":
            return False
        
        # Checking if weapon requires a different specialization:
        original_weapon_names = self.weapon_data.get("Original Weapon Names", [])
        
        for orig_name in original_weapon_names:
            if '(' in orig_name and ')' in orig_name:
                # Extract base name and specialization from LAST parentheses (for weapons):
                paren_start = orig_name.rindex('(')
                paren_end = orig_name.rindex(')')
                base_name = orig_name[:paren_start].strip()
                paren_content = orig_name[paren_start+1:paren_end].strip()
                
                if self.is_specialization_name(paren_content):
                    if base_name == name_val and paren_content != specialization_value:
                        # Resetting weapon selection:
                        self.weapon_data["Variables"]["Name"].set("Select Name")
                        self.weapon_data["Previous Name"] = "Select Name"
                        self.weapon_data["Matched Weapon Name"] = "Select Name"
                        
                        # Clearing weapon attributes, mods, and talents:
                        for key in list(self.weapon_data["Dropdowns"].keys()):
                            if key not in ["Class", "Type", "Name"]:
                                self.weapon_data["Dropdowns"][key].grid_forget()
                                del self.weapon_data["Dropdowns"][key]
                        for key in list(self.weapon_data["Variables"].keys()):
                            if key not in ["Class", "Type", "Name"]:
                                del self.weapon_data["Variables"][key]
                        
                        return True
        
        return False


    def _update_weapon_mods_for_specialization(self, specialization_value):
        """Updating weapon mods to reflect specialization changes:"""
        original_weapon_mods_dict = self.weapon_data.get("Original Weapon Mods", {})
        
        # Resetting individual mods that require different specializations:
        for key, mod_var in list(self.weapon_data["Variables"].items()):
            if key in ["Optics Mod", "Magazine Mod", "Underbarrel Mod", "Muzzle Mod"]:
                mod_val = mod_var.get()
                original_mods = original_weapon_mods_dict.get(key, [])
                
                for orig_mod in original_mods:
                    if '(' in orig_mod and ')' in orig_mod:
                        # Extract base mod and specialization from LAST parentheses:
                        paren_start = orig_mod.rindex('(')
                        paren_end = orig_mod.rindex(')')
                        base_mod = orig_mod[:paren_start].strip()
                        paren_content = orig_mod[paren_start+1:paren_end].strip()
                        
                        if self.is_specialization_name(paren_content):
                            if base_mod == mod_val and paren_content != specialization_value:
                                mod_var.set(f"Select {key}")
                                self.weapon_data[f"Previous {key}"] = f"Select {key}"
                                break
        
        # Updating weapon mod dropdown labels:
        name_val = self.weapon_data["Variables"]["Name"].get()
        class_val = self.weapon_data["Variables"]["Class"].get()
        
        if name_val == "Select Name" or "Matched Weapon Name" not in self.weapon_data:
            return
        
        weapon_sheet = data["Weapon"][class_val]
        type_val = self.weapon_data["Variables"].get("Type", ctk.StringVar(value="Select Type")).get()
        
        if type_val == "Select Type":
            return
        
        lookup_name = self.weapon_data.get("Matched Weapon Name", name_val)
        
        if lookup_name == "Select Name":
            return
        
        weapon_rows = weapon_sheet[(weapon_sheet["Name"] == lookup_name) & (weapon_sheet["Type"] == type_val)]
        if weapon_rows.empty:
            return
        
        weapon_row = weapon_rows.iloc[0]
        
        mod_types = [
            ("Optics Mod", "Optics Mods"),
            ("Magazine Mod", "Magazine Mods"),
            ("Underbarrel Mod", "Underbarrel Mods"),
            ("Muzzle Mod", "Muzzle Mods")
        ]
        
        for i, (column_name, sheet_name) in enumerate(mod_types, start=1):
            if column_name in self.weapon_data["Dropdowns"] and column_name in self.weapon_data["Variables"]:
                mod_cell = weapon_row.get(column_name, pd.NA)
                if not pd.isna(mod_cell) and isinstance(mod_cell, str) and mod_cell.startswith("*"):
                    mod_rail = mod_cell[1:]
                    
                    if sheet_name in data["Weapon"]:
                        mods_df = data["Weapon"][sheet_name]
                        if mod_rail in mods_df.columns:
                            all_mods = mods_df[mods_df[mod_rail] == '✓']['Stats'].tolist()
                            
                            processed_mods = self._process_mods_for_specialization(all_mods, specialization_value)
                            
                            current_mod = self.weapon_data["Variables"][column_name].get()
                            if current_mod not in [f"Select {column_name}"]:
                                available_mods = [m for m in processed_mods if m != current_mod]
                            else:
                                available_mods = processed_mods
                            
                            self.weapon_data["Dropdowns"][column_name].configure(values=available_mods)


    def _update_weapon_for_specialization(self, specialization_value):

        if not hasattr(self, 'weapon_data'):
            return
        
        class_val = self.weapon_data["Variables"]["Class"].get()
        
        if class_val == "Select Class":
            return
        
        # Updating weapon name dropdown if it exists:
        if "Name" in self.weapon_data["Dropdowns"]:
            self._update_weapon_name_dropdown(class_val, specialization_value)
        
        # Checking if selected weapon needs reset:
        if "Name" in self.weapon_data["Variables"]:
            weapon_was_reset = self._reset_weapon_if_specialization_mismatch(specialization_value)
            
            # Updating weapon mods only if weapon wasn't reset:
            if not weapon_was_reset:
                self._update_weapon_mods_for_specialization(specialization_value)


    def update_specialization(self, specialization_variable, specialization_dropdown):
        
        # Retrieving the current specialization selection:
        specialization_value = specialization_variable.get()

        # Updating the specialization dropdown to exclude the current selection:
        specialization_dropdown.configure(values=[specialization for specialization in SPECIALIZATIONS if specialization != specialization_value])
        
        # Propagating specialization changes to skills and weapon:
        self._update_skills_for_specialization(specialization_value)
        self._update_weapon_for_specialization(specialization_value)

    
    def update_all_type_dropdowns(self):
        """Update all gear type dropdowns to handle Exotic exclusivity"""
        if not hasattr(self, 'gear_sections'):
            return
        
        # Check if any gear piece has Exotic selected
        has_exotic = any(
            gear_data["Variables"]["Type"].get() == "Exotic" 
            for gear_data in self.gear_data.values()
        )
        
        # Determine available types
        if has_exotic:
            base_types = ["Improvised", "Brand Set", "Gear Set", "Named"]
        else:
            base_types = ["Improvised", "Brand Set", "Gear Set", "Named", "Exotic"]
        
        # Update each gear's type dropdown
        for gear_name, gear_data in self.gear_data.items():
            current_type = gear_data["Variables"]["Type"].get()
            type_dropdown = gear_data["Dropdowns"]["Type"]
            
            # Start with base types (already accounts for Exotic exclusivity)
            # Then exclude the current selection from the available options
            available_types = [t for t in base_types if t != current_type]
            
            # Update the dropdown values
            type_dropdown.configure(values=available_types)
    
    def update_gear(self, gear_name, gear_data, trigger_point="Type", preserve_selections=True):
        """Dynamically update gear dropdowns based on selections
        
        Args:
            gear_name: Name of the gear slot (e.g., 'Mask', 'Holster')
            gear_data: Dictionary containing gear state and UI elements
            trigger_point: Which dropdown triggered this update ("Type" or "Name")
            preserve_selections: Whether to preserve existing attribute/mod selections
        """
        type_val = gear_data["Variables"]["Type"].get()
        frame = gear_data["Frame"]
        
        # Update ALL type dropdowns to handle Exotic exclusivity
        self.update_all_type_dropdowns()
        
        # Handle type selection (trigger_point == "Type")
        if trigger_point == "Type":
            # Clear existing dropdowns except type
            for key in list(gear_data["Dropdowns"].keys()):
                if key != "Type":
                    gear_data["Dropdowns"][key].grid_forget()
                    del gear_data["Dropdowns"][key]
            for key in list(gear_data["Variables"].keys()):
                if key != "Type":
                    del gear_data["Variables"][key]
            
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
                    command=lambda choice: self.update_gear(gear_name, gear_data, trigger_point="Name"),
                    state="readonly"
                )
            
            name_dropdown.grid(row=row_idx, column=0, padx=5, pady=5, sticky="ew")
            gear_data["Variables"]["Name"] = name_var
            gear_data["Dropdowns"]["Name"] = name_dropdown
            row_idx += 1
            
            # If Improvised, automatically update attributes
            if type_val == "Improvised":
                self._update_gear_attributes(gear_name, gear_data, preserve_selections=False)
        
        # Handle name selection (trigger_point == "Name")
        elif trigger_point == "Name":
            self._update_gear_attributes(gear_name, gear_data, preserve_selections)
    
    def _update_gear_attributes(self, gear_name, gear_data, preserve_selections=True):
        """Internal helper: Update core attributes, attributes, mods, and talents based on gear selection"""
        type_val = gear_data["Variables"]["Type"].get()
        name_val = gear_data["Variables"]["Name"].get()
        frame = gear_data["Frame"]
        
        if name_val in ["Select Name", "Select Type"]:
            return
        
        # Preserve existing selections if requested
        saved_selections = {}
        if preserve_selections:
            for key, var in gear_data["Variables"].items():
                if key not in ["Type", "Name"]:
                    saved_selections[key] = var.get()
        
        # Clear existing attribute dropdowns
        for key in list(gear_data["Dropdowns"].keys()):
            if key not in ["Type", "Name"]:
                gear_data["Dropdowns"][key].grid_forget()
                del gear_data["Dropdowns"][key]
        for key in list(gear_data["Variables"].keys()):
            if key not in ["Type", "Name"]:
                del gear_data["Variables"][key]
        
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
        all_cores = GEAR_CORE_ATTRIBUTES
        for i, cell_val in enumerate(core_cells):
            key = f"Core Attribute {i+1}"
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
            key = f"Attribute {i+1}"
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
            key = f"Mod {i+1}"
            if key in saved_selections and saved_selections[key] not in ["Select Mod"]:
                if pd.isna(cell_val):
                    saved_selections[key] = "Select Mod"
                elif saved_selections[key] not in all_mods:
                    saved_selections[key] = "Select Mod"
        
        # Validate talents
        for i, cell_val in enumerate(talent_cells):
            key = f"Talent {i+1}"
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
            key = f"Core Attribute {i+1}"
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
            key = f"Attribute {i+1}"
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
            key = f"Core Attribute {i+1}"
            saved_value = final_core_values[key]
            core_var = ctk.StringVar(value=saved_value)
            
            if cell_val == "*":
                # User can select any core attribute
                # Exclude final values from OTHER core slots (not this one)
                current_selections = [final_core_values[f"Core Attribute {j + 1}"] 
                                     for j in range(len(core_cells)) 
                                     if j != i and f"Core Attribute {j + 1}" in final_core_values 
                                     and final_core_values[f"Core Attribute {j + 1}"] not in ["Select Core Attribute"]]
                # Also exclude this dropdown's own current selection
                if saved_value not in ["Select Core Attribute"]:
                    current_selections.append(saved_value)
                available_cores = [c for c in GEAR_CORE_ATTRIBUTES 
                                  if c not in current_selections]
                core_dropdown = ctk.CTkComboBox(
                    frame, 
                    values=available_cores, 
                    variable=core_var,
                    command=lambda choice, gear=gear_name, gdata=gear_data: self._update_gear_attributes(gear, gdata, preserve_selections=True),
                    state="readonly"
                )
            elif isinstance(cell_val, str) and cell_val.startswith("!"):
                # Cannot select this specific one
                excluded = cell_val[1:]
                current_selections = [final_core_values[f"Core Attribute {j + 1}"] 
                                     for j in range(len(core_cells)) 
                                     if j != i and f"Core Attribute {j + 1}" in final_core_values 
                                     and final_core_values[f"Core Attribute {j + 1}"] not in ["Select Core Attribute"]]
                # Also exclude this dropdown's own current selection
                if saved_value not in ["Select Core Attribute"]:
                    current_selections.append(saved_value)
                available_cores = [c for c in GEAR_CORE_ATTRIBUTES 
                                  if c != excluded and c not in current_selections]
                core_dropdown = ctk.CTkComboBox(
                    frame, 
                    values=available_cores, 
                    variable=core_var,
                    command=lambda choice, gear=gear_name, gdata=gear_data: self._update_gear_attributes(gear, gdata, preserve_selections=True),
                    state="readonly"
                )
            else:
                # Fixed core attribute
                core_var.set(cell_val)
                core_dropdown = ctk.CTkComboBox(frame, values=[cell_val], variable=core_var, state="disabled")
            
            core_dropdown.grid(row=row_idx, column=0, padx=5, pady=5, sticky="ew")
            gear_data["Variables"][f"Core Attribute {i+1}"] = core_var
            gear_data["Dropdowns"][f"Core Attribute {i+1}"] = core_dropdown
            row_idx += 1
        
        # Attributes
        for i, cell_val in enumerate(attr_cells):
            if pd.isna(cell_val):
                continue
            
            # Use final value for this slot
            key = f"Attribute {i+1}"
            saved_value = final_attr_values[key]
            attr_var = ctk.StringVar(value=saved_value)
            
            if cell_val == "*":
                # User can select any attribute
                # Exclude final values from OTHER attribute slots (not this one)
                current_selections = [final_attr_values[f"Attribute {j + 1}"] 
                                     for j in range(len(attr_cells)) 
                                     if j != i and f"Attribute {j + 1}" in final_attr_values 
                                     and final_attr_values[f"Attribute {j + 1}"] not in ["Select Attribute"]]
                # Also exclude this dropdown's own current selection
                if saved_value not in ["Select Attribute"]:
                    current_selections.append(saved_value)
                available_attrs = [a for a in data["Armor"]["Attributes"]["Stats"].tolist() 
                                  if a not in current_selections]
                attr_dropdown = ctk.CTkComboBox(
                    frame, 
                    values=available_attrs, 
                    variable=attr_var,
                    command=lambda choice, gear=gear_name, gdata=gear_data: self._update_gear_attributes(gear, gdata, preserve_selections=True),
                    state="readonly"
                )
            elif isinstance(cell_val, str) and cell_val.startswith("*") and len(cell_val) > 1:
                # Specific type of attribute (e.g., "*Offensive")
                attr_type = cell_val[1:]
                current_selections = [final_attr_values[f"Attribute {j + 1}"] 
                                     for j in range(len(attr_cells)) 
                                     if j != i and f"Attribute {j + 1}" in final_attr_values 
                                     and final_attr_values[f"Attribute {j + 1}"] not in ["Select Attribute"]]
                # Also exclude this dropdown's own current selection
                if saved_value not in ["Select Attribute"]:
                    current_selections.append(saved_value)
                available_attrs = data["Armor"]["Attributes"][data["Armor"]["Attributes"]["Type"] == attr_type]["Stats"].tolist()
                available_attrs = [a for a in available_attrs if a not in current_selections]
                attr_dropdown = ctk.CTkComboBox(
                    frame, 
                    values=available_attrs, 
                    variable=attr_var,
                    command=lambda choice, gear=gear_name, gdata=gear_data: self._update_gear_attributes(gear, gdata, preserve_selections=True),
                    state="readonly"
                )
            elif isinstance(cell_val, str) and cell_val.startswith("!"):
                # Cannot select this specific one
                excluded = cell_val[1:]
                current_selections = [final_attr_values[f"Attribute {j + 1}"] 
                                     for j in range(len(attr_cells)) 
                                     if j != i and f"Attribute {j + 1}" in final_attr_values 
                                     and final_attr_values[f"Attribute {j + 1}"] not in ["Select Attribute"]]
                # Also exclude this dropdown's own current selection
                if saved_value not in ["Select Attribute"]:
                    current_selections.append(saved_value)
                available_attrs = [a for a in data["Armor"]["Attributes"]["Stats"].tolist() 
                                  if a != excluded and a not in current_selections]
                attr_dropdown = ctk.CTkComboBox(
                    frame, 
                    values=available_attrs, 
                    variable=attr_var,
                    command=lambda choice, gear=gear_name, gdata=gear_data: self._update_gear_attributes(gear, gdata, preserve_selections=True),
                    state="readonly"
                )
            else:
                # Fixed attribute
                attr_var.set(cell_val)
                attr_dropdown = ctk.CTkComboBox(frame, values=[cell_val], variable=attr_var, state="disabled")
            
            attr_dropdown.grid(row=row_idx, column=0, padx=5, pady=5, sticky="ew")
            gear_data["Variables"][f"Attribute {i+1}"] = attr_var
            gear_data["Dropdowns"][f"Attribute {i+1}"] = attr_dropdown
            row_idx += 1
        
        # Mods
        for i, cell_val in enumerate(mod_cells):
            if pd.isna(cell_val):
                continue
            
            saved_value = saved_selections.get(f"Mod {i+1}", "Select Mod")
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
                command=lambda choice, gear=gear_name, gdata=gear_data: self._update_gear_attributes(gear, gdata, preserve_selections=True),
                state="readonly"
            )
            mod_dropdown.grid(row=row_idx, column=0, padx=5, pady=5, sticky="ew")
            gear_data["Variables"][f"Mod {i+1}"] = mod_var
            gear_data["Dropdowns"][f"Mod {i+1}"] = mod_dropdown
            row_idx += 1
        
        # Talents
        for i, cell_val in enumerate(talent_cells):
            if pd.isna(cell_val):
                continue
            
            saved_value = saved_selections.get(f"Talent {i+1}", "Select Talent")
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
                    command=lambda choice, gear=gear_name, gdata=gear_data: self._update_gear_attributes(gear, gdata, preserve_selections=True),
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
            gear_data["Variables"][f"Talent {i+1}"] = talent_var
            gear_data["Dropdowns"][f"Talent {i+1}"] = talent_dropdown
            row_idx += 1
    
    
    def is_specialization_name(self, text):
        """Check if the given text is a valid specialization name"""
        return text in SPECIALIZATIONS
    
    def update_all_skill_class_dropdowns(self):
        """Update all skill class dropdowns to handle class exclusivity"""
        if not hasattr(self, 'skill_sections'):
            return
        
        # Get all base classes
        all_classes = data["Skill"]["Variants"]["Class"].tolist()
        
        # Update each skill's class dropdown
        for skill_name, skill_data in self.skill_data.items():
            current_class = skill_data["Variables"]["Class"].get()
            class_dropdown = skill_data["Dropdowns"]["Class"]
            
            # Start with all classes
            available_classes = all_classes.copy()
            
            # Exclude classes selected in OTHER skills
            for other_skill_name, other_skill_data in self.skill_data.items():
                if other_skill_name != skill_name:
                    other_class = other_skill_data["Variables"]["Class"].get()
                    if other_class != "Select Class" and other_class in available_classes:
                        available_classes.remove(other_class)
            
            # Also exclude this skill's current selection
            if current_class != "Select Class" and current_class in available_classes:
                available_classes.remove(current_class)
            
            # Update the dropdown values
            class_dropdown.configure(values=available_classes)
    
    def update_skill(self, skill_name, skill_data, trigger_point="Class", preserve_selections=True):
        """Dynamically update skill dropdowns based on selections
        
        Args:
            skill_name: Name of the skill slot (e.g., 'Skill 1')
            skill_data: Dictionary containing skill state and UI elements
            trigger_point: Which dropdown triggered this update ("Class", "Name", or "Mod")
            preserve_selections: Whether to preserve existing mod selections
        """
        class_val = skill_data["Variables"]["Class"].get()
        frame = skill_data["Frame"]
        
        # Update ALL skill class dropdowns to handle exclusivity
        self.update_all_skill_class_dropdowns()
        
        if class_val == "Select Class":
            return
        
        # Handle class selection (trigger_point == "Class")
        if trigger_point == "Class":
            # Clear existing dropdowns except class
            for key in list(skill_data["Dropdowns"].keys()):
                if key != "Class":
                    skill_data["Dropdowns"][key].grid_forget()
                    del skill_data["Dropdowns"][key]
            for key in list(skill_data["Variables"].keys()):
                if key not in ["Class", "Expertise"]:
                    del skill_data["Variables"][key]
            
            row_idx = 1
            
            # Name dropdown - get names from Variants sheet
            variant_row = data["Skill"]["Variants"][data["Skill"]["Variants"]["Class"] == class_val].iloc[0]
            names_str = variant_row["Name"]
            names = [n.strip() for n in names_str.split(';')]
            
            # Process names to handle specialization requirements
            processed_names = []
            spec_val = self.specialization_variable.get() if hasattr(self, 'specialization_variable') else "Select Specialization"
            
            for name in names:
                # Check if name has specialization requirement (contains parentheses)
                if '(' in name and ')' in name:
                    # Extract base name and required specialization from LAST parentheses:
                    base_name = name[:name.rindex('(')].strip()
                    required_spec = name[name.rindex('(')+1:name.rindex(')')].strip()
                    
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
                command=lambda choice: self.update_skill(skill_name, skill_data, trigger_point="Name"),
                state="readonly"
            )
            name_dropdown.grid(row=row_idx, column=0, padx=5, pady=5, sticky="ew")
            skill_data["Variables"]["Name"] = name_var
            skill_data["Dropdowns"]["Name"] = name_dropdown
            
            # Store original names for reference
            skill_data['original_names'] = names
            return
        
        # Handle name selection (trigger_point == "Name")
        elif trigger_point == "Name":
            self._handle_skill_name_selection(skill_name, skill_data)
            return
        
        # Handle mod selection (trigger_point == "Mod")
        elif trigger_point == "Mod":
            self._handle_skill_mod_selection(skill_name, skill_data)
            return
    
    def _handle_skill_name_selection(self, skill_name, skill_data):
        """Internal helper: Handle skill name selection and validate specialization requirements"""
        name_val = skill_data["Variables"]["Name"].get()
        class_val = skill_data["Variables"]["Class"].get()
        
        if name_val == "Select Name":
            return
        
        # Check if this is a disabled specialization-locked skill
        if "(Select " in name_val:
            # Revert to previous selection
            prev_name = skill_data.get('prev_name', "Select Name")
            skill_data["Variables"]["Name"].set(prev_name)
            return
        
        # Store the current selection for potential revert
        skill_data['prev_name'] = name_val
        
        # Continue to update mods
        self._update_skill_mods(skill_name, skill_data)
    
    def _update_skill_mods(self, skill_name, skill_data, preserve_selections=True):
        """Internal helper: Update skill mod dropdowns based on class selection"""
        class_val = skill_data["Variables"]["Class"].get()
        name_val = skill_data["Variables"]["Name"].get()
        frame = skill_data["Frame"]
        
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
                # Extract base name from LAST parentheses:
                base_name = orig_name[:orig_name.rindex('(')].strip()
                if base_name == name_val:
                    matched_name = orig_name
                    break
        
        # Update name dropdown to exclude current selection
        variant_row = data["Skill"]["Variants"][data["Skill"]["Variants"]["Class"] == class_val].iloc[0]
        names_str = variant_row["Name"]
        all_names = [n.strip() for n in names_str.split(';')]
        
        # Process names for display (same logic as in update_skill)
        spec_val = self.specialization_variable.get() if hasattr(self, 'specialization_variable') else "Select Specialization"
        processed_names = []
        for name in all_names:
            if '(' in name and ')' in name:
                # Extract base name and required spec from LAST parentheses:
                base_name = name[:name.rindex('(')].strip()
                required_spec = name[name.rindex('(')+1:name.rindex(')')].strip()
                if spec_val == required_spec:
                    processed_names.append(base_name)
                else:
                    processed_names.append(f"{base_name} (Select {required_spec})")
            else:
                processed_names.append(name)
        
        # Exclude current selection from processed names
        available_names = [n for n in processed_names if n != name_val]
        skill_data["Dropdowns"]["Name"].configure(values=available_names)
        
        # Preserve existing selections if requested
        saved_selections = {}
        if preserve_selections:
            for key, var in skill_data["Variables"].items():
                if key not in ["Class", "Name"]:
                    saved_selections[key] = var.get()
        
        # Clear existing mod dropdowns
        for key in list(skill_data["Dropdowns"].keys()):
            if key not in ["Class", "Name"]:
                skill_data["Dropdowns"][key].grid_forget()
                del skill_data["Dropdowns"][key]
        for key in list(skill_data["Variables"].keys()):
            if key not in ["Class", "Name", "Expertise"]:
                del skill_data["Variables"][key]
        
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
            key = f"Mod {i+1}"
            final_mod_values[key] = saved_selections.get(key, f"Select {col_name} Mod")
        
        row_idx = 2
        
        # Create mod dropdowns
        for i, col_name in enumerate(mod_columns):
            key = f"Mod {i+1}"
            saved_value = final_mod_values[key]
            
            mod_var = ctk.StringVar(value=saved_value)
            
            # Get available mods for this column (where column value is '✓')
            all_mods = mods_df[mods_df[col_name] == '✓']['Stats'].tolist()
            
            # Process mods to handle specialization requirements
            spec_val = self.specialization_variable.get() if hasattr(self, 'specialization_variable') else "Select Specialization"
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
                command=lambda choice, sname=skill_name, sdata=skill_data: self._handle_skill_mod_selection(sname, sdata),
                state="readonly"
            )
            mod_dropdown.grid(row=row_idx, column=0, padx=5, pady=5, sticky="ew")
            skill_data["Variables"][key] = mod_var
            skill_data["Dropdowns"][key] = mod_dropdown
            
            # Store original mods and column name for reference
            if 'original_mods' not in skill_data:
                skill_data['original_mods'] = {}
            skill_data['original_mods'][key] = original_mods
            
            if 'mod_column_names' not in skill_data:
                skill_data['mod_column_names'] = {}
            skill_data['mod_column_names'][key] = col_name
            
            row_idx += 1
    
    def _handle_skill_mod_selection(self, skill_name, skill_data):
        """Internal helper: Handle skill mod selection and validate specialization requirements"""
        # Check if any selected mod is locked due to specialization
        for key, mod_var in skill_data["Variables"].items():
            if key.startswith("Mod "):
                mod_val = mod_var.get()
                
                # Check if this is a disabled specialization-locked mod
                if "(Select " in mod_val:
                    # Determine the correct default text based on column name
                    col_name = skill_data.get('mod_column_names', {}).get(key, "Mod")
                    prev_mod = skill_data.get(f"Previous {key}", f"Select {col_name} Mod")
                    mod_var.set(prev_mod)
                    return
                
                # Store the current selection for potential revert
                skill_data[f"Previous {key}"] = mod_val
        
        # Refresh mods to update exclusion lists
        self._update_skill_mods(skill_name, skill_data, preserve_selections=True)

    
    def update_weapon_class_options(self):
        """Update weapon class dropdown to exclude current selection"""
        if not hasattr(self, 'weapon_data'):
            return
        
        current_class = self.weapon_data["Variables"]["Class"].get()
        class_dropdown = self.weapon_data["Dropdowns"]["Class"]
        
        all_classes = ["Assault Rifles", "Light Machineguns", "Marksman Rifles", "Pistols", "Rifles", "Shotguns", "Sub Machine Guns", "Signature Weapons"]
        
        # Exclude current selection
        available_classes = [c for c in all_classes if c != current_class]
        
        class_dropdown.configure(values=available_classes)
    
    def update_weapon(self, weapon_name, weapon_data, trigger_point="Class"):
        """Dynamically update weapon dropdowns based on selections
        
        Args:
            weapon_name: Name of the weapon slot (e.g., 'Weapon')
            weapon_data: Dictionary containing weapon state and UI elements
            trigger_point: Which dropdown triggered this update ("Class", "Type", "Name", "Mod", or 'attribute')
        """
        class_val = weapon_data["Variables"]["Class"].get()
        frame = weapon_data["Frame"]
        
        if class_val == "Select Class":
            return
        
        # Update class dropdown options
        self.update_weapon_class_options()
        
        # Handle class selection (trigger_point == "Class")
        if trigger_point == "Class":
            # Check if class changed - if so, clear name and subsequent dropdowns but preserve type
            prev_class = weapon_data.get("Previous Class", None)
            if prev_class is not None and prev_class != class_val:
                # Class changed - clear name and everything below it
                keys_to_clear = ["Name", "Core Attribute 1", "Core Attribute 2", "Attribute", "Optics Mod", "Magazine Mod", "Underbarrel Mod", "Muzzle Mod", "Talent 1", "Talent 2"]
                for key in keys_to_clear:
                    if key in weapon_data["Dropdowns"]:
                        weapon_data["Dropdowns"][key].grid_forget()
                        del weapon_data["Dropdowns"][key]
                    if key in weapon_data["Variables"]:
                        del weapon_data["Variables"][key]
            weapon_data["Previous Class"] = class_val
            
            # Handle Signature Weapons (no type dropdown needed)
            if class_val == "Signature Weapons":
                # Clear type dropdown if it exists
                if "Type" in weapon_data["Dropdowns"]:
                    weapon_data["Dropdowns"]["Type"].grid_forget()
                    del weapon_data["Dropdowns"]["Type"]
                    del weapon_data["Variables"]["Type"]
                # Proceed directly to name creation
                self._update_weapon_name(weapon_name, weapon_data)
                return
            
            # Create type dropdown if it doesn't exist yet
            if "Type" not in weapon_data["Dropdowns"]:
                row_idx = 1
                type_var = ctk.StringVar(value="Select Type")
                type_values = ["High-End", "Named", "Exotic"]
                type_dropdown = ctk.CTkComboBox(
                    frame,
                    values=type_values,
                    variable=type_var,
                    command=lambda choice: self.update_weapon(weapon_name, weapon_data, trigger_point="Type"),
                    state="readonly"
                )
                type_dropdown.grid(row=row_idx, column=0, padx=5, pady=5, sticky="ew")
                weapon_data["Variables"]["Type"] = type_var
                weapon_data["Dropdowns"]["Type"] = type_dropdown
                return  # Wait for type selection
            else:
                # Type exists - check if selected
                if "Type" in weapon_data["Variables"]:
                    type_val = weapon_data["Variables"]["Type"].get()
                    if type_val != "Select Type":
                        self._update_weapon_name(weapon_name, weapon_data)
            return
        
        # Handle type selection (trigger_point == "Type")
        elif trigger_point == "Type":
            if "Type" in weapon_data["Variables"]:
                type_val = weapon_data["Variables"]["Type"].get()
                if type_val != "Select Type":
                    self._update_weapon_name(weapon_name, weapon_data)
            return
        
        # Handle name selection (trigger_point == "Name")
        elif trigger_point == "Name":
            self._handle_weapon_name_selection(weapon_name, weapon_data)
            return
        
        # Handle mod selection (trigger_point == "Mod")
        elif trigger_point == "mod":
            self._handle_weapon_mod_selection(weapon_name, weapon_data)
            return
        
        # Handle attribute selection (trigger_point == 'attribute')
        elif trigger_point == "attribute":
            self._update_weapon_attributes(weapon_name, weapon_data, preserve_selections=True)
            return
    
    def _update_weapon_name(self, weapon_name, weapon_data):
        """Internal helper: Update weapon name dropdown based on class and type"""
        class_val = weapon_data["Variables"]["Class"].get()
        frame = weapon_data["Frame"]
        
        # Clear existing name dropdown and everything below it (core attributes, attribute, mods, talents)
        keys_to_clear = ["Name", "Core Attribute 1", "Core Attribute 2", "Attribute", "Optics Mod", "Magazine Mod", "Underbarrel Mod", "Muzzle Mod", "Talent 1", "Talent 2"]
        for key in keys_to_clear:
            if key in weapon_data["Dropdowns"]:
                weapon_data["Dropdowns"][key].grid_forget()
                del weapon_data["Dropdowns"][key]
            if key in weapon_data["Variables"]:
                del weapon_data["Variables"][key]
        
        row_idx = 2 if "Type" in weapon_data["Dropdowns"] else 1
        
        name_var = ctk.StringVar()
        
        # Get names from the class sheet
        if class_val == "Signature Weapons":
            # Signature Weapons don't have a Type column - get all names directly
            weapon_sheet = data["Weapon"]["Signature Weapons"]
            all_names = [str(name) for name in weapon_sheet["Name"].tolist()]
        else:
            # Get names from the class sheet filtered by type
            type_val = weapon_data["Variables"]["Type"].get()
            if type_val == "Select Type":
                return
            
            # Update type dropdown to exclude current selection
            all_types = ["High-End", "Named", "Exotic"]
            available_types = [t for t in all_types if t != type_val]
            weapon_data["Dropdowns"]["Type"].configure(values=available_types)
            
            # Get weapon names and convert to strings
            weapon_sheet = data["Weapon"][class_val]
            all_names = [str(name) for name in weapon_sheet[weapon_sheet["Type"] == type_val]["Name"].tolist()]
        
        # Process names to handle specialization requirements
        spec_val = self.specialization_variable.get() if hasattr(self, 'specialization_variable') else "Select Specialization"
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
        weapon_data["Original Weapon Names"] = original_weapon_names
        
        name_var.set("Select Name")
        name_dropdown = ctk.CTkComboBox(
            frame,
            values=processed_names,
            variable=name_var,
                command=lambda choice: self.update_weapon(weapon_name, weapon_data, trigger_point="Name"),
                state="readonly"
            )
        
        name_dropdown.grid(row=row_idx, column=0, padx=5, pady=5, sticky="ew")
        weapon_data["Variables"]["Name"] = name_var
        weapon_data["Dropdowns"]["Name"] = name_dropdown
    
    def _handle_weapon_name_selection(self, weapon_name, weapon_data):
        """Internal helper: Handle weapon name selection to exclude current selection from list"""
        name_val = weapon_data["Variables"]["Name"].get()
        class_val = weapon_data["Variables"]["Class"].get()
        
        if name_val == "Select Name":
            return
        
        # Check if this is a disabled specialization-locked weapon
        if "(Select " in name_val:
            # Revert to previous selection
            prev_name = weapon_data.get("Previous Name", "Select Name")
            weapon_data["Variables"]["Name"].set(prev_name)
            return
        
        # Store the current selection for potential revert
        weapon_data["Previous Name"] = name_val
        
        # For Signature Weapons, update dropdown and create talent dropdown
        if class_val == "Signature Weapons":
            # Match the selected name against original names
            original_weapon_names = weapon_data.get("Original Weapon Names", [])
            matched_name = name_val
            
            # Try to find matching original name
            for orig_name in original_weapon_names:
                if orig_name == name_val:
                    matched_name = orig_name
                    break
                elif '(' in orig_name and ')' in orig_name:
                    paren_start = orig_name.rindex('(')
                    paren_end = orig_name.rindex(')')
                    base_name = orig_name[:paren_start].strip()
                    paren_content = orig_name[paren_start+1:paren_end].strip()
                    
                    if self.is_specialization_name(paren_content) and base_name == name_val:
                        matched_name = orig_name
                        break
            
            # Update weapon_data to use the matched name for lookups
            weapon_data["Matched Weapon Name"] = matched_name
            
            # Get all signature weapon names
            weapon_sheet = data["Weapon"]["Signature Weapons"]
            all_names = [str(name) for name in weapon_sheet["Name"].tolist()]
            
            # Process names for display
            spec_val = self.specialization_variable.get() if hasattr(self, 'specialization_variable') else "Select Specialization"
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
            weapon_data["Dropdowns"]["Name"].configure(values=available_names)
            
            # Create weapon attributes
            self._update_weapon_attributes(weapon_name, weapon_data)
            return
        
        # For non-Signature Weapons, get the type
        type_val = weapon_data["Variables"]["Type"].get()
        
        # Match the selected name against original names
        original_weapon_names = weapon_data.get("Original Weapon Names", [])
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
        weapon_data["Matched Weapon Name"] = matched_name
        
        # Get all names for this class and type, converted to strings
        weapon_sheet = data["Weapon"][class_val]
        all_names = [str(name) for name in weapon_sheet[weapon_sheet["Type"] == type_val]["Name"].tolist()]
        
        # Process names for display
        spec_val = self.specialization_variable.get() if hasattr(self, 'specialization_variable') else "Select Specialization"
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
        weapon_data["Dropdowns"]["Name"].configure(values=available_names)
        
        # Create core attributes and attribute dropdowns
        self._update_weapon_attributes(weapon_name, weapon_data)
    
    def _handle_weapon_mod_selection(self, weapon_name, weapon_data):
        """Internal helper: Handle weapon mod selection and validate specialization requirements"""
        # Check if any selected mod is locked due to specialization
        for key, mod_var in weapon_data["Variables"].items():
            if key in ["Optics Mod", "Magazine Mod", "Underbarrel Mod", "Muzzle Mod"]:
                mod_val = mod_var.get()
                
                # Check if this is a disabled specialization-locked mod
                if "(Select " in mod_val:
                    # Revert to previous selection
                    prev_mod = weapon_data.get(f"Previous {key}", f"Select {key}")
                    mod_var.set(prev_mod)
                    return
                
                # Store the current selection for potential revert
                weapon_data[f"Previous {key}"] = mod_val
        
        # Refresh weapon attributes to update exclusion lists
        self._update_weapon_attributes(weapon_name, weapon_data, preserve_selections=True)
    
    def _update_weapon_attributes(self, weapon_name, weapon_data, preserve_selections=True):
        """Internal helper: Update weapon core attributes and attribute based on name selection"""
        class_val = weapon_data["Variables"]["Class"].get()
        name_val = weapon_data["Variables"]["Name"].get()
        frame = weapon_data["Frame"]
        
        if name_val == "Select Name":
            return
        
        # Preserve existing selections if requested
        saved_selections = {}
        if preserve_selections:
            for key, var in weapon_data["Variables"].items():
                if key not in ["Class", "Type", "Name", "Expertise"]:
                    saved_selections[key] = var.get()
        
        # Clear existing attribute dropdowns
        for key in list(weapon_data["Dropdowns"].keys()):
            if key not in ["Class", "Type", "Name"]:
                weapon_data["Dropdowns"][key].grid_forget()
                del weapon_data["Dropdowns"][key]
        for key in list(weapon_data["Variables"].keys()):
            if key not in ["Class", "Type", "Name", "Expertise"]:
                del weapon_data["Variables"][key]
        
        row_idx = 3  # After class, type, name
        
        # Handle Signature Weapons separately (they only have a talent, no cores/attributes/mods)
        if class_val == "Signature Weapons":
            # Match the displayed name back to the original name in the sheet
            original_weapon_names = weapon_data.get("Original Weapon Names", [])
            lookup_name = name_val
            
            # Find the original name (with specialization suffix)
            for orig_name in original_weapon_names:
                if '(' in orig_name and ')' in orig_name:
                    base_name = orig_name[:orig_name.rindex('(')].strip()
                    if base_name == name_val:
                        lookup_name = orig_name
                        break
                elif orig_name == name_val:
                    lookup_name = orig_name
                    break
            
            sig_weapons = data["Weapon"]["Signature Weapons"]
            sig_row = sig_weapons[sig_weapons["Name"] == lookup_name]
            if not sig_row.empty:
                talent_val = sig_row.iloc[0].get("Talent 1", pd.NA)
                if not pd.isna(talent_val):
                    talent_var = ctk.StringVar(value=talent_val)
                    talent_dropdown = ctk.CTkComboBox(frame, values=[talent_val], variable=talent_var, state="disabled")
                    talent_dropdown.grid(row=row_idx, column=0, padx=5, pady=5, sticky="ew")
                    weapon_data["Variables"]["Talent 1"] = talent_var
                    weapon_data["Dropdowns"]["Talent 1"] = talent_dropdown
            return
        
        # Get the actual weapon name to use for lookups (may differ from displayed name)
        lookup_name = weapon_data.get("Matched Weapon Name", name_val)
        
        # Get weapon row for normal weapons
        weapon_sheet = data["Weapon"][class_val]
        type_val = weapon_data["Variables"]["Type"].get()
        weapon_row = weapon_sheet[(weapon_sheet["Name"] == lookup_name) & (weapon_sheet["Type"] == type_val)].iloc[0]
        
        # Core Attributes (2 slots) - always predetermined or NA
        for i in range(1, 3):
            core_val = weapon_row.get(f"Core Attribute {i}", pd.NA)
            if pd.isna(core_val):
                continue
            
            core_var = ctk.StringVar(value=core_val)
            core_dropdown = ctk.CTkComboBox(frame, values=[core_val], variable=core_var, state="disabled")
            core_dropdown.grid(row=row_idx, column=0, padx=5, pady=5, sticky="ew")
            weapon_data["Variables"][f"Core Attribute {i}"] = core_var
            weapon_data["Dropdowns"][f"Core Attribute {i}"] = core_dropdown
            row_idx += 1
        
        # Attribute (1 slot) - more complex logic
        attr_cell = weapon_row.get("Attribute", pd.NA)
        if not pd.isna(attr_cell):
            saved_value = saved_selections.get("Attribute", "Select Attribute")
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
                    command=lambda choice, wname=weapon_name, wdata=weapon_data: self.update_weapon(wname, wdata, trigger_point="attribute"),
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
                    command=lambda choice, wname=weapon_name, wdata=weapon_data: self.update_weapon(wname, wdata, trigger_point="attribute"),
                    state="readonly"
                )
            else:
                # Fixed attribute
                attr_var.set(attr_cell)
                attr_dropdown = ctk.CTkComboBox(frame, values=[attr_cell], variable=attr_var, state="disabled")
            
            attr_dropdown.grid(row=row_idx, column=0, padx=5, pady=5, sticky="ew")
            weapon_data["Variables"]["Attribute"] = attr_var
            weapon_data["Dropdowns"]["Attribute"] = attr_dropdown
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
            
            saved_value = saved_selections.get(column_name, f"Select {column_name}")
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
                        spec_val = self.specialization_variable.get() if hasattr(self, 'specialization_variable') else "Select Specialization"
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
                            command=lambda choice, wname=weapon_name, wdata=weapon_data: self.update_weapon(wname, wdata, trigger_point="mod"),
                            state="readonly"
                        )
                        
                        # Store original mods for reference
                        if "Original Weapon Mods" not in weapon_data:
                            weapon_data["Original Weapon Mods"] = {}
                        weapon_data["Original Weapon Mods"][column_name] = original_mods
                    else:
                        continue
                else:
                    continue
            else:
                # Fixed mod
                mod_var.set(mod_cell)
                mod_dropdown = ctk.CTkComboBox(frame, values=[mod_cell], variable=mod_var, state="disabled")
            
            mod_dropdown.grid(row=row_idx, column=0, padx=5, pady=5, sticky="ew")
            weapon_data["Variables"][column_name] = mod_var
            weapon_data["Dropdowns"][column_name] = mod_dropdown
            row_idx += 1
        
        # Talents (up to 2 slots for normal weapons)
        for i in range(1, 3):
            talent_cell = weapon_row.get(f"Talent {i}", pd.NA)
            if pd.isna(talent_cell):
                continue
            
            key = f"Talent {i}"
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
                        command=lambda choice, wname=weapon_name, wdata=weapon_data: self.update_weapon(wname, wdata, trigger_point="attribute"),
                        state="readonly"
                    )
                else:
                    continue
            else:
                # Fixed talent
                talent_var.set(talent_cell)
                talent_dropdown = ctk.CTkComboBox(frame, values=[talent_cell], variable=talent_var, state="disabled")
            
            talent_dropdown.grid(row=row_idx, column=0, padx=5, pady=5, sticky="ew")
            weapon_data["Variables"][key] = talent_var
            weapon_data["Dropdowns"][key] = talent_dropdown
            row_idx += 1


#################### SECTION BREAK ####################

##### MAIN EXECUTION #####

if __name__ == "__main__":
    app = DamageCalculatorApp()
    app.mainloop()

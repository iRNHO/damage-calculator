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
from typing import Any, Callable, List, Optional


#################### SECTION BREAK ####################

##### DATA LOADING #####

root_directory = Path(user_data_dir(__title__, __author__))
root_directory = Path.cwd() # Remove this line for production use
assets_directory = root_directory / "assets"
builds_directory = root_directory / "builds"

builds_directory.mkdir(parents=True, exist_ok=True)

spreadsheet = pd.read_excel(root_directory / "iRNHO'S Spreadsheet.xlsx", sheet_name=None)
spreadsheet["Rifles"]["Name"] = spreadsheet["Rifles"]["Name"].astype(str)


#################### SECTION BREAK ####################

##### CONSTANTS #####

SPECIALIZATIONS = spreadsheet["Specializations"]["Name"].tolist()

WEAPON_CLASSES = ["Assault Rifles", "Light Machineguns", "Marksman Rifles", "Pistols", "Rifles", "Shotguns", "Sub Machine Guns", "Signature Weapons"]
WEAPON_TYPES = ["High-End", "Named", "Exotic"]

GEAR_TYPES = ["Improvised", "Brand Set", "Gear Set", "Named", "Exotic"]

SKILL_SLOTS = ["Left", "Right"]
SKILL_CLASSES = spreadsheet["Skills"]["Class"].tolist()




#################### SECTION BREAK ####################

##### UTILITY FUNCTIONS #####

def remove_elements(original_elements: List[Any], elements_to_remove: List[Any]) -> List[Any]:
    """
    Removes specified elements from a given list.

    Parameters:
        original_elements: A list containing the original elements.
        elements_to_remove: A list containing the elements to be removed.

    Returns:
        The list of elements that remain after removing the specified elements.
    
    """

    return [original_element for original_element in original_elements if original_element not in elements_to_remove]


#################### SECTION BREAK ####################

##### GUI CLASS #####

class DamageCalculatorApp(ctk.CTk):

    def __init__(self) -> None:
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

        navigation_bar = ctk.CTkFrame(self, height=60, corner_radius=0)
        navigation_bar.grid(row=0, column=0, padx=0, pady=0, sticky="ew")
        navigation_bar.grid_columnconfigure(1, weight=1)

        burger_button = ctk.CTkButton(navigation_bar, width=30, height=30, fg_color="transparent", text="", image=ctk.CTkImage(Image.open(assets_directory / "burger.png"), size=(25, 25)), command=self.toggle_sidebar)
        burger_button.grid(row=0, column=0, padx=10, pady=10, sticky="w")

        self.reset_button = ctk.CTkButton(navigation_bar, width=30, height=30, fg_color="transparent", text="", image=ctk.CTkImage(Image.open(assets_directory / "reset.png"), size=(25, 25)), command=self.reset_build)
        self.reset_button.grid(row=0, column=2, padx=10, pady=10, sticky="e")

        button_container = ctk.CTkFrame(navigation_bar, fg_color="transparent")
        button_container.grid(row=0, column=1, sticky="")

        self.build_creator_button = ctk.CTkButton(button_container, width=200, height=50, text="Build Creator", font=ctk.CTkFont(size=16, weight="bold"), state="disabled", command=self.show_build_creator)
        self.build_creator_button.pack(padx=5, side="left")

        self.build_tuning_button = ctk.CTkButton(button_container, width=200, height=50, text="Build Tuning", font=ctk.CTkFont(size=16, weight="bold"), command=self.show_build_tuning)
        self.build_tuning_button.pack(padx=5, side="left")

        self.damage_output_button = ctk.CTkButton(button_container, width=200, height=50, text="Damage Output", font=ctk.CTkFont(size=16, weight="bold"), command=self.show_damage_output)
        self.damage_output_button.pack(padx=5, side="left")

        main_container = ctk.CTkFrame(self)
        main_container.grid(row=1, column=0, sticky="nsew")
        main_container.grid_rowconfigure(0, weight=1)
        main_container.grid_columnconfigure(0, weight=1)

        self.build_creator_container = ctk.CTkFrame(main_container)
        self.build_tuning_container = ctk.CTkFrame(main_container)
        self.damage_output_container = ctk.CTkFrame(main_container)
        self.build_creator_container.grid(row=0, column=0, sticky="nsew")

        self.setup_sidebar()
        self.setup_build_creator()
        self.setup_build_tuning()
        self.setup_damage_output()

        self.active_tab_container = self.build_creator_container


    ####################

    def toggle_sidebar(self):
        """
        Toggles the sidebar visibility.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.

        """

        if self.sidebar_is_visible:
            self.sidebar_container.place_forget()
            self.sidebar_is_visible = False
            self.unbind("<Escape>")

        else:
            self.sidebar_container.place(x=0, y=0, relheight=1.0)
            self.sidebar_is_visible = True
            self.bind("<Escape>", lambda event: self.toggle_sidebar())


    def reset_build(self) -> None:
        """
        Resets the current selections to their default state.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.

        """

        self.specialization_variable.set("Select Specialization")
        self.specialization_dropdown.configure(values=SPECIALIZATIONS)

        self.delete_variable_dropdown_pairs(self.weapon_data, ["Class"])
        self.weapon_data["Variable-Dropdown Pairs"]["Class"][0].set("Select Class")
        self.weapon_data["Variable-Dropdown Pairs"]["Class"][1].configure(values=WEAPON_CLASSES)
        self.weapon_data["Variable-Label Pairs"]["Expertise"][0].set(30)
        self.weapon_data["Variable-Label Pairs"]["Expertise"][1].configure(text="Expertise: 30")
        self.weapon_value_history = {}

        for gear_data in self.gear_sections.values():
            self.delete_variable_dropdown_pairs(gear_data, ["Type"])
            gear_data["Variable-Dropdown Pairs"]["Type"][0].set("Select Type")
            gear_data["Variable-Dropdown Pairs"]["Type"][1].configure(values=GEAR_TYPES)

        self.gear_value_history = {}

        for skill_data in self.skill_sections.values():
            self.delete_variable_dropdown_pairs(skill_data, ["Class"])
            skill_data["Variable-Dropdown Pairs"]["Class"][0].set("Select Class")
            skill_data["Variable-Dropdown Pairs"]["Class"][1].configure(values=SKILL_CLASSES)
            skill_data["Variable-Label Pairs"]["Expertise"][0].set(30)
            skill_data["Variable-Label Pairs"]["Expertise"][1].configure(text="Expertise: 30")

        self.skill_value_history = {}


    def show_build_creator(self) -> None:
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


    def show_build_tuning(self) -> None:
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


    def show_damage_output(self) -> None:
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


    def setup_sidebar(self):
        """
        Initializes the sidebar with its UI components.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.

        """

        self.sidebar_container = ctk.CTkFrame(self, width=550, corner_radius=0)
        self.sidebar_container.pack_propagate(False)

        close_button = ctk.CTkButton(self.sidebar_container, width=30, height=30, fg_color="transparent", text="", image=ctk.CTkImage(Image.open(assets_directory / "exit.png"), size=(20, 20)), command=self.toggle_sidebar)
        close_button.place(x=500, y=10)

        self.sidebar_scrollable_container = ctk.CTkScrollableFrame(self.sidebar_container)
        self.sidebar_scrollable_container.pack(fill="both", expand=True, padx=10, pady=(50, 10))

        self.populate_sidebar()

        self.sidebar_is_visible = False


    def setup_build_creator(self) -> None:
        """
        Populates the 'Build Creator' tab with its UI components.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.

        """

        self.build_creator_container.grid_columnconfigure(0, weight=1)
        self.build_creator_container.grid_rowconfigure(0, weight=1)

        scrollable_container = ctk.CTkScrollableFrame(self.build_creator_container)
        scrollable_container.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        scrollable_container.grid_columnconfigure((0, 1, 2), weight=1)

        isolated_container = ctk.CTkFrame(scrollable_container, fg_color="transparent")
        isolated_container.grid(row=0, column=2, rowspan=4, padx=0, pady=0, sticky="new")
        isolated_container.grid_columnconfigure(0, weight=1)

        self.create_specialization_section(isolated_container, "demolitionist.png")
        self.create_weapon_section(isolated_container, "sub_machine_gun.png")

        self.create_gear_section(scrollable_container, 0, 0, "mask.png", "Mask")
        self.create_gear_section(scrollable_container, 1, 0, "body_armor.png", "Body Armor")
        self.create_gear_section(scrollable_container, 2, 0, "holster.png", "Holster")
        self.create_gear_section(scrollable_container, 0, 1, "backpack.png", "Backpack")
        self.create_gear_section(scrollable_container, 1, 1, "gloves.png", "Gloves")
        self.create_gear_section(scrollable_container, 2, 1, "kneepads.png", "Kneepads")

        self.create_skill_section(scrollable_container, 3, 0, "seeker_mine.png", "Left")
        self.create_skill_section(scrollable_container, 3, 1, "seeker_mine.png", "Right")


    def setup_build_tuning(self) -> None:
        """
        Populates the 'Build Tuning' tab with its UI components.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.

        """

        pass


    def setup_damage_output(self) -> None:
        """
        
        Populates the 'Damage Output' tab with its UI components.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.

        """

        pass


    ####################

    def delete_variable_dropdown_pairs(self, item_data: dict, exempt_variables: List[str]) -> None:
        """
        Deletes all variable-dropdown pairs from a given dictionary except for those specified in the exempt list.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            item_data: The dictionary containing variable-dropdown pairs.
            exempt_variables: A list of variable names whose corresponding variable-dropdown pairs should not be deleted.

        """

        for variable_name in remove_elements(list(item_data["Variable-Dropdown Pairs"].keys()), exempt_variables):
            item_data["Variable-Dropdown Pairs"][variable_name][1].destroy()
            del item_data["Variable-Dropdown Pairs"][variable_name]


    def populate_sidebar(self):
        """
        Populates the sidebar with build objects.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.

        """

        for build_object in self.sidebar_scrollable_container.winfo_children():
            build_object.destroy()

        build_dictionaries = []

        for build_path in builds_directory.iterdir():
            with open(build_path, "r") as file:
                build_dictionaries.append(json.load(file))

        for build_dictionary in sorted(build_dictionaries, key=lambda build_dictionary: build_dictionary["Build Name"].lower()):
            build_container = ctk.CTkFrame(self.sidebar_scrollable_container)
            build_container.pack(padx=5, pady=5, fill="x")

            build_label = ctk.CTkLabel(build_container, text=build_dictionary["Build Name"], font=ctk.CTkFont(size=14, weight="bold"), anchor="w")
            build_label.pack(padx=10, pady=10, expand=True, fill="x", side="left")

            button_container = ctk.CTkFrame(build_container, fg_color="transparent")
            button_container.pack(padx=5, pady=5, side="right")

            load_button = ctk.CTkButton(button_container, width=40, height=40, border_width=2, fg_color="transparent", border_color="gray30", text="", image=ctk.CTkImage(Image.open(assets_directory / "load.png"), size=(30, 30)), command=lambda command=self.load_build, build_dictionary=build_dictionary: command(build_dictionary))
            load_button.pack(padx=2, side="left")

            rename_button = ctk.CTkButton(button_container, width=40, height=40, border_width=2, fg_color="transparent", border_color="gray30", text="", image=ctk.CTkImage(Image.open(assets_directory / "rename.png"), size=(30, 30)), command=lambda command=self.rename_build, build_dictionary=build_dictionary: command(build_dictionary))
            rename_button.pack(padx=2, side="left")

            overwrite_button = ctk.CTkButton(button_container, width=40, height=40, border_width=2, fg_color="transparent", border_color="gray30", text="", image=ctk.CTkImage(Image.open(assets_directory / "overwrite.png"), size=(30, 30)), command=lambda command=self.overwrite_build, build_dictionary=build_dictionary: command(build_dictionary))
            overwrite_button.pack(padx=2, side="left")

            delete_button = ctk.CTkButton(button_container, width=40, height=40, border_width=2, fg_color="transparent", border_color="gray30", text="", image=ctk.CTkImage(Image.open(assets_directory / "delete.png"), size=(30, 30)), command=lambda command=self.delete_build, build_dictionary=build_dictionary: command(build_dictionary))
            delete_button.pack(padx=2, side="left")

        new_build_container = ctk.CTkFrame(self.sidebar_scrollable_container, height=80)
        new_build_container.pack(padx=5, pady=10, fill="x")

        new_build_button = ctk.CTkButton(new_build_container, height=60, text="+ Save Build", font=ctk.CTkFont(size=16, weight="bold"), command=self.create_new_build)
        new_build_button.pack(fill="both", expand=True, padx=10, pady=10)


    def create_specialization_section(self, parent_container: ctk.CTkFrame, icon_file: str) -> None:
        """
        Creates the specialization selection section.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            parent_container: The parent container where the specialization section will be added.
            icon_file: The file name of the icon to be displayed in the specialization section.

        """

        specialization_container = ctk.CTkFrame(parent_container, border_width=2)
        specialization_container.pack(padx=10, pady=10, fill="x")
        specialization_container.grid_columnconfigure(1, weight=1)
        specialization_container.grid_rowconfigure(0, weight=1)

        icon_label = ctk.CTkLabel(specialization_container, text="", image=ctk.CTkImage(Image.open(assets_directory / icon_file), size=(40, 40)))
        icon_label.grid(row=0, column=0, padx=10, pady=10, sticky="")

        dropdown_container = ctk.CTkFrame(specialization_container)
        dropdown_container.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        dropdown_container.grid_columnconfigure(0, weight=1)

        specialization_variable = ctk.StringVar(value="Select Specialization")
        specialization_dropdown = ctk.CTkComboBox(dropdown_container, values=SPECIALIZATIONS, state="readonly", variable=specialization_variable, command=lambda choice: self.update_specialization_section(choice))
        specialization_dropdown.grid(row=0, column=0, padx=5, pady=5, sticky="ew")

        self.specialization_variable = specialization_variable
        self.specialization_dropdown = specialization_dropdown


    def create_weapon_section(self, parent_container: ctk.CTkFrame, icon_file: str) -> None:
        """
        Creates the weapon selection section.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            parent_container: The parent container where the weapon section will be added.
            icon_file: The file name of the icon to be displayed in the weapon section.

        """

        weapon_container = ctk.CTkFrame(parent_container, border_width=2)
        weapon_container.pack(padx=10, pady=10, expand=True, fill="both")
        weapon_container.grid_columnconfigure(1, weight=1)
        weapon_container.grid_rowconfigure(0, weight=1)

        icon_label = ctk.CTkLabel(weapon_container, text="", image=ctk.CTkImage(Image.open(assets_directory / icon_file), size=(40, 40)))
        icon_label.grid(row=0, column=0, padx=10, pady=10, sticky="")

        dropdown_container = ctk.CTkFrame(weapon_container)
        dropdown_container.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        dropdown_container.grid_columnconfigure(0, weight=1)

        class_variable = ctk.StringVar(value="Select Class")
        class_dropdown = ctk.CTkComboBox(dropdown_container, values=WEAPON_CLASSES, state="readonly", variable=class_variable, command=lambda choice: self.update_weapon_section("Class", choice))
        class_dropdown.grid(row=0, column=0, padx=5, pady=5, sticky="ew")

        expertise_label = ctk.CTkLabel(dropdown_container, text=f"Expertise: 30")
        expertise_label.grid(row=69, column=0, padx=5, pady=(10, 2), sticky="w")

        expertise_variable = ctk.IntVar(value=30)
        expertise_slider = ctk.CTkSlider(dropdown_container, to=30, number_of_steps=30, command=lambda value: expertise_label.configure(text=f"Expertise: {value:.0f}"), variable=expertise_variable)
        expertise_slider.grid(row=420, column=0, padx=5, pady=(0, 5), sticky="ew")

        self.weapon_data = {
            "Dropdown Container": dropdown_container,
            "Variable-Dropdown Pairs": {
                "Class": (class_variable, class_dropdown)
            },
            "Variable-Label Pairs": {
                "Expertise": (expertise_variable, expertise_label)
            }
        }

        self.weapon_value_history = {}


    def create_gear_section(self, parent_container: ctk.CTkFrame, row_index: int, column_index: int, icon_file: str, gear_slot: str) -> None:
        """
        Creates a gear piece selection section.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            parent_container: The parent container where the gear section will be added.
            row_index: The row index for placing the gear section in the grid.
            column_index: The column index for placing the gear section in the grid.
            icon_file: The file name of the icon to be displayed in the gear section.
            gear_slot: The name of the gear slot.

        """

        gear_container = ctk.CTkFrame(parent_container, border_width=2)
        gear_container.grid(row=row_index, column=column_index, padx=10, pady=10, sticky="nsew")
        gear_container.grid_columnconfigure(1, weight=1)
        gear_container.grid_rowconfigure(0, weight=1)

        icon_label = ctk.CTkLabel(gear_container, text="", image=ctk.CTkImage(Image.open(assets_directory / icon_file), size=(40, 40)))
        icon_label.grid(row=0, column=0, padx=10, pady=10, sticky="")

        dropdown_container = ctk.CTkFrame(gear_container)
        dropdown_container.grid(row=0, column=1, sticky="ew", padx=10, pady=10)
        dropdown_container.grid_columnconfigure(0, weight=1)

        type_variable = ctk.StringVar(value="Select Type")
        type_dropdown = ctk.CTkComboBox(dropdown_container, values=GEAR_TYPES, state="readonly", variable=type_variable, command=lambda choice: self.update_gear_section(gear_slot, "Type", choice))
        type_dropdown.grid(row=0, column=0, padx=5, pady=5, sticky="ew")

        if not hasattr(self, "gear_sections"):
            self.gear_sections = {}

        self.gear_sections[gear_slot] = {
            "Dropdown Container": dropdown_container,
            "Variable-Dropdown Pairs": {
                "Type": (type_variable, type_dropdown)
            }
        }

        if not hasattr(self, "gear_value_history"):
            self.gear_value_history = {}


    def create_skill_section(self, parent_container: ctk.CTkFrame, row_index: int, column_index: int, icon_file: str, skill_slot: str) -> None:
        """
        Creates a skill selection section.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            parent_container: The parent container where the skill section will be added.
            row_index: The row index for placing the skill section in the grid.
            column_index: The column index for placing the skill section in the grid.
            icon_file: The file name of the icon to be displayed in the skill section.
            skill_slot: The name of the skill slot.

        """

        skill_container = ctk.CTkFrame(parent_container, border_width=2)
        skill_container.grid(row=row_index, column=column_index, padx=10, pady=(20, 10), sticky="nsew")
        skill_container.grid_columnconfigure(1, weight=1)
        skill_container.grid_rowconfigure(0, weight=1)

        icon_label = ctk.CTkLabel(skill_container, text="", image=ctk.CTkImage(Image.open(assets_directory / icon_file), size=(40, 40)))
        icon_label.grid(row=0, column=0, padx=10, pady=10, sticky="")

        dropdown_container = ctk.CTkFrame(skill_container)
        dropdown_container.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        dropdown_container.grid_columnconfigure(0, weight=1)

        class_variable = ctk.StringVar(value="Select Class")
        class_dropdown = ctk.CTkComboBox(dropdown_container, values=SKILL_CLASSES, state="readonly", variable=class_variable, command=lambda choice: self.update_skill_section(skill_slot, "Class", choice))
        class_dropdown.grid(row=0, column=0, padx=5, pady=5, sticky="ew")

        expertise_label = ctk.CTkLabel(dropdown_container, text=f"Expertise: 30")
        expertise_label.grid(row=69, column=0, padx=5, pady=(10, 2), sticky="w")

        expertise_variable = ctk.IntVar(value=30)
        expertise_slider = ctk.CTkSlider(dropdown_container, to=30, number_of_steps=30, command=lambda value: expertise_label.configure(text=f"Expertise: {value:.0f}"), variable=expertise_variable)
        expertise_slider.grid(row=420, column=0, padx=5, pady=(0, 5), sticky="ew")

        if not hasattr(self, "skill_sections"):
            self.skill_sections = {}

        self.skill_sections[skill_slot] = {
            "Dropdown Container": dropdown_container,
            "Variable-Dropdown Pairs": {
                "Class": (class_variable, class_dropdown)
            },
            "Variable-Label Pairs": {
                "Expertise": (expertise_variable, expertise_label)
            }
        }

        if not hasattr(self, "skill_value_history"):
            self.skill_value_history = {}


    ####################

    def load_build(self, build_dictionary: dict) -> None:
        """
        Loads a build into the application.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            build_dictionary: The dictionary representing the build to be loaded.

        """

        self.reset_build()

        if "Specialization" in build_dictionary:
            specialization_value = build_dictionary["Specialization"]
            self.specialization_variable.set(specialization_value)
            self.update_specialization_section(specialization_value)

        weapon_dictionary = build_dictionary["Weapon"]

        for variable_name, variable_value in weapon_dictionary.items():
            if variable_name != "Expertise":
                self.weapon_data["Variable-Dropdown Pairs"][variable_name][0].set(variable_value)
                self.update_weapon_section(variable_name, variable_value)

            else:
                self.weapon_data["Variable-Label Pairs"][variable_name][0].set(variable_value)
                self.weapon_data["Variable-Label Pairs"][variable_name][1].configure(text=f"Expertise: {variable_value}")

        for gear_slot, gear_data in self.gear_sections.items():
            if gear_slot in build_dictionary:
                gear_dictionary = build_dictionary[gear_slot]

                for variable_name, variable_value in gear_dictionary.items():
                    gear_data["Variable-Dropdown Pairs"][variable_name][0].set(variable_value)
                    self.update_gear_section(gear_slot, variable_name, variable_value)

        for skill_slot, skill_data in self.skill_sections.items():
            if skill_slot in build_dictionary:
                skill_dictionary = build_dictionary[skill_slot]

                for variable_name, variable_value in skill_dictionary.items():
                    if variable_name != "Expertise":
                        skill_data["Variable-Dropdown Pairs"][variable_name][0].set(variable_value)
                        self.update_skill_section(skill_slot, variable_name, variable_value)

                    else:
                        skill_data["Variable-Label Pairs"][variable_name][0].set(variable_value)
                        skill_data["Variable-Label Pairs"][variable_name][1].configure(text=f"Expertise: {variable_value}")


    def rename_build(self, build_dictionary: dict) -> None:
        """
        Renames a build based on user input.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            build_dictionary: The dictionary representing the build to be renamed.

        """

        build_name = build_dictionary["Build Name"]
        new_name = self.prompt_for_build_name("Enter new build name:", build_name)

        if new_name and new_name != build_name:
            new_name = self.ensure_unique_build_name(new_name, build_name)

            if new_name != build_name:
                (builds_directory / f"{self.sanitize_filename(build_name)}.json").unlink()

                build_dictionary["Build Name"] = new_name
                self.create_build_file(build_dictionary)

                self.populate_sidebar()


    def overwrite_build(self, build_dictionary: dict) -> None:
        """
        Overwrites an existing build with the current selections.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            build_dictionary: The dictionary representing the build to be overwritten.

        """

        build_name = build_dictionary["Build Name"]
        self.create_build_file(self.get_current_selections(build_name))

        self.populate_sidebar()


    def delete_build(self, build_dictionary: dict) -> None:
        """
        Deletes a build file.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            build_dictionary: The dictionary representing the build to be deleted.

        """

        (builds_directory / f"{self.sanitize_filename(build_dictionary["Build Name"])}.json").unlink()

        self.populate_sidebar()


    def create_new_build(self) -> None:
        """
        Creates a new build file using the current selections.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.

        """

        build_name = self.prompt_for_build_name("Enter build name:")

        if build_name:
            build_name = self.ensure_unique_build_name(build_name)

            self.create_build_file(self.get_current_selections(build_name))

            self.populate_sidebar()


    def update_specialization_section(self, user_selection: str) -> None:
        """
        Updates the specialization section dynamically on user input.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            user_selection: The specialization value selected by the user.

        """

        self.specialization_dropdown.configure(values=remove_elements(SPECIALIZATIONS, [user_selection]))

        class_variable_value = self.get_variable_value(self.weapon_data, "Class")

        if class_variable_value != "Select Class":
            self.update_weapon_section("Class", class_variable_value)

        for skill_slot in SKILL_SLOTS:
            class_variable_value = self.get_variable_value(self.skill_sections[skill_slot], "Class")

            if class_variable_value != "Select Class":
                self.update_skill_section(skill_slot, "Class", class_variable_value)


    def update_weapon_section(self, trigger_point: str, user_selection: str) -> None:
        """
        Updates the weapon section dynamically on user input.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            trigger_point: The name of the variable that triggered the update.
            user_selection: The value of the variable as selected by the user.

        """

        if trigger_point == "Class":
            self.delete_variable_dropdown_pairs(self.weapon_data, ["Class"])

            self.update_dropdown_values(self.weapon_data, "Class", remove_elements(WEAPON_CLASSES, [user_selection]))

            if user_selection != "Signature Weapons":
                type_variable_value, type_dropdown_values = self.get_variable_dropdown_values(WEAPON_TYPES, "Type", self.weapon_value_history, "Select Type")
                self.create_variable_dropdown_pair(type_variable_value, type_dropdown_values, self.weapon_data, lambda choice: self.update_weapon_section("Type", choice), 1, "Type")

                if type_variable_value != "Select Type":
                    self.update_weapon_section("Type", type_variable_value)

            else:
                name_variable_value, name_dropdown_values = self.get_variable_dropdown_values(self.get_allowed_weapon_name_values(user_selection, None), "Name", self.weapon_value_history, "Select Name")
                self.create_variable_dropdown_pair(name_variable_value, name_dropdown_values, self.weapon_data, lambda choice: self.update_weapon_section("Name", choice), 1, "Name")

                if name_variable_value != "Select Name":
                    self.update_weapon_section("Name", name_variable_value)

        elif trigger_point == "Type":
            self.delete_variable_dropdown_pairs(self.weapon_data, ["Class", "Type"])

            self.update_value_history("Type", self.weapon_value_history, user_selection)
            self.update_dropdown_values(self.weapon_data, "Type", remove_elements(WEAPON_TYPES, [user_selection]))

            name_variable_value, name_dropdown_values = self.get_variable_dropdown_values(self.get_allowed_weapon_name_values(self.get_variable_value(self.weapon_data, "Class"), user_selection), "Name", self.weapon_value_history, "Select Name")
            self.create_variable_dropdown_pair(name_variable_value, name_dropdown_values, self.weapon_data, lambda choice: self.update_weapon_section("Name", choice), 2, "Name")

            if name_variable_value != "Select Name":
                self.update_weapon_section("Name", name_variable_value)

        elif trigger_point == "Name":
            if "Select" in user_selection:
                self.weapon_data["Variable-Dropdown Pairs"]["Name"][0].set(self.weapon_value_history["Name"][-1])
                return

            self.delete_variable_dropdown_pairs(self.weapon_data, ["Class", "Type", "Name"])

            class_variable_value = self.get_variable_value(self.weapon_data, "Class")
            row_index = 3 if class_variable_value != "Signature Weapons" else 2

            self.update_value_history("Name", self.weapon_value_history, user_selection)
            self.update_dropdown_values(self.weapon_data, "Name", remove_elements(self.get_allowed_weapon_name_values(class_variable_value, self.get_variable_value(self.weapon_data, "Type")), [user_selection]))

            for variable_name in ["Core Attribute 1", "Core Attribute 2"]:
                allowed_core_attribute_values = self.get_allowed_weapon_core_attribute_values(class_variable_value, user_selection, variable_name)

                if allowed_core_attribute_values:
                    self.create_variable_dropdown_pair(*self.get_variable_dropdown_values(allowed_core_attribute_values, variable_name, self.weapon_value_history, "Select Core Attribute"), self.weapon_data, lambda choice, trigger_point=variable_name: self.update_weapon_section(trigger_point, choice), row_index, variable_name)
                    row_index += 1

            allowed_attribute_values = self.get_allowed_weapon_attribute_values(class_variable_value, user_selection)

            if allowed_attribute_values:
                self.create_variable_dropdown_pair(*self.get_variable_dropdown_values(allowed_attribute_values, "Attribute", self.weapon_value_history, "Select Attribute"), self.weapon_data, lambda choice: self.update_weapon_section("Attribute", choice), row_index, "Attribute")
                row_index += 1

            for variable_name in ["Optics Mod", "Magazine Mod", "Underbarrel Mod", "Muzzle Mod"]:
                allowed_mod_values = self.get_allowed_weapon_mod_values(class_variable_value, user_selection, variable_name)

                if allowed_mod_values:
                    self.create_variable_dropdown_pair(*self.get_variable_dropdown_values(allowed_mod_values, variable_name, self.weapon_value_history, f"Select {variable_name}"), self.weapon_data, lambda choice, trigger_point=variable_name: self.update_weapon_section(trigger_point, choice), row_index, variable_name)
                    row_index += 1

            for variable_name in ["Talent 1", "Talent 2"]:
                allowed_talent_values = self.get_allowed_weapon_talent_values(class_variable_value, user_selection, variable_name)

                if allowed_talent_values:
                    self.create_variable_dropdown_pair(*self.get_variable_dropdown_values(allowed_talent_values, variable_name, self.weapon_value_history, "Select Talent"), self.weapon_data, lambda choice, trigger_point=variable_name: self.update_weapon_section(trigger_point, choice), row_index, variable_name)
                    row_index += 1

        elif trigger_point == "Attribute":
            self.update_value_history(trigger_point, self.weapon_value_history, user_selection)
            self.update_dropdown_values(self.weapon_data, trigger_point, remove_elements(self.get_allowed_weapon_attribute_values(self.get_variable_value(self.weapon_data, "Class"), self.get_variable_value(self.weapon_data, "Name")), [user_selection]))

        elif trigger_point in ["Optics Mod", "Magazine Mod", "Underbarrel Mod", "Muzzle Mod"]:
            if "Select" in user_selection:
                self.weapon_data["Variable-Dropdown Pairs"][trigger_point][0].set(self.weapon_value_history[trigger_point][-1])
                return

            self.update_value_history(trigger_point, self.weapon_value_history, user_selection)
            self.update_dropdown_values(self.weapon_data, trigger_point, remove_elements(self.get_allowed_weapon_mod_values(self.get_variable_value(self.weapon_data, "Class"), self.get_variable_value(self.weapon_data, "Name"), trigger_point), [user_selection]))

        else:
            self.update_value_history(trigger_point, self.weapon_value_history, user_selection)
            self.update_dropdown_values(self.weapon_data, trigger_point, remove_elements(self.get_allowed_weapon_talent_values(self.get_variable_value(self.weapon_data, "Class"), self.get_variable_value(self.weapon_data, "Name"), trigger_point), [user_selection]))


    def update_gear_section(self, gear_slot: str, trigger_point: str, user_selection: str) -> None:
        """
        Updates a gear piece section dynamically on user input.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            gear_slot: The name of the gear piece slot.
            trigger_point: The name of the variable that triggered the update.
            user_selection: The value of the variable as selected by the user.

        """

        gear_data = self.gear_sections[gear_slot]

        if trigger_point == "Type":
            base_values = remove_elements(GEAR_TYPES, ["Exotic"]) if any(self.get_variable_value(_gear_data, "Type") == "Exotic" for _gear_data in self.gear_sections.values()) else GEAR_TYPES

            for _gear_data in self.gear_sections.values():
                self.update_dropdown_values(_gear_data, "Type", remove_elements(base_values, [self.get_variable_value(_gear_data, "Type")]))

            self.delete_variable_dropdown_pairs(gear_data, ["Type"])

            name_variable_value, name_dropdown_values = self.get_variable_dropdown_values(self.get_allowed_gear_name_values(user_selection, gear_slot), "Name", self.gear_value_history, "Select Name")
            self.create_variable_dropdown_pair(name_variable_value, name_dropdown_values, gear_data, lambda choice: self.update_gear_section(gear_slot, "Name", choice), 1, "Name")

            if name_variable_value != "Select Name":
                self.update_gear_section(gear_slot, "Name", name_variable_value)

        elif trigger_point == "Name":
            self.delete_variable_dropdown_pairs(gear_data, ["Type", "Name"])

            type_variable_value = self.get_variable_value(gear_data, "Type")
            row_index = 2

            self.update_value_history("Name", self.gear_value_history, user_selection)
            self.update_dropdown_values(gear_data, "Name", remove_elements(self.get_allowed_gear_name_values(type_variable_value, gear_slot), [user_selection]))

            for variable_name in ["Core Attribute 1", "Core Attribute 2", "Core Attribute 3"]:
                allowed_core_attribute_values = self.get_allowed_gear_core_attribute_values(type_variable_value, user_selection, variable_name)

                if allowed_core_attribute_values:
                    self.create_variable_dropdown_pair(*self.get_variable_dropdown_values(allowed_core_attribute_values, variable_name, self.gear_value_history, "Select Core Attribute"), gear_data, lambda choice, trigger_point=variable_name: self.update_gear_section(gear_slot, trigger_point, choice), row_index, variable_name)
                    row_index += 1

            current_values = []
            value_values_pairs = {}

            for variable_name in ["Attribute 1", "Attribute 2", "Attribute 3"]:
                allowed_attribute_values = remove_elements(self.get_allowed_gear_attribute_values(type_variable_value, variable_name, user_selection), current_values)

                if allowed_attribute_values:
                    attribute_variable_value, attribute_dropdown_values = self.get_variable_dropdown_values(allowed_attribute_values, variable_name, self.gear_value_history, "Select Attribute")
                    value_values_pairs[variable_name] = (attribute_variable_value, attribute_dropdown_values)
                    current_values.append(attribute_variable_value)

            for variable_name, (attribute_variable_value, attribute_dropdown_values) in value_values_pairs.items():
                self.create_variable_dropdown_pair(attribute_variable_value, remove_elements(attribute_dropdown_values, current_values) if variable_name != list(value_values_pairs.keys())[-1] else attribute_dropdown_values, gear_data, lambda choice, trigger_point=variable_name: self.update_gear_section(gear_slot, trigger_point, choice), row_index, variable_name)
                row_index += 1

            for variable_name in ["Mod 1", "Mod 2"]:
                allowed_mod_values = self.get_allowed_gear_mod_values(type_variable_value, variable_name, gear_slot, user_selection)

                if allowed_mod_values:
                    self.create_variable_dropdown_pair(*self.get_variable_dropdown_values(allowed_mod_values, variable_name, self.gear_value_history, "Select Mod"), gear_data, lambda choice, trigger_point=variable_name: self.update_gear_section(gear_slot, trigger_point, choice), row_index, variable_name)
                    row_index += 1

            for variable_name in ["Talent 1", "Talent 2"]:
                allowed_talent_values = self.get_allowed_gear_talent_values(type_variable_value, gear_slot, variable_name, user_selection)

                if allowed_talent_values:
                    self.create_variable_dropdown_pair(*self.get_variable_dropdown_values(allowed_talent_values, variable_name, self.gear_value_history, "Select Talent"), gear_data, lambda choice, trigger_point=variable_name: self.update_gear_section(gear_slot, trigger_point, choice), row_index, variable_name)
                    row_index += 1

        elif trigger_point in ["Core Attribute 1", "Core Attribute 2", "Core Attribute 3"]:
            self.update_value_history(trigger_point, self.gear_value_history, user_selection)
            self.update_dropdown_values(gear_data, trigger_point, remove_elements(self.get_allowed_gear_core_attribute_values(self.get_variable_value(gear_data, "Type"), self.get_variable_value(gear_data, "Name"), trigger_point), [user_selection]))

        elif trigger_point in ["Attribute 1", "Attribute 2", "Attribute 3"]:
            self.update_value_history(trigger_point, self.gear_value_history, user_selection)
            current_values = [user_selection]

            for variable_name in remove_elements(["Attribute 1", "Attribute 2", "Attribute 3"], [trigger_point]):
                if variable_name in gear_data["Variable-Dropdown Pairs"]:
                    current_values.append(self.get_variable_value(gear_data, variable_name))

            type_variable_value = self.get_variable_value(gear_data, "Type")
            name_variable_value = self.get_variable_value(gear_data, "Name")

            for variable_name in [variable_name for variable_name in ["Attribute 1", "Attribute 2", "Attribute 3"] if variable_name in gear_data["Variable-Dropdown Pairs"]]:
                self.update_dropdown_values(gear_data, variable_name, remove_elements(self.get_allowed_gear_attribute_values(type_variable_value, variable_name, name_variable_value), current_values))

        elif trigger_point in ["Mod 1", "Mod 2"]:
            self.update_value_history(trigger_point, self.gear_value_history, user_selection)
            self.update_dropdown_values(gear_data, trigger_point, remove_elements(self.get_allowed_gear_mod_values(self.get_variable_value(gear_data, "Type"), trigger_point, gear_slot, self.get_variable_value(gear_data, "Name")), [user_selection]))

        else:
            self.update_value_history(trigger_point, self.gear_value_history, user_selection)
            self.update_dropdown_values(gear_data, trigger_point, remove_elements(self.get_allowed_gear_talent_values(self.get_variable_value(gear_data, "Type"), gear_slot, trigger_point, self.get_variable_value(gear_data, "Name")), [user_selection]))


    def update_skill_section(self, skill_slot: str, trigger_point: str, user_selection: str) -> None:
        """
        Updates a skill section dynamically on user input.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            skill_slot: The name of the skill slot.
            trigger_point: The name of the variable that triggered the update.
            user_selection: The value of the variable as selected by the user.

        """

        skill_data = self.skill_sections[skill_slot]

        if trigger_point == "Class":
            current_values = [self.get_variable_value(_skill_data, "Class") for _skill_data in self.skill_sections.values()]

            for _skill_data in self.skill_sections.values():
                self.update_dropdown_values(_skill_data, "Class", remove_elements(SKILL_CLASSES, current_values))

            self.delete_variable_dropdown_pairs(skill_data, ["Class"])

            name_variable_value, name_dropdown_values = self.get_variable_dropdown_values(self.get_allowed_skill_name_values(self.get_variable_value(skill_data, "Class")), "Name", self.skill_value_history, "Select Name")
            self.create_variable_dropdown_pair(name_variable_value, name_dropdown_values, skill_data, lambda choice: self.update_skill_section(skill_slot, "Name", choice), 1, "Name")

            if name_variable_value != "Select Name":
                self.update_skill_section(skill_slot, "Name", name_variable_value)

        elif trigger_point == "Name":
            if "Select" in user_selection:
                skill_data["Variable-Dropdown Pairs"]["Name"][0].set(self.skill_value_history["Name"][-1])
                return

            self.delete_variable_dropdown_pairs(skill_data, ["Class", "Name"])

            class_variable_value = self.get_variable_value(skill_data, "Class")
            row_index = 2

            self.update_value_history("Name", self.skill_value_history, user_selection)
            self.update_dropdown_values(skill_data, "Name", remove_elements(self.get_allowed_skill_name_values(class_variable_value), [user_selection]))

            for variable_name in ["Mod 1", "Mod 2", "Mod 3"]:
                slot_name, allowed_mod_values = self.get_allowed_skill_mod_values(class_variable_value, variable_name)

                if allowed_mod_values:
                    self.create_variable_dropdown_pair(*self.get_variable_dropdown_values(allowed_mod_values, variable_name, self.skill_value_history, f"Select {slot_name} Mod"), skill_data, lambda choice, trigger_point=variable_name: self.update_skill_section(skill_slot, trigger_point, choice), row_index, variable_name)
                    row_index += 1

        else:
            if "Select" in user_selection:
                skill_data["Variable-Dropdown Pairs"][trigger_point][0].set(self.skill_value_history[trigger_point][-1])
                return

            self.update_value_history(trigger_point, self.skill_value_history, user_selection)
            slot_name, allowed_mod_values = self.get_allowed_skill_mod_values(self.get_variable_value(skill_data, "Class"), trigger_point)
            self.update_dropdown_values(skill_data, trigger_point, remove_elements(allowed_mod_values, [user_selection]))


    ####################

    def prompt_for_build_name(self, user_prompt: str, initial_value: str = "") -> Optional[str]:
        """
        Prompts the user for a build name.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            user_prompt: The prompt text to display to the user.
            initial_value: The initial value to populate the entry field with.

        Returns:
            The build name entered by the user if provided, otherwise 'None'.

        """

        self.unbind("<Escape>")

        dialog_container = ctk.CTkFrame(self, width=400, height=135, corner_radius=10, border_width=2, border_color="gray30")
        dialog_container.place(relx=0.5, rely=0.5, anchor="center")
        dialog_container.pack_propagate(False)

        def close_dialog():
            dialog_container.destroy()

            if self.sidebar_is_visible:
                self.bind("<Escape>", lambda event: self.toggle_sidebar())

        close_button = ctk.CTkButton(dialog_container, width=30, height=30, fg_color="transparent", text="", image=ctk.CTkImage(Image.open(assets_directory / "exit.png"), size=(16, 16)), command=close_dialog)
        close_button.place(x=360, y=10)

        prompt_label = ctk.CTkLabel(dialog_container, text=user_prompt, font=ctk.CTkFont(size=14))
        prompt_label.pack(pady=(20, 5))

        entry_field = ctk.CTkEntry(dialog_container, width=300, height=35)
        entry_field.pack(pady=5)

        if initial_value:
            entry_field.insert(0, initial_value)
            entry_field.select_range(0, "end")

        initial_length = len(initial_value)
        character_counter = ctk.CTkLabel(dialog_container, text=f"{initial_length}/20", font=ctk.CTkFont(size=12), text_color="orange" if initial_length == 20 else "gray60")
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

        result_name = None

        def on_ok():
            nonlocal result_name
            text = entry_field.get().strip()

            if text:
                result_name = text
                close_dialog()

        entry_field.bind("<Return>", lambda event: on_ok())
        entry_field.bind("<Escape>", lambda event: close_dialog())
        dialog_container.bind("<Escape>", lambda event: close_dialog())

        self.wait_window(dialog_container)

        return result_name


    def ensure_unique_build_name(self, build_name: str, acceptable_value: str = None) -> str:
        """
        Ensures that a build name is unique by appending a counter if necessary.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            build_name: The desired build name.
            acceptable_value: An acceptable build name that can be returned even if it already exists.

        Returns:
            A unique build name.

        """

        original_name = build_name
        counter = 1

        while (builds_directory / f"{self.sanitize_filename(build_name)}.json").exists():
            if acceptable_value and acceptable_value == build_name:
                return acceptable_value

            build_name = f"{original_name} ({counter})"
            counter += 1

        return build_name


    def sanitize_filename(self, string: str) -> str:
        """
        Sanitizes a string to be used as a filename by removing invalid characters.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            string: The input string to be sanitized.

        Returns:
            The sanitized filename.

        """

        return "".join(character for character in string if character not in "<>:\"/\\|?*").strip()


    def create_build_file(self, build_dictionary: dict) -> None:
        """
        Creates a build file from a build dictionary.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            build_dictionary: The dictionary representing the build.

        """

        with open(builds_directory / f"{self.sanitize_filename(build_dictionary["Build Name"])}.json", "w") as file:
            json.dump(build_dictionary, file, indent=4)


    def get_current_selections(self, build_name: str) -> dict:
        """
        Retrieves the current state of all selections as a build dictionary.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            build_name: The name for the build.

        Returns:
            The build dictionary representing the current selections.

        """

        build_dictionary = {"Build Name": build_name}

        specialization_value = self.specialization_variable.get()

        if specialization_value != "Select Specialization":
            build_dictionary["Specialization"] = specialization_value

        weapon_dictionary = {}

        for variable_name, variable_dropdown_pair in self.weapon_data["Variable-Dropdown Pairs"].items():
            variable_value = variable_dropdown_pair[0].get()

            if not variable_value.startswith("Select"):
                weapon_dictionary[variable_name] = variable_value

        weapon_dictionary["Expertise"] = self.weapon_data["Variable-Label Pairs"]["Expertise"][0].get()
        build_dictionary["Weapon"] = weapon_dictionary

        for gear_slot, gear_data in self.gear_sections.items():
            gear_dictionary = {}

            for variable_name, variable_dropdown_pair in gear_data["Variable-Dropdown Pairs"].items():
                variable_value = variable_dropdown_pair[0].get()

                if not variable_value.startswith("Select"):
                    gear_dictionary[variable_name] = variable_value

            if gear_dictionary:
                build_dictionary[gear_slot] = gear_dictionary

        for skill_slot, skill_data in self.skill_sections.items():
            skill_dictionary = {}

            for variable_name, variable_dropdown_pair in skill_data["Variable-Dropdown Pairs"].items():
                variable_value = variable_dropdown_pair[0].get()

                if not variable_value.startswith("Select"):
                    skill_dictionary[variable_name] = variable_value

            skill_dictionary["Expertise"] = skill_data["Variable-Label Pairs"]["Expertise"][0].get()
            build_dictionary[skill_slot] = skill_dictionary

        return build_dictionary


    def get_variable_value(self, item_data: dict, variable_name: str) -> Optional[str]:
        """
        Retrieves the value of a specified variable from a given dictionary.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            item_data: The dictionary, necessarily containing variable-dropdown pairs.
            variable_name: The name of the variable.

        Returns:
            The value of the specified variable if present, otherwise 'None'.

        """

        return item_data["Variable-Dropdown Pairs"][variable_name][0].get() if variable_name in item_data["Variable-Dropdown Pairs"] else None


    def update_dropdown_values(self, item_data: dict, variable_name: str, dropdown_values: List[str]) -> None:
        """
        Updates the dropdown values for a specified variable.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            item_data: The dictionary, necessarily containing variable-dropdown pairs.
            variable_name: The name of the variable.
            dropdown_values: The list of new dropdown values.

        """

        item_data["Variable-Dropdown Pairs"][variable_name][1].configure(values=dropdown_values)


    def get_variable_dropdown_values(self, allowed_values: List[str], variable_name: str, value_history: dict, default_value: str) -> tuple[str, List[str]]:
        """
        Retrieves the value and dropdown values for a specified variable.
        
        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            allowed_values: The list of allowed values for the variable.
            variable_name: The name of the variable.
            value_history: The dictionary containing the value history.
            default_value: The default value for the variable.

        Returns:
            A tuple containing the variable value and the list of dropdown values.

        """

        if len(allowed_values) == 1:
            return allowed_values[0], []

        if variable_name in value_history:
            historical_values = value_history[variable_name]

            for index in range(len(historical_values) - 1, -1, -1):
                historical_value = historical_values[index]

                if historical_value in allowed_values:
                    historical_values.append(historical_values.pop(index))
                    return historical_value, remove_elements(allowed_values, [historical_value])

        self.update_value_history(variable_name, value_history, default_value)
        return default_value, allowed_values


    def create_variable_dropdown_pair(self, variable_value: str, dropdown_values: List[str], item_data: dict, command: Callable, row_index: int, variable_name: str) -> None:
        """
        Creates a variable-dropdown pair.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            variable_value: The initial value of the variable.
            dropdown_values: The list of values for the dropdown.
            item_data: The dictionary, necessarily containing variable-dropdown pairs.
            command: The command to execute on selection change.
            row_index: The row index for grid placement.
            variable_name: The name of the variable.

        """

        variable = ctk.StringVar(value=variable_value)
        dropdown = ctk.CTkComboBox(item_data["Dropdown Container"], values=dropdown_values, state="readonly" if dropdown_values else "disabled", variable=variable, command=command)
        dropdown.grid(row=row_index, column=0, padx=5, pady=5, sticky="ew")
        item_data["Variable-Dropdown Pairs"][variable_name] = (variable, dropdown)


    def update_value_history(self, variable_name: str, value_history: dict, new_value: str) -> None:
        """
        Updates the value history for a specified variable.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            variable_name: The name of the variable.
            value_history: The dictionary containing the value history.
            new_value: The new value to add to the history.

        """

        if variable_name not in value_history:
            value_history[variable_name] = []

        if new_value in value_history[variable_name]:
            value_history[variable_name].remove(new_value)

        value_history[variable_name].append(new_value)


    def get_allowed_weapon_name_values(self, weapon_class: str, weapon_type: str) -> List[str]:
        """
        Retrieves the allowed weapon name values for a given weapon class and type.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            weapon_class: The class of the weapon.
            weapon_type: The type of the weapon.

        Returns:
            The list of allowed name values.
        
        """

        if weapon_class != "Signature Weapons":
            return self.emphasize_specialization_dependencies(spreadsheet[weapon_class][spreadsheet[weapon_class]["Type"] == weapon_type]["Name"].tolist())

        speciaization_value = self.specialization_variable.get()

        if speciaization_value != "Select Specialization":
            for weapon_name in spreadsheet["Signature Weapons"]["Name"].tolist():
                if speciaization_value in weapon_name:
                    return [weapon_name.split(" (")[0]]

        return self.emphasize_specialization_dependencies(spreadsheet["Signature Weapons"]["Name"].tolist())


    def get_allowed_weapon_core_attribute_values(self, weapon_class: str, weapon_name: str, variable_name: str) -> List[str]:
        """
        Retrieves the allowed core attribute values for a specified weapon.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            weapon_class: The class of the weapon.
            weapon_name: The name of the weapon.
            variable_name: The name of the variable.

        Returns:
            The list of allowed core attribute values.

        """

        def cell_value_processor(cell_value: str) -> List[str]:
            """
            Processes a spreadsheet cell value to determine the allowed core attribute values.

            Parameters:
                cell_value: The value from the spreadsheet cell.

            Returns:
                The list of allowed core attribute values.

            """

            if pd.isna(cell_value):
                return []

            return [cell_value]

        if weapon_class != "Signature Weapons":
            return cell_value_processor(spreadsheet[weapon_class][spreadsheet[weapon_class]["Name"] == self.de_emphasize_weapon_name(weapon_class, weapon_name)].iloc[0][variable_name])

        return []


    def get_allowed_weapon_attribute_values(self, weapon_class: str, weapon_name: str) -> List[str]:
        """
        Retrieves the allowed attribute values for a specified weapon.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            weapon_class: The class of the weapon.
            weapon_name: The name of the weapon.

        Returns:
            The list of allowed attribute values.

        """

        def cell_value_processor(cell_value: str) -> List[str]:
            """
            Processes a spreadsheet cell value to determine the allowed attribute values.

            Parameters:
                cell_value: The value from the spreadsheet cell.

            Returns:
                The list of allowed attribute values.

            """

            if pd.isna(cell_value):
                return []

            if cell_value == "*":
                return spreadsheet["Weapon Attributes"]["Stats"].tolist()

            if cell_value.startswith("!"):
                return remove_elements(spreadsheet["Weapon Attributes"]["Stats"].tolist(), [cell_value[1:]])

            return [cell_value]

        if weapon_class != "Signature Weapons":
            return cell_value_processor(spreadsheet[weapon_class][spreadsheet[weapon_class]["Name"] == self.de_emphasize_weapon_name(weapon_class, weapon_name)].iloc[0]["Attribute"])

        return []


    def get_allowed_weapon_mod_values(self, weapon_class: str, weapon_name: str, variable_name: str) -> List[str]:
        """
        Retrieves the allowed mod values for a specified weapon.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            weapon_class: The class of the weapon.
            weapon_name: The name of the weapon.
            variable_name: The name of the variable.

        Returns:
            The list of allowed mod values.

        """

        def cell_value_processor(cell_value: str) -> List[str]:
            """
            Processes a spreadsheet cell value to determine the allowed mod values.

            Parameters:
                cell_value: The value from the spreadsheet cell.

            Returns:
                The list of allowed mod values.

            """

            if pd.isna(cell_value):
                return []

            if cell_value.startswith("*"):
                return self.emphasize_specialization_dependencies(spreadsheet[f"{variable_name}s"][spreadsheet[f"{variable_name}s"][cell_value[1:]] == "✓"]["Stats"].tolist())

            return [cell_value]

        if weapon_class != "Signature Weapons":
            return cell_value_processor(spreadsheet[weapon_class][spreadsheet[weapon_class]["Name"] == self.de_emphasize_weapon_name(weapon_class, weapon_name)].iloc[0][variable_name])

        return []


    def get_allowed_weapon_talent_values(self, weapon_class: str, weapon_name: str, variable_name: str) -> List[str]:
        """
        Retrieves the allowed talent values for a specified weapon.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            weapon_class: The class of the weapon.
            weapon_name: The name of the weapon.
            variable_name: The name of the variable.

        Returns:
            The list of allowed talent values.

        """

        def cell_value_processor(cell_value: str) -> List[str]:
            """
            Processes a spreadsheet cell value to determine the allowed talent values.

            Parameters:
                cell_value: The value from the spreadsheet cell.

            Returns:
                The list of allowed talent values.

            """

            if pd.isna(cell_value):
                return []

            if cell_value.startswith("*"):
                return spreadsheet["Weapon Talents"][spreadsheet["Weapon Talents"][weapon_class] == "✓"]["Name"].tolist()

            return [cell_value]

        if weapon_class != "Signature Weapons":
            return cell_value_processor(spreadsheet[weapon_class][spreadsheet[weapon_class]["Name"] == self.de_emphasize_weapon_name(weapon_class, weapon_name)].iloc[0][variable_name])

        return [spreadsheet["Signature Weapons"][spreadsheet["Signature Weapons"]["Name"] == self.de_emphasize_weapon_name(weapon_class, weapon_name)].iloc[0]["Talent"]] if variable_name == "Talent 1" else []


    def get_allowed_gear_name_values(self, gear_type: str, gear_slot: str) -> List[str]:
        """
        Retrieves the allowed gear name values for a given gear type and slot.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            gear_type: The type of the gear piece.
            gear_slot: The name of the gear piece slot.

        Returns:
            The list of allowed name values.

        """

        if gear_type == "Improvised":
            return [f"Improvised {gear_slot}"]

        if gear_type in ["Brand Set", "Gear Set"]:
            return spreadsheet[f"{gear_type}s"]["Name"].tolist()

        return spreadsheet[f"{gear_type} Gear"][spreadsheet[f"{gear_type} Gear"]["Slot"] == gear_slot]["Name"].tolist()


    def get_allowed_gear_core_attribute_values(self, gear_type: str, gear_name: str, variable_name: str) -> List[str]:
        """
        Retrieves the allowed core attribute values for a specified gear piece.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            gear_type: The type of the gear piece.
            gear_name: The name of the gear piece.
            variable_name: The name of the variable.

        Returns:
            The list of allowed core attribute values.

        """

        def cell_value_processor(cell_value: str) -> List[str]:
            """
            Processes a spreadsheet cell value to determine the allowed core attribute values.

            Parameters:
                cell_value: The value from the spreadsheet cell.

            Returns:
                The list of allowed core attribute values.

            """

            if pd.isna(cell_value):
                return []

            if cell_value == "*":
                return spreadsheet["Gear Core Attributes"]["Stats"].tolist()

            if cell_value.startswith("!"):
                return remove_elements(spreadsheet["Gear Core Attributes"]["Stats"].tolist(), [cell_value[1:]])

            return [cell_value]

        if gear_type in ["Improvised", "Brand Set", "Gear Set"]:
            return spreadsheet["Gear Core Attributes"]["Stats"].tolist() if variable_name == "Core Attribute 1" else []

        return cell_value_processor(spreadsheet[f"{gear_type} Gear"][spreadsheet[f"{gear_type} Gear"]["Name"] == gear_name].iloc[0][variable_name])


    def get_allowed_gear_attribute_values(self, gear_type: str, variable_name: str, gear_name: str) -> List[str]:
        """
        Retrieves the allowed attribute values for a specified gear piece.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            gear_type: The type of the gear piece.
            variable_name: The name of the variable.
            gear_name: The name of the gear piece.

        Returns:
            The list of allowed attribute values.

        """

        def cell_value_processor(cell_value: str) -> List[str]:
            """
            Processes a spreadsheet cell value to determine the allowed attribute values.

            Parameters:
                cell_value: The value from the spreadsheet cell.

            Returns:
                The list of allowed attribute values.

            """

            if pd.isna(cell_value):
                return []

            if cell_value == "*":
                return spreadsheet["Gear Attributes"]["Stats"].tolist()

            if cell_value.startswith("*"):
                return spreadsheet["Gear Attributes"][spreadsheet["Gear Attributes"]["Category"] == cell_value[1:]]["Stats"].tolist()

            if cell_value.startswith("!"):
                return remove_elements(spreadsheet["Gear Attributes"]["Stats"].tolist(), [cell_value[1:]])

            return [cell_value]

        if gear_type in ["Improvised", "Brand Set"]:
            return spreadsheet["Gear Attributes"]["Stats"].tolist() if variable_name in ["Attribute 1", "Attribute 2"] else []

        if gear_type == "Gear Set":
            return spreadsheet["Gear Attributes"]["Stats"].tolist() if variable_name == "Attribute 1" else []

        return cell_value_processor(spreadsheet[f"{gear_type} Gear"][spreadsheet[f"{gear_type} Gear"]["Name"] == gear_name].iloc[0][variable_name])


    def get_allowed_gear_mod_values(self, gear_type: str, variable_name: str, gear_slot: str, gear_name: str) -> List[str]:
        """
        Retrieves the allowed mod values for a specified gear piece.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            gear_type: The type of the gear piece.
            variable_name: The name of the variable.
            gear_slot: The name of the gear piece slot.
            gear_name: The name of the gear piece.

        Returns:
            The list of allowed mod values.

        """

        def cell_value_processor(cell_value: str) -> List[str]:
            """
            Processes a spreadsheet cell value to determine the allowed mod values.

            Parameters:
                cell_value: The value from the spreadsheet cell.

            Returns:
                The list of allowed mod values.

            """

            if pd.isna(cell_value):
                return []

            return spreadsheet["Gear Mods"]["Stats"].tolist()

        if gear_type == "Improvised":
            return spreadsheet["Gear Mods"]["Stats"].tolist() if variable_name == "Mod 1" else []

        if gear_type in ["Brand Set", "Gear Set"]:
            return spreadsheet["Gear Mods"]["Stats"].tolist() if variable_name == "Mod 1" and gear_slot in ["Mask", "Body Armor", "Backpack"] else []

        return cell_value_processor(spreadsheet[f"{gear_type} Gear"][spreadsheet[f"{gear_type} Gear"]["Name"] == gear_name].iloc[0][variable_name])


    def get_allowed_gear_talent_values(self, gear_type: str, gear_slot: str, variable_name: str, gear_name: str) -> List[str]:
        """
        Retrieves the allowed talent values for a specified gear piece.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            gear_type: The type of the gear piece.
            gear_slot: The name of the gear piece slot.
            variable_name: The name of the variable.
            gear_name: The name of the gear piece.

        Returns:
            The list of allowed talent values.

        """

        def cell_value_processor(cell_value: str) -> List[str]:
            """
            Processes a spreadsheet cell value to determine the allowed talent values.

            Parameters:
                cell_value: The value from the spreadsheet cell.

            Returns:
                The list of allowed talent values.

            """

            if pd.isna(cell_value):
                return []

            return [cell_value]

        if gear_type in ["Improvised", "Brand Set"]:
            return spreadsheet["Gear Talents"][spreadsheet["Gear Talents"]["Slot"] == gear_slot]["Name"].tolist() if variable_name == "Talent 1" and gear_slot in ["Body Armor", "Backpack"] else []

        if gear_type == "Gear Set":
            return [spreadsheet["Gear Sets"][spreadsheet["Gear Sets"]["Name"] == gear_name].iloc[0][f"{gear_slot} Talent"]] if variable_name == "Talent 1" and gear_slot in ["Body Armor", "Backpack"] else []

        return cell_value_processor(spreadsheet[f"{gear_type} Gear"][spreadsheet[f"{gear_type} Gear"]["Name"] == gear_name].iloc[0][variable_name])


    def get_allowed_skill_name_values(self, skill_class: str) -> List[str]:
        """
        Retrieves the allowed skill name values for a given skill class.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            skill_class: The class of the skill.

        Returns:
            The list of allowed skill name values.

        """

        return self.emphasize_specialization_dependencies(spreadsheet["Skills"][spreadsheet["Skills"]["Class"] == skill_class].iloc[0]["Name"].split("; "))


    def get_allowed_skill_mod_values(self, skill_class: str, variable_name: str) -> tuple[str, List[str]]:
        """
        Retrieves the allowed skill mod values for a given skill class.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            skill_class: The class of the skill.
            variable_name: The name of the variable.

        Returns:
            A tuple containing the mod slot name and the list of allowed mod values.

        """

        inner_column_names = spreadsheet[f"{skill_class} Mods"].columns[1:-1]
        inner_column_index = int(variable_name.split()[-1]) - 1

        if inner_column_index >= len(inner_column_names):
            return None, []

        inner_column_name = inner_column_names[inner_column_index]
        return inner_column_name, self.emphasize_specialization_dependencies(spreadsheet[f"{skill_class} Mods"][spreadsheet[f"{skill_class} Mods"][inner_column_name] == "✓"]["Stats"].tolist())


    ####################

    def emphasize_specialization_dependencies(self, item_names: List[str]) -> List[str]:
        """
        Emphasizes the dependency of certain item names on the specialization selection.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            item_names: The list of item names.

        Returns:
            The list of item names with specialization dependencies emphasized.

        """

        specialization_value = self.specialization_variable.get()
        emphasized_item_names = []

        for item_name in item_names:
            if "(" in item_name:
                base_name = item_name[:item_name.rindex("(") - 1]
                parentheses_content = item_name[item_name.rindex("(") + 1:-1]

                if parentheses_content in SPECIALIZATIONS:
                    emphasized_item_names.append(f"{base_name} (Select {parentheses_content})" if parentheses_content != specialization_value else base_name)
                    continue

            emphasized_item_names.append(item_name)

        return emphasized_item_names


    def de_emphasize_weapon_name(self, weapon_class: str, weapon_name: str) -> str:
        """
        De-emphasizes the specialization dependency for a given weapon name.

        Parameters:
            self: The instance of the 'DamageCalculatorApp' class.
            weapon_class: The class of the weapon.
            weapon_name: The name of the weapon.

        Returns:
            The de-emphasized weapon name.

        """

        for candidate_name in spreadsheet[weapon_class]["Name"].tolist():
            if candidate_name.startswith(weapon_name):
                return candidate_name


#################### SECTION BREAK ####################

##### MAIN EXECUTION #####

if __name__ == "__main__":
    app = DamageCalculatorApp()
    app.mainloop()

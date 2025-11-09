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
data = {name: pd.read_excel(root_directory / "data" / f"{name}.xlsx", sheet_name=None) for name in ["Armor", "Skill", "Talent", "Weapon"]}


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
        
        # Set appearance mode and color theme
        ctk.set_appearance_mode("dark")  # Modes: "system", "light", "dark"
        ctk.set_default_color_theme("blue")  # Themes: "blue", "green", "dark-blue"
        
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
        label = ctk.CTkLabel(
            self.build_creator_frame,
            text="Build Creator",
            font=ctk.CTkFont(size=32, weight="bold")
        )
        label.pack(pady=50)
    
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
        self.build_creator_btn.configure(fg_color=("gray75", "gray25"))
        self.build_tuning_btn.configure(fg_color=["#3B8ED0", "#1F6AA5"])
        self.damage_output_btn.configure(fg_color=["#3B8ED0", "#1F6AA5"])
    
    def show_build_tuning(self):
        """Show the Build Tuning window"""
        if self.current_frame:
            self.current_frame.grid_forget()
        self.build_tuning_frame.grid(row=0, column=0, sticky="nsew")
        self.current_frame = self.build_tuning_frame
        
        # Update button states
        self.build_creator_btn.configure(fg_color=["#3B8ED0", "#1F6AA5"])
        self.build_tuning_btn.configure(fg_color=("gray75", "gray25"))
        self.damage_output_btn.configure(fg_color=["#3B8ED0", "#1F6AA5"])
    
    def show_damage_output(self):
        """Show the Damage Output window"""
        if self.current_frame:
            self.current_frame.grid_forget()
        self.damage_output_frame.grid(row=0, column=0, sticky="nsew")
        self.current_frame = self.damage_output_frame
        
        # Update button states
        self.build_creator_btn.configure(fg_color=["#3B8ED0", "#1F6AA5"])
        self.build_tuning_btn.configure(fg_color=["#3B8ED0", "#1F6AA5"])
        self.damage_output_btn.configure(fg_color=("gray75", "gray25"))


#################### SECTION BREAK ####################

##### MAIN EXECUTION #####

if __name__ == "__main__":
    app = DamageCalculatorApp()
    app.mainloop()

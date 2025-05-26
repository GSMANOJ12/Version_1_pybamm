# -*- coding: utf-8 -*-
"""
Created on Thurs Jan  1 10:02:20 2025

@author: INIYANK
"""

from tkinter import filedialog
from tkinter import messagebox
import customtkinter as CTk
import tkinter as Tk
import os
from PIL import Image

from Cell_Test_Plan_Generator import CellTestPlanGenerator
from Cell_Test_Plan_Generator import ProcessBlockForGUI

class CellTestPlanGeneratorGUI(object):
    def __init__(self, root):
        # Declare the necessary variables
        self.declare_varialbes(root)

        # Load the configuration file and populate the values
        self.load_configuration_file()

        # Build the GUI
        self.build_gui(root)

        # Load the first tab
        self.load_tab(self.current_tab)

    def declare_varialbes(self, root):
        self.gui_title = '' #GUI Title
        self.gui_sideframe_title = '' #Side-frame title
        self.gui_geometry = [] #[ Width, height ]
        self.gui_default_theme = '' #GUI Theme
        self.gui_color = '' #GUI Color
        self.opw_document_matching_string = ''
        self.ini_doc_matching_string = ''

        self.test_spec_file_name = ''
        self.operating_window_file_name = ''
        self.initialization_parameters_file_name = ''

        self.test_procedure_table_list = []
        self.cell_type_list = []
        self.cell_cycler_list = []  # List of the cyclers
        self.registration_parameters_list = []

        self.test_spec_file_input = ''
        self.operating_window_file_input = ''
        self.initialization_parameters_file_input = 'Supporting Documents\Initialization_Parameters_Document.xlsx'
        self.test_procedure_input = ''
        self.cell_type_input = ''
        self.cell_cycler_input = ''
        self.termination_input = 0
        self.multiple_test_plan_input = 0
        self.registration_parameters_input = []

        self.generate_tab_flag = 0

        self.test_plan_status_flag = 0

        self.root = root

    def load_configuration_file(self):
        print('\nLoading GUI Configuration file')

        # Lazy module imports
        import os
        import sys
        import json

        base_dir = 'Configuration File'

        config_path = os.path.join(base_dir, "config.json")

        # Read the config.json file
        try:
            with open(config_path, "r") as file:
                config = json.load(file)
        except FileNotFoundError:
            print(f"Error: config.json not found at {config_path}")
            sys.exit(1)

        # Load the configuration values
        self.gui_title = config["gui_configuration"]["gui_title"]
        self.gui_sideframe_title = config["gui_configuration"]["gui_sideframe_title"]
        self.gui_geometry = config["gui_configuration"]["gui_geometry"]
        self.gui_default_theme = config["gui_configuration"]["gui_default_mode"]
        self.gui_color = config["gui_configuration"]["gui_color"]
        self.opw_document_matching_string = config["doc_matching_string"]["opw_document_matching_string"]
        self.ini_doc_matching_string = config["doc_matching_string"]["ini_doc_matching_string"]

        self.cell_cycler_list = config["database"]["cell_cycler_list"]

    def build_gui(self, root):

        CTk.set_appearance_mode(self.gui_default_theme)
        CTk.set_default_color_theme(self.gui_color)
        CTk.deactivate_automatic_dpi_awareness()

        # Defining the Title name
        self.root.title(self.gui_title)

        # Definig the size of the GUI and disabling resizable option
        self.root.geometry(f"{self.gui_geometry[0]}x{self.gui_geometry[1]}")
        self.root.resizable(False, False)

        # Sidebar setup
        self.sidebar_width = 250
        self.sidebar_frame = CTk.CTkFrame(
            root, width=self.sidebar_width, corner_radius=0)
        self.sidebar_frame.pack(side="left", fill="y")

        # Setting the label
        self.logo_label = CTk.CTkLabel(
            self.sidebar_frame, text=self.gui_sideframe_title, font=("Inter,", 20, "bold"))
        self.logo_label.grid(padx=20, pady=(20, 10))

        # Adding logo
        self.image = CTk.CTkImage(Image.open("Icons\Daimler_logo.png"), size=(240, 60))  # Proper CTkImage usage
        self.image_label = CTk.CTkLabel(
            self.sidebar_frame, text="", image=self.image)
        self.image_label.grid(padx=20, pady=(20, 10))

        # Setting the option menu for changing the apperance
        self.appearance_mode_label = CTk.CTkLabel(
            self.sidebar_frame, 
            text="Appearance Mode:", 
            anchor="w"
        )
        self.appearance_mode_label.grid(padx=20, pady=(10, 0))
        self.appearance_mode_optionemenu = CTk.CTkOptionMenu(self.sidebar_frame, values=["Light", "Dark"],
                                                             command=self.change_appearance_mode_event)
        self.appearance_mode_optionemenu.grid(padx=20, pady=(10, 20))
        # self.sidebar_frame.grid_columnconfigure(0, weight=0)  # Label side (can expand)
        # self.sidebar_frame.grid_columnconfigure(1, weight=0)  # Button side (no expand)

        self.appearance_mode_optionemenu.set("Select Mode")

        self.configuration_label = CTk.CTkLabel(self.sidebar_frame, text="Select the Configuration", font=("Segoe UI,", 16, "bold"))
        self.configuration_label.grid(row=4, column=0, padx=(0, 0), pady=(10, 0))
        self.info_icon = CTk.CTkImage(Image.open("Icons\info_icon.png"), size=(30, 30))
        
        self.configuration_info_button = CTk.CTkButton(self.sidebar_frame, text="", image=self.info_icon, width=30,
                                        height=30, fg_color="transparent", hover_color="#E0E0E0", command=self.simulation_info_popup)
        self.configuration_info_button.grid(row=4, padx=(220, 0), pady=(20, 0))
        self.configuration_combobox = CTk.CTkComboBox(
            self.sidebar_frame, values=["Simulation","test plan"], command=self.select_config, width=150)
        self.configuration_combobox.grid(padx=20, pady=(1, 2))

        self.configuration_combobox.set("----Select Type------")
        self.configuration_textbox = CTk.CTkTextbox(
                    self.sidebar_frame, width=150, height=30)
        self.configuration_textbox.grid(padx=20, pady=(10, 20))

        self.configuration_confirm_button = CTk.CTkButton(
            self.sidebar_frame, text="Confrim your configuration type", command=self.config_func)
        self.configuration_confirm_button.grid(padx=20, pady=(10, 20))

        # Variable to store termination checkbox state across tabs
        self.termination_check = CTk.IntVar(value=0)

        self.multiple_test_plan_check = CTk.IntVar(value=0)

        self.simulation_check=CTk.IntVar(value=0)

        self.conf_type="simulation"

        # Initialize current tab index
        self.current_tab = 0

        # Create a container for main content frame
        self.content_frame = CTk.CTkFrame(root)
        self.content_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # Create a container for navigation frame
        self.create_navigation_frame(root)

        # self.create_navigation_frame(root)


        # Tab content definitions
        self.tabs_all = [
            {"content": self.select_input_fie_tab},
            {"content": self.select_procedure_table_tab},
            {"content": self.select_simulation_tab},
            {"content": self.select_cell_cycler_tab},
            {"content": self.select_registration_tab},
        ]

    def change_appearance_mode_event(self, new_appearance_mode: str):
        CTk.set_appearance_mode(new_appearance_mode)

    def create_navigation_frame(self, root):
        # Create navigation buttons container
        self.navigation_frame = CTk.CTkFrame(root)
        self.navigation_frame.pack(
            side="bottom", fill="x", padx=20, pady=(10, 20))

        # Create back and next button
        self.back_button = CTk.CTkButton(
            self.navigation_frame, text="<Back", command=self.go_back, state="disabled")
        self.next_button = CTk.CTkButton(
            self.navigation_frame, text="Next", command=self.go_next)

        self.back_button.pack(side="left", padx=(5, 0))
        self.next_button.pack(side="right", padx=(0, 5))

        self.generate_button = CTk.CTkButton(
            self.navigation_frame, text="Generate", command=self.process_generate)
        self.generate_button.pack_forget()

    def load_tab(self, index):
        # Clear current content
        for widget in self.content_frame.winfo_children():
            widget.destroy()
        print(self.conf_type)
        if self.conf_type.lower()=="simulation":
               self.tabs=self.tabs_all[:3]
        else:
               self.tabs=self.tabs_all[:2]+self.tabs_all[3:]
        # Load the content for the specified tab
        tab_data = self.tabs[index]
        tab_data["content"]()

        # Update button states
        self.back_button.configure(state="normal" if index > 0 else "disabled")
        self.next_button.configure(
            state="normal" if index < len(self.tabs) - 1 else "disabled")

    def select_input_fie_tab(self):

        self.input_tab_main_label = CTk.CTkLabel(
            self.content_frame, text="Select the Input Files", font=("Segoe UI,", 16, "bold"))
        self.input_tab_main_label.grid(padx=(80, 0), pady=(20, 0))

        self.input_tab_test_spec_label = CTk.CTkLabel(
            self.content_frame, text="Test specification", font=("Segoe UI,", 14, "bold"))
        self.input_tab_test_spec_label.grid(padx=(80, 0), pady=(10, 0))

        self.input_tab_test_spec_textbox = CTk.CTkTextbox(
            self.content_frame, width=300, height=50)
        self.input_tab_test_spec_textbox.grid(padx=(80, 0), pady=(10, 0))

        self.input_tab_opw_label = CTk.CTkLabel(
            self.content_frame, text="Operating Window", font=("Segoe UI,", 14, "bold"))
        self.input_tab_opw_label.grid(padx=(80, 0), pady=(20, 0))

        self.input_tab_opw_textbox = CTk.CTkTextbox(
            self.content_frame, width=300, height=50)
        self.input_tab_opw_textbox.grid(padx=(80, 0), pady=(10, 0))

        self.input_tab_ini_doc_label = CTk.CTkLabel(
            self.content_frame, text="Initialization document", font=("Segoe UI,", 14, "bold"))
        self.input_tab_ini_doc_label.grid(padx=(80, 0), pady=(20, 0))

        self.input_tab_ini_doc_textbox = CTk.CTkTextbox(
            self.content_frame, width=300, height=50)
        self.input_tab_ini_doc_textbox.grid(padx=(80, 0), pady=(10, 0))

        self.combobox_1 = CTk.CTkButton(
            self.content_frame, text=" Browse Files ", command=self.process_browse_files)
        self.combobox_1.grid(padx=(80, 0), pady=(20, 0))
        
        if (self.test_spec_file_name):
            self.input_tab_test_spec_textbox.delete("0.0", "end")
            self.input_tab_test_spec_textbox.insert("0.0", self.test_spec_file_name)
        else:
            self.input_tab_test_spec_textbox.delete("0.0", "end")
            self.input_tab_test_spec_textbox.insert("0.0", "No Test specification selected")

        if (self.operating_window_file_name):
            self.input_tab_opw_textbox.delete("0.0", "end")
            self.input_tab_opw_textbox.insert("0.0", self.operating_window_file_name)
        else:
            self.input_tab_opw_textbox.delete("0.0", "end")
            self.input_tab_opw_textbox.insert("0.0", "No Operating window selected")

        self.initialization_parameters_file_name = os.path.basename(self.initialization_parameters_file_input)
        self.input_tab_ini_doc_textbox.insert(
            "0.0", f"{self.initialization_parameters_file_name}")

        self.check_display_generate()

    def select_procedure_table_tab(self):
        self.selection_tab_test_procedure_label = CTk.CTkLabel(
            self.content_frame, text="Select the Test Procedure", font=("Segoe UI,", 16, "bold"))
        self.selection_tab_test_procedure_label.grid(padx=(80, 0), pady=(20, 0))

        self.selection_tab_test_procedure_textbox = CTk.CTkTextbox(
            self.content_frame, width=300, height=60)
        self.selection_tab_test_procedure_textbox.grid(padx=(80, 0), pady=(10, 0))

        self.selection_tab_test_procedure_combobox = CTk.CTkComboBox(
            self.content_frame, values=self.test_procedure_table_list, width=200, command=self.process_test_procedure_table)
        self.selection_tab_test_procedure_combobox.grid(padx=(80, 0), pady=(20, 0))

        self.selection_tab_cell_type_label = CTk.CTkLabel(
            self.content_frame, text="Select the Cell Type", font=("Segoe UI,", 16, "bold"))
        self.selection_tab_cell_type_label.grid(padx=(80, 0), pady=(40, 0))

        self.selection_tab_cell_type_textbox = CTk.CTkTextbox(
            self.content_frame, width=200, height=30)
        self.selection_tab_cell_type_textbox.grid(padx=(80, 0), pady=(10, 0))

        self.combobox_2_2 = CTk.CTkComboBox(
            self.content_frame, values=self.cell_type_list, width=200, command=self.process_cell_type)
        self.combobox_2_2.grid(padx=(80, 0), pady=(20, 0))

        self.selection_tab_test_procedure_combobox.set("Select Test Procedure")
        self.combobox_2_2.set("Select Cell Type")

        if (self.test_procedure_input):
            self.selection_tab_test_procedure_textbox.delete("0.0", "end")
            self.selection_tab_test_procedure_textbox.insert("0.0", self.test_procedure_input)
        else:
            self.selection_tab_test_procedure_textbox.delete("0.0", "end")
            self.selection_tab_test_procedure_textbox.insert("0.0", "No Test Procedure selected")

        if (self.cell_type_input):
            self.selection_tab_cell_type_textbox.delete("0.0", "end")
            self.selection_tab_cell_type_textbox.insert("0.0", self.cell_type_input)
        else:
            self.selection_tab_cell_type_textbox.delete("0.0", "end")
            self.selection_tab_cell_type_textbox.insert("0.0", "No Cell Type selected")

        self.check_display_generate()

    def select_cell_cycler_tab(self):
        self.cell_cycler_tab_cycler_label = CTk.CTkLabel(
            self.content_frame, text="Select the Cycler Input", font=("Segoe UI,", 16, "bold"))
        self.cell_cycler_tab_cycler_label.grid(padx=(90, 0), pady=(20, 0))

        self.cell_cycler_tab_cycler_textbox = CTk.CTkTextbox(
            self.content_frame, width=150, height=30)
        self.cell_cycler_tab_cycler_textbox.grid(padx=(90, 0), pady=(20, 0))

        self.cell_cycler_tab_cycler_combobox = CTk.CTkComboBox(
            self.content_frame, values=self.cell_cycler_list, command=self.process_cell_cycler)
        self.cell_cycler_tab_cycler_combobox.grid(padx=(90, 0), pady=(20, 0))

        self.cell_cycler_tab_termination_label = CTk.CTkLabel(
            self.content_frame, text="Select the Termination Configuration", font=("Segoe UI,", 16, "bold"))
        self.cell_cycler_tab_termination_label.grid(padx=(90, 0), pady=(40, 0))

        self.cell_cycler_tab_termination_checkbox = CTk.CTkCheckBox(self.content_frame, text="Add Termination",
                                          variable=self.termination_check, command=self.process_termination_input)
        self.cell_cycler_tab_termination_checkbox.grid(row=5, padx=(60, 0), pady=(20, 0))

        self.cell_cycler_tab_test_plan_creation_label = CTk.CTkLabel(
            self.content_frame, text="Select the Test Plan Configuration", font=("Segoe UI,", 16, "bold"))
        self.cell_cycler_tab_test_plan_creation_label.grid(padx=(90, 0), pady=(40, 0))

        self.cell_cycler_tab_test_plan_creation_checkbox = CTk.CTkCheckBox(self.content_frame, text="Create Multiple Test Plan",
                                          variable=self.multiple_test_plan_check, command=self.process_multiple_test_plan)
        self.cell_cycler_tab_test_plan_creation_checkbox.grid(row=7, padx=(60, 0), pady=(20, 0))

        # self.update_termination_checkbox()

        # Information Icon (Replace with your 'info_icon.png')
        self.info_icon = CTk.CTkImage(Image.open(
            "Icons\info_icon.png"), size=(30, 30))
        
        self.termination_info_button = CTk.CTkButton(self.content_frame, text="", image=self.info_icon, width=30,
                                         height=30, fg_color="transparent", hover_color="#E0E0E0", command=self.show_termination_info_popup)
        self.termination_info_button.grid(row=5, padx=(250, 0), pady=(20, 0))

        self.multiple_test_plan_info_button = CTk.CTkButton(self.content_frame, text="", image=self.info_icon, width=30,
                                         height=30, fg_color="transparent", hover_color="#E0E0E0", command=self.show_multiple_test_plan_info_popup)
        self.multiple_test_plan_info_button.grid(row=7, padx=(300, 0), pady=(20, 0))

        self.cell_cycler_tab_cycler_combobox.set("Select Cycler")

        if (self.cell_cycler_input):
            self.cell_cycler_tab_cycler_textbox.delete("0.0", "end")
            self.cell_cycler_tab_cycler_textbox.insert("0.0", self.cell_cycler_input)
        else:
            self.cell_cycler_tab_cycler_textbox.delete("0.0", "end")
            self.cell_cycler_tab_cycler_textbox.insert("0.0", "No Cycler selected")

        self.check_display_generate()

    def select_registration_tab(self):
        self.registration_parameters_list = CellTestPlanGenerator.get_registration_parameters(
            self.initialization_parameters_file_input)

        self.registration_tab_registration_parameters_label = CTk.CTkLabel(
            self.content_frame, text="Select Registration parameters", font=("Segoe UI,", 16, "bold"))
        self.registration_tab_registration_parameters_label.grid(row=0, column=0, columnspan=2, padx=(
            60, 0), pady=(20, 0), sticky="ew")

        # Left Label and Listbox (Available Items)
        self.registration_tab_available_parameters_label = CTk.CTkLabel(
            self.content_frame, text="Available Parameters", font=("Segoe UI,", 14, "bold"))
        self.registration_tab_available_parameters_label.grid(row=1, column=0, padx=(
            0, 0), pady=(20, 0), sticky="ew")

        self.registration_tab_available_parameters_list = Tk.Listbox(
            self.content_frame, selectmode=Tk.MULTIPLE, height=14, exportselection=0)
        self.registration_tab_available_parameters_list.grid(
            row=2, column=0, padx=60, pady=5, sticky="ew")

        # Right Label and Listbox (Selected Items)
        self.registration_tab_selected_parameters_label = CTk.CTkLabel(
            self.content_frame, text="Selected Parameters", font=("Segoe UI,", 14, "bold"))
        self.registration_tab_selected_parameters_label.grid(row=1, column=1, padx=(
            0, 0), pady=(20, 0), sticky="ew")

        self.registration_tab_selected_parameters_list = Tk.Listbox(
            self.content_frame, selectmode=Tk.MULTIPLE, height=14, exportselection=0)
        self.registration_tab_selected_parameters_list.grid(
            row=2, column=1, padx=10, pady=5, sticky="ew")

        self.registration_tab_add_button = CTk.CTkButton(
            self.content_frame, text="Add →", command=self.add_selected_list_box_4)
        self.registration_tab_add_button.grid(row=3, column=0, padx=10, pady=5)

        self.registration_tab_remove_button = CTk.CTkButton(
            self.content_frame, text="← Remove", command=self.remove_selected_listbox_4)
        self.registration_tab_remove_button.grid(row=3, column=1, padx=10, pady=5)

        self.registration_tab_confirm_button = CTk.CTkButton(
            self.content_frame, text="Confirm", command=self.confirm_selected_registration_parameters)
        self.registration_tab_confirm_button.grid(columnspan=2, padx=(50, 0), pady=(10, 0))

        # Insert items into Available Listbox
        for param in self.registration_parameters_list:
            self.registration_tab_available_parameters_list.insert(Tk.END, param)

        self.check_display_generate()


    def select_simulation_tab(self):

    #     self.simulation_checkbox = CTk.CTkCheckBox(self.content_frame, text="Click on this for Simulation", font=("Segoe UI,", 25, "bold"),variable=self.simulation_check,command=self.display_simulation)
    #     self.simulation_checkbox.grid(row=0, column=2, padx=(70, 0), pady=(20, 0))

    # def display_simulation(self):
    #     if(self.simulation_check.get()==1):
            self.simulation_label = CTk.CTkLabel(self.content_frame, text="Select the Simulation Type", font=("Segoe UI,", 16, "bold"))
            self.simulation_label.grid(row=1, column=2, padx=(50, 0), pady=(20, 0))

            self.info_icon_1 = CTk.CTkImage(Image.open("Icons\info_icon.png"), size=(30, 30))
            
            self.simulation_info_button = CTk.CTkButton(self.content_frame, text="", image=self.info_icon_1, width=30,
                                            height=30, fg_color="transparent", hover_color="#E0E0E0", command=self.simulation_info_popup)
            self.simulation_info_button.grid(row=1, column=3, padx=(0, 0), pady=(20, 0))

            self.simulation_combobox = CTk.CTkComboBox(
                self.content_frame, values=["Pybamm","P2D Simulation","other"], command=self.select_simulation, width=250)
            self.simulation_combobox.grid(row=2, column=1, columnspan=2, padx=(90, 0), pady=(20, 0))

            self.simulation_combobox.set("-------------Select Model Type------------")
            self.simulation_textbox = CTk.CTkTextbox(
                        self.content_frame, width=150, height=30)
            self.simulation_textbox.grid(row=3, column=1, columnspan=2, padx=(90, 0), pady=(20, 0))

            self.simulation_confirm_button = CTk.CTkButton(
                self.content_frame, text="Proceed for Simulation", command=self.simulation_func)
            self.simulation_confirm_button.grid(row=4, column=1, columnspan=2, padx=(90, 0), pady=(20, 0))

        # else:
        #     self.simulation_label.grid_forget()
        #     self.simulation_info_button.grid_forget()
        #     self.simulation_combobox.grid_forget()
        #     self.simulation_textbox.grid_forget()
        #     self.simulation_confirm_button.grid_forget()


    def select_simulation(self,sim_type):
        self.sim_type=sim_type
        self.simulation_textbox.delete("0.0", "end")
        self.simulation_textbox.insert("0.0", self.sim_type)
    
    def config_func(self):
        if (self.conf_type.lower()=="simulation"):
            print("okokokokok")
            
    def simulation_func(self):
        if (self.sim_type.lower()=="pybamm"):
           CellTestPlanGenerator.simulation_function("data")
        

    def process_browse_files(self):
        # Initialize the list to store file paths
        self.selected_file_paths = []

        current_directory = os.getcwd()

        file_paths = filedialog.askopenfilenames(initialdir="current_directory", title="Select a File",
                                                 filetypes=(("All files", "*.*"), ("Text files", "*.docx*")))

        # Proceed opening the file when selected (to avoid: PackageNotFoundError)
        if (file_paths):
            # Add selected file paths to the list
            self.selected_file_paths.extend(file_paths)
            for file_path in self.selected_file_paths:
                # Split the extension from the file path
                file_extension = os.path.splitext(file_path)[1]

                allowed_extensions = [".docx", ".xlsx"]

                if file_extension in allowed_extensions[0]:
                    self.test_spec_file_input = file_path
                    self.test_spec_file_name = os.path.basename(file_path)

                    print(
                        f'\nSelected Test Specification: {self.test_spec_file_name}')

                    # Process the test spec document and get the test procedure tables
                    self.test_procedure_table_list = ProcessBlockForGUI.get_test_procedure_inputs(
                        file_path)

                elif file_extension in allowed_extensions[1]:
                    if self.opw_document_matching_string in file_path.lower():
                        self.operating_window_file_input = file_path
                        self.operating_window_file_name = os.path.basename(
                            file_path)

                        print(
                            f'\nSelected Operating Window: {self.operating_window_file_name}')

                        # Process the operating window document and get the cell types
                        self.cell_type_list = ProcessBlockForGUI.read_operating_window_file(
                            file_path)

                    if self.ini_doc_matching_string in file_path.lower():
                        self.initialization_parameters_file_input = file_path
                        self.initialization_parameters_file_name = os.path.basename(
                            file_path)

                        print(
                            f'\nSelected Initialization parameters document: {self.initialization_parameters_file_name}')

        if (self.test_spec_file_name):
            self.input_tab_test_spec_textbox.delete("0.0", "end")
            self.input_tab_test_spec_textbox.insert("0.0", self.test_spec_file_name)
        else:
            self.input_tab_test_spec_textbox.delete("0.0", "end")
            self.input_tab_test_spec_textbox.insert("0.0", "No Test specification selected")

        if (self.operating_window_file_name):
            self.input_tab_opw_textbox.delete("0.0", "end")
            self.input_tab_opw_textbox.insert("0.0", self.operating_window_file_name)
        else:
            self.input_tab_opw_textbox.delete("0.0", "end")
            self.input_tab_opw_textbox.insert("0.0", "No Operating window selected")

        if (self.initialization_parameters_file_name):
            self.input_tab_ini_doc_textbox.delete("0.0", "end")
            self.input_tab_ini_doc_textbox.insert(
                "0.0", self.initialization_parameters_file_name)

        self.check_display_generate()

    def process_test_procedure_table(self, value):
        self.test_procedure_input = value

        print(f'\nSelected Test procedure table: {self.test_procedure_input}')

        self.selection_tab_test_procedure_textbox.delete("0.0", "end")
        self.selection_tab_test_procedure_textbox.insert("0.0", self.test_procedure_input)

        self.check_display_generate()

    def process_cell_type(self, value):
        self.cell_type_input = value

        print(f'\nSelected Cell type: {self.cell_type_input}')

        self.selection_tab_cell_type_textbox.delete("0.0", "end")
        self.selection_tab_cell_type_textbox.insert("0.0", self.cell_type_input)

        self.check_display_generate()

    def process_cell_cycler(self, value):
        self.cell_cycler_input = value

        print(f'\nSelected Cell cycler: {self.cell_cycler_input}')

        self.cell_cycler_tab_cycler_textbox.delete("0.0", "end")
        self.cell_cycler_tab_cycler_textbox.insert("0.0", self.cell_cycler_input)

        self.check_display_generate()

    def update_termination_checkbox(self):
        # Ensure both checkboxes reflect the same state (Checked or Unchecked)
        if self.termination_check.get() == 1:
            self.cell_cycler_tab_termination_checkbox.select()
            self.cell_cycler_tab_termination_checkbox.select()
        else:
            self.cell_cycler_tab_termination_checkbox.deselect()
            self.cell_cycler_tab_termination_checkbox.deselect()
        
        if self.multiple_test_plan_check.get() == 1:
            self.cell_cycler_tab_test_plan_creation_checkbox.select()
            self.cell_cycler_tab_test_plan_creation_checkbox.select()
        else:
            self.cell_cycler_tab_test_plan_creation_checkbox.deselect()
            self.cell_cycler_tab_test_plan_creation_checkbox.deselect()

    def process_termination_input(self):

        self.termination_input = int(self.termination_check.get())

        if (self.termination_input == 1):
            print(f'\nTermination selected')
        else:
            self.termination_input = 0
            print(f'\nTermination not selected')
    
    def process_multiple_test_plan(self):

        self.multiple_test_plan_input = int(self.multiple_test_plan_check.get())

        if (self.multiple_test_plan_input == 1):
            print(f'\nMultiple test plan selected')
        else:
            self.multiple_test_plan_input = 0
            print(f'\nMultiple test plan not selected')
    
    def select_config(self,conf_type):
            self.conf_type=conf_type
            self.configuration_textbox.delete("0.0", "end")
            self.configuration_textbox.insert("0.0", self.conf_type)       

    def simulation_info_popup(self):
        # Display a popup message when the info button is clicked
        messagebox.showinfo(
            "Information", "Simulation:\nPlease select the option in following dropdown model and output graph will be created based accuracy of the model")

    def show_termination_info_popup(self):
        # Display a popup message when the info button is clicked
        messagebox.showinfo(
            "Information", "Termination:\nPlease select the checkbox if no test is planned after this")
    
    def show_multiple_test_plan_info_popup(self):
        # Display a popup message when the info button is clicked
        messagebox.showinfo(
            "Information", "\nPlease select the checkbox if multiple test plan has to be genrated for different temperature\n⚠️Note: Mandatory for DTC-P-6_x_Internal Resistance")

    def add_selected_list_box_4(self):
        selected_indices = self.registration_tab_available_parameters_list.curselection()
        selected_items = [self.registration_tab_available_parameters_list.get(
            i) for i in selected_indices]

        # Add to Selected Listbox if not already present
        existing_items = self.registration_tab_selected_parameters_list.get(0, Tk.END)
        for item in selected_items:
            if item not in existing_items:
                self.registration_tab_selected_parameters_list.insert(Tk.END, item)

    def remove_selected_listbox_4(self):
        # Copy selected indices
        selected_indices = list(self.registration_tab_selected_parameters_list.curselection())
        selected_indices.reverse()  # Reverse to avoid index shifting issues

        for i in selected_indices:
            self.registration_tab_selected_parameters_list.delete(i)

    def confirm_selected_registration_parameters(self):
        self.registration_parameters_input = list(
            self.registration_tab_selected_parameters_list.get(0, Tk.END))

        if not self.registration_parameters_input:
            messagebox.showwarning(
                "Warning!", "Please select at least one parameter for registration")
        else:
            print("\nSelected Registration Parameters:",
                  self.registration_parameters_input)

        self.check_display_generate()

    def go_back(self):
        if self.current_tab > 0:
            self.current_tab -= 1
            self.load_tab(self.current_tab)

            self.check_display_generate()

    def go_next(self):
        if self.current_tab < len(self.tabs) - 1:
            self.current_tab += 1
            self.load_tab(self.current_tab)

            self.check_display_generate()

    def check_display_generate(self):
        self.generate_tab_flag = any(
            self.test_procedure_input and self.cell_cycler_input and self.cell_type_input and self.registration_parameters_input)
        
        if (self.generate_tab_flag):
            self.generate_button.pack(side="left", padx=20, pady=10, expand=True)

    def process_generate(self):
        response_input = messagebox.askokcancel("Confirmation", "Do you want to generate test plan?")

        if (response_input):
            CellTestPlanGenerator.execute_main_generate_function(self.test_spec_file_input, self.operating_window_file_input, self.initialization_parameters_file_input,
                                                             self.test_procedure_input, self.cell_type_input, self.cell_cycler_input, self.termination_input, self.multiple_test_plan_input, self.registration_parameters_input)
        
        flag_input = ProcessBlockForGUI.check_test_plan_status()

        if (flag_input == 1):
            messagebox.showinfo("Information", "✅ Test Plan Generated Successfully!")
        else:
            messagebox.showerror("Error", "Something went wrong!")
        
        self.clear_all_inputs()
        
    def clear_all_inputs(self):
        self.test_spec_file_name = ''
        self.operating_window_file_name = ''

        self.test_spec_file_input = ''
        self.operating_window_file_input = ''
        self.test_procedure_input = ''
        self.cell_type_input = ''
        self.cell_cycler_input = ''
        self.termination_input = 0
        self.multiple_test_plan_input = 0
        self.registration_parameters_input = []

        self.registration_tab_selected_parameters_list.delete(0, Tk.END)

        self.generate_tab_flag = 0

        self.termination_check.set(0)
        self.multiple_test_plan_check.set(0)

        self.generate_button.pack_forget()

def main():
    # Creating an object frame for GUI
    root = CTk.CTk()
    print('\nStarting GUI...')
    CellTestPlanGeneratorGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()

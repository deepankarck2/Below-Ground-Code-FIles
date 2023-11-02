"""
This module contains code for a PyQt5 application that displays a map with markers for each bus. 
The user can edit bus values and run simulations. It also contains code for creating the map using Folium. 
The map is saved as an HTML file, which is then loaded in the PyQt5 application.
There are two CSV files containing bus values and coordinates(values.txt, and sample_for_x_y.txt). Attached are sample CSV files.

The application has the following features:
- A search box with autocomplete for searching buses.
- A dropdown for selecting buses.
- A form for editing bus values.
- A button for submitting changes.
- A button for running simulations.
- A status bar for showing messages to the user.

Instructions To Run:
1) Install the required packages:
    ``` pip install PyQt5 folium ```
2) Download the CSV files and save them in the same directory as this file.
3) Run the code using python qt_folium.py.

"""
import os
import re
import sys
import win32com.client

from PyQt5.QtCore import QUrl, QTimer, Qt
from PyQt5.QtWebEngineWidgets import QWebEngineView
from PyQt5.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QCompleter,
    QDockWidget,
    QDoubleSpinBox,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QPushButton,
    QSizePolicy,
    QSpinBox,
    QVBoxLayout,
    QWidget,
    QDialog,
    QFormLayout,
)
import folium
import random

# Absolute paths for the CSV files
VALUES_CSV_FILE = "A:\\CCNY\\J_Fall_2023\\SD2\\OpenDSS\\load_values.csv"
SAMPLE_XY_FILE = "A:\\CCNY\\J_Fall_2023\\SD2\\OpenDSS\\sample_for_x_y.txt"
MAP_HTML_FILE = "A:\\CCNY\\J_Fall_2023\\SD2\\OpenDSS\\map.html"
GENERATOR_CSV_FILE = "A:\\CCNY\\J_Fall_2023\\SD2\\OpenDSS\\generator_values.csv"
LINES_CSV_FILE = "lines_values.csv"

FILE_PATH = "'A:\CCNY\J_Fall_2023\SD2\OpenDSS\IEEE 30 Bus\Master.dss'"


# Store the current working directory before calling the function
cwd_before = os.getcwd()


def setup_opendss():
    dssObj = win32com.client.Dispatch("OpenDSSEngine.DSS")

    # Start the DSS
    if not dssObj.Start(0):
        print("DSS failed to start!")
        exit()

    dssText = dssObj.Text
    dssCircuit = dssObj.ActiveCircuit
    dssElem = dssCircuit.ActiveCktElement
    dssSolution = dssCircuit.Solution
    dssText.Command = f"compile {FILE_PATH}"  # Load the circuit
    return dssObj, dssText, dssCircuit, dssElem, dssSolution


def load_bus_data(dssCircuit):
    """
    Load bus data from CSV files.

    Returns:
        Tuple: A tuple containing two dictionaries. The first dictionary contains bus values, and the second dictionary contains bus coordinates.
    """

    load_values = {}
    bus_coords = {}
    generator_values = {}
    line_values = {}
    print("CWD before initialization:", os.getcwd())

    with open(LINES_CSV_FILE, "r") as file:
        lines = file.readlines()
        for i, line in enumerate(lines):
            if i == 0:
                continue  # Skip the header line
            sline = line.strip().split(",")
            line_name = sline[0]
            bus1 = sline[1]
            bus2 = sline[2]

            line_values[line_name] = {"Bus1": bus1, "Bus2": bus2}

    load_values = {}  # Initialize empty dictionary

    i = dssCircuit.Loads.First
    while i:
        load_name = dssCircuit.Loads.Name
        bus1 = load_name  # dssCircuit.Loads.Bus1
        kv = (
            dssCircuit.Loads.kv
        )  # This fetches the base kV, ensure this is what you want
        kw = dssCircuit.Loads.kW
        kvar = dssCircuit.Loads.kvar

        # Store the values in the dictionary
        load_values[load_name] = {
            "bus": bus1,
            "kv": kv,
            "kw": kw,
            "kvar": kvar,
        }

        # Move to the next load
        i = dssCircuit.Loads.Next

    print(load_values)

    with open(f"{SAMPLE_XY_FILE}", "r") as file:
        lines = file.readlines()
        for line in lines:
            sline = line.split(",")
            bus_coords[sline[0]] = {"lat": float(sline[1]), "lon": float(sline[2])}

    with open(GENERATOR_CSV_FILE, "r") as file:
        lines = file.readlines()
        for i, line in enumerate(lines):
            if i == 0:  # skip the header line
                continue
            sline = line.strip().split(
                ","
            )  # strip to remove any trailing newline characters
            generator_values[sline[0]] = {
                "Bus1": sline[1],
                "kV": float(sline[2]),
                "kW": float(sline[3]),
                "Model": int(sline[4]),
                "Vpu": float(sline[5]),
                "Maxkvar": float(sline[6]),
                "Minkvar": float(sline[7]),
                "kvar": float(sline[8]),
            }

    return load_values, bus_coords, generator_values, line_values


def create_map(load_values, bus_coords, generator_values, line_values):
    all_lats = []
    all_lons = []

    m = folium.Map(
        location=[
            list(bus_coords.values())[0]["lat"],
            list(bus_coords.values())[0]["lon"],
        ],
        zoom_start=10,
    )

    # draw markers for B6,B8, B12, B28, B13
    for bus, coord in bus_coords.items():
        if bus in ["B6", "B8", "B28", "B13"]:
            folium.Marker(
                [coord["lat"], coord["lon"]],
                popup=bus,
                tooltip=bus,
                # icon=folium.Icon(color="purple"),
                icon=folium.Icon(color="purple", icon="bolt", prefix="fa"),
            ).add_to(m)
            all_lats.append(coord["lat"])
            all_lons.append(coord["lon"])

    # For Transmission Lines
    colors = ["red", "blue", "green", "orange", "purple", "pink"]
    for line, values in line_values.items():
        bus1_coord = bus_coords[values["Bus1"]]
        bus2_coord = bus_coords[values["Bus2"]]

        line_coords = [
            (bus1_coord["lat"], bus1_coord["lon"]),
            (bus2_coord["lat"], bus2_coord["lon"]),
        ]
        folium.PolyLine(line_coords, color=colors[random.randint(0, 4)]).add_to(m)

    # For Buses (Loads)
    for bus, coord in bus_coords.items():
        if bus.lower() in load_values:
            values = load_values[bus.lower()]
            popup_content = (
                f"<div style='font-size: 12px'>"
                f"<strong style='color: blue'>Bus Name: {bus}</strong><br>"
                f"<ul>"
                f"<li>kw: {values['kw']}</li>"
                f"<li>kvar: {values['kvar']}</li>"
                f"<li>kv: {values['kv']}</li>"
                f"</ul>"
                f"</div>"
            )

            marker = folium.Marker(
                [coord["lat"], coord["lon"]],
                popup=folium.Popup(popup_content, max_width=300),
                tooltip=bus,
            )
            m.add_child(marker)
            all_lats.append(coord["lat"])
            all_lons.append(coord["lon"])

    # For Generators
    for gen, values in generator_values.items():
        coord = bus_coords[values["Bus1"]]
        popup_content = (
            f"<div style='font-size: 12px'>"
            f"<strong style='color: red'>Generator Name: {gen}</strong><br>"
            f"<ul>"
            f"<li>kW: {values['kW']}</li>"
            f"<li>kvar: {values['kvar']}</li>"
            f"<li>kV: {values['kV']}</li>"
            f"</ul>"
            f"</div>"
        )

        marker = folium.Marker(
            [coord["lat"], coord["lon"]],
            popup=folium.Popup(popup_content, max_width=300),
            tooltip=gen,
            icon=folium.Icon(color="red"),  # Different icon for generators
        )
        m.add_child(marker)

    # Adjusting map bounds
    sw = [min(all_lats), min(all_lons)]
    ne = [max(all_lats), max(all_lons)]
    m.fit_bounds([sw, ne])

    # Save the map as an HTML file
    m.save(MAP_HTML_FILE)


def extract_numbers(s):
    """Extract all integers from a string and return them as a tuple using regular expression"""
    return tuple(map(int, re.findall(r"\d+", s)))


def custom_sort(item):
    """Custom sorting key that sorts by the extracted numbers."""
    return extract_numbers(item)


def run_simulation(load_values):
    """
    Run a simulation with the given bus values.

    Args:
        load_values (dict): A dictionary containing bus values.

    Returns:
        None
    """
    # For now, just print the values
    print(load_values)


class BusEditor(QWidget):
    """
    A widget for editing bus values and running simulations.
    """

    def __init__(self, load_values, message_label):
        super().__init__()

        self.load_values = load_values
        self.message_label = message_label
        self.temp_changes = {}  # Temporarily store changes before simulation
        self.layout = QVBoxLayout()

        # -------------------------------- Search -------------------------------------------- #
        # Create a new QVBoxLayout for the search box and its label
        search_layout = QVBoxLayout()
        search_layout.setSpacing(10)  # Adjust the value for gap between widgets

        # Create the label and set its size policy
        label = QLabel("Search for a load...")
        label.setSizePolicy(
            QSizePolicy.Preferred, QSizePolicy.Fixed
        )  # Set vertical size policy to Fixed
        search_layout.addWidget(label)

        self.search_box = QLineEdit(self)
        self.completer = QCompleter(list(self.load_values.keys()), self)
        self.search_box.setCompleter(self.completer)
        self.search_box.setPlaceholderText("Type to search...")
        self.search_box.setSizePolicy(
            QSizePolicy.Preferred, QSizePolicy.Fixed
        )  # Set vertical size policy to Fixed
        search_layout.addWidget(self.search_box)

        # Add the new layout to the main layout
        self.layout.addLayout(search_layout)

        # -------------------------------- Dropdown ------------------------------------------ #
        # Create dropdown field for load
        self.load_search = QComboBox(self)
        sorted_loads = sorted(load_values.keys(), key=custom_sort)
        self.load_search.addItems(sorted_loads)
        self.load_search.currentIndexChanged.connect(self.populate_values)
        self.layout.addWidget(QLabel("Select Load:"))
        self.layout.addWidget(self.load_search)

        # ---------------------------------- Show Selected Bus ----------------------------------- #
        # After the dropdown creation code
        self.selected_load_layout = QHBoxLayout()
        self.selected_load_display = QLabel(self)
        self.selected_load_display.setStyleSheet("color: red; font-weight: bold;")
        self.selected_load_layout.addWidget(QLabel("Selected Load:"))
        self.selected_load_layout.addWidget(self.selected_load_display)
        self.layout.addLayout(self.selected_load_layout)

        # ------------------------------------------------------------------------------------- #

        # Add a horizontal line
        hline = QFrame(self)
        hline.setFrameShape(QFrame.HLine)
        hline.setFrameShadow(QFrame.Sunken)
        self.layout.addWidget(hline)

        # ---------------------------------- Change Bus Values --------------------------------------- #
        # Create fields for load attributes
        self.kw_label = QLabel("0.0", self)
        self.kvar_label = QLabel("0.0", self)
        self.kv_label = QLabel("0.0", self)

        self.kw_percent = QDoubleSpinBox(self)
        self.kw_percent.setRange(-100, 100)
        self.kw_percent.setSuffix("%")
        self.kvar_percent = QDoubleSpinBox(self)
        self.kvar_percent.setRange(-100, 100)
        self.kvar_percent.setSuffix("%")
        self.kv_percent = QDoubleSpinBox(self)
        self.kv_percent.setRange(-100, 100)
        self.kv_percent.setSuffix("%")

        # Add widgets to layout
        kw_layout = QHBoxLayout()
        kw_layout.addWidget(QLabel("kw:"))
        kw_layout.addWidget(self.kw_label)
        kw_layout.addWidget(self.kw_percent)
        self.layout.addLayout(kw_layout)

        kvar_layout = QHBoxLayout()
        kvar_layout.addWidget(QLabel("kvar:"))
        kvar_layout.addWidget(self.kvar_label)
        kvar_layout.addWidget(self.kvar_percent)
        self.layout.addLayout(kvar_layout)

        kv_layout = QHBoxLayout()
        kv_layout.addWidget(QLabel("kv:"))
        kv_layout.addWidget(self.kv_label)
        kv_layout.addWidget(self.kv_percent)
        self.layout.addLayout(kv_layout)
        # ------------------------------------------------------------------------------------- #

        # ---------------------------------- Submit Changes --------------------------------------- #
        # Submit button to store edited values temporarily
        self.submit_changes_btn = QPushButton("Submit Changes", self)
        self.submit_changes_btn.clicked.connect(self.submit_changes)
        self.layout.addWidget(self.submit_changes_btn)

        # ---------------------------------- Global Adjustment --------------------------------------- #
        # Add global adjustment functionality
        self.global_adjustment_label = QLabel("Change All Loads:")
        self.layout.addWidget(self.global_adjustment_label)

        # Create checkboxes for selecting attributes
        self.attribute_layout = QHBoxLayout()  # Create a QHBoxLayout for checkboxes

        self.kw_checkbox = QCheckBox("kw", self)
        self.kw_checkbox.setChecked(True)
        self.attribute_layout.addWidget(self.kw_checkbox)

        self.kvar_checkbox = QCheckBox("kvar", self)
        self.kvar_checkbox.setChecked(True)
        self.attribute_layout.addWidget(self.kvar_checkbox)

        self.kv_checkbox = QCheckBox("kv", self)
        self.kv_checkbox.setChecked(True)
        self.attribute_layout.addWidget(self.kv_checkbox)

        # Add the QHBoxLayout to the main QVBoxLayout
        self.layout.addLayout(self.attribute_layout)  # Add the QHBoxLayout

        self.global_percentage_spinbox = QSpinBox(self)
        self.global_percentage_spinbox.setRange(-100, 100)  # Allow reductions too
        self.global_percentage_spinbox.setSuffix("%")
        self.layout.addWidget(self.global_percentage_spinbox)

        self.global_adjustment_btn = QPushButton("Submit all Loads changes", self)
        self.global_adjustment_btn.clicked.connect(self.apply_global_adjustment)
        self.layout.addWidget(self.global_adjustment_btn)

        # Add a horizontal line
        hline = QFrame(self)
        hline.setFrameShape(QFrame.HLine)
        hline.setFrameShadow(QFrame.Sunken)
        self.layout.addWidget(hline)

        # -- PV System Dialog -- #
        self.add_pv_btn = QPushButton("Add PV System", self)

        self.add_pv_btn.clicked.connect(self.show_pv_dialog)
        self.layout.addWidget(self.add_pv_btn)

        # ---------------------------------- Run Simulation --------------------------------------- #
        # Submit button to run simulation
        self.submit_btn = QPushButton("Run Simulation", self)
        self.submit_btn.setStyleSheet("background-color: lightgreen")
        self.submit_btn.clicked.connect(self.run_simulation)
        self.layout.addWidget(self.submit_btn)

        self.setLayout(self.layout)
        self.populate_values()  # populate initial values

    def populate_values(self):
        load = self.load_search.currentText()
        values = self.load_values[load]
        self.kw_label.setText("{:.2f}".format(values["kw"]))
        self.kvar_label.setText("{:.2f}".format(values["kvar"]))
        self.kv_label.setText("{:.2f}".format(values["kv"]))

        # Update the selected load display
        self.selected_load_display.setText(load)

    def on_load_selected(self):
        load = self.search_box.text()
        if load in self.load_values:
            values = self.load_values[load]
            self.kw_input.setText(str(values["kw"]))
            self.kvar_input.setText(str(values["kvar"]))
            self.kv_input.setText(str(values["kv"]))

    def on_load_selected_from_completer(self, selected_load):
        if selected_load in self.load_values:
            # Set the dropdown to the selected load
            index = self.load_search.findText(selected_load)
            if index != -1:
                self.load_search.setCurrentIndex(index)

            values = self.load_values[selected_load]
            self.kw_input.setText(str(values["kw"]))
            self.kvar_input.setText(str(values["kvar"]))
            self.kv_input.setText(str(values["kv"]))

    def submit_changes(self):
        load = self.load_search.currentText()

        # Calculate new values based on percentage
        kw_new_val = float(self.kw_label.text()) + (
            float(self.kw_label.text()) * (self.kw_percent.value() / 100)
        )
        kvar_new_val = float(self.kvar_label.text()) + (
            float(self.kvar_label.text()) * (self.kvar_percent.value() / 100)
        )
        kv_new_val = float(self.kv_label.text()) + (
            float(self.kv_label.text()) * (self.kv_percent.value() / 100)
        )

        self.temp_changes[load] = {
            "kw": kw_new_val,
            "kvar": kvar_new_val,
            "kv": kv_new_val,
        }

        self.populate_values()
        # Show a message to the user
        self.show_temporary_message("Changes submitted successfully!")

    def show_temporary_message(self, message, duration=1000):
        # Set the QLabel content and style
        self.message_label.setText(message)
        self.message_label.setStyleSheet(
            """
                color: white;
                background-color: green;
                padding: 5px;
                border-radius: 3px;
            """
        )

        # Use QTimer to hide the QLabel after the duration
        QTimer.singleShot(duration, self.message_label.clear)

    def apply_global_adjustment(self):
        adjustment_percentage = self.global_percentage_spinbox.value() / 100
        for load, values in self.load_values.items():
            if self.kw_checkbox.isChecked():
                values["kw"] *= 1 + adjustment_percentage
            if self.kvar_checkbox.isChecked():
                values["kvar"] *= 1 + adjustment_percentage
            if self.kv_checkbox.isChecked():
                values["kv"] *= 1 + adjustment_percentage
        self.show_temporary_message(
            f"Values adjusted by {adjustment_percentage*100}% globally!"
        )
        self.populate_values()  # Refresh the displayed values for the currently selected load

    def show_pv_dialog(self):
        load_selected = self.load_search.currentText()
        dialog = PvSystemDialog(self)
        if dialog.exec_():
            # get the values from the dialog
            name_of_pv = dialog.name_of_pv_input.text()
            phases = dialog.phases_input.value()
            kva = dialog.kva_input.text()
            kv = dialog.kv_input.text()
            pmpp = dialog.pmpp_input.text()
            irrad = dialog.irrad_input.text()
            print(load_selected, name_of_pv, phases, kva, kv, pmpp, irrad)

            # Show success message
            self.show_temporary_message(
                f"PV System {name_of_pv} added successfully!", duration=2000
            )
            # Now, you can proceed to add the PV system using the above values and the selected load
            # dssText.command = 'New PVSystem.' + ... (use your logic to construct the command)

    def run_simulation(self):
        # Apply changes to load_values
        for load, values in self.temp_changes.items():
            self.load_values[load] = values

        # Call your existing run_simulation method
        run_simulation(self.load_values)

        # Clear temp_changes after running the simulation
        self.temp_changes = {}


class PvSystemDialog(QDialog):
    def __init__(self, parent=None):
        super(PvSystemDialog, self).__init__(parent)
        self.setWindowTitle("Define PV System")

        # Create layout and widgets
        form_layout = QFormLayout()

        self.name_of_pv_input = QLineEdit(self)
        self.phases_input = QSpinBox(self)
        self.phases_input.setRange(1, 3)  # assuming either 1 or 3 phases
        self.kva_input = QLineEdit(self)
        self.kv_input = QLineEdit(self)
        self.pmpp_input = QLineEdit(self)
        self.irrad_input = QLineEdit(self)

        selec_load_display = QLabel(self)
        selec_load_display.setStyleSheet("color: red; font-weight: bold;")
        selec_load_display.setText(f"Selected Load: {parent.load_search.currentText()}")

        # Add widgets to form layout
        form_layout.addRow(selec_load_display)
        form_layout.addRow("Name of PV:", self.name_of_pv_input)
        form_layout.addRow("Phases:", self.phases_input)
        form_layout.addRow("kVA:", self.kva_input)
        form_layout.addRow("kV:", self.kv_input)
        form_layout.addRow("Pmpp:", self.pmpp_input)
        form_layout.addRow("Irradiance:", self.irrad_input)

        # Add save button
        save_btn = QPushButton("Save", self)
        save_btn.clicked.connect(self.accept)
        form_layout.addWidget(save_btn)

        self.setLayout(form_layout)


# class BusEditor(QWidget):
#     """
#     A widget for editing load values and running simulations.
#     """

#     def __init__(self, load_values, message_label):
#         super().__init__()

#         self.load_values = load_values
#         self.message_label = message_label
#         self.temp_changes = {}  # Temporarily store changes before simulation
#         self.layout = QVBoxLayout()
#         self.layout.setSpacing(5)  # Adjust the value for gap between widgets
#         self.layout.setContentsMargins(
#             10, 10, 10, 10
#         )  # Adjust these values to your preference

#         # Search Box with Autocomplete
#         self.search_box = QLineEdit(self)
#         self.completer = QCompleter(list(self.load_values.keys()), self)
#         self.layout.addWidget(QLabel("Search for a load..."))
#         self.search_box.setCompleter(self.completer)
#         self.search_box.setPlaceholderText("Search for a load...")
#         self.layout.addWidget(self.search_box)

#         # Connect search box to a slot that updates the editor when an item is selected
#         self.completer.activated.connect(self.on_load_selected_from_completer)

#         # Create dropdown field for load
#         self.load_search = QComboBox(self)
#         self.load_search.addItems(sorted(load_values.keys()))
#         self.load_search.currentIndexChanged.connect(self.populate_values)
#         self.layout.addWidget(QLabel("Select Bus:"))
#         self.layout.addWidget(self.load_search)

#         # After the dropdown creation code
#         self.selected_load_display = QLineEdit(self)
#         self.layout.addWidget(QLabel("Selected Bus:"))
#         self.selected_load_display.setPlaceholderText("Selected Bus")
#         self.selected_load_display.setReadOnly(True)  # Make it uneditable
#         self.layout.addWidget(self.selected_load_display)

#         # Add a horizontal line
#         hline = QFrame(self)
#         hline.setFrameShape(QFrame.HLine)
#         hline.setFrameShadow(QFrame.Sunken)
#         self.layout.addWidget(hline)

#         # Create fields for load attributes
#         self.kw_input = QLineEdit(self)
#         self.kvar_input = QLineEdit(self)
#         self.kv_input = QLineEdit(self)

#         # kw
#         # self.kw_percentage =
#         # self.kw_percentage.setRange(0, 100)  # Set range from 0 to 100 for percentage
#         # self.kw_percentage.setSuffix("%")  # Add a suffix to the spinbox

#         self.kw_layout = QHBoxLayout()
#         self.kw_value_label = QLabel(self)
#         self.kw_layout.addWidget(QLabel("kw:"))
#         self.kw_layout.addWidget(self.kw_value_label)
#         self.kw_layout.addWidget(QSpinBox(self))
#         self.layout.addLayout(self.kw_layout)

#         # kvar
#         # self.kvar_percentage = QSpinBox(self)
#         # self.kvar_percentage.setRange(0, 100)
#         # self.kvar_percentage.setSuffix("%")

#         self.kvar_layout = QHBoxLayout()
#         self.kvar_value_label = QLabel(self)
#         self.kvar_layout.addWidget(QLabel("kvar:"))
#         self.kvar_layout.addWidget(self.kvar_value_label)
#         self.kvar_layout.addWidget(QSpinBox(self))
#         self.layout.addLayout(self.kvar_layout)

#         # kv
#         # self.kv_percentage = QSpinBox(self)
#         # self.kv_percentage.setRange(0, 100)
#         # self.kv_percentage.setSuffix("%")

#         self.kv_layout = QHBoxLayout()
#         self.kv_value_label = QLabel(self)
#         self.kv_layout.addWidget(QLabel("kv:"))
#         self.kv_layout.addWidget(self.kv_value_label)
#         self.kv_layout.addWidget(QSpinBox(self))
#         self.layout.addLayout(self.kv_layout)

#         # Submit button to store edited values temporarily
#         self.submit_changes_btn = QPushButton("Submit Changes", self)
#         self.submit_changes_btn.clicked.connect(self.submit_changes)
#         self.layout.addWidget(self.submit_changes_btn)

#         # Add global adjustment functionality
#         self.global_adjustment_label = QLabel("Change All Loads:")
#         self.layout.addWidget(self.global_adjustment_label)

#         self.global_percentage_spinbox = QSpinBox(self)
#         self.global_percentage_spinbox.setRange(-100, 100)  # Allow reductions too
#         self.global_percentage_spinbox.setSuffix("%")
#         self.layout.addWidget(self.global_percentage_spinbox)

#         self.global_adjustment_btn = QPushButton("Apply Global Change", self)
#         self.global_adjustment_btn.clicked.connect(self.apply_global_adjustment)
#         self.layout.addWidget(self.global_adjustment_btn)

#         # Add a horizontal line
#         hline = QFrame(self)
#         hline.setFrameShape(QFrame.HLine)
#         hline.setFrameShadow(QFrame.Sunken)
#         self.layout.addWidget(hline)

#         # Submit button to run simulation
#         self.submit_btn = QPushButton("Run Simulation", self)
#         self.submit_btn.clicked.connect(self.run_simulation)
#         self.layout.addWidget(self.submit_btn)

#         self.setLayout(self.layout)
#         self.populate_values()  # populate initial values

#     def apply_global_adjustment(self):
#         adjustment_percentage = self.global_percentage_spinbox.value() / 100
#         for load, values in self.load_values.items():
#             values["kw"] *= 1 + adjustment_percentage
#             values["kvar"] *= 1 + adjustment_percentage
#             values["kv"] *= 1 + adjustment_percentage
#         self.show_temporary_message(
#             f"Values adjusted by {adjustment_percentage*100}% globally!"
#         )
#         self.populate_values()  # Refresh the displayed values for the currently selected load

#     def populate_values(self):
#         load = self.load_search.currentText()
#         values = self.load_values[load]

#         # Set the values for the labels next to the percentage spin boxes
#         self.kw_value_label.setText(str(values["kw"]))
#         self.kvar_value_label.setText(str(values["kvar"]))
#         self.kv_value_label.setText(str(values["kv"]))

#         # Update the selected load display
#         self.selected_load_display.setText(load)

#     def on_load_selected(self):
#         load = self.search_box.text()
#         if load in self.load_values:
#             values = self.load_values[load]
#             self.kw_input.setText(str(values["kw"]))
#             self.kvar_input.setText(str(values["kvar"]))
#             self.kv_input.setText(str(values["kv"]))

#     def on_load_selected_from_completer(self, selected_load):
#         if selected_load in self.load_values:
#             # Set the dropdown to the selected load
#             index = self.load_search.findText(selected_load)
#             if index != -1:
#                 self.load_search.setCurrentIndex(index)

#             values = self.load_values[selected_load]
#             self.kw_input.setText(str(values["kw"]))
#             self.kvar_input.setText(str(values["kvar"]))
#             self.kv_input.setText(str(values["kv"]))

#     def submit_changes(self):
#         # Store edited load values temporarily
#         load = self.load_search.currentText()
#         original_kw = self.load_values[load]["kw"]
#         original_kvar = self.load_values[load]["kvar"]
#         original_kv = self.load_values[load]["kv"]

#         kw_change = original_kw * (self.kw_percentage.value() / 100)
#         kvar_change = original_kvar * (self.kvar_percentage.value() / 100)
#         kv_change = original_kv * (self.kv_percentage.value() / 100)

#         self.temp_changes[load] = {
#             "kw": original_kw + kw_change,
#             "kvar": original_kvar + kvar_change,
#             "kv": original_kv + kv_change,
#         }

#         # Update input fields to show the modified values
#         self.kw_input.setText(str(original_kw + kw_change))
#         self.kvar_input.setText(str(original_kvar + kvar_change))
#         self.kv_input.setText(str(original_kv + kv_change))

#         # Show a message to the user
#         self.show_temporary_message("Changes submitted successfully!")

#     def show_temporary_message(self, message, duration=1000):
#         # Set the QLabel content and style
#         self.message_label.setText(message)
#         self.message_label.setStyleSheet(
#             """
#                 color: white;
#                 background-color: green;
#                 padding: 5px;
#                 border-radius: 3px;
#             """
#         )

#         # Use QTimer to hide the QLabel after the duration
#         QTimer.singleShot(duration, self.message_label.clear)

#     def run_simulation(self):
#         # Apply changes to load_values
#         for load, values in self.temp_changes.items():
#             self.load_values[load] = values

#         # Call your existing run_simulation method
#         run_simulation(self.load_values)

#         # Clear temp_changes after running the simulation
#         self.temp_changes = {}


if __name__ == "__main__":
    dssObj, dssText, dssCircuit, dssElem, dssSolution = setup_opendss()
    os.chdir(cwd_before)
    load_values, bus_coords, generator_values, line_values = load_bus_data(dssCircuit)
    create_map(load_values, bus_coords, generator_values, line_values)

    # PyQt5 app
    app = QApplication(sys.argv)

    # Main window
    main = QMainWindow()
    main.setWindowTitle("Bus Map")
    main.setGeometry(100, 100, 1000, 600)

    # Create and set a QStatusBar
    status_bar = main.statusBar()

    # Create a custom QLabel for messages
    message_label = QLabel()
    status_bar.addWidget(message_label)

    # Create QWebEngineView
    view = QWebEngineView(main)
    view.load(QUrl.fromLocalFile(os.path.abspath(MAP_HTML_FILE)))
    main.setCentralWidget(view)

    # Add bus editor as a docked panel
    bus_editor = BusEditor(load_values, message_label)
    dock = QDockWidget("Bus Editor", main)
    dock.setWidget(bus_editor)
    main.addDockWidget(Qt.RightDockWidgetArea, dock)

    # Show window
    main.show()
    sys.exit(app.exec_())

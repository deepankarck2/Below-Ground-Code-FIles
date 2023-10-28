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

VALUES_CSV_FILE = "values_csv.txt"  # It contains bus values. See github for sample file
SAMPLE_XY_FILE = "sample_for_x_y.txt"  # This file contains the coordinates for each bus
MAP_HTML_FILE = "map.html"  # No need to create it manually. It be created by Folium.


def load_bus_data():
    """
    Load bus data from CSV files.

    Returns:
        Tuple: A tuple containing two dictionaries. The first dictionary contains bus values, and the second dictionary contains bus coordinates.
    """
    bus_values = {}
    bus_coords = {}

    with open(f"{VALUES_CSV_FILE}", "r") as file:
        lines = file.readlines()
        for line in lines:
            sline = line.split(",")
            bus_values[sline[0]] = {
                "kw": float(sline[1]),
                "kvar": float(sline[2]),
                "kv": float(sline[3]),
            }

    with open(f"{SAMPLE_XY_FILE}", "r") as file:
        lines = file.readlines()
        for line in lines:
            sline = line.split(",")
            bus_coords[sline[0]] = {"lat": float(sline[1]), "lon": float(sline[2])}

    return bus_values, bus_coords


def create_map(bus_values, bus_coords):
    """
    Create a map with markers for each bus.

    Args:
        bus_values (dict): A dictionary containing bus values.
        bus_coords (dict): A dictionary containing bus coordinates.

    Returns:
        None
    """
    # Placeholder for all latitudes and longitudes to determine the map bounds later
    all_lats = []
    all_lons = []

    # Create a base map. You can adjust location and zoom_start as per your dataset.
    m = folium.Map(
        location=[
            list(bus_coords.values())[0]["lat"],
            list(bus_coords.values())[0]["lon"],
        ],
        zoom_start=10,
    )

    # Loop through each bus and add a marker on the map
    for bus, coord in bus_coords.items():
        if bus in bus_values:
            # Get bus attributes
            values = bus_values[bus]

            # Create a popup string
            popup_content = (
                # div with large font
                f"<div style='font-size: 12px'>"
                f"<strong style='color: blue'>Bus Name: {bus}</strong><br>"
                f"<ul>"
                f"<li>kw: {values['kw']}</li>"
                f"<li>kvar: {values['kvar']}</li>"
                f"<li>kv: {values['kv']}</li>"
                f"</ul>"
                f"</div>"
            )

            # Add marker to map
            marker = folium.Marker(
                [coord["lat"], coord["lon"]],
                popup=folium.Popup(popup_content, max_width=300),
                tooltip=bus,
            )
            m.add_child(marker)

            # Append the lat and lon to our lists
            all_lats.append(coord["lat"])
            all_lons.append(coord["lon"])

    # Determine the bounds to fit all markers
    sw = [min(all_lats), min(all_lons)]
    ne = [max(all_lats), max(all_lons)]

    # Adjust the map to fit these bounds
    m.fit_bounds([sw, ne])

    # Save map to an HTML file
    m.save(MAP_HTML_FILE)


def extract_numbers(s):
    """Extract all integers from a string and return them as a tuple using regular expression"""
    return tuple(map(int, re.findall(r"\d+", s)))


def custom_sort(item):
    """Custom sorting key that sorts by the extracted numbers."""
    return extract_numbers(item)


def run_simulation(bus_values):
    """
    Run a simulation with the given bus values.

    Args:
        bus_values (dict): A dictionary containing bus values.

    Returns:
        None
    """
    # For now, just print the values
    print(bus_values)


class BusEditor(QWidget):
    """
    A widget for editing bus values and running simulations.
    """

    def __init__(self, bus_values, message_label):
        super().__init__()

        self.bus_values = bus_values
        self.message_label = message_label
        self.temp_changes = {}  # Temporarily store changes before simulation
        self.layout = QVBoxLayout()

        # -------------------------------- Search -------------------------------------------- #
        # Create a new QVBoxLayout for the search box and its label
        search_layout = QVBoxLayout()
        search_layout.setSpacing(10)  # Adjust the value for gap between widgets

        # Create the label and set its size policy
        label = QLabel("Search for a bus...")
        label.setSizePolicy(
            QSizePolicy.Preferred, QSizePolicy.Fixed
        )  # Set vertical size policy to Fixed
        search_layout.addWidget(label)

        self.search_box = QLineEdit(self)
        self.completer = QCompleter(list(self.bus_values.keys()), self)
        self.search_box.setCompleter(self.completer)
        self.search_box.setPlaceholderText("Type to search...")
        self.search_box.setSizePolicy(
            QSizePolicy.Preferred, QSizePolicy.Fixed
        )  # Set vertical size policy to Fixed
        search_layout.addWidget(self.search_box)

        # Add the new layout to the main layout
        self.layout.addLayout(search_layout)

        # -------------------------------- Dropdown ------------------------------------------ #
        # Create dropdown field for bus
        self.bus_search = QComboBox(self)
        sorted_buses = sorted(bus_values.keys(), key=custom_sort)
        self.bus_search.addItems(sorted_buses)
        self.bus_search.currentIndexChanged.connect(self.populate_values)
        self.layout.addWidget(QLabel("Select Bus:"))
        self.layout.addWidget(self.bus_search)

        # ---------------------------------- Show Selected Bus ----------------------------------- #
        # After the dropdown creation code
        self.selected_bus_layout = QHBoxLayout()
        self.selected_bus_display = QLabel(self)
        self.selected_bus_display.setStyleSheet("color: red; font-weight: bold;")
        self.selected_bus_layout.addWidget(QLabel("Selected Bus:"))
        self.selected_bus_layout.addWidget(self.selected_bus_display)
        self.layout.addLayout(self.selected_bus_layout)

        # ------------------------------------------------------------------------------------- #

        # Add a horizontal line
        hline = QFrame(self)
        hline.setFrameShape(QFrame.HLine)
        hline.setFrameShadow(QFrame.Sunken)
        self.layout.addWidget(hline)

        # ---------------------------------- Change Bus Values --------------------------------------- #
        # Create fields for bus attributes
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
        bus = self.bus_search.currentText()
        values = self.bus_values[bus]
        self.kw_label.setText("{:.2f}".format(values["kw"]))
        self.kvar_label.setText("{:.2f}".format(values["kvar"]))
        self.kv_label.setText("{:.2f}".format(values["kv"]))

        # Update the selected bus display
        self.selected_bus_display.setText(bus)

    def on_bus_selected(self):
        bus = self.search_box.text()
        if bus in self.bus_values:
            values = self.bus_values[bus]
            self.kw_input.setText(str(values["kw"]))
            self.kvar_input.setText(str(values["kvar"]))
            self.kv_input.setText(str(values["kv"]))

    def on_bus_selected_from_completer(self, selected_bus):
        if selected_bus in self.bus_values:
            # Set the dropdown to the selected bus
            index = self.bus_search.findText(selected_bus)
            if index != -1:
                self.bus_search.setCurrentIndex(index)

            values = self.bus_values[selected_bus]
            self.kw_input.setText(str(values["kw"]))
            self.kvar_input.setText(str(values["kvar"]))
            self.kv_input.setText(str(values["kv"]))

    def submit_changes(self):
        bus = self.bus_search.currentText()

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

        self.temp_changes[bus] = {
            "kw": kw_new_val,
            "kvar": kvar_new_val,
            "kv": kv_new_val,
        }

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
        for bus, values in self.bus_values.items():
            if self.kw_checkbox.isChecked():
                values["kw"] *= 1 + adjustment_percentage
            if self.kvar_checkbox.isChecked():
                values["kvar"] *= 1 + adjustment_percentage
            if self.kv_checkbox.isChecked():
                values["kv"] *= 1 + adjustment_percentage
        self.show_temporary_message(
            f"Values adjusted by {adjustment_percentage*100}% globally!"
        )
        self.populate_values()  # Refresh the displayed values for the currently selected bus

    def show_pv_dialog(self):
        bus_selected = self.bus_search.currentText()
        dialog = PvSystemDialog(self)
        if dialog.exec_():
            # get the values from the dialog
            name_of_pv = dialog.name_of_pv_input.text()
            phases = dialog.phases_input.value()
            kva = dialog.kva_input.text()
            kv = dialog.kv_input.text()
            pmpp = dialog.pmpp_input.text()
            irrad = dialog.irrad_input.text()
            print(bus_selected, name_of_pv, phases, kva, kv, pmpp, irrad)

            # Show success message
            self.show_temporary_message(
                f"PV System {name_of_pv} added successfully!", duration=2000
            )
            # Now, you can proceed to add the PV system using the above values and the selected bus
            # dssText.command = 'New PVSystem.' + ... (use your logic to construct the command)

    def run_simulation(self):
        # Apply changes to bus_values
        for bus, values in self.temp_changes.items():
            self.bus_values[bus] = values

        # Call your existing run_simulation method
        run_simulation(self.bus_values)

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

        selec_bus_display = QLabel(self)
        selec_bus_display.setStyleSheet("color: red; font-weight: bold;")
        selec_bus_display.setText(f"Selected Load: {parent.bus_search.currentText()}")

        # Add widgets to form layout
        form_layout.addRow(selec_bus_display)
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

if __name__ == "__main__":
    bus_values, bus_coords = load_bus_data()
    create_map(bus_values, bus_coords)

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
    bus_editor = BusEditor(bus_values, message_label)
    dock = QDockWidget("Bus Editor", main)
    dock.setWidget(bus_editor)
    main.addDockWidget(Qt.RightDockWidgetArea, dock)

    # Show window
    main.show()
    sys.exit(app.exec_())

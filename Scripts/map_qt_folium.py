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
import math
import random
import pandas as pd
import numpy as np

# Absolute paths for the CSV files
# VALUES_CSV_FILE = "A:\\CCNY\\J_Fall_2023\\SD2\\OpenDSS\\load_values.csv"
# SAMPLE_XY_FILE = "A:\\CCNY\\J_Fall_2023\\SD2\\OpenDSS\\sample_for_x_y.txt"
# GENERATOR_CSV_FILE = "A:\\CCNY\\J_Fall_2023\\SD2\\OpenDSS\\generator_values.csv"
# LINES_CSV_FILE = "lines_values.csv"
# FILE_PATH = "'A:\CCNY\J_Fall_2023\SD2\OpenDSS\IEEE 30 Bus\Master.dss'"
# MODEL_NAME = "random_forest_model30.joblib"

LINE_LOAD_VALUE = r"A:\CCNY\J_Fall_2023\SD2\OpenDSS\iee9500linesggdata.csv"  # Check GitHub Sample_Data folder.
MAP_HTML_FILE = "A:\\CCNY\\J_Fall_2023\\SD2\\OpenDSS\\map.html"

if not os.path.exists(MAP_HTML_FILE):
    open(MAP_HTML_FILE, "w").close()


FILE_PATH = (
    "'A:\CCNY\J_Fall_2023\SD2\OpenDSS\ieee9500dss\ieee9500dss\ieee9500_base _copy.dss'"
)

# Store the current working directory before calling the function
cwd_before = os.getcwd()

MAX_ITER = 1000
MAX_CONTROL_ITER = 100


def setup_opendss():
    """
    Set up the OpenDSS engine and circuit.

    Returns:
        Tuple: A tuple containing the OpenDSS objects.
    """
    dssObj = win32com.client.Dispatch("OpenDSSEngine.DSS")

    # Start the DSS
    if not dssObj.Start(0):
        print("DSS failed to start!")
        exit()

    dssText = dssObj.Text
    dssCircuit = dssObj.ActiveCircuit
    dssElement = dssCircuit.ActiveCktElement
    dssSolution = dssCircuit.Solution
    dssText.Command = f"compile {FILE_PATH}"  # Load the circuit
    dssText.Command = f"set maxiterations={MAX_ITER} maxControlIter={MAX_CONTROL_ITER}"

    return dssObj, dssText, dssCircuit, dssElement, dssSolution


def calculate_line_loading(dssCircuit, dssText, dssSolution):
    """
    Calculate the line loading values for the circuit.

    Args:
        dssCircuit (object): The OpenDSS circuit object.
        dssText (object): The OpenDSS text object.
        dssSolution (object): The OpenDSS solution object.

    Returns:
        dict: A dictionary containing the line loading values for each line.
    """
    dssLines = dssCircuit.Lines
    dssMonitors = dssCircuit.Monitors
    dssText.Command = "set mode = daily"
    dssText.Command = "set Number = 1"
    line_values = {}

    # Activate the first Line to start the iteration
    iLine = dssLines.First
    while iLine > 0:
        # Get the current Line's name
        line_name = dssLines.Name

        # Using the Lines COM interface to get Bus1 and Bus2 names
        bus1 = dssLines.Bus1
        bus2 = dssLines.Bus2

        # Initialize loading as None
        line_values[line_name] = {
            "Bus1": bus1,
            "Bus2": bus2,
            "Loading": random.uniform(0, 100),
        }

        # Move to the next Line object
        iLine = dssLines.Next

    # Line Loading Calculation
    monitor_idx = dssMonitors.First
    while monitor_idx > 0:
        line_name = dssMonitors.Name.split(".")[
            1
        ]  # Assuming monitor name format is 'Monitor.line_name'
        if line_name in line_values:
            dssMonitors.Name = f"line_{line_name}"
            I1 = np.array(dssMonitors.Channel(7), dtype=float)  # Current Magnitude
            I2 = np.array(dssMonitors.Channel(9), dtype=float)
            I3 = np.array(dssMonitors.Channel(11), dtype=float)

            # Calculating the total current
            total_current = np.sqrt(I1**2 + I2**2 + I3**2)

            # Get line's normal current rating
            dssLines.Name = line_name
            NormAmps = dssLines.NormAmps

            # Calculate the line loading
            if NormAmps != 0:
                line_loading = np.mean(total_current) * 100 / NormAmps
                line_values[line_name]["Loading"] = random.uniform(0, 100)

        monitor_idx = dssMonitors.Next

    return line_values


def load_bus_data(dssCircuit, dssElement, dssText, dssSolution):
    """
    Load bus data from CSV files.

    Returns:
        Tuple: A tuple containing two dictionaries. The first dictionary contains bus values, and the second dictionary contains bus coordinates.
    """
    load_values = {}
    bus_coords = {}
    generator_values = {}
    line_values = {}

    # ------------------------ LINES ------------------------#
    dssLines = dssCircuit.Lines

    # Activate the first Line to start the iteration
    iLine = dssLines.First
    while iLine > 0:
        # Get the current Line's name
        line_name = dssLines.Name

        # Using the Lines COM interface to get Bus1 and Bus2 names
        bus1 = dssLines.Bus1
        bus2 = dssLines.Bus2

        # Store the data in the dictionary
        line_values[line_name] = {"Bus1": bus1, "Bus2": bus2}

        # Move to the next Line object
        iLine = dssLines.Next

    # ------------------------ Line Loads ------------------------#
    os.chdir(cwd_before)
    # Read the Excel file into a DataFrame
    try:
        df = pd.read_csv(LINE_LOAD_VALUE)
        # Convert the DataFrame to a dictionary with line_name as keys and line_value as values
        line_loading_values = pd.Series(
            df.line_value.values, index=df.line_name
        ).to_dict()

        # Merge line loading values into line_values
        for line_name, loading in line_loading_values.items():
            if line_name in line_values:
                line_values[line_name]["Loading"] = loading
            else:
                print(
                    f"Warning: Line '{line_name}' found in line loading data but not in line values."
                )
    except FileNotFoundError:
        print(
            f"LINE_LOAD_VALUE CSV File not found. Check the file path and name. See beginning of this file for the path."
        )
        for line_name, values in line_values.items():
            if line_name in line_values:
                line_values[line_name]["Loading"] = "nan"
        # Handle the error or exit

    # ------------------------ LOADS ------------------------#

    i = dssCircuit.Loads.First
    while i:
        load_name = dssCircuit.Loads.Name
        # Set dssElement to the active load
        dssCircuit.SetActiveElement(load_name)
        # Fetch the bus to which the load is connected using dssElement
        bus1 = dssCircuit.ActiveElement.Properties("Bus1").Val
        kv = dssCircuit.Loads.kV  # This fetches the base kV for the load
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

    # ------------------------ BUS COORDINATES ------------------------#
    coordinates_missing = False
    for i in range(dssCircuit.NumBuses):
        bus = dssCircuit.Buses(i)
        if bus.x != 0 or bus.y != 0:
            bus_name = bus.Name.lower()
            bus_coords[bus_name] = {"lat": bus.y, "lon": bus.x}
        else:
            coordinates_missing = True
            # print(f"No coordinates found for bus: {bus.Name}")
            break

    if coordinates_missing:
        # Coordinates are missing, read from the CSV file
        try:
            with open(f"{SAMPLE_XY_FILE}", "r") as file:
                lines = file.readlines()
                for line in lines:
                    sline = line.split(",")
                    bus_name = sline[0].lower()
                    bus_coords[bus_name] = {
                        "lat": float(sline[1]),
                        "lon": float(sline[2]),
                    }
        except FileNotFoundError:
            print(
                f"The file {SAMPLE_XY_FILE} was not found. Please check the file path and name."
            )

    # ------------------------ GENERATORS ------------------------#
    dssGenerators = dssCircuit.Generators

    # Activate the first generator to start the iteration
    iGen = dssGenerators.First
    while iGen > 0:
        # Get the current generator's name
        genName = dssGenerators.Name

        bus_val = dssElement.Properties("Bus1").val.lower()

        # Retrieve the properties of the current generator
        generator_values[genName] = {
            "Bus1": bus_val,
            "kV": dssGenerators.kV,
            "kW": dssGenerators.kW,
            "kvar": dssGenerators.kvar,
            "Model": dssGenerators.Model,
            "Vpu": dssElement.Properties("Vpu").val,
            "Maxkvar": dssElement.Properties("Maxkvar").val,
            "Minkvar": dssElement.Properties("Minkvar").val,
        }

        # Move to the next generator
        iGen = dssGenerators.Next

    # Assuming dssCircuit, dssElement are already defined and imported from OpenDSS

    dssTransformers = dssCircuit.Transformers

    # Initialize a dictionary to store transformer values
    transformer_values = {}

    # Activate the first transformer to start the iteration
    iTrans = dssTransformers.First
    while iTrans > 0:
        # Get the current transformer's name
        transName = dssTransformers.Name

        # Retrieve the properties of the current transformer
        transformer_values[transName] = {
            "Wdg": dssTransformers.NumWindings,
            "kVA": dssTransformers.kVA,
            "Tap": dssTransformers.Tap,
            "Xhl": dssTransformers.Xhl,
            "Xht": dssTransformers.Xht,
            "Xlt": dssTransformers.Xlt,
            "Buses": dssElement.Properties("Buses").val,
            "Conn": dssElement.Properties("Conn").val,
            # Add other relevant properties here
        }

        # Move to the next transformer
        iTrans = dssTransformers.Next

    return load_values, bus_coords, generator_values, line_values, transformer_values


def add_pu_feedback_layer(m, bus_coords, given_voltages, threshold=(0.95, 1.05)):
    """
    Add a layer on the map to represent the PU values of each bus.

    Args:
        m (folium.Map): The map object to which the layer is to be added.
        bus_coords (dict): A dictionary containing bus coordinates.
        voltages (dict): A dictionary containing the PU values of the buses.
        threshold (tuple): A tuple with the lower and upper bounds for a good PU value.

    Returns:
        folium.Map: The map object with the PU feedback layer added.
    """
    # Create a feature group for PU feedback
    pu_feedback_layer = folium.FeatureGroup(name="PU Feedback")
    interval = 0
    # Iterate over bus_coords to create markers
    for bus_name, coord in bus_coords.items():
        # Get the PU value for the bus
        pu_value = given_voltages[interval + 3 % len(given_voltages)]
        if pu_value:
            # Determine the color based on the PU value
            color = "green" if threshold[0] <= pu_value <= threshold[1] else "red"

            # color = "green" if random.randint(0, 1) == 0 else "red"
            # Create a marker with the appropriate color
            hollow_circle = folium.Circle(
                location=[coord["lat"], coord["lon"]],
                radius=6,  # Define the radius of the circle
                color=color,
                weight=3,  # Define how thick the circle's border should be
                fill=False,  # Set fill to False to create a hollow circle
                popup=f"Bus: {bus_name}<br>PU: {pu_value:.3f}",
            )
            pu_feedback_layer.add_child(hollow_circle)
            interval += 1
    # Add the layer to the map
    m.add_child(pu_feedback_layer)

    # Add a LayerControl to toggle this layer
    m.add_child(folium.LayerControl())

    return m


# Helper function to strip phase notations from bus names
def get_base_bus_name(bus_name_with_phases):
    """
    Get the base bus name without phase notations.

    Args:
        bus_name_with_phases (str): The bus name with phase notations.

    Returns:
        str: The base bus name without phase notations.
    """
    return (
        (bus_name_with_phases.split(".")[0])
        if "." in bus_name_with_phases
        else bus_name_with_phases
    )


def add_custom_legend(m):
    from branca.element import Template, MacroElement

    # Define the HTML template for the legend
    template = """
    {% macro html(this, kwargs) %}
    <div style="position: fixed; 
                bottom: 20px; left: 30px; width: 150px; height: 90px; 
                border:2px solid grey; z-index:9999; font-size:14px;
                ">&nbsp; Legend <br>
      &nbsp; <i class="fa fa-circle" style="color:blue"></i> Load &nbsp; <br>
      &nbsp; <i class="fa fa-circle" style="color:orange"></i> Generator &nbsp; 
    </div>
    {% endmacro %}
    """

    # Create a MacroElement object to hold the HTML
    macro = MacroElement()
    macro._template = Template(template)

    # Add the legend to the map
    m.get_root().add_child(macro)


def create_map(
    load_values, bus_coords, generator_values, line_values, voltages, transformer_values
):
    all_lats = []
    all_lons = []

    # Assuming MAP_HTML_FILE is defined elsewhere in your script
    MAP_HTML_FILE = "map.html"

    # Initialize map at the first bus coordinates
    first_bus_coords = next(iter(bus_coords.values()))
    m = folium.Map(
        location=[first_bus_coords["lat"], first_bus_coords["lon"]],
        zoom_start=10,
    )

    # For Transmission Lines
    for line, values in line_values.items():
        bus1_base_name = get_base_bus_name(values["Bus1"])
        bus2_base_name = get_base_bus_name(values["Bus2"])

        popup_content = (
            f"<div style='font-size: 14px'>"
            f"<div>Line --> To       -       From </div>"
            f"<div> 1) {values['Bus1']} - {values['Bus2']}<br></div>"
            f"<div> 2) Line Loading: <strong>{values['Loading']}%</strong></div>"
            f"</div>"
        )

        if bus1_base_name in bus_coords and bus2_base_name in bus_coords:
            bus1_coord = bus_coords[bus1_base_name]
            bus2_coord = bus_coords[bus2_base_name]
            line_coords = [
                (bus1_coord["lat"], bus1_coord["lon"]),
                (bus2_coord["lat"], bus2_coord["lon"]),
            ]
            folium.PolyLine(
                line_coords,
                popup=folium.Popup(popup_content, max_width=300),
                # is values['Loading'] is below 80 then green, between 81-99 then yellow, above 100 then red, if its a string then grey
                # is values['Loading'] is below 80 then green, between 81-99 then yellow, above 100 then red, if its a string then grey
                color="grey"
                if values["Loading"] == "nan"
                else "green"
                if values["Loading"] < 80
                else "yellow"
                if values["Loading"] < 100
                else "red",
            ).add_to(m)
        else:
            print(
                f"Coordinates for buses {bus1_base_name} or {bus2_base_name} not found."
            )

    # For Buses (Loads)
    for load_name, values in load_values.items():
        # Convert bus to lowercase for matching with bus_coords keys
        bus_name = get_base_bus_name(values["bus"].lower())

        coord = bus_coords[bus_name]
        popup_content = (
            f"<div style='font-size: 14px'>"
            f"<strong style='color: blue'>Bus Name: {load_name}</strong><br>"
            f"<ul>"
            f"<li>kw: {values['kw']}</li>"
            f"<li>kvar: {values['kvar']}</li>"
            f"<li>kv: {values['kv']}</li>"
            f"</ul>"
            f"</div>"
        )

        # marker = folium.Marker(
        #     [coord["lat"], coord["lon"]],
        #     icon=folium.Icon(prefix="fa", icon="lightbulb", color="blue"),
        #     popup=folium.Popup(popup_content, max_width=300),
        #     tooltip=load_name,
        # )

        marker = folium.CircleMarker(
            location=[coord["lat"], coord["lon"]],
            radius=5,  # Small, non-obtrusive point
            color="blue",
            fill=True,
            fill_color="blue",
            weight=1,
            fill_opacity=0.6,
            tooltip=load_name,
            popup=folium.Popup(popup_content, max_width=300),
        )
        m.add_child(marker)

        all_lats.append(coord["lat"])
        all_lons.append(coord["lon"])

    # For Generators
    for gen, values in generator_values.items():
        # Use the helper function to get the base name without phases
        bus_base_name = get_base_bus_name(values["Bus1"])

        # Check if the base name exists in the coordinates dictionary
        coord = bus_coords[bus_base_name]
        popup_content = (
            f"<div style='font-size: 14px'>"
            f"<strong style='color: red'>Generator Name: {gen}</strong><br>"
            f"<ul>"
            f"<li>kW: {values['kW']}</li>"
            f"<li>kvar: {values['kvar']}</li>"
            f"<li>kV: {values['kV']}</li>"
            f"</ul>"
            f"</div>"
        )

        # marker = folium.Marker(
        #     [coord["lat"], coord["lon"]],
        #     popup=folium.Popup(popup_content, max_width=300),
        #     tooltip=gen,
        #     icon=folium.Icon(
        #         prefix="fa", icon="bolt-lightning", color="orange"
        #     ),  # Different icon for generators
        # )

        marker = folium.CircleMarker(
            location=[coord["lat"], coord["lon"]],
            radius=5,  # Small, non-obtrusive point
            popup=folium.Popup(popup_content, max_width=300),
            color="orange",
            fill=True,
            fill_color="orange",
            fill_opacity=0.7,
            tooltip=gen,
        )

        m.add_child(marker)

    # For Transformers
    for trans, values in transformer_values.items():
        # Use the helper function to get the base name without phases
        # Assuming transformers also have a 'Buses' property that lists connected buses

        # Clean the 'Buses' string to remove unwanted characters
        buses_str = values["Buses"].strip("[] ")
        first_bus = buses_str.split(",")[0]  # Taking the first bus for the location
        bus_base_name = get_base_bus_name(first_bus)

        # Check if the base name exists in the coordinates dictionary
        if bus_base_name in bus_coords:
            coord = bus_coords[bus_base_name]
            popup_content = (
                f"<div style='font-size: 14px'>"
                f"<strong style='color: blue'>Transformer Name: {trans}</strong><br>"
                f"<ul>"
                f"<li>kVA: {values['kVA']}</li>"
                f"<li>Tap: {values['Tap']}</li>"
                f"<li>Winding: {values['Wdg']}</li>"
                f"</ul>"
                f"</div>"
            )

            # Using a different marker for transformers
            marker = folium.CircleMarker(
                location=[coord["lat"], coord["lon"]],
                radius=5,  # Small, non-obtrusive point
                popup=folium.Popup(popup_content, max_width=300),
                color="purple",  # Different color for transformers
                fill=True,
                fill_color="purple",
                fill_opacity=0.7,
                tooltip=trans,
            )

        # m.add_child(marker)

    # Adjusting map bounds
    sw = [min(all_lats), min(all_lons)]
    ne = [max(all_lats), max(all_lons)]
    m.fit_bounds([sw, ne])

    # Add custom legend
    add_custom_legend(m)

    # Save the map as an HTML file
    m.save(MAP_HTML_FILE)

    return m


def extract_numbers(s):
    """Extract all integers from a string and return them as a tuple using regular expression"""
    return tuple(map(int, re.findall(r"\d+", s)))


def custom_sort(item):
    """Custom sorting key that sorts by the extracted numbers."""
    return extract_numbers(item)


def run_simulation(load_values, generator_values, changed_loads, map_obj):
    """
    Run a simulation with the given bus values.

    Args:
        load_values (dict): A dictionary containing bus values.

    Returns:
        None
    """
    # Apply changes to the loads in OpenDSS
    for load_name, values in changed_loads.items():
        # Assuming load_name corresponds to the name of the load in OpenDSS
        dssCircuit.Loads.Name = load_name

        if "kw" in values:
            dssCircuit.Loads.kW = values["kw"]
        if "kvar" in values:
            dssCircuit.Loads.kvar = values["kvar"]

    # Now, solve the circuit with the new values
    dssSolution.Solve()

    # Check if the solution converged and print the result
    if dssSolution.Converged:
        print(f"Solution converged successfully.")
    else:
        print(f"Solution did not converge.")

    new_voltages = dssCircuit.AllBusVmagPu

    # ---------------------------------- Display Back --------------------------------------- #
    # # # Then, create your map and add the PU feedback layer
    map_obj = add_pu_feedback_layer(map_obj, bus_coords, new_voltages)

    # # Save the map as an HTML file (this is assumed to be a global constant)
    map_obj.save(MAP_HTML_FILE)

    # # Trigger the refresh of the map view in your GUI
    refresh_map_view()

    # For now, just print the values
    # print(load_values)


def refresh_map_view():
    """
    Reload the QWebEngineView that contains the map.
    """
    global view  # Assuming 'view' is your QWebEngineView instance
    view.load(QUrl.fromLocalFile(os.path.abspath(MAP_HTML_FILE)))


class BusEditor(QWidget):
    """
    A widget for editing bus values and running simulations.
    """

    def __init__(self, load_values, generator_values, message_label, map_obj):
        super().__init__()

        self.load_values = load_values
        self.generator_values = generator_values
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

        self.kw_percent = QDoubleSpinBox(self)
        self.kw_percent.setRange(-100, 100)
        self.kw_percent.setSuffix("%")
        self.kvar_percent = QDoubleSpinBox(self)
        self.kvar_percent.setRange(-100, 100)
        self.kvar_percent.setSuffix("%")

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

        # Update the selected load display
        self.selected_load_display.setText(load)

    def on_load_selected(self):
        load = self.search_box.text()
        if load in self.load_values:
            values = self.load_values[load]
            self.kw_input.setText(str(values["kw"]))
            self.kvar_input.setText(str(values["kvar"]))

    def on_load_selected_from_completer(self, selected_load):
        if selected_load in self.load_values:
            # Set the dropdown to the selected load
            index = self.load_search.findText(selected_load)
            if index != -1:
                self.load_search.setCurrentIndex(index)

            values = self.load_values[selected_load]
            self.kw_input.setText(str(values["kw"]))
            self.kvar_input.setText(str(values["kvar"]))

    def submit_changes(self):
        load = self.load_search.currentText()

        # Calculate new values based on percentage
        kw_new_val = float(self.kw_label.text()) + (
            float(self.kw_label.text()) * (self.kw_percent.value() / 100)
        )
        kvar_new_val = float(self.kvar_label.text()) + (
            float(self.kvar_label.text()) * (self.kvar_percent.value() / 100)
        )

        self.temp_changes[load] = {
            "kw": kw_new_val,
            "kvar": kvar_new_val,
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
        # Initialize a temporary dictionary to store changes

        for load, values in self.load_values.items():
            # Create a new dictionary to store the adjusted values
            adjusted_values = values.copy()

            if self.kw_checkbox.isChecked():
                adjusted_values["kw"] *= 1 + adjustment_percentage
            if self.kvar_checkbox.isChecked():
                adjusted_values["kvar"] *= 1 + adjustment_percentage

            # Store the adjusted values in the temp_changes dictionary
            self.temp_changes[load] = adjusted_values

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
        run_simulation(
            self.load_values, self.generator_values, self.temp_changes, map_obj
        )

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


if __name__ == "__main__":
    dssObj, dssText, dssCircuit, dssElement, dssSolution = setup_opendss()
    os.chdir(cwd_before)
    voltages = dssCircuit.AllBusVmagPu
    (
        load_values,
        bus_coords,
        generator_values,
        line_values,
        transformer_values,
    ) = load_bus_data(dssCircuit, dssElement, dssText, dssSolution)
    map_obj = create_map(
        load_values,
        bus_coords,
        generator_values,
        line_values,
        voltages,
        transformer_values,
    )

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
    bus_editor = BusEditor(load_values, generator_values, message_label, map_obj)
    dock = QDockWidget("Bus Editor", main)
    dock.setWidget(bus_editor)
    main.addDockWidget(Qt.RightDockWidgetArea, dock)

    # Show window
    main.show()
    sys.exit(app.exec_())

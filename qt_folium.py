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
import sys
from PyQt5.QtWidgets import (
    QApplication,
    QMainWindow,
    QDockWidget,
    QVBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QComboBox,
    QWidget,
    QCompleter,
    QFrame,
)
import folium
import os
from PyQt5.QtWebEngineWidgets import QWebEngineView
from PyQt5.QtCore import QUrl, Qt, QTimer

selected_bus = None


def load_bus_data():
    """
    Load bus data from CSV files.

    Returns:
        Tuple: A tuple containing two dictionaries. The first dictionary contains bus values, and the second dictionary contains bus coordinates.
    """
    bus_values = {}
    bus_coords = {}

    with open("values_csv.txt", "r") as file:
        lines = file.readlines()
        for line in lines:
            sline = line.split(",")
            bus_values[sline[0]] = {
                "kw": float(sline[1]),
                "kvar": float(sline[2]),
                "kv": float(sline[3]),
            }

    with open("sample_for_x_y.txt", "r") as file:
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
                f"Bus name: {bus}<br>"
                f"kw={values['kw']}<br>"
                f"kvar={values['kvar']}<br>"
                f"kv={values['kv']}"
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
    m.save("map.html")


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
        self.layout.setSpacing(5)  # Adjust the value for gap between widgets
        self.layout.setContentsMargins(
            10, 10, 10, 10
        )  # Adjust these values to your preference

        # Search Box with Autocomplete
        self.search_box = QLineEdit(self)
        self.completer = QCompleter(list(self.bus_values.keys()), self)
        self.layout.addWidget(QLabel("Search for a bus..."))
        self.search_box.setCompleter(self.completer)
        self.search_box.setPlaceholderText("Search for a bus...")
        self.layout.addWidget(self.search_box)

        # Connect search box to a slot that updates the editor when an item is selected
        self.completer.activated.connect(self.on_bus_selected_from_completer)

        # Create dropdown field for bus
        self.bus_search = QComboBox(self)
        self.bus_search.addItems(sorted(bus_values.keys()))
        self.bus_search.currentIndexChanged.connect(self.populate_values)
        self.layout.addWidget(QLabel("Select Bus:"))
        self.layout.addWidget(self.bus_search)

        # After the dropdown creation code
        self.selected_bus_display = QLineEdit(self)
        self.layout.addWidget(QLabel("Selected Bus:"))
        self.selected_bus_display.setPlaceholderText("Selected Bus")
        self.selected_bus_display.setReadOnly(True)  # Make it uneditable
        self.layout.addWidget(self.selected_bus_display)

        # Add a horizontal line
        hline = QFrame(self)
        hline.setFrameShape(QFrame.HLine)
        hline.setFrameShadow(QFrame.Sunken)
        self.layout.addWidget(hline)

        # Create fields for bus attributes
        self.kw_input = QLineEdit(self)
        self.kvar_input = QLineEdit(self)
        self.kv_input = QLineEdit(self)
        self.layout.addWidget(QLabel("kw:"))
        self.layout.addWidget(self.kw_input)
        self.layout.addWidget(QLabel("kvar:"))
        self.layout.addWidget(self.kvar_input)
        self.layout.addWidget(QLabel("kv:"))
        self.layout.addWidget(self.kv_input)

        # Submit button to store edited values temporarily
        self.submit_changes_btn = QPushButton("Submit Changes", self)
        self.submit_changes_btn.clicked.connect(self.submit_changes)
        self.layout.addWidget(self.submit_changes_btn)

        # Submit button to run simulation
        self.submit_btn = QPushButton("Run Simulation", self)
        self.submit_btn.clicked.connect(self.run_simulation)
        self.layout.addWidget(self.submit_btn)

        self.setLayout(self.layout)
        self.populate_values()  # populate initial values

    def populate_values(self):
        bus = self.bus_search.currentText()
        values = self.bus_values[bus]
        self.kw_input.setText(str(values["kw"]))
        self.kvar_input.setText(str(values["kvar"]))
        self.kv_input.setText(str(values["kv"]))

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
        # Store edited bus values temporarily
        bus = self.bus_search.currentText()
        self.temp_changes[bus] = {
            "kw": float(self.kw_input.text()),
            "kvar": float(self.kvar_input.text()),
            "kv": float(self.kv_input.text()),
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

    def run_simulation(self):
        # Apply changes to bus_values
        for bus, values in self.temp_changes.items():
            self.bus_values[bus] = values

        # Call your existing run_simulation method
        run_simulation(self.bus_values)

        # Clear temp_changes after running the simulation
        self.temp_changes = {}


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
    view.load(QUrl.fromLocalFile(os.path.abspath("map.html")))
    main.setCentralWidget(view)

    # Add bus editor as a docked panel
    bus_editor = BusEditor(bus_values, message_label)
    dock = QDockWidget("Bus Editor", main)
    dock.setWidget(bus_editor)
    main.addDockWidget(Qt.RightDockWidgetArea, dock)

    # Show window
    main.show()
    sys.exit(app.exec_())

# Below-Ground-Code-FIles

## File 1: `opendss_python_simulation.py`
This module is used to run OpenDSS simulations and modify the simulation by changing loads or applying outages.
It also compares the initial and modified voltages and exports the voltage data to a CSV file.

#### Instructions:
1) Install OpenDSS. Skip this step if OpenDSS is already installed. 
2) **Install necessary packages:** Copy the following command and paste it into the command prompt.
    ```pip install plotly matplotlib pandas pywin32```
    
    Skip if you don't want to plot the data. 
    And comment out the import statements for the packages, and comment out the graph_compare_voltages() function in main.
3) Run the module as follows:
    python opendss_python_simulation.py
Then, follow the instructions on the screen to modify the simulation.
"""

## File: ```qt_folium.py```
This module contains code for a PyQt5 application that displays a map with markers for each bus. 
The user can edit bus values and run simulations. It also contains code for creating the map using Folium. 
The map is saved as an HTML file, which is then loaded in the PyQt5 application.
There are two CSV files containing bus values and coordinates(values.txt, and sample_for_x_y.txt). Attached are sample CSV files.

#### The application has the following features:
- A search box with autocomplete for searching buses.
- A dropdown for selecting buses.
- A form for editing bus values.
- A button for submitting changes.
- A button for running simulations.
- A status bar for showing messages to the user.

#### Instructions To Run:
1) **Install the required packages:** Copy the following command and paste it into the command prompt.
    ``` pip install PyQt5 folium ```
2) Download the CSV files and save them in the same directory as this file.
3) Run the code using python qt_folium.py.

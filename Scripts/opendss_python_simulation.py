""" 
This module is used to run OpenDSS simulations and modify the simulation by changing loads or applying outages.
It also compares the initial and modified voltages and exports the voltage data to a CSV file.

Instructions:
1) Install OpenDSS. Skip this step if OpenDSS is already installed. 
2) Install necessary packages: Copy the following command and paste it in the command prompt.
    ```pip install plotly matplotlib pandas pywin32```
    
    Skip if you don't want to plot the data. 
    And comment out the import statements for the packages, and the graph_compare_voltages() function in main.

Run the module as follows:
    python opendss_python_simulation.py
Then, follow the instructions on the screen to modify the simulation.
"""

import win32com.client
import matplotlib.pyplot as plt
import pandas as pd
import plotly.graph_objects as go

# Please find the path of the Master.dss file or other opendss files in your computer and replace it for FILE_PATH.
FILE_PATH = "'A:\Softwares\OpenDSS\IEEETestCases\IEEE 30 Bus\Master.dss'"
# FILE_PATH = (
#     "'C:\\Users\\dip2l\\Downloads\\ieee9500dss\\ieee9500dss\\ieee9500_base _copy.dss'"
# )


def setup_opendss():
    dssObj = win32com.client.Dispatch("OpenDSSEngine.DSS")

    # Start the DSS
    if not dssObj.Start(0):
        print("DSS failed to start!")
        exit()

    dssText = dssObj.Text
    dssCircuit = dssObj.ActiveCircuit

    return dssObj, dssText, dssCircuit


def run_simulation(dssText, dssCircuit):
    dssText.Command = "export voltages IEEE_30_VLN_Node_Initial.Txt"
    dssText.Command = f"compile {FILE_PATH}"
    dssText.Command = "solve"

    voltages = dssCircuit.AllBusVmagPu
    return voltages


def list_load_names(dssCircuit):
    load_names = []
    i = dssCircuit.Loads.First
    while i:
        load_names.append(dssCircuit.Loads.Name)
        i = dssCircuit.Loads.Next
    return load_names


def list_line_names(dssCircuit):
    """Returns a list of line names from the OpenDSS circuit."""
    return [line_name for line_name in dssCircuit.Lines.AllNames]


def list_transformer_names(dssCircuit):
    """Returns a list of transformer names from the OpenDSS circuit."""
    return [transformer_name for transformer_name in dssCircuit.Transformers.AllNames]


def get_load_details(dssCircuit):
    load_names = []
    load_kw = []
    load_kvar = []

    i = dssCircuit.Loads.First
    while i:
        load_names.append(dssCircuit.Loads.Name)
        load_kw.append(dssCircuit.Loads.kW)
        load_kvar.append(dssCircuit.Loads.kvar)
        i = dssCircuit.Loads.Next

    return load_names, load_kw, load_kvar


def change_one_load(dssCircuit, dssSolution):
    print("Available loads: ", list_load_names(dssCircuit))
    bus = input("Which bus load would you like to change?")
    factor = int(input("How much would you like to change the load?"))
    factor = (1 + factor) / 100

    dssCircuit.Loads.Name = bus  # set active load to the bus specified by the user
    load_properties = dssCircuit.Loads  # get load properties for the active load
    load_kw = load_properties.kw
    load_kvar = load_properties.kvar

    dssCircuit.Loads.kw = load_kw * factor  # update real power of the active load
    dssCircuit.Loads.kvar = (
        load_kvar * factor
    )  # update the reactive power of the active load
    dssSolution.Solve()  # solve the power flow with the updated load values


# Changing multiple loads
newer_kw = []
newer_kvar = []


def change_multiple_loads(dssCircuit, dssSolution):
    factor = int(input("How much would you like to change all of the loads?(%)"))
    factor = (1 + factor) / 100

    load_idx = dssCircuit.Loads.First  # get the index of the first load in the circuit
    while load_idx > 0:
        load_name = dssCircuit.Loads.Name  # get the name of the current load
        dssCircuit.Loads.Name = load_name  # set the active load to the current load
        load_properties = dssCircuit.Loads  # get load properties for the active load

        load_kw = load_properties.kw  # get real power for the active load
        dssCircuit.Loads.kw = load_kw * factor  # updating real power of the active load

        load_kvar = load_properties.kvar  # get the reactive power of the active load
        dssCircuit.Loads.kvar = (
            load_kvar * factor
        )  # updating the reactive power of the active load

        newer_kw.append(dssCircuit.Loads.kw)
        newer_kvar.append(dssCircuit.Loads.kvar)

        load_idx = dssCircuit.Loads.Next

        # print(dssCircuit.Loads.kvar)
        # print(f"\n")
    dssSolution.Solve()


#
def change_multiple_specific_loads(dssCircuit, dssSolution):
    print("Available loads: ", list_load_names(dssCircuit))
    num_loads = int(input("How many loads would you like to change?"))

    for load in range(num_loads):
        bus = input("Which bus load would you like to change? ")
        factor = int(input("How much would you like to change the load?"))
        factor = (1 + factor) / 100
        dssCircuit.Loads.Name = bus  # set the active load to the user specified bus
        load_properties = dssCircuit.Loads  # get load properties for active load
        load_kw = load_properties.kw  # get the real power of the active load
        load_kvar = load_properties.kvar  # get the reactive power of the active load

        dssCircuit.Loads.kw = (
            load_kw * factor
        )  # update the real power of the active load
        dssCircuit.Loads.kvar = (
            load_kvar * factor
        )  # update the reactive power of the active load

    dssSolution.Solve()


def plot_data(voltages, loads):
    # Plot loads
    df_loads = pd.DataFrame({"Name": loads[0], "KW": loads[1], "Kvar": loads[2]})
    df_loads.plot(x="Name", y=["KW", "Kvar"], kind="bar", figsize=(12, 6))
    plt.title("Load Power for IEEE 30-bus system")
    plt.ylabel("Power (kW/kVar)")
    plt.xlabel("Load Name")
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.show()

    # Plot voltages
    plt.figure(figsize=(12, 6))
    plt.plot(voltages)
    plt.title("Bus Voltages (p.u.) for IEEE 30-bus system")
    plt.ylabel("Voltage (p.u.)")
    plt.xlabel("Bus Index")
    plt.grid(True)
    plt.tight_layout()
    plt.show()


def graph_compare_voltages(initial_voltages, modified_voltages):
    # Compare initial and modified voltages
    plt.figure(figsize=(12, 6))
    plt.plot(initial_voltages, label="Initial Voltages", marker="o")
    plt.plot(modified_voltages, label="Modified Voltages", marker="x", linestyle="--")
    plt.title("Bus Voltages (p.u.) for IEEE 30-bus system")
    plt.ylabel("Voltage (p.u.)")
    plt.xlabel("Bus Index")
    plt.grid(True)
    plt.legend()
    plt.tight_layout()
    plt.show()


def graph_compare_voltages_interactive(initial_voltages, modified_voltages):
    fig = go.Figure()

    fig.add_trace(
        go.Scatter(y=initial_voltages, mode="lines+markers", name="Initial Voltages")
    )
    fig.add_trace(
        go.Scatter(
            y=modified_voltages,
            mode="lines+markers",
            name="Modified Voltages",
            line=dict(dash="dot"),
        )
    )

    fig.update_layout(
        title="Bus Voltages (p.u.) for IEEE 30-bus system",
        xaxis_title="Bus Index",
        yaxis_title="Voltage (p.u.)",
        hovermode="closest",
    )

    fig.show()


def apply_multiple_line_outages(dssCircuit, dssText, dssSolution):
    """Applies outages to user-specified lines."""
    print("Available lines: ", list_line_names(dssCircuit))
    line_names = input("Enter line names to outage (comma separated): ").split(",")

    for line_name in line_names:
        if line_name.strip() in list_line_names(dssCircuit):
            dssText.Command = f"disable Line.{line_name.strip()}"
        else:
            print(f"Warning: Line {line_name} not found.")

    dssSolution.Solve()


def apply_multiple_transformer_outages(dssCircuit, dssText, dssSolution):
    """Applies outages to user-specified transformers."""
    print("Available transformers: ", list_transformer_names(dssCircuit))
    transformer_names = input(
        "Enter transformer names to outage (comma separated): "
    ).split(",")

    for transformer_name in transformer_names:
        if transformer_name.strip() in list_transformer_names(dssCircuit):
            dssText.Command = f"disable Transformer.{transformer_name.strip()}"
        else:
            print(f"Warning: Transformer {transformer_name} not found.")

    dssSolution.Solve()


def main():
    """
    This function is the main entry point of the OpenDSS simulation program. It sets up the OpenDSS environment,
    runs the simulation, and provides options to modify the simulation by changing loads or applying outages.
    It also compares the initial and modified voltages and exports the voltage data to a CSV file.

    Args:
        None

    Returns:
        None
    """
    dssObj, dssText, dssCircuit = setup_opendss()
    dssSolution = dssCircuit.Solution

    initial_voltages = run_simulation(dssText, dssCircuit)
    load_names, load_kw, load_kvar = get_load_details(dssCircuit)
    # plot_data(initial_voltages, [load_names, load_kw, load_kvar])

    print("Choose an option:")
    print("0. Exit")
    print("1. Change one load")
    print("2. Change all loads")
    print("3. Change multiple specific loads")
    print("\nSimulate Outages:")
    print("4. Apply one/multiple line outages")
    print("5. Apply one/multiple transformer outages")
    choice = int(input())

    if choice == 0:
        return
    elif choice == 1:
        change_one_load(dssCircuit, dssSolution)
    elif choice == 2:
        change_multiple_loads(dssCircuit, dssSolution)
    elif choice == 3:
        change_multiple_specific_loads(dssCircuit, dssSolution)
    elif choice == 4:
        apply_multiple_line_outages(dssCircuit, dssText, dssSolution)
    elif choice == 5:
        apply_multiple_transformer_outages(dssCircuit, dssText, dssSolution)
    else:
        print("Invalid option")
        return

    # After modification, compare the initial and modified voltages
    modified_voltages = dssCircuit.AllBusVmagPu

    graph_compare_voltages(initial_voltages, modified_voltages)
    # graph_compare_voltages_interactive(initial_voltages, modified_voltages)

    dssText.Command = "export voltages IEEE_30_CHANGED_Result.csv"
    # dssText.Command = "export powers kva elem IEEE_9500_poerws_kva.csv"
    # dssText.Command = r"Show Powers kVA Elem"


if __name__ == "__main__":
    main()

import win32com.client
import csv
import random
import pandas as pd

# Documentation: 
"""
This module is used to generate data for the machine learning model.
It uses the OpenDSS engine to run simulations and collect data.
Users can modify the parameters of loads and generators to generate data.
THe operations are:
    1. Setup OpenDSS
    2. Modify parameters of loads and generators
        i) Randomly select % of loads and generators user wants to modify
        ii) Randomly select loads and generators based on the % selected
        iii) Modify parameters of selected loads and generators randomly based on some threshold(-50% to 50%)
    3. Run simulation and collect data
    4. Store data to CSV
"""

FILE_PATH = "'A:\CCNY\J_Fall_2023\SD2\OpenDSS\IEEE 30 Bus\Master.dss'"
# FILE_PATH = (
#     "'A:\CCNY\J_Fall_2023\SD2\OpenDSS\ieee9500dss\ieee9500dss\ieee9500_base _copy.dss'"
# )

MAX_ITER = 1000
MAX_CONTROL_ITER = 100


def setup_opendss(dss_path):
    print(dss_path)
    dssObj = win32com.client.Dispatch("OpenDSSEngine.DSS")

    if not dssObj.Start(0):
        print("DSS failed to start!")
        exit()

    dssText = dssObj.Text
    dssCircuit = dssObj.ActiveCircuit
    dssSolution = dssCircuit.Solution

    dssText.Command = f"compile {dss_path}"  # Load the circuit
    dssText.Command = f"set maxiterations={MAX_ITER} maxControlIter={MAX_CONTROL_ITER}"

    return dssObj, dssText, dssCircuit, dssSolution


def modify_load_parameters(dssCircuit, load_name, kw_pct_change, kvar_pct_change):
    load = dssCircuit.Loads
    load.Name = load_name

    # Update kW if a percentage change was provided
    if kw_pct_change is not None:
        scaling_factor_kw = 1 + (kw_pct_change / 100)
        load.kW *= scaling_factor_kw

    # Update kVAr if a percentage change was provided
    if kvar_pct_change is not None:
        scaling_factor_kvar = 1 + (kvar_pct_change / 100)
        load.kvar *= scaling_factor_kvar


def modify_generator_parameters(dssCircuit, gen_name, kw_pct_change, kvar_pct_change):
    gen = dssCircuit.Generators
    gen.Name = gen_name

    # Update kW if a percentage change was provided
    if kw_pct_change is not None:
        scaling_factor_kw = 1 + (kw_pct_change / 100)
        gen.kW *= scaling_factor_kw

    # Update kVAr if a percentage change was provided
    if kvar_pct_change is not None:
        scaling_factor_kvar = 1 + (kvar_pct_change / 100)
        gen.kvar *= scaling_factor_kvar


def solve_and_fetch_results(dssCircuit, dssText, dssSolution):
    loads = dssCircuit.Loads
    gens = dssCircuit.Generators
    dssSolution.Solve()

    if not dssSolution.Converged:
        print("Solution did not Converge")
        exit()

    # Extract load data
    load_data = {}
    loads.First  # Set the first load as active
    load_idx = dssCircuit.Loads.First
    while load_idx > 0:
        load_data[loads.Name] = {"kW": loads.kW, "kvar": loads.kvar}
        load_idx = dssCircuit.Loads.Next

    # Extract generator data
    gen_data = {}
    gens.First  # Set the first generator as active
    gen_idx = dssCircuit.Generators.First
    while gen_idx > 0:
        gen_data[gens.Name] = {"kW": gens.kW, "kvar": gens.kvar}
        gens.Next  # Move to the next generator
        gen_idx = dssCircuit.Generators.Next

    buses = dssCircuit.AllBusNames
    voltages = dssCircuit.AllBusVmagPu

    data = {
        "Loads": load_data,
        "Generators": gen_data,
        "Voltages": dict(zip(buses, voltages)),
    }
    return data


def store_original_parameters(dssCircuit):
    original_loads = {}
    loads = dssCircuit.Loads
    i = loads.First
    while i > 0:
        original_loads[loads.Name] = {"kW": loads.kW, "kvar": loads.kvar}
        i = loads.Next

    original_gens = {}
    gens = dssCircuit.Generators
    i = gens.First
    while i > 0:
        original_gens[gens.Name] = {"kW": gens.kW, "kvar": gens.kvar}
        i = gens.Next

    return original_loads, original_gens


def reset_to_original_parameters(dssCircuit, original_loads, original_gens):
    loads = dssCircuit.Loads
    for load_name, params in original_loads.items():
        loads.Name = load_name
        loads.kW = params["kW"]
        loads.kvar = params["kvar"]

    gens = dssCircuit.Generators
    for gen_name, params in original_gens.items():
        gens.Name = gen_name
        gens.kW = params["kW"]
        gens.kvar = params["kvar"]


def collect_data_for_ml(dssCircuit, dssSolution):
    # Solving the circuit
    dssSolution.Solve()

    if not dssSolution.Converged:
        print("Solution did not converge")
        return None, None

    # Extract data for machine learning
    loads = dssCircuit.Loads
    gens = dssCircuit.Generators
    buses = dssCircuit.AllBusNames
    voltages = dssCircuit.AllBusVmagPu

    # Initialize dictionaries to collect features and labels
    features = {}
    labels = {}

    # Collect load and generator parameters as features
    idx = loads.First
    while idx > 0:
        features[f"load_{loads.Name}_kW"] = loads.kW
        features[f"load_{loads.Name}_kvar"] = loads.kvar
        idx = loads.Next

    idx = gens.First
    while idx > 0:
        features[f"gen_{gens.Name}_kW"] = gens.kW
        features[f"gen_{gens.Name}_kvar"] = gens.kvar
        idx = gens.Next

    # Collect voltages as labels
    for bus, voltage in zip(buses, voltages):
        labels[f"bus_{bus}_Vpu"] = voltage

    return features, labels


def store_to_csv(data, file_name):
    with open(file_name, "w", newline="") as csvfile:
        writer = csv.writer(csvfile)
        for key, value in data.items():
            writer.writerow([key])
            for sub_key, sub_value in value.items():
                writer.writerow([sub_key, sub_value])

    print("Data stored to CSV.")


def main():
    dssObj, dssText, dssCircuit, dssSolution = setup_opendss(FILE_PATH)

    number_of_simulations = 500
    percentage_of_loads_to_change = 70  # e.g., 70 percent
    percentage_of_generators_to_change = 30  # e.g., 30 percent
    original_loads, original_gens = store_original_parameters(dssCircuit)

    all_data = []

    # Define your randomization range for loads and generators
    load_kw_range = (-50, 50)  # Percent change
    load_kvar_range = (-10, 10)  # Percent change
    gen_kw_range = (-50, 50)  # Actual change in kW
    gen_kvar_range = (-10, 10)  # Actual change in kvar

    total_loads = len(dssCircuit.Loads.AllNames)
    total_generators = len(dssCircuit.Generators.AllNames)

    number_of_loads_to_change = max(
        1, int((percentage_of_loads_to_change / 100) * total_loads)
    )
    number_of_generators_to_change = max(
        1, int((percentage_of_generators_to_change / 100) * total_generators)
    )

    for i in range(number_of_simulations):
        print(f"Running simulation {i+1}/{number_of_simulations}")

        # Reset loads and generators to original values before each simulation
        reset_to_original_parameters(dssCircuit, original_loads, original_gens)

        # Select random loads and generators to modify
        selected_loads = random.sample(
            dssCircuit.Loads.AllNames, number_of_loads_to_change
        )
        selected_gens = random.sample(
            dssCircuit.Generators.AllNames, number_of_generators_to_change
        )

        # Randomly modify parameters of selected loads
        for load_name in selected_loads:
            modify_load_parameters(
                dssCircuit,
                load_name,
                kw_pct_change=random.uniform(*load_kw_range),
                kvar_pct_change=random.uniform(*load_kvar_range),
            )

        # Randomly modify parameters of selected generators
        for gen_name in selected_gens:
            modify_generator_parameters(
                dssCircuit,
                gen_name,
                kw_pct_change=random.uniform(*gen_kw_range),
                kvar_pct_change=random.uniform(*gen_kvar_range),
            )
        # Collect data
        features, labels = collect_data_for_ml(dssCircuit, dssSolution)
        if features is not None:
            all_data.append({**features, **labels})

    # Create a Pandas DataFrame
    df = pd.DataFrame(all_data)

    # Save DataFrame to CSV
    df.to_csv("training_data.csv", index=False)


if __name__ == "__main__":
    main()

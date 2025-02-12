from dataclasses import dataclass
import pandas as pd

# This file is used to create classes to store data for each type of component.
# The idea is to allow us to then input this into the generate_tech_profile as 
# a parameter rather than a long array

# Class definitions
@dataclass
class Component:
    model: str
    qty: int

@dataclass
class Compute:
    server: Component
    cpu: Component
    memory: str

# @dataclass
# class Storage:
#     storageModel: str
#     capacity: int

# @dataclass
# class Backup:
#     backupModel: str
#     backupDesc: str

# @dataclass
# class DR:
#     drApp: str
#     drDesc: str

# @dataclass
# class Virtualization:
#     virtName: str
#     virtDesc: str

# Component definitions

def gen_switch_data():
    # Networking
    df = pd.read_csv("switch_data.csv")

    switches = {}

    for _, row in df.iterrows():
        model = row['Switch Model']
        speed = row['Speed']
        ports = row['Num Ports']
        switches[model] = {
            'speed': speed,
            'ports': ports
        }

    # Print database (for debugging)
    # for name, specs in switches.items():
    #     print(f"Name of Switch: {name}")
    #     print(f"Speed of Switch: {specs['speed']}")
    #     print(f"Number of Ports: {specs['ports']}")

    return switches

from dataclasses import dataclass

# This file is used to create classes to store data for each type of component.
# The idea is to allow us to then input this into the generate_tech_profile as 
# a parameter rather than a long array

# Class definitions
@dataclass
class Networking:
    swithModel: str
    qty: int
    speed: str
    numPorts: int


# Component definitions

# Networking
switch_S4128F = Networking("PowerSwitch S4128F", 2, "10 GbE SFP+", 28)
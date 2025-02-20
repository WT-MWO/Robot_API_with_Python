import clr
import numpy as np

clr.AddReference(r"C:\Program Files\Autodesk\Robot Structural Analysis Professional 2023\Exe\Interop.RobotOM.dll")

from RobotOM import *
import RobotOM as rbt

# Input data
tower_h = 36
h_interval = 3
base_width = 6.1
slope = 0.05  # %

n_intervals = tower_h/h_interval


# Connect to Robot
# app = RobotApplication()

# project = app.Projectc
# structure = project.Structure #IRobotStructure

# app.Project.Structure.Clear  # Clears previous structure

# data structure

p1 = (0, 0, 0)
p2 = (base_width, 0, 0)
p3 = (base_width/2, 0, 0)


choord_1_x = np.arange(0, tower_h*slope, h_interval*slope)
choord_1_y = np.full(int(n_intervals), 0)


print(choord_1_x.size)
print(choord_1_y.size)
# Point creation
# p1 = [x,y,z]


# bar creation
# bar 1 connecting p1 and p4
# bar 2 connecting p2 and p5
# bar n connecting n and n+3

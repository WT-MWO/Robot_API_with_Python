import clr
# import numpy as np

clr.AddReference(r"C:\Program Files\Autodesk\Robot Structural Analysis Professional 2023\Exe\Interop.RobotOM.dll")

from RobotOM import *
import RobotOM as rbt

# Connect to Robot
# app = RobotApplication()

# project = app.Project
# structure = project.Structure #IRobotStructure

# app.Project.Structure.Clear  # Clears previous structure

# data structure
# [p1,p2,p3,p4,p5,p6]
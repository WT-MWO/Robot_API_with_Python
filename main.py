import clr
# import numpy as np

clr.AddReference(r"C:\Program Files\Autodesk\Robot Structural Analysis Professional 2023\Exe\Interop.RobotOM.dll")

from RobotOM import *
import RobotOM as rbt

# x_coords = np.arange(0, 12, 3, dtype=np.double)


app = RobotApplication()

project = app.Project
structure = project.Structure #IRobotStructure

# Create nodes
nodes = structure.Nodes  # IRobotNodeServer

nodes.Create(1, 0, 0, 0)
nodes.Create(2, 3, 0, 0)
nodes.Create(3, 6, 0, 0)
nodes.Create(4, 9, 0, 0)

# Create bars
bars = structure.Bars  # IRobotBarServer
bars.Create(1, 1, 2)
bars.Create(2, 2, 3)
bars.Create(3, 3, 4)

# Apply supports

labels = structure.Labels  # IRobotLabelServer

# Apply predefined pinned support
node = nodes.Get(1)
node.SetLabel(rbt.IRobotLabelType.I_LT_SUPPORT, 'Pinned')

# Define a new, custom, roller support
support_name = 'Roller'
support_label = labels.Create(rbt.IRobotLabelType.I_LT_SUPPORT, 'Roller')
support_data = rbt.IRobotNodeSupportData(support_label.Data)
print(support_data)
support_data.UX = 0
support_data.UY = 1
support_data.UZ = 1
support_data.RX = 0
support_data.RY = 0
support_data.RZ = 0
labels.Store(support_label)

# Apply roller support to nodes
for n in range(2,5):
    node = nodes.Get(n)
    node.SetLabel(rbt.IRobotLabelType.I_LT_SUPPORT, 'Roller')

# Apply section
all_bars = bars.GetAll()
print(all_bars)

for index in range(1, all_bars.Count+1):
    bar = rbt.IRobotBar(all_bars.Get(index))
    bar.SetLabel(rbt.IRobotLabelType.I_LT_BAR_SECTION, "IPE 100")


# Create loadcases



# Create load combination

# Calculate the model

# Display bending moment

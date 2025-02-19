import clr
# import numpy as np

clr.AddReference(r"C:\Program Files\Autodesk\Robot Structural Analysis Professional 2023\Exe\Interop.RobotOM.dll")

from RobotOM import *
import RobotOM as rbt

# Connect to Robot
app = RobotApplication()

project = app.Project
structure = project.Structure #IRobotStructure

app.Project.Structure.Clear  # Clears previous structure

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
support_label = labels.Create(rbt.IRobotLabelType.I_LT_SUPPORT, support_name)
support_data = rbt.IRobotNodeSupportData(support_label.Data)
support_data.UX = 0
support_data.UY = 1
support_data.UZ = 1
support_data.RX = 0
support_data.RY = 0
support_data.RZ = 0
labels.Store(support_label)

# Apply roller support to nodes
for n in range(2, 5):
    node = nodes.Get(n)
    node.SetLabel(rbt.IRobotLabelType.I_LT_SUPPORT, 'Roller')

# Apply section
all_bars = bars.GetAll()
for index in range(1, all_bars.Count+1):
    bar = rbt.IRobotBar(all_bars.Get(index))
    bar.SetLabel(rbt.IRobotLabelType.I_LT_BAR_SECTION, "IPE 100")

# Create loadcases
# Self-weight
sw_number = 1
sw_name = 'Self-weight'
nature = rbt.IRobotCaseNature.I_CN_PERMANENT
solver = rbt.IRobotCaseAnalizeType.I_CAT_STATIC_LINEAR
structure.Cases.CreateSimple(sw_number, sw_name, nature, solver)

# Live load
ll_number = 2
ll_name = 'Live load'
nature = rbt.IRobotCaseNature.I_CN_EXPLOATATION
structure.Cases.CreateSimple(ll_number, ll_name, nature, solver)


# Apply load
# Self-weight
case = rbt.IRobotSimpleCase(structure.Cases.Get(sw_number))
record_index = case.Records.New(rbt.IRobotLoadRecordType.I_LRT_DEAD)  # Crate new record - returns an index of the record
record = rbt.IRobotLoadRecord(case.Records.Get(record_index))  # Get newly created record by it's index
record.SetValue(2, -1)  # Sign for Z
record.SetValue(3, 1)  # Load factor
record.Objects.FromText('all')

# Live uniform load
ll_value = -500  # N/m
case = rbt.IRobotSimpleCase(structure.Cases.Get(ll_number))
record_index = case.Records.New(rbt.IRobotLoadRecordType.I_LRT_BAR_UNIFORM)
record = rbt.IRobotLoadRecord(case.Records.Get(record_index))
record.SetValue(2, ll_value)
record.Objects.FromText('all')

# Create load combination
comb_number = 3  # Or use: cases.FreeNumber
comb_name = "ULS 1"
comb_type = rbt.IRobotCombinationType.I_CBT_ULS
comb_nature = rbt.IRobotCaseNature.I_CN_PERMANENT
comb_analize_type = rbt.IRobotCaseAnalizeType.I_CAT_COMB
combination = structure.Cases.CreateCombination(comb_number, comb_name, comb_type, comb_nature, comb_analize_type)
case_factor_mng = combination.CaseFactors
case_factor_mng.New(1, 1.35)
case_factor_mng.New(2, 1.5)

# Calculate the model
app.Project.CalcEngine.Calculate()

# Display bending moment
view = rbt.IRobotView3(project.ViewMngr.GetView(1))
view.ParamsDiagram.Descriptions = rbt.IRobotViewDiagramDescriptionType.I_VDDT_LABELS  # IRobotViewDisplayParams
view.ParamsDiagram.Set(rbt.IRobotViewDiagramResultType.I_VDRT_NTM_MY, True)


print("Done.")

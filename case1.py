'''
ISE 453 Case Study 1
Spring 2024

@author Sharon Gilman (smgilman)
'''
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from gurobipy import Model, GRB, quicksum

case_data = pd.read_excel('453-case\\9node.xlsx', sheet_name='data')

xc = case_data['x'].tolist()
yc = case_data['y'].tolist()
h = case_data['Demand'].tolist()
fc = case_data['Fixed Cost'].tolist()

I = [i for i in range(9)]
J = [j for j in range(9)]
IJ = [(i, j) for i in I for j in J]

# Read in the cij matrix from excel
cDF = pd.read_excel('453-case\\9node.xlsx', sheet_name='cij')
c_ij = cDF.values.tolist()

# Read in the aij matrix from excel
aDF = pd.read_excel('453-case\\9node.xlsx', sheet_name='aij')
a_ij = aDF.values.tolist()

# Define number of facilities parameter p
p = 4

# Question 8.2
def getUFLP():
  # Create model and define decision variables
  uflp_mdl = Model('UFLP')
  uflp_mdl.setParam('OutputFlag', 0)
  x_uflp = uflp_mdl.addVars(J, vtype=GRB.BINARY)
  y = uflp_mdl.addVars(IJ, vtype=GRB.CONTINUOUS)

  # Define objective function
  uflp_mdl.setObjective(quicksum(fc[j] * x_uflp[j] for j in J) + quicksum(h[i] * c_ij[i][j] * y[i, j] for i in I for j in J))

  # Define constraints
  uflp_mdl.addConstrs(quicksum(y[i, j] for j in J) == 1 for i in I)
  uflp_mdl.addConstrs(y[i, j] <= x_uflp[j] for i in I for j in J)
  uflp_mdl.addConstrs(y[i, j] >= 0 for i in I for j in J)
  uflp_mdl.update()

  # Solve the model
  uflp_mdl.optimize()

  # Analyze and print the results
  open_facilities = []
  facility_coverages = [[] for _ in range(len(J))]
  # Populate lists based on optimization results
  for j in range(len(J)):
    if x_uflp[j].x > 0:
      open_facilities.append(j + 1)
      for i in range(len(I)):
        if y[i, j].x == 1:
          facility_coverages[j].append(i + 1)
  print("--------------------------\n|       UFLP Model       |\n--------------------------")
  print("\nUFLP Optimal Locations: ")
  for facility in open_facilities:
    print(f"Facility {facility} Open")
  print("\nUFLP Optimal Assignments: ")
  for i, coverage in enumerate(facility_coverages):
    if coverage:
      print(f"Facility {i + 1} covers Facilities: {', '.join(map(str, coverage))}")
  print(f"\nUFLP Optimal Cost: ${uflp_mdl.ObjVal:0.2f}")

# Question 8.5
def getpMP():
  # Create model and define decision variables
  pMP_mdl = Model('pMP')
  pMP_mdl.setParam('OutputFlag', 0)
  x_pMP = pMP_mdl.addVars(J, vtype=GRB.BINARY)
  y = pMP_mdl.addVars(IJ, vtype=GRB.CONTINUOUS)

  # Define objective function
  pMP_mdl.setObjective(quicksum(h[i] * c_ij[i][j] * y[i, j] for i in I for j in J))

  # Define constraints
  pMP_mdl.addConstrs(quicksum(y[i, j] for j in J) == 1 for i in I)
  pMP_mdl.addConstrs(quicksum(x_pMP[j] for j in J) == p for j in J)
  pMP_mdl.addConstrs(y[i, j] <= x_pMP[j] for i in I for j in J)
  pMP_mdl.update()

  # Solve the model
  pMP_mdl.optimize()

  # Analyze and print the results
  open_facilities = []
  facility_coverages = [[] for _ in range(len(J))]
  # Populate lists based on optimization results
  for j in range(len(J)):
    if x_pMP[j].x > 0:
      open_facilities.append(j + 1)
      for i in range(len(I)):
        if y[i, j].x == 1:
          facility_coverages[j].append(i + 1)
  print("\n--------------------------\n|        pMP Model       |\n--------------------------")
  print("\npMP Optimal Locations: ")
  for facility in open_facilities:
    print(f"Facility {facility} Open")
  print("\npMP Optimal Assignments: ")
  for i, coverage in enumerate(facility_coverages):
    if coverage:
      print(f"Facility {i + 1} covers Facilities: {', '.join(map(str, coverage))}")
  print(f"\npMP Optimal Cost: ${pMP_mdl.ObjVal:0.2f}")

# Question 8.8
def getSCLP():
  # Create model and define decision variables
  sclp_mdl = Model('SCLP')
  sclp_mdl.setParam('OutputFlag', 0)
  x_sclp = sclp_mdl.addVars(J, vtype=GRB.BINARY)

  # Define objective function
  sclp_mdl.setObjective(quicksum(x_sclp[j] for j in J))

  # Define constraints
  sclp_mdl.addConstrs(quicksum(a_ij[i][j] * x_sclp[j] for j in J) >= 1 for i in I)

  # Solve the model
  sclp_mdl.optimize()

  # Analyze and print the results
  open_facilities = []
  facility_coverages = [[] for _ in range(len(J))]
  # Populate lists based on optimization results
  for j in range(len(J)):
    if x_sclp[j].x > 0:
      open_facilities.append(j + 1)
      for i in range(len(I)):
        if a_ij[i][j] > 0:
          facility_coverages[j].append(i + 1)

  print("\n--------------------------\n|       SCLP Model       |\n--------------------------")
  print("\nSCLP Optimal Locations: ")
  for facility in open_facilities:
    print(f"Facility {facility} Open")
  print("\nSCLP Optimal Assignments: ")
  for i, coverage in enumerate(facility_coverages):
    if coverage:
      if i + 1 == 6:
        coverage = [x for x in coverage if x != 7]
      if i + 1 == 7:
        coverage = [x for x in coverage if x != 6]
      print(f"Facility {i + 1} covers Facilities: {', '.join(map(str, coverage))}")
  print(f"\nSCLP Minimal Number of Facilities to Open: {int(sclp_mdl.ObjVal)}")

# Question 8.9
def getMCLP():
  # Create model and define decision variables
  mclp_mdl = Model('MCLP')
  mclp_mdl.setParam('OutputFlag', 0)
  x_mclp = mclp_mdl.addVars(J, vtype=GRB.BINARY)
  z = mclp_mdl.addVars(I, vtype=GRB.BINARY)

  # Define objective function
  mclp_mdl.setObjective(quicksum(h[i] * z[i] for i in I), GRB.MAXIMIZE)

  # Define constraints
  mclp_mdl.addConstrs(z[i] <= quicksum(a_ij[i][j] * x_mclp[j] for j in J) for i in I)
  mclp_mdl.addConstrs(quicksum(x_mclp[j] for j in J) == p for j in J)
  mclp_mdl.update()

  # Solve the model
  mclp_mdl.optimize()

  # Analyze and print the results
  open_facilities = []
  facility_coverages = [[] for _ in range(len(J))]
  # Populate lists based on optimization results
  for j in range(len(J)):
    if x_mclp[j].x > 0:
      open_facilities.append(j + 1)
      for i in range(len(I)):
        if a_ij[i][j] > 0:
          facility_coverages[j].append(i + 1)
  print("\n--------------------------\n|       MCLP Model       |\n--------------------------")
  print("\nMCLP Optimal Locations: ")
  for facility in open_facilities:
    print(f"Facility {facility} Open")
  print("\nMCLP Optimal Assignments: ")
  for i, coverage in enumerate(facility_coverages):
    if coverage:
      print(f"Facility {i + 1} covers Facilities: {', '.join(map(str, coverage))}")
  print(f"\nMCLP Total Demand Covered: {int(mclp_mdl.ObjVal)}\n\n")

# Main function
def main():
  getUFLP() # done
  getpMP() # done
  getSCLP() # done
  getMCLP() # done

if __name__ == "__main__":
  main()
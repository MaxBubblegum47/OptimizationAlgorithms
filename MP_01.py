# -*- coding: utf-8 -*-
import gurobipy as gp
from gurobipy import GRB
import os
import xlrd # <-- this not works so well, better use directly CSV
from tabulate import tabulate


def read_xlsx(file_path, sheet_name='Foglio1'):
    """

    Args:
        :param file_path: (str) path to the input file.
        :param sheet_name: (str) name of the excel sheet to read
    :return:
    """
    products, machines, periods = [], [], []
    numMachines = []

    # Open the .xlsx file
    book = xlrd.open_workbook(file_path)
    sh = book.sheet_by_name(sheet_name)

    # Reads the products (first row:  ...)
    col = 1
    while True:
        try:
            products.append(sh.cell_value(0, col))
            col = col + 1
        except IndexError:
            break
    products = products[:-1]  # removes the last col
    #print("Products = ", products)

    # Reads the profits
    profits = {j : 0 for j in products}
    col = 1
    for j in products:
        profits[j] = sh.cell_value(1, col)
        col = col + 1
    #print("Profits = ", profits)


    # Reads the machines 
    row = 2
    m = str(sh.cell_value(row, 0))
    while m != "":
        try:
            machines.append(m)
            row = row + 1
            m = str(sh.cell_value(row, 0))
        except IndexError:
            break
    #print("Machines = ", machines)


    # Reads the matrix of production times
    A = {(m, j): 0 for m in machines for j in products}
    for row_idx, m in enumerate(machines):
        for col_idx, j in enumerate(products):
            A[m, j] = float(sh.cell_value(row_idx + 2, col_idx + 1 ))

    #reads the number of machines        
    col = len(products) + 1
    for row_idx, m in enumerate(machines):
        numMachines.append(sh.cell_value(row_idx + 2, col))
    #print("NumMachines = ", numMachines)

    # Reads the periods
    row = 2
    t = str(sh.cell_value(row, 0))
    while t.lower() != "max sales":
        try:
            row = row + 1
            t = str(sh.cell_value(row, 0))
        except IndexError:
            break
    
    rowSales = row
    row = row + 1
    t = str(sh.cell_value(row, 0))
    while t != '':
        try:
            periods.append(t)
            row = row + 1
            t = str(sh.cell_value(row, 0))
        except IndexError:
            break

    #print("Periods = ", periods)

    # Reads the matrix of maximum demands
    MAXS = {(t, j): 0 for t in periods for j in products}
    for row_idx, t in enumerate(periods):
        for col_idx, j in enumerate(products):
            MAXS[t, j] = float(sh.cell_value(row_idx + rowSales + 1, col_idx + 1))

 
 
    # Reads the unavailability and prepare matrix MC
    row = 2
    t = str(sh.cell_value(row, 0))
    while t.lower() != "unavailable":
        try:
            row = row + 1
            t = str(sh.cell_value(row, 0))
        except IndexError:
            break

    rowUnav = row 
    
    MC = {(m, t): 0 for m in machines for t in periods}   
    
    for row_idx, t in enumerate(periods):
        for col_idx, m in enumerate(machines):
            MC[m, t] = numMachines[col_idx] - float(sh.cell_value(row_idx + rowUnav + 1, col_idx + 1))
    
    #print("Available machines",MC)
     
    return products, machines, periods, A, profits, MAXS, MC


def prev_month(t):
    if t == 'January':
        return 'January'

    if t == 'February':
        return 'January'
    
    if t == 'March':
        return 'February'
    
    if t == 'April':
        return 'March'
    
    if t == 'May':
        return 'April'
    
    if t == 'June':
        return 'May'



def make_model(products, machines, periods, A, profits, MAXS, MC, model_name='MP_01'):
    
    m = gp.Model(model_name)
    
    # Variables
    # Product Made
    x = m.addVars(periods, products, vtype=GRB.INTEGER, name="x")

    # Product Sold
    s = m.addVars(periods, products, vtype=GRB.INTEGER, name="sales")
    
    # Product stored in inventory
    i = m.addVars(periods, products, vtype=GRB.INTEGER, name="inventory")

    # Constraints
    # flow conservation - me, now I need to say that in january is all zero
    m.addConstrs(i[t,j] == i[prev_month(t),j] + x[t,j] - s[t,j]  if t != 'January' else (i[t,j]  -x[t,j] + s[t,j] == 0) for t in periods for j in products)
    m.addConstrs(i[t,j] == 50 for t in periods for j in products if t == 'June' )
    
    #flow conservation - teacher
    # m.addConstrs((i[t,j]  -x[t,j] + s[t,j] == 0 for t in periods[0:1] for j in products))
    # m.addConstrs((i[t,j] - i[periods[periods.index(t)-1], j] -x[t,j] + s[t,j] == 0 for t in periods[1:] for j in products))
    
    #inventory
    m.addConstrs((i[t,j] <= 100 for t in periods for j in products))
    m.addConstrs(i[periods[len(periods)-1], j] == 50  for j in products)
    
    #sales
    m.addConstrs((s[t,j] <= MAXS[t,j] for t in periods for j in products))
    
    #time
    m.addConstrs(((gp.quicksum(A[i,j] * x[t,j] for j in products) <= 16*24*MC[i,t]) 
                  for t in periods for i in machines)) 

    # Objective
    m.setObjective(gp.quicksum(profits[j] * s[t,j] for t in periods for j in products) 
                   -gp.quicksum(0.5 * i[t,j] for t in periods for j in products), GRB.MAXIMIZE)
    

    m.write("MP_01.lp")
    
    return m


if __name__ == '__main__':
    # Read the data from the input file
    products, machines, periods, A, profits, MAXS, MC = read_xlsx(os.path.join('/home/maxbubblegum/Desktop/OptimizationAlgorithms', 'MP_01.xls'))

    # Make the model and solve
    model = make_model(products, machines, periods, A, profits, MAXS, MC)
    
    model.optimize()
    
    model.write(os.path.join('/home/maxbubblegum/Desktop/OptimizationAlgorithms', 'MP_01.lp'))

    # Print in a better way, use tabulate
    if model.status == GRB.Status.OPTIMAL:      
        print(f'\nProfit: {model.objVal:.6f}\n')
        # Print single variables
        for i in model.getVars():
            if i.x > 0.0001:
                print(f'Var. {i.varName:22s} = {i.x:6.1f}')
            
    else:
        print(">>> No feasible solution")
        
    

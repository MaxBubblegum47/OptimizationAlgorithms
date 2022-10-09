# -*- coding: utf-8 -*-
"""
Created on Sun Oct  9 11:22:02 2022

@author: loren
"""
import gurobipy as gp
from gurobipy import GRB
import os
import xlrd



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
    print("Products = ", products)

    # Reads the profits
    profits = {j : 0 for j in products}
    col = 1
    for j in products:
        profits[j] = sh.cell_value(1, col)
        col = col + 1
    print("Profits = ", profits)


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
    print("Machines = ", machines)


    # Reads the matrix of production times
    A = {(m, j): 0 for m in machines for j in products}
    for row_idx, m in enumerate(machines):
        for col_idx, j in enumerate(products):
            A[m, j] = float(sh.cell_value(row_idx + 2, col_idx + 1 ))

    #reads the number of machines        
    col = len(products) + 1
    for row_idx, m in enumerate(machines):
        numMachines.append(sh.cell_value(row_idx + 2, col))
    print("NumMachines = ", numMachines)

    # Reads the periods
    row = 2
    t = str(sh.cell_value(row, 0))
    while t.lower() != "max sales":
        try:
            row = row + 1
            t = str(sh.cell_value(row, 0))
        except IndexError:
            break
    #
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

    print("Periods = ", periods)

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
    #
    MC = {(m, t): 0 for m in machines for t in periods}   
    
    for row_idx, t in enumerate(periods):
        for col_idx, m in enumerate(machines):
            MC[m, t] = numMachines[col_idx] - float(sh.cell_value(row_idx + rowUnav + 1, col_idx + 1))
    #
    print("Available machines",MC)
     
    return products, machines, periods, A, profits, MAXS, MC


def make_model(products, machines, periods, A, profits, MAXS, MC, model_name='MP01'):
    
    # create the model
    m = gp.Model(model_name)
    
    # variables
    x = m.addVars(periods, product, vtype=GRB.CONTINUOUS, name="x")
    
    s = m.addVars(periods, product, vtype=GRB.CONTINUOUS, name="sales")
    
    i = m.addVars(periods, product, vtype=GRB.CONTINUOUS, name="inventory")
    
    
    # constraints
    
    # flow balance
    # idk : the main idea is that i[t,j] = i[t-1,j] + x[t,j] - s[t,j]
    # how can I represent time - 1?
    
    # Inventory limits
    m.addConstrs(i[t,j] <= 100 for t in periods for j in products)
    m.addConstrs(i[0,j] == 0 for j in products)
    m.addConstrs(i[periods[len(periods)-1],j] == 50 for j in products)
      
    # maximum sales
    m.addConstrs(s[t,j] <= MAXS[t,j] for t in periods for j in products)
    
    # production time
    m.addConstrs(quicksum(A[i,j]*x[t,j]) <= 16*24*MC[i,t] for i in machines for t in periods)    
    
    
    return m


if __name__ == '__main__':
    # Read the data from the input file
    products, machines, periods, A, profits, MAXS, MC = read_xlsx(os.path.join('C:\\Users\\loren\\Desktop', 'MP_01.xlsx'))
    
    model = make_model(products, machines, periods, A, profits, MAXS, MC)

    model.optimize()
    
    model.write(os.path.join('C:\\Users\\loren\\Desktop', '05_productionMix.lp'))

    if model.status == GRB.Status.OPTIMAL:      
        print(f'\nProfit: {model.objVal:.6f}\n')
        # Print single variables
        for i in model.getVars():
            if i.x > 0.0001:
                print(f'Var. {i.varName:22s} = {i.x:6.1f}')
                
    print()
    
    for t in periods:
        for i in machines:
            s = 'TC[' + t + ','+i+']'
            cnt = model.getConstrByName(s)
            print(f'{cnt.ConstrName:32s} slack = {cnt.slack:10.3f} RHS = {cnt.RHS}')
            
    
    
    
    
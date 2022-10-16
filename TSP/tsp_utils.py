# -*- coding: utf-8 -*-
"""
Created on Mon Oct 10 15:00:58 2022

@author: Mauro
"""
import math
import random
from csv import reader
import matplotlib.pyplot as plt
import tsplib95 
import gurobipy as gb
from gurobipy import GRB


def plot_tour(points, edges, title="Tour", figsize=(12, 12), save_fig=None):
    """
    Plot a tour.

    :param points: list of points.
    :param edges: list of selected edges
    :param title: title of the figure.
    :param figsize: width and height of the figure
    :param save_fig: if provided, path to file in which the figure will be save.
    :return: None
    """

    plt.figure(figsize=figsize)
    plt.xlabel("X")
    plt.ylabel("Y")
    plt.title(title, fontsize=15)
    plt.grid()
    x = [pnt[0] for pnt in points]
    y = [pnt[1] for pnt in points]
    plt.scatter(x, y, s=60)

    n=len(points)

    # Add label to points
    for i, label in enumerate(points):
        plt.annotate('{}'.format(i), (x[i]+0.1, y[i]+0.1), size=25)

   # Add the edges
    for i in range(n):
        for j in range(n):
            if j < i:
                plt.plot([points[i][0], points[j][0]], [points[i][1], points[j][1]], 'b', alpha=.5)
    # plots the selected edges (sub)tours
    for (i, j) in edges:
       plt.plot([points[i][0], points[j][0]], [points[i][1], points[j][1]], 'r', alpha=1.,linewidth=3)
       
    if save_fig:
        plt.savefig(save_fig)
    else:
        plt.show()
        
def read_csv_points(file_path, sep=',', has_headers=True, dtype=float):
    """
    Read a csv file containing 2D points.

    :param file_path: path to the csv file containing the points
    :param sep: csv separator (default ',')
    :param has_headers: whether the file has headers (default True)
    :param dtype: data type of values in the input file (default float)
    :return: list of points
    """
    with open(file_path, 'r') as f:
        csv_r = reader(f, delimiter=sep)

        if has_headers:
            headers = next(csv_r)
            print('Headers:', headers)

        points = [tuple(map(dtype, line)) for line in csv_r]
        print(points)

    return points

def EuclDist(points):
    """
    generates a dictionary of Euclidean distances between pairs of points    

    Parameters
    ----------
    points : list of pair of coordinates

    """
    # Dictionary of Euclidean distance between each pair of points
    dist = {(i, j):
            math.sqrt(sum((points[i][k]-points[j][k])**2 for k in range(2)))
            for i in range(len(points)) for j in range(len(points))}
            #for i in range(len(points)) for j in range(i)}
    return dist        

def randomEuclGraph (n, maxcoord):
    """
    generates an instance of an Euclidean graph
    Parameters
    ----------
    n number of vertices
    maxcoord maximum value of a coordinate of a vertex
    """
    points = [(random.randint(0, maxcoord), random.randint(0, maxcoord)) for i in range(n)]
    dist = EuclDist(points)
    print(points)
    print(dist)
    return points, dist
    
def readTSPLIB(file_path):
    problem = tsplib95.load(file_path)
    n = problem.dimension
    nodes = list(problem.get_nodes())
    points = tuple(problem.node_coords.values())
    if len(points) == 0:
        points = tuple(problem.display_data.values())
    # shift nodes to start from 0
    nodes = [x-1 for x in nodes]
    #shift nodes to start from 0
    edges = list(problem.get_edges())
    edges = [(i-1,j-1) for (i,j) in edges]
    dist = {(i,j) : 0 for  (i,j) in edges}
    for (i,j) in edges:
        dist[i,j] = problem.get_weight(i+1, j+1)
    
    return n, points, dist

def make_model(n, points, dist):
    m = gb.Model("atsp")

    x = m.addVars(dist.keys(), vtype=GRB.BINARY, name = "x")
    u = m.addVars(n, vtype=GRB.INTEGER, name = "u")

    M = 100000
    m.addConstrs(quicksum(x[i,j] == 1 for j in n) for i in n)
    m.addConstrs(quicksum(x[i,j] == 1 for i in n) for j in n)
    m.addConstrs(u[1] == 0)
    m.addConstrs(u[j] - u[i] >= 1 - M(1 - x[i,j]) for i,j in A and j != 1)
    m.addConstrs(u[i] >= 0 for i in n)


if __name__ == '__main__':
    n = 5
    maxcoord = 10
    print(n)

    points, dist = randomEuclGraph(n, maxcoord)

    print(points)

    model = make_model(n, points, dist)

    model.optimize()

    model.write(os.path.join('/home/maxbubblegum47/Desktop', 'tsp.lp'))


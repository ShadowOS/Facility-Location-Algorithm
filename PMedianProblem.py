# -*- coding: utf-8 -*-
"""
Created on Tue Apr 16 18:56:20 2019

@author: Karnika
"""

# -*- coding: utf-8 -*-
"""
Created on Thu Apr 11 00:38:26 2019

@author: Karnika
"""

from gurobipy import*
import os
import xlrd
import xlwt
import xlsxwriter

book = xlrd.open_workbook(os.path.join("d.xlsx"))

sh = book.sheet_by_name("Sheet3")
Node = []
demand={}
MinNumOfFacility=5

i = 1
while True:
    try:
        sp = sh.cell_value(i,0)
        Node.append(sp)
        demand[sp]=sh.cell_value(i,14)
        i = i + 1
        
    except IndexError:
        break
Cost = {}
i = 1
for x in Node:
    j = 1
    for y in Node:
        Cost[x,y] = sh.cell_value(i,j)
        j += 1
    i += 1

sh = book.sheet_by_name("Sheet4")

Aij = {}
i = 1
for x in Node:
    j = 1
    for y in Node:
        Aij[x,y] = sh.cell_value(i,j)
        j += 1
    i += 1
def minimumDistance(p,q):
    m=Model("d")    
    X=m.addVars(Node,Node,vtype=GRB.BINARY,name="X_ij")
    m.modelSense=GRB.MINIMIZE    
    m.setObjective(sum(Cost[i,j]*X[i,j] for i in Node for j in Node)) 
    for i in Node:
        if i == p:
            m.addConstr(sum(X[j,i] for j in Node if  Aij[j, i] == 1) - sum(X[i,k] for k in Node if  Aij[i, k] == 1) == -1)
        elif i == q:
            m.addConstr(sum(X[j,i] for j in Node if  Aij[j, i] == 1) - sum(X[i,k] for k in Node if  Aij[i, k] == 1) == 1)
        else:
            m.addConstr(sum(X[j,i] for j in Node if  Aij[j, i] == 1) - sum(X[i,k] for k in Node if  Aij[i, k] == 1) == 0)
    
    print('Origin node = ', p)
    print('Destination node = ', q)           
    m.optimize()
    m.write('md1.lp')                            
    for v in m.getVars():
        if v.x > 0.01:
            print(v.varName, v.x)
    print('Objective:',m.objVal)
    return(m.objVAl)
md={}            
for i in Node:
    for j in Node:
        if i!=j and j > i:
            md[i,j] = minimumDistance(i,j)
            
for i in Node:
    for j in Node:
        if i==j:
            md[i,j]= 0
for i in Node:
    for j in Node:
        if i>j:
            md[i,j]=md[j,i]
X=MinNumOfFacility-1
P=0
            
demandweightedmatrix={}
for i in Node:
    for j in Node:
        demandweightedmatrix[i,j]=md[i,j]*demand[j]        
a={}
b={}        
def sumdwm(p):
    a[p]=sum(demandweightedmatrix[p,i] for i in Node )
    return(a)
for i in Node:
    b[i]=sumdwm(i)
min_a = min(a[x] for x in a)
FacilityLocatedLocation = [key for key, value in a.items() if value == min_a]
P=1
P=P-1
for x in FacilityLocatedLocation: 
    if P<X :
        p=FacilityLocatedLocation[P]
        for i in Node:
            for j in Node:
                if demandweightedmatrix[i,j] < demandweightedmatrix[p,j]:
                    demandweightedmatrix[i,j]=demandweightedmatrix[i,j]
                elif demandweightedmatrix[i,j] > demandweightedmatrix[p,j] :
                    demandweightedmatrix[i,j]=demandweightedmatrix[p,j]
                else:
                    demandweightedmatrix[i,j]=demandweightedmatrix[i,j]
        for i in Node:
            b[i]=sumdwm(i)
        min_a = min(a[x] for x in a)
        FacilityLocatedLocation2 = [key for key, value in a.items() if value == min_a]
        sp=FacilityLocatedLocation2[0]
        FacilityLocatedLocation.append(sp)
        P+=1
workbook=xlsxwriter.Workbook('emd.xlsx')
worksheet=workbook.add_worksheet('result')
row=0
col=0
for i in b:
    c=a[i]
    d=Node[row]
    worksheet.write(row,0,d)
    worksheet.write(row,1,c)
    row+=1
worksheet=workbook.add_worksheet('md')
i=1
k=0
for x in Node:
    j=1
    l=0
    for y in Node:
        e=md[x,y]
        worksheet.write(i,j,e)
        f=Node[k]
        worksheet.write(i,0,f)
        g=Node[l]
        worksheet.write(0,j,g)
        l+=1
        j+=1    
    i+=1
    k+=1        
worksheet=workbook.add_worksheet('demandweightedmatrix')
i=1
k=0
for x in Node:
    j=1
    l=0
    for y in Node:
        e=demandweightedmatrix[x,y]
        worksheet.write(i,j,e)
        f=Node[k]
        worksheet.write(i,0,f)
        g=Node[l]
        worksheet.write(0,j,g)
        l+=1
        j+=1    
    i+=1
    k+=1
worksheet=workbook.add_worksheet('finalresult')
row=0
col=0
for i in FacilityLocatedLocation:
    c=FacilityLocatedLocation[row]
    worksheet.write(row,0,c)
    row+=1
workbook.close()

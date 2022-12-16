# -*- coding: utf-8 -*-
"""
Created on Fri Dec 16 10:38:14 2022

@author: baret
"""

import itertools
import numpy as np 
import xlsxwriter

def calc_sol():
    L=[[10,11,12],[20,21,22],[30,31,32]]
    L1=[0,1,2]
    L2=[0,1,2]
    L3=[0,1,2]
    
    N1=(list(itertools.combinations(L1, 1)))
    N2=list(itertools.combinations(L2, 1)) 
    N3=(list(itertools.combinations(L3, 1)))
    
    Sol=[]
    
    for k1 in range(0,3):
        for k2 in range(0,3):
            for k3 in range(0,3):
               New=N1[k1]+N2[k2]+N3[k3]
               Sol.append(New)
    return Sol 

def print_Sol(filename,Sol,nbF,nbL):
    Sol2=[]
    #transformer la solution en binaire
    for i in range(len(Sol)):
        Size=nbF*nbL
        print(Size)
        z=np.zeros(Size)
        for j in range(len(Sol[i])):
            print(i*nbL+(Sol[i][j]))
            z[j*nbL+(Sol[i][j])]=1

            print(z)
        Sol2.append(z)
        
    print(Sol2)
    
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet('Result')
    
    for k1 in range(len(Sol2)):
        for k2 in range(len(Sol2[k1])):
            print('write')
            worksheet.write(k1,k2,Sol2[k1][k2])
        workbook.define_name(f'Solution{k1}',  f'=Result!$A${k1}:$I${k1}')
    workbook.close()
            
filename ='C:/Users/baret/Documents/Markov/sol_possibles.xlsx'

Sol=calc_sol()
  
print_Sol(filename, Sol, 3, 3)           
# -*- coding: utf-8 -*-
"""
Created on Fri Dec 16 14:54:31 2022

@author: baret
"""

# -*- coding: utf-8 -*-
"""
Created on Tue Nov 15 15:54:16 2022

@author: baret
"""


import pandas as pd
import xlsxwriter
import numpy as np

def aggr_result(directory,  instance_size,model):
    file_result=directory + '/Resultats_compare.xlsx'
    workbook = xlsxwriter.Workbook(file_result)
    worksheet = workbook.add_worksheet('Result')
    worksheet.write('A1','Instance')
    worksheet.write('C1','H1')
    worksheet.write('D1','H2')

    for i in range(0,instance_size) :
        fichier= f'{directory}/Instance/instance{i}.xlsx'
        fichierSol= (directory + f'/Resultats/resultat_{i}.xlsx')
        #donnees générales 
            
       # Donnees=pd.read_excel(fichier,sheet_name="Donnees", index_col=0 )      
        #données facilities 
 #       Capacites=pd.read_excel(fichier,sheet_name="Congestions", index_col=0 )
                
    #    Do=Donnees.to_numpy()
               
         # données solutions   
        Resolution_time=pd.read_excel(fichierSol,sheet_name="Sheet1", index_col=0)
        RT=Resolution_time.to_numpy()

        R=RT.size
        if R>0:
            print(i)
            ResT=RT[0,3]
            worksheet.write(i+1,1,ResT)
            H1=RT[2,2]
            print(H1)
            worksheet.write(i+1,2,H1)
            H2=RT[2,3]
            print(H2)
            worksheet.write(i+1,3,H2)
            #BUD=RT[2,7]
        #    worksheet.write(i+1,4,BUD)
       # nbClients=Do[0]
      #  nbFacilities=Do[1]
        worksheet.write(i+1,0,i)
        #worksheet.write(i+1,1,nbClients)
       # worksheet.write(i+1,2,nbFacilities)
        
    workbook.close()
    return 
        


directory1='C:/Users/baret/Documents/GitHub/PMFP-Markov-Simulation/2_niveaux'
model='C'  
RT=aggr_result(directory1, 27,model)


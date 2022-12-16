# -*- coding: utf-8 -*-
"""
Created on Fri Dec 16 11:38:02 2022

@author: baret
"""

import  os
import lingo_relation 
import xlsxwriter

def resolution(instance_debut,instance_fin, directory):
    #3
    instances = [(i) for i in range(instance_debut,instance_fin)]

    for instance in instances:
       lingo_relation.create_lingo_ltf_file(directory,instance)
       print('running instance '+str(instance) )
       fileSol= directory+'/Resultats/resultat_'+str(instance)+'.xlsx'
       # print(fileSol)
       workbook = xlsxwriter.Workbook(fileSol)

       workbook.close()
       filemodel=directory+'/Modeles/model_'+str(instance)+'.ltf'
       print(filemodel)
       os.system(f'runlingo {filemodel}')

    return 

directory='C:/Users/baret/Documents/GitHub/PMFP-Markov-Simulation/2_niveaux'
resolution(20,30,directory)
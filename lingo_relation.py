# -*- coding: utf-8 -*-
"""
Created on Fri Dec 16 13:27:27 2022

@author: baret
"""

def create_lingo_ltf_file(directory, instance_number):
   print('crea file lingo ')
   instance_folder_path = directory
   instance = f'/Instances/instance{instance_number}.xlsx'
   
   lin_model = open(f'{directory}/Modeles/model_{instance_number}.ltf','w')
   print(lin_model)
   lin_model.writelines('set default\nset echoin 1\n\n')

   """
      # First data section
   """

   lin_model.writelines('MODEL:\n\n')
     
   """
      # Sets section
   """
   write_sets(instance_folder_path,instance,lin_model )

   """
      # Data section
   
   """
   lin_model.writelines(f'DATA:\n')
   write_data(instance_folder_path, instance, lin_model)

   write_results_data(instance_folder_path,instance_number,lin_model,)
   lin_model.writelines(f'ENDDATA\n\n')

   """
      # Objective function 
   """

   write_the_objective_function_init(lin_model)

    
   lin_model.writelines(f'END\n\n')

   """
     # Add command for runlingo
   """

   lin_model.writelines(f'set terseo 1\n')
  # lin_model.writelines(f'set timlim 120\n')
   lin_model.writelines(f'set MULTIS 3\n ')
   lin_model.writelines(f'go\n')
   lin_model.writelines(f'nonz volume\n')
   lin_model.writelines(f'quit\n')
   
   lin_model.close()

   return 

def write_sets(instance_folder_path, instance,lingo_model):

   lingo_model.writelines(f'SETS:\n')

   lingo_model.writelines(f'Clients: demand, EC, penalty;\n')
   lingo_model.writelines(f'Facilities: Cap;\n')
   lingo_model.writelines('F;\n')
   lingo_model.writelines(f'Levels;\n')
   lingo_model.writelines(f'SORTED(Clients,Facilities):sort, positions, P, distance;\n')
   lingo_model.writelines('LINK2(Clients,F): proba;\n')
   lingo_model.writelines(f'LINKS(Facilities, levels): cost,z;\n')

   lingo_model.writelines(f'ENDSETS\n\n')
   return 

def write_data(instance_folder_path, instance, lingo_model):
 
   lingo_model.writelines(f'Number_clients=@ole(\'{instance_folder_path}{instance}\',\'NbClients\');\n')
   lingo_model.writelines(f'Number_facilities=@ole(\'{instance_folder_path}{instance}\',\'NbFacilities\');\n')  
   lingo_model.writelines(f'Number_levels=@ole(\'{instance_folder_path}{instance}\',\'NbLevels\');\n')  
   lingo_model.writelines(f'N2=@ole(\'{instance_folder_path}{instance}\',\'n\');\n')
   lingo_model.writelines( f'Clients=1..Number_clients;\n')
   lingo_model.writelines(f'Facilities=1..Number_facilities; \n')
   lingo_model.writelines('F=1..N2;\n')
   lingo_model.writelines(f'Levels=1..Number_levels; \n')

   lingo_model.writelines(f'demand = @ole(\'{instance_folder_path}{instance}\',\'Demandes\');\n')
   lingo_model.writelines(f'penalty = @ole(\'{instance_folder_path}{instance}\',\'Penalites\');\n')
   lingo_model.writelines(f'distance = @ole(\'{instance_folder_path}{instance}\',\'Distances\');\n')
   lingo_model.writelines(f'cost = @ole(\'{instance_folder_path}{instance}\',\'Cout\');\n')

   lingo_model.writelines(f'proba = @ole(\'{instance_folder_path}{instance}\',\'Proba\');\n')

   return


def write_results_data(instance_folder_path,instance_number,lingo_model):

   results_folder_path = instance_folder_path+'/Resultats/'
   results = f'resultat_{instance_number}.xlsx'
   #results_folder_path = 'test'
   #results = '1'
  

   lingo_model.writelines('!Results;\n')
  
   lingo_model.writelines(f'@ole(\'{results_folder_path}{results}\',\'B2\')=@WRITE(\'Objectif\');\n')
   lingo_model.writelines(f'@ole(\'{results_folder_path}{results}\',\'C2\')=@WRITE(H1+H2);\n')

   lingo_model.writelines(f'@ole(\'{results_folder_path}{results}\',\'B3\')=@WRITEFOR(LINKS(k,l):z(k,l));\n')
   lingo_model.writelines(f'@ole(\'{results_folder_path}{results}\',\'D2\')=@WRITE(\' The resolution time is : \',@TIME(),\' seconds\',@NEWLINE(1));\n\n')        
   lingo_model.writelines(f'@ole(\'{results_folder_path}{results}\',\'C3\')=@WRITEFOR(Facilities(k):Cap(k));\n')
   
   lingo_model.writelines(f'@ole(\'{results_folder_path}{results}\',\'D3\')=@WRITE(\'H1\');\n')
   lingo_model.writelines(f'@ole(\'{results_folder_path}{results}\',\'D4\')=@WRITE(H1);\n')
   
   lingo_model.writelines(f'@ole(\'{results_folder_path}{results}\',\'E3\')=@WRITE(\'H2\');\n')
   lingo_model.writelines(f'@ole(\'{results_folder_path}{results}\',\'E4\')=@WRITE(H2);\n')
 
   
   return 

    
def write_the_objective_function_init(lingo_model):

   lingo_model.writelines(f'!Objective function;\n')
   lingo_model.writelines(f'[obj] Min=H1+H2;\n\n')
#   lingo_model.writelines('N2=Number_facilities+1;\n')
   lingo_model.writelines('H1=(@sum(Clients(i):demand(i)*@sum(SORTED(i,k1): proba(i,k1)*distance(i,k1))));\n')
   lingo_model.writelines('H2=(@sum(Clients(i):proba(i,N2)*penalty(i)*demand(i)));\n')
   return 



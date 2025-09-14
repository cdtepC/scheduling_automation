#!/usr/bin/env python
# coding: utf-8

# # Personalplanung MW als GAP
# 
# ## 1. Erstellung eines Schichtplans
# ### 1.1 Daten beziehen und verarbeiten
# ### 1.2 Modell bauen und lösen
# ### 1.3 Dokument erstellen und ausgeben
# ### 1.4 Programm für den Echteinsatz
# ## 2. Evaluation des Modells
# ### 2.1 Simulation vorbereiten
# ### 2.2 Probleme simulieren, lösen und Testwerte erheben
# ### 2.3 Ergebnisse als Dokumente ausgeben
# ### 2.4 Versuchsreihe 

# In[1]:


import sys
try:
    import docplex.mp
except:
    if hasattr(sys, 'real_prefix'):
        #we are in a virtual env.
        get_ipython().system('pip install docplex')
    else:
        get_ipython().system('pip install --user docplex')
from docplex.mp.model import Model

# echteinsatz
import pandas as pd
import regex as rg
import re
import math
import numpy as np
import os
import xlwt 
from xlwt import Workbook

# analysis
import json
from tqdm import tqdm
import time as t
import random




class tutor:
    def __init__(self,prename,shift_min,vgA,new,support):
        self.prename = prename
        self.shift_min = shift_min # minimale Schichtdauer
        self.vgA = vgA # vertraglich geregelte Arbeitszeit 
        self.new = new # das erstse Semester dabei? (boolean)
        self.support = support # unterstützte Zeitfenster + Label (dictionary)
    def print(self):
        print("Vorname: ", self.prename)
        print("Minimale Schicht-Dauer: ", self.shift_min)
        print("Vertraglich vereinbarte h/Woche: ", self.vgA)
        if self.new == 1:
            print("Das erste Semester dabei!")
        else:
            print("Nicht das erste Semester dabei! (Schon eingearbeitet)")
        print("Präferenzen bzgl. Arbeitszeiten:")
        for day in dict.keys(self.support):
            print(day+":", [self.support[day][time] for time in dict.keys(self.support[day])])


# # input und preprocessing

# In[2]:


def raw_to_tutor(raw_tutor): # raw_tutor = pd.read_excel(f)
    temp = raw_tutor['Unnamed: 3']
    
    # Anfahrtszeit -> minimale Schichtdauer
    name = temp[5]
    if temp[9] == "bis zu einer Stunde":
        shift_min = 2
    else:
        shift_min = 3
        
    # vertraglich geregelte Arbeitszeit
    vgA = int(temp[10])
    
    # das erste Semester Teil des Teams
    if temp[11] == "stimmt":
        new = True
    else:
        new = False
    

    alltimes = ['10-11','11-12','12-13','13-14','14-15','15-16','16-17']
    days = ['Montag','Dienstag','Mittwoch','Donnerstag','Freitag']
    table = {}
    
    # unterstützte Zeitfenster + Label
    for count,day in enumerate(days):
        column = 'Unnamed: '+str(2*count+3)
        table[day] = {}
        for line0,time in enumerate(alltimes):
            table[day][time] = (raw_tutor[column][line0+16],raw_tutor['Unnamed: '+str(2*count+4)][line0+16])
        
    return tutor(name,shift_min,vgA,new,table)


# table: dictionary mit table[day] = dictionary mit table[day][time] = (präferenz,d/p-block)

# In[3]:


def inputFunction():
    
    current_directory = os.getcwd()
    os.chdir('\\'.join((current_directory,"Daten")))
    
    tutors = {}
    key = 1
    
    for filename in os.listdir(os.getcwd()):
        if 'Informationsblatt' in filename:
            with open(filename,'rb') as f:
                raw_tutor = pd.read_excel(f)
                tutors[key] = raw_to_tutor(raw_tutor)
                key+=1
   
            
    os.chdir(current_directory)
    
    return tutors


# ##### Es ist: 
# - 6 präsenz Stunden
# - 5 online Stunden
# an Montagen, Mittwochen, Freitagen: 
# - präsenz 2 fach besetzt
# - online einfach
# an Dienstagen und Donnerstagen:
# - präsenz 2 fach, 12-14 Uhr sogar dreifach
# - online einfach
# zusammen:
# - an Montagen, Mittwochen, Freitagen: 17h
# - an Dienstagen und Donnerstagen: 19h
# insgesamt: 89
# ##### Das Ergebnis ist valide!

# # all together: *main function*

# # input und preprocessing

# In[4]:


# Wie viele Lehrkräfte sollen pro Zeitfenster (in Präsenz) arbeiten?
def get_planned_h():

    current_directory = os.getcwd()
    os.chdir('\\'.join((current_directory,"Daten")))
    plan_df = pd.read_excel("Wochenplan.xlsx")
    os.chdir(current_directory)

    planned_h = {}

    days = []

    for (key,column) in plan_df.items():

        if key != "Unnamed: 0":
            if re.match("Unnamed",key)==None:
                day = key
                days.append(day)
                planned_h[day] = {}
                planned_h[day]["p"] = {}
                for (numero,time) in enumerate(plan_df["Unnamed: 0"][1:10]):
                    time = re.sub(r' Uhr','',time)
                    planned_h[day]["p"][time] = plan_df[day][numero+1]

            else:
                planned_h[day]["d"] = {}
                for (numero,time) in enumerate(plan_df["Unnamed: 0"][1:10]):
                    time = re.sub(r' Uhr','',time)
                    planned_h[day]["d"][time] = plan_df[key][numero+1]
        
    return [planned_h,days]

# planned_h ist ein dictionary mit: planned_h[Tag][Format][Zeit] = Anzahl an Lehrkräften, die arbeiten sollen


# In[5]:


# Welche Schichteinheiten (Aufgaben) gibt es, wie viele Lehrkräfte sollen jeweils arbeiten?
def get_demand(planned_h,parameters):
# planned_h
    
    days = parameters["days"]
    modes = parameters["modes"]
    times = parameters["times"]
    
    key = 1
    
    demand = {}
    total_demand = 0
    
    for day in days:
        for mod in modes:
            for time in times[mod]:
                demand[key] = (day,mod,time,planned_h[day][mod][time])
                total_demand += planned_h[day][mod][time] 
                key += 1
                
    return [demand,total_demand]
    
# demand[key] = (Tag,Format,Zeit,Anzahl an Lehrkräften)
# total_demand = Anzahl an Stunden, die insgesamt pro Woche geleistet werden müssen 
# das heißt: total_demand ist nicht gleich der Eingabegröße n = Anzahl der Aufgaben (demand)


# In[6]:


def set_parameters(tutors):
    
    parameters = {}
    
    modes = {"d","p"}
    parameters["modes"] = modes
    times = {
        "p": ["10-11","11-12","12-13","13-14","14-15","15-16"],
        "d": ["12-13","13-14","14-15","15-16","16-17"]
    }
    parameters["times"] = times
    
    ####
    # Wie viele Lehrkräfte sollen pro Zeitfenster (in Präsenz) arbeiten?
    [planned_h,days] = get_planned_h()
    parameters["days"] = days
    # Wie soll der Schichtplan gemacht werden, welche Schichteinheiten (Aufgaben) gibt es?
    [demand,total_demand] = get_demand(planned_h,parameters)
    # Input-Größe n = total_demand
    parameters["total_demand"] = total_demand
    ###              
                  
    # Präferenzen und E-Label merken
    alltimes = sorted(list(set(times["p"]).union(set(times["d"]))))
    parameters["alltimes"] = alltimes
    
    p = {} # Präferenzen / Profitkoeffizienten pro Zeitfenster
    E = {} # E-Label pro Zeitfenster
    for i in tutors.keys():
        p[i] = {}
        E[i] = {}
        for day in days:
            for time in alltimes:
                val = str(tutors[i].support[day][time][0])
                if val[0]=='E':
                    p[i][(day,time)] = int(val[1])
                    E[i][(day,time)] = True
                else:
                    p[i][(day,time)] = int(val)
                    E[i][(day,time)] = False
         
    parameters["p"] = p
    parameters["E"] = E
    # E, profit done
    
    # welche Nebenbedingungen sollen "streng" gelten, und welche in die Zielfunktion?
    # standardmäßig alle "streng", optional kann ausgewählt werden
    strict_constraints=[1,2,3,4,5,6,7]
    objective_constraints = []
    choose = 0 # wenn choose = 1 kann gewählt werden
    if choose == 1:
        strict_constraints = []
        print("Welche NB sollen streng sein? Nacheinander NB eingeben und mit 0 abschließen.")
        nb = 1
        while nb != 0:
            nb = int(input("NB: "))
            strict_constraints.append(nb)
        print("ok, ", [i for i in strict_constraints if i!=0])
        print("Welche NB soll in der Zielfunktion berücksichtigt sein?")
        objective_constraints = []
        nb = 1
        while nb != 0:
            nb = int(input("NB: "))
            objective_constraints.append(nb)
        print("ok, ", [i for i in objective_constraints if i!=0])
        parameters["objective_constraints"] = objective_constraints
    
    parameters["strict_constraints"] = strict_constraints
    parameters["objective_constraints"] = objective_constraints
    
    return [parameters,demand]


# # Modell bauen

# In[7]:


def make_constraints(tutors,demand,parameters,model,formulation):
    
    
    days = parameters["days"]
    times = parameters["times"]
    modes = parameters["modes"]
    alltimes = parameters["alltimes"]
    E = parameters["E"]
    strict_constraints = parameters["strict_constraints"]
    objective_constraints = parameters["objective_constraints"]
    
    # Hilfsparameter
    similar_demand = {}
    for day in parameters["days"]:
        similar_demand[day] = {}
        for mode in parameters["modes"]:
            similar_demand[day][mode] = {}
            for time in parameters["times"][mode]:
                nec = [dem[3] for dem in demand.values() if (dem[0]==day and (dem[1] == mode and dem[2] == time))]
                similar_demand[day][mode][time] = nec[0]
    
    
# Variablen

    # MW-FD-1
    x = { i: { (dem[0],dem[1],dem[2]): model.binary_var(name="x_tutor:{0}_day:{1}_mode:{2}_time:{3}_demandindex{4}".format(i,dem[0],dem[1],dem[2],j)) for (j,dem) in demand.items() } for i in tutors.keys() }
    
    # MW-FD-R4
    if 4 in strict_constraints or 4 in objective_constraints:
        y4 = { i: { day: model.binary_var(name="y_NB4_tutor:{0}_day:{1}".format(i,day)) for day in days} for i in tutors.keys() if tutors[i].vgA <= 13}
    
    if formulation == "x":
        
        # MW-FD-X-R3    
        if 3 in strict_constraints or 3 in objective_constraints:  
            y3 = { i: { (day,"p",times["p"][t]): model.binary_var(name="y3_tutor:{0}_day:{1}_mode:p_time:{2}".format(i,day,times["p"][t]))  for day in days for t in range(len(times["p"])-1) } for i in tutors.keys() }
        
        # MW-FD-X-R6-2
        if 6 in strict_constraints or 6 in objective_constraints:
            y6 = { i: {(day,mode): model.binary_var(name="y6_tutor:{0}_day:{1}_mode:{2}".format(i,day,mode)) for day in days for mode in modes } for i in tutors.keys() if tutors[i].shift_min == 3 }
    
    elif formulation == "z":
        
        # MW-FD-Z
        if (3 in strict_constraints or 3 in objective_constraints) or (6 in strict_constraints or 6 in objective_constraints): 
            z = { i: { (dem[0],dem[1],dem[2]): model.binary_var(name="z_tutor:{0}_day:{1}_mode:{2}_time:{3}".format(i,dem[0],dem[1],dem[2])) for (j,dem) in demand.items() } for i in tutors.keys() }#if not(dem[1]=="p" and dem[2] != "15-16") } for i in tutors.keys() }
        
        
# Einfache Nebenbedingungen

# Rucksackbedingungen   
    # MW-FC-1
    for i in tutors.keys():
        model.add_constraint( sum([x[i][(day,mode,time)] for (day,mode,time,b) in demand.values()]) <= tutors[i].vgA )

# Mindestauslastung der Lehrkräfte
    # MW-FC-2
    for i in tutors.keys():
        model.add_constraint( sum([x[i][(day,mode,time)] for day in days for mode in modes for time in times[mode] ]) >= (2/3)*tutors[i].vgA )
        
# p/d-Label blocks
    # MW-FC-3
    for i in tutors.keys():
        for day in days:
            for time in alltimes:
                if tutors[i].support[day][time][0]==0:
                    if time in times["p"]:
                        model.add_constraint(x[i][(day,"p",time)]==0)
                    if time in times["d"]:
                        model.add_constraint(x[i][(day,"d",time)]==0)
                elif tutors[i].support[day][time][1]=="p" and time in times["d"]:    
                    model.add_constraint(x[i][(day,"d",time)]==0)
                elif tutors[i].support[day][time][1]=="d" and time in times["p"]:
                    model.add_constraint(x[i][(day,"p",time)]==0)             
        
# modifizierte Zuordnungsbedingungen
    # MW-FC-4
    for day in days:
        # präsenz
        for t in range(len(times["p"])):
            model.add_constraint( sum([x[i][(day,"p",times["p"][t])] for i in tutors.keys() ]) == similar_demand[day]["p"][times["p"][t]] )
        # online
        for t in range(len(times["d"])):
            model.add_constraint( sum([x[i][(day,"d",times["d"][t])] for i in tutors.keys()]) == similar_demand[day]["d"][times["d"][t]])
            
# nicht gleichzeitig arbeiten
    # MW-FC-5
    for i in tutors.keys():
        if tutors[i].shift_min == 3:
            for day in days:
                for t in range(2,len(times["p"])):
                    model.add_constraint( x[i][(day,"p",times["p"][t])] + x[i][(day,"d",times["d"][t-2])] <= 1)

            
# Nebenbedingungen für die Hilfsvariablen
        
    # MW-FC-DR4 
    if 4 in strict_constraints or 4 in objective_constraints:                                                                                                                     
        for i in tutors.keys():
            if tutors[i].vgA <= 13:
                for day in days:
                    # (D.4.1)
                    model.add_constraint( (1/11)*(sum([x[i][(day,"p",time)] for time in times["p"] ]) + sum([x[i][(day,"d",time)] for time in times["d"] ]) ) <= y4[i][day] ) 
                    # (D.4.2)
                    model.add_constraint( (1/11)*(sum([x[i][(day,"p",time)] for time in times["p"] ]) + sum([x[i][(day,"d",time)] for time in times["d"] ]) + 10) >= y4[i][day] )     
         
    
    if formulation == "x":
        
        # MW-FD-X-R3
        if 3 in strict_constraints or 3 in objective_constraints:  
            for i in tutors.keys():
                for day in days:
                    for t in range(1,len(times["p"])-1):
                        # (D.3.1)
                        model.add_constraint( y3[i][(day,"p",times["p"][t])] <= x[i][(day,"p",times["p"][t])] + x[i][(day,"p",times["p"][t+1])] )
                        # (D.3.2)
                        model.add_constraint( y3[i][(day,"p",times["p"][t])] >= x[i][(day,"p",times["p"][t])] )
                        # (D.3.3)
                        model.add_constraint( y3[i][(day,"p",times["p"][t])] >= x[i][(day,"p",times["p"][t+1])] )
                        
        # MW-FD-X-R6-2
        if 6 in strict_constraints or 6 in objective_constraints:
            for i in tutors.keys():
                if tutors[i].shift_min == 3:
                    for day in days:
                        # (D.6.1)
                        model.add_constraint( sum([ x[i][(day,"p",times["p"][t])] for t in range(len(times["p"])) ]) <= 6*(1-y6[i][(day,"d")]) )                                                                                                                             
                        # (D.6.2)
                        model.add_constraint( sum([ x[i][(day,"d",times["d"][t])] for t in range(len(times["d"])-1) ]) <= 6*(1-y6[i][(day,"p")]) )     

    elif formulation == "z" and ((3 in strict_constraints or 3 in objective_constraints) or (6 in strict_constraints or 6 in objective_constraints)): 
        # MW-FD-Z
        for i in tutors.keys():
            for day in days:
                for t in range(1,len(times["p"])-1):
                    # (D.z.1) präsenz
                    model.add_constraint( z[i][(day,"p",times["p"][t])] + x[i][(day,"p",times["p"][t-1])] <= 1 ) 
                    # (D.z.2) präsenz   nicht für Spezialfall
                    model.add_constraint( x[i][(day,"p",times["p"][t])] - x[i][(day,"p",times["p"][t-1])] <= z[i][(day,"p",times["p"][t])] ) 
                for t in range(1,len(times["d"])-2):
                    # (D.z.1) digital
                    model.add_constraint( z[i][(day,"d",times["d"][t])] + x[i][(day,"d",times["d"][t-1])] <= 1 ) 
                    # (D.z.2) digital   nicht für Spezialfall
                    model.add_constraint( x[i][(day,"d",times["d"][t])] - x[i][(day,"d",times["d"][t-1])] <= z[i][(day,"d",times["d"][t])] ) 
                # (D.z.5)
                model.add_constraint( z[i][(day,"d",times["d"][len(times["d"])-1])] - x[i][(day,"d",times["d"][len(times["d"])-2])] - x[i][(day,"p",times["p"][len(times["p"])-1])] <= z[i][(day,"d",times["d"][len(times["d"])-1])] )
                for mode in modes: 
                    # (D.z.3)
                    model.add_constraint( z[i][(day,mode,times[mode][0])]==x[i][(day,mode,times[mode][0])] )   
                # (D.z.4)
                model.add_constraint( x[i][(day,"d",times["d"][len(times["d"])-1])] - x[i][(day,"d",times["d"][len(times["d"])-2])]                                      - x[i][(day,"p",times["p"][len(times["p"])-1])] <= z[i][(day,"d",times["d"][len(times["d"])-1])] )    
                # (D.z.5)
                # model.add_constraint( z[i][(day,"d",times["d"][len(times["d"])-1])] + x[i][(day,"d",times["d"][len(times["d"])-2])] \
                #                     - x[i][(day,"p",times["p"][len(times["p"])-1])] <= 1 )    
                
            
            
# 7 Regeln


    # Regel 1
    if 1 in strict_constraints:
        for i in tutors.keys():
            for day in days:
                mode = "p"
                
                # MW-FC-R1-1
                model.add_constraint( sum( [x[i][(day,"p",t)] for t in times["p"][0:5] ] ) <= 4 )
                model.add_constraint( sum( [x[i][(day,"p",t)] for t in times["p"][1:6] ] ) <= 4 )
                model.add_constraint( sum( [x[i][(day,"p",t)] for t in times["p"][1:6] ] ) + x[i][(day,"d",times["d"][4])] <= 4 )
                
                # MW-FC-R1-2
                model.add_constraint( sum( [x[i][(day,"p",t)] for t in times["p"] ] ) + x[i][(day,"d",times["d"][4])]  <= 5 )
                
                
    # Regel 2
    if 2 in strict_constraints:
        for i in tutors.keys():
            for day in days:
                mode = "p"
                
                # MW-FC-R2-1
                # (2.1.1)
                for t in range(1,len(times[mode])-1):
                    if not E[i][(day,times[mode][t])]==True:
                        model.add_constraint( x[i][(day,mode,times[mode][t])] - x[i][(day,mode,times[mode][t-1])] - x[i][(day,mode,times[mode][t+1])] <= 0 )
                # (2.1.2)
                if not E[i][(day,times[mode][0])]==True:
                    model.add_constraint( x[i][(day,mode,times[mode][0])] - x[i][(day,mode,times[mode][1])] <= 0 )
                # (2.1.3)
                if not E[i][(day,times[mode][5])]==True:
                    model.add_constraint( x[i][(day,mode,times[mode][5])] - x[i][(day,mode,times[mode][4])] - x[i][(day,"d",times["d"][4])] <= 0 )
                
                # MW-FC-R2-2
                if tutors[i].shift_min==3:
                    # (2.2.1)
                    for t in range(1,len(times[mode])-2):
                        model.add_constraint( 5*(x[i][(day,mode,times[mode][t])] + x[i][(day,mode,times[mode][t+1])]                                             - x[i][(day,mode,times[mode][t-1])] - x[i][(day,mode,times[mode][t+2])])                                              <= 5 + sum([x[i][(day,mode,times[mode][c])] for c in range(len(times[mode]))]) +x[i][(day,"d",times["d"][0])] )
                    # (2.2.2)
                    model.add_constraint( x[i][(day,mode,times[mode][0])] + x[i][(day,mode,times[mode][1])]                                         -x[i][(day,mode,times[mode][2])] <= 1 )
                    # (2.2.3)
                    model.add_constraint( x[i][(day,mode,times[mode][5])] + x[i][(day,"d",times["d"][4])]                                         -x[i][(day,mode,times[mode][4])] <= 1 )
                    # (2.2.4)
                    model.add_constraint( x[i][(day,mode,times[mode][4])] + x[i][(day,mode,times[mode][5])]                                         - x[i][(day,"d",times["d"][4])]-x[i][(day,mode,times[mode][3])] <= 1 )
                    
                    
    # Regel 3
    if 3 in strict_constraints:
        
        # MW-FC-X-R3                  
        if formulation == "x":
            for day in days:
                for t in range(len(times["p"])-1):
                    a = similar_demand[day]["p"][times["p"][t]] + similar_demand[day]["p"][times["p"][t+1]]-1
                    model.add_constraint( sum( [y3[i][(day,"p",times["p"][t])] for i in tutors.keys() ] ) <= a )    
        
        # MW-FC-Z-R3
        elif formulation == "z":
            for day in days:
                # MW-FC-Z-R3
                for t in range(1,len(times["p"])):
                    a = similar_demand[day]["p"][times["p"][t]]-1
                    model.add_constraint( sum( [z[i][(day,"p",times["p"][t])] for i in tutors.keys() ] ) <= a )  
                           
                                 
    # Regel 4
    if 4 in strict_constraints:
        
        # MW-FC-R4
        for i in tutors.keys():
            if tutors[i].vgA <= 8:
                model.add_constraint( sum([y4[i][day] for day in days ]) <= 3 )
            elif tutors[i].vgA <= 13:
                model.add_constraint( sum([y4[i][day] for day in days ]) <= 4 )  
                                 
    
    # Regel 5
    if 5 in strict_constraints:
        
        # MW-FC-R5
        for i in tutors.keys():
            for day in days:
                # (5.1.1)   
                # präsenz
                if not(E[i][(day,times["p"][1])]==True and E[i][(day,times["p"][5])]==True):
                    model.add_constraint(x[i][(day,"p",times["p"][1])] - x[i][(day,"p",times["p"][2])] - x[i][(day,"p",times["p"][3])]                                           - x[i][(day,"p",times["p"][4])] + x[i][(day,"p",times["p"][5])] <= 1)
                # digital
                if not(E[i][(day,times["d"][0])]==True and E[i][(day,times["d"][4])]==True):
                    model.add_constraint(x[i][(day,"d",times["d"][0])] - x[i][(day,"d",times["d"][1])] - x[i][(day,"d",times["d"][2])]                                           - x[i][(day,"d",times["d"][3])] + x[i][(day,"d",times["d"][4])] <= 1) 
                # (5.1.2)
                if tutors[i].shift_min==3:
                    if not(E[i][(day,times["p"][2])]==True and E[i][(day,times["d"][4])]==True):
                        model.add_constraint(x[i][(day,"p",times["p"][2])] - x[i][(day,"p",times["p"][3])] - x[i][(day,"p",times["p"][4])]                                               - x[i][(day,"p",times["p"][5])] + x[i][(day,"d",times["d"][4])] <= 1)
                # (5.2.1)
                # präsenz
                for t in range(1,len(times["p"])-3):
                    if not(E[i][(day,times["p"][t])]==True and E[i][(day,times["p"][t+3])]==True):
                          model.add_constraint(x[i][(day,"p",times["p"][t])] - x[i][(day,"p",times["p"][t+1])]                                           - x[i][(day,"p",times["p"][t+2])] + x[i][(day,"p",times["p"][t+3])] <= 1)
                # digital
                for t in range(0,len(times["d"])-3):
                    if not(E[i][(day,times["d"][t])]==True and E[i][(day,times["d"][t+3])]==True):
                          model.add_constraint(x[i][(day,"d",times["d"][t])] - x[i][(day,"d",times["d"][t+1])]                                           - x[i][(day,"d",times["d"][t+2])] + x[i][(day,"d",times["d"][t+3])] <= 1)
                # (5.2.2)
                if tutors[i].shift_min==3:
                    if not(E[i][(day,times["p"][3])]==True and E[i][(day,times["d"][4])]==True):
                        model.add_constraint(x[i][(day,"p",times["p"][3])] - x[i][(day,"p",times["p"][4])]                                               - x[i][(day,"p",times["p"][5])] + x[i][(day,"d",times["d"][4])] <= 1)                 
                # (5.3.1)
                # präsenz
                for t in range(1,len(times["p"])-2):
                    if not(E[i][(day,times["p"][t])]==True and E[i][(day,times["p"][t+2])]==True):
                        model.add_constraint(5*x[i][(day,"p",times["p"][t])] - 5*x[i][(day,"p",times["p"][t+1])] + 5*x[i][(day,"p",times["p"][t+2])]                                              <= 5 + sum([x[i][(day,"p",time)] for time in times["p"]]) + x[i][(day,"d",times["d"][len(times["d"])-1])] )
                # digital
                for t in range(len(times["d"])-2):
                    if not(E[i][(day,times["d"][t])]==True and E[i][(day,times["d"][t+2])]==True):
                        model.add_constraint(5*x[i][(day,"d",times["d"][t])] - 5*x[i][(day,"d",times["d"][t+1])] + 5*x[i][(day,"d",times["d"][t+2])]                                              <= 5 + sum([x[i][(day,"p",time)] for time in times["p"]]) + x[i][(day,"d",times["d"][len(times["d"])-1])] )
                # (5.3.2)
                if not(E[i][(day,times["p"][4])]==True and E[i][(day,times["d"][len(times["d"])-1])]==True):
                    model.add_constraint(5*x[i][(day,"p",times[mode][4])] - 5*x[i][(day,"d",times[mode][5])] + 5*x[i][(day,"d",times[mode][4])]                                              <= 5 + sum([x[i][(day,"p",time)] for time in times["p"]]) + x[i][(day,"d",times["d"][len(times["d"])-1])] )      
                
                                 
    # Regel 6
    if 6 in strict_constraints:    
        for i in tutors.keys():
            for day in days:
                
                # MW-FC-R6-1
                if tutors[i].shift_min == 2:
                    # (6.1.1)
                    for t in range(len(times["p"])-1):
                        model.add_constraint( 5*x[i][(day,"p",times["p"][t])] + sum([x[i][(day,"d",times["d"][T])] for T in range(max(0,(t-2)-2),min((t-2)+2,len(times["d"]))) ]) <= 5 )
                    model.add_constraint( 5*x[i][(day,"p",times["p"][t])] + sum([x[i][(day,"d",times["d"][T])] for T in range(1,len(times["d"])-1) ]) <= 5 )
                
                if tutors[i].shift_min == 3:
                    if formulation == "x":
                        # MW-FC-X-R6-2
                        model.add_constraint( y6[i][(day,"p")] + y6[i][(day,"d")] >= 1 )
                    elif formulation == "z":
                        # MW-FC-Z-R6-2
                        model.add_constraint( sum([z[i][(day,"p",times["p"][t])] for t in range(len(times["p"]))])                                              + 2*sum([z[i][(day,"d",times["d"][t])] for t in range(len(times["d"]))])                                              <= 2 )
    
    
    # Regel 7
    if 7 in strict_constraints:
        
        # MW-FC-R7
        for t in range(len(times["p"])):
            model.add_constraint( sum([x[i][(day,"p",times["p"][t])] for i in tutors.keys() if tutors[i].new == True ]) >= 1 )                                                                                                                          
           


    if formulation == "x":    
        return [model,x,y4,[y3,y6]]
    elif formulation == "z":
        return [model,x,y4,z]


# In[8]:


# hier kann das Projekt ggf. erweitert werden
# für NB in der Zielfunktion werden die anderen Variablen und Parameter gebraucht
def make_objective(tutors,demand,model,formulation,parameters,x,y4,var):#y3,y6):
    
    if formulation == "x":
        y3 = var[0]
        y6 = var[1]
    if formulation == "z":
        z = var
    
    p = parameters["p"]
    model.maximize( sum([p[i][(day,time)]*x[i][(day,mode,time)]                          for i in tutors.keys() for (day,mode,time,h) in demand.values()]) )
    
    return model


# In[9]:


def build_model(tutors,demand,parameters,formulation):
    
    model = Model(name="MW")
    [model,x,y4,var] = make_constraints(tutors,demand,parameters,model,formulation)
    model = make_objective(tutors,demand,model,formulation,parameters,x,y4,var)
    
    return model


# # Lösung als Datei ausgeben

# In[10]:


def solution_to_dataframe(solution,parameters,tutors,demand):
    # create data frame
    days = parameters["days"]
    modes = parameters["modes"]
    times = parameters["times"]
    
    # get the relevant x-variables
    itersol = set()
    for dvar in solution.iter_variables():
        if dvar.to_string()[0] == "x":
            itersol.add(dvar)

    # assign names from variables to demand
    table = { Tag: { mode: { time: [] for time in set(times["p"]).union(set(times["d"])) } for mode in modes } for Tag in days }
    for dvar in itersol:
        varstr = dvar.to_string()
        for (Tag,mode,time,b) in demand.values():
            if ''.join([Tag,"_mode:",mode,"_time:",time]) in varstr:
                key = int(varstr[8])
                prename = tutors[key].prename
                table[Tag][mode][time].append(prename)

    # get information into dataframe Table        
    Table = {
        'Tag': ["Montag",None,"Dienstag",None,"Mittwoch",None,"Donnerstag",None,"Freitag",None],
        'Modus': ["präsenz","online","präsenz","online","präsenz","online","präsenz","online","präsenz","online"]    
    }
    Table = { 
        'Zeiten': [time for time in sorted(list(set(times["p"]).union(set(times["d"]))))]
            }

    for Tag in days:
        Table[Tag] = [ '\n'.join([''.join(["präsenz: ",' ,'.join(table[Tag]["p"][time])]),''.join(["online: ",' ,'.join(table[Tag]["d"][time])])]) for time in sorted(list(set(times["p"]).union(set(times["d"])))) ]


    # create the data frame
    df = pd.DataFrame(Table)

    # create excel writer object
    #writer = pd.ExcelWriter("output.xlsx")

    # write dataframe to excel
    #df.to_excel(writer)

    # save the document
    #writer.save()
    
    return [df,table]


# In[11]:


def dataframe_to_document(df,table,parameters):
    days = parameters["days"]
    alltimes = parameters["alltimes"]
    import xlwt 
    from xlwt import Workbook
            
    wbk = Workbook()
    sheet = wbk.add_sheet('Sheet 1')

    for (i,day) in enumerate(days):
        # Überschriften
        sheet.write(0, 1+2*i, day)
        sheet.write(1, 1+2*i, "präsenz")
        sheet.write(1, 2+2*i, "online")
        for (line,time) in enumerate(alltimes):
            # Uhrzeiten

            # Besetzung
            sheet.write(2+line, 1+2*i, ' ,'.join(table[day]["p"][time]))
            sheet.write(2+line, 2+2*i, ' ,'.join(table[day]["d"][time]))

    for (line,time) in enumerate(alltimes):
        sheet.write(2+line,0, time) 

    sheet.col(0).width = 6*256
    for i in range(5):
        sheet.col(1+2*i).width = 20*256
        sheet.col(2+2*i).width = 11*256


    current_directory = os.getcwd()
    if not "output" in os.getcwd(): #os.path.exists(current_directory+"//output"):
        #print("check")
        os.chdir(current_directory+"\\Daten\\Schichtplan Ausgabe")
    wbk.save("Schichtplan.odt")
    os.chdir(current_directory)
    
    print("Es wurde erfolgreich ein Schichtplan erstellt.")

    return
    


# # Programm für den Echteinsatz (mit Ausgabe eines Schichtplans)

# In[12]:


# für den Normalbetrieb / Exceltabellen als Input, keine künstlichen Agenten
def real_case(formulation):
    tutors = inputFunction()
    
    print("Anzahl an Informationsblättern: ",len(tutors))
    print("Insgesamt verfügbare Stunden: ", sum([tutors[i].vgA for i in tutors.keys()]))
    
    [parameters,demand] = set_parameters(tutors)
    
    print("Insgesamt benötigte Stunden: ", parameters["total_demand"])

    model = build_model(tutors,demand,parameters,formulation)

    model.print_information()
    
    
    if model.solve() == None:
        print("*** Das Problem ist nicht lösbar")
    else:
        solution = model.solution

        [df,table] = solution_to_dataframe(solution,parameters,tutors,demand)
        
        dataframe_to_document(df,table,parameters)


# In[13]:


solution = real_case("z")


# # Simulation und Numerische Tests

# # Datengrundlage für Simulation

# In[14]:


# beziehe aus den Informationsblättern Tagesabläufe (Präferenzen pro Tag) 
def sample_day_tables(tutors):
    
    alltimes = ['10-11','11-12','12-13','13-14','14-15','15-16','16-17']
    days = ['Montag','Dienstag','Mittwoch','Donnerstag','Freitag']
    
    str_day_tables = set()
    count = 0
    for tutor in tutors.values():
        for day in days:
            str_day_tables.add(json.dumps(tutor.support[day],sort_keys=True))
            count += 1
    
    sample_days = []
    for string in list(str_day_tables):
        sample_days.append(json.loads(string))
    
    return sample_days


# # simuliere Lehrkräfte 

# In[15]:


# bei vorgegebenen h/Woche (vertraglich geregelter Arbeitszeit) <- isolated_h_per_tutor = True
def create_tutor(h_week,sampled_days):
# h_week: vertraglich geregelte Arbeitszeit
# sampled_days: Menge der verschiedenen Tagesverläufe 
    
    prename = "agent"
    if random.randint(0,1) == 0:
        shift_min = 2
    else:
        shift_min = 3
    new = random.randint(0,1)
    
    # Präferenzen zu Zeitfenstern des Schichtplans müssen passen
    # Die verfügbaren Wochenstunden (, die nicht die Präferenz 0 haben,) müssen die vertraglich geregelten übersteigen.

    # zeros_per_week gibt an, zu wie vielen Zeitfenstern die Präferenz 0 gegeben ist
    zeros_per_week = h_week+6
    
    while zeros_per_week > h_week+5:
        zeros_per_week = 0
        support = {}
        for day in ['Montag','Dienstag','Mittwoch','Donnerstag','Freitag']:
            support[day] = sampled_days[random.randint(0,len(sampled_days)-1)]
            for preference in support[day].values():
                if preference == 0:
                    zeros_per_week += 1
            
            
    return tutor(prename,shift_min,h_week,new,support)


# In[16]:


# so viele Lehrkräfte, sodass 10% mehr Arbeitszeit pro Woche verfügbar ist als nötig
def create_tutors(h,sampled_days,total_demand,isolate_h_per_tutor):
# h_week: vertraglich geregelte Arbeitszeit
# sampled_days: Menge der verschiedenen Tagesverläufe 
# total demand = n (Wie viele Schichteinheiten gibt es?)

    tutors = {}
    
    if isolate_h_per_tutor == True:
        # Dann ist schon vorher klar, wie viele Lehrkräfte gebraucht werden -> for Schleife
        for nr in range(0,int((total_demand*1.1)/h)):
            tutors[nr+1] = create_tutor(h,sampled_days)
            tutors[nr+1].prename = tutors[nr+1].prename + str(nr+1)
        
    else:
        # Hier ist nicht klar, wie viele Lehrkräfte gebraucht werden -> while Schleife
        total_supply = 0 # verfügbare Arbeitszeit pro Woche
        nr = 1
        while total_supply < int(total_demand*1.1):
            h_week = random.randint(3,19)
            tutors[nr] = create_tutor(h_week,sampled_days)
            tutors[nr].prename = tutors[nr].prename + str(nr)
            total_supply += h_week
            nr += 1
        
    return tutors


# # Probleme simulieren, lösen und numerische Daten erheben

# In[17]:


def collect_numerical_data(sampled_days,counter,isolate_h_per_tutor,formulation):
    
    # Vorbereitung
    
    input_table = {} # Ergebnisse werden gespeichert
    
    parameters = {}

    modes = {"d","p"}
    parameters["modes"] = modes

    times = {
        "p": ["10-11","11-12","12-13","13-14","14-15","15-16"],
        "d": ["12-13","13-14","14-15","15-16","16-17"]
    }
    parameters["times"] = times

    # Wie viele Lehrkräfte pro Zeitfenster? -> planned_h
    [planned_h,days] = get_planned_h()
    parameters["days"] = days

    [demand,total_demand] = get_demand(planned_h,parameters)
    #####################################################
    
    # Welchen Einfluss hat die wöchentliche Arbeitszeit der Lehrkräfte auf die Komplexität und Zulässigkeit?
    if isolate_h_per_tutor == True:

        for i in range(3,20):
            input_table[i] = {
                "Bedarf": 0, # nur vorbereitend
                "m": 0, #
                "Modellierungszeit": {},
                "Lösungszeit": {},
                "Zeit insgesamt": {},
                "Lösbarkeit": {},
                "vgA/verfügbare h": {},
                "Errors": {},
                "Lösungszeit zulässiger Problme": {},
                "Modellierungszeit zulässiger Problme": {}
            }

        start = t.time() # 
        
        # Versuchsgröße
        for k in tqdm(range(0,counter)):
            
            # vertraglich geregelte Arbeitszeit pro Woche
            for h in range(3,20):
                
                tutors = create_tutors(h,sampled_days,total_demand,isolate_h_per_tutor)
                [parameters,demand] = set_parameters(tutors) # update Parameter -> Labels, ...

                # speichere erste Daten
                input_table[h]["Bedarf"] = total_demand
                input_table[h]["m"] = int((total_demand*1.1)/h)

                
                # manchmal gibt es "trivially unfeasible constraints" -> crash
                # falls das Modellbauen zu einer Fehlermeldung führt, geht es mit except weiter
                try:
                    start_model = t.time()
                    model = build_model(tutors,demand,parameters,formulation)
                    end_model = t.time()

                    start_solving = t.time()
                    solution = model.solve()
                    end_solving = t.time()

                    if solution == None:
                        solvable = False
                    else:
                        solvable = True

                    # wie flexibel ist eine Lehrkraft hinsichtlich vertraglich geregelter Arbeitszeit und verfügbarer Arbeitszeit?
                    inverse_flexibility = 0
                    for i in tutors.keys(): 
                        supplied_h = 0
                        for day in tutors[i].support.keys():
                            for time in tutors[i].support[day].keys():
                                if tutors[i].support[day][time]!=0:
                                    supplied_h+=1
                        inverse_flexibility += (tutors[i].vgA/supplied_h)/len(tutors)

                    # speichere Daten
                    input_table[h]["Modellierungszeit"][k] = end_model-start_model
                    input_table[h]["Lösungszeit"][k] = end_solving-start_solving
                    input_table[h]["Zeit insgesamt"][k] = (end_model-start_model) + (end_solving-start_solving)
                    input_table[h]["Lösbarkeit"][k] = int(solvable)
                    input_table[h]["vgA/verfügbare h"][k] = inverse_flexibility

                except:
                    input_table[h]["Modellierungszeit"][k] = "error"
                    input_table[h]["Lösungszeit"][k] = "error"
                    input_table[h]["Zeit insgesamt"][k] = "error"
                    input_table[h]["Lösbarkeit"][k] = "error"
                    input_table[h]["vgA/verfügbare h"][k] = "error"
                    
    # Wie ist es mit zufälligen Eingaben, mit nur kontrolliertem Bedarf?
    # -> notwendige Arbeitszeit pro Woche wird gegeben, mit "zufällig" vielen Lehrkräften
    else: 
        # analog 
        input_table = {
                "Bedarf": 0,
                "m": {},
                "Modellierungszeit": {},
                "Lösungszeit": {},
                "Zeit insgesamt": {},
                "Lösbarkeit": {},
                "vgA/verfügbare h": {},
                "Errors": {},
                "durchschn. vgA": {},
                "Lösungszeit zulässiger Problme": {},
                "Modellierungszeit zulässiger Problme": {}
            }

        start = t.time()

        # Versuchsgröße
        for k in tqdm(range(0,counter)):
        
            # analog
            tutors = create_tutors(0,sampled_days,total_demand,isolate_h_per_tutor)
            [parameters,demand] = set_parameters(tutors)
            
            input_table["Bedarf"] = total_demand
            input_table["m"][k] = len(tutors)

            try:
                start_model = t.time()
                model = build_model(tutors,demand,parameters,formulation)
                end_model = t.time()

                start_solving = t.time()
                solution = model.solve()
                end_solving = t.time()

                if solution == None:
                    solvable = False
                else:
                    solvable = True

                inverse_flexibility = 0
                for i in tutors.keys(): 
                    supplied_h = 0
                    for day in tutors[i].support.keys():
                        for time in tutors[i].support[day].keys():
                            if tutors[i].support[day][time]!=0:
                                supplied_h+=1
                    inverse_flexibility += (tutors[i].vgA/supplied_h)/len(tutors)

                input_table["Modellierungszeit"][k] = end_model-start_model
                input_table["Lösungszeit"][k] = end_solving-start_solving
                input_table["Zeit insgesamt"][k] = (end_model-start_model) + (end_solving-start_solving)
                input_table["Lösbarkeit"][k] = int(solvable)
                input_table["vgA/verfügbare h"][k] = inverse_flexibility
                input_table["durchschn. vgA"][k] = sum([tutors[i].vgA for i in tutors.keys()])/len(tutors)

            except:
                input_table["Modellierungszeit"][k] = "error"
                input_table["Lösungszeit"][k] = "error"
                input_table["Zeit insgesamt"][k] = "error"
                input_table["Lösbarkeit"][k] = "error"
                input_table["vgA/verfügbare h"][k] = "error"
                input_table["durchschn. vgA"][k] = "error"
        
    end = t.time()

    
    return [input_table,end-start,total_demand]


# # Ausgabe von Dokumente

# In[18]:


def numerical_data_to_documents(input_table,counter,total_demand,isolate_h_per_tutor,formulation):

    cwd = os.getcwd()
    os.chdir('\\'.join([cwd,"Neue Auswertungen"]))
    if isolate_h_per_tutor == True:
    # input_table ist pro Arbeitszeit pro Woche -> (19-2) mal so viele Daten wie sonst
        for h in range(3,20):
            #input_table[h]["Mittelwert"][] 
            input_table[h]["Modellierungszeit"][counter+2] = sum([input_table[h]["Modellierungszeit"][k] for k in range(counter) if input_table[h]["Modellierungszeit"][k]!= "error"])/counter
            input_table[h]["Lösungszeit"][counter+2] = sum([input_table[h]["Lösungszeit"][k] for k in range(counter) if input_table[h]["Lösungszeit"][k]!= "error"])/counter
            input_table[h]["Zeit insgesamt"][counter+2] = sum([input_table[h]["Zeit insgesamt"][k] for k in range(counter) if input_table[h]["Zeit insgesamt"][k]!="error"])/counter
            input_table[h]["Lösbarkeit"][counter+2] = sum([input_table[h]["Lösbarkeit"][k] for k in range(counter) if input_table[h]["Lösbarkeit"][k]!= "error"])/counter
            input_table[h]["vgA/verfügbare h"][counter+2] = sum([input_table[h]["vgA/verfügbare h"][k] for k in range(counter) if input_table[h]["vgA/verfügbare h"][k]!= "error"])/counter
            input_table[h]["Errors"][counter+2] = len([input_table[h]["Zeit insgesamt"][k] for k in range(counter) if input_table[h]["Zeit insgesamt"][k]=="error"])
            if len([input_table[h]["Lösungszeit"][k] for k in range(counter) if (input_table[h]["Lösungszeit"][k]!= "error" and input_table[h]["Lösbarkeit"][k]!= 0)]) != 0:
                input_table[h]["Lösungszeit zulässiger Problme"][counter+2] = sum([input_table[h]["Lösungszeit"][k] for k in range(counter) if (input_table[h]["Lösungszeit"][k]!= "error" and input_table[h]["Lösbarkeit"][k]!= 0)])/len([input_table[h]["Lösungszeit"][k] for k in range(counter) if (input_table[h]["Lösungszeit"][k]!= "error" and input_table[h]["Lösbarkeit"][k]!= 0)])
            else:
                input_table[h]["Lösungszeit zulässiger Problme"][counter+2] = "-"
            if len([input_table[h]["Modellierungszeit"][k] for k in range(counter) if (input_table[h]["Modellierungszeit"][k]!= "error" and input_table[h]["Lösbarkeit"][k]!= 0)]) != 0:
                input_table[h]["Modellierungszeit zulässiger Problme"][counter+2] = sum([input_table[h]["Modellierungszeit"][k] for k in range(counter) if (input_table[h]["Modellierungszeit"][k]!= "error" and input_table[h]["Lösbarkeit"][k]!= 0)])/len([input_table[h]["Modellierungszeit"][k] for k in range(counter) if (input_table[h]["Modellierungszeit"][k]!= "error" and input_table[h]["Lösbarkeit"][k]!= 0)])
            else:
                input_table[h]["Modellierungszeit zulässiger Problme"][counter+2] = "-"
            
            for k in range(counter):
                if input_table[h]["Lösbarkeit"][k] == "error":
                    input_table[h]["Errors"][counter+2] += 1


            df = pd.DataFrame(input_table[h])

            last = df.index[-1]
            df = df.rename(index={last: 'Durchschnitt:'})
            # pro h eine Tabelle
            with pd.ExcelWriter("demand_{0} {1}_formulation numerical results h{2}.xlsx".format(total_demand,formulation,h)) as writer:
                df.to_excel(writer)
                writer.save()
        # -> eine Tabelle mit Durchschnittswerten
        numerical_overview(input_table[3].keys(),total_demand,counter,formulation)

    else: 
        input_table["Modellierungszeit"][counter+2] = sum([input_table["Modellierungszeit"][k] for k in range(counter) if input_table["Modellierungszeit"][k]!= "error"])/counter
        input_table["Lösungszeit"][counter+2] = sum([input_table["Lösungszeit"][k] for k in range(counter) if input_table["Lösungszeit"][k]!= "error"])/counter
        input_table["Zeit insgesamt"][counter+2] = sum([input_table["Zeit insgesamt"][k] for k in range(counter) if input_table["Zeit insgesamt"][k]!="error"])/counter
        input_table["Lösbarkeit"][counter+2] = sum([input_table["Lösbarkeit"][k] for k in range(counter) if input_table["Lösbarkeit"][k]!= "error"])/counter
        input_table["vgA/verfügbare h"][counter+2] = sum([input_table["vgA/verfügbare h"][k] for k in range(counter) if input_table["vgA/verfügbare h"][k]!= "error"])/counter
        input_table["Errors"][counter+2] = 0
        if len([input_table["Lösungszeit"][k] for k in range(counter) if (input_table["Lösungszeit"][k]!= "error" and input_table["Lösbarkeit"][k]!= 0)])!=0:
            input_table["Lösungszeit zulässiger Problme"][counter+2] = sum([input_table["Lösungszeit"][k] for k in range(counter) if (input_table["Lösungszeit"][k]!= "error" and input_table["Lösbarkeit"][k]!= 0)])/len([input_table["Lösungszeit"][k] for k in range(counter) if (input_table["Lösungszeit"][k]!= "error" and input_table["Lösbarkeit"][k]!= 0)])
        else:
            input_table["Lösungszeit zulässiger Problme"][counter+2] = "-"
        if len([input_table["Modellierungszeit"][k] for k in range(counter) if (input_table["Modellierungszeit"][k]!= "error" and input_table["Lösbarkeit"][k]!= 0)])!=0:
            input_table["Modellierungszeit zulässiger Problme"][counter+2] = sum([input_table["Modellierungszeit"][k] for k in range(counter) if (input_table["Modellierungszeit"][k]!= "error" and input_table["Lösbarkeit"][k]!= 0)])/len([input_table["Modellierungszeit"][k] for k in range(counter) if (input_table["Modellierungszeit"][k]!= "error" and input_table["Lösbarkeit"][k]!= 0)])
        else:
            input_table["Modellierungszeit zulässiger Problme"][counter+2] = "-"
        for k in range(counter):
            if input_table["Lösbarkeit"][k] == "error":
                input_table["Errors"][counter+2] += 1


        df = pd.DataFrame(input_table)

        last = df.index[-1]
        df = df.rename(index={last: 'Durchschnitt:'})

        with pd.ExcelWriter("demand_{0} {1}_formulation numerical results.xlsx".format(total_demand,formulation)) as writer:
            df.to_excel(writer)
            writer.save()

#dataframe = pd.DataFrame(input_table)
#dataframe.to_excel('numerical results.xlsx')


# In[19]:


# falls isolate_h_per_tutor=True
def numerical_overview(keys,total_demand,counter,formulation):

    results = {}

    results["h/Woche"] = {}
    #for key in input_table[3].keys():
    #    results[key] = {}
    for key in keys:
        results[key] = {}
        
    row = 1

    for h in range(3,20):
                
        r = pd.read_excel("demand_{0} {1}_formulation numerical results h{2}.xlsx".format(total_demand,formulation,h))
        durchschnitt = r.index[-1]
        results["h/Woche"][row] = h
        for key in r.keys():
            if key != "Unnamed: 0":
                results[key][row] = r[key][durchschnitt]
        row += 1

    row += 1
    results["h/Woche"][row] = counter
    results_df = pd.DataFrame(results)
    #last = df.index[-1]
    results_df = results_df.rename(index={row: 'Versuchsgröße:'})
    

    with pd.ExcelWriter("demand_{0} {1}_formulation numerical results overview.xlsx".format(total_demand,formulation)) as writer:
        #results_df = pd.DataFrame(results)
        results_df.to_excel(writer)
        writer.save()


# # Versuchsreihe

# In[20]:


# counter: Versuchsgröße (Wiederholungen)
# isolate_h_per_tutor: {
#    True: jede Lehrkraft (künstlicher Agent) hat die selbe vertraglich geregelte Arbeitszeit
#    False: Lehrkräfte (künstliche Agenten) haben zufällige Arbeitszeiten

def numerical_evaluation(counter,isolate_h_per_tutor,formulation):    
    
    cwd = os.getcwd()
    # beziehe aus den Informationsblättern Tage (verschiedene Präferenzen)
    sampled_days = sample_day_tables(inputFunction())
    
    # Versuche werden ausgeführt
    [input_table,taken_time,total_demand] = tqdm(collect_numerical_data(sampled_days,counter,isolate_h_per_tutor,formulation))
    
    # Ergebnisse werden zusammengefasst und in Output Dateien verwandelt
    numerical_data_to_documents(input_table,counter,total_demand,isolate_h_per_tutor,formulation)

    os.chdir(cwd)


# In[21]:


evaluation = input("Sollen Untersuchungen vorgenommen werden? (Nein/Teams nach Bedarf/Teams nach vgA) \n")
if evaluation != "Nein":
    if evaluation == "Teams nach Bedarf":
        isolate_h_per_tutor = False
        formulation = input("x- oder z-Formulierung? (x/z)")
        size = input("Versuchsgröße? (ganze Zahl größer 0)")
        numerical_evaluation(int(size),isolate_h_per_tutor,formulation)
    elif evaluation == "Teams nach vgA":
        isolate_h_per_tutor = False
        formulation = input("x- oder z-Formulierung? (x/z)")
        size = input("Versuchsgröße? (ganze Zahl größer 0)")
        numerical_evaluation(int(size),isolate_h_per_tutor,formulation)
    else:
        print("Unzulässige Eingabe")


# In[ ]:





# In[ ]:





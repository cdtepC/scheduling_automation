#!/usr/bin/env python
# coding: utf-8

# In[15]:


import sys
try:
    import docplex.mp
except:
    if hasattr(sys, 'real_prefix'):
        get_ipython().system('pip install docplex')
    else:
        get_ipython().system('pip install --user docplex')
from docplex.mp.model import Model

import pandas as pd
import regex as rg
import math
import numpy as np
import os

# analysis
import time 
from tqdm import tqdm
import random


class teacher:
    def __init__(self,key,h_per_week,row,nr):
        self.key = key # Schlüssel für die Datenstruktur über alle Lehrkräfte
        self.h_per_week = h_per_week # Soll-Wert
        self.row = row # Reihe in der Excel-Tabelle
        self.nr = nr
    def print(self):
        print("Nr.: ", self.key)
        print("Gewünschte h/Woche: ", self.h_per_week)
        print("Excel Reihe:", self.row)


# In[16]:


def inputFunction():
    
    current_directory = os.getcwd()
    
    demand = {} # dictionary über alle Nachfragen
    teachers = {} # dictionary über alle Lehrkräfte
    key = 1
    
    #os.chdir(''.join((current_directory,"\\Data")))
    raw_data = pd.read_excel('\\'.join((current_directory,"Data","Lehrkräfte.xlsx")))
    
    
    # import Lehrkräfte
    for row in range(1,len(raw_data)-1):
        h_per_week = raw_data["Stundenzahl Soll"][row]
        nr = raw_data["Nummer"]
        if h_per_week > 0:
            teachers[key] = teacher(row+1,h_per_week,row,nr)
            key += 1     
            
            
    key = 1
    # import Nachfrage
    raw_data = pd.read_excel('\\'.join((current_directory,"Data","Nachfrage.xlsx")))
    for row in range(len(raw_data)-2):
        student_nr = raw_data["Schüler"][row]
        if raw_data["Deutsch"][row]>0:
            demand[key] = (row+2,"Deutsch",raw_data["Deutsch"][row],student_nr)
            key += 1
        if raw_data["Mathe"][row]>0:
            demand[key] = (row+2,"Mathe",raw_data["Mathe"][row],student_nr)
            key += 1
        if raw_data["Englisch"][row]>0:
            demand[key] = (row+2,"Englisch",raw_data["Englisch"][row],student_nr)
            key += 1
    
    return [teachers,demand]


# In[17]:


def get_splines(teachers,n,L,K):

    a = {}
    b = {}
    t = {}
    gamma_s_0 = np.concatenate((np.zeros(n),[1,0],np.zeros(L)))
    gamma_d_0 = np.concatenate((np.zeros(n),[0,1],np.zeros(L)))
    
    for (i,tutor) in teachers.items():
        
        c_i = tutor.h_per_week
        
        # Diskretisierungsstellen in quadratischen Abständen: 
        # Bsp.: K=10, c_i = 100, t[i] = [0,1,4,9,16,25,36,49,64,81]
        t[i] = np.linspace(0, np.sqrt(c_i), num=int(K), endpoint=False)**2
        
        # Koeffizienten des Interpolanden auf t[i] <= t <= t[i+1]
        # a wird direkt auf tilde{a} für die Epigraphform gebracht, wobei 
        a[i] = [ (t[i][k+1]+t[i][k] )*(gamma_s_0 + gamma_d_0) for k in range(K-1) ]
        b[i] = [ (-t[i][k]*t[i][k+1]) for k in range(K-1) ]

    return[a,b]
    


# In[18]:


def build_model(teachers,demand,L,formulation,K,students):

    model = Model(name="LS")
    
    n = len(demand)
    m = len(teachers)
    w = { i: { j: demand[j][2] for j in demand.keys() } for i in teachers.keys() }
    
    # Variablen unabhängig von linear oder quadratisch
    # LS-FD-1
    x = { i: { j: model.binary_var(name='x_teacher:{0}_demand:{1}'.format(i,j)) for j in demand.keys() } for i in teachers.keys() }
    # LS-FD-2
    s = { i: model.continuous_var(name='s_{0}'.format(i)) for i in teachers.keys() }
    d = { i: model.continuous_var(name='d_{0}'.format(i)) for i in teachers.keys() }
    # LS-FD 3
    y = { i: { l: model.binary_var(name='y_{0}_{1}'.format(i,l)) for l in students.keys() } for i in teachers.keys() }

    # Zuordnungsbedingungen
    # LS-FC-1
    for j in demand.keys():
        model.add_constraint(sum(np.array([x[i][j] for i in teachers.keys()]))==1)

    # elastische Rucksackbedingungen
    # LS-FC-2
    for i in teachers.keys():
        model.add_constraint(sum(np.array([w[i][j]*x[i][j] for j in demand.keys() ])) + d[i] - s[i] == teachers[i].h_per_week)
        model.add_constraint(d[i]>= 0)
        model.add_constraint(s[i]>= 0)

    # Extremfälle
    # LS-FC-3
    for i in teachers.keys():
        model.add_constraint(s[i]<= teachers[i].h_per_week)

    # Schüler pro Lehrer
    for student_nr in students.keys():
        # LS-FC-D3
        for i in teachers.keys():
            model.add_constraint(y[i][student_nr] >= (1/n)*sum([x[i][j] for j in students[student_nr]])  )
            model.add_constraint(y[i][student_nr] <= sum([x[i][j] for j in students[student_nr]])  )
        # LS-FC-4
        model.add_constraint(sum([y[i][student_nr] for i in teachers.keys()])==1)


    if formulation == "linear":
        # Epigraphform 
        
        # LS-FE-D
        z = { i: model.continuous_var(name='z_{0}'.format(i)) for i in teachers.keys() } # bzw. t 
        
        # Parameter
        #K=11
        [a,b] = get_splines(teachers,n,L,K)
        # tilde{p}
        p_i = np.concatenate((np.ones(n),[0,0],np.zeros(L)))
        
        # LS-FE-C
        for i in teachers.keys():
            x_i_tilde = np.concatenate(([x[i][j] for j in demand.keys()],[s[i],d[i]],[y[i][student_nr] for student_nr in students.keys()]))
            for k in range(K-1):
                model.add_constraint( (a[i][k]-p_i).dot(x_i_tilde) + b[i][k] <= z[i] )  
                
        # LS-FE-G
        model.minimize(sum(z.values()))
        
    elif formulation == "quadratisch":
        # ohne Epigraphform
        
        # p und x
        p = np.ones(m*n)
        get_x = {}
        for i in teachers.keys():
            get_x[i] = np.array([x[i][j] for j in demand.keys()])
        x = np.concatenate([get_x[i] for i in teachers.keys()])
        
        model.maximize(p.dot(x) - sum([pow(d[i]+s[i],2) for i in teachers.keys()]) )

    return [model,x]


# In[19]:


def get_students(demand):
    
    students = { d[3]: [] for d in demand.values() }
    for (key,d) in demand.items():
        students[d[3]].append(key)

    return students


# # Skript 

# In[22]:


[teachers,demand] = inputFunction()
students = get_students(demand)
L = len(students)
########################## Umfang der Daten
print("Anzahl der Lehrkräfte: m =", len(teachers))
print("Anzahl der SchülerInnen: L= ", L)
print("Umfang der Nachfrage: n =", len(demand))
summe = 0
for dem in demand.values():
    summe += dem[2]
print("Summe nachgefragter Stunden: ",summe)
summe = 0
for t in teachers.values():
    summe += t.h_per_week
print("Summe der Soll-Werte: ", summe )
########################## 
n = len(demand.keys())
#[a,b] = get_splines(teachers,n,L,K)
## 
formulation = "linear"
K = 11
[model,x] = build_model(teachers,demand,L,formulation,K,students)

model.print_information()




start = time.time()
solution = model.solve()
end = time.time()

print(end-start)

if solution != None:
    print("Das Problem ist lösbar")
    
    for (i,t) in teachers.items():
        print("Lehrkraft Nr. {0} arbeitet {1} von {2} h/Woche".format(i,sum([demand[j][2]*x[i][j].solution_value for j in demand.keys()]),t.h_per_week)) 
        
proof = { l: set() for l in students.keys() } # lehrer pro Schüler
for l in students.keys():
    for j in students[l]:
        for i in x.keys():
            if x[i][j].solution_value == 1:
                #print("Nachfrage j={0} kommt von Schüler {1} und wird von Lehrer {2} unterrichtet".format(l,j,i))
                proof[l].add(i)
                
    print("Nachhilfeschüler {0} wird von den Lehrern {1} unterrichtet.".format(l,proof[l]))


# In[ ]:





# In[ ]:





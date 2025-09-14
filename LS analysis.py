#!/usr/bin/env python
# coding: utf-8

# In[12]:


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

import pandas as pd
import regex as rg
import math
import numpy as np
#pifrom ortools.linear_solver import pywraplp
import os

# analysis
import time 
from tqdm import tqdm
import random
import datetime


class teacher:
    def __init__(self,key,h_per_week,row,nr):
        self.key = key # Schlüssel für die Datenstruktur über alle Lehrkräfte
        self.h_per_week = h_per_week # Soll-Wert
        self.row = row # Reihe in der Excel-Tabelle
        self.nr = nr
    def print(self):
        print("Nr.: ", self.key)
        print("Gewünschte h/Woche: ", self.h_per_week)
        print("original row:", self.row)


# In[13]:


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
    


# In[14]:


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

    return model


# In[15]:


def get_students(demand):
    
    students = { d[3]: [] for d in demand.values() }
    for (key,d) in demand.items():
        students[d[3]].append(key)

    return students


# In[16]:


def random_demand_by_n(n):
    
    objectives = ["Deutsch","Mathe","Englisch"]
    demand = {}
    #obj = 0
    #l = 0
    stud_keys = list(range(1,n+1))
    studs = { l: ["Deutsch","Mathe","Englisch"] for l in stud_keys }
    
    for j in range(1,n+1):
        if len(demand)==0:
            obj = objectives[random.randint(0,2)]
            l = random.randint(1,n+1)
            # d = (Excel Tabellen Verweis, Fach, h/Woche, Schüler Schlüssel)
            d = (0,obj,random.randint(1,3),l)  
        else:
            #while (obj,l) in [(objective,student) for (a,objective,b,student) in demand.values()]:
            l = stud_keys[random.randint(0,len(stud_keys)-1)]
            if len(studs[l]) > 1:
                # wenn SchülerIn l noch nicht alle 3 Fächer nachgefragt hat
                obj = studs[l].pop()
            else: 
                obj = studs[l].pop()
                helpe = set(stud_keys)
                helpe.remove(l)
                stud_keys = list(helpe)

                # stud_keys = list(stud_keys)
            d = (0,obj,random.randint(1,3),l)  
            
        demand[j] = d
    
    return demand


# In[17]:


def random_teachers_by_m(m):
    
    teachers = {}
    
    for i in range(1,m+1):
        teachers[i] = teacher(0,random.randint(1,19),0,i)
    
    return teachers


# In[18]:


def teachers_by_demand(demand):
    
    teachers = {}
    key = 1
    
    d = sum([demand[i][2] for i in demand.keys()]) 
    supply = 0
    
    while d > supply:
        teachers[key] = teacher(0,random.randint(1,19),0,key)
        supply += teachers[key].h_per_week
        key += 1
        
    return teachers


# In[19]:


def measure_time(teachers,demand,L,formulation,K,students):
    
    results = {}
    
    start_model = time.time()
    model = build_model(teachers,demand,L,formulation,K,students)
    end_model = time.time()
    
    start_solving = time.time()
    #print("start solver:", datetime.datetime.now())
    solution = model.solve()
    #print("end solver:", datetime.datetime.now())
    end_solving = time.time()
    
    if solution == None:
        results["lösbar"] = 0
    else:
        results["lösbar"] = 1

    results["Modellierungszeit"] = end_model-start_model
    results["Lösungszeit"] = end_solving-start_solving
    results["Insgesamt Zeit"] = (end_model-start_model) + (end_solving-start_solving)

        
    return results


# In[20]:


def random_demand_by_L(L):
    
    objectives = ["Deutsch","Mathe","Englisch"]
    demand = {}
    key = 1
    
    for l in range(1,L+1):
        a = random.randint(0,1)
        if a == True:
            d = (0,objectives[0],random.randint(1,3),l)
            demand[key] = d
            key += 1
            
        b = random.randint(0,1)
        if b == True:
            d = (0,objectives[1],random.randint(1,3),l)
            demand[key] = d
            key += 1
            
        if a+b == 0:
            c = 1
        else:
            c = random.randint(0,1)
        if c == True:
            d = (0,objectives[2],random.randint(1,3),l)
            demand[key] = d
            key += 1
        
    return demand


# In[24]:


def data_to_document(data,nv,Lv,Kv,mv):
    
        #####################################################
    df = {
        "m": {},
        "n": {},
        "L": {},
        "K": {},
        "Nachgefragte h": {},
        "durchschn. Soll-Wert": {},
        "durchschn. Lösungszeit": {},
        "durchschn. Modellierungszeit": {},
        "Lösbarkeit": {}
    }
    # n fest, K fest
    wd = os.getcwd()
    
    print("data, counters:",len(data))
    
    for m in mv:
        
        if len(Lv)==7:
            nv = Lv
        
        for n in nv:
            try:
                os.chdir('\\'.join([wd,"Neue Auswertungen"]))
            except:
                os.mkdir('\\'.join([wd,"Neue Auswertungen"]))
                os.chdir('\\'.join([wd,"Neue Auswertungen"]))

            for K in Kv:

            #for L in Lv:
            #for n in [data[counter][K][L]["data"]["n"] for L in Lv]
                #for counter in range(len(data)):
                row = 1


                no_error = [counter for counter in range(len(data)) if data[counter][K][m][n]["results"]["Lösungszeit"]!="Error"]


                for counter in no_error:
                    df["m"][row] = data[counter][K][m][n]["data"]["m"] 
                    df["n"][row] = data[counter][K][m][n]["data"]["n"] 
                    df["L"][row] = data[counter][K][m][n]["data"]["L"] 
                    df["K"][row] = data[counter][K][m][n]["data"]["K"] 
                    df["Nachgefragte h"][row] = data[counter][K][m][n]["data"]["Nachgefragte h"] 
                    df["durchschn. Soll-Wert"][row] = data[counter][K][m][n]["data"]["durchschn. Soll-Wert"] 
                    df["durchschn. Lösungszeit"][row] = data[counter][K][m][n]["results"]["Lösungszeit"] 
                    df["durchschn. Modellierungszeit"][row] = data[counter][K][m][n]["results"]["Modellierungszeit"] 
                    df["Lösbarkeit"][row] = data[counter][K][m][n]["results"]["lösbar"] 
                    row +=1

                df["m"][row] = sum([data[counter][K][m][n]["data"]["m"] for counter in no_error])/len(no_error)
                df["n"][row] = sum([data[counter][K][m][n]["data"]["n"] for counter in no_error])/len(no_error)
                df["L"][row] = sum([data[counter][K][m][n]["data"]["L"] for counter in no_error])/len(no_error)
                df["K"][row] = sum([data[counter][K][m][n]["data"]["K"] for counter in no_error])/len(no_error)
                df["Nachgefragte h"][row] = sum([data[counter][K][m][n]["data"]["Nachgefragte h"] for counter in no_error])/len(no_error)
                df["durchschn. Soll-Wert"][row] = sum([data[counter][K][m][n]["data"]["durchschn. Soll-Wert"] for counter in no_error])/len(no_error)
                df["durchschn. Lösungszeit"][row] = sum([data[counter][K][m][n]["results"]["Lösungszeit"] for counter in no_error])/len(no_error)
                df["durchschn. Modellierungszeit"][row] = sum([data[counter][K][m][n]["results"]["Modellierungszeit"] for counter in no_error])/len(no_error)
                df["Lösbarkeit"][row] = sum([data[counter][K][m][n]["results"]["lösbar"] for counter in no_error])/len(no_error)

                DF = pd.DataFrame(df)

                DF = DF.rename(index={row: 'Durchschnitt:'})
                
                if len(Lv)==7:
                    with pd.ExcelWriter("L{0}_LS_K{1}_m{2}_Ergebnisse.xlsx".format(n,K,m)) as writer:
                        #results_df = pd.DataFrame(results)
                        DF.to_excel(writer)
                        writer.save()
                        #print("it was written")
                else:

                    with pd.ExcelWriter("n{0}_LS_K{1}_m{2}_Ergebnisse.xlsx".format(n,K,m)) as writer:
                        #results_df = pd.DataFrame(results)
                        DF.to_excel(writer)
                        writer.save()
                        #print("it was written")

    os.chdir(wd)
    
     


# In[22]:


formulation = "linear"
#Kv = [40]
#Kv = [20]
Kv = [10]
#mv = [100]
#mv = [50]
mv = [24]
#nv = [320]
#nv = [160]
nv = [120]
#Lv = [160]
#Lv = [80]
Lv = [60]


p = input("Was wird verändert? (K/m/n/L)")

if p=="K":
    #Kv = [10,20,30,40,50,60,70]
    #Kv = [5,10,15,20,25,30,35]
    Kv = [3,5,8,10,13,15,18]
if p=="m":
    #mv = [25,50,75,100,125,150,175]
    #mv = [12,25,37,50,62,75,92]
    mv = [6,12,18,24,30,36,42]
if p=="n":
    #nv = [80,160,240,320,400,480,560]
    nv = [30,60,90,120,150,180,210]
if p=="L":
    #Lv = [40,80,120,160,200,240,280]
    #Lv = [20,40,60,80,100,120,140]
    Lv = [15,30,45,60,75,90,105]
print("Kv:",Kv)
print("mv:",mv)
print("nv:",nv)
print("Lv:",Lv)


# In[23]:



d = {}
c = 5

data = {}

key = 0
data[key] = {
"data": {},
"results": {}
}

for K in tqdm(Kv):
    data[K] = {}

    for m in tqdm(mv):
        data[K][m]={}
        for n in tqdm(nv):
            for L in tqdm(Lv):
                if len(Lv)==1:
                    data[K][m][n] = {
                        "data": {},
                        "results": {}
                    }
                if len(Lv)==7:
                    data[K][m][L] = {
                        "data": {},
                        "results": {}
                    }
                # team and demand 

                teachers = random_teachers_by_m(m)
                if len(Lv)==1:
                    demand = random_demand_by_n(n)
                elif len(Lv) == 7:
                    demand = random_demand_by_L(L)
                #demand = random_demand_by_L(Lv[0])
                students = get_students(demand)
                L = len(students)
                #teachers = teachers_by_demand(demand)
                #teachers = random_teachers_by_m(m)
                m = len(teachers)
                n = len(demand)

                # data
                summe_d = 0
                for dem in demand.values():
                    summe_d += dem[2]
                summe_t = 0
                for t in teachers.values():
                    summe_t += t.h_per_week

                    
                n = len(demand)
                
                
                if len(Lv)==1:
                    
                    data[K][m][n]["data"]["m"] = m
                    print("m=",m)
                    data[K][m][n]["data"]["n"] = n
                    print("n=",n)
                    data[K][m][n]["data"]["L"] = L
                    print("L=",L)
                    data[K][m][n]["data"]["K"] = K
                    print("K=",K)
                    data[K][m][n]["data"]["durchschn. Soll-Wert"] = summe_t/m
                    data[K][m][n]["data"]["Nachgefragte h"] = summe_d

                    for counter in tqdm(range(c)):
                        #d[counter] =
                    # results
                        print(len(teachers),len(demand),L,len(students),K)
                        data[K][m][n]["results"] = measure_time(teachers,demand,L,formulation,K,students)
                        d[counter] = data
                    
                    
                elif len(Lv) == 7:

                    print(L)
                    data[K][m][L]["data"]["m"] = m
                    print("m=",m)
                    data[K][m][L]["data"]["n"] = n
                    print("n=",n)
                    data[K][m][L]["data"]["L"] = L
                    print("L=",L)
                    data[K][m][L]["data"]["K"] = K
                    print("K=",K)
                    data[K][m][L]["data"]["durchschn. Soll-Wert"] = summe_t/m
                    data[K][m][L]["data"]["Nachgefragte h"] = summe_d

                    for counter in tqdm(range(c)):
                        #d[counter] =
                    # results
                        data[K][m][L]["results"] = measure_time(teachers,demand,L,formulation,K,students)
                        d[counter] = data

                        
data_to_document(d,nv,Lv,Kv,mv)


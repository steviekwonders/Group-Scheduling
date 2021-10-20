#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Sep  2 15:24:26 2021

@author: student
"""
# import gurobipy as gp
from gurobipy import quicksum, Model, GRB
import pandas as pd
import os

class group_schedule():
    
    def __init__(self, directory, new_hire_excel_file):
        
        self.attendees, self.groups, self.days, self.pairs_attendees = self.read_in_data(directory, new_hire_excel_file)
        
    def read_in_data(self, directory, new_hire_excel_file):
        
        os.chdir(directory)
        new_hire_df = pd.read_excel(new_hire_excel_file)
        
        attendees = list(new_hire_df['Attendees'])
        groups = ['g'+str(i+1) for i in range(int(new_hire_df['# of Groups'][0]))]
        days = ['d'+str(i+1) for i in range(int(new_hire_df['# of Days'][0]))]
        pairs_attendees = [(a, b) for idx, a in enumerate(attendees) for b in attendees[idx + 1:]]
        
        
        return attendees, groups, days, pairs_attendees
    
    def create_group_schedule(self, max_pairs, max_group_size):
        # --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ### BUILD AND SOLVE MODEL
        
        attendees = self.attendees
        groups = self.groups
        days = self.days
        pairs_attendees = self.pairs_attendees
        
        ## DEFINE the Model
        schedule_dispatch = Model('New Hire Scheduling')
        
        ## DECISION VARIABLES
        x = schedule_dispatch.addVars(groups, days, attendees, vtype=GRB.BINARY, name = 'Placement')
        y = schedule_dispatch.addVars(groups, days, pairs_attendees, vtype=GRB.BINARY, name = 'Pairs')
        z = schedule_dispatch.addVars(pairs_attendees, vtype=GRB.BINARY, name='RealPairs')
        
        ## OBJECTIVE
        schedule_dispatch.ModelSense = GRB.MAXIMIZE
        schedule_dispatch.setObjective(0)
        
        
        schedule_dispatch.addConstr(quicksum(z[pair[0], pair[1]] for pair in pairs_attendees) == max_pairs)
        
        # CONSTRAINTS
        for d in days:
            for a in attendees:
                schedule_dispatch.addConstr(quicksum(x[g,d,a] for g in groups) == 1)
                
        for g in groups:
            for d in days:
                schedule_dispatch.addConstr(quicksum(x[g,d,a] for a in attendees) <= max_group_size)
                
        for pair in pairs_attendees:
            schedule_dispatch.addConstr(quicksum(y[g,d, pair[0], pair[1]] for g in groups for d in days) == z[pair[0], pair[1]])
            #schedule_dispatch.addConstr(quicksum(y[g,d, pair[0], pair[1]] for g in groups for d in days) <= 1)
            # schedule_dispatch.addConstr(0.99989999 + quicksum(y[g,d, pair[0], pair[1]] for g in groups for d in days)/(len(groups)*len(days)) >= z[pair[0], pair[1]])
            
        for g in groups:
            for d in days:
                for pair in pairs_attendees:
                    schedule_dispatch.addConstr(y[g,d,pair[0], pair[1]] <= x[g,d,pair[0]])
                    schedule_dispatch.addConstr(y[g,d,pair[0], pair[1]] <= x[g,d,pair[1]])
                    schedule_dispatch.addConstr(y[g,d,pair[0], pair[1]] >= x[g,d,pair[0]] + x[g,d,pair[1]] - 1)
        
        schedule_dispatch.write('NewHireSchedule.lp') 
                  
        ## Solve
        schedule_dispatch.optimize()
        
        # --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        ### EXCEL OUTPUT: Create an output table that shows the group assignments for each day
        
        for a in attendees:
            for g in groups:
                for d in days:
                    if x[g,d,a].x > 0.5:
                        print(a, g, d)
                        
        # Write out to excel
        group_schedule_df = pd.DataFrame()
        
        gr = []
        att = []
        da = []
                       
        for d in days:
            for g in groups:
                for a in attendees:
                    if x[g,d,a].x > 0.5:
                    
                        att.append(a)
                        gr.append(g)
                        da.append(d)
                    
                 
        group_schedule_df['Day'] = da
        group_schedule_df['Group'] = gr
        group_schedule_df['Attendee'] = att
        
        group_schedule_df.to_excel('Group Schedule.xlsx')

# --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------                       
                        
if __name__ == "__main__":
    schedule_1 = group_schedule('/Users/student/Desktop/SUNYBuffalo', 'NewHires.xlsx')
    schedule_1.create_group_schedule(210, 4)
        
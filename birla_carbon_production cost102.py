from pulp import *
import pandas as pd

#hellos
prod_cost=pd.read_excel("Master.xlsx",sheet_name='Prod cost')
log_cost_converted = pd.read_excel("Master.xlsx",sheet_name = 'logistics_cost_converted')
cp=pd.read_excel("Master.xlsx",sheet_name='cp')
plant_cap=pd.read_excel("Master.xlsx",sheet_name=1)
fixed_cost=pd.read_excel("Master.xlsx",sheet_name = 'Fixedcost')
import_dutyy=pd.read_excel("Master.xlsx",sheet_name = 'Import Duty')

types =['HB','SB']
links = pd.DataFrame()
for typee in types:
    prob=LpProblem("Production",LpMinimize)
    hb=prod_cost[prod_cost['type']==typee]
    plant=list(plant_cap['Plant'])
    customer=list(cp['Customer_Country'])
    demand=dict(zip(customer,cp['Qty']))
    supply=dict(zip(plant,plant_cap['Capacity']))
    pcosts=dict(zip(plant,hb['Prod cost/ton']))
    fixed_cost=fixed_cost.filter(['Plant','Customer_Country','FixedCost'],axis=1)
    dict_fixed_cost={}
    for i in fixed_cost.Plant.unique():
        dict_fixed_cost[i] = {fixed_cost.Customer_Country[i]: fixed_cost.FixedCost[i] for i in fixed_cost[fixed_cost.Plant==i].index}
    
    import_duty=import_dutyy.filter(['Plant','Customer_Country','import_duty'],axis=1)
    dict_import_duty = {}
    for i in import_duty.Plant.unique():
        dict_import_duty[i] = {import_duty.Customer_Country[i]: import_duty.import_duty[i] for i in import_duty[import_duty.Plant==i].index}
    extra_import=import_dutyy.filter(['Plant','Customer_Country','import_extra_per_ton'],axis=1)
    dict_extra_import={}
    for i in extra_import.Plant.unique():
        dict_extra_import[i] = {extra_import.Customer_Country[i]: extra_import.import_extra_per_ton[i] for i in extra_import[extra_import.Plant==i].index}
        
    log_cost_converted = log_cost_converted.filter(['Plant','Customer_Country','log_cost'],axis=1)
    dict_log_cost_converted = {}
    for i in log_cost_converted.Plant.unique():
        dict_log_cost_converted[i] = {log_cost_converted.Customer_Country[i]: log_cost_converted.log_cost[i] for i in log_cost_converted[log_cost_converted.Plant==i].index}
        
    
    plant_vars=LpVariable.dicts("Plant",(plant,customer),0)
    #objective fn
    prob+=lpSum((pcosts[i]*demand[j]*plant_vars[i][j]) for i in plant for j in customer)+lpSum((dict_log_cost[i][j]*demand[j]*plant_vars[i][j]) for i in plant for j in customer)+lpSum((fixed_cost[i]*demand[j]*plant_vars[i][j]) for i in plant for j in customer)+lpSum((pcosts[i]*demand[j]*plant_vars[i][j]+dict_log_cost[i][j]*demand[j]*plant_vars[i][j]+fixed_cost[i]*demand[j]*plant_vars[i][j])*import_duty[i][j] for i in plant for j in customer)
    for j in customer :
        prob += lpSum(plant_vars[i][j] for i in plant) == 1
    for i in plant :
        prob += lpSum(plant_vars[i][j]*demand[j] for j in customer) <= supply[i]
    prob.solve()
    print('Status:',LpStatus[prob.status])
    cols=['Plant','Customer_Country','Value']
    df1=pd.DataFrame(columns=cols)
    for v in prob.variables():
        if v.varValue>0.0001:
            temp=v.name
            words=temp.split('_')
            plant = words[1]
            country = words[2]
            val = v.varValue
            df1 = df1.append([{'Plant': plant,'Customer_Country': country,'Value': val}], ignore_index=True)
    
    x=df1.merge(cp,how='left',on='Customer_Country')
    x['Final_qty']=x.Qty*x.Value
    x['type']=typee
   
    links=links.append(x)
    print(links)
    


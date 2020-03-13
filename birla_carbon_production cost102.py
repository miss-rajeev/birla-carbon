import pandas as pd
import numpy as np
 
existing = pd.read_excel("Master_1.xlsx",sheet_name = 'Existing')

# add production cost
prod_cost = pd.read_excel("Master_1.xlsx",sheet_name = 'Prod cost')
log_cost = pd.read_excel("Master_1.xlsx",sheet_name = 'Logistics freight')
import_duty = pd.read_excel("Master_1.xlsx",sheet_name = 'Import Duty')
plant_capacity = pd.read_excel("Master_1.xlsx",sheet_name = 'Plant Capacity')
fixed_cost = pd.read_excel("Master_1.xlsx",sheet_name = 'Fixedcost')

existing = pd.merge(existing,prod_cost,how='inner',left_on = ['Plant','type'],right_on = ['Plant','type'])
existing = pd.merge(existing,log_cost,how='inner',left_on = ['Plant_Continent','Customer_Continent'],right_on = ['Originn','Destination']).drop(['Originn','Destination'],axis=1)
existing = pd.merge(existing,import_duty.filter(['Plant','Customer_Country','import_duty','import_extra_per_ton'],axis=1),how='inner',left_on = ['Plant','Customer_Country'],right_on = ['Plant','Customer_Country'])


existing['FS_Qty'] = existing['FS_Yield']*existing['Qty AMJ']
existing['FS_total_cost'] = existing['FS_Qty']*existing['FS_cost_per_ton']
existing['NG_cost_per_ton'] = existing['NG_m3_per_ton']*existing['NG_cost_per_m3']
existing['NG_total_cost'] = existing['Qty AMJ']*existing['NG_cost_per_ton']
existing['Energy_total_cost'] = existing['Qty AMJ']*existing['Energy_cost_per_ton_CB']
existing['Production_total_cost'] = existing['FS_total_cost']+existing['NG_total_cost']-existing['Energy_total_cost']
existing['Logistics_total_cost'] = existing['log_cost_per_ton']*existing['Qty AMJ'] 
existing['Import_duty_cost'] = (existing['Production_total_cost']+existing['Logistics_total_cost']+400*existing['Qty AMJ'])*existing['import_duty'] + existing['Qty AMJ']*existing['import_extra_per_ton']
existing['Total_cost'] = existing['Production_total_cost']+existing['Logistics_total_cost']+existing['Import_duty_cost']

existing = existing.filter(['Plant', 'Plant_Country', 'Plant_Continent', 'type',
       'Customer_Country', 'Customer_Continent', 'Qty AMJ','FS_Yield',
       'NG_m3_per_ton', 'FS_cost_per_ton', 'NG_cost_per_m3',
       'Energy_cost_per_ton_CB','log_cost_per_ton', 'import_duty',
       'import_extra_per_ton', 'FS_Qty', 'FS_total_cost', 'NG_cost_per_ton',
       'NG_total_cost', 'Energy_total_cost', 'Production_total_cost',
       'Logistics_total_cost', 'Import_duty_cost', 'Total_cost'],axis=1)

existing.to_excel('Existing_Scenario_3.xlsx',index=False)



######### Reallocation
# import_duty dict
dict_import_duty = {}
for i in import_duty.Plant.unique():
    dict_import_duty[i] = {import_duty.Customer_Country[i]: import_duty.import_duty[i] for i in import_duty[import_duty.Plant==i].index}

# import_duty_extra dict
dict_import_duty_extra = {}
for i in import_duty.Plant.unique():
    dict_import_duty_extra[i] = {import_duty.Customer_Country[i]: import_duty.import_extra_per_ton[i] for i in import_duty[import_duty.Plant==i].index}

# logistics cost dict
log_cost_converted = pd.read_excel("Master_file_upd2.xlsx",sheet_name = 'logistics_cost_converted')
log_cost_converted = log_cost_converted.filter(['Plant','Customer_Country','log_cost'],axis=1)
dict_log_cost = {}
for i in log_cost_converted.Plant.unique():
    dict_log_cost[i] = {log_cost_converted.Customer_Country[i]: log_cost_converted.log_cost[i] for i in log_cost_converted[log_cost_converted.Plant==i].index}

# fixed cost dict
dict_fixed_cost = {}
for i in fixed_cost.Plant.unique(): 
    dict_fixed_cost[i] = {fixed_cost.Customer_Country[i]: fixed_cost.FixedCost[i] for i in fixed_cost[fixed_cost.Plant==i].index}

types = ['HB','SB']
links=pd.DataFrame()

existing1 = existing[existing['Qty AMJ'] > 100]
existing2 = existing[existing['Qty AMJ'] <= 100]

for typee in types:
    ex2 = existing2[existing2.type == typee]
    ex2_aggr = ex2.groupby(['Plant'],as_index=False)['Qty AMJ'].sum()
    p=plant_capacity[plant_capacity.type==typee]
    p = pd.merge(p,ex2_aggr,how = 'left',left_on=['Plant'],right_on=['Plant'])
    p['Qty AMJ'].fillna(0,inplace=True)
    p.Capacity = p.Capacity - p['Qty AMJ']
    
    plants = p['Plant'].tolist()
   
    
    PlantCapacity = p.set_index('Plant')['Capacity'].to_dict() 
    PlantCapacity
    c=existing1[existing1.type==typee]
    
    cust_level = c.groupby(['Customer_Country'], as_index=False)['Qty AMJ'].sum()
    customers = cust_level.Customer_Country.tolist()
    demand = {}
    demand = cust_level.set_index('Customer_Country')['Qty AMJ'].to_dict()
    
    inter = prod_cost[prod_cost.type == typee]
    inter['NG_cost_per_ton'] = inter['NG_m3_per_ton']*inter['NG_cost_per_m3']
    dict_FS_yield = {}
    dict_FS_yield = inter.set_index('Plant')['FS_Yield'].to_dict()
    dict_FS_cost_per_ton = {}
    dict_FS_cost_per_ton = inter.set_index('Plant')['FS_cost_per_ton'].to_dict()
    dict_NG_cost_per_ton = {}
    dict_NG_cost_per_ton = inter.set_index('Plant')['NG_cost_per_ton'].to_dict()
    dict_Energy_cost_per_ton = {}
    dict_Energy_cost_per_ton = inter.set_index('Plant')['Energy_cost_per_ton_CB'].to_dict()
                                                                                             
    # optimization                                                                                           
    from pulp import *
    prob = LpProblem('Reallocation', LpMinimize)
    serv_vars = LpVariable.dicts("Service",[(k,i) for i in customers for k in plants],0)
    prob += lpSum(serv_vars[(k,i)]*demand[i]*dict_FS_yield[k]*dict_FS_cost_per_ton[k] for k in plants for i in customers) + lpSum(serv_vars[(k,i)]*demand[i]*dict_NG_cost_per_ton[k] for k in plants for i in customers) - lpSum(serv_vars[(k,i)]*demand[i]*dict_Energy_cost_per_ton[k] for k in plants for i in customers) + lpSum(serv_vars[(k,i)]*demand[i]*dict_log_cost[k][i] for k in plants for i in customers) + lpSum(serv_vars[(k,i)]*dict_fixed_cost[k][i] for k in plants for i in customers) + lpSum(serv_vars[(k,i)]*((demand[i]*dict_FS_yield[k]*dict_FS_cost_per_ton[k]+demand[i]*dict_NG_cost_per_ton[k]-demand[i]*dict_Energy_cost_per_ton[k]+demand[i]*dict_log_cost[k][i]+dict_fixed_cost[k][i]+400*demand[i])*dict_import_duty[k][i]+demand[i]*dict_import_duty_extra[k][i]) for k in plants for i in customers)
    
    # customer demand fulfillment         
    for i in customers:
        prob += lpSum(serv_vars[(k,i)] for k in plants) == 1
    # plant capacity contraint
    for k in plants:
        prob += lpSum(serv_vars[(k,i)]*demand[i] for i in customers) <=  PlantCapacity[k]
        
    prob.solve()
    print("Status:",LpStatus[prob.status])
    print("Cost of production in a year:", value(prob.objective)/1000000)
    
    cols = ['Plant','Customer_Country','Value']
    df1 =  pd.DataFrame(columns=cols)
    for v in prob.variables():
        if v.varValue > 0.0001 and v.name.startswith('Service'):
            temp = v.name
            string = temp.replace('Service_','').replace("'","").replace('_','-').replace('(','').replace(')','').split(',')
            df1 = df1.append([{'Plant': string[0],'Customer_Country': string[1],'Value': v.varValue}], ignore_index=True)
    df1['Customer_Country'] = df1['Customer_Country'].str[1:]
    df1['type']=typee
    
    links=links.append(df1)


# attach demand
cust_demand = existing1.groupby(['Customer_Country','type'], as_index=False)['Qty AMJ'].sum()
links2 = pd.merge(links,cust_demand,left_on=['Customer_Country','type'],right_on=['Customer_Country','type'],how='left')
links2['Qty_AMJ_final'] = links2['Value']*links2['Qty AMJ']
    
realloc = pd.merge(links2,prod_cost,how='inner',left_on = ['Plant','type'],right_on = ['Plant','type'])

#mapping
mapping = existing.filter(['Customer_Country','Customer_Continent'])
mapping = mapping.drop_duplicates()
realloc = pd.merge(realloc,mapping,how='inner',left_on = ['Customer_Country'],right_on = ['Customer_Country'])

mapping = existing.filter(['Plant','Plant_Country','Plant_Continent'])
mapping = mapping.drop_duplicates()
realloc = pd.merge(realloc,mapping,how='inner',left_on = ['Plant'],right_on = ['Plant'])

realloc = pd.merge(realloc,log_cost,how='inner',left_on = ['Plant_Continent','Customer_Continent'],right_on = ['Originn','Destination']).drop(['Originn','Destination'],axis=1)
realloc = pd.merge(realloc,import_duty.filter(['Plant','Customer_Country','import_duty','import_extra_per_ton'],axis=1),how='inner',left_on = ['Plant','Customer_Country'],right_on = ['Plant','Customer_Country'])
realloc = pd.merge(realloc,fixed_cost.filter(['Plant','Customer_Country','FixedCost']),how='inner',left_on = ['Plant','Customer_Country'],right_on = ['Plant','Customer_Country'])

realloc['FS_Qty'] = realloc['FS_Yield']*realloc['Qty_AMJ_final']
realloc['FS_total_cost'] = realloc['FS_Qty']*realloc['FS_cost_per_ton']
realloc['NG_cost_per_ton'] = realloc['NG_m3_per_ton']*realloc['NG_cost_per_m3']
realloc['NG_total_cost'] = realloc['Qty_AMJ_final']*realloc['NG_cost_per_ton']
realloc['Energy_total_cost'] = realloc['Qty_AMJ_final']*realloc['Energy_cost_per_ton_CB']
realloc['Production_total_cost'] = realloc['FS_total_cost']+realloc['NG_total_cost']-realloc['Energy_total_cost']
realloc['Logistics_total_cost'] = realloc['log_cost_per_ton']*realloc['Qty_AMJ_final'] 
realloc['Import_duty_cost'] = (realloc['Production_total_cost']+realloc['Logistics_total_cost']+realloc['FixedCost']+400*realloc['Qty_AMJ_final'])*realloc['import_duty']+realloc['Qty_AMJ_final']*realloc['import_extra_per_ton']
realloc['Total_cost'] = realloc['Production_total_cost']+realloc['Logistics_total_cost']+realloc['FixedCost']+realloc['Import_duty_cost']

realloc = realloc.filter(['Plant', 'Plant_Country', 'Plant_Continent', 'type',
       'Customer_Country', 'Customer_Continent', 'Qty_AMJ_final','FS_Yield',
       'NG_m3_per_ton', 'FS_cost_per_ton', 'NG_cost_per_m3',
       'Energy_cost_per_ton_CB','log_cost_per_ton', 'import_duty',
       'import_extra_per_ton', 'FS_Qty', 'FS_total_cost', 'NG_cost_per_ton',
       'NG_total_cost', 'Energy_total_cost', 'Production_total_cost',
       'Logistics_total_cost', 'Import_duty_cost', 'FixedCost' ,'Total_cost'],axis=1)

existing2['FixedCost'] = 0
existing2['Qty_AMJ_final'] = existing2['Qty AMJ']
existing2 = existing2.drop(['Qty AMJ'],axis=1)

existing2 = existing2.filter(['Plant', 'Plant_Country', 'Plant_Continent', 'type',
       'Customer_Country', 'Customer_Continent', 'Qty_AMJ_final','FS_Yield',
       'NG_m3_per_ton', 'FS_cost_per_ton', 'NG_cost_per_m3',
       'Energy_cost_per_ton_CB','log_cost_per_ton', 'import_duty',
       'import_extra_per_ton', 'FS_Qty', 'FS_total_cost', 'NG_cost_per_ton',
       'NG_total_cost', 'Energy_total_cost', 'Production_total_cost',
       'Logistics_total_cost', 'Import_duty_cost', 'FixedCost' ,'Total_cost'],axis=1)

realloc = realloc.append(existing2)

realloc.to_excel('Realloc_Scenario.xlsx',index=False)
    
    
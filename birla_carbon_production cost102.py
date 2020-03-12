from pulp import *
import pandas as pd

prod_cost=pd.read_excel("Master.xlsx",sheet_name='Prod cost')
log_cost_converted = pd.read_excel("Master.xlsx",sheet_name = 'logistics_cost_converted')
cp=pd.read_excel("Master.xlsx",sheet_name='cp')
plant_cap=pd.read_excel("Master.xlsx",sheet_name='kk')
fixed_cost=pd.read_excel("Master.xlsx",sheet_name = 'Fixedcost')
import_dutyy=pd.read_excel("Master.xlsx",sheet_name = 'Import Duty')
existing=pd.read_excel("Master.xlsx",sheet_name='existing')
log_= pd.read_excel("Master.xlsx",sheet_name = 'Logistics freight')
existing = pd.merge(existing,prod_cost,how='inner',left_on = ['Plant','type'],right_on = ['Plant','type'])
existing = pd.merge(existing,log_,how='inner',left_on = ['Plant_Continent','Customer_Continent'],right_on = ['Originn','Destination']).drop(['Originn','Destination'],axis=1)
existing = pd.merge(existing,import_dutyy.filter(['Plant','Customer_Country','import_duty','import_extra_per_ton'],axis=1),how='inner',left_on = ['Plant','Customer_Country'],right_on = ['Plant','Customer_Country'])


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

existing.to_excel('Existing_Scenario2.xlsx',index=False)

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
        
dict_log_cost_converted = {}
for i in log_cost_converted.Plant.unique():
    dict_log_cost_converted[i] = {log_cost_converted.Customer_Country[i]: log_cost_converted.log_cost[i] for i in log_cost_converted[log_cost_converted.Plant==i].index}


types =['HB','SB']
links = pd.DataFrame()
existing1 = existing[existing['Qty AMJ'] > 100]
existing2 = existing[existing['Qty AMJ'] <= 100]
for typee in types:
    ex2=existing2[existing2['type']==typee]
    ex2_aggr=ex2.groupby(['Plant'],as_index=False)['Qty AMJ'].sum()
    p=plant_cap[plant_cap.type==typee]
    p=pd.merge(p,ex2_aggr,how ='left',left_on=['Plant'],right_on=['Plant'])
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
    
    inter=prod_cost[prod_cost.type == typee]
    inter['NG_cost_per_ton'] = inter['NG_m3_per_ton']*inter['NG_cost_per_m3']
    dict_FS_yield = {}
    dict_FS_yield = inter.set_index('Plant')['FS_Yield'].to_dict()
    dict_FS_cost_per_ton = {}
    dict_FS_cost_per_ton = inter.set_index('Plant')['FS_cost_per_ton'].to_dict()
    dict_NG_cost_per_ton = {}
    dict_NG_cost_per_ton = inter.set_index('Plant')['NG_cost_per_ton'].to_dict()
    dict_Energy_cost_per_ton = {}
    dict_Energy_cost_per_ton = inter.set_index('Plant')['Energy_cost_per_ton_CB'].to_dict()
    
    prob=LpProblem("Reallocation",LpMinimize)
    serv_vars=LpVariable.dicts("Service",[(k,i) for i in customers for k in plants],0)
    #objective fn
    prob += lpSum(serv_vars[(k,i)]*demand[i]*dict_FS_yield[k]*dict_FS_cost_per_ton[k] for k in plants for i in customers) + lpSum(serv_vars[(k,i)]*demand[i]*dict_NG_cost_per_ton[k] for k in plants for i in customers) - lpSum(serv_vars[(k,i)]*demand[i]*dict_Energy_cost_per_ton[k] for k in plants for i in customers) + lpSum(serv_vars[(k,i)]*demand[i]*dict_log_cost[k][i] for k in plants for i in customers) + lpSum(serv_vars[(k,i)]*dict_fixed_cost[k][i] for k in plants for i in customers) + lpSum(serv_vars[(k,i)]*((demand[i]*dict_FS_yield[k]*dict_FS_cost_per_ton[k]+demand[i]*dict_NG_cost_per_ton[k]-demand[i]*dict_Energy_cost_per_ton[k]+demand[i]*dict_log_cost[k][i]+dict_fixed_cost[k][i]+400*demand[i])*dict_import_duty[k][i]+demand[i]*dict_import_duty_extra[k][i]) for k in plants for i in customers)
    for i in customers:
        prob+= lpSum(serv_vars[(k,i)] for k in plants) == 1
    for k in plants:
        prob+= lpSum(serv_vars[(k,i)]*demand[j] for i in customers) <= PlantCapacity[k]
    prob.solve()
    print('Status:',LpStatus[prob.status])
    print("Cost in a year:", value(prob.objective)/1000000)
    
    cols=['Plant','Customer_Country','Value',]
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
    #print(links)

# attach demand
cust_demand = links.groupby(['Customer_Country','type'], as_index=False)['Qty'].sum()
links2 = pd.merge(links,cust_demand,left_on=['Customer_Country','type'],right_on=['Customer_Country','type'],how='left')
links2['Qty_AMJ_final'] = links2['Value']*links2['Qty_x']
    
realloc = pd.merge(links2,prod_cost,how='inner',left_on = ['Plant','type'],right_on = ['Plant','type'])

#mapping
mapping = existing.filter(['Customer_Country','Customer_Continent'])
mapping = mapping.drop_duplicates()
realloc = pd.merge(realloc,mapping,how='inner',left_on = ['Customer_Country'],right_on = ['Customer_Country'])

mapping = existing.filter(['Plant','Plant_Country','Plant_Continent'])
mapping = mapping.drop_duplicates()
realloc = pd.merge(realloc,mapping,how='inner',left_on = ['Plant'],right_on = ['Plant'])

realloc = pd.merge(realloc,log_,how='inner',left_on = ['Plant_Continent','Customer_Continent'],right_on = ['Originn','Destination']).drop(['Originn','Destination'],axis=1)
realloc = pd.merge(realloc,import_dutyy.filter(['Plant','Customer_Country','import_duty','import_extra_per_ton'],axis=1),how='inner',left_on = ['Plant','Customer_Country'],right_on = ['Plant','Customer_Country'])
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





realloc.to_excel('Realloc_Scenario_1.xlsx',index=False)
    
    
   
   
    
    


    
  
    
    
    
    
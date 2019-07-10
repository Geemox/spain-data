import pandas as pd # For creating efficient data structures
import numpy as np # For linear algebra operations on N-dimentional arrays
import os # For path manipulation functions

# Assign spreadsheet filename to `file`
file = 'METIS_Data_processing v1.xlsm'

# Load spreadsheet
xl = pd.ExcelFile(file)

# Load a sheet into a DataFrame by name: df1
dfg = xl.parse('generation', header=1)
dfv = xl.parse('Decision Variables Urb-Sem-Rur', header=0)

s1 = set(dfg.index[dfg['Country ID'] == 'ES'].tolist())
s2 = set(dfg.index[dfg['Asset'] == 'Solar availability'].tolist())
indx_es_solar = list(s1.intersection(s2))[0]

gen = dfg.iloc[indx_es_solar][5:]
g = np.array(gen.values[:-3])

dfv.set_index('Column1',inplace=True, drop=True)

dx_coef=0.5

total_consumers = dfv.loc['Urban','Total Consumers'].values[0] + dfv.loc['Semi-urban','Total Consumers'].values[0] + dfv.loc['Rural','Total Consumers'].values[0]

lv_urban_coef= dfv.loc['Urban','Consumers <1kV'].values[0]/ total_consumers
mv_urban_coef= dfv.loc['Urban','Consumers  1-100 kV'].values[0]/ total_consumers
hv_urban_coef= dfv.loc['Urban','Consumers >100kV'].values[0]/ total_consumers

Sup_urban_ntw= dfv.loc['Urban','Sup Country (km2)'].values[0]

g_urban_lv = g*dx_coef*lv_urban_coef/Sup_urban_ntw
g_urban_mv = g*dx_coef*mv_urban_coef/Sup_urban_ntw
g_urban_hv = g*dx_coef*hv_urban_coef/Sup_urban_ntw

gen_export = pd.DataFrame({'VL1': g_urban_lv,
					'VL2': g_urban_mv,
					'VL3': g_urban_hv})
gen_export.to_csv('generation.csv',index=False, sep=';')

#print(gen.index.tolist()[:-3])

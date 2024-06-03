# -*- coding: utf-8 -*-
"""
Created on Tue Mar 12 15:19:32 2024

@author: runan
"""
import pandas as pd

def get_floor_span():
    span = float(input("Please enter the span of the floor in meters:"))
    return span

floor_span = get_floor_span()
print("The span of the floor is:", floor_span,"m")

def get_floor_width():
    width = float(input("Please enter the width of the floor in meter (=length of beams):"))
    return width

floor_width = get_floor_width()
print("The width of the floor is:", floor_width,"m")

def get_structure_height():
    height = float(input("Please enter the height of the structure in meters:"))
    return height

structure_height = get_structure_height()
print("The height of the structure is:", structure_height,"m")

floor_area = floor_span*floor_width

def get_permanent_load():
    permanent_load = float(input("Please enter the uniformly spread permanent load on your floor in kN/m²:"))
    return permanent_load

floor_permanent_load = get_permanent_load()

def get_variable_load():
    variable_load = float(input("Please enter the uniformly spread variable load on your floor in kN/m²:"))
    return variable_load

floor_variable_load = get_variable_load()

def get_additional_permanent_load_beams():
    additional_permanent_beams = float(input("Please enter the additional permanent spread load on your beams in kN/m:"))
    return additional_permanent_beams

additional_permanent_load_beams = get_additional_permanent_load_beams()


def get_additional_variable_load_beams():
    additional_variable_beams = float(input("Please enter the additional variable spread load on your beams in kN/m:"))
    return additional_variable_beams

additional_variable_load_beams = get_additional_variable_load_beams()

def get_additional_permanent_load_columns():
    additional_permanent_columns = float(input("Please enter the additional permanent point load on your columns in kN:"))
    return additional_permanent_columns

additional_permanent_load_columns = get_additional_permanent_load_columns()

def get_additional_variable_load_columns():
    additional_variable_columns = float(input("Please enter the additional variable point load on your columns in kN:"))
    return additional_variable_columns

additional_variable_load_columns = get_additional_variable_load_columns()

load_ULS = 1.35*floor_permanent_load + 1.50*floor_variable_load
load_SLS = floor_permanent_load + floor_variable_load
load_ULS_HCS = 1.35*(floor_permanent_load+1.2)*1.196 + 1.50*floor_variable_load*1.196
print(load_ULS)
print(load_SLS)

acting_moment_CLT = load_ULS * floor_span**2/8
acting_moment_HCS = load_ULS_HCS * floor_span**2/8




#HOLLOW CORE SLABS
def find_closest_greater_value_HCS(excel_file, sheet_name_HCS, search_column_HCS, search_value_HCS, target_column_HCS):
    df = pd.read_excel(excel_file, sheet_name=sheet_name_HCS)
    
    while True:
        df_filtered = df[df[search_column_HCS] >= search_value_HCS]
        
        if df_filtered.empty:
            return None
        
        index = df_filtered[search_column_HCS].idxmin()       
        target_value = df.loc[index, target_column_HCS]
        
        load_ULS_HCS_loop = 1.35 * (floor_permanent_load + 1.2) * 1.196 + 1.35 * target_value + 1.50 * floor_variable_load * 1.196
        moment_ULS_loop = load_ULS_HCS_loop * floor_span ** 2 / 8
        
        if moment_ULS_loop == search_value_HCS:
            return moment_ULS_loop
        
        df[search_column_HCS] = df[search_column_HCS].astype(float)
        df.at[index, search_column_HCS] = moment_ULS_loop
      
        search_value_HCS = moment_ULS_loop

excel_file = r"C:\Users\runan\Documents\Masterproef\2024 April-mei\Structural materials.xlsx"
sheet_name_HCS = "HCS - Floor"
search_column_HCS = "M_Rd (with 50mm compression layer) [kNm]" 
initial_search_value_HCS = acting_moment_HCS  
target_column_HCS = "Self weight [kN/m]" 

result_HCS = find_closest_greater_value_HCS(excel_file, sheet_name_HCS, search_column_HCS, initial_search_value_HCS, target_column_HCS)


def find_closest_greater_value_HCS2(excel_file, sheet_name_HCS2, search_column_HCS2, search_value_HCS2, target_column_HCS2):
    df = pd.read_excel(excel_file, sheet_name=sheet_name_HCS2)
    df_filtered = df[df[search_column_HCS2] >= search_value_HCS2]
    
    if df_filtered.empty:
        return None
    
    differences = (df_filtered[search_column_HCS2] - search_value_HCS2).abs()
    
    index = differences.idxmin()
    
    return df.loc[index, target_column_HCS2]

excel_file = r"C:\Users\runan\Documents\Masterproef\2024 April-mei\Structural materials.xlsx"
sheet_name_HCS2 = "HCS - Floor"
search_column_HCS2 = "M_Rd (with 50mm compression layer) [kNm]"
search_value_HCS2 = result_HCS

area_concrete_HCS = find_closest_greater_value_HCS2(excel_file, sheet_name_HCS2, search_column_HCS2, search_value_HCS2, "A_concrete [mm²]")  
area_steel_HCS = find_closest_greater_value_HCS2(excel_file, sheet_name_HCS2, search_column_HCS2, search_value_HCS2, "A_steel [mm²]")  
thickness_HCS = find_closest_greater_value_HCS2(excel_file, sheet_name_HCS2, search_column_HCS2, search_value_HCS2, "Thickness [mm]")  
self_weight_HCS = find_closest_greater_value_HCS2(excel_file, sheet_name_HCS2, search_column_HCS2, search_value_HCS2, "Self weight [kN/m]")  
type_HCS = find_closest_greater_value_HCS2(excel_file, sheet_name_HCS2, search_column_HCS2, search_value_HCS2, "Type")
print(thickness_HCS)
print(type_HCS) 




#IN SITU CAST CONCRETE
E_concrete = 10000000000

from sympy.solvers import solve
from sympy import Symbol

x = Symbol("x")
thickness_concrete = solve(((1500*(load_ULS+25*x)*1000*(floor_span**3))/(32*E_concrete*(x**3)))-1, x)

first_thickness_concrete=thickness_concrete[0]

floor_depth_concrete = float(first_thickness_concrete)
print(floor_depth_concrete)

import math

def round_up_to_two_decimals(number):
    return math.ceil(number * 100) / 100

thickness_concrete_rounded = round_up_to_two_decimals(floor_depth_concrete)
print(thickness_concrete_rounded)




#CROSS-LAMINATED TIMBER 
limit_deflection_floor_CLT = (floor_span*1000)/250
E_CLT_floor = 12000/1.25

def find_closest_greater_value_floor_CLT(excel_file, sheet_name_concrete, check_threshold):
    df = pd.read_excel(excel_file, sheet_name=sheet_name_concrete)
    df = df.sort_values(by='I [mm^4]')
    
    for index, row in df.iterrows():
        secondmomentofarea_floor_CLT = row['I [mm^4]']
        self_weight_floor_CLT = row['Self weight [kN/m]']
        thickness_floor_CLT = row['Thickness [mm]']
        type_floor_CLT = row['CLT type']
        
        deflection_floor_CLT_char = (5/384) * ((floor_permanent_load + floor_variable_load + self_weight_floor_CLT) * (floor_span * 1000)**4) / (E_CLT_floor * secondmomentofarea_floor_CLT)
        deflection_floor_CLT_qp = (5/384) * ((floor_permanent_load + 0.3 * floor_variable_load + self_weight_floor_CLT) * (floor_span * 1000)**4) / (E_CLT_floor * secondmomentofarea_floor_CLT)
        deflection_floor_CLT_total = deflection_floor_CLT_char + 0.8 * deflection_floor_CLT_qp
        if check_threshold >= deflection_floor_CLT_total:
            return deflection_floor_CLT_total, thickness_floor_CLT, type_floor_CLT, self_weight_floor_CLT
    print("End of DataFrame reached without finding a suitable result.")
    return None, None, None

sheet_name_floor_CLT = "CLT - Floor"

deflection_floor_CLT, thickness_CLT, type_floor_CLT, self_weight_floor_CLT = find_closest_greater_value_floor_CLT(excel_file, sheet_name_floor_CLT, limit_deflection_floor_CLT)
print(deflection_floor_CLT)
print(thickness_CLT)
print(type_floor_CLT)
print(self_weight_floor_CLT)

thickness_CLT_m = thickness_CLT/1000

self_weight_HCS_total = (self_weight_HCS/1.196) * floor_area
self_weight_CLT_total = self_weight_floor_CLT * floor_area
self_weight_concrete_total = 25 * thickness_concrete_rounded * floor_area




#VOLUMES FLOORS
Volume_concrete_one_HCS = area_concrete_HCS/(1000*1000)*floor_span
Volume_steel_one_HCS = area_steel_HCS/(1000*1000)*floor_span
Volume_concrete_floor_HCS = (floor_width/1.196) * Volume_concrete_one_HCS
Volume_steel_floor_HCS = (floor_width/1.196) * Volume_steel_one_HCS
Volume_compression_layer_HCS = 0.05 * floor_area

Volume_in_situ_concrete_total = floor_depth_concrete*floor_area
Volume_reinforcement_in_situ_concrete = Volume_in_situ_concrete_total*0.015
Volume_in_situ_concrete = Volume_in_situ_concrete_total - Volume_reinforcement_in_situ_concrete

Volume_CLT = (thickness_CLT/1000)*floor_area




#BEAMS STEEL & GLT - HCS FLOOR
ULS_load_beam_HCS = (1.35*self_weight_HCS_total + load_ULS*floor_area)/(2*floor_width) + 1.35*additional_permanent_load_beams + 1.50*additional_variable_load_beams
print(f"ULS_load_beam_HCS: {ULS_load_beam_HCS}")
acting_moment_beam_HCS = ULS_load_beam_HCS * floor_width ** 2 / 8

def find_closest_greater_value_beam_HCS(excel_file, sheet_name_HCS, search_column_HCS, search_value_HCS, target_column_HCS):
    df = pd.read_excel(excel_file, sheet_name=sheet_name_HCS)
    
    while True:
        df_filtered = df[df[search_column_HCS] >= search_value_HCS]
        
        if df_filtered.empty:
            return None
        
        index = df_filtered[search_column_HCS].idxmin()
        target_value = df.loc[index, target_column_HCS]
        
        load_ULS_HCS_loop = ULS_load_beam_HCS  + 1.35 * target_value
        moment_ULS_loop_HCS = load_ULS_HCS_loop * floor_width ** 2 / 8
        
        if moment_ULS_loop_HCS == search_value_HCS:
            return moment_ULS_loop_HCS
        
        df[search_column_HCS] = df[search_column_HCS].astype(float)
        df.at[index, search_column_HCS] = moment_ULS_loop_HCS
        
        search_value_HCS = moment_ULS_loop_HCS

 
excel_file = r"C:\Users\runan\Documents\Masterproef\2024 April-mei\Structural materials.xlsx"
sheet_name_beam_steel = "Steel - Beam"
sheet_name_beam_concrete = "Concrete - Beam"
sheet_name_beam_GLT = "GLT - Beam"
search_column_beam_steel_MRd = "M_Rd [kNm]" 
search_column_beam_GLT_MRd = "M_Rd [kNm] ULS flexural"
initial_search_value_beam_HCS = acting_moment_beam_HCS
target_column_self_weight = "Self weight [kN/m]" 

MEd_beam_steel_HCS = find_closest_greater_value_beam_HCS(excel_file, sheet_name_beam_steel, search_column_beam_steel_MRd, initial_search_value_beam_HCS, target_column_self_weight)

MEd_beam_GLT_HCS = find_closest_greater_value_beam_HCS(excel_file, sheet_name_beam_GLT, search_column_beam_GLT_MRd, initial_search_value_beam_HCS, target_column_self_weight)

def find_corresponding_value_other_column(excel_file, sheet_name_HCS2, search_column_HCS2, search_value_HCS2, target_column_HCS2):
    df = pd.read_excel(excel_file, sheet_name=sheet_name_HCS2)
    df_filtered = df[df[search_column_HCS2] >= search_value_HCS2]
    
    if df_filtered.empty:
        return None
    
    differences = (df_filtered[search_column_HCS2] - search_value_HCS2).abs()
    index = differences.idxmin()
    
    return df.loc[index, target_column_HCS2]

area_beam_steel_HCS = find_corresponding_value_other_column(excel_file, sheet_name_beam_steel, search_column_beam_steel_MRd, MEd_beam_steel_HCS, "Area [mm²]")  
self_weight_beam_steel_HCS = find_corresponding_value_other_column(excel_file, sheet_name_beam_steel, search_column_beam_steel_MRd, MEd_beam_steel_HCS, "Self weight [kN/m]")  
profile_beam_steel_HCS = find_corresponding_value_other_column(excel_file, sheet_name_beam_steel, search_column_beam_steel_MRd, MEd_beam_steel_HCS, "Profile")  
height_beam_steel_HCS = find_corresponding_value_other_column(excel_file, sheet_name_beam_steel, search_column_beam_steel_MRd, MEd_beam_steel_HCS, "Height [mm]")  
print(profile_beam_steel_HCS) 

area_beam_GLT_HCS = find_corresponding_value_other_column(excel_file, sheet_name_beam_GLT, search_column_beam_GLT_MRd, MEd_beam_GLT_HCS, "Area [mm²]")  
self_weight_beam_GLT_HCS = find_corresponding_value_other_column(excel_file, sheet_name_beam_GLT, search_column_beam_GLT_MRd, MEd_beam_GLT_HCS, "Self weight [kN/m]")  
profile_beam_GLT_HCS = find_corresponding_value_other_column(excel_file, sheet_name_beam_GLT, search_column_beam_GLT_MRd, MEd_beam_GLT_HCS, "Dimensions")  
height_beam_GLT_HCS = find_corresponding_value_other_column(excel_file, sheet_name_beam_GLT, search_column_beam_GLT_MRd, MEd_beam_GLT_HCS, "height [mm]")  
print(profile_beam_GLT_HCS) 

Volume_beam_steel_HCS = area_beam_steel_HCS / (1000 * 1000) * floor_width
Volume_beam_GLT_HCS = area_beam_steel_HCS/(1000*1000)*floor_width





#BEAMS CONCRETE - HCS FLOOR
E_concrete_beam = 10000
limit_deflection_beam_concrete = floor_width * 1000 / 300

load_SLS_beam_HCS = load_SLS * floor_area/(2*floor_width) + self_weight_HCS_total/(2*floor_width) + additional_permanent_load_beams + additional_variable_load_beams

def find_closest_greater_value_beam_concrete_HCS(excel_file, sheet_name_concrete, check_threshold):
    df = pd.read_excel(excel_file, sheet_name=sheet_name_concrete)
    df = df.sort_values(by='I_y [mm^4]')
    
    for index, row in df.iterrows():
        secondmomentofarea_concrete = row['I_y [mm^4]']
        self_weight_beam_concrete = row['self_weight [kN/m]']
        area_beam_concrete = row['Area [mm²]']
        profile_beam_concrete = row['Dimensions']
        height_beam_concrete = row['height [mm]']
        
        deflection_beam_concrete = (5/384) * ((load_SLS_beam_HCS + self_weight_beam_concrete) * (floor_width * 1000)**4) / (E_concrete_beam * secondmomentofarea_concrete)
        
        if check_threshold >= deflection_beam_concrete:
            return deflection_beam_concrete, area_beam_concrete, profile_beam_concrete, self_weight_beam_concrete, height_beam_concrete
    
    print("End of DataFrame reached without finding a suitable result.")
    return None, None, None

excel_file = r"C:\Users\runan\Documents\Masterproef\2024 April-mei\Structural materials.xlsx"
sheet_name_beam_concrete = "Concrete - Beam"

deflection_beam_concrete_HCS, area_beam_concrete_HCS, profile_beam_concrete_HCS, self_weight_beam_concrete_HCS, height_beam_concrete_HCS = find_closest_greater_value_beam_concrete_HCS(excel_file, sheet_name_beam_concrete, limit_deflection_beam_concrete)
print(profile_beam_concrete_HCS)

Volume_beam_concrete_HCS = area_beam_concrete_HCS / (1000 * 1000) * floor_width





#BEAMS STEEL & GLT - CLT FLOOR
ULS_load_beam_CLT =(1.35*self_weight_CLT_total + load_ULS*floor_area)/(2*floor_width) + 1.35*additional_permanent_load_beams + 1.50*additional_variable_load_beams
print(f"ULS_load_beam_CLT: {ULS_load_beam_CLT}")
acting_moment_beam_CLT = ULS_load_beam_CLT * floor_width ** 2 / 8

def find_closest_greater_value_beam_CLT(excel_file, sheet_name_CLT, search_column_CLT, search_value_CLT, target_column_CLT):
    df = pd.read_excel(excel_file, sheet_name=sheet_name_CLT)
    
    while True:
        df_filtered = df[df[search_column_CLT] >= search_value_CLT]
        
        if df_filtered.empty:
            return None
        
        index = df_filtered[search_column_CLT].idxmin()
        target_value = df.loc[index, target_column_CLT]
        
        load_ULS_CLT_loop = ULS_load_beam_CLT + 1.35 * target_value
        moment_ULS_loop_CLT = load_ULS_CLT_loop * floor_width ** 2 / 8
        
        if moment_ULS_loop_CLT == search_value_CLT:
            return moment_ULS_loop_CLT
        
        df[search_column_CLT] = df[search_column_CLT].astype(float)
        df.at[index, search_column_CLT] = moment_ULS_loop_CLT
        
        search_value_CLT = moment_ULS_loop_CLT

 
excel_file = r"C:\Users\runan\Documents\Masterproef\2024 April-mei\Structural materials.xlsx"
sheet_name_beam_steel = "Steel - Beam"
sheet_name_beam_concrete = "Concrete - Beam"
sheet_name_beam_GLT = "GLT - Beam"
search_column_beam_steel_MRd = "M_Rd [kNm]" 
search_column_beam_GLT_MRd = "M_Rd [kNm] ULS flexural"
initial_search_value_beam_CLT = acting_moment_beam_CLT 
target_column_self_weight = "Self weight [kN/m]" 

MEd_beam_steel_CLT = find_closest_greater_value_beam_CLT(excel_file, sheet_name_beam_steel, search_column_beam_steel_MRd, initial_search_value_beam_CLT, target_column_self_weight)

MEd_beam_GLT_CLT = find_closest_greater_value_beam_CLT(excel_file, sheet_name_beam_GLT, search_column_beam_GLT_MRd, initial_search_value_beam_CLT, target_column_self_weight)


area_beam_steel_CLT = find_corresponding_value_other_column(excel_file, sheet_name_beam_steel, search_column_beam_steel_MRd, MEd_beam_steel_CLT, "Area [mm²]")   
self_weight_beam_steel_CLT = find_corresponding_value_other_column(excel_file, sheet_name_beam_steel, search_column_beam_steel_MRd, MEd_beam_steel_CLT, "Self weight [kN/m]")  
profile_beam_steel_CLT = find_corresponding_value_other_column(excel_file, sheet_name_beam_steel, search_column_beam_steel_MRd, MEd_beam_steel_CLT, "Profile")   
height_beam_steel_CLT = find_corresponding_value_other_column(excel_file, sheet_name_beam_steel, search_column_beam_steel_MRd, MEd_beam_steel_CLT, "Height [mm]")  
print(profile_beam_steel_CLT)

area_beam_GLT_CLT = find_corresponding_value_other_column(excel_file, sheet_name_beam_GLT, search_column_beam_GLT_MRd, MEd_beam_GLT_CLT, "Area [mm²]")  
self_weight_beam_GLT_CLT = find_corresponding_value_other_column(excel_file, sheet_name_beam_GLT, search_column_beam_GLT_MRd, MEd_beam_GLT_CLT, "Self weight [kN/m]")  
profile_beam_GLT_CLT = find_corresponding_value_other_column(excel_file, sheet_name_beam_GLT, search_column_beam_GLT_MRd, MEd_beam_GLT_CLT, "Dimensions")  
height_beam_GLT_CLT = find_corresponding_value_other_column(excel_file, sheet_name_beam_GLT, search_column_beam_GLT_MRd, MEd_beam_GLT_CLT, "height [mm]")  
print(profile_beam_GLT_CLT)

Volume_beam_steel_CLT = area_beam_steel_CLT / (1000 * 1000) * floor_width
Volume_beam_GLT_CLT = area_beam_steel_CLT/(1000*1000)*floor_width





#BEAMS CONCRETE - CLT FLOOR
load_SLS_beam_CLT = load_SLS * floor_area/(2*floor_width) + self_weight_CLT_total/(2*floor_width) + additional_permanent_load_beams + additional_variable_load_beams

def find_closest_greater_value_beam_concrete_CLT(excel_file, sheet_name_concrete, check_threshold):
    df = pd.read_excel(excel_file, sheet_name=sheet_name_concrete)
    df = df.sort_values(by='I_y [mm^4]')
    
    for index, row in df.iterrows():
        secondmomentofarea_concrete = row['I_y [mm^4]']
        self_weight_beam_concrete = row['self_weight [kN/m]']
        area_beam_concrete = row['Area [mm²]']
        profile_beam_concrete = row['Dimensions']
        height_beam_concrete = row['height [mm]']
        
        deflection_beam_concrete = (5/384) * ((load_SLS_beam_CLT + self_weight_beam_concrete) * (floor_width * 1000)**4) / (E_concrete_beam * secondmomentofarea_concrete)
        
        if check_threshold >= deflection_beam_concrete:
            return deflection_beam_concrete, area_beam_concrete, profile_beam_concrete, self_weight_beam_concrete, height_beam_concrete
    
    print("End of DataFrame reached without finding a suitable result.")
    return None, None, None

excel_file = r"C:\Users\runan\Documents\Masterproef\2024 April-mei\Structural materials.xlsx"
sheet_name_beam_concrete = "Concrete - Beam"

deflection_beam_concrete_CLT, area_beam_concrete_CLT, profile_beam_concrete_CLT, self_weight_beam_concrete_CLT, height_beam_concrete_CLT = find_closest_greater_value_beam_concrete_CLT(excel_file, sheet_name_beam_concrete, limit_deflection_beam_concrete)
print(profile_beam_concrete_CLT)

Volume_beam_concrete_CLT = area_beam_concrete_CLT / (1000 * 1000) * floor_width





#BEAMS STEEL & GLT - CONCRETE FLOOR
ULS_load_beam_concrete =(1.35*self_weight_concrete_total + load_ULS*floor_area)/(2*floor_width) + 1.35*additional_permanent_load_beams + 1.50*additional_variable_load_beams
print(f"ULS_load_beam_concrete: {ULS_load_beam_concrete}")
acting_moment_beam_concrete = ULS_load_beam_concrete * floor_width ** 2 / 8

def find_closest_greater_value_beam_concrete(excel_file, sheet_name_concrete, search_column_concrete, search_value_concrete, target_column_concrete):
    df = pd.read_excel(excel_file, sheet_name=sheet_name_concrete)
    
    while True:
        df_filtered = df[df[search_column_concrete] >= search_value_concrete]
        
        if df_filtered.empty:
            return None
        
        index = df_filtered[search_column_concrete].idxmin()
        
        target_value = df.loc[index, target_column_concrete]
        
        load_ULS_concrete_loop = ULS_load_beam_concrete + 1.35 * target_value
        moment_ULS_loop_concrete = load_ULS_concrete_loop * floor_width ** 2 / 8
        
        if moment_ULS_loop_concrete == search_value_concrete:
            return moment_ULS_loop_concrete
        
        df[search_column_concrete] = df[search_column_concrete].astype(float)  
        df.at[index, search_column_concrete] = moment_ULS_loop_concrete
        
        search_value_concrete = moment_ULS_loop_concrete

 
excel_file = r"C:\Users\runan\Documents\Masterproef\2024 April-mei\Structural materials.xlsx"
sheet_name_beam_steel = "Steel - Beam"
sheet_name_beam_GLT = "GLT - Beam"
search_column_beam_steel_MRd = "M_Rd [kNm]" 
search_column_beam_GLT_MRd = "M_Rd [kNm] ULS flexural"
initial_search_value_beam_concrete = acting_moment_beam_concrete  
target_column_self_weight = "Self weight [kN/m]"  

MEd_beam_steel_concrete = find_closest_greater_value_beam_concrete(excel_file, sheet_name_beam_steel, search_column_beam_steel_MRd, initial_search_value_beam_concrete, target_column_self_weight)

MEd_beam_GLT_concrete = find_closest_greater_value_beam_concrete(excel_file, sheet_name_beam_GLT, search_column_beam_GLT_MRd, initial_search_value_beam_concrete, target_column_self_weight)


area_beam_steel_concrete = find_corresponding_value_other_column(excel_file, sheet_name_beam_steel, search_column_beam_steel_MRd, MEd_beam_steel_concrete, "Area [mm²]")   
self_weight_beam_steel_concrete = find_corresponding_value_other_column(excel_file, sheet_name_beam_steel, search_column_beam_steel_MRd, MEd_beam_steel_concrete, "Self weight [kN/m]")  
profile_beam_steel_concrete = find_corresponding_value_other_column(excel_file, sheet_name_beam_steel, search_column_beam_steel_MRd, MEd_beam_steel_concrete, "Profile")   
height_beam_steel_concrete = find_corresponding_value_other_column(excel_file, sheet_name_beam_steel, search_column_beam_steel_MRd, MEd_beam_steel_concrete, "Height [mm]")  
print(profile_beam_steel_concrete)

area_beam_GLT_concrete = find_corresponding_value_other_column(excel_file, sheet_name_beam_GLT, search_column_beam_GLT_MRd, MEd_beam_GLT_concrete, "Area [mm²]")  
self_weight_beam_GLT_concrete = find_corresponding_value_other_column(excel_file, sheet_name_beam_GLT, search_column_beam_GLT_MRd, MEd_beam_GLT_concrete, "Self weight [kN/m]")  
profile_beam_GLT_concrete = find_corresponding_value_other_column(excel_file, sheet_name_beam_GLT, search_column_beam_GLT_MRd, MEd_beam_GLT_concrete, "Dimensions")  
height_beam_GLT_concrete = find_corresponding_value_other_column(excel_file, sheet_name_beam_GLT, search_column_beam_GLT_MRd, MEd_beam_GLT_concrete, "height [mm]")  
print(profile_beam_GLT_concrete)

Volume_beam_steel_concrete = area_beam_steel_concrete / (1000 * 1000) * floor_width
Volume_beam_GLT_concrete = area_beam_steel_concrete/(1000*1000)*floor_width




#BEAMS CONCRETE - CONCRETE FLOOR
load_SLS_beam_concrete = load_SLS * floor_area/(2*floor_width) + self_weight_concrete_total/(2*floor_width) + additional_permanent_load_beams + additional_variable_load_beams

def find_closest_greater_value_beam_concrete_concrete(excel_file, sheet_name_concrete, check_threshold):
    df = pd.read_excel(excel_file, sheet_name=sheet_name_concrete)
    df = df.sort_values(by='I_y [mm^4]')
    
    for index, row in df.iterrows():
        secondmomentofarea_concrete = row['I_y [mm^4]']
        self_weight_beam_concrete = row['self_weight [kN/m]']
        area_beam_concrete = row['Area [mm²]']
        profile_beam_concrete = row['Dimensions']
        height_beam_concrete = row['height [mm]']
        
        deflection_beam_concrete = (5/384) * ((load_SLS_beam_concrete + self_weight_beam_concrete) * (floor_width * 1000)**4) / (E_concrete_beam * secondmomentofarea_concrete)
        
        if check_threshold >= deflection_beam_concrete:
            return deflection_beam_concrete, area_beam_concrete, profile_beam_concrete, self_weight_beam_concrete, height_beam_concrete
    
    print("End of DataFrame reached without finding a suitable result.")
    return None, None, None

excel_file = r"C:\Users\runan\Documents\Masterproef\2024 April-mei\Structural materials.xlsx"
sheet_name_beam_concrete = "Concrete - Beam"

deflection_beam_concrete_concrete, area_beam_concrete_concrete, profile_beam_concrete_concrete, self_weight_beam_concrete_concrete, height_beam_concrete_concrete = find_closest_greater_value_beam_concrete_concrete(excel_file, sheet_name_beam_concrete, limit_deflection_beam_concrete)
print(profile_beam_concrete_concrete)

Volume_beam_concrete_concrete = area_beam_concrete_concrete / (1000 * 1000) * floor_width



pi = math.pi

def adjust_value(value):
    if value > 1:
        return 1
    else:
        return value


#STEEL COLUMN - HCS FLOOR
spread_load_columns_steel_HCS = 1.35 * self_weight_beam_steel_HCS + ULS_load_beam_HCS
point_load_column_steel_HCS = (spread_load_columns_steel_HCS * floor_width / 2) + 1.35*additional_permanent_load_columns + 1.50*additional_variable_load_columns
print(f"point_load_column_steel_HCS: {point_load_column_steel_HCS}")

column_height_steel_HCS = structure_height - (height_beam_steel_HCS / 1000) - (thickness_HCS/1000)

def find_closest_greater_value_column_steel(excel_file, sheet_name_steel, check_column, check_threshold):
    df = pd.read_excel(excel_file, sheet_name=sheet_name_steel)
    df[check_column] = df[check_column].astype(float)
    
    for index, row in df.iterrows():
        secondmomentofarea_steel = row['I_z [mm^4]']
        area_steel = row['Area [mm²]']
        fy_steel = row['Steel grade']
        alpha_steel = row['alpha']
        elasticitymodulus_steel = row['E [N/mm²]']
        resistance_steel = row['N_Rd [kN]']
        
        N_cr = ((elasticitymodulus_steel * secondmomentofarea_steel * (pi**2)) / ((column_height_steel_HCS * 1000)**2))
        lambdastreep = ((area_steel * fy_steel) / N_cr)**(1/2)
        phi = 0.5 * (1 + alpha_steel * (lambdastreep - 0.2) + lambdastreep**2)
        chi = 1 / (phi + (phi**2 - lambdastreep**2)**(0.5))
        chi_adjusted = adjust_value(chi)
        new_resistance_steel = resistance_steel * chi_adjusted
        
        if new_resistance_steel >= point_load_column_steel_HCS:
            return new_resistance_steel, resistance_steel, row['N_Rd [kN]']
    
    print("End of DataFrame reached without finding a suitable result.")
    return None, None, None

sheet_name_column_steel = "Steel - Column"
compression_column_steel = "N_Rd [kN]"   

new_resistance_steel_HCS, resistance_steel_HCS, matching_N_Rd_steel_HCS = find_closest_greater_value_column_steel(excel_file, sheet_name_column_steel, compression_column_steel, point_load_column_steel_HCS)
area_column_steel_HCS = find_corresponding_value_other_column(excel_file, sheet_name_column_steel, compression_column_steel, resistance_steel_HCS, "Area [mm²]")  
profile_column_steel_HCS = find_corresponding_value_other_column(excel_file, sheet_name_column_steel, compression_column_steel, resistance_steel_HCS, "Profile")  
print(profile_column_steel_HCS) 

Volume_column_steel_HCS = area_column_steel_HCS/(1000*1000) * column_height_steel_HCS





#GLT COLUMN - HCS FLOOR
spread_load_columns_GLT_HCS = 1.35 * self_weight_beam_GLT_HCS + ULS_load_beam_HCS
point_load_column_GLT_HCS = (spread_load_columns_GLT_HCS*floor_width/2) + 1.35*additional_permanent_load_columns + 1.50*additional_variable_load_columns
print(f"point_load_column_GLT_HCS: {point_load_column_GLT_HCS}")

column_height_GLT_HCS = structure_height - (height_beam_GLT_HCS/1000) - (thickness_HCS/1000)

def find_closest_greater_value_column_GLT(excel_file, sheet_name_GLT, check_column, check_threshold):
    df = pd.read_excel(excel_file, sheet_name=sheet_name_GLT)
    df[check_column] = df[check_column].astype(float)
    
    for index, row in df.iterrows():
        secondmomentofarea_GLT_HCS = row['I_z [mm^4]']
        area_GLT_HCS = row['Area [mm²]']
        fy_GLT_HCS = row['f_c,0,d [N/mm²]']
        alpha_GLT_HCS = row['alpha_buckling']
        elasticitymodulus_GLT_HCS = row['E [N/mm²]']
        resistance_GLT_HCS = row['N_Rd [kN]']
        
        N_cr = ((elasticitymodulus_GLT_HCS * secondmomentofarea_GLT_HCS * (pi**2)) / ((column_height_GLT_HCS * 1000)**2))
        lambdastreep = ((area_GLT_HCS * fy_GLT_HCS) / N_cr)**(1/2)
        phi = 0.5 * (1 + alpha_GLT_HCS * (lambdastreep - 0.2) + lambdastreep**2)
        chi = 1 / (phi + (phi**2 - lambdastreep**2)**(0.5))
        chi_adjusted = adjust_value(chi)
        new_resistance_GLT_HCS = resistance_GLT_HCS * chi_adjusted
        
        if new_resistance_GLT_HCS >= point_load_column_GLT_HCS:
            return new_resistance_GLT_HCS, resistance_GLT_HCS, row['N_Rd [kN]']
    
    print("End of DataFrame reached without finding a suitable result.")
    return None, None, None

sheet_name_column_GLT = "GLT - Column"
compression_column_GLT = "N_Rd [kN]"

new_resistance_GLT_HCS, resistance_GLT_HCS, matching_N_Rd_GLT_HCS = find_closest_greater_value_column_GLT(excel_file, sheet_name_column_GLT, compression_column_GLT, point_load_column_GLT_HCS)
area_column_GLT_HCS = find_corresponding_value_other_column(excel_file, sheet_name_column_GLT, compression_column_GLT, resistance_GLT_HCS, "Area [mm²]")  
profile_column_GLT_HCS = find_corresponding_value_other_column(excel_file, sheet_name_column_GLT, compression_column_GLT, resistance_GLT_HCS, "Dimensions")  
print(profile_column_GLT_HCS) 

Volume_column_GLT_HCS = area_column_GLT_HCS/(1000*1000) * column_height_GLT_HCS





#CONCRETE COLUMN - HCS FLOOR
spread_load_columns_concrete_HCS = 1.35 * self_weight_beam_concrete_HCS + ULS_load_beam_HCS
point_load_column_concrete_HCS = (spread_load_columns_concrete_HCS * floor_width / 2) + 1.35*additional_permanent_load_columns + 1.50*additional_variable_load_columns
print(f"point_load_column_concrete_HCS: {point_load_column_concrete_HCS}")

column_height_concrete_HCS = structure_height - (height_beam_concrete_HCS / 1000) - (thickness_HCS/1000)

def find_closest_greater_value_column_concrete(excel_file, sheet_name_concrete, check_column, check_threshold):
    df = pd.read_excel(excel_file, sheet_name=sheet_name_concrete)
    df[check_column] = df[check_column].astype(float)
    
    for index, row in df.iterrows():
        secondmomentofarea_concrete = row['I_z [mm^4]']
        area_concrete = row['Area [mm²]']
        fy_concrete = row['f_combined [N/mm²]']
        alpha_concrete = row['alpha_buckling']
        elasticitymodulus_concrete = row['E [N/mm²]']
        resistance_concrete = row['N_Rd [kN]']
        
        N_cr = ((elasticitymodulus_concrete * secondmomentofarea_concrete * (pi**2)) / ((column_height_concrete_HCS * 1000)**2))
        lambdastreep = ((area_concrete * fy_concrete) / N_cr)**(1/2)
        phi = 0.5 * (1 + alpha_concrete * (lambdastreep - 0.2) + lambdastreep**2)
        chi = 1 / (phi + (phi**2 - lambdastreep**2)**(0.5))
        chi_adjusted = adjust_value(chi)
        new_resistance_concrete = resistance_concrete * chi_adjusted
        
        if new_resistance_concrete >= point_load_column_concrete_HCS:
            return new_resistance_concrete, resistance_concrete, row['N_Rd [kN]']
    
    print("End of DataFrame reached without finding a suitable result.")
    return None, None, None

sheet_name_column_concrete = "Concrete - Column"
compression_column_concrete = "N_Rd [kN]"   

new_resistance_concrete_HCS, resistance_concrete_HCS, matching_N_Rd_concrete_HCS = find_closest_greater_value_column_concrete(excel_file, sheet_name_column_concrete, compression_column_concrete, point_load_column_concrete_HCS)
area_column_concrete_HCS = find_corresponding_value_other_column(excel_file, sheet_name_column_concrete, compression_column_concrete, resistance_concrete_HCS, "Area [mm²]")  
profile_column_concrete_HCS = find_corresponding_value_other_column(excel_file, sheet_name_column_concrete, compression_column_concrete, resistance_concrete_HCS, "Dimensions")  
print(profile_column_concrete_HCS) 

Volume_column_concrete_HCS = area_column_concrete_HCS/(1000*1000) * column_height_concrete_HCS





#STEEL COLUMN - CLT FLOOR
spread_load_columns_steel_CLT = 1.35 * self_weight_beam_steel_CLT + ULS_load_beam_CLT
point_load_column_steel_CLT = (spread_load_columns_steel_CLT * floor_width / 2) + 1.35*additional_permanent_load_columns + 1.50*additional_variable_load_columns
print(f"point_load_column_steel_CLT: {point_load_column_steel_CLT}")

column_height_steel_CLT = structure_height - (height_beam_steel_CLT / 1000) - (thickness_CLT/1000) 

new_resistance_steel_CLT, resistance_steel_CLT, matching_N_Rd_steel_CLT = find_closest_greater_value_column_steel(excel_file, sheet_name_column_steel, compression_column_steel, point_load_column_steel_CLT)

area_column_steel_CLT = find_corresponding_value_other_column(excel_file, sheet_name_column_steel, compression_column_steel, resistance_steel_CLT, "Area [mm²]")  
profile_column_steel_CLT = find_corresponding_value_other_column(excel_file, sheet_name_column_steel, compression_column_steel, resistance_steel_CLT, "Profile")  
print(profile_column_steel_CLT) 

Volume_column_steel_CLT = area_column_steel_CLT/(1000*1000) * column_height_steel_CLT




#GLT COLUMN - CLT FLOOR
spread_load_columns_GLT_CLT = 1.35 * self_weight_beam_GLT_CLT + ULS_load_beam_CLT
point_load_column_GLT_CLT = (spread_load_columns_GLT_CLT*floor_width/2) + 1.35*additional_permanent_load_columns + 1.50*additional_variable_load_columns
print(f"point_load_column_GLT_CLT: {point_load_column_GLT_CLT}")

column_height_GLT_CLT = structure_height - (height_beam_GLT_CLT/1000) - (thickness_CLT/1000)

new_resistance_GLT_CLT, resistance_GLT_CLT, matching_N_Rd_GLT_CLT = find_closest_greater_value_column_GLT(excel_file, sheet_name_column_GLT, compression_column_GLT, point_load_column_GLT_CLT)

area_column_GLT_CLT = find_corresponding_value_other_column(excel_file, sheet_name_column_GLT, compression_column_GLT, resistance_GLT_CLT, "Area [mm²]")    
profile_column_GLT_CLT = find_corresponding_value_other_column(excel_file, sheet_name_column_GLT, compression_column_GLT, resistance_GLT_CLT, "Dimensions")  
print(profile_column_GLT_CLT) 

Volume_column_GLT_CLT = area_column_GLT_CLT/(1000*1000) * column_height_GLT_CLT




#CONCRETE COLUMN - CLT FLOOR
spread_load_columns_concrete_CLT = 1.35 * self_weight_beam_concrete_CLT + ULS_load_beam_CLT
point_load_column_concrete_CLT = (spread_load_columns_concrete_CLT * floor_width / 2) + 1.35*additional_permanent_load_columns + 1.50*additional_variable_load_columns
print(f"point_load_column_concrete_CLT: {point_load_column_concrete_CLT}")

column_height_concrete_CLT = structure_height - (height_beam_concrete_CLT / 1000) - (thickness_CLT/1000)

new_resistance_concrete_CLT, resistance_concrete_CLT, matching_N_Rd_concrete_CLT = find_closest_greater_value_column_concrete(excel_file, sheet_name_column_concrete, compression_column_concrete, point_load_column_concrete_CLT)

area_column_concrete_CLT = find_corresponding_value_other_column(excel_file, sheet_name_column_concrete, compression_column_concrete, resistance_concrete_CLT, "Area [mm²]")  
profile_column_concrete_CLT = find_corresponding_value_other_column(excel_file, sheet_name_column_concrete, compression_column_concrete, resistance_concrete_CLT, "Dimensions")  
print(profile_column_concrete_CLT) 

Volume_column_concrete_CLT = area_column_concrete_CLT/(1000*1000) * column_height_concrete_CLT




#STEEL COLUMN - CONCRETE FLOOR
spread_load_columns_steel_concrete = 1.35 * self_weight_beam_steel_concrete + ULS_load_beam_concrete
point_load_column_steel_concrete = (spread_load_columns_steel_concrete * floor_width / 2) + 1.35*additional_permanent_load_columns + 1.50*additional_variable_load_columns
print(f"point_load_column_steel_concrete: {point_load_column_steel_concrete}")

column_height_steel_concrete = structure_height - (height_beam_steel_concrete / 1000) - (thickness_concrete_rounded/1000) 

new_resistance_steel_concrete, resistance_steel_concrete, matching_N_Rd_steel_concrete = find_closest_greater_value_column_steel(excel_file, sheet_name_column_steel, compression_column_steel, point_load_column_steel_concrete)

area_column_steel_concrete = find_corresponding_value_other_column(excel_file, sheet_name_column_steel, compression_column_steel, resistance_steel_concrete, "Area [mm²]")  
profile_column_steel_concrete = find_corresponding_value_other_column(excel_file, sheet_name_column_steel, compression_column_steel, resistance_steel_concrete, "Profile")  
print(profile_column_steel_concrete) 

Volume_column_steel_concrete = area_column_steel_concrete/(1000*1000) * column_height_steel_concrete




#GLT COLUMN - CONCRETE FLOOR
spread_load_columns_GLT_concrete = 1.35 * self_weight_beam_GLT_concrete + ULS_load_beam_concrete
point_load_column_GLT_concrete = (spread_load_columns_GLT_concrete*floor_width/2) + 1.35*additional_permanent_load_columns + 1.50*additional_variable_load_columns
print(f"point_load_column_GLT_concrete: {point_load_column_GLT_concrete}")

column_height_GLT_concrete = structure_height - (height_beam_GLT_concrete/1000) - (thickness_concrete_rounded/1000)

new_resistance_GLT_concrete, resistance_GLT_concrete, matching_N_Rd_GLT_concrete = find_closest_greater_value_column_GLT(excel_file, sheet_name_column_GLT, compression_column_GLT, point_load_column_GLT_concrete)

area_column_GLT_concrete = find_corresponding_value_other_column(excel_file, sheet_name_column_GLT, compression_column_GLT, resistance_GLT_concrete, "Area [mm²]")    
profile_column_GLT_concrete = find_corresponding_value_other_column(excel_file, sheet_name_column_GLT, compression_column_GLT, resistance_GLT_concrete, "Dimensions")  
print(profile_column_GLT_concrete) 

Volume_column_GLT_concrete = area_column_GLT_concrete/(1000*1000) * column_height_GLT_concrete




#CONCRETE COLUMN - CONCRETE FLOOR
spread_load_columns_concrete_concrete = 1.35 * self_weight_beam_concrete_concrete + ULS_load_beam_concrete
point_load_column_concrete_concrete = (spread_load_columns_concrete_concrete * floor_width / 2) + 1.35*additional_permanent_load_columns + 1.50*additional_variable_load_columns
print(f"point_load_column_concrete_concrete: {point_load_column_concrete_concrete}")

column_height_concrete_concrete = structure_height - (height_beam_concrete_concrete / 1000) - (thickness_concrete_rounded/1000)

new_resistance_concrete_concrete, resistance_concrete_concrete, matching_N_Rd_concrete_concrete = find_closest_greater_value_column_concrete(excel_file, sheet_name_column_concrete, compression_column_concrete, point_load_column_concrete_concrete)

area_column_concrete_concrete = find_corresponding_value_other_column(excel_file, sheet_name_column_concrete, compression_column_concrete, resistance_concrete_concrete, "Area [mm²]")  
profile_column_concrete_concrete = find_corresponding_value_other_column(excel_file, sheet_name_column_concrete, compression_column_concrete, resistance_concrete_concrete, "Dimensions")  
print(profile_column_concrete_concrete) 

Volume_column_concrete_concrete = area_column_concrete_concrete/(1000*1000) * column_height_concrete_concrete




#CALCULATE IMPACT
def retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_impact_category, column_name_material):
    df = pd.read_excel(excel_file_LCA, sheet_name=sheet_name_LCA, index_col=0)
    
    try:
        value = df.loc[row_name_impact_category, column_name_material]
        return value
    except KeyError:
        print("Row or column not found.")
        return None

excel_file_LCA = r"C:\Users\runan\Documents\Masterproef\2024 April-mei\SimaPro LCA database.xlsx"
sheet_name_LCA = "EN 15804+A2"
row_name_milipoint = "milipoint"
row_name_GWP = "Climate change"
row_name_OzoneD = "Ozone depletion"
row_name_IonisingR = "Ionising radiation"
row_name_photochemicalOF = "Photochemical ozone formation"
row_name_ParticulatM = "Particulate matter"
row_name_HumanTNC = "Human toxicity, non-cancer"
row_name_HumanTC = "Human toxicity, cancer"
row_name_Acidification = "Acidification"
row_name_EutrophicationF = "Eutrophication, freshwater"
row_name_EutrophicationM = "Eutrophication, marine"
row_name_EutrophicationT = "Eutrophication, terrestrial"
row_name_EcotoxicityF = "Ecotoxicity, freshwater"
row_name_LandUse = "Land use"
row_name_WaterUse = "Water use"
row_name_ResourceUseF = "Resource use, fossils"
row_name_ResourceUseMM = "Resource use, minerals and metals"
row_name_ClimateCF = "Climate change - Fossil"
row_name_ClimateCB = "Climate change - Biogenic"
row_name_ClimateCLU = "Climate change - Land use and LU change"
row_name_HumanTNCO = "Human toxicity, non-cancer - organics"
row_name_HumanTNCI = "Human toxicity, non-cancer - inorganics"
row_name_HumanTNCM = "Human toxicity, non-cancer - metals"
row_name_HumanTCO = "Human toxicity, cancer - organics"
row_name_HumanTCI = "Human toxicity, cancer - inorganics"
row_name_HumanTCM = "Human toxicity, cancer - metals"
row_name_ExotoxicityFO = "Ecotoxicity, freshwater - organics"
row_name_ExotoxicityFI = "Ecotoxicity, freshwater - inorganics"
row_name_ExotoxicityFM = "Ecotoxicity, freshwater - metals"
column_name_concrete_C50_60 = "Concrete C50/60 [1m³]"
column_name_concrete_C30_37 = "Concrete C30/37 [1m³]"
column_name_reinforcement_steel = "Reinforcement steel [1m³]"
column_name_CLT = "Cross laminated timber [1m³]"
column_name_GLT = "Glued laminated timber [1m³]"
column_name_structural_steel = "Structural steel [1m³]"


mPt_concrete_C50_60 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_milipoint, column_name_concrete_C50_60)
mPt_concrete_C30_37 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_milipoint, column_name_concrete_C30_37)
mPt_reinforcement =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_milipoint, column_name_reinforcement_steel)
mPt_CLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_milipoint, column_name_CLT)
mPt_GLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_milipoint, column_name_GLT)
mPt_structural_steel =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_milipoint, column_name_structural_steel)

mPt_score_HCS = mPt_concrete_C30_37*Volume_compression_layer_HCS + mPt_concrete_C50_60*Volume_concrete_floor_HCS + mPt_reinforcement*Volume_steel_floor_HCS
mPt_score_in_situ_concrete = mPt_concrete_C30_37*Volume_in_situ_concrete + mPt_reinforcement*Volume_reinforcement_in_situ_concrete
mPt_score_CLT = mPt_CLT*Volume_CLT
mPt_score_support_steel_HCS = (Volume_beam_steel_HCS * mPt_structural_steel)*2 + (Volume_column_steel_HCS * mPt_structural_steel)*4
mPt_score_support_GLT_HCS = (Volume_beam_GLT_HCS * mPt_GLT)*2 + (Volume_column_GLT_HCS * mPt_GLT)*4
mPt_score_support_concrete_HCS = (Volume_beam_concrete_HCS * 0.965 * mPt_concrete_C30_37 + Volume_beam_concrete_HCS * 0.035 * mPt_reinforcement)*2 + (Volume_column_concrete_HCS * 0.97 * mPt_concrete_C30_37 + Volume_column_concrete_HCS * 0.03 * mPt_reinforcement)*4
mPt_score_support_steel_CLT = (Volume_beam_steel_CLT * mPt_structural_steel)*2 + (Volume_column_steel_CLT * mPt_structural_steel)*4
mPt_score_support_GLT_CLT = (Volume_beam_GLT_CLT * mPt_GLT)*2 + (Volume_column_GLT_CLT * mPt_GLT)*4
mPt_score_support_concrete_CLT = (Volume_beam_concrete_CLT * 0.965 * mPt_concrete_C30_37 + Volume_beam_concrete_CLT * 0.035 * mPt_reinforcement)*2 + (Volume_column_concrete_CLT * 0.97 * mPt_concrete_C30_37 + Volume_column_concrete_CLT * 0.03 * mPt_reinforcement)*4
mPt_score_support_steel_concrete = (Volume_beam_steel_concrete * mPt_structural_steel)*2 + (Volume_column_steel_concrete * mPt_structural_steel)*4
mPt_score_support_GLT_concrete = (Volume_beam_GLT_concrete * mPt_GLT)*2 + (Volume_column_GLT_concrete * mPt_GLT)*4
mPt_score_support_concrete_concrete = (Volume_beam_concrete_concrete * 0.965 * mPt_concrete_C30_37 + Volume_beam_concrete_concrete * 0.035 * mPt_reinforcement)*2 + (Volume_column_concrete_concrete * 0.97 * mPt_concrete_C30_37 + Volume_column_concrete_concrete * 0.03 * mPt_reinforcement)*4

GWP_concrete_C50_60 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_GWP, column_name_concrete_C50_60)
GWP_concrete_C30_37 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_GWP, column_name_concrete_C30_37)
GWP_reinforcement =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_GWP, column_name_reinforcement_steel)
GWP_CLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_GWP, column_name_CLT)
GWP_GLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_GWP, column_name_GLT)
GWP_structural_steel =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_GWP, column_name_structural_steel)

GWP_score_HCS = GWP_concrete_C50_60*Volume_concrete_floor_HCS + GWP_concrete_C30_37*Volume_compression_layer_HCS + GWP_reinforcement*Volume_steel_floor_HCS
GWP_score_in_situ_concrete = GWP_concrete_C30_37*Volume_in_situ_concrete + GWP_reinforcement* Volume_reinforcement_in_situ_concrete
GWP_score_CLT = GWP_CLT*Volume_CLT
GWP_score_support_steel_HCS = (Volume_beam_steel_HCS * GWP_structural_steel)*2 + (Volume_column_steel_HCS * GWP_structural_steel)*4
GWP_score_support_GLT_HCS = (Volume_beam_GLT_HCS * GWP_GLT)*2 + (Volume_column_GLT_HCS * GWP_GLT)*4
GWP_score_support_concrete_HCS = (Volume_beam_concrete_HCS * 0.965 * GWP_concrete_C30_37 + Volume_beam_concrete_HCS * 0.035 * GWP_reinforcement)*2 + (Volume_column_concrete_HCS * 0.97 * GWP_concrete_C30_37 + Volume_column_concrete_HCS * 0.03 * GWP_reinforcement)*4
GWP_score_support_steel_CLT = (Volume_beam_steel_CLT * GWP_structural_steel)*2 + (Volume_column_steel_CLT * GWP_structural_steel)*4
GWP_score_support_GLT_CLT = (Volume_beam_GLT_CLT * GWP_GLT)*2 + (Volume_column_GLT_CLT * GWP_GLT)*4
GWP_score_support_concrete_CLT = (Volume_beam_concrete_CLT * 0.965 * GWP_concrete_C30_37 + Volume_beam_concrete_CLT * 0.035 * GWP_reinforcement)*2 + (Volume_column_concrete_CLT * 0.97 * GWP_concrete_C30_37 + Volume_column_concrete_CLT * 0.03 * GWP_reinforcement)*4
GWP_score_support_steel_concrete = (Volume_beam_steel_concrete * GWP_structural_steel)*2 + (Volume_column_steel_concrete * GWP_structural_steel)*4
GWP_score_support_GLT_concrete = (Volume_beam_GLT_concrete * GWP_GLT)*2 + (Volume_column_GLT_concrete * GWP_GLT)*4
GWP_score_support_concrete_concrete = (Volume_beam_concrete_concrete * 0.965 * GWP_concrete_C30_37 + Volume_beam_concrete_concrete * 0.035 * GWP_reinforcement)*2 + (Volume_column_concrete_concrete * 0.97 * GWP_concrete_C30_37 + Volume_column_concrete_concrete * 0.03 * GWP_reinforcement)*4

OzoneD_concrete_C50_60 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_OzoneD, column_name_concrete_C50_60)
OzoneD_concrete_C30_37 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_OzoneD, column_name_concrete_C30_37)
OzoneD_reinforcement =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_OzoneD, column_name_reinforcement_steel)
OzoneD_CLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_OzoneD, column_name_CLT)
OzoneD_GLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_OzoneD, column_name_GLT)
OzoneD_structural_steel =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_OzoneD, column_name_structural_steel)

OzoneD_score_HCS = OzoneD_concrete_C50_60*Volume_concrete_floor_HCS + OzoneD_concrete_C30_37*Volume_compression_layer_HCS + OzoneD_reinforcement*Volume_steel_floor_HCS
OzoneD_score_in_situ_concrete = OzoneD_concrete_C30_37*Volume_in_situ_concrete + OzoneD_reinforcement * Volume_reinforcement_in_situ_concrete
OzoneD_score_CLT = OzoneD_CLT*Volume_CLT
OzoneD_score_support_steel_HCS = (Volume_beam_steel_HCS * OzoneD_structural_steel)*2 + (Volume_column_steel_HCS * OzoneD_structural_steel)*4
OzoneD_score_support_GLT_HCS = (Volume_beam_GLT_HCS * OzoneD_GLT)*2 + (Volume_column_GLT_HCS * OzoneD_GLT)*4
OzoneD_score_support_concrete_HCS = (Volume_beam_concrete_HCS * 0.965 * OzoneD_concrete_C30_37 + Volume_beam_concrete_HCS * 0.035 * OzoneD_reinforcement)*2 + (Volume_column_concrete_HCS * 0.97 * OzoneD_concrete_C30_37 + Volume_column_concrete_HCS * 0.03 * OzoneD_reinforcement)*4
OzoneD_score_support_steel_CLT = (Volume_beam_steel_CLT * OzoneD_structural_steel)*2 + (Volume_column_steel_CLT * OzoneD_structural_steel)*4
OzoneD_score_support_GLT_CLT = (Volume_beam_GLT_CLT * OzoneD_GLT)*2 + (Volume_column_GLT_CLT * OzoneD_GLT)*4
OzoneD_score_support_concrete_CLT = (Volume_beam_concrete_CLT * 0.965 * OzoneD_concrete_C30_37 + Volume_beam_concrete_CLT * 0.035 * OzoneD_reinforcement)*2 + (Volume_column_concrete_CLT * 0.97 * OzoneD_concrete_C30_37 + Volume_column_concrete_CLT * 0.03 * OzoneD_reinforcement)*4
OzoneD_score_support_steel_concrete = (Volume_beam_steel_concrete * OzoneD_structural_steel)*2 + (Volume_column_steel_concrete * OzoneD_structural_steel)*4
OzoneD_score_support_GLT_concrete = (Volume_beam_GLT_concrete * OzoneD_GLT)*2 + (Volume_column_GLT_concrete * OzoneD_GLT)*4
OzoneD_score_support_concrete_concrete = (Volume_beam_concrete_concrete * 0.965 * OzoneD_concrete_C30_37 + Volume_beam_concrete_concrete * 0.035 * OzoneD_reinforcement)*2 + (Volume_column_concrete_concrete * 0.97 * OzoneD_concrete_C30_37 + Volume_column_concrete_concrete * 0.03 * OzoneD_reinforcement)*4

IonisingR_concrete_C50_60 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_IonisingR, column_name_concrete_C50_60)
IonisingR_concrete_C30_37 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_IonisingR, column_name_concrete_C30_37)
IonisingR_reinforcement =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_IonisingR, column_name_reinforcement_steel)
IonisingR_CLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_IonisingR, column_name_CLT)
IonisingR_GLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_IonisingR, column_name_GLT)
IonisingR_structural_steel =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_IonisingR, column_name_structural_steel)

IonisingR_score_HCS = IonisingR_concrete_C50_60*Volume_concrete_floor_HCS + IonisingR_concrete_C30_37*Volume_compression_layer_HCS + IonisingR_reinforcement*Volume_steel_floor_HCS
IonisingR_score_in_situ_concrete = IonisingR_concrete_C30_37*Volume_in_situ_concrete + IonisingR_reinforcement* Volume_reinforcement_in_situ_concrete
IonisingR_score_CLT = IonisingR_CLT*Volume_CLT
IonisingR_score_support_steel_HCS = (Volume_beam_steel_HCS * IonisingR_structural_steel)*2 + (Volume_column_steel_HCS * IonisingR_structural_steel)*4
IonisingR_score_support_GLT_HCS = (Volume_beam_GLT_HCS * IonisingR_GLT)*2 + (Volume_column_GLT_HCS * IonisingR_GLT)*4
IonisingR_score_support_concrete_HCS = (Volume_beam_concrete_HCS * 0.965 * IonisingR_concrete_C30_37 + Volume_beam_concrete_HCS * 0.035 * IonisingR_reinforcement)*2 + (Volume_column_concrete_HCS * 0.97 * IonisingR_concrete_C30_37 + Volume_column_concrete_HCS * 0.03 * IonisingR_reinforcement)*4
IonisingR_score_support_steel_CLT = (Volume_beam_steel_CLT * IonisingR_structural_steel)*2 + (Volume_column_steel_CLT * IonisingR_structural_steel)*4
IonisingR_score_support_GLT_CLT = (Volume_beam_GLT_CLT * IonisingR_GLT)*2 + (Volume_column_GLT_CLT * IonisingR_GLT)*4
IonisingR_score_support_concrete_CLT = (Volume_beam_concrete_CLT * 0.965 * IonisingR_concrete_C30_37 + Volume_beam_concrete_CLT * 0.035 * IonisingR_reinforcement)*2 + (Volume_column_concrete_CLT * 0.97 * IonisingR_concrete_C30_37 + Volume_column_concrete_CLT * 0.03 * IonisingR_reinforcement)*4
IonisingR_score_support_steel_concrete = (Volume_beam_steel_concrete * IonisingR_structural_steel)*2 + (Volume_column_steel_concrete * IonisingR_structural_steel)*4
IonisingR_score_support_GLT_concrete = (Volume_beam_GLT_concrete * IonisingR_GLT)*2 + (Volume_column_GLT_concrete * IonisingR_GLT)*4
IonisingR_score_support_concrete_concrete = (Volume_beam_concrete_concrete * 0.965 * IonisingR_concrete_C30_37 + Volume_beam_concrete_concrete * 0.035 * IonisingR_reinforcement)*2 + (Volume_column_concrete_concrete * 0.97 * IonisingR_concrete_C30_37 + Volume_column_concrete_concrete * 0.03 * IonisingR_reinforcement)*4

photochemicalOF_concrete_C50_60 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_photochemicalOF, column_name_concrete_C50_60)
photochemicalOF_concrete_C30_37 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_photochemicalOF, column_name_concrete_C30_37)
photochemicalOF_reinforcement =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_photochemicalOF, column_name_reinforcement_steel)
photochemicalOF_CLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_photochemicalOF, column_name_CLT)
photochemicalOF_GLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_photochemicalOF, column_name_GLT)
photochemicalOF_structural_steel =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_photochemicalOF, column_name_structural_steel)

photochemicalOF_score_HCS = photochemicalOF_concrete_C50_60*Volume_concrete_floor_HCS + photochemicalOF_concrete_C30_37*Volume_compression_layer_HCS + photochemicalOF_reinforcement*Volume_steel_floor_HCS
photochemicalOF_score_in_situ_concrete = photochemicalOF_concrete_C30_37*Volume_in_situ_concrete + photochemicalOF_reinforcement* Volume_reinforcement_in_situ_concrete
photochemicalOF_score_CLT = photochemicalOF_CLT*Volume_CLT
photochemicalOF_score_support_steel_HCS = (Volume_beam_steel_HCS * photochemicalOF_structural_steel)*2 + (Volume_column_steel_HCS * photochemicalOF_structural_steel)*4
photochemicalOF_score_support_GLT_HCS = (Volume_beam_GLT_HCS * photochemicalOF_GLT)*2 + (Volume_column_GLT_HCS * photochemicalOF_GLT)*4
photochemicalOF_score_support_concrete_HCS = (Volume_beam_concrete_HCS * 0.965 * photochemicalOF_concrete_C30_37 + Volume_beam_concrete_HCS * 0.035 * photochemicalOF_reinforcement)*2 + (Volume_column_concrete_HCS * 0.97 * photochemicalOF_concrete_C30_37 + Volume_column_concrete_HCS * 0.03 * photochemicalOF_reinforcement)*4
photochemicalOF_score_support_steel_CLT = (Volume_beam_steel_CLT * photochemicalOF_structural_steel)*2 + (Volume_column_steel_CLT * photochemicalOF_structural_steel)*4
photochemicalOF_score_support_GLT_CLT = (Volume_beam_GLT_CLT * photochemicalOF_GLT)*2 + (Volume_column_GLT_CLT * photochemicalOF_GLT)*4
photochemicalOF_score_support_concrete_CLT = (Volume_beam_concrete_CLT * 0.965 * photochemicalOF_concrete_C30_37 + Volume_beam_concrete_CLT * 0.035 * photochemicalOF_reinforcement)*2 + (Volume_column_concrete_CLT * 0.97 * photochemicalOF_concrete_C30_37 + Volume_column_concrete_CLT * 0.03 * photochemicalOF_reinforcement)*4
photochemicalOF_score_support_steel_concrete = (Volume_beam_steel_concrete * photochemicalOF_structural_steel)*2 + (Volume_column_steel_concrete * photochemicalOF_structural_steel)*4
photochemicalOF_score_support_GLT_concrete = (Volume_beam_GLT_concrete * photochemicalOF_GLT)*2 + (Volume_column_GLT_concrete * photochemicalOF_GLT)*4
photochemicalOF_score_support_concrete_concrete = (Volume_beam_concrete_concrete * 0.965 * photochemicalOF_concrete_C30_37 + Volume_beam_concrete_concrete * 0.035 * photochemicalOF_reinforcement)*2 + (Volume_column_concrete_concrete * 0.97 * photochemicalOF_concrete_C30_37 + Volume_column_concrete_concrete * 0.03 * photochemicalOF_reinforcement)*4

ParticulatM_concrete_C50_60 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_ParticulatM, column_name_concrete_C50_60)
ParticulatM_concrete_C30_37 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_ParticulatM, column_name_concrete_C30_37)
ParticulatM_reinforcement =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_ParticulatM, column_name_reinforcement_steel)
ParticulatM_CLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_ParticulatM, column_name_CLT)
ParticulatM_GLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_ParticulatM, column_name_GLT)
ParticulatM_structural_steel =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_ParticulatM, column_name_structural_steel)

ParticulatM_score_HCS = ParticulatM_concrete_C50_60*Volume_concrete_floor_HCS + ParticulatM_concrete_C30_37*Volume_compression_layer_HCS + ParticulatM_reinforcement*Volume_steel_floor_HCS
ParticulatM_score_in_situ_concrete = ParticulatM_concrete_C30_37*Volume_in_situ_concrete + ParticulatM_reinforcement* Volume_reinforcement_in_situ_concrete
ParticulatM_score_CLT = ParticulatM_CLT*Volume_CLT
ParticulatM_score_support_steel_HCS = (Volume_beam_steel_HCS * ParticulatM_structural_steel)*2 + (Volume_column_steel_HCS * ParticulatM_structural_steel)*4
ParticulatM_score_support_GLT_HCS = (Volume_beam_GLT_HCS * ParticulatM_GLT)*2 + (Volume_column_GLT_HCS * ParticulatM_GLT)*4
ParticulatM_score_support_concrete_HCS = (Volume_beam_concrete_HCS * 0.965 * ParticulatM_concrete_C30_37 + Volume_beam_concrete_HCS * 0.035 * ParticulatM_reinforcement)*2 + (Volume_column_concrete_HCS * 0.97 * ParticulatM_concrete_C30_37 + Volume_column_concrete_HCS * 0.03 * ParticulatM_reinforcement)*4
ParticulatM_score_support_steel_CLT = (Volume_beam_steel_CLT * ParticulatM_structural_steel)*2 + (Volume_column_steel_CLT * ParticulatM_structural_steel)*4
ParticulatM_score_support_GLT_CLT = (Volume_beam_GLT_CLT * ParticulatM_GLT)*2 + (Volume_column_GLT_CLT * ParticulatM_GLT)*4
ParticulatM_score_support_concrete_CLT = (Volume_beam_concrete_CLT * 0.965 * ParticulatM_concrete_C30_37 + Volume_beam_concrete_CLT * 0.035 * ParticulatM_reinforcement)*2 + (Volume_column_concrete_CLT * 0.97 * ParticulatM_concrete_C30_37 + Volume_column_concrete_CLT * 0.03 * ParticulatM_reinforcement)*4
ParticulatM_score_support_steel_concrete = (Volume_beam_steel_concrete * ParticulatM_structural_steel)*2 + (Volume_column_steel_concrete * ParticulatM_structural_steel)*4
ParticulatM_score_support_GLT_concrete = (Volume_beam_GLT_concrete * ParticulatM_GLT)*2 + (Volume_column_GLT_concrete * ParticulatM_GLT)*4
ParticulatM_score_support_concrete_concrete = (Volume_beam_concrete_concrete * 0.965 * ParticulatM_concrete_C30_37 + Volume_beam_concrete_concrete * 0.035 * ParticulatM_reinforcement)*2 + (Volume_column_concrete_concrete * 0.97 * ParticulatM_concrete_C30_37 + Volume_column_concrete_concrete * 0.03 * ParticulatM_reinforcement)*4

HumanTNC_concrete_C50_60 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_HumanTNC, column_name_concrete_C50_60)
HumanTNC_concrete_C30_37 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_HumanTNC, column_name_concrete_C30_37)
HumanTNC_reinforcement =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_HumanTNC, column_name_reinforcement_steel)
HumanTNC_CLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_HumanTNC, column_name_CLT)
HumanTNC_GLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_HumanTNC, column_name_GLT)
HumanTNC_structural_steel =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_HumanTNC, column_name_structural_steel)

HumanTNC_score_HCS = HumanTNC_concrete_C50_60*Volume_concrete_floor_HCS + HumanTNC_concrete_C30_37*Volume_compression_layer_HCS + HumanTNC_reinforcement*Volume_steel_floor_HCS
HumanTNC_score_in_situ_concrete = HumanTNC_concrete_C30_37*Volume_in_situ_concrete + HumanTNC_reinforcement* Volume_reinforcement_in_situ_concrete
HumanTNC_score_CLT = HumanTNC_CLT*Volume_CLT
HumanTNC_score_support_steel_HCS = (Volume_beam_steel_HCS * HumanTNC_structural_steel)*2 + (Volume_column_steel_HCS * HumanTNC_structural_steel)*4
HumanTNC_score_support_GLT_HCS = (Volume_beam_GLT_HCS * HumanTNC_GLT)*2 + (Volume_column_GLT_HCS * HumanTNC_GLT)*4
HumanTNC_score_support_concrete_HCS = (Volume_beam_concrete_HCS * 0.965 * HumanTNC_concrete_C30_37 + Volume_beam_concrete_HCS * 0.035 * HumanTNC_reinforcement)*2 + (Volume_column_concrete_HCS * 0.97 * HumanTNC_concrete_C30_37 + Volume_column_concrete_HCS * 0.03 * HumanTNC_reinforcement)*4
HumanTNC_score_support_steel_CLT = (Volume_beam_steel_CLT * HumanTNC_structural_steel)*2 + (Volume_column_steel_CLT * HumanTNC_structural_steel)*4
HumanTNC_score_support_GLT_CLT = (Volume_beam_GLT_CLT * HumanTNC_GLT)*2 + (Volume_column_GLT_CLT * HumanTNC_GLT)*4
HumanTNC_score_support_concrete_CLT = (Volume_beam_concrete_CLT * 0.965 * HumanTNC_concrete_C30_37 + Volume_beam_concrete_CLT * 0.035 * HumanTNC_reinforcement)*2 + (Volume_column_concrete_CLT * 0.97 * HumanTNC_concrete_C30_37 + Volume_column_concrete_CLT * 0.03 * HumanTNC_reinforcement)*4
HumanTNC_score_support_steel_concrete = (Volume_beam_steel_concrete * HumanTNC_structural_steel)*2 + (Volume_column_steel_concrete * HumanTNC_structural_steel)*4
HumanTNC_score_support_GLT_concrete = (Volume_beam_GLT_concrete * HumanTNC_GLT)*2 + (Volume_column_GLT_concrete * HumanTNC_GLT)*4
HumanTNC_score_support_concrete_concrete = (Volume_beam_concrete_concrete * 0.965 * HumanTNC_concrete_C30_37 + Volume_beam_concrete_concrete * 0.035 * HumanTNC_reinforcement)*2 + (Volume_column_concrete_concrete * 0.97 * HumanTNC_concrete_C30_37 + Volume_column_concrete_concrete * 0.03 * HumanTNC_reinforcement)*4

HumanTC_concrete_C50_60 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_HumanTC, column_name_concrete_C50_60)
HumanTC_concrete_C30_37 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_HumanTC, column_name_concrete_C30_37)
HumanTC_reinforcement =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_HumanTC, column_name_reinforcement_steel)
HumanTC_CLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_HumanTC, column_name_CLT)
HumanTC_GLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_HumanTC, column_name_GLT)
HumanTC_structural_steel =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_HumanTC, column_name_structural_steel)

HumanTC_score_HCS = HumanTC_concrete_C50_60*Volume_concrete_floor_HCS + HumanTC_concrete_C30_37*Volume_compression_layer_HCS + HumanTC_reinforcement*Volume_steel_floor_HCS
HumanTC_score_in_situ_concrete = HumanTC_concrete_C30_37*Volume_in_situ_concrete + HumanTC_reinforcement* Volume_reinforcement_in_situ_concrete
HumanTC_score_CLT = HumanTC_CLT*Volume_CLT
HumanTC_score_support_steel_HCS = (Volume_beam_steel_HCS * HumanTC_structural_steel)*2 + (Volume_column_steel_HCS * HumanTC_structural_steel)*4
HumanTC_score_support_GLT_HCS = (Volume_beam_GLT_HCS * HumanTC_GLT)*2 + (Volume_column_GLT_HCS * HumanTC_GLT)*4
HumanTC_score_support_concrete_HCS = (Volume_beam_concrete_HCS * 0.965 * HumanTC_concrete_C30_37 + Volume_beam_concrete_HCS * 0.035 * HumanTC_reinforcement)*2 + (Volume_column_concrete_HCS * 0.97 * HumanTC_concrete_C30_37 + Volume_column_concrete_HCS * 0.03 * HumanTC_reinforcement)*4
HumanTC_score_support_steel_CLT = (Volume_beam_steel_CLT * HumanTC_structural_steel)*2 + (Volume_column_steel_CLT * HumanTC_structural_steel)*4
HumanTC_score_support_GLT_CLT = (Volume_beam_GLT_CLT * HumanTC_GLT)*2 + (Volume_column_GLT_CLT * HumanTC_GLT)*4
HumanTC_score_support_concrete_CLT = (Volume_beam_concrete_CLT * 0.965 * HumanTC_concrete_C30_37 + Volume_beam_concrete_CLT * 0.035 * HumanTC_reinforcement)*2 + (Volume_column_concrete_CLT * 0.97 * HumanTC_concrete_C30_37 + Volume_column_concrete_CLT * 0.03 * HumanTC_reinforcement)*4
HumanTC_score_support_steel_concrete = (Volume_beam_steel_concrete * HumanTC_structural_steel)*2 + (Volume_column_steel_concrete * HumanTC_structural_steel)*4
HumanTC_score_support_GLT_concrete = (Volume_beam_GLT_concrete * HumanTC_GLT)*2 + (Volume_column_GLT_concrete * HumanTC_GLT)*4
HumanTC_score_support_concrete_concrete = (Volume_beam_concrete_concrete * 0.965 * HumanTC_concrete_C30_37 + Volume_beam_concrete_concrete * 0.035 * HumanTC_reinforcement)*2 + (Volume_column_concrete_concrete * 0.97 * HumanTC_concrete_C30_37 + Volume_column_concrete_concrete * 0.03 * HumanTC_reinforcement)*4

Acidification_concrete_C50_60 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_Acidification, column_name_concrete_C50_60)
Acidification_concrete_C30_37 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_Acidification, column_name_concrete_C30_37)
Acidification_reinforcement =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_Acidification, column_name_reinforcement_steel)
Acidification_CLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_Acidification, column_name_CLT)
Acidification_GLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_Acidification, column_name_GLT)
Acidification_structural_steel =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_Acidification, column_name_structural_steel)

Acidification_score_HCS = Acidification_concrete_C50_60*Volume_concrete_floor_HCS + Acidification_concrete_C30_37*Volume_compression_layer_HCS + Acidification_reinforcement*Volume_steel_floor_HCS
Acidification_score_in_situ_concrete = Acidification_concrete_C30_37*Volume_in_situ_concrete + Acidification_reinforcement* Volume_reinforcement_in_situ_concrete
Acidification_score_CLT = Acidification_CLT*Volume_CLT
Acidification_score_support_steel_HCS = (Volume_beam_steel_HCS * Acidification_structural_steel)*2 + (Volume_column_steel_HCS * Acidification_structural_steel)*4
Acidification_score_support_GLT_HCS = (Volume_beam_GLT_HCS * Acidification_GLT)*2 + (Volume_column_GLT_HCS * Acidification_GLT)*4
Acidification_score_support_concrete_HCS = (Volume_beam_concrete_HCS * 0.965 * Acidification_concrete_C30_37 + Volume_beam_concrete_HCS * 0.035 * Acidification_reinforcement)*2 + (Volume_column_concrete_HCS * 0.97 * Acidification_concrete_C30_37 + Volume_column_concrete_HCS * 0.03 * Acidification_reinforcement)*4
Acidification_score_support_steel_CLT = (Volume_beam_steel_CLT * Acidification_structural_steel)*2 + (Volume_column_steel_CLT * Acidification_structural_steel)*4
Acidification_score_support_GLT_CLT = (Volume_beam_GLT_CLT * Acidification_GLT)*2 + (Volume_column_GLT_CLT * Acidification_GLT)*4
Acidification_score_support_concrete_CLT = (Volume_beam_concrete_CLT * 0.965 * Acidification_concrete_C30_37 + Volume_beam_concrete_CLT * 0.035 * Acidification_reinforcement)*2 + (Volume_column_concrete_CLT * 0.97 * Acidification_concrete_C30_37 + Volume_column_concrete_CLT * 0.03 * Acidification_reinforcement)*4
Acidification_score_support_steel_concrete = (Volume_beam_steel_concrete * Acidification_structural_steel)*2 + (Volume_column_steel_concrete * Acidification_structural_steel)*4
Acidification_score_support_GLT_concrete = (Volume_beam_GLT_concrete * Acidification_GLT)*2 + (Volume_column_GLT_concrete * Acidification_GLT)*4
Acidification_score_support_concrete_concrete = (Volume_beam_concrete_concrete * 0.965 * Acidification_concrete_C30_37 + Volume_beam_concrete_concrete * 0.035 * Acidification_reinforcement)*2 + (Volume_column_concrete_concrete * 0.97 * Acidification_concrete_C30_37 + Volume_column_concrete_concrete * 0.03 * Acidification_reinforcement)*4

EutrophicationF_concrete_C50_60 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_EutrophicationF, column_name_concrete_C50_60)
EutrophicationF_concrete_C30_37 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_EutrophicationF, column_name_concrete_C30_37)
EutrophicationF_reinforcement =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_EutrophicationF, column_name_reinforcement_steel)
EutrophicationF_CLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_EutrophicationF, column_name_CLT)
EutrophicationF_GLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_EutrophicationF, column_name_GLT)
EutrophicationF_structural_steel =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_EutrophicationF, column_name_structural_steel)

EutrophicationF_score_HCS = EutrophicationF_concrete_C50_60*Volume_concrete_floor_HCS + EutrophicationF_concrete_C30_37*Volume_compression_layer_HCS + EutrophicationF_reinforcement*Volume_steel_floor_HCS
EutrophicationF_score_in_situ_concrete = EutrophicationF_concrete_C30_37*Volume_in_situ_concrete + EutrophicationF_reinforcement* Volume_reinforcement_in_situ_concrete
EutrophicationF_score_CLT = EutrophicationF_CLT*Volume_CLT
EutrophicationF_score_support_steel_HCS = (Volume_beam_steel_HCS * EutrophicationF_structural_steel)*2 + (Volume_column_steel_HCS * EutrophicationF_structural_steel)*4
EutrophicationF_score_support_GLT_HCS = (Volume_beam_GLT_HCS * EutrophicationF_GLT)*2 + (Volume_column_GLT_HCS * EutrophicationF_GLT)*4
EutrophicationF_score_support_concrete_HCS = (Volume_beam_concrete_HCS * 0.965 * EutrophicationF_concrete_C30_37 + Volume_beam_concrete_HCS * 0.035 * EutrophicationF_reinforcement)*2 + (Volume_column_concrete_HCS * 0.97 * EutrophicationF_concrete_C30_37 + Volume_column_concrete_HCS * 0.03 * EutrophicationF_reinforcement)*4
EutrophicationF_score_support_steel_CLT = (Volume_beam_steel_CLT * EutrophicationF_structural_steel)*2 + (Volume_column_steel_CLT * EutrophicationF_structural_steel)*4
EutrophicationF_score_support_GLT_CLT = (Volume_beam_GLT_CLT * EutrophicationF_GLT)*2 + (Volume_column_GLT_CLT * EutrophicationF_GLT)*4
EutrophicationF_score_support_concrete_CLT = (Volume_beam_concrete_CLT * 0.965 * EutrophicationF_concrete_C30_37 + Volume_beam_concrete_CLT * 0.035 * EutrophicationF_reinforcement)*2 + (Volume_column_concrete_CLT * 0.97 * EutrophicationF_concrete_C30_37 + Volume_column_concrete_CLT * 0.03 * EutrophicationF_reinforcement)*4
EutrophicationF_score_support_steel_concrete = (Volume_beam_steel_concrete * EutrophicationF_structural_steel)*2 + (Volume_column_steel_concrete * EutrophicationF_structural_steel)*4
EutrophicationF_score_support_GLT_concrete = (Volume_beam_GLT_concrete * EutrophicationF_GLT)*2 + (Volume_column_GLT_concrete * EutrophicationF_GLT)*4
EutrophicationF_score_support_concrete_concrete = (Volume_beam_concrete_concrete * 0.965 * EutrophicationF_concrete_C30_37 + Volume_beam_concrete_concrete * 0.035 * EutrophicationF_reinforcement)*2 + (Volume_column_concrete_concrete * 0.97 * EutrophicationF_concrete_C30_37 + Volume_column_concrete_concrete * 0.03 * EutrophicationF_reinforcement)*4

EutrophicationM_concrete_C50_60 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_EutrophicationM, column_name_concrete_C50_60)
EutrophicationM_concrete_C30_37 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_EutrophicationM, column_name_concrete_C30_37)
EutrophicationM_reinforcement =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_EutrophicationM, column_name_reinforcement_steel)
EutrophicationM_CLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_EutrophicationM, column_name_CLT)
EutrophicationM_GLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_EutrophicationM, column_name_GLT)
EutrophicationM_structural_steel =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_EutrophicationM, column_name_structural_steel)

EutrophicationM_score_HCS = EutrophicationM_concrete_C50_60*Volume_concrete_floor_HCS + EutrophicationM_concrete_C30_37*Volume_compression_layer_HCS + EutrophicationM_reinforcement*Volume_steel_floor_HCS
EutrophicationM_score_in_situ_concrete = EutrophicationM_concrete_C30_37*Volume_in_situ_concrete + EutrophicationM_reinforcement* Volume_reinforcement_in_situ_concrete
EutrophicationM_score_CLT = EutrophicationM_CLT*Volume_CLT
EutrophicationM_score_support_steel_HCS = (Volume_beam_steel_HCS * EutrophicationM_structural_steel)*2 + (Volume_column_steel_HCS * EutrophicationM_structural_steel)*4
EutrophicationM_score_support_GLT_HCS = (Volume_beam_GLT_HCS * EutrophicationM_GLT)*2 + (Volume_column_GLT_HCS * EutrophicationM_GLT)*4
EutrophicationM_score_support_concrete_HCS = (Volume_beam_concrete_HCS * 0.965 * EutrophicationM_concrete_C30_37 + Volume_beam_concrete_HCS * 0.035 * EutrophicationM_reinforcement)*2 + (Volume_column_concrete_HCS * 0.97 * EutrophicationM_concrete_C30_37 + Volume_column_concrete_HCS * 0.03 * EutrophicationM_reinforcement)*4
EutrophicationM_score_support_steel_CLT = (Volume_beam_steel_CLT * EutrophicationM_structural_steel)*2 + (Volume_column_steel_CLT * EutrophicationM_structural_steel)*4
EutrophicationM_score_support_GLT_CLT = (Volume_beam_GLT_CLT * EutrophicationM_GLT)*2 + (Volume_column_GLT_CLT * EutrophicationM_GLT)*4
EutrophicationM_score_support_concrete_CLT = (Volume_beam_concrete_CLT * 0.965 * EutrophicationM_concrete_C30_37 + Volume_beam_concrete_CLT * 0.035 * EutrophicationM_reinforcement)*2 + (Volume_column_concrete_CLT * 0.97 * EutrophicationM_concrete_C30_37 + Volume_column_concrete_CLT * 0.03 * EutrophicationM_reinforcement)*4
EutrophicationM_score_support_steel_concrete = (Volume_beam_steel_concrete * EutrophicationM_structural_steel)*2 + (Volume_column_steel_concrete * EutrophicationM_structural_steel)*4
EutrophicationM_score_support_GLT_concrete = (Volume_beam_GLT_concrete * EutrophicationM_GLT)*2 + (Volume_column_GLT_concrete * EutrophicationM_GLT)*4
EutrophicationM_score_support_concrete_concrete = (Volume_beam_concrete_concrete * 0.965 * EutrophicationM_concrete_C30_37 + Volume_beam_concrete_concrete * 0.035 * EutrophicationM_reinforcement)*2 + (Volume_column_concrete_concrete * 0.97 * EutrophicationM_concrete_C30_37 + Volume_column_concrete_concrete * 0.03 * EutrophicationM_reinforcement)*4

EutrophicationT_concrete_C50_60 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_EutrophicationT, column_name_concrete_C50_60)
EutrophicationT_concrete_C30_37 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_EutrophicationT, column_name_concrete_C30_37)
EutrophicationT_reinforcement =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_EutrophicationT, column_name_reinforcement_steel)
EutrophicationT_CLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_EutrophicationT, column_name_CLT)
EutrophicationT_GLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_EutrophicationT, column_name_GLT)
EutrophicationT_structural_steel =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_EutrophicationT, column_name_structural_steel)

EutrophicationT_score_HCS = EutrophicationT_concrete_C50_60*Volume_concrete_floor_HCS + EutrophicationT_concrete_C30_37*Volume_compression_layer_HCS + EutrophicationT_reinforcement*Volume_steel_floor_HCS
EutrophicationT_score_in_situ_concrete = EutrophicationT_concrete_C30_37*Volume_in_situ_concrete + EutrophicationT_reinforcement* Volume_reinforcement_in_situ_concrete
EutrophicationT_score_CLT = EutrophicationT_CLT*Volume_CLT
EutrophicationT_score_support_steel_HCS = (Volume_beam_steel_HCS * EutrophicationT_structural_steel)*2 + (Volume_column_steel_HCS * EutrophicationT_structural_steel)*4
EutrophicationT_score_support_GLT_HCS = (Volume_beam_GLT_HCS * EutrophicationT_GLT)*2 + (Volume_column_GLT_HCS * EutrophicationT_GLT)*4
EutrophicationT_score_support_concrete_HCS = (Volume_beam_concrete_HCS * 0.965 * EutrophicationT_concrete_C30_37 + Volume_beam_concrete_HCS * 0.035 * EutrophicationT_reinforcement)*2 + (Volume_column_concrete_HCS * 0.97 * EutrophicationT_concrete_C30_37 + Volume_column_concrete_HCS * 0.03 * EutrophicationT_reinforcement)*4
EutrophicationT_score_support_steel_CLT = (Volume_beam_steel_CLT * EutrophicationT_structural_steel)*2 + (Volume_column_steel_CLT * EutrophicationT_structural_steel)*4
EutrophicationT_score_support_GLT_CLT = (Volume_beam_GLT_CLT * EutrophicationT_GLT)*2 + (Volume_column_GLT_CLT * EutrophicationT_GLT)*4
EutrophicationT_score_support_concrete_CLT = (Volume_beam_concrete_CLT * 0.965 * EutrophicationT_concrete_C30_37 + Volume_beam_concrete_CLT * 0.035 * EutrophicationT_reinforcement)*2 + (Volume_column_concrete_CLT * 0.97 * EutrophicationT_concrete_C30_37 + Volume_column_concrete_CLT * 0.03 * EutrophicationT_reinforcement)*4
EutrophicationT_score_support_steel_concrete = (Volume_beam_steel_concrete * EutrophicationT_structural_steel)*2 + (Volume_column_steel_concrete * EutrophicationT_structural_steel)*4
EutrophicationT_score_support_GLT_concrete = (Volume_beam_GLT_concrete * EutrophicationT_GLT)*2 + (Volume_column_GLT_concrete * EutrophicationT_GLT)*4
EutrophicationT_score_support_concrete_concrete = (Volume_beam_concrete_concrete * 0.965 * EutrophicationT_concrete_C30_37 + Volume_beam_concrete_concrete * 0.035 * EutrophicationT_reinforcement)*2 + (Volume_column_concrete_concrete * 0.97 * EutrophicationT_concrete_C30_37 + Volume_column_concrete_concrete * 0.03 * EutrophicationT_reinforcement)*4

EcotoxicityF_concrete_C50_60 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_EcotoxicityF, column_name_concrete_C50_60)
EcotoxicityF_concrete_C30_37 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_EcotoxicityF, column_name_concrete_C30_37)
EcotoxicityF_reinforcement =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_EcotoxicityF, column_name_reinforcement_steel)
EcotoxicityF_CLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_EcotoxicityF, column_name_CLT)
EcotoxicityF_GLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_EcotoxicityF, column_name_GLT)
EcotoxicityF_structural_steel =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_EcotoxicityF, column_name_structural_steel)

EcotoxicityF_score_HCS = EcotoxicityF_concrete_C50_60*Volume_concrete_floor_HCS + EcotoxicityF_concrete_C30_37*Volume_compression_layer_HCS + EcotoxicityF_reinforcement*Volume_steel_floor_HCS
EcotoxicityF_score_in_situ_concrete = EcotoxicityF_concrete_C30_37*Volume_in_situ_concrete + EcotoxicityF_reinforcement* Volume_reinforcement_in_situ_concrete
EcotoxicityF_score_CLT = EcotoxicityF_CLT*Volume_CLT
EcotoxicityF_score_support_steel_HCS = (Volume_beam_steel_HCS * EcotoxicityF_structural_steel)*2 + (Volume_column_steel_HCS * EcotoxicityF_structural_steel)*4
EcotoxicityF_score_support_GLT_HCS = (Volume_beam_GLT_HCS * EcotoxicityF_GLT)*2 + (Volume_column_GLT_HCS * EcotoxicityF_GLT)*4
EcotoxicityF_score_support_concrete_HCS = (Volume_beam_concrete_HCS * 0.965 * EcotoxicityF_concrete_C30_37 + Volume_beam_concrete_HCS * 0.035 * EcotoxicityF_reinforcement)*2 + (Volume_column_concrete_HCS * 0.97 * EcotoxicityF_concrete_C30_37 + Volume_column_concrete_HCS * 0.03 * EcotoxicityF_reinforcement)*4
EcotoxicityF_score_support_steel_CLT = (Volume_beam_steel_CLT * EcotoxicityF_structural_steel)*2 + (Volume_column_steel_CLT * EcotoxicityF_structural_steel)*4
EcotoxicityF_score_support_GLT_CLT = (Volume_beam_GLT_CLT * EcotoxicityF_GLT)*2 + (Volume_column_GLT_CLT * EcotoxicityF_GLT)*4
EcotoxicityF_score_support_concrete_CLT = (Volume_beam_concrete_CLT * 0.965 * EcotoxicityF_concrete_C30_37 + Volume_beam_concrete_CLT * 0.035 * EcotoxicityF_reinforcement)*2 + (Volume_column_concrete_CLT * 0.97 * EcotoxicityF_concrete_C30_37 + Volume_column_concrete_CLT * 0.03 * EcotoxicityF_reinforcement)*4
EcotoxicityF_score_support_steel_concrete = (Volume_beam_steel_concrete * EcotoxicityF_structural_steel)*2 + (Volume_column_steel_concrete * EcotoxicityF_structural_steel)*4
EcotoxicityF_score_support_GLT_concrete = (Volume_beam_GLT_concrete * EcotoxicityF_GLT)*2 + (Volume_column_GLT_concrete * EcotoxicityF_GLT)*4
EcotoxicityF_score_support_concrete_concrete = (Volume_beam_concrete_concrete * 0.965 * EcotoxicityF_concrete_C30_37 + Volume_beam_concrete_concrete * 0.035 * EcotoxicityF_reinforcement)*2 + (Volume_column_concrete_concrete * 0.97 * EcotoxicityF_concrete_C30_37 + Volume_column_concrete_concrete * 0.03 * EcotoxicityF_reinforcement)*4

LandUse_concrete_C50_60 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_LandUse, column_name_concrete_C50_60)
LandUse_concrete_C30_37 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_LandUse, column_name_concrete_C30_37)
LandUse_reinforcement =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_LandUse, column_name_reinforcement_steel)
LandUse_CLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_LandUse, column_name_CLT)
LandUse_GLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_LandUse, column_name_GLT)
LandUse_structural_steel =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_LandUse, column_name_structural_steel)

LandUse_score_HCS = LandUse_concrete_C50_60*Volume_concrete_floor_HCS + LandUse_concrete_C30_37*Volume_compression_layer_HCS + LandUse_reinforcement*Volume_steel_floor_HCS
LandUse_score_in_situ_concrete = LandUse_concrete_C30_37*Volume_in_situ_concrete + LandUse_reinforcement* Volume_reinforcement_in_situ_concrete
LandUse_score_CLT = LandUse_CLT*Volume_CLT
LandUse_score_support_steel_HCS = (Volume_beam_steel_HCS * LandUse_structural_steel)*2 + (Volume_column_steel_HCS * LandUse_structural_steel)*4
LandUse_score_support_GLT_HCS = (Volume_beam_GLT_HCS * LandUse_GLT)*2 + (Volume_column_GLT_HCS * LandUse_GLT)*4
LandUse_score_support_concrete_HCS = (Volume_beam_concrete_HCS * 0.965 * LandUse_concrete_C30_37 + Volume_beam_concrete_HCS * 0.035 * LandUse_reinforcement)*2 + (Volume_column_concrete_HCS * 0.97 * LandUse_concrete_C30_37 + Volume_column_concrete_HCS * 0.03 * LandUse_reinforcement)*4
LandUse_score_support_steel_CLT = (Volume_beam_steel_CLT * LandUse_structural_steel)*2 + (Volume_column_steel_CLT * LandUse_structural_steel)*4
LandUse_score_support_GLT_CLT = (Volume_beam_GLT_CLT * LandUse_GLT)*2 + (Volume_column_GLT_CLT * LandUse_GLT)*4
LandUse_score_support_concrete_CLT = (Volume_beam_concrete_CLT * 0.965 * LandUse_concrete_C30_37 + Volume_beam_concrete_CLT * 0.035 * LandUse_reinforcement)*2 + (Volume_column_concrete_CLT * 0.97 * LandUse_concrete_C30_37 + Volume_column_concrete_CLT * 0.03 * LandUse_reinforcement)*4
LandUse_score_support_steel_concrete = (Volume_beam_steel_concrete * LandUse_structural_steel)*2 + (Volume_column_steel_concrete * LandUse_structural_steel)*4
LandUse_score_support_GLT_concrete = (Volume_beam_GLT_concrete * LandUse_GLT)*2 + (Volume_column_GLT_concrete * LandUse_GLT)*4
LandUse_score_support_concrete_concrete = (Volume_beam_concrete_concrete * 0.965 * LandUse_concrete_C30_37 + Volume_beam_concrete_concrete * 0.035 * LandUse_reinforcement)*2 + (Volume_column_concrete_concrete * 0.97 * LandUse_concrete_C30_37 + Volume_column_concrete_concrete * 0.03 * LandUse_reinforcement)*4

WaterUse_concrete_C50_60 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_WaterUse, column_name_concrete_C50_60)
WaterUse_concrete_C30_37 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_WaterUse, column_name_concrete_C30_37)
WaterUse_reinforcement =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_WaterUse, column_name_reinforcement_steel)
WaterUse_CLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_WaterUse, column_name_CLT)
WaterUse_GLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_WaterUse, column_name_GLT)
WaterUse_structural_steel =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_WaterUse, column_name_structural_steel)

WaterUse_score_HCS = WaterUse_concrete_C50_60*Volume_concrete_floor_HCS + WaterUse_concrete_C30_37*Volume_compression_layer_HCS + WaterUse_reinforcement*Volume_steel_floor_HCS
WaterUse_score_in_situ_concrete = WaterUse_concrete_C30_37*Volume_in_situ_concrete + WaterUse_reinforcement* Volume_reinforcement_in_situ_concrete
WaterUse_score_CLT = WaterUse_CLT*Volume_CLT
WaterUse_score_support_steel_HCS = (Volume_beam_steel_HCS * WaterUse_structural_steel)*2 + (Volume_column_steel_HCS * WaterUse_structural_steel)*4
WaterUse_score_support_GLT_HCS = (Volume_beam_GLT_HCS * WaterUse_GLT)*2 + (Volume_column_GLT_HCS * WaterUse_GLT)*4
WaterUse_score_support_concrete_HCS = (Volume_beam_concrete_HCS * 0.965 * WaterUse_concrete_C30_37 + Volume_beam_concrete_HCS * 0.035 * WaterUse_reinforcement)*2 + (Volume_column_concrete_HCS * 0.97 * WaterUse_concrete_C30_37 + Volume_column_concrete_HCS * 0.03 * WaterUse_reinforcement)*4
WaterUse_score_support_steel_CLT = (Volume_beam_steel_CLT * WaterUse_structural_steel)*2 + (Volume_column_steel_CLT * WaterUse_structural_steel)*4
WaterUse_score_support_GLT_CLT = (Volume_beam_GLT_CLT * WaterUse_GLT)*2 + (Volume_column_GLT_CLT * WaterUse_GLT)*4
WaterUse_score_support_concrete_CLT = (Volume_beam_concrete_CLT * 0.965 * WaterUse_concrete_C30_37 + Volume_beam_concrete_CLT * 0.035 * WaterUse_reinforcement)*2 + (Volume_column_concrete_CLT * 0.97 * WaterUse_concrete_C30_37 + Volume_column_concrete_CLT * 0.03 * WaterUse_reinforcement)*4
WaterUse_score_support_steel_concrete = (Volume_beam_steel_concrete * WaterUse_structural_steel)*2 + (Volume_column_steel_concrete * WaterUse_structural_steel)*4
WaterUse_score_support_GLT_concrete = (Volume_beam_GLT_concrete * WaterUse_GLT)*2 + (Volume_column_GLT_concrete * WaterUse_GLT)*4
WaterUse_score_support_concrete_concrete = (Volume_beam_concrete_concrete * 0.965 * WaterUse_concrete_C30_37 + Volume_beam_concrete_concrete * 0.035 * WaterUse_reinforcement)*2 + (Volume_column_concrete_concrete * 0.97 * WaterUse_concrete_C30_37 + Volume_column_concrete_concrete * 0.03 * WaterUse_reinforcement)*4

ResourceUseF_concrete_C50_60 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_ResourceUseF, column_name_concrete_C50_60)
ResourceUseF_concrete_C30_37 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_ResourceUseF, column_name_concrete_C30_37)
ResourceUseF_reinforcement =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_ResourceUseF, column_name_reinforcement_steel)
ResourceUseF_CLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_ResourceUseF, column_name_CLT)
ResourceUseF_GLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_ResourceUseF, column_name_GLT)
ResourceUseF_structural_steel =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_ResourceUseF, column_name_structural_steel)

ResourceUseF_score_HCS = ResourceUseF_concrete_C50_60*Volume_concrete_floor_HCS + ResourceUseF_concrete_C30_37*Volume_compression_layer_HCS + ResourceUseF_reinforcement*Volume_steel_floor_HCS
ResourceUseF_score_in_situ_concrete = ResourceUseF_concrete_C30_37*Volume_in_situ_concrete + ResourceUseF_reinforcement* Volume_reinforcement_in_situ_concrete
ResourceUseF_score_CLT = ResourceUseF_CLT*Volume_CLT
ResourceUseF_score_support_steel_HCS = (Volume_beam_steel_HCS * ResourceUseF_structural_steel)*2 + (Volume_column_steel_HCS * ResourceUseF_structural_steel)*4
ResourceUseF_score_support_GLT_HCS = (Volume_beam_GLT_HCS * ResourceUseF_GLT)*2 + (Volume_column_GLT_HCS * ResourceUseF_GLT)*4
ResourceUseF_score_support_concrete_HCS = (Volume_beam_concrete_HCS * 0.965 * ResourceUseF_concrete_C30_37 + Volume_beam_concrete_HCS * 0.035 * ResourceUseF_reinforcement)*2 + (Volume_column_concrete_HCS * 0.97 * ResourceUseF_concrete_C30_37 + Volume_column_concrete_HCS * 0.03 * ResourceUseF_reinforcement)*4
ResourceUseF_score_support_steel_CLT = (Volume_beam_steel_CLT * ResourceUseF_structural_steel)*2 + (Volume_column_steel_CLT * ResourceUseF_structural_steel)*4
ResourceUseF_score_support_GLT_CLT = (Volume_beam_GLT_CLT * ResourceUseF_GLT)*2 + (Volume_column_GLT_CLT * ResourceUseF_GLT)*4
ResourceUseF_score_support_concrete_CLT = (Volume_beam_concrete_CLT * 0.965 * ResourceUseF_concrete_C30_37 + Volume_beam_concrete_CLT * 0.035 * ResourceUseF_reinforcement)*2 + (Volume_column_concrete_CLT * 0.97 * ResourceUseF_concrete_C30_37 + Volume_column_concrete_CLT * 0.03 * ResourceUseF_reinforcement)*4
ResourceUseF_score_support_steel_concrete = (Volume_beam_steel_concrete * ResourceUseF_structural_steel)*2 + (Volume_column_steel_concrete * ResourceUseF_structural_steel)*4
ResourceUseF_score_support_GLT_concrete = (Volume_beam_GLT_concrete * ResourceUseF_GLT)*2 + (Volume_column_GLT_concrete * ResourceUseF_GLT)*4
ResourceUseF_score_support_concrete_concrete = (Volume_beam_concrete_concrete * 0.965 * ResourceUseF_concrete_C30_37 + Volume_beam_concrete_concrete * 0.035 * ResourceUseF_reinforcement)*2 + (Volume_column_concrete_concrete * 0.97 * ResourceUseF_concrete_C30_37 + Volume_column_concrete_concrete * 0.03 * ResourceUseF_reinforcement)*4

ResourceUseMM_concrete_C50_60 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_ResourceUseMM, column_name_concrete_C50_60)
ResourceUseMM_concrete_C30_37 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_ResourceUseMM, column_name_concrete_C30_37)
ResourceUseMM_reinforcement =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_ResourceUseMM, column_name_reinforcement_steel)
ResourceUseMM_CLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_ResourceUseMM, column_name_CLT)
ResourceUseMM_GLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_ResourceUseMM, column_name_GLT)
ResourceUseMM_structural_steel =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_ResourceUseMM, column_name_structural_steel)

ResourceUseMM_score_HCS = ResourceUseMM_concrete_C50_60*Volume_concrete_floor_HCS + ResourceUseMM_concrete_C30_37*Volume_compression_layer_HCS + ResourceUseMM_reinforcement*Volume_steel_floor_HCS
ResourceUseMM_score_in_situ_concrete = ResourceUseMM_concrete_C30_37*Volume_in_situ_concrete + ResourceUseMM_reinforcement* Volume_reinforcement_in_situ_concrete
ResourceUseMM_score_CLT = ResourceUseMM_CLT*Volume_CLT
ResourceUseMM_score_support_steel_HCS = (Volume_beam_steel_HCS * ResourceUseMM_structural_steel)*2 + (Volume_column_steel_HCS * ResourceUseMM_structural_steel)*4
ResourceUseMM_score_support_GLT_HCS = (Volume_beam_GLT_HCS * ResourceUseMM_GLT)*2 + (Volume_column_GLT_HCS * ResourceUseMM_GLT)*4
ResourceUseMM_score_support_concrete_HCS = (Volume_beam_concrete_HCS * 0.965 * ResourceUseMM_concrete_C30_37 + Volume_beam_concrete_HCS * 0.035 * ResourceUseMM_reinforcement)*2 + (Volume_column_concrete_HCS * 0.97 * ResourceUseMM_concrete_C30_37 + Volume_column_concrete_HCS * 0.03 * ResourceUseMM_reinforcement)*4
ResourceUseMM_score_support_steel_CLT = (Volume_beam_steel_CLT * ResourceUseMM_structural_steel)*2 + (Volume_column_steel_CLT * ResourceUseMM_structural_steel)*4
ResourceUseMM_score_support_GLT_CLT = (Volume_beam_GLT_CLT * ResourceUseMM_GLT)*2 + (Volume_column_GLT_CLT * ResourceUseMM_GLT)*4
ResourceUseMM_score_support_concrete_CLT = (Volume_beam_concrete_CLT * 0.965 * ResourceUseMM_concrete_C30_37 + Volume_beam_concrete_CLT * 0.035 * ResourceUseMM_reinforcement)*2 + (Volume_column_concrete_CLT * 0.97 * ResourceUseMM_concrete_C30_37 + Volume_column_concrete_CLT * 0.03 * ResourceUseMM_reinforcement)*4
ResourceUseMM_score_support_steel_concrete = (Volume_beam_steel_concrete * ResourceUseMM_structural_steel)*2 + (Volume_column_steel_concrete * ResourceUseMM_structural_steel)*4
ResourceUseMM_score_support_GLT_concrete = (Volume_beam_GLT_concrete * ResourceUseMM_GLT)*2 + (Volume_column_GLT_concrete * ResourceUseMM_GLT)*4
ResourceUseMM_score_support_concrete_concrete = (Volume_beam_concrete_concrete * 0.965 * ResourceUseMM_concrete_C30_37 + Volume_beam_concrete_concrete * 0.035 * ResourceUseMM_reinforcement)*2 + (Volume_column_concrete_concrete * 0.97 * ResourceUseMM_concrete_C30_37 + Volume_column_concrete_concrete * 0.03 * ResourceUseMM_reinforcement)*4

ClimateCF_concrete_C50_60 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_ClimateCF, column_name_concrete_C50_60)
ClimateCF_concrete_C30_37 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_ClimateCF, column_name_concrete_C30_37)
ClimateCF_reinforcement =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_ClimateCF, column_name_reinforcement_steel)
ClimateCF_CLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_ClimateCF, column_name_CLT)
ClimateCF_GLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_ClimateCF, column_name_GLT)
ClimateCF_structural_steel =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_ClimateCF, column_name_structural_steel)

ClimateCF_score_HCS = ClimateCF_concrete_C50_60*Volume_concrete_floor_HCS + ClimateCF_concrete_C30_37*Volume_compression_layer_HCS + ClimateCF_reinforcement*Volume_steel_floor_HCS
ClimateCF_score_in_situ_concrete = ClimateCF_concrete_C30_37*Volume_in_situ_concrete + ClimateCF_reinforcement* Volume_reinforcement_in_situ_concrete
ClimateCF_score_CLT = ClimateCF_CLT*Volume_CLT
ClimateCF_score_support_steel_HCS = (Volume_beam_steel_HCS * ClimateCF_structural_steel)*2 + (Volume_column_steel_HCS * ClimateCF_structural_steel)*4
ClimateCF_score_support_GLT_HCS = (Volume_beam_GLT_HCS * ClimateCF_GLT)*2 + (Volume_column_GLT_HCS * ClimateCF_GLT)*4
ClimateCF_score_support_concrete_HCS = (Volume_beam_concrete_HCS * 0.965 * ClimateCF_concrete_C30_37 + Volume_beam_concrete_HCS * 0.035 * ClimateCF_reinforcement)*2 + (Volume_column_concrete_HCS * 0.97 * ClimateCF_concrete_C30_37 + Volume_column_concrete_HCS * 0.03 * ClimateCF_reinforcement)*4
ClimateCF_score_support_steel_CLT = (Volume_beam_steel_CLT * ClimateCF_structural_steel)*2 + (Volume_column_steel_CLT * ClimateCF_structural_steel)*4
ClimateCF_score_support_GLT_CLT = (Volume_beam_GLT_CLT * ClimateCF_GLT)*2 + (Volume_column_GLT_CLT * ClimateCF_GLT)*4
ClimateCF_score_support_concrete_CLT = (Volume_beam_concrete_CLT * 0.965 * ClimateCF_concrete_C30_37 + Volume_beam_concrete_CLT * 0.035 * ClimateCF_reinforcement)*2 + (Volume_column_concrete_CLT * 0.97 * ClimateCF_concrete_C30_37 + Volume_column_concrete_CLT * 0.03 * ClimateCF_reinforcement)*4
ClimateCF_score_support_steel_concrete = (Volume_beam_steel_concrete * ClimateCF_structural_steel)*2 + (Volume_column_steel_concrete * ClimateCF_structural_steel)*4
ClimateCF_score_support_GLT_concrete = (Volume_beam_GLT_concrete * ClimateCF_GLT)*2 + (Volume_column_GLT_concrete * ClimateCF_GLT)*4
ClimateCF_score_support_concrete_concrete = (Volume_beam_concrete_concrete * 0.965 * ClimateCF_concrete_C30_37 + Volume_beam_concrete_concrete * 0.035 * ClimateCF_reinforcement)*2 + (Volume_column_concrete_concrete * 0.97 * ClimateCF_concrete_C30_37 + Volume_column_concrete_concrete * 0.03 * ClimateCF_reinforcement)*4

ClimateCB_concrete_C50_60 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_ClimateCB, column_name_concrete_C50_60)
ClimateCB_concrete_C30_37 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_ClimateCB, column_name_concrete_C30_37)
ClimateCB_reinforcement =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_ClimateCB, column_name_reinforcement_steel)
ClimateCB_CLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_ClimateCB, column_name_CLT)
ClimateCB_GLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_ClimateCB, column_name_GLT)
ClimateCB_structural_steel =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_ClimateCB, column_name_structural_steel)

ClimateCB_score_HCS = ClimateCB_concrete_C50_60*Volume_concrete_floor_HCS + ClimateCB_concrete_C30_37*Volume_compression_layer_HCS + ClimateCB_reinforcement*Volume_steel_floor_HCS
ClimateCB_score_in_situ_concrete = ClimateCB_concrete_C30_37*Volume_in_situ_concrete + ClimateCB_reinforcement* Volume_reinforcement_in_situ_concrete
ClimateCB_score_CLT = ClimateCB_CLT*Volume_CLT
ClimateCB_score_support_steel_HCS = (Volume_beam_steel_HCS * ClimateCB_structural_steel)*2 + (Volume_column_steel_HCS * ClimateCB_structural_steel)*4
ClimateCB_score_support_GLT_HCS = (Volume_beam_GLT_HCS * ClimateCB_GLT)*2 + (Volume_column_GLT_HCS * ClimateCB_GLT)*4
ClimateCB_score_support_concrete_HCS = (Volume_beam_concrete_HCS * 0.965 * ClimateCB_concrete_C30_37 + Volume_beam_concrete_HCS * 0.035 * ClimateCB_reinforcement)*2 + (Volume_column_concrete_HCS * 0.97 * ClimateCB_concrete_C30_37 + Volume_column_concrete_HCS * 0.03 * ClimateCB_reinforcement)*4
ClimateCB_score_support_steel_CLT = (Volume_beam_steel_CLT * ClimateCB_structural_steel)*2 + (Volume_column_steel_CLT * ClimateCB_structural_steel)*4
ClimateCB_score_support_GLT_CLT = (Volume_beam_GLT_CLT * ClimateCB_GLT)*2 + (Volume_column_GLT_CLT * ClimateCB_GLT)*4
ClimateCB_score_support_concrete_CLT = (Volume_beam_concrete_CLT * 0.965 * ClimateCB_concrete_C30_37 + Volume_beam_concrete_CLT * 0.035 * ClimateCB_reinforcement)*2 + (Volume_column_concrete_CLT * 0.97 * ClimateCB_concrete_C30_37 + Volume_column_concrete_CLT * 0.03 * ClimateCB_reinforcement)*4
ClimateCB_score_support_steel_concrete = (Volume_beam_steel_concrete * ClimateCB_structural_steel)*2 + (Volume_column_steel_concrete * ClimateCB_structural_steel)*4
ClimateCB_score_support_GLT_concrete = (Volume_beam_GLT_concrete * ClimateCB_GLT)*2 + (Volume_column_GLT_concrete * ClimateCB_GLT)*4
ClimateCB_score_support_concrete_concrete = (Volume_beam_concrete_concrete * 0.965 * ClimateCB_concrete_C30_37 + Volume_beam_concrete_concrete * 0.035 * ClimateCB_reinforcement)*2 + (Volume_column_concrete_concrete * 0.97 * ClimateCB_concrete_C30_37 + Volume_column_concrete_concrete * 0.03 * ClimateCB_reinforcement)*4

ClimateCLU_concrete_C50_60 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_ClimateCLU, column_name_concrete_C50_60)
ClimateCLU_concrete_C30_37 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_ClimateCLU, column_name_concrete_C30_37)
ClimateCLU_reinforcement =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_ClimateCLU, column_name_reinforcement_steel)
ClimateCLU_CLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_ClimateCLU, column_name_CLT)
ClimateCLU_GLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_ClimateCLU, column_name_GLT)
ClimateCLU_structural_steel =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_ClimateCLU, column_name_structural_steel)

ClimateCLU_score_HCS = ClimateCLU_concrete_C50_60*Volume_concrete_floor_HCS + ClimateCLU_concrete_C30_37*Volume_compression_layer_HCS + ClimateCLU_reinforcement*Volume_steel_floor_HCS
ClimateCLU_score_in_situ_concrete = ClimateCLU_concrete_C30_37*Volume_in_situ_concrete + ClimateCLU_reinforcement* Volume_reinforcement_in_situ_concrete
ClimateCLU_score_CLT = ClimateCLU_CLT*Volume_CLT
ClimateCLU_score_support_steel_HCS = (Volume_beam_steel_HCS * ClimateCLU_structural_steel)*2 + (Volume_column_steel_HCS * ClimateCLU_structural_steel)*4
ClimateCLU_score_support_GLT_HCS = (Volume_beam_GLT_HCS * ClimateCLU_GLT)*2 + (Volume_column_GLT_HCS * ClimateCLU_GLT)*4
ClimateCLU_score_support_concrete_HCS = (Volume_beam_concrete_HCS * 0.965 * ClimateCLU_concrete_C30_37 + Volume_beam_concrete_HCS * 0.035 * ClimateCLU_reinforcement)*2 + (Volume_column_concrete_HCS * 0.97 * ClimateCLU_concrete_C30_37 + Volume_column_concrete_HCS * 0.03 * ClimateCLU_reinforcement)*4
ClimateCLU_score_support_steel_CLT = (Volume_beam_steel_CLT * ClimateCLU_structural_steel)*2 + (Volume_column_steel_CLT * ClimateCLU_structural_steel)*4
ClimateCLU_score_support_GLT_CLT = (Volume_beam_GLT_CLT * ClimateCLU_GLT)*2 + (Volume_column_GLT_CLT * ClimateCLU_GLT)*4
ClimateCLU_score_support_concrete_CLT = (Volume_beam_concrete_CLT * 0.965 * ClimateCLU_concrete_C30_37 + Volume_beam_concrete_CLT * 0.035 * ClimateCLU_reinforcement)*2 + (Volume_column_concrete_CLT * 0.97 * ClimateCLU_concrete_C30_37 + Volume_column_concrete_CLT * 0.03 * ClimateCLU_reinforcement)*4
ClimateCLU_score_support_steel_concrete = (Volume_beam_steel_concrete * ClimateCLU_structural_steel)*2 + (Volume_column_steel_concrete * ClimateCLU_structural_steel)*4
ClimateCLU_score_support_GLT_concrete = (Volume_beam_GLT_concrete * ClimateCLU_GLT)*2 + (Volume_column_GLT_concrete * ClimateCLU_GLT)*4
ClimateCLU_score_support_concrete_concrete = (Volume_beam_concrete_concrete * 0.965 * ClimateCLU_concrete_C30_37 + Volume_beam_concrete_concrete * 0.035 * ClimateCLU_reinforcement)*2 + (Volume_column_concrete_concrete * 0.97 * ClimateCLU_concrete_C30_37 + Volume_column_concrete_concrete * 0.03 * ClimateCLU_reinforcement)*4

HumanTNCO_concrete_C50_60 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_HumanTNCO, column_name_concrete_C50_60)
HumanTNCO_concrete_C30_37 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_HumanTNCO, column_name_concrete_C30_37)
HumanTNCO_reinforcement =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_HumanTNCO, column_name_reinforcement_steel)
HumanTNCO_CLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_HumanTNCO, column_name_CLT)
HumanTNCO_GLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_HumanTNCO, column_name_GLT)
HumanTNCO_structural_steel =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_HumanTNCO, column_name_structural_steel)

HumanTNCO_score_HCS = HumanTNCO_concrete_C50_60*Volume_concrete_floor_HCS + HumanTNCO_concrete_C30_37*Volume_compression_layer_HCS + HumanTNCO_reinforcement*Volume_steel_floor_HCS
HumanTNCO_score_in_situ_concrete = HumanTNCO_concrete_C30_37*Volume_in_situ_concrete + HumanTNCO_reinforcement* Volume_reinforcement_in_situ_concrete
HumanTNCO_score_CLT = HumanTNCO_CLT*Volume_CLT
HumanTNCO_score_support_steel_HCS = (Volume_beam_steel_HCS * HumanTNCO_structural_steel)*2 + (Volume_column_steel_HCS * HumanTNCO_structural_steel)*4
HumanTNCO_score_support_GLT_HCS = (Volume_beam_GLT_HCS * HumanTNCO_GLT)*2 + (Volume_column_GLT_HCS * HumanTNCO_GLT)*4
HumanTNCO_score_support_concrete_HCS = (Volume_beam_concrete_HCS * 0.965 * HumanTNCO_concrete_C30_37 + Volume_beam_concrete_HCS * 0.035 * HumanTNCO_reinforcement)*2 + (Volume_column_concrete_HCS * 0.97 * HumanTNCO_concrete_C30_37 + Volume_column_concrete_HCS * 0.03 * HumanTNCO_reinforcement)*4
HumanTNCO_score_support_steel_CLT = (Volume_beam_steel_CLT * HumanTNCO_structural_steel)*2 + (Volume_column_steel_CLT * HumanTNCO_structural_steel)*4
HumanTNCO_score_support_GLT_CLT = (Volume_beam_GLT_CLT * HumanTNCO_GLT)*2 + (Volume_column_GLT_CLT * HumanTNCO_GLT)*4
HumanTNCO_score_support_concrete_CLT = (Volume_beam_concrete_CLT * 0.965 * HumanTNCO_concrete_C30_37 + Volume_beam_concrete_CLT * 0.035 * HumanTNCO_reinforcement)*2 + (Volume_column_concrete_CLT * 0.97 * HumanTNCO_concrete_C30_37 + Volume_column_concrete_CLT * 0.03 * HumanTNCO_reinforcement)*4
HumanTNCO_score_support_steel_concrete = (Volume_beam_steel_concrete * HumanTNCO_structural_steel)*2 + (Volume_column_steel_concrete * HumanTNCO_structural_steel)*4
HumanTNCO_score_support_GLT_concrete = (Volume_beam_GLT_concrete * HumanTNCO_GLT)*2 + (Volume_column_GLT_concrete * HumanTNCO_GLT)*4
HumanTNCO_score_support_concrete_concrete = (Volume_beam_concrete_concrete * 0.965 * HumanTNCO_concrete_C30_37 + Volume_beam_concrete_concrete * 0.035 * HumanTNCO_reinforcement)*2 + (Volume_column_concrete_concrete * 0.97 * HumanTNCO_concrete_C30_37 + Volume_column_concrete_concrete * 0.03 * HumanTNCO_reinforcement)*4

HumanTNCI_concrete_C50_60 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_HumanTNCI, column_name_concrete_C50_60)
HumanTNCI_concrete_C30_37 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_HumanTNCI, column_name_concrete_C30_37)
HumanTNCI_reinforcement =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_HumanTNCI, column_name_reinforcement_steel)
HumanTNCI_CLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_HumanTNCI, column_name_CLT)
HumanTNCI_GLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_HumanTNCI, column_name_GLT)
HumanTNCI_structural_steel =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_HumanTNCI, column_name_structural_steel)

HumanTNCI_score_HCS = HumanTNCI_concrete_C50_60*Volume_concrete_floor_HCS + HumanTNCI_concrete_C30_37*Volume_compression_layer_HCS + HumanTNCI_reinforcement*Volume_steel_floor_HCS
HumanTNCI_score_in_situ_concrete = HumanTNCI_concrete_C30_37*Volume_in_situ_concrete + HumanTNCI_reinforcement* Volume_reinforcement_in_situ_concrete
HumanTNCI_score_CLT = HumanTNCI_CLT*Volume_CLT
HumanTNCI_score_support_steel_HCS = (Volume_beam_steel_HCS * HumanTNCI_structural_steel)*2 + (Volume_column_steel_HCS * HumanTNCI_structural_steel)*4
HumanTNCI_score_support_GLT_HCS = (Volume_beam_GLT_HCS * HumanTNCI_GLT)*2 + (Volume_column_GLT_HCS * HumanTNCI_GLT)*4
HumanTNCI_score_support_concrete_HCS = (Volume_beam_concrete_HCS * 0.965 * HumanTNCI_concrete_C30_37 + Volume_beam_concrete_HCS * 0.035 * HumanTNCI_reinforcement)*2 + (Volume_column_concrete_HCS * 0.97 * HumanTNCI_concrete_C30_37 + Volume_column_concrete_HCS * 0.03 * HumanTNCI_reinforcement)*4
HumanTNCI_score_support_steel_CLT = (Volume_beam_steel_CLT * HumanTNCI_structural_steel)*2 + (Volume_column_steel_CLT * HumanTNCI_structural_steel)*4
HumanTNCI_score_support_GLT_CLT = (Volume_beam_GLT_CLT * HumanTNCI_GLT)*2 + (Volume_column_GLT_CLT * HumanTNCI_GLT)*4
HumanTNCI_score_support_concrete_CLT = (Volume_beam_concrete_CLT * 0.965 * HumanTNCI_concrete_C30_37 + Volume_beam_concrete_CLT * 0.035 * HumanTNCI_reinforcement)*2 + (Volume_column_concrete_CLT * 0.97 * HumanTNCI_concrete_C30_37 + Volume_column_concrete_CLT * 0.03 * HumanTNCI_reinforcement)*4
HumanTNCI_score_support_steel_concrete = (Volume_beam_steel_concrete * HumanTNCI_structural_steel)*2 + (Volume_column_steel_concrete * HumanTNCI_structural_steel)*4
HumanTNCI_score_support_GLT_concrete = (Volume_beam_GLT_concrete * HumanTNCI_GLT)*2 + (Volume_column_GLT_concrete * HumanTNCI_GLT)*4
HumanTNCI_score_support_concrete_concrete = (Volume_beam_concrete_concrete * 0.965 * HumanTNCI_concrete_C30_37 + Volume_beam_concrete_concrete * 0.035 * HumanTNCI_reinforcement)*2 + (Volume_column_concrete_concrete * 0.97 * HumanTNCI_concrete_C30_37 + Volume_column_concrete_concrete * 0.03 * HumanTNCI_reinforcement)*4

HumanTNCM_concrete_C50_60 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_HumanTNCM, column_name_concrete_C50_60)
HumanTNCM_concrete_C30_37 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_HumanTNCM, column_name_concrete_C30_37)
HumanTNCM_reinforcement =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_HumanTNCM, column_name_reinforcement_steel)
HumanTNCM_CLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_HumanTNCM, column_name_CLT)
HumanTNCM_GLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_HumanTNCM, column_name_GLT)
HumanTNCM_structural_steel =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_HumanTNCM, column_name_structural_steel)

HumanTNCM_score_HCS = HumanTNCM_concrete_C50_60*Volume_concrete_floor_HCS + HumanTNCM_concrete_C30_37*Volume_compression_layer_HCS + HumanTNCM_reinforcement*Volume_steel_floor_HCS
HumanTNCM_score_in_situ_concrete = HumanTNCM_concrete_C30_37*Volume_in_situ_concrete + HumanTNCM_reinforcement* Volume_reinforcement_in_situ_concrete
HumanTNCM_score_CLT = HumanTNCM_CLT*Volume_CLT
HumanTNCM_score_support_steel_HCS = (Volume_beam_steel_HCS * HumanTNCM_structural_steel)*2 + (Volume_column_steel_HCS * HumanTNCM_structural_steel)*4
HumanTNCM_score_support_GLT_HCS = (Volume_beam_GLT_HCS * HumanTNCM_GLT)*2 + (Volume_column_GLT_HCS * HumanTNCM_GLT)*4
HumanTNCM_score_support_concrete_HCS = (Volume_beam_concrete_HCS * 0.965 * HumanTNCM_concrete_C30_37 + Volume_beam_concrete_HCS * 0.035 * HumanTNCM_reinforcement)*2 + (Volume_column_concrete_HCS * 0.97 * HumanTNCM_concrete_C30_37 + Volume_column_concrete_HCS * 0.03 * HumanTNCM_reinforcement)*4
HumanTNCM_score_support_steel_CLT = (Volume_beam_steel_CLT * HumanTNCM_structural_steel)*2 + (Volume_column_steel_CLT * HumanTNCM_structural_steel)*4
HumanTNCM_score_support_GLT_CLT = (Volume_beam_GLT_CLT * HumanTNCM_GLT)*2 + (Volume_column_GLT_CLT * HumanTNCM_GLT)*4
HumanTNCM_score_support_concrete_CLT = (Volume_beam_concrete_CLT * 0.965 * HumanTNCM_concrete_C30_37 + Volume_beam_concrete_CLT * 0.035 * HumanTNCM_reinforcement)*2 + (Volume_column_concrete_CLT * 0.97 * HumanTNCM_concrete_C30_37 + Volume_column_concrete_CLT * 0.03 * HumanTNCM_reinforcement)*4
HumanTNCM_score_support_steel_concrete = (Volume_beam_steel_concrete * HumanTNCM_structural_steel)*2 + (Volume_column_steel_concrete * HumanTNCM_structural_steel)*4
HumanTNCM_score_support_GLT_concrete = (Volume_beam_GLT_concrete * HumanTNCM_GLT)*2 + (Volume_column_GLT_concrete * HumanTNCM_GLT)*4
HumanTNCM_score_support_concrete_concrete = (Volume_beam_concrete_concrete * 0.965 * HumanTNCM_concrete_C30_37 + Volume_beam_concrete_concrete * 0.035 * HumanTNCM_reinforcement)*2 + (Volume_column_concrete_concrete * 0.97 * HumanTNCM_concrete_C30_37 + Volume_column_concrete_concrete * 0.03 * HumanTNCM_reinforcement)*4

HumanTCO_concrete_C50_60 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_HumanTCO, column_name_concrete_C50_60)
HumanTCO_concrete_C30_37 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_HumanTCO, column_name_concrete_C30_37)
HumanTCO_reinforcement =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_HumanTCO, column_name_reinforcement_steel)
HumanTCO_CLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_HumanTCO, column_name_CLT)
HumanTCO_GLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_HumanTCO, column_name_GLT)
HumanTCO_structural_steel =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_HumanTCO, column_name_structural_steel)

HumanTCO_score_HCS = HumanTCO_concrete_C50_60*Volume_concrete_floor_HCS + HumanTCO_concrete_C30_37*Volume_compression_layer_HCS + HumanTCO_reinforcement*Volume_steel_floor_HCS
HumanTCO_score_in_situ_concrete = HumanTCO_concrete_C30_37*Volume_in_situ_concrete + HumanTCO_reinforcement* Volume_reinforcement_in_situ_concrete
HumanTCO_score_CLT = HumanTCO_CLT*Volume_CLT
HumanTCO_score_support_steel_HCS = (Volume_beam_steel_HCS * HumanTCO_structural_steel)*2 + (Volume_column_steel_HCS * HumanTCO_structural_steel)*4
HumanTCO_score_support_GLT_HCS = (Volume_beam_GLT_HCS * HumanTCO_GLT)*2 + (Volume_column_GLT_HCS * HumanTCO_GLT)*4
HumanTCO_score_support_concrete_HCS = (Volume_beam_concrete_HCS * 0.965 * HumanTCO_concrete_C30_37 + Volume_beam_concrete_HCS * 0.035 * HumanTCO_reinforcement)*2 + (Volume_column_concrete_HCS * 0.97 * HumanTCO_concrete_C30_37 + Volume_column_concrete_HCS * 0.03 * HumanTCO_reinforcement)*4
HumanTCO_score_support_steel_CLT = (Volume_beam_steel_CLT * HumanTCO_structural_steel)*2 + (Volume_column_steel_CLT * HumanTCO_structural_steel)*4
HumanTCO_score_support_GLT_CLT = (Volume_beam_GLT_CLT * HumanTCO_GLT)*2 + (Volume_column_GLT_CLT * HumanTCO_GLT)*4
HumanTCO_score_support_concrete_CLT = (Volume_beam_concrete_CLT * 0.965 * HumanTCO_concrete_C30_37 + Volume_beam_concrete_CLT * 0.035 * HumanTCO_reinforcement)*2 + (Volume_column_concrete_CLT * 0.97 * HumanTCO_concrete_C30_37 + Volume_column_concrete_CLT * 0.03 * HumanTCO_reinforcement)*4
HumanTCO_score_support_steel_concrete = (Volume_beam_steel_concrete * HumanTCO_structural_steel)*2 + (Volume_column_steel_concrete * HumanTCO_structural_steel)*4
HumanTCO_score_support_GLT_concrete = (Volume_beam_GLT_concrete * HumanTCO_GLT)*2 + (Volume_column_GLT_concrete * HumanTCO_GLT)*4
HumanTCO_score_support_concrete_concrete = (Volume_beam_concrete_concrete * 0.965 * HumanTCO_concrete_C30_37 + Volume_beam_concrete_concrete * 0.035 * HumanTCO_reinforcement)*2 + (Volume_column_concrete_concrete * 0.97 * HumanTCO_concrete_C30_37 + Volume_column_concrete_concrete * 0.03 * HumanTCO_reinforcement)*4

HumanTCI_concrete_C50_60 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_HumanTCI, column_name_concrete_C50_60)
HumanTCI_concrete_C30_37 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_HumanTCI, column_name_concrete_C30_37)
HumanTCI_reinforcement =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_HumanTCI, column_name_reinforcement_steel)
HumanTCI_CLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_HumanTCI, column_name_CLT)
HumanTCI_GLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_HumanTCI, column_name_GLT)
HumanTCI_structural_steel =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_HumanTCI, column_name_structural_steel)

HumanTCI_score_HCS = HumanTCI_concrete_C50_60*Volume_concrete_floor_HCS + HumanTCI_concrete_C30_37*Volume_compression_layer_HCS + HumanTCI_reinforcement*Volume_steel_floor_HCS
HumanTCI_score_in_situ_concrete = HumanTCI_concrete_C30_37*Volume_in_situ_concrete + HumanTCI_reinforcement* Volume_reinforcement_in_situ_concrete
HumanTCI_score_CLT = HumanTCI_CLT*Volume_CLT
HumanTCI_score_support_steel_HCS = (Volume_beam_steel_HCS * HumanTCI_structural_steel)*2 + (Volume_column_steel_HCS * HumanTCI_structural_steel)*4
HumanTCI_score_support_GLT_HCS = (Volume_beam_GLT_HCS * HumanTCI_GLT)*2 + (Volume_column_GLT_HCS * HumanTCI_GLT)*4
HumanTCI_score_support_concrete_HCS = (Volume_beam_concrete_HCS * 0.965 * HumanTCI_concrete_C30_37 + Volume_beam_concrete_HCS * 0.035 * HumanTCI_reinforcement)*2 + (Volume_column_concrete_HCS * 0.97 * HumanTCI_concrete_C30_37 + Volume_column_concrete_HCS * 0.03 * HumanTCI_reinforcement)*4
HumanTCI_score_support_steel_CLT = (Volume_beam_steel_CLT * HumanTCI_structural_steel)*2 + (Volume_column_steel_CLT * HumanTCI_structural_steel)*4
HumanTCI_score_support_GLT_CLT = (Volume_beam_GLT_CLT * HumanTCI_GLT)*2 + (Volume_column_GLT_CLT * HumanTCI_GLT)*4
HumanTCI_score_support_concrete_CLT = (Volume_beam_concrete_CLT * 0.965 * HumanTCI_concrete_C30_37 + Volume_beam_concrete_CLT * 0.035 * HumanTCI_reinforcement)*2 + (Volume_column_concrete_CLT * 0.97 * HumanTCI_concrete_C30_37 + Volume_column_concrete_CLT * 0.03 * HumanTCI_reinforcement)*4
HumanTCI_score_support_steel_concrete = (Volume_beam_steel_concrete * HumanTCI_structural_steel)*2 + (Volume_column_steel_concrete * HumanTCI_structural_steel)*4
HumanTCI_score_support_GLT_concrete = (Volume_beam_GLT_concrete * HumanTCI_GLT)*2 + (Volume_column_GLT_concrete * HumanTCI_GLT)*4
HumanTCI_score_support_concrete_concrete = (Volume_beam_concrete_concrete * 0.965 * HumanTCI_concrete_C30_37 + Volume_beam_concrete_concrete * 0.035 * HumanTCI_reinforcement)*2 + (Volume_column_concrete_concrete * 0.97 * HumanTCI_concrete_C30_37 + Volume_column_concrete_concrete * 0.03 * HumanTCI_reinforcement)*4

HumanTCM_concrete_C50_60 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_HumanTCM, column_name_concrete_C50_60)
HumanTCM_concrete_C30_37 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_HumanTCM, column_name_concrete_C30_37)
HumanTCM_reinforcement =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_HumanTCM, column_name_reinforcement_steel)
HumanTCM_CLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_HumanTCM, column_name_CLT)
HumanTCM_GLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_HumanTCM, column_name_GLT)
HumanTCM_structural_steel =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_HumanTCM, column_name_structural_steel)

HumanTCM_score_HCS = HumanTCM_concrete_C50_60*Volume_concrete_floor_HCS + HumanTCM_concrete_C30_37*Volume_compression_layer_HCS + HumanTCM_reinforcement*Volume_steel_floor_HCS
HumanTCM_score_in_situ_concrete = HumanTCM_concrete_C30_37*Volume_in_situ_concrete + HumanTCM_reinforcement* Volume_reinforcement_in_situ_concrete
HumanTCM_score_CLT = HumanTCM_CLT*Volume_CLT
HumanTCM_score_support_steel_HCS = (Volume_beam_steel_HCS * HumanTCM_structural_steel)*2 + (Volume_column_steel_HCS * HumanTCM_structural_steel)*4
HumanTCM_score_support_GLT_HCS = (Volume_beam_GLT_HCS * HumanTCM_GLT)*2 + (Volume_column_GLT_HCS * HumanTCM_GLT)*4
HumanTCM_score_support_concrete_HCS = (Volume_beam_concrete_HCS * 0.965 * HumanTCM_concrete_C30_37 + Volume_beam_concrete_HCS * 0.035 * HumanTCM_reinforcement)*2 + (Volume_column_concrete_HCS * 0.97 * HumanTCM_concrete_C30_37 + Volume_column_concrete_HCS * 0.03 * HumanTCM_reinforcement)*4
HumanTCM_score_support_steel_CLT = (Volume_beam_steel_CLT * HumanTCM_structural_steel)*2 + (Volume_column_steel_CLT * HumanTCM_structural_steel)*4
HumanTCM_score_support_GLT_CLT = (Volume_beam_GLT_CLT * HumanTCM_GLT)*2 + (Volume_column_GLT_CLT * HumanTCM_GLT)*4
HumanTCM_score_support_concrete_CLT = (Volume_beam_concrete_CLT * 0.965 * HumanTCM_concrete_C30_37 + Volume_beam_concrete_CLT * 0.035 * HumanTCM_reinforcement)*2 + (Volume_column_concrete_CLT * 0.97 * HumanTCM_concrete_C30_37 + Volume_column_concrete_CLT * 0.03 * HumanTCM_reinforcement)*4
HumanTCM_score_support_steel_concrete = (Volume_beam_steel_concrete * HumanTCM_structural_steel)*2 + (Volume_column_steel_concrete * HumanTCM_structural_steel)*4
HumanTCM_score_support_GLT_concrete = (Volume_beam_GLT_concrete * HumanTCM_GLT)*2 + (Volume_column_GLT_concrete * HumanTCM_GLT)*4
HumanTCM_score_support_concrete_concrete = (Volume_beam_concrete_concrete * 0.965 * HumanTCM_concrete_C30_37 + Volume_beam_concrete_concrete * 0.035 * HumanTCM_reinforcement)*2 + (Volume_column_concrete_concrete * 0.97 * HumanTCM_concrete_C30_37 + Volume_column_concrete_concrete * 0.03 * HumanTCM_reinforcement)*4

ExotoxicityFO_concrete_C50_60 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_ExotoxicityFO, column_name_concrete_C50_60)
ExotoxicityFO_concrete_C30_37 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_ExotoxicityFO, column_name_concrete_C30_37)
ExotoxicityFO_reinforcement =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_ExotoxicityFO, column_name_reinforcement_steel)
ExotoxicityFO_CLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_ExotoxicityFO, column_name_CLT)
ExotoxicityFO_GLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_ExotoxicityFO, column_name_GLT)
ExotoxicityFO_structural_steel =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_ExotoxicityFO, column_name_structural_steel)

ExotoxicityFO_score_HCS = ExotoxicityFO_concrete_C50_60*Volume_concrete_floor_HCS + ExotoxicityFO_concrete_C30_37*Volume_compression_layer_HCS + ExotoxicityFO_reinforcement*Volume_steel_floor_HCS
ExotoxicityFO_score_in_situ_concrete = ExotoxicityFO_concrete_C30_37*Volume_in_situ_concrete + ExotoxicityFO_reinforcement* Volume_reinforcement_in_situ_concrete
ExotoxicityFO_score_CLT = ExotoxicityFO_CLT*Volume_CLT
ExotoxicityFO_score_support_steel_HCS = (Volume_beam_steel_HCS * ExotoxicityFO_structural_steel)*2 + (Volume_column_steel_HCS * ExotoxicityFO_structural_steel)*4
ExotoxicityFO_score_support_GLT_HCS = (Volume_beam_GLT_HCS * ExotoxicityFO_GLT)*2 + (Volume_column_GLT_HCS * ExotoxicityFO_GLT)*4
ExotoxicityFO_score_support_concrete_HCS = (Volume_beam_concrete_HCS * 0.965 * ExotoxicityFO_concrete_C30_37 + Volume_beam_concrete_HCS * 0.035 * ExotoxicityFO_reinforcement)*2 + (Volume_column_concrete_HCS * 0.97 * ExotoxicityFO_concrete_C30_37 + Volume_column_concrete_HCS * 0.03 * ExotoxicityFO_reinforcement)*4
ExotoxicityFO_score_support_steel_CLT = (Volume_beam_steel_CLT * ExotoxicityFO_structural_steel)*2 + (Volume_column_steel_CLT * ExotoxicityFO_structural_steel)*4
ExotoxicityFO_score_support_GLT_CLT = (Volume_beam_GLT_CLT * ExotoxicityFO_GLT)*2 + (Volume_column_GLT_CLT * ExotoxicityFO_GLT)*4
ExotoxicityFO_score_support_concrete_CLT = (Volume_beam_concrete_CLT * 0.965 * ExotoxicityFO_concrete_C30_37 + Volume_beam_concrete_CLT * 0.035 * ExotoxicityFO_reinforcement)*2 + (Volume_column_concrete_CLT * 0.97 * ExotoxicityFO_concrete_C30_37 + Volume_column_concrete_CLT * 0.03 * ExotoxicityFO_reinforcement)*4
ExotoxicityFO_score_support_steel_concrete = (Volume_beam_steel_concrete * ExotoxicityFO_structural_steel)*2 + (Volume_column_steel_concrete * ExotoxicityFO_structural_steel)*4
ExotoxicityFO_score_support_GLT_concrete = (Volume_beam_GLT_concrete * ExotoxicityFO_GLT)*2 + (Volume_column_GLT_concrete * ExotoxicityFO_GLT)*4
ExotoxicityFO_score_support_concrete_concrete = (Volume_beam_concrete_concrete * 0.965 * ExotoxicityFO_concrete_C30_37 + Volume_beam_concrete_concrete * 0.035 * ExotoxicityFO_reinforcement)*2 + (Volume_column_concrete_concrete * 0.97 * ExotoxicityFO_concrete_C30_37 + Volume_column_concrete_concrete * 0.03 * ExotoxicityFO_reinforcement)*4

ExotoxicityFI_concrete_C50_60 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_ExotoxicityFI, column_name_concrete_C50_60)
ExotoxicityFI_concrete_C30_37 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_ExotoxicityFI, column_name_concrete_C30_37)
ExotoxicityFI_reinforcement =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_ExotoxicityFI, column_name_reinforcement_steel)
ExotoxicityFI_CLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_ExotoxicityFI, column_name_CLT)
ExotoxicityFI_GLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_ExotoxicityFI, column_name_GLT)
ExotoxicityFI_structural_steel =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_ExotoxicityFI, column_name_structural_steel)

ExotoxicityFI_score_HCS = ExotoxicityFI_concrete_C50_60*Volume_concrete_floor_HCS + ExotoxicityFI_concrete_C30_37*Volume_compression_layer_HCS + ExotoxicityFI_reinforcement*Volume_steel_floor_HCS
ExotoxicityFI_score_in_situ_concrete = ExotoxicityFI_concrete_C30_37*Volume_in_situ_concrete + ExotoxicityFI_reinforcement* Volume_reinforcement_in_situ_concrete
ExotoxicityFI_score_CLT = ExotoxicityFI_CLT*Volume_CLT
ExotoxicityFI_score_support_steel_HCS = (Volume_beam_steel_HCS * ExotoxicityFI_structural_steel)*2 + (Volume_column_steel_HCS * ExotoxicityFI_structural_steel)*4
ExotoxicityFI_score_support_GLT_HCS = (Volume_beam_GLT_HCS * ExotoxicityFI_GLT)*2 + (Volume_column_GLT_HCS * ExotoxicityFI_GLT)*4
ExotoxicityFI_score_support_concrete_HCS = (Volume_beam_concrete_HCS * 0.965 * ExotoxicityFI_concrete_C30_37 + Volume_beam_concrete_HCS * 0.035 * ExotoxicityFI_reinforcement)*2 + (Volume_column_concrete_HCS * 0.97 * ExotoxicityFI_concrete_C30_37 + Volume_column_concrete_HCS * 0.03 * ExotoxicityFI_reinforcement)*4
ExotoxicityFI_score_support_steel_CLT = (Volume_beam_steel_CLT * ExotoxicityFI_structural_steel)*2 + (Volume_column_steel_CLT * ExotoxicityFI_structural_steel)*4
ExotoxicityFI_score_support_GLT_CLT = (Volume_beam_GLT_CLT * ExotoxicityFI_GLT)*2 + (Volume_column_GLT_CLT * ExotoxicityFI_GLT)*4
ExotoxicityFI_score_support_concrete_CLT = (Volume_beam_concrete_CLT * 0.965 * ExotoxicityFI_concrete_C30_37 + Volume_beam_concrete_CLT * 0.035 * ExotoxicityFI_reinforcement)*2 + (Volume_column_concrete_CLT * 0.97 * ExotoxicityFI_concrete_C30_37 + Volume_column_concrete_CLT * 0.03 * ExotoxicityFI_reinforcement)*4
ExotoxicityFI_score_support_steel_concrete = (Volume_beam_steel_concrete * ExotoxicityFI_structural_steel)*2 + (Volume_column_steel_concrete * ExotoxicityFI_structural_steel)*4
ExotoxicityFI_score_support_GLT_concrete = (Volume_beam_GLT_concrete * ExotoxicityFI_GLT)*2 + (Volume_column_GLT_concrete * ExotoxicityFI_GLT)*4
ExotoxicityFI_score_support_concrete_concrete = (Volume_beam_concrete_concrete * 0.965 * ExotoxicityFI_concrete_C30_37 + Volume_beam_concrete_concrete * 0.035 * ExotoxicityFI_reinforcement)*2 + (Volume_column_concrete_concrete * 0.97 * ExotoxicityFI_concrete_C30_37 + Volume_column_concrete_concrete * 0.03 * ExotoxicityFI_reinforcement)*4

ExotoxicityFM_concrete_C50_60 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_ExotoxicityFM, column_name_concrete_C50_60)
ExotoxicityFM_concrete_C30_37 =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_ExotoxicityFM, column_name_concrete_C30_37)
ExotoxicityFM_reinforcement =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_ExotoxicityFM, column_name_reinforcement_steel)
ExotoxicityFM_CLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_ExotoxicityFM, column_name_CLT)
ExotoxicityFM_GLT =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_ExotoxicityFM, column_name_GLT)
ExotoxicityFM_structural_steel =  retrieve_value(excel_file_LCA, sheet_name_LCA, row_name_ExotoxicityFM, column_name_structural_steel)

ExotoxicityFM_score_HCS = ExotoxicityFM_concrete_C50_60*Volume_concrete_floor_HCS + ExotoxicityFM_concrete_C30_37*Volume_compression_layer_HCS + ExotoxicityFM_reinforcement*Volume_steel_floor_HCS
ExotoxicityFM_score_in_situ_concrete = ExotoxicityFM_concrete_C30_37*Volume_in_situ_concrete + ExotoxicityFM_reinforcement* Volume_reinforcement_in_situ_concrete
ExotoxicityFM_score_CLT = ExotoxicityFM_CLT*Volume_CLT
ExotoxicityFM_score_support_steel_HCS = (Volume_beam_steel_HCS * ExotoxicityFM_structural_steel)*2 + (Volume_column_steel_HCS * ExotoxicityFM_structural_steel)*4
ExotoxicityFM_score_support_GLT_HCS = (Volume_beam_GLT_HCS * ExotoxicityFM_GLT)*2 + (Volume_column_GLT_HCS * ExotoxicityFM_GLT)*4
ExotoxicityFM_score_support_concrete_HCS = (Volume_beam_concrete_HCS * 0.965 * ExotoxicityFM_concrete_C30_37 + Volume_beam_concrete_HCS * 0.035 * ExotoxicityFM_reinforcement)*2 + (Volume_column_concrete_HCS * 0.97 * ExotoxicityFM_concrete_C30_37 + Volume_column_concrete_HCS * 0.03 * ExotoxicityFM_reinforcement)*4
ExotoxicityFM_score_support_steel_CLT = (Volume_beam_steel_CLT * ExotoxicityFM_structural_steel)*2 + (Volume_column_steel_CLT * ExotoxicityFM_structural_steel)*4
ExotoxicityFM_score_support_GLT_CLT = (Volume_beam_GLT_CLT * ExotoxicityFM_GLT)*2 + (Volume_column_GLT_CLT * ExotoxicityFM_GLT)*4
ExotoxicityFM_score_support_concrete_CLT = (Volume_beam_concrete_CLT * 0.965 * ExotoxicityFM_concrete_C30_37 + Volume_beam_concrete_CLT * 0.035 * ExotoxicityFM_reinforcement)*2 + (Volume_column_concrete_CLT * 0.97 * ExotoxicityFM_concrete_C30_37 + Volume_column_concrete_CLT * 0.03 * ExotoxicityFM_reinforcement)*4
ExotoxicityFM_score_support_steel_concrete = (Volume_beam_steel_concrete * ExotoxicityFM_structural_steel)*2 + (Volume_column_steel_concrete * ExotoxicityFM_structural_steel)*4
ExotoxicityFM_score_support_GLT_concrete = (Volume_beam_GLT_concrete * ExotoxicityFM_GLT)*2 + (Volume_column_GLT_concrete * ExotoxicityFM_GLT)*4
ExotoxicityFM_score_support_concrete_concrete = (Volume_beam_concrete_concrete * 0.965 * ExotoxicityFM_concrete_C30_37 + Volume_beam_concrete_concrete * 0.035 * ExotoxicityFM_reinforcement)*2 + (Volume_column_concrete_concrete * 0.97 * ExotoxicityFM_concrete_C30_37 + Volume_column_concrete_concrete * 0.03 * ExotoxicityFM_reinforcement)*4




def create_and_export_data(output_excel):
    data = {
        'Impact category': ['Type or thickness', 'milipoint', 'Climate change', 'Ozone depletion', 'Ionising radiation', 'Photochemical ozone formation', 'Particulate matter', 'Human toxicity, non-cancer', 'Human toxicity, cancer', 'Acidification', 'Eutrophication, freshwater', 'Eutrophication, marine', 'Eutrophication, terrestrial', 'Ecotoxicity, freshwater', 'Land use', 'Water use', 'Resource use, fossils', 'Resource use, minerals and metals', 'Climate change - Fossil', 'Climate change - Biogenic', 'Climate change - Land use and LU change', 'Human toxicity, non-cancer - organics', 'Human toxicity, non-cancer - inorganics', 'Human toxicity, non-cancer - metals', 'Human toxicity, cancer - organics', 'Human toxicity, cancer - inorganics', 'Human toxicity, cancer - metals', 'Ecotoxicity, freshwater - organics', 'Ecotoxicity, freshwater - inorganics', 'Ecotoxicity, freshwater - metals'],
        'Unit': ['m', 'mPt', 'kg CO2 eq', 'kg CFC11 eq', 'kBq U-235 eq', 'kg NMVOC eq', 'disease inc.', 'CTUh', 'CTUh', 'mol H+ eq', 'kg P eq', 'kg N eq', 'mol N eq', 'CTUe', 'Pt', 'm3 depriv.', 'MJ', 'kg Sb eq', 'kg CO2 eq', 'kg CO2 eq', 'kg CO2 eq', 'CTUh', 'CTUh', 'CTUh', 'CTUh', 'CTUh', 'CTUh', 'CTUe', 'CTUe', 'CTUe'],
        'Hollow-core slab floor': [type_HCS, mPt_score_HCS, GWP_score_HCS, OzoneD_score_HCS, IonisingR_score_HCS, photochemicalOF_score_HCS, ParticulatM_score_HCS, HumanTNC_score_HCS, HumanTC_score_HCS, Acidification_score_HCS, EutrophicationF_score_HCS, EutrophicationM_score_HCS, EutrophicationT_score_HCS, EcotoxicityF_score_HCS, LandUse_score_HCS, WaterUse_score_HCS, ResourceUseF_score_HCS, ResourceUseMM_score_HCS, ClimateCF_score_HCS, ClimateCB_score_HCS, ClimateCLU_score_HCS, HumanTNCO_score_HCS,HumanTNCI_score_HCS, HumanTNCM_score_HCS, HumanTCO_score_HCS, HumanTCI_score_HCS, HumanTCM_score_HCS, ExotoxicityFO_score_HCS, ExotoxicityFI_score_HCS, ExotoxicityFM_score_HCS],
        'Cross-laminated timber floor': [type_floor_CLT, mPt_score_CLT, GWP_score_CLT, OzoneD_score_CLT, IonisingR_score_CLT, photochemicalOF_score_CLT, ParticulatM_score_CLT, HumanTNC_score_CLT, HumanTC_score_CLT,  Acidification_score_CLT, EutrophicationF_score_CLT, EutrophicationM_score_CLT, EutrophicationT_score_CLT, EcotoxicityF_score_CLT, LandUse_score_CLT, WaterUse_score_CLT, ResourceUseF_score_CLT, ResourceUseMM_score_CLT, ClimateCF_score_CLT, ClimateCB_score_CLT, ClimateCLU_score_CLT, HumanTNCO_score_CLT,HumanTNCI_score_CLT, HumanTNCM_score_CLT, HumanTCO_score_CLT, HumanTCI_score_CLT, HumanTCM_score_CLT, ExotoxicityFO_score_CLT, ExotoxicityFI_score_CLT, ExotoxicityFM_score_CLT],
        'Cast-in-situ concrete floor': [thickness_concrete_rounded, mPt_score_in_situ_concrete, GWP_score_in_situ_concrete, OzoneD_score_in_situ_concrete, IonisingR_score_in_situ_concrete, photochemicalOF_score_in_situ_concrete, ParticulatM_score_in_situ_concrete, HumanTNC_score_in_situ_concrete, HumanTC_score_in_situ_concrete,  Acidification_score_in_situ_concrete, EutrophicationF_score_in_situ_concrete, EutrophicationM_score_in_situ_concrete, EutrophicationT_score_in_situ_concrete, EcotoxicityF_score_in_situ_concrete, LandUse_score_in_situ_concrete, WaterUse_score_in_situ_concrete, ResourceUseF_score_in_situ_concrete, ResourceUseMM_score_in_situ_concrete, ClimateCF_score_in_situ_concrete, ClimateCB_score_in_situ_concrete, ClimateCLU_score_in_situ_concrete, HumanTNCO_score_in_situ_concrete,HumanTNCI_score_in_situ_concrete, HumanTNCM_score_in_situ_concrete, HumanTCO_score_in_situ_concrete, HumanTCI_score_in_situ_concrete, HumanTCM_score_in_situ_concrete, ExotoxicityFO_score_in_situ_concrete, ExotoxicityFI_score_in_situ_concrete, ExotoxicityFM_score_in_situ_concrete],
        'HCS - Steel supports': [f"Beams: {profile_beam_steel_HCS} and columns: {profile_column_steel_HCS}", mPt_score_support_steel_HCS, GWP_score_support_steel_HCS, OzoneD_score_support_steel_HCS, IonisingR_score_support_steel_HCS, photochemicalOF_score_support_steel_HCS, ParticulatM_score_support_steel_HCS, HumanTNC_score_support_steel_HCS, HumanTC_score_support_steel_HCS, Acidification_score_support_steel_HCS, EutrophicationF_score_support_steel_HCS, EutrophicationM_score_support_steel_HCS, EutrophicationT_score_support_steel_HCS, EcotoxicityF_score_support_steel_HCS, LandUse_score_support_steel_HCS, WaterUse_score_support_steel_HCS, ResourceUseF_score_support_steel_HCS, ResourceUseMM_score_support_steel_HCS, ClimateCF_score_support_steel_HCS, ClimateCB_score_support_steel_HCS, ClimateCLU_score_support_steel_HCS, HumanTNCO_score_support_steel_HCS,HumanTNCI_score_support_steel_HCS, HumanTNCM_score_support_steel_HCS, HumanTCO_score_support_steel_HCS, HumanTCI_score_support_steel_HCS, HumanTCM_score_support_steel_HCS, ExotoxicityFO_score_support_steel_HCS, ExotoxicityFI_score_support_steel_HCS, ExotoxicityFM_score_support_steel_HCS],
        'HCS - GLT supports': [f"Beams: {profile_beam_GLT_HCS} and columns: {profile_column_GLT_HCS}", mPt_score_support_GLT_HCS, GWP_score_support_GLT_HCS, OzoneD_score_support_GLT_HCS, IonisingR_score_support_GLT_HCS, photochemicalOF_score_support_GLT_HCS, ParticulatM_score_support_GLT_HCS, HumanTNC_score_support_GLT_HCS, HumanTC_score_support_GLT_HCS, Acidification_score_support_GLT_HCS, EutrophicationF_score_support_GLT_HCS, EutrophicationM_score_support_GLT_HCS, EutrophicationT_score_support_GLT_HCS, EcotoxicityF_score_support_GLT_HCS, LandUse_score_support_GLT_HCS, WaterUse_score_support_GLT_HCS, ResourceUseF_score_support_GLT_HCS, ResourceUseMM_score_support_GLT_HCS, ClimateCF_score_support_GLT_HCS, ClimateCB_score_support_GLT_HCS, ClimateCLU_score_support_GLT_HCS, HumanTNCO_score_support_GLT_HCS,HumanTNCI_score_support_GLT_HCS, HumanTNCM_score_support_GLT_HCS, HumanTCO_score_support_GLT_HCS, HumanTCI_score_support_GLT_HCS, HumanTCM_score_support_GLT_HCS, ExotoxicityFO_score_support_GLT_HCS, ExotoxicityFI_score_support_GLT_HCS, ExotoxicityFM_score_support_GLT_HCS],
        'HCS - Cast-in-situ concrete supports': [f"Beams: {profile_beam_concrete_HCS} and columns: {profile_column_concrete_HCS}", mPt_score_support_concrete_HCS, GWP_score_support_concrete_HCS, OzoneD_score_support_concrete_HCS, IonisingR_score_support_concrete_HCS, photochemicalOF_score_support_concrete_HCS, ParticulatM_score_support_concrete_HCS, HumanTNC_score_support_concrete_HCS, HumanTC_score_support_concrete_HCS, Acidification_score_support_concrete_HCS, EutrophicationF_score_support_concrete_HCS, EutrophicationM_score_support_concrete_HCS, EutrophicationT_score_support_concrete_HCS, EcotoxicityF_score_support_concrete_HCS, LandUse_score_support_concrete_HCS, WaterUse_score_support_concrete_HCS, ResourceUseF_score_support_concrete_HCS, ResourceUseMM_score_support_concrete_HCS, ClimateCF_score_support_concrete_HCS, ClimateCB_score_support_concrete_HCS, ClimateCLU_score_support_concrete_HCS, HumanTNCO_score_support_concrete_HCS,HumanTNCI_score_support_concrete_HCS, HumanTNCM_score_support_concrete_HCS, HumanTCO_score_support_concrete_HCS, HumanTCI_score_support_concrete_HCS, HumanTCM_score_support_concrete_HCS, ExotoxicityFO_score_support_concrete_HCS, ExotoxicityFI_score_support_concrete_HCS, ExotoxicityFM_score_support_concrete_HCS],
        'CLT - Steel supports': [f"Beams: {profile_beam_steel_CLT} and columns: {profile_column_steel_CLT}", mPt_score_support_steel_CLT, GWP_score_support_steel_CLT, OzoneD_score_support_steel_CLT, IonisingR_score_support_steel_CLT, photochemicalOF_score_support_steel_CLT, ParticulatM_score_support_steel_CLT, HumanTNC_score_support_steel_CLT, HumanTC_score_support_steel_CLT, Acidification_score_support_steel_CLT, EutrophicationF_score_support_steel_CLT, EutrophicationM_score_support_steel_CLT, EutrophicationT_score_support_steel_CLT, EcotoxicityF_score_support_steel_CLT, LandUse_score_support_steel_CLT, WaterUse_score_support_steel_CLT, ResourceUseF_score_support_steel_CLT, ResourceUseMM_score_support_steel_CLT, ClimateCF_score_support_steel_CLT, ClimateCB_score_support_steel_CLT, ClimateCLU_score_support_steel_CLT, HumanTNCO_score_support_steel_CLT,HumanTNCI_score_support_steel_CLT, HumanTNCM_score_support_steel_CLT, HumanTCO_score_support_steel_CLT, HumanTCI_score_support_steel_CLT, HumanTCM_score_support_steel_CLT, ExotoxicityFO_score_support_steel_CLT, ExotoxicityFI_score_support_steel_CLT, ExotoxicityFM_score_support_steel_CLT],
        'CLT - GLT supports': [f"Beams: {profile_beam_GLT_CLT} and columns: {profile_column_GLT_CLT}", mPt_score_support_GLT_CLT, GWP_score_support_GLT_CLT, OzoneD_score_support_GLT_CLT, IonisingR_score_support_GLT_CLT, photochemicalOF_score_support_GLT_CLT, ParticulatM_score_support_GLT_CLT, HumanTNC_score_support_GLT_CLT, HumanTC_score_support_GLT_CLT,Acidification_score_support_GLT_CLT, EutrophicationF_score_support_GLT_CLT, EutrophicationM_score_support_GLT_CLT, EutrophicationT_score_support_GLT_CLT, EcotoxicityF_score_support_GLT_CLT, LandUse_score_support_GLT_CLT, WaterUse_score_support_GLT_CLT, ResourceUseF_score_support_GLT_CLT, ResourceUseMM_score_support_GLT_CLT, ClimateCF_score_support_GLT_CLT, ClimateCB_score_support_GLT_CLT, ClimateCLU_score_support_GLT_CLT, HumanTNCO_score_support_GLT_CLT,HumanTNCI_score_support_GLT_CLT, HumanTNCM_score_support_GLT_CLT, HumanTCO_score_support_GLT_CLT, HumanTCI_score_support_GLT_CLT, HumanTCM_score_support_GLT_CLT, ExotoxicityFO_score_support_GLT_CLT, ExotoxicityFI_score_support_GLT_CLT, ExotoxicityFM_score_support_GLT_CLT],
        'CLT - Cast-in-situ concrete supports': [f"Beams: {profile_beam_concrete_CLT} and columns: {profile_column_concrete_CLT}", mPt_score_support_concrete_CLT, GWP_score_support_concrete_CLT, OzoneD_score_support_concrete_CLT, IonisingR_score_support_concrete_CLT, photochemicalOF_score_support_concrete_CLT, ParticulatM_score_support_concrete_CLT, HumanTNC_score_support_concrete_CLT, HumanTC_score_support_concrete_CLT, Acidification_score_support_concrete_CLT, EutrophicationF_score_support_concrete_CLT, EutrophicationM_score_support_concrete_CLT, EutrophicationT_score_support_concrete_CLT, EcotoxicityF_score_support_concrete_CLT, LandUse_score_support_concrete_CLT, WaterUse_score_support_concrete_CLT, ResourceUseF_score_support_concrete_CLT, ResourceUseMM_score_support_concrete_CLT, ClimateCF_score_support_concrete_CLT, ClimateCB_score_support_concrete_CLT, ClimateCLU_score_support_concrete_CLT, HumanTNCO_score_support_concrete_CLT,HumanTNCI_score_support_concrete_CLT, HumanTNCM_score_support_concrete_CLT, HumanTCO_score_support_concrete_CLT, HumanTCI_score_support_concrete_CLT, HumanTCM_score_support_concrete_CLT, ExotoxicityFO_score_support_concrete_CLT, ExotoxicityFI_score_support_concrete_CLT, ExotoxicityFM_score_support_concrete_CLT],
        'Concrete - Steel supports': [f"Beams: {profile_beam_steel_concrete} and columns: {profile_column_steel_concrete}", mPt_score_support_steel_concrete, GWP_score_support_steel_concrete, OzoneD_score_support_steel_concrete, IonisingR_score_support_steel_concrete, photochemicalOF_score_support_steel_concrete, ParticulatM_score_support_steel_concrete, HumanTNC_score_support_steel_concrete, HumanTC_score_support_steel_concrete, Acidification_score_support_steel_concrete, EutrophicationF_score_support_steel_concrete, EutrophicationM_score_support_steel_concrete, EutrophicationT_score_support_steel_concrete, EcotoxicityF_score_support_steel_concrete, LandUse_score_support_steel_concrete, WaterUse_score_support_steel_concrete, ResourceUseF_score_support_steel_concrete, ResourceUseMM_score_support_steel_concrete, ClimateCF_score_support_steel_concrete, ClimateCB_score_support_steel_concrete, ClimateCLU_score_support_steel_concrete, HumanTNCO_score_support_steel_concrete,HumanTNCI_score_support_steel_concrete, HumanTNCM_score_support_steel_concrete, HumanTCO_score_support_steel_concrete, HumanTCI_score_support_steel_concrete, HumanTCM_score_support_steel_concrete, ExotoxicityFO_score_support_steel_concrete, ExotoxicityFI_score_support_steel_concrete, ExotoxicityFM_score_support_steel_concrete],
        'Concrete - GLT supports': [f"Beams: {profile_beam_GLT_concrete} and columns: {profile_column_GLT_concrete}", mPt_score_support_GLT_concrete, GWP_score_support_GLT_concrete, OzoneD_score_support_GLT_concrete, IonisingR_score_support_GLT_concrete, photochemicalOF_score_support_GLT_concrete, ParticulatM_score_support_GLT_concrete, HumanTNC_score_support_GLT_concrete, HumanTC_score_support_GLT_concrete, Acidification_score_support_GLT_concrete, EutrophicationF_score_support_GLT_concrete, EutrophicationM_score_support_GLT_concrete, EutrophicationT_score_support_GLT_concrete, EcotoxicityF_score_support_GLT_concrete, LandUse_score_support_GLT_concrete, WaterUse_score_support_GLT_concrete, ResourceUseF_score_support_GLT_concrete, ResourceUseMM_score_support_GLT_concrete, ClimateCF_score_support_GLT_concrete, ClimateCB_score_support_GLT_concrete, ClimateCLU_score_support_GLT_concrete, HumanTNCO_score_support_GLT_concrete,HumanTNCI_score_support_GLT_concrete, HumanTNCM_score_support_GLT_concrete, HumanTCO_score_support_GLT_concrete, HumanTCI_score_support_GLT_concrete, HumanTCM_score_support_GLT_concrete, ExotoxicityFO_score_support_GLT_concrete, ExotoxicityFI_score_support_GLT_concrete, ExotoxicityFM_score_support_GLT_concrete],
        'Concrete - Cast-in-situ concrete supports': [f"Beams: {profile_beam_concrete_concrete} and columns: {profile_column_concrete_concrete}", mPt_score_support_concrete_concrete, GWP_score_support_concrete_concrete, OzoneD_score_support_concrete_concrete, IonisingR_score_support_concrete_concrete, photochemicalOF_score_support_concrete_concrete, ParticulatM_score_support_concrete_concrete, HumanTNC_score_support_concrete_concrete, HumanTC_score_support_concrete_concrete, Acidification_score_support_concrete_concrete, EutrophicationF_score_support_concrete_concrete, EutrophicationM_score_support_concrete_concrete, EutrophicationT_score_support_concrete_concrete, EcotoxicityF_score_support_concrete_concrete, LandUse_score_support_concrete_concrete, WaterUse_score_support_concrete_concrete, ResourceUseF_score_support_concrete_concrete, ResourceUseMM_score_support_concrete_concrete, ClimateCF_score_support_concrete_concrete, ClimateCB_score_support_concrete_concrete, ClimateCLU_score_support_concrete_concrete, HumanTNCO_score_support_concrete_concrete,HumanTNCI_score_support_concrete_concrete, HumanTNCM_score_support_concrete_concrete, HumanTCO_score_support_concrete_concrete, HumanTCI_score_support_concrete_concrete, HumanTCM_score_support_concrete_concrete, ExotoxicityFO_score_support_concrete_concrete, ExotoxicityFI_score_support_concrete_concrete, ExotoxicityFM_score_support_concrete_concrete]
    }
    
    df = pd.DataFrame(data)
    
    df.to_excel(output_excel, index=False)
    print(f"Data successfully exported to {output_excel}")

output_excel = r"C:\Users\runan\Documents\Masterproef\2024 April-mei\Environmental_results.xlsx"

create_and_export_data(output_excel)







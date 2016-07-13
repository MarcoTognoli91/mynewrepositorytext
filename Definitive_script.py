from openpyxl import *
from opaque_surface_function import *
from fenestration_function import *
#from BelowGrade_upd_Salvatore import findData
#from BelowGrade_upd_Salvatore import BelowGradeSimple
from BelowGrade import *
from vent_inf_group4 import *
from final import *
import os
#cd c:/Users/Marco/Dropbox/Building system - Group B work/EETBS-Group Assignment 1-RLF Implementation in Python
os.chdir("C:/Users/Marco/Dropbox/Building system - Group B work/EETBS-Group Assignment 1-RLF Implementation in Python")

#define the location characteristic dictionary
Name_location=["Stockholm","Munich","Naples","Heraklion","Cairo"]
T_set_heating=[20,20,20,20,20]
T_set_cooling=[24,24,24,24,24]
T_out_heating=[-5,-2,0,2,5]
T_out_cooling=[27,29,32,35,40]
w_in=[0.0093,0.0093,0.0093,0.0093,0.0093]#humidity level inside the building in summer condition
w_out=[0.103,0.110,0.123,0.130,0.150]#humidity level outside the building in summer condition. Note that humidity is here defined just
#for cooling calculations. Even tought the vent/inf. function deserve some humidity value even for winter computation.
#In that case just put the summer values and do not consider the latent heat computed!
h_in_summer=[47.8,47.8,47.8,47.8,47.8]
h_out_summer=[52,54,56,60,64]
#Also in this case we will put the summer enthalpy value even for winter calculation and avoid the latent heat due to vent/inf as a result

DR=[9,8,10,11,7]
Latitude=[60,50,45,40,30]
A_floor=150 #m^2 floor area
N_occ=5 #number of occupants
input_duct="Attic"
input_insulation=0.7 #insulation of the ducts
input_leakage=5 #%
input_conditioning_heating="H/F"
input_stories=1 #number of stories
Situations=list()
for index in range(0,len(T_set_heating)):
    name=Name_location[index]    
    t_set_heating=T_set_heating[index]
    t_set_cooling=T_set_cooling[index]
    t_out_heating=T_out_heating[index]
    t_out_cooling=T_out_cooling[index]
    w_in_cooling=w_in[index]
    w_out_cooling=w_out[index]
    h_in_cooling=h_in_summer[index]
    h_out_cooling=h_out_summer[index]
    dr=DR[index]
    latitude=Latitude[index]
    Situation = {"T_set_heating":t_set_heating,"T_set_cooling":t_set_cooling,"T_out_heating":t_out_heating,"T_out_cooling":t_out_cooling,
    "DR":dr,"w_in_cooling":w_in_cooling,"w_out_cooling":w_out_cooling,"h_in_cooling":h_in_cooling, "h_out_cooling":h_out_cooling, "Latitude":latitude, "Location Name":name}
    Situations.append(Situation)
ExcelFileResults = load_workbook("Overall_Results.xlsx")
#internal gains-equal for every case
Q_int_sensible=136+2.2*A_floor+22*N_occ
Q_int_latent=20+0.22*A_floor+12*N_occ

#for index in range(0,1):
for index in range(0,len(T_set_heating)):
    season=["Heating","Cooling"]
    opaque_surface_function(Situations[index]["T_set_heating"], Situations[index]["T_set_cooling"], Situations[index]["T_out_heating"], Situations[index]["T_out_cooling"], Situations[index]["DR"])

    ExcelFileOpaque = load_workbook("OpaqueSurfaceCharacteristic.xlsx")
    OpaqueResults = ExcelFileOpaque.get_sheet_by_name("HouseCharacteristics")
    OpaqueResultsValueHeating=OpaqueResults.columns[11][1:]
    OpaqueResultsValueCooling=OpaqueResults.columns[13][1:]
    Results = ExcelFileResults.get_sheet_by_name("Results-Situation "+str(index+1))
    Q_tot_surface_heating=0
    Q_tot_surface_cooling=0
    Q_tot_fenestration_heating=0
    Q_tot_fenestration_cooling=0

    for i in range(0,len(OpaqueResultsValueHeating)):
        Q_tot_surface_heating=Q_tot_surface_heating+OpaqueResultsValueHeating[i].value
        Q_tot_surface_cooling=Q_tot_surface_cooling+OpaqueResultsValueCooling[i].value
        Results.columns[2][i+2].value=OpaqueResultsValueHeating[i].value
        Results.columns[3][i+2].value=OpaqueResultsValueCooling[i].value
        ExcelFileResults.save("Overall_Results.xlsx")
    Windows=fenestration_function(Situations[index]["T_set_heating"], Situations[index]["T_set_cooling"], Situations[index]["T_out_heating"], Situations[index]["T_out_cooling"], Situations[index]["DR"], Situations[index]["Latitude"], "WindowList2.xlsx")
    for window in Windows:
        Q_tot_fenestration_heating=Q_tot_fenestration_heating+window["Heating flux"]
        Q_tot_fenestration_cooling=Q_tot_fenestration_cooling+window["Cooling flux"]
    Results.columns[2][6].value=Q_tot_surface_heating
    Results.columns[3][6].value=Q_tot_surface_cooling
    Results.columns[2][8].value=Q_tot_fenestration_heating
    Results.columns[3][8].value=Q_tot_fenestration_cooling
    q_below_grade=range(0,2)
    q_below_grade[0]={"Value":0}
    q_below_grade[1]={"Value":0}    #q_below_grade=BelowGradeSimple()#pay attention at the values imposed and at the one added inside the function itself
    #Results.columns[2][10].value=q_below_grade[0]["Value"]
    #Results.columns[2][11].value=q_below_grade[1]["Value"]
    #Results.columns[2][12].value=q_below_grade[0]["Value"]+q_below_grade[1]["Value"]#<---below grade calculation refers to an ambient under soil level, its
    #heat flux affects JUST the heating calculation. If present consider a variation in opaque surface calculation of CF_floor as adjacent to an ambient
    #and no more close to a crawlspace/slab floor.
    #moreover pay attention at the temperature imposed inside belowgrade calculation set inside the related excel worksheet
    RES_heating=vent_infilt_load(Situations[index]["T_set_heating"],Situations[index]["T_out_heating"],
    Situations[index]["w_in_cooling"],Situations[index]["w_out_cooling"],Situations[index]["h_in_cooling"],Situations[index]["h_out_cooling"],season[0])
    RES_cooling=vent_infilt_load(Situations[index]["T_set_cooling"],Situations[index]["T_out_cooling"],
    Situations[index]["w_in_cooling"],Situations[index]["w_out_cooling"],Situations[index]["h_in_cooling"],Situations[index]["h_out_cooling"],season[1])
    Results.columns[2][14].value=RES_heating["q_vi_sensible"]
    Results.columns[3][14].value=RES_cooling["q_vi_sensible"]
    Results.columns[3][15].value=RES_cooling["q_vi_latent"]
    Results.columns[3][17].value=Q_int_sensible
    Results.columns[3][18].value=Q_int_latent
    Subtotal_cooling=Q_tot_surface_cooling+Q_tot_fenestration_cooling+RES_cooling["q_vi_sensible"]+Q_int_sensible
    Subtotal_heating=Q_tot_surface_heating+Q_tot_fenestration_heating+RES_heating["q_vi_sensible"]+q_below_grade[0]["Value"]+q_below_grade[1]["Value"]
    Results.columns[2][20].value=Subtotal_heating
    Results.columns[3][20].value=Subtotal_cooling
    Q_dist_heating=distributionLosses(input_duct, input_insulation, input_leakage, input_conditioning_heating, input_stories, Subtotal_heating)
    Q_dist_cooling=distributionLosses(input_duct, input_insulation, input_leakage, "C", input_stories, Subtotal_cooling)
    Results.columns[2][22].value=Q_dist_heating
    Results.columns[3][22].value=Q_dist_cooling
    Results.columns[2][24].value=Q_dist_heating+Subtotal_heating
    Results.columns[3][24].value=Q_dist_cooling+Subtotal_cooling
    Results.columns[3][25].value=Q_int_latent+RES_cooling["q_vi_latent"]
    #after the results it is better to report the location characteristics inside the result worksheet
    Results.columns[7][1].value=Situations[index]["Location Name"]
    Results.columns[7][2].value=Situations[index]["Latitude"]
    Results.columns[7][3].value=Situations[index]["T_out_heating"]
    Results.columns[7][4].value=Situations[index]["T_out_cooling"]
    Results.columns[7][5].value=Situations[index]["w_out_cooling"]
    Results.columns[7][6].value=Situations[index]["DR"]
    Results.columns[7][7].value=Situations[index]["h_out_cooling"]
    Results.columns[7][9].value=Situations[index]["T_set_heating"]
    Results.columns[7][10].value=Situations[index]["T_set_cooling"]
    Results.columns[7][11].value=Situations[index]["w_in_cooling"]
    Results.columns[7][12].value=Situations[index]["h_in_cooling"]

    ExcelFileResults.save("Overall_Results.xlsx")

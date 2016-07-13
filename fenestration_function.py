# -*- coding: utf-8 -*-
#Note for the user:
#1)Insert the Excel file using the same module as WindowList2.xlsx and be sure that the Python directory is the same as the Excel one
#2)Close the Excel file before any Python computation
#3)Check if you have installed openpyxl component and scipy one

from openpyxl import *


def fenestration_function(T_set_heating, T_set_cooling, T_out_heating, T_out_cooling, DR, Latitude, Excelname):
    """This function calculates the Heating Factor and Cooling Factor of fenestrations (trasparent walls). 
    The input argument are  first of all to fill out the Excel table ("WindowList2.xlsx") and to add the 
    temperature variable and latitude in this order:
    T_set_heating, T_set_cooling, T_out_heating, T_out_cooling, DR [Daily cooling Range], Latitude [as XX.XX],
    the Excel file position and name has to be provided into the function call as a string.
    The function exploit the fenestrations parameters, these are added to the Excel file and a dictionary
    of each single window property is exported.
    """
    import os
    
    #os.chdir("C:/Users/Marco/Dropbox\Building system - Group B work/May 26")#<----pay attention here as the function must rely on the same excel folder
    ExcelFile = load_workbook(Excelname)#Once our workbook it is defined it will be called several times
    WindowData = ExcelFile.get_sheet_by_name("WindowData")#with this command we extract the Excel sheet of interest
    name_cells = WindowData.columns[0][1:]#The windows data are arranged on the same sheet split in columns
    
    
    print name_cells[0].value #we use .value to extract the value from the selected tuple (cell)
    
    for cell in name_cells:
        name = cell.value
        print name
    windows = []#This is the main dictionary of the script, it collects ALL the windows data for each window
    
    Number_of_windows = len(name_cells)
    windows=list()
    for index in range(1,Number_of_windows+1):#Here we are defining the basic windows dictionary structure
        name = (WindowData.columns[0][index].value).encode('utf-8')
        direction = (WindowData.columns[1][index].value).encode('utf-8')  
        Height = float(WindowData.columns[2][index].value) #pay attention in these values since they are
        Width = float(WindowData.columns[3][index].value) #directly converted into float
        Frame_type = (WindowData.columns[4][index].value).encode('utf-8')
        Frame_Material = (WindowData.columns[5][index].value).encode('utf-8')
        ID = (WindowData.columns[6][index].value).encode('utf-8')
        Attachment = (WindowData.columns[7][index].value).encode('utf-8')
        Overhang_length = float(WindowData.columns[8][index].value)
        Overhang_height = float((WindowData.columns[9][index].value))
        Int_shade_type = (WindowData.columns[10][index].value).encode('utf-8')
        Int_shade_material = (WindowData.columns[11][index].value).encode('utf-8')
        Int_shade_close = (WindowData.columns[12][index].value).encode('utf-8')
        window = {"name":name,"direction":direction,"Height":Height,"Width":Width,"Frame_type":Frame_type, "Frame_Material":Frame_Material, "ID":ID
        ,"Attachment":Attachment ,"Overhang_length":Overhang_length ,"Overhang_height": Overhang_height,"Int_shade_type":Int_shade_type,"Int_shade_material":Int_shade_material, "Int_shade_close": Int_shade_close}
        windows.append(window)
        
        
        
       
    Results= ExcelFile.get_sheet_by_name("Results")#Results is the output Excel sheet
    
    U_fixed  = ExcelFile.get_sheet_by_name("U_Fixed_Fenestration")
    
    U_Fixed_Fenestration_IDs = U_fixed.columns[2][2:] #These for cycles convert the list of cells in a list of values, in this case a string type
    U_Fixed_IDs =list() 
    for cell in U_Fixed_Fenestration_IDs:
        cell_ID= (cell.value).encode('utf-8')
        U_Fixed_IDs.append(cell_ID)
        
    U_Fixed_Fenestration_Frame_IDs =U_fixed.rows[1][3:] #Must be repeated also for rows 
    U_Fixed_frame_IDs = list()
    for cell in U_Fixed_Fenestration_Frame_IDs:
        cell_frame_ID = (cell.value).encode('utf-8')
        U_Fixed_frame_IDs.append(cell_frame_ID)
        
    U_Operable  = ExcelFile.get_sheet_by_name("U_Operable_Fenestration")#Windows are splitted into operables side and fixed side
    
    U_Operable_Fenestration_IDs = U_Operable.columns[2][2:] 
    U_Operable_IDs =list() 
    for cell in U_Operable_Fenestration_IDs:
        cell_ID= (cell.value).encode('utf-8')
        U_Operable_IDs.append(cell_ID)
        
    U_Operable_Fenestration_Frame_IDs =U_Operable.rows[1][3:]
    U_Operable_frame_IDs = list()
    for cell in U_Operable_Fenestration_Frame_IDs:
        cell_frame_ID = (cell.value).encode('utf-8')
        U_Operable_frame_IDs.append(cell_frame_ID)    
        
    deltaT_Heating = T_set_heating-T_out_heating
    deltaT_cooling = T_out_cooling-T_set_cooling
    
    for window in windows:#This main for cycle detect, for fixed and operable windows, the proper U value based on the correct
        window_ID = window["ID"]#indexes. In this case the condition it is based on ID code and frame type
        window_Frame_ID = window["Frame_Material"]
        if (window["Frame_type"]=="Fixed") or (window["Frame_type"]=="fixed"):
            window_index_ID = U_Fixed_IDs.index(window_ID)+2 #just remind that indexes are shifted, in this case by two positions
            window_index_Frame_ID =  U_Fixed_frame_IDs.index(window_Frame_ID)+3
            U_window = float((U_fixed.columns[window_index_Frame_ID][window_index_ID].value).encode('utf-8'))
        elif (window["Frame_type"]=="Operable") or (window["Frame_type"]=="operable"):
            window_index_ID = U_Operable_IDs.index(window_ID)+2
            window_index_Frame_ID =  U_Operable_frame_IDs.index(window_Frame_ID)+3
            U_window = float((U_Operable.columns[window_index_Frame_ID][window_index_ID].value).encode('utf-8'))
        else:
            print "The frame type was not acceptable a fixed one was used instead " #in case our frame it si not acceptable
            window_index_ID = U_Fixed_IDs.index(window_ID)+2 #just use the fixed one
            window_index_Frame_ID =  U_Fixed_frame_IDs.index(window_Frame_ID)+3
            U_window = float((U_fixed.columns[window_index_Frame_ID][window_index_ID].value).encode('utf-8'))
            
        HF_window = U_window*deltaT_Heating#The methood prescribe for heating operation the calculation of the so called
                                        #Heating Factor that it is simply U*deltaT
        window["U"] = U_window #Finally U values and HF are assigned into windows dictionary
        window["HF"] = HF_window
        
    
    
    
    for window in windows: #To compute the Excel results we select the final columns of interest into the Results sheet
        window_index= windows.index(window)
        U = window["U"]
        HF = window["HF"]
        Results.columns[2][window_index+1].value = U#Again the correct index it is shifted due to table position
        Results.columns[3][window_index+1].value = HF#inside the result sheet
        ExcelFile.save(Excelname)
    
    #SHGC
    SHGC_fixed  = ExcelFile.get_sheet_by_name("SHGC_Fixed_Fenestration")#Here the steps are right the same as the U calculation
    
    SHGC_Fixed_Fenestration_IDs = SHGC_fixed.columns[2][2:] 
    SHGC_Fixed_IDs =list() 
    for cell in SHGC_Fixed_Fenestration_IDs:
        cell_ID= (cell.value).encode('utf-8')
        SHGC_Fixed_IDs.append(cell_ID)
        
    SHGC_Fixed_Fenestration_Frame_IDs =SHGC_fixed.rows[1][3:]
    SHGC_Fixed_frame_IDs = list()
    for cell in SHGC_Fixed_Fenestration_Frame_IDs:
        cell_frame_ID = (cell.value).encode('utf-8')
        SHGC_Fixed_frame_IDs.append(cell_frame_ID)
        
    SHGC_operable  = ExcelFile.get_sheet_by_name("SHGC_Operable_Fenestration")
    
    SHGC_Operable_Fenestration_IDs = SHGC_operable.columns[2][2:] 
    SHGC_Operable_IDs =list() 
    for cell in SHGC_Operable_Fenestration_IDs:
        cell_ID= (cell.value).encode('utf-8')
        SHGC_Operable_IDs.append(cell_ID)
        
    SHGC_Operable_Fenestration_Frame_IDs =SHGC_operable.rows[1][3:]
    SHGC_Operable_frame_IDs = list()
    for cell in SHGC_Operable_Fenestration_Frame_IDs:
        cell_frame_ID = cell.value
        SHGC_Operable_frame_IDs.append(cell_frame_ID)    
    
    for window in windows:
        window_ID = window["ID"]
        window_Frame_ID = window["Frame_Material"]
        if (window["Frame_type"]=="Fixed") or (window["Frame_type"]=="fixed"):
            window_index_ID = SHGC_Fixed_IDs.index(window_ID)+2 #just remind that indexes are shifted, in this case by two positions
            window_index_Frame_ID =  SHGC_Fixed_frame_IDs.index(window_Frame_ID)+3
            SHGC_window = float((SHGC_fixed.columns[window_index_Frame_ID][window_index_ID].value).encode('utf-8'))
        elif (window["Frame_type"]=="Operable") or (window["Frame_type"]=="operable"):
            window_index_ID = SHGC_Operable_IDs.index(window_ID)+2
            window_index_Frame_ID =  SHGC_Operable_frame_IDs.index(window_Frame_ID)+3
            SHGC_window = float((SHGC_operable.columns[window_index_Frame_ID][window_index_ID].value).encode('utf-8'))
        else:
            print "The frame type was not acceptable a fixed one was used instead " #in case our frame it si not acceptable
            window_index_ID = SHGC_Fixed_IDs.index(window_ID)+2 #just use the fixed one
            window_index_Frame_ID =  SHGC_Fixed_frame_IDs.index(window_Frame_ID)+3
            SHGC_window = float((SHGC_fixed.columns[window_index_Frame_ID][window_index_ID].value).encode('utf-8'))
            
            
        window["SHGC"] = SHGC_window   
    
    
    Results= ExcelFile.get_sheet_by_name("Results")
    for window in windows: 
        window_index= windows.index(window)
        SHGC = window["SHGC"]
        Results.columns[7][window_index+1].value = SHGC
        ExcelFile.save(Excelname)    
    
    #T_x
    Tx_none=1#Definition of T_x values depending on attachment type
    Tx_insect=0.64#If an insect screen is present
    
    
    for window in windows:#At each window it is assigned the corrispondnat T_x value
        window_Attachment = window["Attachment"]
        if (window["Attachment"]=="None"):
            Attachment_window = Tx_none
        elif (window["Attachment"]=="Insect_screen"):
            Attachment_window = Tx_insect
        else:
            print "The Attachment type was not acceptable no insect screen was used instead "
            Attachment_window = Tx_none
            
            
        window["Tx"] = Attachment_window   
    
    
    Results= ExcelFile.get_sheet_by_name("Results")
    for window in windows: 
        window_index= windows.index(window)
        Tx = window["Tx"]
        Results.columns[4][window_index+1].value = Tx
        ExcelFile.save(Excelname)
    
    #PXI
    Peak_irr  = ExcelFile.get_sheet_by_name("Peak_Irradiance")#To get the real peak irradiance value the vertical unshaded values are needed
    SLFs  = ExcelFile.get_sheet_by_name("Shade_Line_Factor")#Shade line factors are as well derived by table
    
    #Calculation of Shade Line Factors
    
    SLFs_exposure = SLFs.columns[0][4:] 
    SLFs_IDs =list() 
    for cell in SLFs_exposure:
        cell_ID= (cell.value).encode('utf-8')
        SLFs_IDs.append(cell_ID)#we extract the list of exposure
        
    SLFs_coordinate =SLFs.rows[3][1:]
    SLFs_coordinate_IDs = list()
    for cell in SLFs_coordinate:
        cell_frame_ID = cell.value
        SLFs_coordinate_IDs.append(cell_frame_ID)#extraction of latitude list

    #SLFs_coordinate_interp_min=list()
    #for index in SLFs_coordinate_IDs:
    #        SLFs_coordinate_interp_min[index]=SLFs_coordinate_IDs-Latitude
    #SLFs_coordinate_interp_minval=min(SLFs_coordinate_IDs_min)
    from scipy.interpolate import interp1d#recalling the numeric addin to introduce the interpolation operation
    
    
    for window in windows:
        window_direction = window["direction"]
        window_index_direction = SLFs_IDs.index(window_direction)+4
        window_direction_SLFs_tuple=SLFs.rows[window_index_direction][1:]#to get the right SLFs value at that specific latitude we need
        window_direction_SLFs=list()#to interpolate, so we extract the SLFs values from that specific row that means per each window exposition        
        for index in window_direction_SLFs_tuple:
            index_value = index.value
            window_direction_SLFs.append(index_value)#window_direction_SLFs_tuple is a list of cell, by this command the list of float is extracted
        SLFs_function=interp1d(SLFs_coordinate_IDs, window_direction_SLFs,kind='cubic')#this cubic function is created to further determine a value in between the list
        #window_index_direction_coordinate =  SLFs_coordinate_IDs.index(Latitude)+1
        SLFs_window = float(SLFs_function(Latitude))
            
        window["SLFs"] = SLFs_window
    #print(windows["SLFs"])
    
    #Calculation of Shading Fraction
    
    for window in windows:#Shading graction is the effect caused by total shading caused by a overhang element
        window_overhang_length = window["Overhang_length"]
        window_overhang_height = window["Overhang_height"]
        SLFs_window=window["SLFs"]
        height=window["Height"]
        F_shd=min(1,max(0,((SLFs_window*window_overhang_length-window_overhang_height)/height)))#the formula accounts also for
        #non shaded windows, in that case F_shd=0
        
        window["F_shd"] = F_shd
    
    Results= ExcelFile.get_sheet_by_name("Results")    
    for window in windows: 
        window_index= windows.index(window)
        F_shd = window["F_shd"]
        Results.columns[5][window_index+1].value = F_shd
        #ExcelFile.save(Excelname)
    
    #PXI calculation
    
    T_x=Results.columns[4][1:]
    PXI=ExcelFile.get_sheet_by_name("Peak_Irradiance")
    
    PXI_orientation=list()
    for cell in range(4,29,3):#Peak irradiance table provide values for direct, diffuse and total irradiance so we select the right indexes by a step of three
        PXI_orientation_cell=PXI.columns[0][cell].value.encode('utf-8')
        PXI_orientation.append(PXI_orientation_cell)
    
    PXI_latitude=PXI.rows[3][2:]
    PXI_latitude_IDs = list()
    for cell in PXI_latitude:
        cell_ID = cell.value
        PXI_latitude_IDs.append(cell_ID)
    T_x=Results.columns[4][1:]
    
    for window in windows:
        window_orientation = window["direction"]
        window_Overhang_length = window["Overhang_length"]
        window_F_shd=window["F_shd"]
        window_index= windows.index(window)
        window_T_x=T_x[window_index].value
        if window_Overhang_length>0:#for windows with external overhang
            PXI_index = PXI_orientation.index(window_orientation)#once again we need to interpolate the values on latitude scale
            #window_latitude =  PXI_latitude_IDs.index(Latitude)
            window_direction_PXI_tuple=PXI.rows[PXI_index*3+5][2:]#the problem here is the difference in direct, diffuse and total irradiance
            window_direction_PXI_direct_tuple=PXI.rows[PXI_index*3+4][2:]#so we have to compute the list of values for each of them at each window exposition
            window_direction_PXI=list()        
            for index in window_direction_PXI_tuple:
                index_value = index.value
                window_direction_PXI.append(index_value)
            PXI_function_diffuse=interp1d(SLFs_coordinate_IDs, window_direction_PXI,kind='cubic')
            window_direction_PXI_direct=list()        
            for index in window_direction_PXI_direct_tuple:
                index_value = index.value
                window_direction_PXI_direct.append(index_value)
            PXI_function_direct=interp1d(SLFs_coordinate_IDs, window_direction_PXI_direct,kind='cubic')
            PXI_results=float(window_T_x*float(PXI_function_diffuse(Latitude))+(1-window_F_shd)*float((PXI_function_direct(Latitude))))
        elif window_Overhang_length==0:#for windows with no overhang
            window_direction_PXI_total_tuple=PXI.rows[PXI_index*3+6][2:]
            window_direction_PXI_total=list()        
            for index in window_direction_PXI_total_tuple:
                index_value = index.value
                window_direction_PXI_total.append(index_value)
            PXI_function_total=interp1d(SLFs_coordinate_IDs, window_direction_PXI_total,kind='cubic')#here we need to create a function for total irradiance
            PXI_results=float(window_T_x*(float(PXI_function_total(Latitude))))
        window["PXI"]=PXI_results
    
    for window in windows: 
        window_index= windows.index(window)
        PXI_result = window["PXI"]
        Results.columns[6][window_index+1].value = PXI_result
        #ExcelFile.save(Excelname)
    
    #SLF or also called FFs, they account for solar irradiance that affects the peak cooling load
    SLF  = ExcelFile.get_sheet_by_name("Solar_Load_Factor")
    
    SLF_exposure = SLF.columns[0][2:] 
    SLF_IDs =list() 
    for cell in SLF_exposure:
        cell_ID= (cell.value).encode('utf-8')
        SLF_IDs.append(cell_ID)
        
    SLF_type =SLF.rows[1][1:3]
    SLF_type_IDs = list()
    for cell in SLF_type:
        cell_frame_ID = cell.value
        SLF_type_IDs.append(cell_frame_ID)
    
    SLF_building="Single Family Detached"
    
    for window in windows:
        window_direction = window["direction"]
        if (SLF_building=="Single Family Detached") or (SLF_building=="Multifamily"):#for multifamily buildings or single family ones
            window_index_direction = SLF_IDs.index(window_direction)+2
            window_index_direction_type =  SLF_type_IDs.index(SLF_building)+1
            SLF_window = float((SLF.columns[window_index_direction_type][window_index_direction].value).encode('utf-8'))
        else:
            print "The building type was not acceptable a multifamily one was used instead " 
            window_index_direction = SLF_IDs.index(window_direction)+2
            window_index_direction_type =  SLF_type_IDs.index("Multifamily")+1
            SLF_window = float((SLF.columns[window_index_direction_type][window_index_direction].value).encode('utf-8'))
            
            
                
        window["SLF"] = SLF_window
    
    Results= ExcelFile.get_sheet_by_name("Results")    
    for window in windows: 
        window_index= windows.index(window)
        SLF = window["SLF"]
        Results.columns[9][window_index+1].value = SLF #in the table SLF are indicated also as FFS
        #ExcelFile.save(Excelname)
    
    #IAC
    IACs=ExcelFile.get_sheet_by_name("Interior_Attenuation_Coefficien")
    #this computation accounts for interior shading system like drapes
    IACs_IDs_cell=IACs.columns[2][6:]
    IACs_ID=list()
    for cell in IACs_IDs_cell:
        IAC_ID_cell=cell.value
        IACs_ID.append(IAC_ID_cell)
    
    IACs_types_cell=IACs.rows[4][3:]
    IACs_type=list()
    for cell in IACs_types_cell:
        IAC_type_cell=cell.value
        IACs_type.append(IAC_type_cell)
    
    Results= ExcelFile.get_sheet_by_name("Results")
    F_shd_dict=Results.columns[5][1:]
    F_shd=list()
    for cell in F_shd_dict:
        F_shd_cell=cell.value
        F_shd.append(F_shd_cell)
    
    for window in windows:
        window_ID = window["ID"]
        window_index= windows.index(window)
        window_F_shd = F_shd[window_index]
        if window["Int_shade_material"]=="Open_weave":#the tables are set for drapes and roller shaders, each of them with specific configuration
            IACcl_ID_index = IACs_ID.index(window_ID)+4 #just remind that indexes are shifted
            IACcl_type_index =  IACs_type.index(window_F_shd)+3
            IAC=float(1+window_F_shd*(float(IACs.columns[IACcl_type_index][IACcl_ID_index].value.encode('utf-8')-1)))
        elif window["Int_shade_material"]=="Closed_weave_dark":
            IACcl_ID_index = IACs_ID.index(window_ID)+4 #just remind that indexes are shifted
            IACcl_type_index = 4
            IAC=float(1+window_F_shd*(float(IACs.columns[IACcl_type_index][IACcl_ID_index].value.encode('utf-8')-1)))
        elif window["Int_shade_material"]=="Closed_weave_light":
            IACcl_ID_index = IACs_ID.index(window_ID)+4 #just remind that indexes are shifted
            IACcl_type_index =5
            IAC=float(1+window_F_shd*(float(IACs.columns[IACcl_type_index][IACcl_ID_index].value.encode('utf-8')-1)))
        elif window["Int_shade_material"]=="Closed_weave_medium":#the so called medium terms means that the corresponding IAC factor is a mean average between dark and light weave
            IACcl_ID_index = IACs_ID.index(window_ID)+4 #just remind that indexes are shifted
            IACcl_type_index_light =4
            IACcl_type_index_dark =5
            IAC=float(1+window_F_shd*(((float(IACs.columns[IACcl_type_index_dark][IACcl_ID_index].value.encode('utf-8'))+float(IACs.columns[IACcl_type_index_light][IACcl_ID_index].value.encode('utf-8')))/2)-1))
        elif window["Int_shade_material"]=="Opaque_dark":
            IACcl_ID_index = IACs_ID.index(window_ID)+4 #just remind that indexes are shifted
            IACcl_type_index =6
            IAC=float(1+window_F_shd*(float(IACs.columns[IACcl_type_index][IACcl_ID_index].value.encode('utf-8')-1)))
        elif window["Int_shade_material"]=="Opaque_light":
            IACcl_ID_index = IACs_ID.index(window_ID)+4 #just remind that indexes are shifted
            IACcl_type_index =7
            IAC=float(1+window_F_shd*(float(IACs.columns[IACcl_type_index][IACcl_ID_index].value.encode('utf-8')-1)))
        elif window["Int_shade_material"]=="Opaque_medium":
            IACcl_ID_index = IACs_ID.index(window_ID)+4 #just remind that indexes are shifted
            IACcl_type_index_light =6
            IACcl_type_index_dark =7
            IAC=float(1+window_F_shd*(((float(IACs.columns[IACcl_type_index_dark][IACcl_ID_index].value.encode('utf-8'))+float(IACs.columns[IACcl_type_index_light][IACcl_ID_index].value.encode('utf-8')))/2)-1))
        elif window["Int_shade_material"]=="Traslucent":
            IACcl_ID_index = IACs_ID.index(window_ID)+4 #just remind that indexes are shifted
            IACcl_type_index =  IACs_type.index(window_F_shd)+3
            IAC=float(1+window_F_shd*(float(IACs.columns[IACcl_type_index][IACcl_ID_index].value.encode('utf-8')-1)))
        elif window["Int_shade_material"]=="Blinds_medium":
            IACcl_ID_index = IACs_ID.index(window_ID)+4 #just remind that indexes are shifted
            IACcl_type_index =  IACs_type.index(window_F_shd)+3
            IAC=float(1+window_F_shd*(float(IACs.columns[IACcl_type_index][IACcl_ID_index].value.encode('utf-8')-1)))
        elif window["Int_shade_material"]=="Blinds_white":
            IACcl_ID_index = IACs_ID.index(window_ID)+4 #just remind that indexes are shifted
            IACcl_type_index =  IACs_type.index(window_F_shd)+3
            IAC=float(1+window_F_shd*(float(IACs.columns[IACcl_type_index][IACcl_ID_index].value.encode('utf-8')-1)))
        else:
            print "The internal shading layer it is not acceptable /n Please write it accoarding to the available nomenclature" #in case our frame it si not acceptable
            IAC="NaN"
        window["IAC"]=IAC
    
    for window in windows: 
        window_index= windows.index(window)
        IAC_result = window["IAC"]
        Results.columns[8][window_index+1].value = IAC_result
        #ExcelFile.save(Excelname)
    
    #CF final calculation of cooling condition
    U_dict=Results.columns[2][1:]
    T_x_dict=Results.columns[4][1:]
    F_shd_dict=Results.columns[5][1:]
    PXI_dict=Results.columns[6][1:]
    SHGC_dict=Results.columns[7][1:]
    IAC_dict=Results.columns[8][1:]
    FFS_dict=Results.columns[9][1:]
    #despite HF CF requests all these variables
    U=list()
    T_x=list()
    F_shd=list()
    PXI=list()
    SHGC=list()
    IAC=list()
    FFS=list()
    
    for cell in U_dict:
        cell=cell.value
        U.append(cell)
    for cell in T_x_dict:
        cell=cell.value
        T_x.append(cell)
    for cell in F_shd_dict:
        F_shd_cell=cell.value
        F_shd.append(F_shd_cell)
    for cell in PXI_dict:
        cell=cell.value
        PXI.append(cell)
    for cell in SHGC_dict:
        cell=cell.value
        SHGC.append(cell)
    for cell in IAC_dict:
        cell=cell.value
        IAC.append(cell)
    for cell in FFS_dict:
        cell=cell.value
        FFS.append(cell)
    #so here components are arranged in list accoarding to each window
    CF=list()
    for cell in range(0,len(U)):
        CF_cell=U[cell]*((-T_set_cooling+T_out_cooling)-0.46*DR)+PXI[cell]*SHGC[cell]*IAC[cell]*FFS[cell]
        CF.append(CF_cell)
    
    for cell in range(0,len(U)): 
        Results.columns[10][cell+1].value = CF[cell]
        ExcelFile.save(Excelname)
        windows[cell]["CF"]=CF[cell]
    
        
    #Heat/Cooling fluxes calculation, as it is simply HF/CF*delta_T
    for window in windows:
        window["Heating flux"]=window["Height"]*window["Width"]*window["HF"]
        window["Cooling flux"]=window["Height"]*window["Width"]*window["CF"]
    #Cooling and heating flux are present in the dictionaries while they are not exported to the result Excel file
    print windows
    return windows
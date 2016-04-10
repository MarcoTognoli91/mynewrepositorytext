glassProp = {"name":"glass", "k":0.86}
brickProp ={"name":"brick", "k": 0.89}
cement ={"name":"cement", "k": 1.4}
material_list = [ glassProp,brickProp,cement]
Ti =20
To= -10
Ri={"name":"Ri","type":"conv","area":0.25,"hConv":10}
R1={"name":"R1","type":"cond","length":0.03,"area":0.25,"k":0.026}
R2={"name":"R2","type":"cond","length":0.02,"area":0.25,"k":0.22}
R3={"name":"R3","type":"cond","length":0.16,"area":0.015,"k":0.22}
R4={"name":"R4","type":"cond","length":0.16,"area":0.22,"k":0.72}
R5={"name":"R5","type":"cond","length":0.16,"area":0.015,"k":0.22}
R6={"name":"R6","type":"cond","length":0.02,"area":0.25,"k":0.22}
Ro={"name":"Ro","type":"conv","area":0.25,"hConv":25}
parallelSet = [R3,R4,R5]
seriesSet = [R1,R2,R6]

def materialSensitivity(material_List,parallelSet,seriesSet,Ti,To,Ri,Ro,R3,R4,R5):
    #for material in material_List
    heatTransfer={}
    for innermaterial in material_list:
        R4["k"]=innermaterial["k"]
        parallelSet=[R3,R4,R5]
        from wallCalculation import wallHeatTransfer
        # you need to import wallHeatTransfer from wallCalculation Script
        nameLayer=innermaterial["name"]
        heatTransfer[nameLayer] = wallHeatTransfer(seriesSet,parallelSet,Ri,Ro, Ti,To)
        heatTransferResults=[heatTransfer]
    return heatTransferResults
heatTransfer=materialSensitivity(material_list,parallelSet,seriesSet,Ti,To,Ri,Ro,R3,R4,R5)
print(heatTransfer)

# sensitivityMultipleResultss={"material=glass, size=0.3":2.786,etc.....}
sizeList=[0.2,0.35,0.75]

def multipleMaterialSensitivity(material_list,sizeList,parallelSet,seriesSet,Ti,To,Ri,Ro,R3,R4,R5):
    #for material in material_List
    heatTransfer={}
    for innermaterial in material_list:
        for size in sizeList:
            R4["k"]=innermaterial["k"]
            R4["length"]=size
            parallelSet=[R3,R4,R5]
            from wallCalculation import wallHeatTransfer
            # you need to import wallHeatTransfer from wallCalculation Script
            nameLayer=innermaterial["name"]
            lengthLayer=size
            data=[nameLayer,lengthLayer]
            dataStr='material='+str(data[0])+' size='+str(data[1])
            heatTransfer[dataStr] = wallHeatTransfer(seriesSet,parallelSet,Ri,Ro, Ti,To)
            multipleHeatTransferResults=[heatTransfer]
    return multipleHeatTransferResults
multipleHeatTransfer=multipleMaterialSensitivity(material_list,sizeList,parallelSet,seriesSet,Ti,To,Ri,Ro,R3,R4,R5)
print(multipleHeatTransfer)
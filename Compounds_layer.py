A=raw_input("Please Enter the area of the wall (in m2): ")
material=raw_input("Please Enter the material of the layer: ")

if material=="glass":
    type_glas=raw_input("which type of glass do you mean:window=1, wool insulation=2 \n")
    if int(type_glas)==1:
        L=0.04
        k=str(0.96) #W/mK
    else:
        L=raw_input("Please Enter the length of the layer (in m): ")
        k= str(0.04) #W/mK
elif material=="brick":
    L=raw_input("Please Enter the length of the layer (in m): ")
    k=(raw_input("Please Enter the conductivity of the layer(in W/(m K): "))
else:
    print("I do not have the properties of this material")
    L=raw_input("Please Enter the length of the layer (in m): ")
    k=(raw_input("Please Enter the conductivity of the layer(in W/(m K): "))
print("\n you just said "+ "L= " + L+ " m "+ "A= " + A+ " m2 "+ "k= "+ k +" W/(m*K) \n")
R=float(L)/(float(k)*float(A))
print("Well the Thermal Resistnace is "+ str(R)+ " degC/W")
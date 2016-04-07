def Rcalculator():    
    A_string=raw_input("Please Enter the area of the wall (in m2):\n ")
    R_tot=0
    U=0
    n_layer_string=raw_input("Please Enter the total number of layers \n ")
    n_parallel_string=raw_input("Please Enter the total number of parallel layers \n ")
    n_layer=int(n_layer_string)
    A=float(A_string)
    n_parallel=int(n_parallel_string)
    if n_parallel==0:
        answer=raw_input("Does it contain any heat convective term? Y/N \n")
        if answer=="Y":
            for index in range(n_layer):
                
                answer2=raw_input("Is this one the convective term? Y/N \n")
                if answer2=="Y":
                    h=(raw_input("Please Enter the convective term(in W/(m2 K): "))
                    R=1/(float(h)*A)
                elif answer2=="N":
                    L=raw_input("Please Enter the length of the layer (in m): ")
                    k=(raw_input("Please Enter the conductivity of the layer(in W/(m K): "))
                    R=float(L)/(float(k)*A)
                R_tot=R_tot+R
        elif answer=='N':
            for index in range(n_layer):
                print("Layer number "+str(index+1)+"\n")
                L=raw_input("Please Enter the length of the layer (in m): ")
                k=(raw_input("Please Enter the conductivity of the layer(in W/(m K): "))
                R=float(L)/(float(k)*A)
                R_tot=R_tot+R
        print("This is the global wall thermal conductivity resistance "+str(R_tot))
    else:
        for index in range(n_layer):
            print("Layer number "+str(index+1)+"\n")
            answer3=raw_input("Is this one of the parallel layer? Y/N\n")
            if answer3=="Y":
                n_parallel_elements_string=raw_input("How many parallel elements are resent?\n")
                n_parallel_elements=int(n_parallel_elements_string)
                for index2 in range(n_parallel_elements):
                    print("Parallel element number "+str(index2+1)+"\n")
                    answer2=raw_input("Is this one the convective term? Y/N \n")
                    if answer2=="Y":
                        h=(raw_input("Please Enter the convective term(in W/(m2 K): "))
                        R=1/(float(h)*A)
                        U=(1/R)+U
                    elif answer2=="N":
                        L=raw_input("Please Enter the length of the layer (in m): ")
                        k=(raw_input("Please Enter the conductivity of the layer(in W/(m K): "))
                        R=float(L)/(float(k)*A)
                        U=(1/R)+U
                R_tot=R_tot+(1/U)
            else:
                answer2=raw_input("Is this one the convective term? Y/N \n")
                if answer2=="Y":
                    h=(raw_input("Please Enter the convective term(in W/(m2 K): "))
                    R=1/(float(h)*A)
                elif answer2=="N":
                    L=raw_input("Please Enter the length of the layer (in m): ")
                    k=(raw_input("Please Enter the conductivity of the layer(in W/(m K): "))
                    R=float(L)/(float(k)*A)
                R_tot=R_tot+R
        print("This is the global wall thermal conductivity resistance "+str(R_tot))
    return(R_tot)
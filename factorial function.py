def Factorial(n):
    number=float(n)
    result=float(1)
    while number>0:
        result=result*number
        print(result)
        print(number)
        number=number-1
    return result
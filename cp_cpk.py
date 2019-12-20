import numpy as np
import random

class calcul_cp_cpk():

    def Cp(mylist, usl, lsl):

        arr = np.array(mylist)
        arr = arr.ravel()
        sigma = np.std(arr)
        Cp = float(usl - lsl) / (6*sigma)
        return Cp

    def Cpk(mylist, usl, lsl):

        arr = np.array(mylist)
        arr = arr.ravel()
        sigma = np.std(arr)
        m = np.mean(arr)
        Cpu = float(usl - m) / (3*sigma)
        Cpl = float(m - lsl) / (3*sigma)
        Cpk = np.min([Cpu, Cpl])
        return Cpk

#my_list = []
#i=1

#for row in range(100):
#    my_list.insert(i, random.uniform(108.27, 108.07))
#    i += 1
#print(str(Cp(my_list, 108.22, 108.12)))

#print(str(Cpk(my_list, 108.22, 108.12)))



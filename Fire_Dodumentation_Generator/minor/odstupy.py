from chunks import chunks
import os
import sys
import numpy as np
sys.path.insert(0, 'c:/Users/Honza/Google Drive/Work/Generator_zprav/minor/')
from read_results import read_results


def odstup(wb, data_dir):
    odstup = wb.get_sheet_by_name('Odstupové vzdálenosti')
    # Variables, aneb co počítat:
    From_POP = odstup["B2"].value
    POP_singl = odstup["B3"].value

    ''' Varianta 1: zadání otvorů jednotlivě B2 = NE, B3 = ANO'''
    # Načtení zadaných dat
    data = []
    Name_data = []
    PU_data = []
    p_v = []
    h = []
    b = []
    fasada = []
    Nazev_result = []
    p_v_result = []

    odstup_row = odstup.max_row  # determines maximum number of rows
    for row in odstup.iter_rows(min_row=5, max_row=odstup_row,
                                min_col=1, max_col=6):
        for cell in row:
            data.append(cell.value)

    # Zadat data z excelu do jednoho listu
    Odstup_data = list(chunks(data, 6))

    for i in range(0, len(Odstup_data)):
        Name_data.append((Odstup_data[i])[0])
        PU_data.append((Odstup_data[i])[1])
        p_v.append((Odstup_data[i])[2])
        h.append((Odstup_data[i])[3])
        b.append((Odstup_data[i])[4])
        fasada.append((Odstup_data[i])[5])

    ''' Načtení informací o objektu a požárních úsecích '''
    os.chdir(data_dir)
    info1 = read_results('raw_data_info.csv')
    data_PU = read_results('results.csv')
    info = info1[0][0]
    for i in range(0, len(data_PU)):
        Nazev_result.append(data_PU[i][0])
        p_v_result.append(data_PU[i][5])

    if POP_singl == 'ANO' and From_POP == 'NE':
        for i in range(0, len(PU_data)):
            if p_v[i] is None and Nazev_result[i] is not None:
                id = Nazev_result.index(PU_data[i])
                p_v[i] = p_v_result[id]
            if p_v[i] is not None:
                pass
            elif p_v[i] is None and Nazev_result[i] is None:
                raise NameError('!!! Nejsou zadány hodnoty pro výpočet !!!')

    # Změna písmen na čísla + připočtení za hořlavost systému
    p_v = [float(x) for x in p_v]

    if info == 'smíšený':
        p_v = [x+5 for x in p_v]
    elif info == 'hořlavý':
        p_v = [x+10 for x in p_v]
    elif info == 'hořlavý DP3':
        p_v = [x+15 for x in p_v]

    ''' Hustota tepleného toku '''
    p_v_array = np.asarray(p_v)
    print(p_v_array)
    epsilon = 1.00
    T_n = 20 + (345 * np.log10((8 * p_v_array[0]) + 1))
    I_tok = epsilon * np.power((T_n + 273), 4) * 5.67 * 10**-11

    # Configuration factor
    H_f = 5
    W_f = 10

    krit = 18.5
    pol_crit = krit / I_tok
    print(pol_crit)
    res = 1
    r = 0.1


    x = H_f / (r * 2)
    y = W_f / (r * 2)
    first = (x / (np.sqrt((1+x**2)))) / (2*3.14159)
    second = np.arctan((y / (np.sqrt((1+x**2)))))
    A = (first * second) / (2*360)
    third = (y / (np.sqrt((1+y**2))))
    fourth = np.arctan((x / (np.sqrt((1+y**2)))))
    B = (third * fourth) / (2*360)

    while res > pol_crit:
        r += 0.05
        x = H_f / (r * 2)
        y = W_f / (r * 2)
        first = (x / (np.sqrt((1+x**2)))) / (2*3.14159)
        second = np.arctan((y / (np.sqrt((1+x**2)))))
        A = (first * second) / (2*360)
        third = (y / (np.sqrt((1+y**2))))
        fourth = np.arctan((x / (np.sqrt((1+y**2)))))
        B = (third * fourth) / (2*360)
        res = A + B
        print(r, res, pol_crit)

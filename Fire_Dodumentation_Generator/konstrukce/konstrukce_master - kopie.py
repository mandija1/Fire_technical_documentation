from pandas import read_excel, read_csv
import numpy as np
from pylatex import Document, NoEscape, Section, Itemize, Subsection
from pylatex.base_classes.command import Options

###############################################################################
''' Reading input file '''
df = read_excel('./konstrukce_input.xlsx', sheet_name='input konstrukce')
df_clean = df.dropna(subset=['Nazev konstrukce'])
df_input = read_excel('./konstrukce_input.xlsx', sheet_name='Input')


class SPB_obj(object):
    def __init__(self, storage, kcsys):
        self.storage = storage
        self.kcsys = kcsys

    def get_SPB(self):
        if self.storage == 1:
            SPB = 'I'
        if 1 < self.storage <= 3:
            if self.kcsys == "nehořlavý" or self.kcsys == "smíšený":
                SPB = 'II'
            else:
                SPB = 'III'
        return SPB

    def get_odolnost(self, SPB, type, dict):
        df = read_csv('./database/%s_SPB.csv' % SPB)
        idx = dict.get(type)[0]
        df = df.iloc[idx, :]
        if self.kcsys == 'nehořlavý':
            df_druh = read_csv('./database/druh_nehorlavy.csv')
            if self.storage > 1:
                df_druh['NP'][3] = 'DP3'
            df_druh = df_druh.iloc[idx, :]
        if self.kcsys == 'smíšený':
            df_druh = read_csv('./database/druh_smiseny.csv')
            if self.storage > 1 and type != 'E.1 PDK_Stěna':
                df_druh['NP'][0] = 'DP2'
                df_druh['NP'][1] = 'DP2'
                df_druh['NP'][2] = 'DP2'
            df_druh = df_druh.iloc[idx, :]
        if self.kcsys == 'hořlavý DP2' or self.kcsys == 'hořlavý DP3':
            df_druh = read_csv('./database/druh_horlavy.csv')
            df_druh = df_druh.iloc[idx, :]
        PDK_nadzem = [df['NP'], df_druh['NP']]
        PDK_podzem = [df['PP'], df_druh['PP']]
        PDK_posl = [df['pNP'], df_druh['pNP']]
        PDK_mezi = [df['mezi'], df_druh['mezi']]
        return PDK_nadzem, PDK_podzem, PDK_posl, PDK_mezi


dict = {'E.1 PDK_Stěna': [0, '(R)EI'], 'E.1 PDK_Strop': [0, '(R)EI'],
        'E.3 Obvod stěna': [2, '(R)EW'],
        'E.4 Střecha nosné': [3, 'R'],
        'E.5 vnitr nosné kce zajišťující stabilitu': [4, 'R'],
        'E.6 vnější nosné kce zajišťující stab': [5, 'R'],
        'E.9 Konstrukce schodišť mimo CHÚC': [6, 'R'],
        'E.11 Střešní pláště': [7, 'EI'],
        }

RD = SPB_obj(df_input['Hodnota'][1], df_input['Hodnota'][0])
SPB = RD.get_SPB()
PDK_nadzem, PDK_podzem, PDK_posl, PDK_mezi = RD.get_odolnost(SPB, 'E.1 PDK_Stěna', dict)


def cihlyPO(b, interpolate, i):
    myArr = np.asarray(interpolate)
    '''Určení lokace v tabulce'''
    if b[i] < interpolate[0]:
        raise NameError('!!! Mimo rozsah tabulkek !!!')
    if b[i] == interpolate[0]:
        min1 = interpolate[0]
    elif b[i] in interpolate:
        min1 = b[i]
    elif b[i] < interpolate[5] and b[i] not in interpolate:
        min1 = myArr[myArr < b[i]].max()
    elif b[i] >= interpolate[5]:
        min1 = interpolate[5]

    '''Určení požární odolnosti'''
    idx = []
    for k in range(0, len(interpolate)):
        if min1 == interpolate[k]:
            idx.append(k)
    idd = max(idx)
    if idd == 0:
        check = 30
    if idd == 1:
        check = 45
    if idd == 2:
        check = 60
    if idd == 3:
        check = 90
    if idd == 4:
        check = 120
    if idd == 5:
        check = 180
    return check


##############################################################################
''' Požárně dělící stěny'''
spec_dict_zb = {'ŽB. desky – výztuž v jednom směru':
                [[60, 70, 80, 100, 120, 150], [10, 15, 20, 30, 40, 55]],
                'ŽB. nosné stěny':
                [[120, 125, 130, 140, 160, 210], [10, 10, 10, 25, 35, 50]],
                'ŽB. lokálně podepřená deska':
                [[150, 170, 180, 200, 200, 200], [10, 15, 15, 25, 35, 45]]
                }


def zb_odolnost(spec_dict, frame, i):
    frame_minor = frame.iloc[i]
    list_tlouska = spec_dict.get(frame_minor['Specifikace'])[0]
    list_vyztuz = spec_dict.get(frame_minor['Specifikace'])[1]
    if frame['b'][0] < list_tlouska[0] and frame['a / rho'][0] < list_vyztuz[0]:
        raise NameError(
            '!!! Nedostatečná požární odolnost konstrukce {} !!!'
            .format(frame['Nazev konstrukce'][i])
        )
    if list_vyztuz[0] <= frame['a / rho'][0] < list_vyztuz[1]:
        if frame['b'][0] >= list_tlouska[0]:
            check = 30
    if list_vyztuz[1] <= frame['a / rho'][0] < list_vyztuz[2]:
        if frame['b'][0] >= list_tlouska[1]:
            check = 45
    if list_vyztuz[2] <= frame['a / rho'][0] < list_vyztuz[3]:
        if frame['b'][0] >= list_tlouska[2]:
            check = 60
    if list_vyztuz[3] <= frame['a / rho'][0] < list_vyztuz[4]:
        if frame['b'][0] >= list_tlouska[3]:
            check = 90
    if list_vyztuz[4] <= frame['a / rho'][0] < list_vyztuz[5]:
        if frame['b'][0] >= list_tlouska[4]:
            check = 120
    if list_vyztuz[5] <= frame['a / rho'][0]:
        if frame['b'][0] >= list_tlouska[5]:
            check = 180
    if frame['Specifikace'][0] == 'ŽB. nenosné stěny':
        token = 'EI'
    else:
        token = 'REI'
    return check, token

odolnost = []
token_spec = []
check_already = []
df_PDK_steny = read_csv("./database/tvarnice.csv")
df_PDK_stropy = read_csv("./database/stropy.csv")

for i in range(len(df_clean['Nazev konstrukce'])):
    if df_clean['Typ konstrukce'][i] == 'E.1 PDK_Stěna':
        if df_clean['Specifikace'][i] == 'Cihly plné pálené (skupina 1S)':
            interpolate = [90, 90, 90, 100, 140, 190]
            token = 'REI'
        if df_clean['Specifikace'][i] == 'Neznámé zdivo (skupina 4)':
            interpolate = [240, 240, 240, 300, 365, 425]
            token = 'REI'
        if df_clean['Specifikace'][i] == 'Nenosné (všechny skupiny)':
            interpolate = [70, 70, 70, 100, 140, 140]
            token = 'EI'
        if df_clean['Specifikace'][i] == 'ŽB. nenosné stěny':
            if df_clean['b'][i] < 60:
                raise NameError('!!! Mimo rozsah tabulkek !!!')
            if 60 <= df_clean['b'][i] < 70:
                check = 30
            if 70 <= df_clean['b'][i] < 80:
                check = 45
            if 80 <= df_clean['b'][i] < 100:
                check = 60
            if 100 <= df_clean['b'][i] < 120:
                check = 90
            if 120 <= df_clean['b'][i] < 150:
                check = 120
            if df_clean['b'][i] >= 150:
                check = 180
            token = 'EI'
            check_already.append(df_clean['Specifikace'][i])

        if df_clean['Specifikace'][i] == 'ŽB. nosné stěny':
            check, token = zb_odolnost(spec_dict_zb, df_clean, i)
            check_already.append(df_clean['Specifikace'][i])

        if df_clean['Specifikace'][i] == 'ŽB nosné stěny – neznámé a':
            if df_clean['b'][i] >= 120:
                check = 30
                token = 'REI'
                check_already.append(df_clean['Specifikace'][i])
            else:
                raise NameError(
                    '!!! Nedostatečná požární odolnost konstrukce {} !!!'
                    .format(df_clean['Nazev konstrukce'][i])
                )
        if df_clean['Specifikace'][i] == 'Vápenopískové nenosné':
            interpolate = [70, 90, 90, 100, 140, 170]
            token = 'EI'
        if df_clean['Specifikace'][i] == 'Vápenopískové nosné':
            interpolate = [100, 100, 100, 100, 200, 240]
            token = 'REI'
        if df_clean['Specifikace'][i] == 'Betonové tvárnice nenosné':
            interpolate = [50, 70, 100, 100, 200, 200]
            token = 'EI'
        if df_clean['Specifikace'][i] == 'Betonové tvárnice nosné':
            interpolate = [170, 170, 170, 170, 190, 240]
            token = 'REI'
        if df_clean['Specifikace'][i] == 'Pórobetonové tvárnice nenosné':
            interpolate = [65, 70, 75, 100, 100, 150]
            token = 'EI'
        if df_clean['Specifikace'][i] == 'Pórobetonové tvárnice nosné':
            interpolate = [115, 115, 140, 200, 225, 300]
            token = 'REI'

        if df_clean['Specifikace'][i] in df_PDK_steny['Nazev'].values:
            index = df_PDK_steny.index[df_PDK_steny['Nazev'] == df_clean['Specifikace'][i]].tolist()
            check = df_PDK_steny['odolnost'][index[0]]
            token = df_PDK_steny['klasifikace'][index[0]]
            check_already.append(df_clean['Specifikace'][i])

        if df_clean['Specifikace'][i] not in check_already:
            check = cihlyPO(df_clean['b'], interpolate, i)

    ###########################################################################
    ''' Stropy'''
    ###########################################################################
    if df_clean['Typ konstrukce'][i] == 'E.1 PDK_Strop':
        if df_clean['Specifikace'][i] == 'ŽB. desky – výztuž v jednom směru':
            check, token = zb_odolnost(spec_dict_zb, df_clean, i)
            check_already.append(df_clean['Specifikace'][i])
        if df_clean['Specifikace'][i] == 'ŽB. desky – neznámá krycí tloušťka a':
            if df_clean['b'][i] >= 60:
                check = 30
                token = 'REI'
                check_already.append(df_clean['Specifikace'][i])
            else:
                raise NameError(
                    '!!! Nedostatečná požární odolnost konstrukce {} !!!'
                    .format(df_clean['Nazev konstrukce'][i])
                )
        if df_clean['Specifikace'][i] == 'ŽB. lokálně podepřená deska':
            check, token = zb_odolnost(spec_dict_zb, df_clean, i)
            check_already.append(df_clean['Specifikace'][i])

        df_PDK = read_csv("./database/stropy.csv")

        if df_clean['Specifikace'][i] in df_PDK_stropy['Nazev'].values:
            index = df_PDK_stropy.index[df_PDK_stropy['Nazev'] == df_clean['Specifikace'][i]].tolist()
            check = df_PDK_stropy['odolnost'][index[0]]
            token = df_PDK_stropy['klasifikace'][index[0]]
            check_already.append(df_clean['Specifikace'][i])

    odolnost.append(check)
    token_spec.append(token)
    
    ###########################################################################
    ''' Ověření požární odolnosti'''

    if df_clean['Pozn.'][i] == 'Nadzemní podlaží':
        if odolnost[i] >= PDK_nadzem[0]:
            pass
        else:
            raise NameError(
                '!!! Nedostatečná požární odolnost konstrukce {} !!!'
                .format(df_clean['Nazev konstrukce'][i])
            )
    if df_clean['Pozn.'][i] == 'Podzemní podlaží':
        if odolnost[i] >= PDK_podzem[0]:
            pass
        else:
            raise NameError(
                '!!! Nedostatečná požární odolnost konstrukce {} !!!'
                .format(df_clean['Nazev konstrukce'][i])
            )
    if df_clean['Pozn.'][i] == 'Poslední nadzemní podlaží':
        if odolnost[i] >= PDK_posl[0]:
            pass
        else:
            raise NameError(
                '!!! Nedostatečná požární odolnost konstrukce {} !!!'
                .format(df_clean['Nazev konstrukce'][i])
            )
    if df_clean['Pozn.'][i] == 'Mezi objekty':
        if odolnost[i] >= PDK_mezi[0]:
            pass
        else:
            raise NameError(
                '!!! Nedostatečná požární odolnost konstrukce {} !!!'
                .format(df_clean['Nazev konstrukce'][i])
            )

print(odolnost)
print(token_spec)

df_clean = df_clean.drop(df_clean.iloc[:, 6:], axis=1)
df_clean['odolnost'] = odolnost
df_clean['token_spec'] = token_spec


class E_generator(object):
    def __init__(self, frame, RD, dict):
        self.frame = frame
        self.RD = RD
        self.token = dict

    def preamble(self):
        geometry_options = {
            "margin": "0.5cm",
            "includeheadfoot": True
            }
        doc = Document(page_numbers=True, geometry_options=geometry_options)
        doc.preamble.append(NoEscape(r'\usepackage[czech]{babel}'))
        doc.documentclass.options = Options('10pt')
        return doc

    def E_1(self):
        df_popis = read_excel('./konstrukce_input.xlsx', sheet_name='Databaze_PKD')
        with doc.create(Section('Zhodnocení požární odolnosti konstrukcí')):
            SPB = self.RD.get_SPB()
            uniques = self.frame['Typ konstrukce'].unique().tolist()
            if 'E.1 PDK_Stěna' in uniques or 'E.1 PDK_Strop' in uniques:
                with doc.create(Subsection('Požárně dělící konstrukce')):
                    frame_choose = self.frame.loc[self.frame['Typ konstrukce'].isin(['E.1 PDK_Stěna', 'E.1 PDK_Strop'])]
                with doc.create(Itemize()) as itemize:
                    for n in range(len(frame_choose['Typ konstrukce'])):
                        PDK_nadzem, PDK_podzem, PDK_posl, PDK_mezi = RD.get_odolnost(SPB, self.frame['Typ konstrukce'][n], dict)
                        ozn = {'Nadzemní podlaží': ['v NP', PDK_nadzem],
                               'Podzemní podlaží': ['v PP', PDK_podzem],
                               'Poslední nadzemní podlaží': ['v posledním NP', PDK_posl],
                               'Mezi objekty': ['mezi objekty', PDK_mezi],
                               }
                        ozn_def = ozn.get(frame_choose['Pozn.'][n])[0]
                        token_need = self.token.get(frame_choose['Typ konstrukce'][n])[1]
                        odolnost_need = int(ozn.get(frame_choose['Pozn.'][n])[1][0])
                        druh_need = ozn.get(frame_choose['Pozn.'][n])[1][1]
                        itemize.add_item("Normový požadavek {} - {}.SPB - {} {} {}".format(ozn_def, SPB, token_need, odolnost_need, druh_need))
            for n in range(len(frame_choose['Typ konstrukce'])):
                doc.append(NoEscape(r'\textbf{%s}: ' % self.frame['Nazev konstrukce'][n]))
                popis = df_popis.loc[df_popis['Název'] == self.frame['Specifikace'][n]]
                doplnit = ['y {} mm', 'e {} mm', 'a = {} mm', '{} D']
                adder = []
                counter = []
                for k in range(len(doplnit)):
                    if doplnit[k] in popis['Popis'].values[0]:
                        if doplnit[k] == 'y {} mm' or doplnit[k] == 'e {} mm':
                            adder.append(self.frame['b'].values[n])
                            counter.append(k)
                        if doplnit[k] == 'a = {} mm':
                            adder.append(self.frame['a / rho'].values[n])
                            counter.append(k)
                        if doplnit[k] == '{} D':
                            adder.append(self.frame['odolnost'].values[n])
                            counter.append(k)
                        else:
                            pass
                if len(adder) == 2:
                    text = popis['Popis'].values[0].format(adder[0], adder[1])
                if len(adder) == 3:
                    text = popis['Popis'].values[0].format(adder[0], adder[1], adder[2])
                if len(adder) == 0:
                    text = popis['Popis'].values[0]
                doc.append(NoEscape(text))
                doc.append(NoEscape(r'\par'))


        doc.generate_pdf("E.1", clean_tex=False)


E1top = E_generator(df_clean, RD, dict)
doc = E1top.preamble()
E1top.E_1()

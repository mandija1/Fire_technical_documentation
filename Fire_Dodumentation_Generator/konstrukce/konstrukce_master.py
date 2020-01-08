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


def cihlyPO(b, interpolate, i):
    myArr = np.asarray(interpolate)
    '''Určení lokace v tabulce'''
    if b.values[i] < interpolate[0]:
        raise NameError('!!! Mimo rozsah tabulkek !!!')
    if b.values[i] == interpolate[0]:
        min1 = interpolate[0]
    elif b.values[i] in interpolate:
        min1 = b.values[i]
    elif b.values[i] < interpolate[5] and b.values[i] not in interpolate:
        min1 = myArr[myArr < b[i]].max()
    elif b.values[i] >= interpolate[5]:
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
    '''Stanovení požární odolnosti železobetonových konstrukcí
        provedeno podle tabulkového hodnocení Zoufal a spol.
        '''
    frame_minor = frame.iloc[i]
    list_tlouska = spec_dict.get(frame_minor['Specifikace'])[0]
    list_vyztuz = spec_dict.get(frame_minor['Specifikace'])[1]
    if frame['b'].values[0] < list_tlouska[0] and frame['a / rho'].values[0] < list_vyztuz[0]:
        raise NameError(
            '!!! Nedostatečná požární odolnost konstrukce {} !!!'
            .format(frame['Nazev konstrukce'][i])
        )
    if list_vyztuz[0] <= frame['a / rho'].values[0] < list_vyztuz[1]:
        if frame['b'].values[0] >= list_tlouska[0]:
            check = 30
    if list_vyztuz[1] <= frame['a / rho'].values[0] < list_vyztuz[2]:
        if frame['b'].values[0] >= list_tlouska[1]:
            check = 45
    if list_vyztuz[2] <= frame['a / rho'].values[0] < list_vyztuz[3]:
        if frame['b'].values[0] >= list_tlouska[2]:
            check = 60
    if list_vyztuz[3] <= frame['a / rho'].values[0] < list_vyztuz[4]:
        if frame['b'].values[0] >= list_tlouska[3]:
            check = 90
    if list_vyztuz[4] <= frame['a / rho'].values[0] < list_vyztuz[5]:
        if frame['b'].values[0] >= list_tlouska[4]:
            check = 120
    if list_vyztuz[5] <= frame['a / rho'].values[0]:
        if frame['b'].values[0] >= list_tlouska[5]:
            check = 180
    if frame['Specifikace'].values[0] == 'ŽB. nenosné stěny':
        token = 'EI'
    else:
        token = 'REI'
    return check, token


def pozarni_odolnost_sten(i, frame, check_already, data):
    if frame['Specifikace'].values[i] == 'Cihly plné pálené (skupina 1S)':
        interpolate = [90, 90, 90, 100, 140, 190]
        token = 'REI'
    if frame['Specifikace'].values[i] == 'Neznámé zdivo (skupina 4)':
        interpolate = [240, 240, 240, 300, 365, 425]
        token = 'REI'
    if frame['Specifikace'].values[i] == 'Nenosné (všechny skupiny)':
        interpolate = [70, 70, 70, 100, 140, 140]
        token = 'EI'
    if frame['Specifikace'].values[i] == 'ŽB. nenosné stěny':
        if frame['b'].values[i] < 60:
            raise NameError('!!! Mimo rozsah tabulkek !!!')
        if 60 <= frame['b'][i] < 70:
            check = 30
        if 70 <= frame['b'][i] < 80:
            check = 45
        if 80 <= frame['b'][i] < 100:
            check = 60
        if 100 <= frame['b'][i] < 120:
            check = 90
        if 120 <= frame['b'][i] < 150:
            check = 120
        if frame['b'][i] >= 150:
            check = 180
        token = 'EI'
        check_already.append(frame['Specifikace'][i])

    if frame['Specifikace'].values[i] == 'ŽB. nosné stěny':
        check, token = zb_odolnost(spec_dict_zb, frame, i)
        check_already.append(frame['Specifikace'].values[i])

    if frame['Specifikace'].values[i] == 'ŽB nosné stěny – neznámé a':
        if frame['b'][i] >= 120:
            check = 30
            token = 'REI'
            check_already.append(frame['Specifikace'][i])
        else:
            raise NameError(
                '!!! Nedostatečná požární odolnost konstrukce {} !!!'
                .format(frame['Nazev konstrukce'][i])
            )
    if frame['Specifikace'].values[i] == 'Vápenopískové nenosné':
        interpolate = [70, 90, 90, 100, 140, 170]
        token = 'EI'
    if frame['Specifikace'].values[i] == 'Vápenopískové nosné':
        interpolate = [100, 100, 100, 100, 200, 240]
        token = 'REI'
    if frame['Specifikace'].values[i] == 'Betonové tvárnice nenosné':
        interpolate = [50, 70, 100, 100, 200, 200]
        token = 'EI'
    if frame['Specifikace'].values[i] == 'Betonové tvárnice nosné':
        interpolate = [170, 170, 170, 170, 190, 240]
        token = 'REI'
    if frame['Specifikace'].values[i] == 'Pórobetonové tvárnice nenosné':
        interpolate = [65, 70, 75, 100, 100, 150]
        token = 'EI'
    if frame['Specifikace'].values[i] == 'Pórobetonové tvárnice nosné':
        interpolate = [115, 115, 140, 200, 225, 300]
        token = 'REI'

    if frame['Specifikace'].values[i] in data['Nazev'].values:
        index = data.index[data['Nazev'] == frame['Specifikace'][i]].tolist()
        check = data['odolnost'][index[0]]
        token = data['klasifikace'][index[0]]
        check_already.append(frame['Specifikace'][i])

    if frame['Specifikace'].values[i] not in check_already and (frame['Typ konstrukce'].values[i] == 'E.1 PDK_Stěna' or frame['Typ konstrukce'].values[i] == 'E.3 Obvod stěna'):
        check = cihlyPO(frame['b'], interpolate, i)

    return check, check_already, token


def check_odolnost(df, name, odolnost, token_spec):
    '''Porovnává požadovanou a stanovenou požární odolnost
        '''
    PDK_nadzem, PDK_podzem, PDK_posl, PDK_mezi = RD.get_odolnost(SPB, name, dict)
    if df['Pozn.'].values[i] == 'Nadzemní podlaží':
        if odolnost[i] >= PDK_nadzem[0]:
            pass
        else:
            raise NameError(
                '!!! Nedostatečná požární odolnost konstrukce {} !!!'
                .format(df['Nazev konstrukce'][i])
            )
    if df['Pozn.'].values[i] == 'Podzemní podlaží':
        if odolnost[i] >= PDK_podzem[0]:
            pass
        else:
            raise NameError(
                '!!! Nedostatečná požární odolnost konstrukce {} !!!'
                .format(df['Nazev konstrukce'][i])
            )
    if df['Pozn.'].values[i] == 'Poslední nadzemní podlaží':
        if odolnost[i] >= PDK_posl[0]:
            pass
        else:
            raise NameError(
                '!!! Nedostatečná požární odolnost konstrukce {} !!!'
                .format(df['Nazev konstrukce'][i])
            )
    if df['Pozn.'].values[i] == 'Mezi objekty':
        if odolnost[i] >= PDK_mezi[0]:
            pass
        else:
            raise NameError(
                '!!! Nedostatečná požární odolnost konstrukce {} !!!'
                .format(df['Nazev konstrukce'][i])
            )


###############################################################################
###############################################################################
''' Načtení dat a rozdělení dat do jednotlivých dataframes'''
if 'E.1 PDK_Stěna' in df_clean['Typ konstrukce'].values and 'E.1 PDK_Stěna' in df_clean['Typ konstrukce'].values:
    df_E1 = df_clean.loc[df_clean['Typ konstrukce'].isin(['E.1 PDK_Stěna', 'E.1 PDK_Strop'])]
    df_PDK_steny = read_csv("./database/tvarnice.csv")
    df_PDK_stropy = read_csv("./database/stropy.csv")
elif 'E.1 PDK_Stěna' in df_clean['Typ konstrukce'].values and 'E.1 PDK_Stěna' not in df_clean['Typ konstrukce'].values:
    df_E1 = df_clean.loc[df_clean['Typ konstrukce'] == 'E.1 PDK_Stěna']
    df_PDK_steny = read_csv("./database/tvarnice.csv")
elif 'E.1 PDK_Stěna' not in df_clean['Typ konstrukce'].values and 'E.1 PDK_Strop' in df_clean['Typ konstrukce'].values:
    df_E1 = df_clean.loc[df_clean['Typ konstrukce'] == 'E.1 PDK_Strop']
    df_PDK_stropy = read_csv("./database/stropy.csv")
if 'E.3 Obvod stěna' in df_clean['Typ konstrukce'].values:
    df_E3 = df_clean.loc[df_clean['Typ konstrukce'] == 'E.3 Obvod stěna']
    df_obvod_steny = read_csv("./database/tvarnice.csv")
###############################################################################

'''Obvodové stěny'''
try:
    df_E3
except NameError:
    df_E3 = None
if df_E3 is not None:
    odolnost_E3 = []
    token_spec_E3 = []
    check_already_E3 = []
    for i in range(len(df_E3['Nazev konstrukce'])):
        check, check_already_E1, token = pozarni_odolnost_sten(i, df_E3, check_already_E3, df_obvod_steny)
        odolnost_E3.append(check)
        token_spec_E3.append(token)
    # Ověření požární odolnosti
    check_odolnost(df_E3, 'E.3 Obvod stěna', odolnost_E3, token_spec_E3)
    # Vyčistnění a připravení dat
    df_E3 = df_E3.drop(df_E3.iloc[:, 6:], axis=1)
    df_E3['odolnost'] = odolnost_E3
    df_E3['token_spec'] = token_spec_E3

'''Požární uzávěry'''
try:
    df_E2
except NameError:
    df_E2 = None

'''Požárně dělící stěny'''
try:
    df_E1
except NameError:
    df_E1 = None
if df_E1 is not None:
    odolnost_E1 = []
    token_spec_E1 = []
    check_already_E1 = []
    for i in range(len(df_E1['Nazev konstrukce'])):
        if df_E1['Typ konstrukce'].values[i] == 'E.1 PDK_Stěna':
            check, check_already_E1, token = pozarni_odolnost_sten(i, df_E1, check_already_E1, df_PDK_steny)

        #######################################################################
        ''' Stropy'''
        #######################################################################
        if df_E1['Specifikace'].values[i] == 'ŽB. desky – výztuž v jednom směru':
            check, token = zb_odolnost(spec_dict_zb, df_E1, i)
            check_already_E1.append(df_E1['Specifikace'][i])
        if df_E1['Specifikace'].values[i] == 'ŽB. desky – neznámá krycí tloušťka a':
            if df_E1['b'][i] >= 60:
                check = 30
                token = 'REI'
                check_already.append(df_E1['Specifikace'][i])
            else:
                raise NameError(
                    '!!! Nedostatečná požární odolnost konstrukce {} !!!'
                    .format(df_E1['Nazev konstrukce'][i])
                )
        if df_E1['Specifikace'].values[i] == 'ŽB. lokálně podepřená deska':
            check, token = zb_odolnost(spec_dict_zb, df_E1, i)
            check_already_E1.append(df_E1['Specifikace'][i])

        df_PDK = read_csv("./database/stropy.csv")

        if df_E1['Specifikace'].values[i] in df_PDK_stropy['Nazev'].values:
            index = df_PDK_stropy.index[df_PDK_stropy['Nazev'] == df_E1['Specifikace'][i]].tolist()
            check = df_PDK_stropy['odolnost'][index[0]]
            token = df_PDK_stropy['klasifikace'][index[0]]
            check_already_E1.append(df_E1['Specifikace'][i])

        odolnost_E1.append(check)
        token_spec_E1.append(token)

        #######################################################################
        ''' Ověření požární odolnosti'''
        check_odolnost(df_E1, 'E.1 PDK_Stěna', odolnost_E1, token_spec_E1)

    if df_E1 is not None:
        df_E1 = df_E1.drop(df_E1.iloc[:, 6:], axis=1)
        df_E1['odolnost'] = odolnost_E1
        df_E1['token_spec'] = token_spec_E1


class E_generator(object):
    def __init__(self, RD, dict):
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
        with doc.create(Section('Zhodnocení požární odolnosti konstrukcí')):
            doc.append(NoEscape(r'Normový požadavek na požární odolnost konstrukcí je stanoven na základě stupně požární bezpečnosti jednotlivých požárních úseků (viz tabulka \ref{PU}), a to pomocí tabulky 12 v normě ČSN 73 0802.'))
        return doc

    def E_podkapitola(self, frame):
        df_popis = read_excel('./konstrukce_input.xlsx', sheet_name='Databaze_PKD')
        SPB = self.RD.get_SPB()
        subsection_names = {'E.1 PDK_Stěna': 'Požární stěny a požární stropy',
                            'E.3 Obvod stěna': 'Obvodové stěny'
                            }
        subsection_name_choose = subsection_names.get(frame['Typ konstrukce'].values[0])
        with doc.create(Subsection(subsection_name_choose)):
            with doc.create(Itemize()) as itemize:
                for n in range(len(frame['Typ konstrukce'])):
                    PDK_nadzem, PDK_podzem, PDK_posl, PDK_mezi = RD.get_odolnost(SPB, frame['Typ konstrukce'].values[n], dict)
                    ozn = {'Nadzemní podlaží': ['v NP', PDK_nadzem],
                           'Podzemní podlaží': ['v PP', PDK_podzem],
                           'Poslední nadzemní podlaží': ['v posledním NP', PDK_posl],
                           'Mezi objekty': ['mezi objekty', PDK_mezi],
                           }
                    ozn_def = ozn.get(frame['Pozn.'].values[n])[0]
                    token_need = self.token.get(frame['Typ konstrukce'].values[n])[1]
                    odolnost_need = int(ozn.get(frame['Pozn.'].values[n])[1][0])
                    druh_need = ozn.get(frame['Pozn.'].values[n])[1][1]
                    itemize.add_item("Normový požadavek {} - {}.SPB - {} {} {}".format(ozn_def, SPB, token_need, odolnost_need, druh_need))
            for n in range(len(frame['Typ konstrukce'])):
                doc.append(NoEscape(r'\textbf{%s}: ' % frame['Nazev konstrukce'].values[n]))
                popis = df_popis.loc[df_popis['Název'] == frame['Specifikace'].values[n]]
                doplnit = ['y {} mm', 'e {} mm', 'a = {} mm', '{} D']
                adder = []
                counter = []
                for k in range(len(doplnit)):
                    if doplnit[k] in popis['Popis'].values[0]:
                        if doplnit[k] == 'y {} mm' or doplnit[k] == 'e {} mm':
                            adder.append(frame['b'].values[n])
                            counter.append(k)
                        if doplnit[k] == 'a = {} mm':
                            adder.append(frame['a / rho'].values[n])
                            counter.append(k)
                        if doplnit[k] == '{} D':
                            adder.append(frame['odolnost'].values[n])
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

    def E_podkapitola_none(self, name):
        subsection_names = {'E.1 PDK_Stěna': ['Požární stěny a požární stropy', r'Objekt je jedním požárním úsekem. Požární stěny ani požární stropy tak nejsou navrženy. Tento bod je dále považován za vyhovující.'],
                            'E.3 Obvod stěna': ['Obvodové stěny', r'Obvodové konstrukce nemusí v souladu s čl. 8.4.3 normy ČSN 73 0802 vykazovat požární odolnost. V souladu se stejným článkem jsou obvodové stěny dále považovány za zcela požárně otevřené plochy.'],
                            'E.2 Otvory': ['Požární uzávěry', r'V objektu nejsou naprženy průchody přes požární úseky. Dveře s požární odolností tak nebudou v objektu instalovány. Tento bod je tak možné dále považovat za vyhovující.']
                            }
        subsection_name_choose = subsection_names.get(name)
        with doc.create(Subsection(subsection_name_choose[0])):
            doc.append(NoEscape(subsection_name_choose[1]))


Etop = E_generator(RD, dict)
doc = Etop.preamble()
if df_E1 is None:
    Etop.E_podkapitola_none('E.1 PDK_Stěna')
else:
    Etop.E_podkapitola(df_E1)
    doc.append(NoEscape(r'\textbf{Styk požární stěny s požárním stropem:} '))
    if df_input['Hodnota'][4] == 'ANO':
        doc.append(NoEscape(r'Výše popsané požární stěny a požární stropy s požární odolností jsou druhu DP1 a plně se stýkají, což je v souladu s čl. 8.2.4 normy ČSN 73 0802 plně stýkají.'))
    if df_input['Hodnota'][4] == 'NE':
        doc.append(NoEscape(r'Objekt je dle návrhu samostatně stojící budovou. Styk požární stěny s požárním stropem pro zabránění přenosu požáru na jiný objekt tak není nutné podlaž čl. 8.2.4 normy ČSN 73 0802 zajišťovat.'))
    if df_input['Hodnota'][4] == 'ANO – podhled':
        doc.append(NoEscape(r'V objektu je nutné splnit požadavky čl. 8.2.4 normy ČSN 73 0802. Vzhledem k tomu, že převýšení střešního pláště požárními stěnami není navrženo, musí navrhované sádrokarto\-nové podhledy ve 2.NP vykazovat společně s nosnou konstrukcí střechy požární odolnost minimálně REI 15 (hodnota z tabulky 12 normy ČSN 73 0802 pro poslední nadzemní podlaží) a střešní krytina musí vykazovat klasifikaci B$_{ROOF}$(t3).'))
        doc.append(NoEscape(r'\par'))
        doc.append(NoEscape(r'Při splnění těchto požadavků bude možné konstrukci střechy považovat v souladu s čl. 3.2.4 normy ČSN 73 0810 za konstrukci druhu DP2. Konstrukce bude zároveň vykazovat požadovanou požární odolnost. Navrhovanou střešní krytinu (plechová krytina) je možné z hlediska tohoto požadavku dále považovat za vyhovující. Podhled musí vykazovat minimálně 15-ti minutovou požární odolnost. Nový podhled bude proveden v souladu s požadavky kapitoly \ref{cha-12}.'))
    if df_input['Hodnota'][4] == 'ANO – přesah':
        doc.append(NoEscape(r'Střešní plášť je na hranici požárních úseků převýšen výše popsanou požární stěnou o více než 300 mm. Toto řešení je v souladu s čl. 8.2.4 normy ČSN 73 0802 dále považováno za vyhovující.'))
if df_E2 is None:
    Etop.E_podkapitola_none('E.2 Otvory')
else:
    Etop.E_podkapitola(df_E2)

if df_E3 is None:
    Etop.E_podkapitola_none('E.3 Obvod stěna')
else:
    Etop.E_podkapitola(df_E3)

doc.generate_pdf("E.1", clean_tex=False)

import os
# from openpyxl import load_workbook
import pandas as pd
import numpy as np
import warnings
from pylatex import Document, LongTable, NoEscape, Command, Section, Subsection, Figure, MultiColumn, utils
from pylatex.base_classes.command import Options
from tisk_odstup import tisk_odstup

dir_path = os.path.dirname(os.path.realpath(__file__))
os.chdir(os.path.dirname(__file__))
pd.options.mode.chained_assignment = None  # default='warn'

###############################################################################
''' Reading input file '''
# file = 'odstup_zadani.xlsx'  # Name of the file
df = pd.read_excel('input.xlsx', sheet_name='Odstupy_zadani')
df_input = pd.read_excel('input.xlsx', sheet_name='Input')
df_clean = df.dropna(subset=['Fasada'])

''' Pracovní hodnoty pro výpočet '''
p_v = df_input["Hodnota"][1]
kcni_system = df_input["Hodnota"][0]

if kcni_system == 'hořlavý DP3':
    add_pv = 15
    p_v = p_v + add_pv
elif kcni_system == 'hořlavý DP2':
    add_pv = 10
    p_v = p_v + add_pv
elif kcni_system == 'smíšený':
    add_pv = 5
    p_v = p_v + add_pv
else:
    p_v = p_v

epsilon = 1.0
r = 0.1
krit = 18.5

''' Hustota tepleného toku '''
T_n = 20 + (345 * np.log10((8 * p_v) + 1))
I_tok = epsilon * np.power((T_n + 273), 4) * 5.67 * 10**-11

''' Konfigurační faktor '''
bi = df_clean.loc[:, 'sirka'].div(2, axis=0)
hi = df_clean.loc[:, 'vyska'].div(2, axis=0)

df_clean['bi'] = bi
df_clean['hi'] = hi

pol_crit = krit / I_tok

d_results = []
d_results_kraj = []
index_store = []
POP_store = []


def odstup(hi, bi, crit, i):
    res = 1
    d = 0.1
    while res > crit:
        A = hi / d
        B = bi / d
        first = (A / (np.sqrt((1+A**2))))
        second = (B / (np.sqrt((1+A**2))))
        A1 = (first * np.arctan(second))
        third = (B / (np.sqrt((1+B**2))))
        fourth = (A / (np.sqrt((1+B**2))))
        B1 = (third * np.arctan(fourth))
        res = (A1 + B1) / (2*np.pi) * 4
        d += 0.05
    return round(d, 2)


def odstup_kraj(hi, bi, crit, i):
    d_kraj = 0.1
    res_1 = 1
    while res_1 > crit:
        A = hi / d_kraj
        B = bi / d_kraj * 2
        first = (A / (np.sqrt((1+A**2))))
        second = (B / (np.sqrt((1+A**2))))
        A1 = (first * np.arctan(second))
        third = (B / (np.sqrt((1+B**2))))
        fourth = (A / (np.sqrt((1+B**2))))
        B1 = (third * np.arctan(fourth))
        res_1 = (A1 + B1) / (2*np.pi) * 2
        d_kraj += 0.05
    return round(d_kraj, 2)


POP_store_list = []
d_results_za_krajem = []
I_tok_list = []
pv_list = []
for i in range(len(bi)):
    if df_clean['Samostatne'][i] == 'ANO':
        d = odstup(hi[i], bi[i], pol_crit, i)
        d_results.append(d)
        d_kraj = odstup_kraj(hi[i], bi[i], pol_crit, i)
        d_results_kraj.append(d_kraj)
        POP_store_list.append(100.00)
        pv_list.append(p_v)
        I_tok_list.append(I_tok)
    if df_clean['Samostatne'][i] == 'Přístřešky pro auta DP3':
        T_n_p = 20 + (345 * np.log10((8 * 30) + 1))
        I_tok_p = epsilon * np.power((T_n_p + 273), 4) * 5.67 * 10**-11
        pol_crit_pristresek = krit / I_tok_p
        hi_p = 1.5 / 2
        d = odstup(hi_p, bi[i], pol_crit_pristresek, i)
        d_results.append(d)
        d_kraj = odstup_kraj(hi_p, bi[i], pol_crit_pristresek, i)
        d_results_kraj.append(d_kraj)
        POP_store_list.append(100.00)
        df_clean['vyska'].iloc[[i]] = 1.5
        pv_list.append(30)
        I_tok_list.append(I_tok_p)
    if df_clean['Samostatne'][i] == 'Pergoly / přístřešky':
        T_n_perg = 658.0
        I_tok_perg = epsilon * np.power((T_n_perg + 273), 4) * 5.67 * 10**-11
        pol_crit_perg = krit / I_tok_perg
        d = odstup(hi[i], bi[i], pol_crit_perg, i)
        d_results.append(d)
        d_kraj = odstup_kraj(hi[i], bi[i], pol_crit_perg, i)
        d_results_kraj.append(d_kraj)
        POP_store_list.append(100.00)
        pv_list.append(8.71)
        I_tok_list.append(I_tok_perg)
    if df_clean['Samostatne'][i] == 'NE – Fasada':
        S_store_window = 0.0
        S_store = 0.0
        POP_store = 0.0
        index_store.append(i)
        if df_clean['Samostatne'][i+1] == 'NE – Okno':
            if df_clean['Nazev/fasada'][i] == df_clean['Nazev/fasada'][i+1]:
                S_store = df_clean['vyska'][i] * df_clean['sirka'][i]
                for o in (range(i+1, len(bi))):
                    if df_clean['Nazev/fasada'][i] == df_clean['Nazev/fasada'][o]:
                        S_store_window += (df_clean['vyska'][o] * df_clean['sirka'][o])
                if S_store_window > S_store:
                    raise ValueError('Error: Plocha oken je vetsi nez plocha fasady {}'.format(df_clean['Nazev/fasada'][i]))
                else:
                    POP_store = S_store_window/S_store
                if POP_store < 0.4:
                    warnings.warn('Pozor hodnota POP u {} je pouze {} %. Vhodnejsi je pocitat otvory samostatne'.format(df_clean['Nazev/fasada'][i], POP_store))
                    POP_store = 0.4
                # POP_store.append(S_store_window/S_store)
            else:
                raise ValueError('Nazev fasady neni konzistentni mezi radky {} a {}'.format(i+1, i+2))
        pol_crit_POP = krit / (I_tok * POP_store)
        d = odstup(hi[i], bi[i], pol_crit_POP, i)
        d_results.append(d)
        d_kraj = d
        d_results_kraj.append(d_kraj)
        POP_store_list.append(round(POP_store * 100, 2))
        pv_list.append(p_v)
        I_tok_list.append(I_tok)

for i in range(len(d_results_kraj)):
    d_results_za_krajem.append(round(d_results_kraj[i] / 2, 2))
    I_tok_list[i] = round(I_tok_list[i] * POP_store_list[i] / 100, 2)

df_result = df_clean.loc[df_clean['Samostatne'].isin(['ANO', 'NE – Fasada', 'Přístřešky pro auta DP3', 'Pergoly / přístřešky'])][['Nazev/fasada', 'vyska', 'sirka', 'Samostatne']]
df_result['POP'] = POP_store_list
df_result['p_v'] = pv_list
df_result['I_avrg'] = I_tok_list
df_result['d'] = d_results
df_result['d\''] = d_results_kraj
df_result['d\'s'] = d_results_za_krajem

df_popis = df_clean['Fasada']
df_popis = df_popis[df_popis.index.isin(df_result.index)].reset_index(drop=True)
df_restult = df_result.reset_index(drop=True)

###############################################################################
''' POP tabulka data'''
df_POP_init = df_clean.loc[df_clean['Samostatne'].isin(['ANO', 'NE – Fasada', 'NE – Okno', 'Přístřešky pro auta DP3', 'Pergoly / přístřešky'])][['Fasada', 'Samostatne', 'Nazev/fasada', 'vyska', 'sirka']]
bolean = df_clean.Samostatne.isin(['NE – Fasada'])
bolean_list = bolean.values.tolist()

if True in bolean.values:
    df_POP_count = df_clean.index[df_POP_init['Samostatne'] == 'NE – Fasada'].tolist()
    POP_area = []
    for item in df_POP_count:
        item += 1
        POP_area_local = []
        while df_POP_init['Samostatne'][item] == 'NE – Okno':
            POP_area_local.append(df_POP_init['vyska'][item] * df_POP_init['sirka'][item])
            item += 1
            if item > (len(df_POP_init)-1):
                break
        POP_area.append(sum(POP_area_local))
        df_POP = df_clean.loc[df_clean['Samostatne'].isin(['NE – Fasada'])][['Samostatne', 'Nazev/fasada', 'vyska', 'sirka']]
    df_POP['Plocha_celk'] = df_POP['vyska'] * df_POP['sirka']
    df_POP['Plocha_POP'] = POP_area
    df_POP = df_POP[['Nazev/fasada', 'Plocha_celk', 'Plocha_POP']]
    df_POP['procento'] = round(df_POP['Plocha_POP'] / df_POP['Plocha_celk'] * 100, 2)
    decision_POP = []
    for i in df_POP['procento']:
        if i >= 40:
            decision_POP.append('NE')
        else:
            decision_POP.append('ANO')
    df_POP['jednotlive'] = decision_POP
else:
    pass

df_popis_POP = df_POP_init['Fasada']
if 'NE – Fasada' in df['Samostatne'].values:
    df_popis_POP = df_popis_POP[df_popis_POP.index.isin(df_POP.index)].reset_index(drop=True)

###############################################################################

geometry_options = {
    "margin": "0.5cm",
    "includeheadfoot": True
    }
doc = Document(page_numbers=True, geometry_options=geometry_options)
doc.preamble.append(NoEscape(r'\definecolor{Hex}{RGB}{239,239,239}'))
doc.preamble.append(NoEscape(r'\usepackage[czech]{babel}'))
doc.preamble.append(NoEscape(r'\usepackage{threeparttablex}'))
doc.documentclass.options = Options('10pt')

df_count_fasada = df.groupby('Fasada').nunique()['Nazev/fasada']
count_fasada = df_count_fasada.tolist()


def POP_ne_generator():
    with doc.create(Section('Vymezení požárně nebezpečného prostoru')):
        with doc.create(Subsection('Procenta požárně otevřených ploch')):
            doc.append(NoEscape(r'Každý okenní otvor je dále spočten jako jednotlivá 100 \% požárně otevřená plocha.'))


def POP_generator(df):
    data_used = df.values.tolist()
    for i in range(0, len(data_used)):
        data_used[i][1] = str(("%.2f" % data_used[i][1])).replace(".", ",")
        data_used[i][2] = str(("%.2f" % data_used[i][2])).replace(".", ",")
        data_used[i][3] = str(("%.2f" % data_used[i][3])).replace(".", ",")

    with doc.create(Section('Vymezení požárně nebezpečného prostoru')):
        with doc.create(Subsection('Procenta požárně otevřených ploch')):
            doc.append(NoEscape(r'Dále je určeno, zda je možné jednotlivé otvory ve fasádě posuzovat samostatně, a to výpočtem procenta požárně otevřených ploch z celkové plochy každé fasády. Způsob odečtení plochy fasády a plochy otevřených ploch je patrný z obrázku \ref{SchemaPOP}.'))
            with doc.create(Figure(position='htb!')) as POP:
                POP.add_image('images/POP.jpg', width='200px')
                POP.add_caption('Odečtení požárně otevřených ploch')
                POP.append(Command('label', 'SchemaPOP'))
            doc.append(NoEscape(r'Procentuální výsledky jednotlivých fasád s požárně otevřenými plochami jsou patrné z tabulky \ref{POP}.'))

        with doc.create(LongTable("lcccc", pos=['htb'])) as data_table:
            doc.append(Command('caption', 'Odstupové vzdálenosti od objektu'))
            doc.append(Command('label', 'POP'))
            doc.append(Command('\ '))
            data_table.append
            data_table.add_hline()
            data_table.add_row(["Popis", "Celková plocha", "Plocha POP",
                                "Procento POP", "POP jednotlivě"])
            data_table.add_row([" ", NoEscape('[m$^2$]'), NoEscape('[m$^2$]'),
                                NoEscape('[\%]'), NoEscape(r'\textless 40 \%')])
            data_table.add_hline()
            data_table.end_table_header()
            for i in range(0, len(data_used)):
                if i == 0:
                    # \multicolumn{9}{l}{\textbf{1. nadzemní podlaží - POP společné}}\\
                    data_table.add_row((MultiColumn(5, align='l', data=utils.bold(df_popis_POP[i])),))
                if i > 0 and df_popis_POP[i] != df_popis_POP[i-1]:
                    data_table.add_hline()
                    data_table.add_row((MultiColumn(5, align='l', data=utils.bold(df_popis_POP[i])),))
                if (i) % 2 == 0:
                    data_table.add_row(data_used[i], color="Hex")
                elif (i) % 2 != 0:
                    data_table.add_row(data_used[i])
            data_table.add_hline()

        doc.append(NoEscape(r'Ostatní v tabulce výše nezmíněné otvory na fasádách objektu jsou spočteny jako samostatné 100 \% otevřené požární plochy.'))

###############################################################################

###############################################################################


def odstup_generator(df):
    ''' Funkce k vytvoření kapitoly s vypočtenými odstupy

        Vstupní data musí být DataFrame o specifickém formátu o deseti
        sloupcích viz níže:

        df_result['Nazev/fasada', 'vyska', 'sirka', 'Samostatne']]
        df_result['POP'] = POP_store_list
        df_result['p_v'] = pv_list
        df_result['I_avrg'] = I_tok_list
        df_result['d'] = d_results
        df_result['d\''] = d_results_kraj
        df_result['d\'s'] = d_results_za_krajem

        Funkce tvoří přímo tabulku upravenou na míru pro prezentování
        odstupových vzdáleností v PBŘ zprávě
    '''

    '''Poznámky k tabulce'''
    list_avaiable = ['Přístřešky pro auta DP3', 'Pergoly / přístřešky']

    '''Práce s daty '''
    data_popis = df['Samostatne'].values.tolist()  # Data pro přidání popisků
    df = df.drop(columns=['Samostatne'])
    data_used = df.values.tolist()  # Data pro tvoření tabulky

    '''Přidání popisků na správné místo'''
    check = []
    if not data_popis:
        # Zjišťuje zda je nutné přidat popisky
        pass
    else:
        save_var = []
        save_id = []
        save_name = []
        save_id_check = []
        for i in range(0, len(data_popis)):
            if data_popis[i] in list_avaiable:
                check.append(i)
                save_var.append(data_popis[i])
                save_id_check.append(i)
        count = list(range(1, len(save_var)+1))
        pre = ' $^{'
        suf = ')}$'
        for i in range(0, len(save_var)):
            count[i] = pre + str(count[i])
            count[i] = str(count[i]) + suf
        for n in range(0, len(save_var)):
            if n == 0:
                new_dict = dict(zip([save_var[n]], [count[n]]))
                save_id.append(save_id_check[n])
                save_name.append(save_var[n])
                # new_dict = dict(zip([save_var[n]], str(count[n])))
            if n > 0:
                if save_var[n] in new_dict:
                    save_id.append(save_id_check[n])
                    save_name.append(save_var[n])
                else:
                    len_dict = len(new_dict)+1
                    len_dict = pre + str(len_dict)
                    len_dict = str(len_dict) + suf
                    save_id.append(save_id_check[n])
                    save_name.append(save_var[n])
                    new_dict[save_var[n]] = len_dict
        for i in range(0, len(save_id)):
            data_used[save_id[i]][0] = NoEscape(data_used[save_id[i]][0]
                                                + new_dict[save_name[i]])

    '''Přeměna čísel s . na čísla s ,'''
    for i in range(0, len(data_used)):
        data_used[i][1] = str(("%.2f" % data_used[i][1])).replace(".", ",")
        data_used[i][2] = str(("%.2f" % data_used[i][2])).replace(".", ",")
        data_used[i][3] = str(("%.2f" % data_used[i][3])).replace(".", ",")
        data_used[i][4] = str(("%.2f" % data_used[i][4])).replace(".", ",")
        data_used[i][5] = str(("%.2f" % data_used[i][5])).replace(".", ",")
        data_used[i][6] = str(("%.2f" % data_used[i][6])).replace(".", ",")
        data_used[i][7] = str(("%.2f" % data_used[i][7])).replace(".", ",")
        data_used[i][8] = str(("%.2f" % data_used[i][8])).replace(".", ",")

    '''Textová část kapitoly '''
    with doc.create(Subsection('Stanovení odstupových vzdáleností')):
        if kcni_system != 'nehořlavý':
            doc.append(NoEscape(r'Požární zatížení je v souladu s čl. 10.4.4 normy ČSN 73 0802 vzhledem k zatřídění konstrukčního systému nově navrhované budovy ({}; zatřídění viz tabulka '.format(kcni_system)))
            doc.append(NoEscape(r'\ref{PTCH}) '))
            doc.append(NoEscape(r' dále zvýšeno o {} kg/m$^2$.'.format(add_pv)))
            doc.append(NoEscape('\par'))
        doc.append(NoEscape(r'Odstupové vzdálenosti jsou stanoveny určením vzdáleností kritického tepelného toku 18,5 kW/m$^2$ od vnějšího líce požárně otevřené plochy, a to pomocí programu pro výpočet odstupových vzdáleností, který je zmíněn v kapitole \ref{cha-1}). Výsledné hodnoty odstupových vzdáleností z tohoto programu jsou názorně zakresleny v následujícím obrázku \ref{SchemaPOP2}.'))
        with doc.create(Figure(position='htb!')) as Odstup:
            Odstup.add_image('images/odstup.jpg', width='160px')
            Odstup.add_caption('Znázornění vypočtených hodnot')
            Odstup.append(Command('label', 'SchemaPOP2'))
        doc.append(NoEscape(r'Vypočtené odstupové vzdále\-nosti jsou patrné z tabulky \ref{odstup}.'))

        '''Ověření zda je možné přidat popisky + přidání popisků'''
        if len(check) > 0:
            doc.append(Command('begin', 'ThreePartTable'))
            doc.append(Command('begin', 'TableNotes'))
            doc.append(NoEscape('\small'))
            for key in new_dict.keys():
                pre_str = '\item '
                if key == 'Pergoly / přístřešky':
                    loc_str = ' Odstupová vzdálenost z domnělých stěn otevřených přístřešků je spočtena pomocí normové křivky pro vnější požár. Maximální teplota požáru je tedy v tomto případě 658$^o$C. Požární zatížení je stanoveno pomocí čl. 10.4.4 normy ČSN 73 0802.'
                    string = new_dict[key]
                    doc.append(NoEscape(pre_str + string + loc_str))
                if key == 'Přístřešky pro auta DP3':
                    loc_str = ' Odstupová vzdálenost od otevřeného přístřešku pro auta je stanovena na základě čl. I.3.1 normy ČSN 73 0804. Uvažováno je s ekvivalentní dobou požáru 30 minut (zde označeno jako p$_v$ = 30 kg/m$^2$). s výškou přístřešku 1,5 m a skutečnou délkou přístřešku.'
                    string = new_dict[key]
                    doc.append(NoEscape(pre_str + string + loc_str))

        if len(check) > 0:
            doc.append(Command('end', 'TableNotes'))
        with doc.create(LongTable("lcccccccc", pos=['htb'])) as data_table:
            doc.append(Command('caption', 'Odstupové vzdálenosti od objektu'))
            doc.append(Command('label', 'odstup'))
            doc.append(Command('\ '))
            data_table.append
            data_table.add_hline()
            data_table.add_row([" ", "Výška", "Šířka", "POP",
                                NoEscape('$p_v$'), NoEscape('$I_{avrg}$'), "d", "d\'", NoEscape("d\'$_s$")])
            data_table.add_row([" ", "[m]", "[m]", NoEscape('[\%]'), NoEscape('[kg/m$^2$]'),
                                NoEscape('[kW/m$^2$]'), "[m]", "[m]", "[m]"])
            data_table.add_hline()
            data_table.end_table_header()
            for i in range(0, len(data_used)):
                if i == 0:
                    # \multicolumn{9}{l}{\textbf{1. nadzemní podlaží - POP společné}}\\
                    data_table.add_row((MultiColumn(9, align='l', data=utils.bold(df_popis[i])),))
                if i > 0 and df_popis[i] != df_popis[i-1]:
                    data_table.add_hline()
                    data_table.add_row((MultiColumn(9, align='l', data=utils.bold(df_popis[i])),))
                if (i) % 2 == 0:
                    data_table.add_row(data_used[i], color="Hex")
                elif (i) % 2 != 0:
                    data_table.add_row(data_used[i])
            data_table.add_hline()
            if len(check) > 0:
                doc.append(NoEscape('\insertTableNotes'))
        if len(check) > 0:
            doc.append(Command('end', 'ThreePartTable'))


'''Volání funkcí'''
if df_input['Hodnota'][2] == 'ANO':
    if True in bolean.values:
        POP_generator(df_POP)
    else:
        POP_ne_generator()
if df_input['Hodnota'][3] == 'ANO':
    odstup_generator(df_result)
if df_input['Hodnota'][3] == 'ANO' or df_input['Hodnota'][2] == 'ANO':
    doc.generate_pdf("Odstupy", clean_tex=False)
if df_input['Hodnota'][4] == 'ANO':
    df_rotate = df_clean.loc[
                df_clean['Samostatne'].isin(
                                            ['ANO',
                                             'NE – Fasada',
                                             'Přístřešky pro auta DP3',
                                             'Pergoly / přístřešky'])][['Uhel']]
    tisk_odstup(df_result, df_input, df_rotate)

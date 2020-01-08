
from pylatex import Document, LongTable, NoEscape, Command, MultiColumn
from pylatex.base_classes.command import Options
from pathlib import PurePath
from read_results import read_results, chunks
import os


def appendix(data_dir, vystup_dir):
    os.chdir(data_dir)
    data_PU = read_results('raw_data_PU.csv')
    data_POP = read_results('raw_data_POP.csv')
    data_mid = read_results('results.csv')
    PurePath(vystup_dir)
    geometry_options = {
        "margin": "0.5cm",
        "includeheadfoot": True
        }
    doc = Document(page_numbers=True, geometry_options=geometry_options)
    doc.preamble.append(NoEscape(r'\definecolor{Hex}{RGB}{239,239,239}'))
    doc.documentclass.options = Options('10pt')

    doc.append(Command('textbf', 'Zadání hodnot pro výpočet: '))
    doc.append(NoEscape(r'Tabulka \ref{rozmery} je souhrnem zadaných hodnot \
                        potřebných k výpočtu požárně výpočtového zatížení.'))

    def genenerate_longtabu(data):
        if data == data_PU:
            with doc.create(LongTable("l l c c c c c c l")) as data_table:
                doc.append(Command('caption', 'Zadání místností pro výpočet'))
                doc.append(Command('label', 'rozmery'))
                doc.append(Command('\ '))
                data_table.append
                data_table.add_hline()
                data_table.add_row(["PU", "Místnost", "Plocha",
                                    NoEscape('h$_s$'), NoEscape('a$_n$'),
                                    NoEscape('p$_n$'), NoEscape('p$_s$'),
                                    "c", "Tab. A.1"])
                data_table.add_hline()
                data_table.end_table_header()
                data_table.add_hline()
                doc.append(Command('endfoot'))
                doc.append(Command('endlastfoot'))
                # data_table.end_table_footer()
                data_table.add_hline()
        if data == data_POP:
            with doc.create(LongTable("l l c c c c", pos=['l'])) as data_table:
                doc.append(Command('caption', 'Zadání okenních otvorů'))
                doc.append(Command('label', 'okna'))
                doc.append(Command('\ '))
                data_table.append
                data_table.add_hline()
                data_table.add_row(["PU", "Místnost", "n otvorů",
                                    "šířka otvorů", "výška otvorů", "Plocha"])
                data_table.add_row([" ", " ", " ", "[m]",
                                    "[m]", NoEscape('m$^2$')])
                data_table.add_hline()
                data_table.end_table_header()
                data_table.add_hline()
                data_table.add_hline()
        for i in range(0, len(data)):
            if i % 2 == 0:
                data_table.add_row(data[i])
            else:
                data_table.add_row(data[i], color="Hex")
        data_table.add_hline()

    for i in range(0, len(data_PU)):
        data_PU[i][2] = data_PU[i][2].replace(".", ",")
        data_PU[i][3] = data_PU[i][3].replace(".", ",")
        data_PU[i][4] = data_PU[i][4].replace(".", ",")
        data_PU[i][5] = data_PU[i][5].replace(".", ",")
        data_PU[i][6] = data_PU[i][6].replace(".", ",")
        data_PU[i][7] = data_PU[i][7].replace(".", ",")
    genenerate_longtabu(data_PU)
    doc.append(NoEscape(r'Okenní otvory nutné pro výpočet součinitele b jsou\
                        uvedeny v následující tabulce \ref{okna}'))
    for i in range(0, len(data_POP)):
        data_POP[i][3] = data_POP[i][3].replace(".", ",")
        data_POP[i][4] = data_POP[i][4].replace(".", ",")
        data_POP[i][5] = data_POP[i][5].replace(".", ",")
    genenerate_longtabu(data_POP)
    doc.append(Command('textbf', 'Mezivýsledky a výsledky: '))
    doc.append(NoEscape(r'Mezivýsledky nutné pro stanovení parametru b jsou\
                        patrné z tabulky \ref{mezivysledky}'))

    #########################################################################
    ''' Mezivýsledky parametr B '''
    data_mid_ar = []
    for i in range(0, len(data_mid)):
        data_mid_ar.append(data_mid[i][0])
        data_mid_ar.append(data_mid[i][1])
        data_mid_ar.append(data_mid[i][8])
        data_mid_ar.append(data_mid[i][9])
        data_mid_ar.append(data_mid[i][10])
        data_mid_ar.append(data_mid[i][14])
        data_mid_ar.append(data_mid[i][15])
    data_mid_ar = list(chunks(data_mid_ar, 7))
    for i in range(0, len(data_mid_ar)):
        data_mid_ar[i][1] = data_mid_ar[i][1].replace(".", ",")
        data_mid_ar[i][2] = data_mid_ar[i][2].replace(".", ",")
        data_mid_ar[i][3] = data_mid_ar[i][3].replace(".", ",")
        data_mid_ar[i][4] = data_mid_ar[i][4].replace(".", ",")
        data_mid_ar[i][5] = data_mid_ar[i][5].replace(".", ",")
        data_mid_ar[i][6] = data_mid_ar[i][6].replace(".", ",")

    with doc.create(LongTable("l c c c c c c", pos=['l'])) as data_table:
        doc.append(Command('caption', 'Mezivýsledky pro paramter b'))
        doc.append(Command('label', 'mezivysledky'))
        doc.append(Command('\ '))
        data_table.append
        data_table.add_hline()
        data_table.add_row(["PÚ", "S", "S0", "hs",
                            "h0", "n", "k"])
        data_table.add_row([" ", NoEscape('m$^2$'), NoEscape('m$^2$'), "[m]",
                            "[m]", "[-]", "[-]"])
        data_table.add_hline()
        data_table.end_table_header()
        data_table.add_hline()
        data_table.add_hline()
        for i in range(0, len(data_mid_ar)):
            if i % 2 == 0:
                data_table.add_row(data_mid_ar[i])
            else:
                data_table.add_row(data_mid_ar[i], color="Hex")
        data_table.add_hline()

    #########################################################################
    ''' Vysledky '''

    doc.append(NoEscape(r'Přehled výsledků požárních úseků je patrný z tabulky\
                        \ref{Vysledky}'))

    data_res = []
    for i in range(0, len(data_mid)):
        data_res.append(data_mid[i][0])
        data_res.append(data_mid[i][1])
        data_res.append(data_mid[i][2])
        data_res.append(data_mid[i][3])
        data_res.append(data_mid[i][4])
        data_res.append(data_mid[i][13])  # a_n
        data_res.append(data_mid[i][12])  # p_n
        data_res.append(data_mid[i][11])  # p_s
        data_res.append(data_mid[i][7])   # p_only
        data_res.append(data_mid[i][5])   # p_v
    data_res = list(chunks(data_res, 10))

    for i in range(0, len(data_res)):
        data_res[i][1] = data_res[i][1].replace(".", ",")
        data_res[i][2] = data_res[i][2].replace(".", ",")
        data_res[i][3] = data_res[i][3].replace(".", ",")
        data_res[i][4] = data_res[i][4].replace(".", ",")
        data_res[i][5] = data_res[i][5].replace(".", ",")
        data_res[i][6] = data_res[i][6].replace(".", ",")
        data_res[i][7] = data_res[i][7].replace(".", ",")
        data_res[i][8] = data_res[i][8].replace(".", ",")
        data_res[i][9] = data_res[i][9].replace(".", ",")
    with doc.create(LongTable("l c c c c c c c c c ", pos=['l'])) as data_table:
        doc.append(Command('caption', 'Přehled požárních úseků a jejich\
                            výsledky'))
        doc.append(Command('label', 'Vysledky'))
        doc.append(Command('\ '))
        data_table.append
        data_table.add_hline()
        data_table.add_row(["PÚ", "Plocha", "a", "b", "c", NoEscape('a$_n$'),
                            NoEscape('p$_n$'), NoEscape('p$_s$'),
                            NoEscape('p'), NoEscape('p$_v$')])
        data_table.add_row([" ", NoEscape('m$^2$'), "[-]",
                            "[-]", "[-]",
                            "[-]", NoEscape('[kg/m$^2$]'),
                            NoEscape('[kg/m$^2$]'), NoEscape('[kg/m$^2$]'),
                            NoEscape('[kg/m$^2$]')])

        data_table.end_table_header()
        data_table.add_hline()
        for i in range(0, len(data_res)):
            if i % 2 == 0:
                data_table.add_row(data_res[i])
            else:
                data_table.add_row(data_res[i], color="Hex")
        data_table.add_hline()
    os.chdir(vystup_dir)
    doc.generate_pdf("Appendix", clean_tex=False)

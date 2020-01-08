import os
from read_results import read_results, chunks
from pylatex import Document, LongTable, NoEscape, Command, Subsection
from pylatex.base_classes.command import Options


def d1_generator(data_dir, vystup_dir):
    os.chdir(data_dir)
    data_PU = read_results('results.csv')
    info_PU = read_results('raw_data_info.csv')
    list_names = []
    for i in range(0, len(data_PU)):
        list_names.append(data_PU[i][19])
    # Arrange data to desired shape
    data_used = []
    forbiden = ['CHÚC-A', 'CHÚC-B', 'CHÚC-C', 'OB2 byt']
    for i in range(0, len(data_PU)):
        if list_names[i] not in forbiden:
            data_used.append(data_PU[i][0])
            data_used.append(data_PU[i][2])
            data_used.append("%.2f" % float(data_PU[i][17]))
            A = float(data_PU[i][17])
            data_used.append("%.2f" % float(data_PU[i][18]))
            B = float(data_PU[i][18])
            C = A * B
            data_used.append("%.2f" % C)
            data_used.append(data_PU[i][1])
            if (A * B) > float(data_PU[i][1]):
                data_used.append('ANO')
            else:
                data_used.append('NE')
                print()
                print('D.1 Pozor nevyhovuje - ' + str(data_PU[i][0]))
                print()

    data_used = list(chunks(data_used, 7))

    geometry_options = {
        "margin": "0.5cm",
        "includeheadfoot": True
        }
    doc = Document(page_numbers=True, geometry_options=geometry_options)
    doc.preamble.append(NoEscape(r'\usepackage[czech]{babel}'))
    doc.preamble.append(NoEscape(r'\usepackage{threeparttablex}'))
    doc.preamble.append(NoEscape(r'\usepackage{pdfpages}'))
    doc.preamble.append(NoEscape(r'\definecolor{Hex}{RGB}{239,239,239}'))
    doc.documentclass.options = Options('10pt')

    with doc.create(Subsection('Mezní rozměry požárních úseků')):
        doc.append(NoEscape(r'\textbf{Mezní půdorysné rozměry}: Posouzení\
                            rozměrů požárního úseku je patrné z tabulky\
                            \ref{rozmery}.'))
        if 'OB2 byt' in list_names:
            doc.append('Mezní rozměry bytových jednotek se v souladu\
                        s čl. 5.1.5 normy ČSN 73 0833 nestanovují.')

        if info_PU[0] == ['nehořlavý']:
            str1 = str(9)
        if info_PU[0] == ['smíšený']:
            str1 = str(10)
        if info_PU[0] == ['hořlavý']:
            str1 = str(11)
        string1 = 'Mezní délky a mezní šířky jednotlivých požárních\
                   úseků byly stanoveny tabulkou '
        string2 = ' normy ČSN 73 0802 na základě součinitele rychlosti\
                    odhořívání požárního úseků.'
        string = string1 + str1 + string2
        doc.append(string)
        for i in range(0, len(data_used)):
            data_used[i][1] = data_used[i][1].replace(".", ",")
            data_used[i][2] = data_used[i][2].replace(".", ",")
            data_used[i][3] = data_used[i][3].replace(".", ",")
            data_used[i][4] = data_used[i][4].replace(".", ",")
            data_used[i][5] = data_used[i][5].replace(".", ",")

        with doc.create(LongTable("l c c c c c c", pos=['htb!'])) as data_table:
            doc.append(Command('caption', 'Ověření rozměrů požárních úseků'))
            doc.append(Command('label', 'rozmery'))
            doc.append(Command('\ '))
            data_table.append
            data_table.add_hline()
            data_table.add_row(["PÚ", "Souč.", "Mězní délka",
                                "Mezní šířka", "Mezní plocha", "Plocha",
                                "Vyhovuje"])
            data_table.add_row([" ", "a", "[m]", "[m]",
                                NoEscape('[m$^2$]'), NoEscape('[m$^2$]'), ""])
            data_table.add_hline()
            data_table.end_table_header()
            for i in range(0, len(data_used)):
                if i % 2 != 0:
                    data_table.add_row(data_used[i], color="Hex")
                else:
                    data_table.add_row(data_used[i])
            data_table.add_hline()
            os.chdir(vystup_dir)
    doc.generate_pdf("D1_PU", clean_tex=False)

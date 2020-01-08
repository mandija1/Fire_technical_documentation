import os
import sys
import math
from read_results import read_results, chunks
from pylatex import Document, LongTable, NoEscape, Command, MultiColumn, Section, Subsection
from pylatex.base_classes.command import Options

geometry_options = {
    "margin": "0.5cm",
    "includeheadfoot": True
    }
doc = Document(page_numbers=True, geometry_options=geometry_options)
doc.preamble.append(NoEscape(r'\usepackage[czech]{babel}'))
doc.preamble.append(NoEscape(r'\usepackage{threeparttablex}'))
doc.preamble.append(NoEscape(r'\usepackage{pdfpages}'))
doc.preamble.append(NoEscape(r'\definecolor{Hex}{RGB}{239,239,239}'))
doc.preamble.append(NoEscape(r'\usepackage{threeparttablex}'))
doc.documentclass.options = Options('10pt')

def k_generator(data_dir, vystup_dir):
    os.chdir(data_dir)
    data_PU = read_results('results.csv')

    data_used = []
    for i in range(0, len(data_PU)):
        data_used.append(data_PU[i][0])
        data_used.append(data_PU[i][1])
        data_used.append(data_PU[i][2])
        data_used.append(data_PU[i][4])
        a = 0.15*((float(data_PU[i][1])*float(data_PU[i][2])*float(data_PU[i][4])) ** 0.5)
        # print(float(data_PU[i][1]), float(data_PU[i][2]), float(data_PU[i][4]))
        data_used.append(str(round(a, 2)))
        data_used.append('{} PG6 21A/183B'. format(math.ceil(a)))

    data_used = list(chunks(data_used, 6))

    with doc.create(Section('Hasicí přístroje')):
        with doc.create(Subsection('Přehled hasicích přístrojů')):
            doc.append(NoEscape(r'Přehled počtu a druhu všech hasicích přístrojů,\
                                které budou v objektu osazeny, je patrný z\
                                tabulky \ref{PHP_stanoveni}.'))

        for i in range(0, len(data_used)):
            data_used[i][1] = data_used[i][1].replace(".", ",")
            data_used[i][2] = data_used[i][2].replace(".", ",")
            data_used[i][3] = data_used[i][3].replace(".", ",")
            data_used[i][4] = data_used[i][4].replace(".", ",")

        with doc.create(LongTable("l c c c c l", pos=['htb'])) as data_table:
            doc.append(Command('caption', 'Počet a druh hasicích přístrojů'))
            doc.append(Command('label', NoEscape('PHP_stanoveni')))
            doc.append(Command('\ '))
            data_table.append
            data_table.add_hline()
            data_table.add_row(["Požární úsek", "S", "a", "c",
                                NoEscape('$n_r$'), "Počet PHP - typ"])
            data_table.add_row([" ", NoEscape('[m$^2$]'), "-", "-",
                                "-", " "])
            data_table.add_hline()
            data_table.end_table_header()
            for i in range(0, len(data_used)):
                if i % 2 != 0:
                    data_table.add_row(data_used[i], color="Hex")
                else:
                    data_table.add_row(data_used[i])
            data_table.add_hline()
            # doc.append(NoEscape('\insertTableNotes'))
            os.chdir(vystup_dir)
    doc.generate_pdf("K_hasicaky", clean_tex=False)

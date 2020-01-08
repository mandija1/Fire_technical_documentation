import os
from read_results import read_results
from pylatex import Document, NoEscape, Command, Section, Itemize
from pylatex.base_classes.command import Options


def c_generator(data_dir, vystup_dir):
    os.chdir(data_dir)
    data_PU = read_results('results.csv')
    info_PU = read_results('raw_data_info.csv')
    list_names = []
    for i in range(0, len(data_PU)):
        list_names.append(data_PU[i][19])

    geometry_options = {
        "margin": "0.5cm",
        "includeheadfoot": True
        }
    doc = Document(page_numbers=True, geometry_options=geometry_options)
    doc.documentclass.options = Options('10pt')
    with doc.create(Section('Mezní rozměry požárních úseků')):
        doc.append(NoEscape(r'\textbf{Rozdělení do požárních úseků}: Objekt\
                            bude rozdělen do požárních úseků podle požadavků\
                            norem. V souladu s požadavky budou požární úseky\
                            tvořit především tyto provozy/prostory.'))
        with doc.create(Itemize()) as itemize:
            if 'OB2 byt' in list_names:
                itemize.add_item("Dle normy ČSN 73 0833 čl. 3.6 musí každá\
                                 bytová buňka v objektech skupiny OB2 tvořit\
                                 samostatný požární úsek.")
            if '15.10 c)' in list_names:
                itemize.add_item("Kotelna III. kategorie podle normy\
                                  ČSN 07 0703 bude tvořit samostatný požární\
                                  úsek")
            if 'CHÚC-A' in list_names or 'CHÚC-B' in list_names or\
               'CHÚC-C' in list_names:
                itemize.add_item("Chráněné únikové cesty tvoří samostatné\
                                  požární úseky")
            if ['ANO'] == info_PU[3]:
                itemize.add_item("Výtahové šachty budou tvořit samostatné\
                                  požární useky.")
            if ['ANO'] == info_PU[4]:
                itemize.add_item("Instalační šachty budou tvořit samostatné\
                                  požární úseky.")
            if 'OB2 sklad' in list_names:
                itemize.add_item("Prostory určené pro skladování věcí pro\
                                  domácnost budou tvořit samostatné požární\
                                  úseky.")
            if 'OB3' in list_names:
                itemize.add_item("Každá ubytovací buňka tvoří samostatný\
                                  požární úsek.")
            if 'OB4 ubyt' in list_names:
                itemize.add_item("Ubytovací (lůžkové) části budovy skupiny\
                                 OB4 budou tvořit samostatné požární úseky.")
            if 'OB4 sklad' in list_names:
                itemize.add_item("Prostory určené pro skladování věcí různých\
                                  potřeb pro provoz ubytovacích částí budovy\
                                  budou tvořit samostatné požární úsek.y")
            if 'AZ1 Ordi.' in list_names or 'AZ2 Ordi' in list_names:
                itemize.add_item("Ordinace lékaře bude tvořit samostatný\
                                  požární úsek.")
            if 'AZ1 Lék.' in list_names or 'AZ2 Lék' in list_names or\
               'LZ2 Lék' in list_names:
                itemize.add_item("Lékárna včetně zázemí bude tvořit samostatný\
                                  požární úsek.")
            if 'AZ2 vyšet.' in list_names:
                itemize.add_item("Vyšetřovací a léčebné části objektu budou\
                                  tvořit samostatný požární úsek.")
            if 'LZ1' in list_names or 'LZ2 lůž' in list_names:
                itemize.add_item("Lůžkové části pro pacienty v objektu budou\
                                  tvořit samostatné požární úseky")
            if 'LZ2 int.péče' in list_names:
                itemize.add_item("Jednotky intenzivní péče, anesteziologicko\
                                  resistutační oddělení a operační sály budou\
                                  tvořit samostatný požární úseky.")
            if 'LZ2 biochem' in list_names:
                itemize.add_item("Oddělení klinické biochemie bude tvořit\
                                  samostatný požární úsek.")
            if 'peč. Služ' in list_names:
                itemize.add_item("Každá bytová jednotka v domě s pečovatelskou\
                                  službou bude tvořit samostatný požární\
                                  úsek.")
            if 'soc.péče.ošetř' in list_names:
                itemize.add_item("Ošetřovny v domovech sodiální poče budou\
                                  tvořit samostatné požární úsek.")
            if 'soc.péče.lůž' in list_names:
                itemize.add_item("Lůžkové prostory v domovech sociální péče\
                                  budou tvořit samostatné požární úseky.")
            if 'soc.péče.byt' in list_names:
                itemize.add_item("Každá bytová jednotka v domě sociální péče\
                                  bude tvořit samostatný požární úsek.")
            if 'Jesle' in list_names:
                itemize.add_item("Prostory pro děti do tří let (Jesle)\
                                  budou tvořit samostatný požární úsek.")
    os.chdir(vystup_dir)
    doc.generate_pdf("C", clean_tex=False)

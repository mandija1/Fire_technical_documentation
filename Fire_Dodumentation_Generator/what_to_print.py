import sys
sys.path.insert(0, '../Fire_Dodumentation_Generator/table_generator/')
from appendix_generator import appendix
from chunks import chunks
from D_generator import d_generator
from D1_generator import d1_generator
from C_generator import c_generator
from E1_generator import E1_generator
from k_generator import k_generator
sys.path.insert(0, '../minor/')
from open_konst import open_konst
from odstupy import odstup
from konst_data_prep import konst_data_prep
import openpyxl


def what_to_print(name, Cesta, konst, soubor, wb):
    dire = 'data'
    data_dir = Cesta
    for i in range(0, len(dire)):
        data_dir += dire[i]
    dire_v = 'vystupy'
    vystup_dir = Cesta
    for i in range(0, len(dire_v)):
        vystup_dir += dire_v[i]
    if name == 'Z':
        appendix(data_dir, vystup_dir)
    if name == 'D':
        d_generator(data_dir, vystup_dir)
    if name == 'D.1':
        d1_generator(data_dir, vystup_dir)
    if name == 'C':
        c_generator(data_dir, vystup_dir)
    if name == 'C':
        c_generator(data_dir, vystup_dir)
    Konstrukce = ['E', 'E.1', 'E.2', 'E.3', 'E.4', 'E.5', 'E.6', 'E.7', 'E.8',
                  'E.9', 'E.10', 'E.11', 'E.1']
    if name in Konstrukce:
        data_konst = open_konst(konst, vystup_dir)
        if name == 'E.1':
            E1_generator(soubor, Cesta, data_konst, vystup_dir, data_dir)
    if name == 'H.1':
        odstup(wb, data_dir)
    if name == 'K':
        k_generator(data_dir, vystup_dir)

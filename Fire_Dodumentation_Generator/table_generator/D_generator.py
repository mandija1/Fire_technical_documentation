import os
import sys
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

doc.documentclass.options = Options('10pt')


def d_generator(data_dir, vystup_dir):
    os.chdir(data_dir)
    data_PU = read_results('results.csv')
    info_PU = read_results('raw_data_info.csv')


    # Arrange data to desired shape
    data_check = []
    data_used = []
    data_replace_p = []
    data_check_p = []
    data_replace_a = []
    l_names = []
    p_stale = []
    for i in range(0, len(data_PU)):
        data_used.append(data_PU[i][0])
        data_used.append(data_PU[i][16])
        data_used.append(data_PU[i][1])
        data_used.append(data_PU[i][2])
        data_used.append(data_PU[i][3])
        data_used.append(data_PU[i][4])
        data_used.append(data_PU[i][5])
        data_used.append(data_PU[i][6])
        data_check.append(data_PU[i][19])
        data_replace_p.append(data_PU[i][12])
        data_replace_a.append(data_PU[i][13])
        data_check_p.append(data_PU[i][7])
        l_names.append(data_PU[i][19])
        p_stale.append(data_PU[i][20])

    data_used = list(chunks(data_used, 8))
    list_avaiable = ['B.1 pol.1', 'B.1 pol.2', 'B.1 pol.3',
                     'B.1 pol.4', 'B.1 pol.5', 'B.1 pol.6',
                     'B.1 pol.7', 'B.1 pol.8', 'B.1 pol.9',
                     'B.1 pol.10', 'B.1 pol.11', 'B.1 pol.12',
                     'B.1 pol.13']
    list_avaiable2 = ['AZ1 Ordi.', 'AZ1 Lék.', 'AZ2 Ordi', 'AZ2 vyšet.',
                      'AZ2 Lék.', 'LZ1', 'LZ2 lůž', 'LZ2 int.péče', 'LZ2 Lék',
                      'LZ2 biochem', 'peč. Služ', 'soc.péče.ošetř.',
                      'soc.péče.lůž.', 'soc.péče.byt', 'Jesle']
    check = []
    for i in range(0, len(data_check)):
        if data_check[i] in list_avaiable or\
           data_check[i] == 'OB2 byt' or\
           data_check[i] == 'OB3' or data_check[i] == 'OB4 ubyt.' or data_check[i] == 'OB4 sklad' or\
           data_check[i] == 'CHÚC-A' or data_check[i] == 'CHÚC-B' or\
           data_check[i] == 'CHÚC-C' or data_check[i] in list_avaiable2:
            check.append(i)
    sys.path.insert(0,
                    "c:/Users/Honza/Google Drive/Work/Generator_zprav/minor/")
    from stupen import spb_def
    type_sys = info_PU[0]
    h_p = float(info_PU[1][0])
    podlazi = float(info_PU[2][0])
    if len(check) > 0:
        for item in check:
            data_used[item][4] = '-'
            '''data_used[item][4] = '-'
            if data_check[item] == 'OB2 byt':
                if (float(data_check_p[item]) - float(data_replace_p[item])) >= 0:
                    data_used[item][6] = '%.2f' % 40
                    spb_fix = spb_def(h_p, type_sys[0], [40], [1.00], podlazi, str(data_PU[item][0]))
                    data_used[item][7] = spb_fix[0]
                if (float(data_check_p[item]) - float(data_replace_p[item])) > 5:
                    data_used[item][6] = '%.2f' % 45.00
                    spb_fix = spb_def(h_p, type_sys[0], [45], [1.00], podlazi, data_PU[item][0])
                    data_used[item][7] = spb_fix[0]
                if (float(data_check_p[item]) - float(data_replace_p[item])) >= 15:
                    data_used[item][6] = '%.2f' % 50.00
                    spb_fix = spb_def(h_p, type_sys[0], [50], [1.00], podlazi, data_PU[item][0])
                    data_used[item][7] = spb_fix[0]
                data_used[item][3] = str('%.2f' % 1.00)
                # pozn = ' $^{1)}$'
                # data_used[item][0] = NoEscape(data_used[item][0] + pozn)
            if data_check[item] in list_avaiable:
                if 0 <= float(p_stale[item]) <= 5:
                    data_used[item][6] = '%.2f' % float(data_replace_p[item])
                if float(p_stale[item]) > 5:
                    data_replace_p[item] = ((float(p_stale[item]) - 5) * 1.15)\
                                            + float(data_replace_p[item])
                    data_used[item][6] = '%.2f' % data_replace_p[item]
                spb_fix = spb_def(h_p, type_sys[0],
                                  [float(data_replace_p[item])],
                                  [1.00], podlazi,
                                  data_PU[item][0])
                data_used[item][7] = spb_fix[0]
                data_used[item][3] = str('%.2f' % 1.00)
            if data_check[item] in list_avaiable2:
                data_used[item][6] = '%.2f' % float(data_replace_p[item])
                data_used[item][3] = data_replace_a[item]
                spb_fix = spb_def(h_p, type_sys[0], [float(data_used[item][6])], [float(data_used[item][3])], podlazi, data_PU[item][0])
                data_used[item][7] = spb_fix[0]
            if data_check[item] == 'OB3' or data_check[i] == 'OB4 ubyt.':
                data_used[item][6] = '%.2f' % float(data_replace_p[item])
                spb_fix = spb_def(h_p, type_sys[0],
                                  [30],
                                  [1.00], podlazi,
                                  data_PU[item][0])
                data_used[item][7] = spb_fix[0]
                data_used[item][3] = str('%.2f' % 1.00)
            if data_check[item] == 'OB4 sklad':
                data_used[item][6] = '%.2f' % float(data_replace_p[item])
                spb_fix = spb_def(h_p, type_sys[0],
                                  [60],
                                  [1.00], podlazi,
                                  data_PU[item][0])
                data_used[item][7] = spb_fix[0]
                data_used[item][3] = str('%.2f' % 1.00)'''
            if data_check[item] == 'CHÚC-A' or data_check[item] == 'CHÚC-B' or\
               data_check[item] == 'CHÚC-C':
                data_used[item][3] = '-'
                data_used[item][4] = '-'
                data_used[item][5] = '-'
                data_used[item][6] = '-'
                if h_p < 30:
                    data_used[item][7] = 'II.'
                if h_p >= 30:
                    data_used[item][7] = 'III.'
                if h_p >= 45:
                    data_used[item][7] = 'IV.'

    with doc.create(Section('Posouzení požárních úseků')):
        with doc.create(Subsection('Vyhodnocení požárních úseků')):
            doc.append(NoEscape(r'Z tabulky \ref{PU} je patrné požární riziko\
                                a stupeň požární bezpečnosti všech řešených\
                                požárních úseků. Není-li v poznámce pod\
                                tabulkou uvedeno jinak, je požární riziko PÚ\
                                stanoveno výpočtem (viz příloha \ref{A-1}:\
                                Výpočet požárního rizika).'))
            save_var = []
            save_id = []
            save_name = []
            save_id_check = []
            for i in range(0, len(data_check)):
                if data_check[i] == 'B.1 pol.1' or data_check[i] == 'B.1 pol.2' or\
                   data_check[i] == 'B.1 pol.3' or data_check[i] == 'B.1 pol.4' or\
                   data_check[i] == 'B.1 pol.5' or data_check[i] == 'B.1 pol.6' or\
                   data_check[i] == 'B.1 pol.7' or data_check[i] == 'B.1 pol.8' or\
                   data_check[i] == 'B.1 pol.9' or data_check[i] == 'B.1 pol.10' or\
                   data_check[i] == 'B.1 pol.11' or data_check[i] == 'B.1 pol.12' or\
                   data_check[i] == 'B.1 pol.13' or data_check[i] == 'OB2 byt' or\
                   data_check[i] == 'CHÚC-A' or data_check[i] == 'CHÚC-B' or\
                   data_check[i] == 'CHÚC-C' or data_check[i] == 'OB3' or\
                   data_check[i] == 'OB4 ubyt.' or\
                   data_check[i] == 'OB4 sklad' or\
                   data_check[i] in list_avaiable2:
                    save_var.append(data_check[i])
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
                data_used[save_id[i]][1] = NoEscape(data_used[save_id[i]][1]
                                                    + new_dict[save_name[i]])

            doc.append(Command('begin', 'ThreePartTable'))
            if len(check) > 0:
                doc.append(Command('begin', 'TableNotes'))
                doc.append(NoEscape('\small'))
                for key in new_dict.keys():
                    pre_str = '\item'
                    if key == 'OB2 byt':
                        idx = l_names.index(key)
                        suf_str = ' Uvažováno s požárně výpočetním zatížením (p$_v$)\
                                 podle normy ČSN 73 0833 čl. 5.1.2.'
                        string = new_dict[key]
                        if 0 <= (float(data_check_p[idx]) -
                           float(data_replace_p[idx])) < 5:
                            loc_str = pre_str + string + suf_str
                            doc.append(NoEscape(loc_str))
                        if (float(data_check_p[idx]) -
                           float(data_replace_p[idx])) > 5:
                            suf_str = suf_str + ' Stálé požární\
                            zatížení v požárním úseku je p$_s$ = 10 kg/m$^2$.\
                            Je tak přihlédnuto k poznámce stejného článku,\
                            která stanovuje požární zatížení na hodnotu\
                            p$_v$ = 45 kg/m$^2$.'
                            loc_str = pre_str + string + suf_str
                            doc.append(NoEscape(loc_str))

                    if key in list_avaiable:
                        string = new_dict[key]
                        idx = l_names.index(key)
                        polozka = str(list_avaiable.index(key) + 1)
                        suf_str_f = ' Hodnota požárně výpočtového zatížení je\
                                    stanovena paušálně z položky '
                        suf_str_ff = ' tabulky B.1 normy ČSN 73 0802.'
                        suf_str_b = ' Stálé požární zatížení je p$_s$ = '
                        suf_str_bb = ' kg/m$^2$. Při stanovení požárně výpočtového\
                                    zaížení je tak přihlédnuto k čl. B.1.2.'
                        if 0 <= (float(data_check_p[idx]) -
                           float(data_replace_p[idx])) <= 5:
                            suf_str_f = suf_str_f + polozka + suf_str_ff
                            loc_str = pre_str + string + suf_str_f
                            doc.append(NoEscape(loc_str))
                        if 5 < (float(data_check_p[idx]) -
                           float(data_replace_p[idx])):
                            suf_str_num = str(p_stale[idx])
                            suf_str_f = suf_str_f + polozka + suf_str_ff
                            suf_str_b = suf_str_b + suf_str_num + suf_str_bb
                            loc_str = pre_str + string + suf_str_f + suf_str_b
                            doc.append(NoEscape(loc_str))

                    if key in list_avaiable2:
                        string = new_dict[key]
                        idx = l_names.index(key)
                        suf_str_f = ' Hodnota požárně výpočtového zatížení je\
                                    stanovena paušálně z '
                        suf_str_ff = ' z normy ČSN 73 0835.'
                        if key == 'AZ1 Ordi.':
                            odkaz = 'čl. 5.3.1'
                            suf_str_b = ' Jedná se o zařízení lékařské péče\
                                     zařazené do skupiny AZ1'
                            loc_str = pre_str + string + suf_str_f + odkaz +\
                                      suf_str_ff + suf_str_b
                            doc.append(NoEscape(loc_str))
                        if key == 'AZ1 Lék.':
                            odkaz = 'čl. 5.3.1'
                            suf_str_b = ' Jedná se o lékárenské zařízení\
                                     zařazené do skupiny AZ1'
                            loc_str = pre_str + string + suf_str_f + odkaz +\
                                      suf_str_ff + suf_str_b
                            doc.append(NoEscape(loc_str))
                        if key == 'AZ2 Ordi':
                            odkaz = 'čl. 6.2.1'
                            suf_str_b = ' Jedná se o lékárenské pracoviště\
                                     zařazené do skupiny AZ2'
                            loc_str = pre_str + string + suf_str_f + odkaz +\
                                      suf_str_ff + suf_str_b
                            doc.append(NoEscape(loc_str))
                        if key == 'AZ2 vyšet.':
                            odkaz = 'čl. 6.2.1'
                            suf_str_b = ' Jedná se o vyšetřovací nebo léčebnou\
                                     část budovy zařazené do skupiny AZ2'
                            loc_str = pre_str + string + suf_str_f + odkaz +\
                                      suf_str_ff + suf_str_b
                            doc.append(NoEscape(loc_str))
                        if key == 'AZ2 Lék.':
                            odkaz = 'čl. 6.2.1'
                            suf_str_b = ' Jedná se o lékárenské zařízení\
                                        zařazené do skupiny AZ2'
                            loc_str = pre_str + string + suf_str_f + odkaz +\
                                      suf_str_ff + suf_str_b
                            doc.append(NoEscape(loc_str))
                        if key == 'LZ1':
                            odkaz = 'čl. 7.2.1'
                            suf_str_b = ' Jedná se o požární úsek, který je\
                                        součástí budovy skupiny LZ1'
                            loc_str = pre_str + string + suf_str_f + odkaz +\
                                      suf_str_ff + suf_str_b
                            doc.append(No.Escape(loc_str))
                        if key == 'LZ2 lůž':
                            odkaz = 'čl. 8.2.1'
                            suf_str_b = ' Jedná se o lůžkové jednotky v\
                                        budově skupiny LZ2'
                            loc_str = pre_str + string + suf_str_f + odkaz +\
                                      suf_str_ff + suf_str_b
                            doc.append(NoEscape(loc_str))
                        if key == 'LZ2 int.péče':
                            odkaz = 'čl. 8.2.1'
                            suf_str_b = ' Jedná se o jednotky intenzivní péče,\
                                        ansteziologicko resustitační oddělení,\
                                        nebo o operační oddělení v budově\
                                        zařazené do skupiny LZ2'
                            loc_str = pre_str + string + suf_str_f + odkaz +\
                                      suf_str_ff + suf_str_b
                            doc.append(NoEscape(loc_str))
                        if key == 'LZ2 Lék':
                            odkaz = 'čl. 8.2.1'
                            suf_str_b = ' Jedná se o lékárenské zařízení v\
                                        budově skupiny LZ2'
                            loc_str = pre_str + string + suf_str_f + odkaz +\
                                      suf_str_ff + suf_str_b
                            doc.append(NoEscape(loc_str))
                        if key == 'LZ2 biochem':
                            odkaz = 'čl. 8.2.1'
                            suf_str_b = ' Jedná se o oddělení klinické biochemie\
                                        v budově skupiny LZ2'
                            loc_str = pre_str + string + suf_str_f + odkaz +\
                                      suf_str_ff + suf_str_b
                            doc.append(NoEscape(loc_str))
                        if key == 'peč. Služ':
                            odkaz = 'čl. 9.3.1'
                            suf_str_b = ' Jedná se o bytovou jednotku v domě\
                                        s pečovatelskou službou'
                            loc_str = pre_str + string + suf_str_f + odkaz +\
                                      suf_str_ff + suf_str_b
                            doc.append(NoEscape(loc_str))
                        if key == 'soc.péče.oštř.':
                            odkaz = 'čl. 10.3.1'
                            suf_str_b = ' Jedná se o ošetřovatelské oddělení v\
                                        budově sociální péče'
                            loc_str = pre_str + string + suf_str_f + odkaz +\
                                      suf_str_ff + suf_str_b
                            doc.append(NoEscape(loc_str))
                        if key == 'soc.péče.lůž.':
                            odkaz = 'čl. 10.3.1'
                            suf_str_b = ' Jedná se o lůžkovou část ústavu\
                                        sociální péče'
                            loc_str = pre_str + string + suf_str_f + odkaz +\
                                      suf_str_ff + suf_str_b
                            doc.append(NoEscape(loc_str))
                        if key == 'soc.péče.byt.':
                            odkaz = 'čl. 10.3.1'
                            suf_str_b = ' Jedná se o bytové jednotky v budově\
                                        sociální péče.'
                            loc_str = pre_str + string + suf_str_f + odkaz +\
                                      suf_str_ff + suf_str_b
                            doc.append(NoEscape(loc_str))
                        if key == 'Jesle.':
                            odkaz = 'čl. 10.3.1'
                            suf_str_b = ' Jedná se o zrdavotnické zařízení\
                                        pro děti - jesle'
                            loc_str = pre_str + string + suf_str_f + odkaz +\
                                      suf_str_ff + suf_str_b
                            doc.append(NoEscape(loc_str))
                    if key == 'OB3':
                        string = new_dict[key]
                        pre_str = '\item'
                        suf_str = ' Požární riziko  pro ubytovací jednotku bylo\
                                    stanoveno paušálně pomocí čl. 6.1.1 normy\
                                    ČSN 73 0833.'
                        loc_str = pre_str + string + suf_str
                        doc.append(NoEscape(loc_str))

                    if key == 'OB4 ubyt.':
                        string = new_dict[key]
                        pre_str = '\item'
                        suf_str = ' Požární riziko pro ubytovací jednotku v budově OB4 bylo\
                                    stanoveno paušálně pomocí čl. 7.1.1 normy\
                                    ČSN 73 0833.'
                        loc_str = pre_str + string + suf_str
                        doc.append(NoEscape(loc_str))

                    if key == 'OB4 sklad':
                        string = new_dict[key]
                        pre_str = '\item'
                        suf_str = ' Požární riziko pro ubytovací jednotku v budově OB4 bylo\
                                    stanoveno paušálně pomocí čl. 7.1.3 normy\
                                    ČSN 73 0833.'
                        loc_str = pre_str + string + suf_str
                        doc.append(NoEscape(loc_str))

                    if key == 'CHÚC-A' or key == 'CHÚC-B' or key == 'CHÚC-C':
                        string = new_dict[key]
                        pre_str = '\item'
                        suf_str = ' CHÚC je zatříděna v souladu s čl. 9.3.2 normy\
                                    ČSN 73 0802.'
                        loc_str = pre_str + string + suf_str
                        doc.append(NoEscape(loc_str))
            if info_PU[3] == ['ANO']:
                try:
                    new_dict
                except NameError:
                    var_exist = False
                else:
                    var_exist = True
                if var_exist is False:
                    new_dict = []
                    pre_str = '\item'
                    doc.append(Command('begin', 'TableNotes'))
                    doc.append(NoEscape('\small'))
                if info_PU[5] == ['osobní výtahy, malé nákladní výtahy']:
                    if h_p <= 22.5:
                        s_vytah = ['Šv', 'Výtahové šachty', '-', '-', '-',
                                   '-', '-', 'II.']
                    if 22.5 < h_p <= 45.0:
                        s_vytah = ['Šv', 'Výtahové šachty', '-', '-', '-',
                                   '-', '-', 'III.']
                    if h_p > 45.0:
                        s_vytah = ['Šv', 'Výtahové šachty', '-', '-', '-',
                                   '-', '-', 'IV.']
                    len_vytah = len(new_dict) + 1
                    app_vytah = pre + str(len_vytah) + suf
                    suf_str_v = ' Výtahová šachta odpovídá čl. 8.10.2 a) normy\
                                  ČSN 73 0802. Šachta slouží pro přepravu\
                                  osob, nebo jako malý nákladní výtah.'
                    loc_str_v = pre_str + app_vytah + suf_str_v
                    doc.append(NoEscape(loc_str_v))
                if info_PU[5] == ['osobně-nákladní, nákladní výtahy']:
                    if h_p <= 30:
                        s_vytah = ['Šv', 'Výtahové šachty', '-', '-', '-',
                                   '-', '-', 'III.']
                    if h_p > 30:
                        s_vytah = ['Šv', 'Výtahové šachty', '-', '-', '-',
                                   '-', '-', 'IV.']
                    len_vytah = len(new_dict) + 1
                    app_vytah = pre + str(len_vytah) + suf
                    suf_str_v = ' Výtahová šachta odpovídá čl. 8.10.2 normy\
                                  ČSN 73 0802 bodu b). Šachta slouží jako\
                                  osobo nákladní nebo nákladní výtah.'
                    loc_str_v = pre_str + app_vytah + suf_str_v
                    doc.append(NoEscape(loc_str_v))
                app_this = pre + str(len_vytah) + suf
                s_vytah[1] = NoEscape(s_vytah[1] + app_this)
            if info_PU[4] == ['ANO']:
                try:
                    new_dict
                except NameError:
                    var_exist = False
                else:
                    var_exist = True
                if var_exist is False:
                    new_dict = []
                    pre_str = '\item'
                    doc.append(Command('begin', 'TableNotes'))
                    doc.append(NoEscape('\small'))
                if info_PU[6] == ['rozvody hořlavých látek – max 1000 mm2']:
                    if h_p <= 22.5:
                        s_inst = ['Ši', 'Instalační šachty', '-', '-', '-',
                                  '-', '-', 'II.']
                    if 22.5 < h_p <= 45.0:
                        s_inst = ['Ši', 'Instalační šachty', '-', '-', '-',
                                  '-', '-', 'III.']
                    if h_p > 45.0:
                        s_inst = ['Ši', 'Instalační šachty', '-', '-', '-',
                                  '-', '-', 'IV.']
                    if info_PU[3] == ['ANO']:
                        len_inst = len_vytah + 1
                        app_vytah = pre + str(len_inst) + suf
                        suf_str_i = ' Instalační šachty jsou zatříděny v souladu s\
                                    čl. 8.12.2 c). Šachty jsou dimenzovány\
                                    pro rozvody hořlavých látek o celkovém\
                                    průřezu 1000 m$^2$'
                        loc_str_i = pre_str + app_vytah + suf_str_i
                        doc.append(NoEscape(loc_str_i))
                if info_PU[6] == ['rozvody nehořlavých látek – potrubí A1,\
                                  A2']:
                    s_inst = ['Ši', 'Instalační šachty', '-', '-', '-',
                              '-', '-', 'I.']
                if info_PU[6] == ['rozvody nehořlavých látek – potrubí B-F']:
                    s_inst = ['Ši', 'Instalační šachty', '-', '-', '-', '-',
                              '-', 'II.']
                if info_PU[6] == ['rozvody hořlavých látek – 1000-8000 mm2']:
                    if h_p <= 45.0:
                        s_inst = ['Ši', 'Instalační šachty', '-', '-', '-', '-',
                                  '-', 'IV.']
                    if h_p > 45.0:
                        s_inst = ['Ši', 'Instalační šachty', '-', '-', '-', '-',
                                  '-', 'V.']
                if info_PU[6] == ['rozvody hořlavých látek – více než 8000 mm2']:
                    s_inst = ['Ši', 'Instalační šachty', '-', '-', '-', '-',
                              '-', 'VI.']
                if info_PU[3] == ['ANO']:
                    len_inst = len_vytah + 1
                else:
                    len_inst = len(new_dict) + 1
                app_this = pre + str(len_inst) + suf
                s_inst[1] = NoEscape(s_inst[1] + app_this)
                data_used.append(s_vytah)
                data_used.append(s_inst)
            if len(check) > 0 or info_PU[4] == ['ANO'] or info_PU[3] == ['ANO']:
                doc.append(Command('end', 'TableNotes'))
            for i in range(0, len(data_used)):
                data_used[i][2] = data_used[i][2].replace(".", ",")
                data_used[i][3] = data_used[i][3].replace(".", ",")
                data_used[i][4] = data_used[i][4].replace(".", ",")
                data_used[i][5] = data_used[i][5].replace(".", ",")
                data_used[i][6] = data_used[i][6].replace(".", ",")
            with doc.create(LongTable("l l c c c c c c", pos=['htb'])) as data_table:
                doc.append(Command('caption', 'Přehled požárních úselků a jejich SPB'))
                doc.append(Command('label', 'PU'))
                doc.append(Command('\ '))
                data_table.append
                data_table.add_hline()
                data_table.add_row(["Číslo", "Popis", "Plocha",
                                    MultiColumn(3, data='Součinitelé'),
                                    NoEscape('p$_v$'), "SPB"])
                data_table.add_row([" ", " ", NoEscape('[m$^2$]'), "a",
                                    "b", "c", NoEscape('[kg/m$^2$]'), '[-]'])
                data_table.add_hline()
                data_table.end_table_header()
                for i in range(0, len(data_used)):
                    if i % 2 != 0:
                        data_table.add_row(data_used[i], color="Hex")
                    else:
                        data_table.add_row(data_used[i])
                data_table.add_hline()
                if len(check) > 0 or info_PU[4] == ['ANO'] or info_PU[3] == ['ANO']:
                    doc.append(NoEscape('\insertTableNotes'))
                os.chdir(vystup_dir)
            doc.append(Command('end', 'ThreePartTable'))
    doc.generate_pdf("D_PU", clean_tex=False)

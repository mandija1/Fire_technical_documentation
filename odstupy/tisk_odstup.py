import unicodedata
from dxfwrite import DXFEngine as dxf
from dxfwrite.dimlines import dimstyles, LinearDimension
import os
import numpy as np
from math import atan2, degrees, cos, sin, radians, atan


def define_circle(p1, p2, p3):
    """
    Returns the center and radius of the circle passing the given 3 points.
    In case the 3 points form a line, returns (None, infinity).
    """
    temp = p2[0] * p2[0] + p2[1] * p2[1]
    bc = (p1[0] * p1[0] + p1[1] * p1[1] - temp) / 2
    cd = (temp - p3[0] * p3[0] - p3[1] * p3[1]) / 2
    det = (p1[0] - p2[0]) * (p2[1] - p3[1]) - (p2[0] - p3[0]) * (p1[1] - p2[1])

    if abs(det) < 1.0e-6:
        return (None, np.inf)

    # Center of circle
    cx = (bc*(p2[1] - p3[1]) - cd*(p1[1] - p2[1])) / det
    cy = ((p1[0] - p2[0]) * cd - (p2[0] - p3[0]) * bc) / det

    radius = np.sqrt((cx - p1[0])**2 + (cy - p1[1])**2)
    return ((cx, cy), radius)


def define_angle(p1, p2):
    xDiff = p2[0] - p1[0]
    yDiff = p2[1] - p1[1]
    return degrees(atan2(yDiff, xDiff))


def distance(x, y):
    dis = ((y[0] - x[0]) ** 2 + (y[1] - x[1]) ** 2) ** 0.5
    return dis


def tisk_odstup(df_result, df_input, df_rotate):
    dir_path = os.path.dirname(os.path.realpath(__file__))
    os.chdir(os.path.dirname(__file__))
    print(dir_path)

    ###################################
    drawing = dxf.drawing('drawing.dxf')
    drawing.header['$LTSCALE'] = 500
    # dimensionline setup:
    # add block and layer definition to drawing
    if df_input['Hodnota'][5] == 'ANO':
        dimstyles.setup(drawing)
        dimstyles.new("dots", tickfactor=df_input['Hodnota'][6] * 2.5, height=df_input['Hodnota'][6] * 2.5, scale=1.,  textabove=df_input['Hodnota'][6])

    # Přidání hladin do výkresu
    drawing.add_layer('LINES', color=5, linetype='DIVIDE')

    df_result_tisk = df_result[['sirka', 'd', 'd\'', 'd\'s']] * 1000
    df_result_text = df_result['Nazev/fasada']

    '''df_result_text ='''
    odstup = df_result_tisk.values.tolist()
    rotation = df_rotate.values.tolist()
    text = df_result_text.values.tolist()

    x = 0
    for i in range(len(odstup)):
        # Středová kružnice
        des_B = distance([x, 0], [x + odstup[i][0]/2, odstup[i][1]])
        des_C = distance([x, 0], [x + odstup[i][0], odstup[i][2]])
        deg_A = radians(90)
        deg_B = atan(odstup[i][1] / (odstup[i][0]/2))
        deg_C = atan(odstup[i][2] / (odstup[i][0]))

        A = [x + cos(radians(rotation[i][0]) + deg_A) * odstup[i][2],
             0 + sin(radians(rotation[i][0]) + deg_A) * odstup[i][2]]
        B = [x + cos(radians(rotation[i][0]) + deg_B) * des_B,
             0 + sin(radians(rotation[i][0]) + deg_B) * des_B]
        C = [x + cos(radians(rotation[i][0]) + deg_C) * des_C,
             0 + sin(radians(rotation[i][0]) + deg_C) * des_C]

        center, radius = define_circle(A, B, C)
        if center is not None:
            angle_end = define_angle(center, A)
            angle_begin = define_angle(center, C)

        # Vzdálenost od počátku ke středu boční pravé kružnice
        des_C_r = distance([x, 0], [x + odstup[i][0], odstup[i][2]/2])
        # Úhel od počátku ke středu boční pravé kružnice
        deg_C_r = atan(odstup[i][2]/2 / (odstup[i][0]))

        # Souřanice středu bočních kružnic
        A_l = [x + cos(radians(rotation[i][0]) + deg_A) * odstup[i][2]/2,
               0 + sin(radians(rotation[i][0]) + deg_A) * odstup[i][2]/2]
        A_r = [x + cos(radians(rotation[i][0]) + deg_C_r) * des_C_r,
               0 + sin(radians(rotation[i][0]) + deg_C_r) * des_C_r]

        # Vykreslení kružnic
        arc = dxf.arc(odstup[i][3], A_l, 90 + rotation[i][0], 270 + rotation[i][0], layer='LINES')
        drawing.add(arc)
        arc = dxf.arc(odstup[i][3], A_r, 270 + rotation[i][0], 90 + rotation[i][0], layer='LINES')
        drawing.add(arc)

        if center is not None:
            arc = dxf.arc(radius, center, angle_begin, angle_end, layer='LINES')
            drawing.add(arc)
        else:
            line = dxf.line(A, C, layer='LINES')
            drawing.add(line)
        text_to_add = text[i]
        plain_text = unicodedata.normalize('NFKD', text_to_add)
        text_added = dxf.text(plain_text, (x, odstup[i][1] + 1000), height=500)
        drawing.add(text_added)
        if df_input['Hodnota'][5] == 'ANO':

            des_boc = distance([x, 0], [x - odstup[i][3], odstup[i][2]/2])
            deg_boc = atan(odstup[i][2]/2 / (-(odstup[i][3])))

            A_boc = [x - cos(radians(rotation[i][0]) + deg_boc) * des_boc,
                   0 - sin(radians(rotation[i][0]) + deg_boc) * des_boc]

            angll = rotation[i][0]
            if 90 < rotation[i][0] < 270:
                angll = rotation[i][0]-180

            des_center = distance([x, 0], [x + odstup[i][0]/2, odstup[i][1]/2])
            deg_center = atan(odstup[i][1]/2 / (odstup[i][0]/2))

            center_text = [x + cos(radians(rotation[i][0]) + deg_center) * des_center,
                           0 + sin(radians(rotation[i][0]) + deg_center) * des_center]

            des_top = distance([x, 0], [x + odstup[i][0]/2, odstup[i][1]])
            deg_top = atan(odstup[i][1] / (odstup[i][0]/2))

            top_point = [x + cos(radians(rotation[i][0]) + deg_top) * des_top,
                         0 + sin(radians(rotation[i][0]) + deg_top) * des_top]

            des_bot = distance([x, 0], [x + odstup[i][0]/2, 0])
            deg_bot = atan(0 / (odstup[i][0]/2))

            bot_point = [x + cos(radians(rotation[i][0]) + deg_bot) * des_bot,
                         0 + sin(radians(rotation[i][0]) + deg_bot) * des_bot]

            angllee = rotation[i][0]
            if 0 < rotation[i][0] < 180:
                angllee = rotation[i][0]-180

            drawing.add(LinearDimension(A_l, [A_l, A_boc], dimstyle='dots', angle=angll))
            drawing.add(LinearDimension(center_text, [top_point, bot_point], dimstyle='dots', angle=angllee + 90))
        x += 15000

    drawing.save()

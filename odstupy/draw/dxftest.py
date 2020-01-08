def tisk_odstup(df_result):
    import ezdxf
    import os
    import numpy as np
    from math import atan2, degrees

    dir_path = os.path.dirname(os.path.realpath(__file__))
    os.chdir(os.path.dirname(__file__))
    print(dir_path)

    ###################################
    dwg = ezdxf.new('AC1015')  # hatch requires the DXF R2000 (AC1015) format or later
    msp = dwg.modelspace()  # adding entities to the model space


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


    odstup = [[5000, 3750, 2600, 1300], [6000, 3750, 2600, 1300]]
    x = 0

    for i in range(len(odstup)):
        print(odstup[i])
        # Středová kružnice
        A = [x, odstup[i][2]]
        B = [x + odstup[i][0]/2, odstup[i][1]]
        C = [x + odstup[i][0], odstup[i][2]]

        center, radius = define_circle(A, B, C)
        angle_end = define_angle(center, A)
        angle_begin = define_angle(center, C)

        msp.add_arc([x, odstup[i][2]/2], odstup[i][3], 90, 270)
        msp.add_arc([x+odstup[i][0], odstup[i][2]/2], odstup[i][3], 270, 90)
        msp.add_arc(center, radius, angle_begin, angle_end)
        x += 10000

    dwg.saveas("solid_hatch_polyline_path.dxf")

def chunks(l, n):  # This definition splits the list into desired chunks
    # For item i in a range that is a length of l,
    for i in range(0, len(l), n):
        # Create an index range for l of n items:
        yield l[i:i+n]


def konst_data_prep(name, data):
    Int = 9
    if 'E.1 PKD_Stěna' in data:
        indices = [i for i, x in enumerate(data)
                   if x == name]
        indices[:] = [x / Int for x in indices]
        for i in range(0, len(indices)):
            indices[i] = (int(indices[i])+1)*9
        ay = []
        for i in range(0, len(indices)):
            ay.append(range(indices[i]-9, indices[i]))
        ax = []
        for item in ay:
            for p in item:
                ax.append(p)
        data_prep = []
        for item in ax:
            data_prep.append(data[item])
        data_prep = list(chunks(data_prep, 9))
        Nazev_kce = []
        Specif1 = []
        Specif2 = []
        b = []
        a = []
        h = []
        Ly = []
        Lx = []
        zeb_a = []
        for i in range(0, len(data_prep)):
            Nazev_kce.append((data_prep[i])[0])
            Specif1.append((data_prep[i])[1])
            Specif2.append((data_prep[i])[2])
            b.append((data_prep[i])[3])
            a.append((data_prep[i])[4])
            h.append((data_prep[i])[5])
            Ly.append((data_prep[i])[6])
            Lx.append((data_prep[i])[7])
            zeb_a.append((data_prep[i])[8])
        return Nazev_kce, Specif1, Specif2, b, a, h, Ly, Lx, zeb_a

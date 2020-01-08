def chunks(l, n):  # This definition splits the list into desired chunks
    # For item i in a range that is a length of l,
    for i in range(0, len(l), n):
        # Create an index range for l of n items:
        yield l[i:i+n]


def open_konst(konst, vystup_dir):
    data = []  # empty list for storing data
    for row in konst.iter_rows(min_row=4, max_row=41,
                               min_col=1, max_col=1):
        for cell in row:
            data.append(cell.value)
    data_len = 0
    for i in range(0, len(data)):
        if data[i] is not None:
            data_len += 1
        else:
            break
    data = []  # empty list for storing data
    for row in konst.iter_rows(min_row=4, max_row=data_len+3,
                               min_col=1, max_col=9):
        for cell in row:
            data.append(cell.value)
    '''if len(data) > 9:
        data = list(chunks(data, 9))'''
    return data

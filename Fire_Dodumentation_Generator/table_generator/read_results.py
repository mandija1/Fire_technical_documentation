import csv


def read_results(dir):
    with open(dir, 'r', encoding='utf-8') as f:
        reader = csv.reader(f, dialect='excel', delimiter='\t')
        data = []
        for row in reader:
            data.append(row)
        data = [x for x in data if x != []]
    return data


def chunks(l, n):  # This definition splits the list into desired chunks
    # For item i in a range that is a length of l,
    for i in range(0, len(l), n):
        # Create an index range for l of n items:
        yield l[i:i+n]

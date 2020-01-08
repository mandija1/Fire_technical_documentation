import os
import csv

def read_results(dir):
    # os.chdir("c:/Users/Honza/Google Drive/Práce/Automatizace požární zprávy/")
    with open(dir, 'r', encoding='utf-8') as f:
        reader = csv.reader(f, dialect='excel', delimiter='\t')
        data = []
        for row in reader:
            data.append(row)
        data = [x for x in data if x != []]
    return data

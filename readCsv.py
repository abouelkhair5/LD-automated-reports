import csv
with open('reports/Bing.csv') as bingFile:
    reader = csv.DictReader(bingFile.readlines()[3:])
    branded_totals = {'clicks': 0, 'impr': 0, 'spend': 0}
    non_branded_totals = {'clicks': 0, 'impr': 0, 'spend': 0}
    for row in reader:
        if 'Non-Brand' in row['Campaign']:
            non_branded_totals['clicks'] += int(row['Clicks'])
            non_branded_totals['impr'] += int(row['Impr.'])
            non_branded_totals['spend'] += float(row['Spend'])
        elif 'Branded' in row['Campaign']:
            branded_totals['clicks'] += int(row['Clicks'])
            branded_totals['impr'] += int(row['Impr.'])
            branded_totals['spend'] += float(row['Spend'])
        print row
    print non_branded_totals
    print branded_totals

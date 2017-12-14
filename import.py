import csv
from re import sub
from decimal import Decimal
from xlrd import open_workbook
from xlutils.copy import copy


google_file = open('reports/Google Adwords.csv')
bing_file = open('reports/Bing.csv')

google_reader = csv.reader(google_file)
bing_reader = csv.reader(bing_file)

# skipping the header of the google file
for x in range(0, 3):
    next(google_reader)

summary = dict()
summary['Branded'] = dict()
summary['Non-Brand'] = dict()
summary['Branded']['clicks'] = 0
summary['Branded']['impressions'] = 0
summary['Branded']['cost'] = 0
summary['Non-Brand']['clicks'] = 0
summary['Non-Brand']['impressions'] = 0
summary['Non-Brand']['cost'] = 0

for row in google_reader:
    campaign = row[1].split(' ')[0]
    clicks = int(row[9].replace(',', ''))
    impressions = int(row[10].replace(',', ''))
    cost = Decimal(sub(r'[^\d.]', '', row[13]))
    if campaign:
        if campaign not in summary.keys():
            summary[campaign] = dict()
            summary[campaign]['clicks'] = 0
            summary[campaign]['impressions'] = 0
            summary[campaign]['cost'] = 0
        summary[campaign]['clicks'] += clicks
        summary[campaign]['impressions'] += impressions
        summary[campaign]['cost'] += cost
print(summary)

Week = 48
base_row = Week * 4 - 2
branded_row = base_row + 1
non_brand_row = base_row + 2
cost_col = 9
clicks_col = 33
impr_col = 37
search_sheet = 4

rb = open_workbook("reports/BeautyBoutique Paid Marketing Tracker (Dec 1, 2017).xlsx")
wb = copy(rb)

s = wb.get_sheet(search_sheet)
s.write(branded_row, clicks_col, summary['Branded']['clicks'])
s.write(branded_row, impr_col, summary['Branded']['impressions'])
s.write(branded_row, cost_col, summary['Branded']['cost'])
s.write(non_brand_row, clicks_col, summary['Non-Brand']['clicks'])
s.write(non_brand_row, impr_col, summary['Non-Brand']['impressions'])
s.write(non_brand_row, cost_col, summary['Non-Brand']['cost'])
wb.save('names.xlsx')

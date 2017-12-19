import csv
import sys
from re import sub
from decimal import Decimal
from openpyxl import load_workbook

if len(sys.argv) != 5:
    print "usage: auto-report <report-src> <google-src> <bing-src> <omniture-src> opt:<display-src> <dest>"
    exit(1)

report_file_name = sys.argv[1]
google_file_name = sys.argv[2]
bing_file_name = sys.argv[3]
omniture_file_name = sys.argv[4]

summary = dict()
summary['Branded'] = dict()
summary['Non-Brand'] = dict()
summary['Affiliate'] = dict()
summary['Display'] = dict()
summary['Branded']['clicks'] = 0
summary['Branded']['impressions'] = 0
summary['Branded']['cost'] = 0
summary['Non-Brand']['clicks'] = 0
summary['Non-Brand']['impressions'] = 0
summary['Non-Brand']['cost'] = 0

def headers_from_sheet(sheet):
    return map(lambda col: col[0].value, sheet.iter_cols(max_row=1))

def find_row_by_week(sheet, week_code):
    index = headers_from_sheet(sheet).index('FW')
    return map(lambda item: item.value, tuple(sheet.columns)[index]).index(week_code) + 1

with open(google_file_name) as google_file:
    reader = csv.DictReader(google_file.readlines()[2:])
    for row in reader:
        clicks = int(row['Clicks'].replace(',', ''))
        impressions = int(row['Impr.'].replace(',', ''))
        cost = Decimal(sub(r'[^\d.]', '', row['Cost']))
        if 'Non-Brand' in row['Campaign']:
            summary['Non-Brand']['clicks'] += clicks
            summary['Non-Brand']['impressions'] += impressions
            summary['Non-Brand']['cost'] += cost
        elif 'Branded' in row['Campaign']:
            summary['Branded']['clicks'] += clicks
            summary['Branded']['impressions'] += impressions
            summary['Branded']['cost'] += cost


with open(bing_file_name) as bing_file:
    reader = csv.DictReader(bing_file.readlines()[3:])
    for row in reader:
        clicks = int(row['Clicks'].replace(',', ''))
        impressions = int(row['Impr.'].replace(',', ''))
        cost = Decimal(sub(r'[^\d.]', '', row['Spend']))
        if 'Non-Brand' in row['Campaign']:
            summary['Non-Brand']['clicks'] += clicks
            summary['Non-Brand']['impressions'] += impressions
            summary['Non-Brand']['cost'] += cost
        elif 'Branded' in row['Campaign']:
            summary['Branded']['clicks'] += clicks
            summary['Branded']['impressions'] += impressions
            summary['Branded']['cost'] += cost

with open(omniture_file_name) as omniture_file:
    reader = csv.DictReader(omniture_file.readlines()[26:])
    for row in reader:
        sheet_name = row['Last Touch Channel']
        revenue = int(row['Revenue'])
        visits = int(row['Visits'])
        orders = int(row['Orders'])
        if sheet_name == 'Affiliate' :
            summary['Affiliate']['revenue'] = revenue
            summary['Affiliate']['visits'] = visits
            summary['Affiliate']['orders'] = orders
        elif sheet_name == 'Display':
            summary['Display']['revenue'] = revenue
            summary['Display']['visits'] = visits
            summary['Display']['orders'] = orders
        elif sheet_name == 'Paid Search Branded':
            summary['Branded']['revenue'] = revenue
            summary['Branded']['visits'] = visits
            summary['Branded']['orders'] = orders
        elif sheet_name == 'Paid Search Unbranded':
            summary['Non-Brand']['revenue'] = revenue
            summary['Non-Brand']['visits'] = visits
            summary['Non-Brand']['orders'] = orders



print summary

wb = load_workbook(sys.argv[1])
search_sheet = wb['Search']

Week = 'WK48'
headers = headers_from_sheet(search_sheet)

base_row = find_row_by_week(search_sheet, Week)
branded_row = base_row + 1
non_brand_row = base_row + 2
cost_col = headers.index('Spend') + 1
clicks_col = headers.index('Clicks') + 1
impr_col = headers.index('Impressions') + 1
revenue_col = headers.index('Revenue') + 1
visits_col = headers.index('Visits') + 1
orders_col = headers.index('Orders') + 1

search_sheet.cell(row=branded_row, column=clicks_col, value = summary['Branded']['clicks'])
search_sheet.cell(row=branded_row, column=impr_col, value = summary['Branded']['impressions'])
search_sheet.cell(row=branded_row, column=cost_col, value = summary['Branded']['cost'])
search_sheet.cell(row=branded_row, column=revenue_col, value = summary['Branded']['revenue'])
search_sheet.cell(row=branded_row, column=visits_col, value = summary['Branded']['visits'])
search_sheet.cell(row=branded_row, column=orders_col, value = summary['Branded']['orders'])
search_sheet.cell(row=non_brand_row, column=clicks_col, value = summary['Non-Brand']['clicks'])
search_sheet.cell(row=non_brand_row, column=impr_col, value = summary['Non-Brand']['impressions'])
search_sheet.cell(row=non_brand_row, column=cost_col, value = summary['Non-Brand']['cost'])
search_sheet.cell(row=non_brand_row, column=revenue_col, value = summary['Non-Brand']['revenue'])
search_sheet.cell(row=non_brand_row, column=visits_col, value = summary['Non-Brand']['visits'])
search_sheet.cell(row=non_brand_row, column=orders_col, value = summary['Non-Brand']['orders'])

affiliate_sheet = wb['Affiliate']
headers = headers_from_sheet(affiliate_sheet)

base_row = find_row_by_week(affiliate_sheet, Week)
revenue_col = headers.index('Revenue') + 1
visits_col = headers.index('Visits') + 1
orders_col = headers.index('Orders') + 1

affiliate_sheet.cell(row=base_row, column=revenue_col, value = summary['Affiliate']['revenue'])
affiliate_sheet.cell(row=base_row, column=visits_col, value = summary['Affiliate']['visits'])
affiliate_sheet.cell(row=base_row, column=orders_col, value = summary['Affiliate']['orders'])

display_sheet = wb['Display']
headers = headers_from_sheet(display_sheet)

base_row = find_row_by_week(display_sheet, Week)
revenue_col = headers.index('Revenue') + 1
visits_col = headers.index('Visits') + 1
orders_col = headers.index('Orders') + 1

display_sheet.cell(row=base_row, column=revenue_col, value = summary['Display']['revenue'])
display_sheet.cell(row=base_row, column=visits_col, value = summary['Display']['visits'])
display_sheet.cell(row=base_row, column=orders_col, value = summary['Display']['orders'])

wb.save('names.xlsx')

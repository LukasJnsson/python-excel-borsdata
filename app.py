import os
from dotenv import load_dotenv
import requests
import pandas as pd
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import Workbook
load_dotenv()


# API KEY
api_key = os.environ.get('API_KEY')


# API Endpoint
large_mid_cap_url = f'https://apiservice.borsdata.se/v1/instruments?authKey={api_key}'


# Fetch all companies listed on Large and Mid Cap
r = requests.get(large_mid_cap_url)
d = r.json()


# Array with Large and Mid Cap companies
large_mid_cap_companies = []

for largeCompany in d['instruments']:
    if largeCompany['marketId'] == 1:
        large_mid_cap_companies.append(largeCompany)

for midCompany in d['instruments']:
    if midCompany['marketId'] == 2:
        large_mid_cap_companies.append(midCompany)


# Map each company id with name
large_mid_cap_companies_dict = {company['insId']: company['name'] for company in large_mid_cap_companies}


# Each raw annual report
large_mid_cap_companies_reports = []

# Fetch annual report for each company
for company in large_mid_cap_companies:
    companyId = company['insId']
    c_url = f'https://apiservice.borsdata.se/v1/instruments/{companyId}/reports/year?authKey={api_key}'
    response = requests.get(c_url)
    data = response.json()
    large_mid_cap_companies_reports.append(data)


# Excel header
df = pd.DataFrame(columns=["År", "Namn", "ROIC", "Börsvärde", "Totala tillgångar", "Totala skulder", "Skuldsättningsgrad"])


# Iterate and select the desired key figures from each company
for p in large_mid_cap_companies_reports:
    
    instrument_id = p['instrument']
    instrument_name = large_mid_cap_companies_dict[instrument_id] # map name from arr

    for report in p['reports']:

        year = report['year']
        if year >= 2017 and year <= 2022: # Select reports in range 2017 - 2022

            # ROIC https://github.com/Borsdata-Sweden/API/wiki/KPI-History (KpiId 37)
            roic_url = f"https://apiservice.borsdata.se/v1/Instruments/{instrument_id}/kpis/37/year/mean/history?authKey={api_key}"
            roic_response = requests.get(roic_url)
            roic_data = roic_response.json()

            # Get ROIC for the current year
            for item in roic_data['values']:
                if item['y'] == year:
                    #ROIC
                    roic = item['v']
                    break


            # Market value
            stock_Price_Average = int(report['stock_Price_Average'])
            number_Of_Shares = int(report['number_Of_Shares'])
            market_value = ( stock_Price_Average * number_Of_Shares )
            

            # Total assets
            total_assets = int(report['total_Assets'])


            # Total liabilities
            current_Liabilities = int(report['current_Liabilities'])
            non_Current_liabilities = int(report['non_Current_Liabilities'])
            
            total_liabilities = ( non_Current_liabilities + current_Liabilities )


            # Debt ratio
            total_equity = int(report['total_Equity'])
            try: 
                debt_ratio = ( (non_Current_liabilities + current_Liabilities ) / total_equity )
            
            except ZeroDivisionError:
                debt_ratio = 0


            # Append the data to the dataframe
            df = df.append({
                "År": year,
                "Namn": instrument_name,
                "ROIC": roic,
                "Börsvärde": market_value,
                "Totala tillgångar": total_assets,
                "Totala skulder": total_liabilities,
                "Skuldsättningsgrad": debt_ratio
            }, ignore_index=True)


# Write the data to the Excel file
wb = Workbook()
ws = wb.active
for r in dataframe_to_rows(df, index=False, header=True):
    ws.append(r)
    # Name of the created Excel file
wb.save("borsdata.xlsx")
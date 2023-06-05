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
nordic_market_companies_url = f'https://apiservice.borsdata.se/v1/instruments?authKey={api_key}'


def fetchNordicMarketCompanies(nordic_market_companies_url):
    """
    Function that fetch all companies listed on the Nordic Market
    """
    companies = requests.get(nordic_market_companies_url)
    json_companies = companies.json()

    return json_companies


def fetchLargeMidCap(nordic_market_companies_url):
    """
    Function that fetch all companies listed OMX Large and Mid Cap
        marketId === 1 (Large cap)
        marketId === 2 (Mid cap)
    """
    companies = fetchNordicMarketCompanies(nordic_market_companies_url)

    # Array with Large and Mid Cap companies
    large_mid_cap_companies = []
    
    for largeCompany in companies['instruments']:
        if largeCompany['marketId'] == 1:
            large_mid_cap_companies.append(largeCompany)

    for midCompany in companies['instruments']:
        if midCompany['marketId'] == 2:
            large_mid_cap_companies.append(midCompany)

    return large_mid_cap_companies


def mapEachCompany(nordic_market_companies_url):
    """
    Function that map each company 'insId' with 'name'
    """
    large_mid_cap_companies = fetchLargeMidCap(nordic_market_companies_url)
    large_mid_cap_companies_dict = {company['insId']: company['name'] for company in large_mid_cap_companies}

    return large_mid_cap_companies_dict


def fetchAnnualReports(nordic_market_companies_url):
    """
    Function that fetch all the annual reports for each company
    """
    large_mid_cap_companies = fetchLargeMidCap(nordic_market_companies_url)

    # Array with each raw report
    large_mid_cap_companies_reports = []

    for company in large_mid_cap_companies:
        companyId = company['insId']
        company_url = f'https://apiservice.borsdata.se/v1/instruments/{companyId}/reports/year?authKey={api_key}'
        response = requests.get(company_url)
        data = response.json()
        large_mid_cap_companies_reports.append(data)

    return large_mid_cap_companies_reports


def fetchKeyFigures(nordic_market_companies_url):
    """
    Funtion that fetch the selected key figures and creates Excel file
    """
    large_mid_cap_companies_dict = mapEachCompany(nordic_market_companies_url)
    large_mid_cap_companies_reports = fetchAnnualReports(nordic_market_companies_url)


    # Excel column headers
    data_frame = pd.DataFrame(columns=["År", "Namn", "ROIC", "Börsvärde", "Totala tillgångar", "Totala skulder", "Skuldsättningsgrad"])


    # Iterate and select the desired key figures from each company
    for company in large_mid_cap_companies_reports:
        
        instrument_id = company['instrument']
        instrument_name = large_mid_cap_companies_dict[instrument_id] # map name from array

        for report in company['reports']:

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
                        try:
                            roic = item['v']
                        except Exception as e:
                            print(f'Error {e}')
                        break

                # Market value
                try:
                    stock_Price_Average = int(report['stock_Price_Average'])
                    number_Of_Shares = int(report['number_Of_Shares'])
                    market_value = ( stock_Price_Average * number_Of_Shares )
                except Exception as e:
                    print(f'Error {e}')

                # Total assets
                try:
                    total_assets = int(report['total_Assets'])
                except Exception as e:
                    print(f'Error {e}')

                # Total liabilities
                try:
                    current_Liabilities = int(report['current_Liabilities'])
                    non_Current_liabilities = int(report['non_Current_Liabilities'])
                    total_liabilities = ( non_Current_liabilities + current_Liabilities )
                except Exception as e:
                    print(f'Error {e}')

                # Debt ratio
                try:
                    total_equity = int(report['total_Equity'])
                    debt_ratio = ( (non_Current_liabilities + current_Liabilities ) / total_equity )
                
                except Exception as e:
                    print(f'Error {e}')

                # Append the data to the dataframe
                data_frame = data_frame.append({
                    "År": year,
                    "Namn": instrument_name,
                    "ROIC": roic,
                    "Börsvärde": market_value,
                    "Totala tillgångar": total_assets,
                    "Totala skulder": total_liabilities,
                    "Skuldsättningsgrad": debt_ratio
                }, ignore_index=True)

    try:
        # Write the data to the Excel file
        workbook = Workbook()
        ws = workbook.active
        for row in dataframe_to_rows(data_frame, index=False, header=True):
            ws.append(row)
            # Name of the created Excel file
        workbook.save("borsdata.xlsx")

    except Exception as e:
        print(f'Error {e}')

fetchKeyFigures(nordic_market_companies_url)
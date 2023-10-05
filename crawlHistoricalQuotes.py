from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support import expected_conditions as EC

from datetime import date, timedelta
import time
from datetime import datetime
import csv
import os

ERROR_CODE = []

def openChrome():
    # path = 'http://en.stockbiz.vn/Stocks/'+code+'/HistoricalQuotes.aspx'
    chrome_driver_path = ChromeDriverManager().install()
    chrome_service = Service(executable_path=chrome_driver_path)
    browser = webdriver.Chrome(service=chrome_service)
    browser.maximize_window()
    # browser.get(path)
    return browser


def write_to_csv(filename, browser, code):
    while True:
        try:
            rows = browser.find_element(
                'xpath', '/html/body/form/table/tbody/tr/td[2]/div/div[3]/table/tbody/tr[2]/td[1]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td/div/div[2]/table').find_elements(By.TAG_NAME, 'tr')

            for i in range(2, len(rows)):
                value = []
                row = browser.find_elements(
                    "xpath", '/html/body/form/table/tbody/tr/td[2]/div/div[3]/table/tbody/tr[2]/td[1]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td/div/div[2]/table/tbody/tr['+str(i)+']/td')
                date = row[0].text
                close = row[5].text
                volumn = row[8].text.replace(',', '')
                value.append(code)
                value.append(date)
                value.append(close)
                value.append(volumn)

                with open(filename, 'a', encoding='UTF8', newline='') as f:
                    writer = csv.writer(f)
                    writer.writerow(value)
                f.close()
            if 'Next' in rows[-1].text:
                rows[-1].find_elements(By.TAG_NAME, "a")[-1].click()
                time.sleep(0.3)
            else:
                break
        except:
            continue


def crawl(filename, CODES, start, end):
    browser = openChrome()
    for code in CODES:
        path = 'http://en.stockbiz.vn/Stocks/'+code+'/HistoricalQuotes.aspx'
        try:
            browser.get(path)
            # time.sleep(0.3)
            # set start_date
            start_date = browser.find_element(
                'xpath', '/html/body/form/table/tbody/tr/td[2]/div/div[3]/table/tbody/tr[2]/td[1]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td/div/div[1]/table/tbody/tr/td[2]/table/tbody/tr/td[1]/input[3]')
            start_date.clear()
            start_date.send_keys(start)

            # set end_date
            end_date = browser.find_element(
                'xpath', '/html/body/form/table/tbody/tr/td[2]/div/div[3]/table/tbody/tr[2]/td[1]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td/div/div[1]/table/tbody/tr/td[4]/table/tbody/tr/td[1]/input[3]')
            end_date.clear()
            end_date.send_keys(end)

            # click View button
            browser.find_element(
                'xpath', '/html/body/form/table/tbody/tr/td[2]/div/div[3]/table/tbody/tr[2]/td[1]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td/div/div[1]/table/tbody/tr/td[5]/input').click()

            # write to csv
            write_to_csv(filename, browser, code)
        except:
            ERROR_CODE.append(code)
            continue
    if len(ERROR_CODE) != 0:
        print("Error code: ", ERROR_CODE)

#####################################################


if __name__ == "__main__":
    filename = 'HistoricalQuotes_data.csv'
    START_DATE = "01/01/2008"
    END_DATE = " 08/31/2023"
    CODES = ["AAA", "AAM", "AAT", "ABR", "ABS", "ABT", "ACB", "ACC", "ACG", "ACL", "ADG", "ADP", "ADS", "AGG", "AGM", "AGR", "ANV", "APC", "APG", "APH", "ASG", "ASM", "ASP", "AST", "ATP", "BAF", "BBC", "BCE", "BCG", "BCM", "BFC", "BHN", "BIC", "BID", "BKG", "BMC", "BMI", "BMP", "BRC", "BSI", "BTP", "BTT", "BVH", "BWE", "C32", "C47", "CAV", "CCI", "CCL", "CDC", "CHP", "CIG", "CII", "CKG", "CLC", "CLL", "CLW", "CMG", "CMV", "CMX", "CNG", "COM", "CRC", "CRE", "CRV", "CSM", "CSV", "CTD", "CTF", "CTG", "CTI", "CTR", "CTS", "CVT", "D2D", "DAG", "DAH", "DAT", "DBC", "DBD", "DBT", "DC4", "DCL", "DCM", "DDB", "DGC", "DGW", "DHA", "DHC", "DHG", "DHM", "DIG", "DLG", "DMC", "DPG", "DPM", "DPR", "DQC", "DRC", "DRH", "DRL", "DSN", "DTA", "DTL", "DTT", "DVP", "DXG", "DXS", "DXV", "EIB", "ELC", "EVE", "EVF", "EVG", "FCM", "FCN", "FDC", "FIR", "FIT", "FMC", "FPT", "FRT", "FTS", "GAS", "GDT", "GEG", "GEX", "GIL", "GMC", "GMD", "GMH", "GSP", "GTA", "GVR", "HAG", "HAH", "HAP", "HAR", "HAS", "HAX", "HBC", "HCD", "HCM", "HDB", "HDC", "HDG", "HHP", "HHS", "HHV", "HID", "HII", "HMC", "HNG", "HPG", "HPX", "HQC", "HRC", "HSG", "HSL", "HT1", "HTI", "HTL", "HTN", "HTV", "HU1", "HUB", "HVH", "HVN", "HVX", "IBC", "ICT", "IDI", "IJC", "ILB", "IMP", "ITA", "ITC", "ITD", "JVC", "KBC", "KDC", "KDH", "KHG", "KHP", "KMR", "KOS", "KPF", "KSB", "L10", "LAF", "LBM", "LCG", "LDG", "LEC", "LGC", "LGL", "LHG", "LIX", "LM8",
             "LPB", "LSS", "MBB", "MCP", "MDG", "MHC", "MIG", "MSB", "MSH", "MSN", "MWG", "NAF", "NAV", "NBB", "NCG", "NCT", "NHA", "NHH", "NHT", "NKG", "NLG", "NNC", "NO1", "NSC", "NT2", "NTL", "NVL", "NVT", "OCB", "OGC", "OPC", "ORS", "PAC", "PAN", "PC1", "PDN", "PDR", "PET", "PGC", "PGD", "PGI", "PGV", "PHC", "PHR", "PIT", "PJT", "PLP", "PLX", "PMG", "PNC", "PNJ", "POM", "POW", "PPC", "PSH", "PTB", "PTC", "PTL", "PVD", "PVP", "PVT", "QBS", "QCG", "RAL", "RDP", "REE", "S4A", "SAB", "SAM", "SAV", "SBA", "SBG", "SBT", "SBV", "SC5", "SCD", "SCR", "SCS", "SFC", "SFG", "SFI", "SGN", "SGR", "SGT", "SHA", "SHB", "SHI", "SHP", "SIP", "SJD", "SJF", "SJS", "SKG", "SMA", "SMB", "SMC", "SPM", "SRC", "SRF", "SSB", "SSC", "SSI", "ST8", "STB", "STG", "STK", "SVC", "SVD", "SVI", "SVT", "SZC", "SZL", "TBC", "TCB", "TCD", "TCH", "TCL", "TCM", "TCO", "TCR", "TCT", "TDC", "TDG", "TDH", "TDM", "TDP", "TDW", "TEG", "TGG", "THG", "TIP", "TIX", "TLD", "TLG", "TLH", "TMP", "TMS", "TMT", "TN1", "TNA", "TNC", "TNH", "TNI", "TNT", "TPB", "TPC", "TRA", "TRC", "TSC", "TTA", "TTB", "TTE", "TTF", "TV2", "TVB", "TVS", "TVT", "TYA", "UIC", "VAF", "VCA", "VCB", "VCF", "VCG", "VCI", "VDP", "VDS", "VFG", "VGC", "VHC", "VHM", "VIB", "VIC", "VID", "VIP", "VIX", "VJC", "VMD", "VND", "VNE", "VNG", "VNL", "VNM", "VNS", "VOS", "VPB", "VPD", "VPG", "VPH", "VPI", "VPS", "VRC", "VRE", "VSC", "VSH", "VSI", "VTB", "VTO", "YBM", "YEG"]
    


    crawl(os.path.dirname(os.path.abspath(__file__)) +
          '/'+filename, CODES, START_DATE, END_DATE)

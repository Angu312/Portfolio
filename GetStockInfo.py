from selenium import webdriver
import time
import os
import pandas as pd
import matplotlib.pyplot as plt

TickerSymbol = str(input("What is the company's ticker symbol on the NY Stock Exchange? "))

KeyRatios = "http://financials.morningstar.com/ratios/r.html?t=" + TickerSymbol + "&region=usa&culture=en-US"
BalanceSheet = "http://financials.morningstar.com/balance-sheet/bs.html?t=" + TickerSymbol + "&region=usa&culture=en-US"
PERatio = "http://financials.morningstar.com/valuation/price-ratio.html?t=" + TickerSymbol + "&region=usa&culture=en-US"

ClickRobot = webdriver.Chrome(executable_path=r'C:\Users\Angu312\Anaconda3\Scripts\chromedriver.exe')
ClickRobot.get(KeyRatios)
ClickRobot.find_element_by_css_selector('.large_button').click()
time.sleep(4)
ClickRobot.quit()

ClickRobot = webdriver.Chrome(executable_path=r'C:\Users\Angu312\Anaconda3\Scripts\chromedriver.exe')
ClickRobot.get(BalanceSheet)
ClickRobot.find_element_by_css_selector('.rf_export').click()
time.sleep(4)
ClickRobot.quit()

ClickRobot = webdriver.Chrome(executable_path=r'C:\Users\Angu312\Anaconda3\Scripts\chromedriver.exe')
ClickRobot.get(PERatio)

os.chdir(r'C:/Users/Angu312/Documents/')
BalanceSheetData = pd.read_csv(TickerSymbol + ' Balance Sheet.csv', skiprows=1, index_col=0)
KeyRatiosData = pd.read_csv(TickerSymbol + ' Key Ratios.csv', skiprows=2, index_col=0)

InterestCoverage = KeyRatiosData.loc["Interest Coverage"].tolist()
AnnualPeriod = KeyRatiosData.loc["Profitability"].tolist()
dictionary = dict(zip(InterestCoverage, AnnualPeriod))
for i in InterestCoverage:
    if float(i) >= 6:
        print('Interest Coverage Ratio is GOOD at ' + str(i) + ' on ' + dictionary[i])
    else:
        print('Interest Coverage Ratio is BAD at ' + str(i)+ ' on ' + dictionary[i])

Liabilities = BalanceSheetData.loc["Total liabilities"].tolist()
MostRecentLiability = int(Liabilities[-1])
Assets = BalanceSheetData.loc["Total assets"].tolist()
MostRecentAsset = int(Assets[-1])

if MostRecentLiability/MostRecentAsset < 0.75:
    print('Debt-to-Assets Ratio is currently GOOD at ' + str(MostRecentLiability/MostRecentAsset))
else:
    print('Debt-to-Assets Ratio is currently BAD at ' + str(MostRecentLiability/MostRecentAsset))

FreeCashFlow = KeyRatiosData.loc["Free Cash Flow USD Mil"].str.replace(',', '').tolist()
MostRecentCashFlow = int(FreeCashFlow[-1])
LongTermDebt = BalanceSheetData.loc["Long-term debt"].tolist()
MostRecentLongTermDebt = int(LongTermDebt[-1])

if MostRecentLongTermDebt/MostRecentCashFlow <= 3:
    print('Debt Payback Time is GOOD at ' + str(MostRecentLongTermDebt/MostRecentCashFlow) + ' years')
else:
    print('Debt Payback Time is BAD at ' + str(MostRecentLongTermDebt/MostRecentCashFlow) + ' years')

ROIC = KeyRatiosData.loc["Return on Invested Capital %"].tolist()
AnnualPeriod = KeyRatiosData.loc["Profitability"].tolist()
dictionary = dict(zip(ROIC, AnnualPeriod))
for i in ROIC:
    if float(i) < 10:
        print('Return on Invested Capital is BAD at ' + str(i) + ' on ' + dictionary[i])
    else:
        print('Return on Invested Capital is GOOD at ' + str(i) + ' on ' + dictionary[i])

print('Do not forget! If the Price/Earnings ratio is less than 15, you are getting a deal!')

AnnualPeriod = KeyRatiosData.loc["Profitability"].tolist()
EPS = KeyRatiosData.loc["Earnings Per Share USD"].tolist()

left = [i+1 for i in range(len(AnnualPeriod))]
height = EPS
tick_label = AnnualPeriod
plt.bar(left, height, tick_label=tick_label, width=0.8, color=['orange'])
plt.xlabel('Annual Period')
plt.ylabel('Earnings Per Share (EPS)')
plt.title('Earnings Per Share')
plt.show()

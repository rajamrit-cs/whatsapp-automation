import pyautogui as pg
import webbrowser as web
import time
import pandas as pd
import os

# os.chdir("/home/amrit/Downloads/")
# print(os.getcwd())
df = pd.read_excel("/home/amrit/Downloads/sample.xlsx", sheet_name=None)
name = []
mobile = []
message = '''Name is here on the behalf of Bajaj Finance pesabazar✨ Earlier I had called you from Paisa Baza💥r at that time you could not pick up the phone because of this I am sending you some proposals. Overdraft Facility⬇️ Overdraft UNSECURE ONLY TWO NBFC PROVIDEING 📈➕
TATACAPITAL*BAJAJ FINANCE📊THAT IS UNSECURE  & NON EMI RESERVE FUND FOR FUTURE*
If get your max amount eligibility behalf of your these details we got soft approval (calculation) of 8L-10Lin  OD Limit*Right now we are offering you Flexi Loan on the behalf your * CIBIL SCORE

Sir Our Normal  Rate or charges are...

ROI - 14.75 % p.a.
Processing Fee - 2%
Flexi Fee - Rs. 4999/-CURRENTLY RUNNING OPEN MARKET OF CHARGES ,, bajaj finance over draft 
ROI--14.25% on reducing 
PF--1% negotiable +
Flexi fees --4999 >>ON 15LAC *💸AMOUNT,,,,But Sir Right now we are offering You

If you go for⏩ 10 Lacs or  above & Your Company  listed in Super Cat A or A
& credit report is 700above 


Rate of Interest - 14.5% p.a.

PF - 1% (of sanctioned amount)
Flexi Fee - Rs 4999/-
(Both are one time charges in 7 years)

Below 10L

Rate of Interest - 14.75% p.a.

PF - 1.50%
Flexi Fee - Rs 5999/- 
(Both are one time charges in 7 years)

PF & Flexi Fee are negotiable according to your profile & No Hidden ChargesFOR *PERSONAL LOAN  OPEN MARKET CHARGES ,,
ACORDING TO COSTUMER'S PROFILE 
CAT A》10.50%or10.75 % roi & pf >0.5%  +  insurance amount. 
CAT B》11.25%or10.99% roi & pf > 1% + insurance amount. 
CAT C》12.5%or13.5% roi & pf > 1.5% + insurance amount. 
CAT D》13.99%or14.99% roi & pf > 2% + insurance amount. 
CAT E》15.99% or 18% roi & pf > 2.5% + insurance amount.Sir If you looking for such Financial Backup Amount (Flexi Loan)

Required Documents📃📄📃

1) Latest 3 Months  PaySlips (may/june/july)

2) Updated Bank Statement last six months to- Till Date

3) Pan Card Copy

4) Pass port Size Photo

5) Current Residence Proof with Aadhar Card
6) 3 Leave Cheques at time of Agreement is mean after approval

Regards,
RAHUL KUMAR
7683017438'''

for index, row in df.items():
    name.append(row["Name"])
    mobile.append(row["Mobile"])

for i in name:
    print(i)

exit(0)
combo = zip(mobile, name)
message = 'Dear ' + str(name[0]) + ',\n' + message
combo = zip(["91_7683017438"], [message])
print(combo)

first = True
for lead, message in combo:
    time.sleep(4)
    web.open("https://web.whatsapp.com/send?phone="+lead+"&text="+message)
    if first:
        time.sleep(6)
        first = False
    width, height = pg.size()
    pg.click(width/2, height/2)
    time.sleep(4)
    pg.press('enter')
    time.sleep(8)
    pg.hotkey('ctrl', 'w')

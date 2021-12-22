import openpyxl
wb = openpyxl.load_workbook('static/uploads/test.xlsx')
ws = wb.active
val = 0
for x in (range(6,ws.max_row+1)):
    try:
        
        val = val + ws.cell(row=x, column=8).value 
    except:
        continue

val=0
AidsUSD =0
AidsLevy = 0
TaxUSD = 0
GrossIncomeUSD = 0
totalFringeB = 0
AidsZWL =0
PAYETax= 0
PAYETaxZWL=0
totalEarningsUSD=0
totalEarningsZWL = 0
cols = [6,8,9,10,11,12,13,14,15,16]
for i in cols:
    for x in (range(6,ws.max_row+1)):
        if i == 6:
            try:
                AidsUSD = AidsUSD + ws.cell(row=x, column=i).value
                
            except:
                    continue
        if i == 8:
            try:
                GrossIncomeUSD = GrossIncomeUSD + ws.cell(row=x, column=i).value
                
            except:
                    continue
        if i == 9:
            try:
                TaxUSD = TaxUSD + ws.cell(row=x, column=i).value
                
            except:
                    continue
        if i == 10:
            try:
                AidsZWL = AidsZWL + ws.cell(row=x, column=i).value
            except:
                    continue

        if i == 11:
            try:
                PAYETax = PAYETax + ws.cell(row=x, column=i).value
            except:
                    continue
        if i == 12:
            try:
                totalEarningsUSD = totalEarningsUSD + ws.cell(row=x, column=i).value
            except:
                    continue
        if i == 13:
            try:
                AidsLevy = AidsLevy + ws.cell(row=x, column=i).value
            except:
                    continue
        if i == 14:
            try:
                PAYETaxZWL = PAYETaxZWL + ws.cell(row=x, column=i).value
            except:
                    continue

        if i == 15:
            try:
                totalEarningsZWL = totalEarningsZWL + ws.cell(row=x, column=i).value
            except:
                    continue

        if i == 16:
            try:
                totalFringeB = totalFringeB + ws.cell(row=x, column=i).value
            except:
                    continue


    for x in (range(6,ws.max_row+1)):
            try:
                val = val + ws.cell(row=x, column=i).value 
            except:
                continue
          
print('TaxUSD: ',int(TaxUSD))
print('GrossIncomeUSD: ',int(GrossIncomeUSD))    
print('AidsUSD: ',int(AidsUSD))
print('AidsZWL: ',int(AidsZWL))
print('PAYETax: ',int(PAYETax))
print('totalEarningsUSD: ',int(totalEarningsUSD))
print('AidsLevy: ',int(AidsLevy))
print('PAYETaxZWL: ',int(PAYETaxZWL))
print('totalEarningsZWL: ',int(totalEarningsZWL))
print('totalFringeB: ',int(totalFringeB))
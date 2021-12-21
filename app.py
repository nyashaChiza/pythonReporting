from flask import Flask, render_template, request, url_for, session, send_from_directory
from mailmerge import MailMerge
import os
import random
import xlwings as xw
app = Flask(__name__)

def get_data(path):
    wb = xw.Book(path)
    sht = wb.sheets
    data1 = sht[0].range('F6:F500').value
    data2 = sht[0].range('I6:I500').value
    data3 = sht[0].range('J6:J500').value
    data4 = sht[0].range('K6:K500').value
    data5 = sht[0].range('L6:L500').value
    data6 = sht[0].range('M6:M500').value
    data7 = sht[0].range('N6:N500').value
    data8 = sht[0].range('O6:O500').value
    data9 = sht[0].range('P6:P500').value
    data10 = sht[0].range('H6:H500').value

    TaxUSD = 0
    AidsUSD = 0
    AidsZWL= 0
    PAYETax = 0
    totalEarningsUSD = 0
    AidsLevy = 0
    PAYETaxZWL= 0
    totalEarningsZWL = 0
    totalFringeB = 0
    ne = 0
    GrossIncomeUSD=0
    #--------------------------------------------------
    for item in data1:
            if item != None:
                TaxUSD = TaxUSD  + item

    for item in data2:
            if item != None:
                AidsUSD = AidsUSD  + item

    for item in data3:
            if item != None:
                AidsZWL = AidsZWL  + item
     
    for item in data4:
            if item != None:
                PAYETax = PAYETax  + item
            
    for item in data4:
            if item != None:
                PAYETax = PAYETax  + item
    for item in data5:
            if item != None:
                totalEarningsUSD = totalEarningsUSD  + item
    for item in data6:
            if item != None:
                AidsLevy = AidsLevy  + item
    for item in data7:
            if item != None:
                PAYETaxZWL = PAYETaxZWL  + item
    for item in data8:
            if item != None:
                totalEarningsZWL = totalEarningsZWL  + item
        
    for item in data9:
            if item != None:
                totalFringeB = totalFringeB  + item
                ne = ne + 1
    for item in data10:
            if item != None:
                GrossIncomeUSD = GrossIncomeUSD  + item

#--------------------------------------------------
    TaxUSD = TaxUSD/2
    AidsUSD = AidsUSD/2
    AidsZWL = AidsZWL/2
    totalFringeB =totalFringeB/2
    totalEarningsZWL=totalEarningsZWL/2
    AidsLevy= AidsLevy/2
    totalEarningsUSD =totalEarningsUSD/2
    PAYETax =PAYETax/2
#--------------------------------------------------
    values = {}
    values['tr1'] = int(totalEarningsUSD+totalEarningsZWL+totalFringeB)
    values['tr2'] =  int(totalEarningsZWL+totalFringeB)
    values['tr3'] = int(totalEarningsUSD)
    values['tr4'] = int(GrossIncomeUSD)
    values['ne'] = ne
    values['gp1'] = int(PAYETax+ PAYETaxZWL)
    values['gp2'] = int(PAYETaxZWL)
    values['gp3'] = int(PAYETax)
    values['gp4'] = int(TaxUSD)
    values['al1'] = int(AidsZWL+AidsLevy)
    values['al2'] = int(AidsLevy)
    values['al3'] = int(AidsZWL)
    values['al4'] = int(AidsUSD)
    values['tt1'] = int(PAYETax+PAYETaxZWL+AidsZWL+AidsLevy)
    values['tt2'] = int(PAYETaxZWL+AidsLevy)
    values['tt3'] = int(PAYETax+AidsZWL)
    values['tt4'] = int(TaxUSD+AidsUSD)
    return values

def gen_report(data1, data2):
    document = MailMerge('static/templates/temp.docx')
    document.merge(
            ename = data1.get('ename'),
            tname = data1.get('tname'),
            btnum = data1.get('btnum'),
            paye = data1.get('paye'),
            tin = data1.get('tin'),
            address = data1.get('address'),
            postal=data1.get('postal'),
            cell= data1.get('cell'),
            email= data1.get('email'),
            tax_period= data1.get('tax_period'),
            due_date= data1.get('due_date'),
            tr1 = str(data2.get('tr1')),
            tr2 = str(data2.get('tr2')),
            tr3 = str(data2.get('tr3')),
            tr4 = str(data2.get('tr4')),
            ne = str(data2.get('ne')),
            gp1 = str(data2.get('gp1')),
            gp2 = str(data2.get('gp2')),
            gp3 = str(data2.get('gp3')),
            gp4 = str(data2.get('gp4')),
            al1 = str(data2.get('al1')),
            al2 = str(data2.get('al2')),
            al3 = str(data2.get('al3')),
            al4 = str(data2.get('al4')),
            tt1 = str(data2.get('tt1')),
            tt2 = str(data2.get('tt2')),
            tt3 = str(data2.get('tt3')),
            tt4 = str(data2.get('tt4')),
        )
    doc_path = 'static/reports/'+str(data1.get('ename'))+str('.docx')
    try:
        os.remove(os.path.join(doc_path))
        document.write(doc_path)
    except:
        document.write(doc_path)
    return str(data1.get('ename'))+str('.docx')


@app.route('/', methods=['POST','GET'])
def index():
    if request.files and request.method == 'POST' :
        file = request.files['file']
        try:
            os.remove(os.path.join("static/uploads/", file.filename))
            file.save(os.path.join("static/uploads/", file.filename))
        except:
            file.save(os.path.join("static/uploads/", file.filename))
        data = request.form.to_dict()
        doc_path =  gen_report(data, get_data(os.path.join("static/uploads/", file.filename)))
        return reports()
    template = 'index.html'
    return render_template(template)

@app.route("/reports")
def reports():
    onlyfiles = [f for f in os.listdir('static/reports')]
    template = 'reports.html'
    return render_template(template,files=onlyfiles)


@app.route("/download/<string:name>")
def download(name):
    download_file = name
    print('downloading ',name)
    return send_from_directory(directory='static/reports',filename=download_file) 


if __name__ == '__main__':
    app.run(debug=True, port=5000)

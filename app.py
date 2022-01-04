from flask import Flask, render_template, request, url_for, session, send_from_directory
from mailmerge import MailMerge
import os
import random
import openpyxl
app = Flask(__name__)

def get_data(path):
    wb = openpyxl.load_workbook(path)
    ws = wb.active
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
#--------------------------------------------------
#--------------------------------------------------
 #   TaxUSD = TaxUSD/2
  #  TaxUSD = TaxUSD/2
   # AidsUSD = AidsUSD/2
    #AidsZWL = AidsZWL/2
    #totalFringeB =totalFringeB/2
    #totalEarningsZWL=totalEarningsZWL/2
    #AidsLevy= AidsLevy/2
    #totalEarningsUSD =totalEarningsUSD/2
    #PAYETax =PAYETax/2
#--------------------------------------------------
    values = {}
    values['tr1'] = '{:.2f}'.format((totalEarningsUSD+totalEarningsZWL+totalFringeB)/2)
    values['tr2'] =  '{:.2f}'.format((totalEarningsZWL+totalFringeB)/2)
    values['tr3'] = '{:.2f}'.format((totalEarningsUSD)/2)
    values['tr4'] = '{:.2f}'.format((GrossIncomeUSD)/2)
    values['ne'] = ws.max_row-6
    values['gp1'] = '{:.2f}'.format((PAYETax+ PAYETaxZWL)/2)
    values['gp2'] = '{:.2f}'.format((PAYETaxZWL)/2)
    values['gp3'] = '{:.2f}'.format((PAYETax)/2)
    values['gp4'] = '{:.2f}'.format((TaxUSD)/2)
    values['al1'] = '{:.2f}'.format((AidsZWL+AidsLevy)/2)
    values['al2'] = '{:.2f}'.format((AidsLevy)/2)
    values['al3'] = '{:.2f}'.format((AidsZWL)/2)
    values['al4'] = '{:.2f}'.format((AidsUSD)/2)
    values['tt1'] = '{:.2f}'.format((PAYETax+PAYETaxZWL+AidsZWL+AidsLevy)/2)
    values['tt2'] = '{:.2f}'.format((PAYETaxZWL+AidsLevy)/2)
    values['tt3'] = '{:.2f}'.format((PAYETax+AidsZWL)/2)
    values['tt4'] = '{:.2f}'.format((TaxUSD+AidsUSD)/2)
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
            rate= data1.get('rate'),
            region= data1.get('region'),
            station= data1.get('station'),
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
#-------------------------------------------------------------
@app.route('/', methods=['POST','GET'])
def index():
    error = 'none'
    if request.method == 'POST':
        email = request.form.get('email')
        password = request.form.get('password')
        if email =='client@petalmafrica.com' and password == 'client@123':
            return home()
        else:
            error = 'invalid log in details'
    template = 'sign-in.html'
    return render_template(template, error = error)

@app.route('/home/', methods=['POST','GET'])
def home():
    if request.files and request.method == 'POST' :
        file = request.files['file']
        try:
            os.remove(os.path.join("static/uploads/", file.filename))
            file.save(os.path.join("static/uploads/", file.filename))
        except:
            file.save(os.path.join("static/uploads/", file.filename))
        data = request.form.to_dict()
        try:
            gen_report(data, get_data(os.path.join("static/uploads/", file.filename)))
            return reports('Document was saved successfully')
        except Exception as e:
            print(e)
            return reports('failed to save document')
    template = 'index.html'
    return render_template(template)

@app.route("/reports/<string:error>/")
def reports(error):
    onlyfiles = [f for f in os.listdir('static/reports')]
    template = 'reports.html'
    return render_template(template,files=onlyfiles, error = error)


@app.route("/download/<string:name>")
def download(name):
    download_file = name
    print('downloading ',name)
    try:
        return send_from_directory(directory='static/reports',filename=download_file) 
    except Exception as e:
        print(e)
        return reports('failed to download file')
        

@app.route("/delete/<string:name>")
def delete(name):
    error = 'Document Deleted Successfully'
    try:
        os.remove('static/reports/'+name)
    except Exception as e:
        print(e)
        error = 'Failed to delete File'
    print('deleting ',name)
    return reports(error)

@app.route("/logout/")
def logout():
    #session['user'] = None
    return index() 


if __name__ == '__main__':
    app.run(debug=True, port=5000)

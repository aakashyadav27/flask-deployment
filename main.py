from flask import Flask,redirect,request,url_for,jsonify, render_template
import requests
import pandas as pd
import re
import string
from collections import Counter
from flask import Flask
import xlwt
import re
import os
from datetime import datetime
cwd = os.getcwd()

app = Flask(__name__)

@app.route("/",methods=['POST','GET'])
def hello_world():
        return render_template('index.html')

@app.route('/test', methods=['GET', 'POST'])
def connect_management():
    user = request.form.get('Projects')
    with requests.session() as s:
        url = "http://plwbox.i2econsulting.com/app/plw/127.0.0.1:8100/home/Workbox/Workbox"
        r = s.get(url, headers=headers)
        r = s.post(url, data=login_data, headers=headers)
        url1 = "http://plwbox.i2econsulting.com/sat0/OPX2/127.0.0.1:8400/odata/startsession()"
        r = s.get(url1)
        applet = "applet(" + str(r.content).split("\\")[1] + "')"
        project_name = '"' + user + '"'
        Odata_url = "http://plwbox.i2econsulting.com/sat0/OPX2/127.0.0.1:8400/odata/applet('9430697dad94f4af71a8e41e852f1ef9')/ACTIVITY?$filter=PROJECT eq {}&$select=NAME,PROJECT,_PM_DA_S_LINE_ID,_PM_DA_S_MSP_PREDECESSORS,_PM_DA_S_MSP_SUCCESSORS,_pm_sf_sync_ps_target,_pm_sf_sync_pf_target".format(
            project_name)
        Odata_url = Odata_url.replace(Odata_url.split("/")[7].strip(), applet)
        r = s.get(Odata_url)
        data = r.json()
        Odata_url = Odata_url.replace(Odata_url.split("/")[8].strip(), 'stopsession()')
        r = s.get(Odata_url)
    df = pd.DataFrame(data['value'])
    df.columns = ['id', 'Task', 'project', 'Line_indetifier',
                  'predecessors', 'successors',
                  'start_date', 'end_date']
    """df['start_date'] = df['start_date'].apply(lambda x: datetimeConversion(x))
    df['end_date'] = df['end_date'].apply(lambda x: datetimeConversion(x))"""
    df.to_excel('{}.xls'.format(user))
    return "You are file has been saved at your local file system location:{}".format(cwd)

if __name__=='__main__':
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/95.0.4638.69 Safari/537.36",
        "Connection": "keep-alive",
        "Cache-Control": "max-age=0",
        "Upgrade-Insecure-Requests": "1",
        "Origin": "http://plwbox.i2econsulting.com",
        "Content-Type": "application/x-www-form-urlencoded",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9",
        "Accept-Language": "en-GB,en-US;q=0.9,en;q=0.8,hi;q=0.7"
    }
    # login data
    login_data = {
        "httpd_username": "intranet",
        "httpd_password": "admin@123",
        "login": ""
    }
    def datetimeConversion(d):
        d = re.split('\(|\)', d)[1][:10]
        d = int(d[:10])
        return datetime.fromtimestamp(d).strftime('%Y-%m-%d')

    app.run(debug=True)
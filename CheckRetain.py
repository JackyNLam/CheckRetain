from selenium import webdriver
import xmltodict,json,time,subprocess,sys,tkinter
from datetime import datetime,timedelta
from win10toast import ToastNotifier

def getWallChartXML(browser):
    try:
        ChartXML = browser.execute_script('return fetch("https://retain-asiapac.ey.net/IWRetainWeb/IWISAPIRedirect.dll/rwpajax/getWallchartGroupData?cmpid=WEBPLANNERJOBS&maxres=12&isres=1&sdate=' +LastSat+ '&edate=' +NextSat+ '&nameindex=0&inctimes=1&isprojview=1&list=resListSingleId&sellist=resList&refreshlist=1&singleid=' +PersonID+ '&scrolltoid=NaN&expandscrollid=0&projnameindex=0&projlevelindex=0&nameid=-1&isforecast=0&expids=&req_date=Tue%20Jun%2022%202021%2014%3A12%3A05%20GMT%2B0800%20(Hong%20Kong%20Standard%20Time)&rtn_http_method=get", {"body":"extquery=&namesparent=&selquery=&bkgquery=&colsch=Default&extfields=<extfields><field>RES_DESCR</field><field>RES_BUU_ID</field><field>RES_SMU_ID</field></extfields>&extbkgfields=&bkgbarfields=%3Cbkg_bar%3E%3Cfield%3EBKG_JOB_ID%3C%2Ffield%3E%3Cfield%3EBkgTmeHrs%3C%2Ffield%3E%3C%2Fbkg_bar%3E&bkgseltables=%3Ctables%3E%3C%2Ftables%3E&sortorder=<xml><field><name>RES_DESCR</name><order>0</order></field></xml>&resetsort=0&showbrq=0&showbkg=1&showghost=0&showtimesheet=0&tblname=RES&bkgtable=BKG&namesbkgtable=&filtermasterrec=0&viewtype=month&startgrouppos=undefined&scrollindex=0&newscrollindex=undefined&forecastcols=0&projnamexml=&expandedgroup=null&groups=<groups><group><table>RES</table><field>RES_CUR_GRD_ID</field><isid>true</isid></group></groups>&expxml==","method": "POST",}).then(response=>response.text()).then(data=>data);')
        return ChartXML
    except:
        time.sleep(0.5)
        return getWallChartXML(browser)

def getoutput(searchfor,wincommand):
    tempText = str(subprocess.check_output(wincommand,shell=True, stderr=subprocess.PIPE, stdin=subprocess.PIPE))
    return (searchfor in tempText)
    
if getoutput("WiFiFirst","netsh wlan show interfaces") or getoutput("EY Remote Connect","rasdial"):
    pass
else:
    sys.exit()
LastSat=datetime.strftime(datetime.today().date()-timedelta(days=(datetime.today().weekday()+2)),"%d/%m/%Y")
NextSat=datetime.strftime(datetime.strptime(LastSat,"%d/%m/%Y")+timedelta(days=365),"%d/%m/%Y")
PersonID = "1349716"
browser = webdriver.Edge(r"C:\Users\VR787FC\msedgedriver.exe")
browser.minimize_window()
browser.get('https://retain-asiapac.ey.net/IWRetainWeb/IWISAPIRedirect.dll/Files/static/html/webportal_6_18.html')
ChartXML = getWallChartXML(browser)
browser.quit()
Chartdic= xmltodict.parse(ChartXML)
try:
    with open("RetainJSON",'r') as f:
        ChartdicOLD = json.loads(f.read())
    if ChartdicOLD == Chartdic: 
        with open("RetainDiff.txt",'a') as ftwo:
             ftwo.write("\n{}\nNo update from retain.".format(datetime.now()))               
    else:
        with open("RetainDiff.txt",'a') as ftwo:
            ftwo.write("\n{}\n{}".format(datetime.now(),[x for x in ChartdicOLD["result"]["bkgs"]["bkg"] if x not in Chartdic["result"]["bkgs"]["bkg"]]+[x for x in Chartdic["result"]["bkgs"]["bkg"] if x not in ChartdicOLD["result"]["bkgs"]["bkg"]]))
        with open("RetainJSON",'w') as f:
            f.write(json.dumps(Chartdic))

        toaster = ToastNotifier()
        toaster.show_toast("Retain Notification","Retain is updated!!!",icon_path="python.ico",duration=2,threaded=True)
        time.sleep(2.1)
        toaster.show_toast("Retain Notification","Retain is updated!!!",icon_path="python.ico",duration=3,threaded=True)
        time.sleep(3.1)
        toaster.show_toast("Retain Notification","Retain is updated!!!",icon_path="python.ico",duration=5,threaded=True)
        root = tkinter.Tk()
        root.geometry('280x120')
        root.title("Retain Notification")
        root['background']='#dce1fe'
        tkinter.Label(root,text="Retain is updated!",font=(None,16),bg="#dce1fe").pack(pady=25)
        tkinter.Button(root, text ="Got It", command = root.destroy,padx=30,bg='brown',fg='white').pack()
        root.mainloop()
except Exception as e:
    print(e)
    with open("RetainJSON",'w') as f:
        f.write(json.dumps(Chartdic))
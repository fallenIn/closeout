# -*- coding: utf-8 -*-
import smtplib
import datetime
import csv,os,sys
#import pymssql
import uuid
#import _mssql
import decimal
import tkinter as tk
import win32
from tkinter import ttk
from tkinter import *
from tkinter import scrolledtext
import tkinter.filedialog
from tkinter.filedialog import askdirectory
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
decimal.__version__
uuid.ctypes.__version__
#_mssql.__version__


win = tk.Tk()
win.title("Closeout Mail Tool")    # 添加标题
#win.withdraw()
#fpath = filedialog.askopenfilename()

ttk.Label(win, text="Chooes a Upper").grid(column=1, row=0)    # 添加一个标签，并将其列设置为1，行设置为0
ttk.Label(win, text="Enter a Date:").grid(column=0, row=0)      # 设置其在界面中出现的位置  column代表列   row 代表行
#ttk.Label(win, text="目标文件路径:").grid(column=0, row=2)       #选择文件路径
tex = scrolledtext.ScrolledText(win, width=40, height=5,font=("微软雅黑",8))#添加滚动文本框
tex.place(x=10, y=80)
#global row1,row
def selectPath():
    global filepath,filepath
    path_ = tkinter.filedialog.askopenfilename()
    #paths_= askdirectory()
    path.set(path_)
    #path.set(paths_)
    filepath = path_
    #filepaths = paths_
    tex.insert(END,filepath+'\n')
    #print(filepath获取文件路径)

def clickMe():
    global Newdir,attname,count,name,a,b,c,d,Markertmapping,Codemapping,row1,row
    months=[]
    a=[]
    b=[]
    c=[]
    d=[]
    row=[]
    # 当acction被点击时,该函数则生效
    Fdate = name.get()   # 设置button显示的内容
    FUpper = number.get()
    print(FUpper)
    action.configure(state='active')
    count = Fdate
    #print(filepath)
    if FUpper == 'PHL':
        months = ['','F','G','H','J','K','M','N','Q','U','V','X','Z']
        a = ['APEX','BMD','HKEX','CBOT','CBOE','CME','COMEX','EUREX','ICUS','LIFFE','LME','NYMEX','SGX','TOCOM']
        b = ['APEX','BMD','HKFE','CBOT','CBOE','CME','COMEX','EUREX','NYBOT','IPE','LME','CME','SGXQ','TOCOM']
        c = ['YW','YK','YC','UB','TN','EH','C','XBT','VX','UPO','POL','MHI','PF','UC', 'KW', 'O', 'S', 'W', 'YM', 'ZB', 'ZF', 'ZL', 'ZM', 'ZN', 'ZR', 'ZT', 'ZQ','OC', 'OS', 'OW', 'OZL', 'OZM', 'AD', 'BP', 'CD', 'EC', 'ED', 'ES', 'FC', 'JY', 'LB', 'LC', 'LN', 'M6A', 'M6E', 'MP', 'NE', 'NIY', 'NKD', 'NQ', 'SF', 'OAD', 'OBP', 'OCD', 'OEC', 'OES', 'OJY', 'ONQ', 'GC', 'HG', 'SI', 'OGC', 'DAX', 'ESX', 'GBL', 'GBM', 'GBS', 'CC', 'CT', 'DX', 'KC', 'NTF', 'OJ', 'SB11', 'YG', 'YI', 'LW', 'RC', 'Z', 'AH', 'CA', 'NI', 'PB', 'SN', 'ZS', 'CL', 'HO', 'NG', 'PA', 'PL', 'QM', 'RB', 'TIO', 'OCL', 'ONG', 'CN', 'FE', 'JB', 'MSG', 'NK', 'NU', 'TF', 'CH', 'JRU', 'STF', 'SFE']
        d = ['YW','XB','YC','UBE','TN','EH','C','XBT','VIX','FUPO','FPOL','MHI','PF','QUSDCNY','KCBT WHEAT', 'O', 'S', 'W', 'YM', 'US', 'FV', 'BO', 'SM', 'TY', 'RR', 'TU','ZQ', 'CORN O', 'SOYBEAN O', 'WHEAT O', 'SOYOIL O', 'SOYMEAL O', 'AUD', 'GBP', 'CAN ', 'Euro FX', 'GLBX EURO', 'MINI S&P', 'F CATTLE', 'JPY', 'RL LUMBER', 'L CATTLE', 'LEAN HOG', 'MIC AUDUSD', 'MIC EURUSD', 'MXP', 'NZD', 'NIY', 'NKD', 'MINI NSDQ', 'SWF', 'AUD FX O', 'GBP FX O', 'CAD FX O', 'EUR FX O', 'MINI S&P O', 'JPF FX O', 'MINI NQ O', 'CMX GLD', 'CMX COP', 'CMX SIL', 'CMX GLD  O', 'DAX', 'DJEST50', 'BUND', 'BOBL', 'SHAZ', 'NB COCO', 'NB COTT', 'NB DLR', 'NB COFY', 'MINIRL2000', 'NB FCOJ', 'NB SU11', 'MINIGOLD', 'MINISILVER', 'SUGAR', 'R COFFEE', 'FTSE100', 'AH', 'CA', 'NI', 'PB', 'SN', 'ZS', 'CRUDE', 'HEATOIL', 'NATGAS', 'NYM PAL', 'NYM PLA', 'MINICRUDE', 'GASOLINE', 'IRONORE', 'CRUDE O', 'NATGAS O', 'QXINHUA 50', 'IRON ORE', 'QJG10YR', 'QSG', 'QNK', 'QNU', 'TSR20', 'QCH', 'RUBBER', 'TSR20 SP', 'IRONORE SP']

        Markertmapping = {'CBOT':'CME_CBT','CME':'CME','COMEX':'CME','EUREX':'EUREX','ICUS':'NYBOT','LIFFE':'IPE','LIFFE1':'ICE_UK','LME':'LME','NYMEX':'CME','SGX':'SGXQ','TOCOM':'TOCOM'}
        Codemapping = {'FC': 'F CATTLE', 'NTF': 'MINIRL2000', 'OBP': 'GBP FX O', 'OCD': 'CAD FX O', 'GC': 'CMX GLD', 'NI': 'NI', 'CD': 'CAN ', 'SF': 'SWF', 'OJ': 'NB FCOJ', 'YG': 'MINIGOLD', 'NKD': 'NKD', 'STF': 'TSR20 SP', 'PL': 'NYM PLA', 'CL': 'CRUDE', 'YM': 'DJIA5', 'CN': 'QXINHUA 50', 'SB11': 'NB SU11', 'OW': 'WHEAT O', 'JY': 'JPY', 'RC': 'R COFFEE', 'OZL': 'SOYOIL O', 'TIO': 'IRONORE', 'ZN': '10Y TN', 'SN': 'SN', 'KC': 'NB COFY', 'LN': 'LEAN HOG', 'OS': 'SOYBEAN O', 'OAD': 'AUD FX O', 'YI': 'MINISILVER', 'OGC': 'CMX GLD  O', 'O': 'OATS', 'DX': 'NB DLR', 'OEC': 'EUR FX O', 'W': 'WHEAT', 'OES': 'MINI S&P O', 'RB': 'GASOLINE', 'LW': 'SUGAR', 'CH': 'QCH', 'ZL': 'SOYOIL', 'DAX': 'DAX', 'ZF': '5Y TN', 'OC': 'CORN O', 'MSG': 'QSG', 'AH': 'AH', 'OCL': 'CRUDE O', 'PB': 'PB', 'M6A': 'MIC AUDUSD', 'MP': 'MXP', 'C': 'CORN', 'CC': 'NB COCO', 'ES': 'MINI S&P', 'GBS': 'SHAZ', 'ED': 'GLBX EURO', 'ZM': 'SOYMEAL', 'ZT': '2Y TN', 'HG': 'CMX COP', 'GBL': 'BUND', 'PA': 'NYM PAL', 'NU': 'QNU', 'JRU': 'RUBBER', 'ONQ': 'MINI NQ O', 'NG': 'NATGAS', 'EC': 'Euro FX', 'Z': 'FTSE100', 'AD': 'AUD', 'OZM': 'SOYMEAL O', 'ZB': '30Y TB', 'FE': 'IRON ORE', 'GBM': 'BOBL', 'LB': 'RL LUMBER', 'KW': 'KCBT WHEAT', 'S': 'SOYBEAN', 'ZR': 'RICE', 'OJY': 'JPF FX O', 'NIY': 'NIY', 'CT': 'NB COTT', 'LC': 'L CATTLE', 'HO': 'HEATOIL', 'SI': 'CMX SIL', 'JB': 'QJG10YR', 'BP': 'GBP', 'NK': 'QNK', 'ZS': 'ZS', 'NQ': 'MINI NSDQ', 'QM': 'MINICRUDE', 'NE': 'NZD', 'SFE': 'IRONORE SP', 'ESX': 'DJEST50', 'M6E': 'MIC EURUSD', 'TF': 'TSR20', 'ONG': 'NATGAS O', 'CA': 'CA'}
        def tr_s(s):
            return months[int(s[2:])]+'20'+s[:2]

        def tr_s1(s):
            return months[int(s[2:5])]+'20'+s[0:2]
        def tr_s2(s):
            return s[-7:-2]
        def tr_s3(s):
            return s[-1]
        row = []
        try:
            with open(filepath,'r',encoding='utf-8') as csvfile:
                reader = csv.reader(csvfile)
                for line in reader:
                    if line[0] == 'PHLSG':
                        row.append(line)
                    if line[0] == 'PHLSGCQG':
                        row.append(line)
                    if line[0] == 'PHLSGCQG02':
                        row.append(line)
                    else:
                        pass
                        #print(row2)
        except Exception:
            sys.stderr.write('读取文件发生IO异常！\n')
        finally:
            csvfile.close()
            sys.stderr.write('finnaly读取文件成功！\n')
        print(row)
        tex.insert(END,row)
        for row1 in row:
            if row1[0] == 'PHLSG':
                row1[0] = 'CFT8000'
            if row1[0] == 'PHLSGCQG':
                row1[0] = 'CFT8001'
            if row1[0] == 'PHLSGCQG02':
                row1[0] = 'CFT8002'
            else:
                pass
            #row1[0] = 'CFT8000'
            row1.insert(1,'F')
            row1.append('')
            row1.append('')
            row1.append('')
            row1[6] = ''
            row1[7] = ''
            row1[8] = ''
            row1[10] = row1[5]
            row1[5] = ''
            #print(row1)
            for j in row1:
                if j == row1[2]:
                    if row1[2] in a:#市场code转换
                        #print(row1[2])
                        if row1[3] in ["FTSE100"]:
                            row1[2]= "ICE_UK"
                        if row1[2] == "LME":
                            pass
                        else:
                            t = a.index(row1[2])
                            row1[2]= b[t]

                    else:
                        pass
                else:
                    pass
                if j == row1[3]:
                    if row1[3] in c:#产品code转换
                        t2 = c.index(row1[3])
                        row1[3] = d[t2]
                if len(row1[4]) > 6:#期权处理
                    p1 = tr_s1(row1[4])
                    p5 = tr_s2(row1[4])
                    row1[5] = p5
                    p6 = tr_s3(row1[4])
                    row1[6] = p6
                    row1[4] = p1
                    row1[1] = "O"
                elif len(row1[4]) == 4:
                    p1 = tr_s(row1[4])
                    row1[4] = p1
            print(row1)
            tex.insert(END,row1)
        print(row)
        sFilename = os.path.join(os.path.dirname(filepath),'closeout_PHL.csv')#创建文件路径

        eFile = open(sFilename,'w',newline='')

        eWriter = csv.writer(eFile,delimiter=',',lineterminator='\r\n')
        eWriter.writerow(['client_no','Com_Type','Exch_cd','Com_cd','Contract_Month','Strike_Price','Call_Put','Val_Date','Trade_Date_Buy','Trade_Date_Sell','Traded_Qty','Traded_Price_Buy','Traded_Price_Sell','Traded_Premium_Buy','Traded_Premium_Sell'])
        for clr in row:
            eWriter.writerow(clr)
        eFile.close()


    if FUpper == 'EDFMAN':
        months = ['','F','G','H','J','K','M','N','Q','U','V','X','Z']
        a = ['APEX','BMD','HKEX','CBOT','CBOE','CME','COMEX','EUREX','ICUS','LIFFE','LME','NYMEX','SGX','TOCOM']
        b = ['APEX','BMD','HKFE','CBOT','CBOE','CME','COMEX','EUREX','NYBOT','IPE','LME','CME','SGXQ','TOCOM']
        c = ['YW','YK','YC','UB','TN','EH','C','XBT','VX','UPO','POL','MHI','PF','UC', 'KW', 'O', 'S', 'W', 'YM', 'ZB', 'ZF', 'ZL', 'ZM', 'ZN', 'ZR', 'ZT', 'ZQ','OC', 'OS', 'OW', 'OZL', 'OZM', 'AD', 'BP', 'CD', 'EC', 'ED', 'ES', 'FC', 'JY', 'LB', 'LC', 'LN', 'M6A', 'M6E', 'MP', 'NE', 'NIY', 'NKD', 'NQ', 'SF', 'OAD', 'OBP', 'OCD', 'OEC', 'OES', 'OJY', 'ONQ', 'GC', 'HG', 'SI', 'OGC', 'DAX', 'ESX', 'GBL', 'GBM', 'GBS', 'CC', 'CT', 'DX', 'KC', 'NTF', 'OJ', 'SB11', 'YG', 'YI', 'LW', 'RC', 'Z', 'AH', 'CA', 'NI', 'PB', 'SN', 'ZS', 'CL', 'HO', 'NG', 'PA', 'PL', 'QM', 'RB', 'TIO', 'OCL', 'ONG', 'CN', 'FE', 'JB', 'MSG', 'NK', 'NU', 'TF', 'CH', 'JRU', 'STF', 'SFE']
        d = ['YW','XB','YC','UBE','TN','EH','C','XBT','VIX','FUPO','FPOL','MHI','PF','QUSDCNY','KCBT WHEAT', 'O', 'S', 'W', 'YM', 'US', 'FV', 'BO', 'SM', 'TY', 'RR', 'TU','ZQ', 'CORN O', 'SOYBEAN O', 'WHEAT O', 'SOYOIL O', 'SOYMEAL O', 'AUD', 'GBP', 'CAN ', 'Euro FX', 'GLBX EURO', 'MINI S&P', 'F CATTLE', 'JPY', 'RL LUMBER', 'L CATTLE', 'LEAN HOG', 'MIC AUDUSD', 'MIC EURUSD', 'MXP', 'NZD', 'NIY', 'NKD', 'MINI NSDQ', 'SWF', 'AUD FX O', 'GBP FX O', 'CAD FX O', 'EUR FX O', 'MINI S&P O', 'JPF FX O', 'MINI NQ O', 'CMX GLD', 'CMX COP', 'CMX SIL', 'CMX GLD  O', 'DAX', 'DJEST50', 'BUND', 'BOBL', 'SHAZ', 'NB COCO', 'NB COTT', 'NB DLR', 'NB COFY', 'MINIRL2000', 'NB FCOJ', 'NB SU11', 'MINIGOLD', 'MINISILVER', 'SUGAR', 'R COFFEE', 'FTSE100', 'AH', 'CA', 'NI', 'PB', 'SN', 'ZS', 'CRUDE', 'HEATOIL', 'NATGAS', 'NYM PAL', 'NYM PLA', 'MINICRUDE', 'GASOLINE', 'IRONORE', 'CRUDE O', 'NATGAS O', 'QXINHUA 50', 'IRON ORE', 'QJG10YR', 'QSG', 'QNK', 'QNU', 'TSR20', 'QCH', 'RUBBER', 'TSR20 SP', 'IRONORE SP']

        Markertmapping = {'CBOT':'CME_CBT','CME':'CME','COMEX':'CME','EUREX':'EUREX','ICUS':'NYBOT','LIFFE':'IPE','LIFFE1':'ICE_UK','LME':'LME','NYMEX':'CME','SGX':'SGXQ','TOCOM':'TOCOM'}
        Codemapping = {'FC': 'F CATTLE', 'NTF': 'MINIRL2000', 'OBP': 'GBP FX O', 'OCD': 'CAD FX O', 'GC': 'CMX GLD', 'NI': 'NI', 'CD': 'CAN ', 'SF': 'SWF', 'OJ': 'NB FCOJ', 'YG': 'MINIGOLD', 'NKD': 'NKD', 'STF': 'TSR20 SP', 'PL': 'NYM PLA', 'CL': 'CRUDE', 'YM': 'DJIA5', 'CN': 'QXINHUA 50', 'SB11': 'NB SU11', 'OW': 'WHEAT O', 'JY': 'JPY', 'RC': 'R COFFEE', 'OZL': 'SOYOIL O', 'TIO': 'IRONORE', 'ZN': '10Y TN', 'SN': 'SN', 'KC': 'NB COFY', 'LN': 'LEAN HOG', 'OS': 'SOYBEAN O', 'OAD': 'AUD FX O', 'YI': 'MINISILVER', 'OGC': 'CMX GLD  O', 'O': 'OATS', 'DX': 'NB DLR', 'OEC': 'EUR FX O', 'W': 'WHEAT', 'OES': 'MINI S&P O', 'RB': 'GASOLINE', 'LW': 'SUGAR', 'CH': 'QCH', 'ZL': 'SOYOIL', 'DAX': 'DAX', 'ZF': '5Y TN', 'OC': 'CORN O', 'MSG': 'QSG', 'AH': 'AH', 'OCL': 'CRUDE O', 'PB': 'PB', 'M6A': 'MIC AUDUSD', 'MP': 'MXP', 'C': 'CORN', 'CC': 'NB COCO', 'ES': 'MINI S&P', 'GBS': 'SHAZ', 'ED': 'GLBX EURO', 'ZM': 'SOYMEAL', 'ZT': '2Y TN', 'HG': 'CMX COP', 'GBL': 'BUND', 'PA': 'NYM PAL', 'NU': 'QNU', 'JRU': 'RUBBER', 'ONQ': 'MINI NQ O', 'NG': 'NATGAS', 'EC': 'Euro FX', 'Z': 'FTSE100', 'AD': 'AUD', 'OZM': 'SOYMEAL O', 'ZB': '30Y TB', 'FE': 'IRON ORE', 'GBM': 'BOBL', 'LB': 'RL LUMBER', 'KW': 'KCBT WHEAT', 'S': 'SOYBEAN', 'ZR': 'RICE', 'OJY': 'JPF FX O', 'NIY': 'NIY', 'CT': 'NB COTT', 'LC': 'L CATTLE', 'HO': 'HEATOIL', 'SI': 'CMX SIL', 'JB': 'QJG10YR', 'BP': 'GBP', 'NK': 'QNK', 'ZS': 'ZS', 'NQ': 'MINI NSDQ', 'QM': 'MINICRUDE', 'NE': 'NZD', 'SFE': 'IRONORE SP', 'ESX': 'DJEST50', 'M6E': 'MIC EURUSD', 'TF': 'TSR20', 'ONG': 'NATGAS O', 'CA': 'CA'}

        def tr_s(s):
            return months[int(s[2:])]+'20'+s[:2]

        def tr_s1(s):
            return months[int(s[2:5])]+'20'+s[0:2]
        def tr_s2(s):
            return s[-7:-2]
        def tr_s3(s):
            return s[-1]
        row = []
        try:
            with open(filepath,'r',encoding='utf-8') as csvfile:
                reader = csv.reader(csvfile)
                for line in reader:
                    if line[0] == 'EDFMAN':
                        row.append(line)
                    if line[0] == 'EDFMANCME':
                        row.append(line)
                    else:
                        pass
                        #print(row2)
        except Exception:
            sys.stderr.write('读取文件发生IO异常！\n')
        finally:
            csvfile.close()
            sys.stderr.write('finnaly读取文件成功！\n')
        #print(row)
        tex.insert(END,row)

        for row1 in row:
            if row1[0] == 'EDFMANCME':
                row1[0] = 'UHK50010'
            if row1[0] == 'EDFMAN' and row1[1] == 'ICEU':
                row1[0] = 'UHK50010'
            if row1[0] == 'EDFMAN' and row1[1] == 'ICUS':
                row1[0] = 'UHK50010'
            if row1[0] == 'EDFMAN' and row1[1] == 'LME':
                row1[0] = 'LHLL0656'
            if row1[0] == 'EDFMAN' and row1[1] == 'EURONEXT':
                row1[0] = 'LHLL0656'
            if row1[0] == 'EDFMAN' and row1[1] == 'EUREX':
                row1[0] = 'LHLL0656'
            else:
                pass
            row1.insert(1,'F')
            row1.append('')
            row1.append('')
            row1.append('')
            row1[6] = ''
            row1[7] = ''
            row1[8] = ''
            row1[10] = row1[5]
            row1[5] = ''
            #print(row1)
            for j in row1:
                if j == row1[2]:
                    if row1[2] in a:#市场code转换
                        #print(row1[2])
                        if row1[3] in ["FTSE100"]:
                            row1[2]= "ICE_UK"
                        if row1[2] == "LME":
                            pass
                        else:
                            t = a.index(row1[2])
                            row1[2]= b[t]

                    else:
                        pass
                else:
                    pass
                if j == row1[3]:
                    if row1[3] in c:#产品code转换
                        t2 = c.index(row1[3])
                        row1[3] = d[t2]
                if len(row1[4]) > 6:#期权处理
                    p1 = tr_s1(row1[4])
                    p5 = tr_s2(row1[4])
                    row1[5] = p5
                    p6 = tr_s3(row1[4])
                    row1[6] = p6
                    row1[4] = p1
                    row1[1] = "O"
                elif len(row1[4]) == 4:
                    p1 = tr_s(row1[4])
                    row1[4] = p1
            print(row1)
            tex.insert(END,row1)
        #print(row)
        sFilename = os.path.join(os.path.dirname(filepath),'closeout_EDFMAN.csv')#创建文件路径

        eFile = open(sFilename,'w',newline='')

        eWriter = csv.writer(eFile,delimiter=',',lineterminator='\r\n')
        eWriter.writerow(['client_no','Com_Type','Exch_cd','Com_cd','Contract_Month','Strike_Price','Call_Put','Val_Date','Trade_Date_Buy','Trade_Date_Sell','Traded_Qty','Traded_Price_Buy','Traded_Price_Sell','Traded_Premium_Buy','Traded_Premium_Sell'])
        for clr in row:
            eWriter.writerow(clr)
        eFile.close()

    if FUpper == 'ADM':
        Markertmapping = {'CBOT':'CME_CBT','CME':'CME','COMEX':'CME','EUREX':'EUREX','ICUS':'NYBOT','LIFFE':'IPE','LIFFE1':'ICE_UK','LME':'LME','NYMEX':'CME','SGX':'SGXQ','TOCOM':'TOCOM'}
        Codemapping = {'FC': 'F CATTLE', 'NTF': 'MINIRL2000', 'OBP': 'GBP FX O', 'OCD': 'CAD FX O', 'GC': 'CMX GLD', 'NI': 'NI', 'CD': 'CAN ', 'SF': 'SWF', 'OJ': 'NB FCOJ', 'YG': 'MINIGOLD', 'NKD': 'NKD', 'STF': 'TSR20 SP', 'PL': 'NYM PLA', 'CL': 'CRUDE', 'YM': 'DJIA5', 'CN': 'QXINHUA 50', 'SB11': 'NB SU11', 'OW': 'WHEAT O', 'JY': 'JPY', 'RC': 'R COFFEE', 'OZL': 'SOYOIL O', 'TIO': 'IRONORE', 'ZN': '10Y TN', 'SN': 'SN', 'KC': 'NB COFY', 'LN': 'CME LEAN HOGS', 'OS': 'SOYBEAN O', 'OAD': 'AUD FX O', 'YI': 'MINISILVER', 'OGC': 'CMX GLD  O', 'O': 'OATS', 'DX': 'NB DLR', 'OEC': 'EUR FX O', 'W': 'WHEAT', 'OES': 'MINI S&P O', 'RB': 'GASOLINE', 'LW': 'SUGAR', 'CH': 'QCH', 'ZL': 'SOYOIL', 'DAX': 'DAX', 'ZF': '5Y TN', 'OC': 'CORN O', 'MSG': 'QSG', 'AH': 'AH', 'OCL': 'CRUDE O', 'PB': 'PB', 'M6A': 'MIC AUDUSD', 'MP': 'MXP', 'C': 'CORN', 'CC': 'NB COCO', 'ES': 'MINI S&P', 'GBS': 'SHAZ', 'ED': 'GLBX EURO', 'ZM': 'SOYMEAL', 'ZT': '2Y TN', 'HG': 'CMX COP', 'GBL': 'BUND', 'PA': 'NYM PAL', 'NU': 'QNU', 'JRU': 'RUBBER', 'ONQ': 'MINI NQ O', 'NG': 'NATGAS', 'EC': 'Euro FX', 'Z': 'FTSE100', 'AD': 'AUD', 'OZM': 'SOYMEAL O', 'ZB': '30Y TB', 'FE': 'IRON ORE', 'GBM': 'BOBL', 'LB': 'RL LUMBER', 'KW': 'KCBT WHEAT', 'S': 'SOYBEAN', 'ZR': 'RICE', 'OJY': 'JPF FX O', 'NIY': 'NIY', 'CT': 'NB COTT', 'LC': 'L CATTLE', 'HO': 'HEATOIL', 'SI': 'CMX SIL', 'JB': 'QJG10YR', 'BP': 'GBP', 'NK': 'QNK', 'ZS': 'ZS', 'NQ': 'MINI NSDQ', 'QM': 'MINICRUDE', 'NE': 'NZD', 'SFE': 'IRONORE SP', 'ESX': 'DJEST50', 'M6E': 'MIC EURUSD', 'TF': 'TSR20', 'ONG': 'NATGAS O', 'CA': 'CA'}
        monthsmapping = {'01': 'JAN', '02': 'FEB', '03': 'MAR', '04': 'APR', '05': 'MAY', '06': 'JUN', '07': 'JUL', '08': 'AUG', '09': 'SEP', '10': 'OCT', '11': 'NOV', '12': 'DEC'}
        months = ['','JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC']

        new_dict = {v : k for k, v in monthsmapping.items()}

        '''CHICAGO CLOSE OUT:-   2'''

        #print(new_dict)
        #global moon
        def MonthTr_s(s,code):#s:合约代码，code:esunny产品代码
            moon=s[2:]
            #print(moon)
            return  monthsmapping[moon]+ ' '+s[:2] + ' '+Codemapping[code]

        def Qty_Closeout(q):
            return 'CHICAGO CLOSE OUT:-   '+q

        def OptionTr_s(s,code):
            moon=s[2:4]
            sprice=s[4:10]
            direc=s[-1]
            if direc == 'P':
               direc='PUT'
            else:
               direc = 'CALL'
            return direc+' '+monthsmapping[moon]+ ' '+s[:2] + ' '+Codemapping[code]+sprice
        
        row = []
        try:
            with open(filepath,'r',encoding='utf-8') as csvfile:
                reader = csv.reader(csvfile)
                for line in reader:
                    if line[0] == 'ADMISCQG':
                        row.append(line)
                        #row.append(line[2])
                        #row.append(line[3])
                        #row.append(line[4])
                    if line[0] == 'ADMISDMA':
                        row.append(line)
                        #row.append(line[2])
                        #row.append(line[3])
                        #row.append(line[4])
                    else:
                        pass
            print(row)
        except Exception:
            sys.stderr.write('读取文件发生IO异常！\n')
        finally:
            csvfile.close()
            sys.stderr.write('finnaly读取文件成功！\n')
        #print(row)
        tex.insert(END,row)
        for row1 in row:
            if row1[0] == 'ADMISCQG':
                row1[0] = 'CHN C1065'
            if row1[0] == 'ADMISDMA':
                row1[0] = 'CHN C1065'
            else:
                pass
            print(row1)
            if len(row1[3]) == 4:#期货处理
                p = MonthTr_s(row1[3],row1[2])
                row1[1] =p
                row1[2] = Qty_Closeout(row1[4])
                #print(row1)
                del row1[6]
                del row1[5]
                #print(row1)
                del row1[4]
                del row1[3]
                #print(row1)
            elif len(row1[3]) > 6:#期权处理
                p = OptionTr_s(row1[3],row1[2])
                row1[1] =p
                row1[2] = Qty_Closeout(row1[4])
                #print(row1)
                del row1[6]
                del row1[5]
                #print(row1)
                del row1[4]
                del row1[3]
            print(row1)
            tex.insert(END,row1)
        #print(row)
        sFilename = os.path.join(os.path.dirname(filepath),'closeout_ADM.csv')#创建文件路径

        eFile = open(sFilename,'w',newline='')

        eWriter = csv.writer(eFile,delimiter=',',lineterminator='\r\n')
        eWriter.writerow(['AccountOffice Number','GMI Description   ([PSDSC1])','Qty'])
        for clr in row:
            eWriter.writerow(clr)
        eFile.close()

    if FUpper == 'MAREX':
        months = ['','F','G','H','J','K','M','N','Q','U','V','X','Z']
        a = ['APEX','BMD','HKEX','CBOT','CBOE','CME','COMEX','EUREX','ICUS','LIFFE','LME','NYMEX','SGX','TOCOM']
        b = ['APEX','BMD','HKFE','CBOT','CBOE','CME','COMEX','EUREX','NYBOT','IPE','LME','CME','SGXQ','TOCOM']
        c = ['YW','YK','YC','UB','TN','EH','C','XBT','VX','UPO','POL','MHI','PF','UC', 'KW', 'O', 'S', 'W', 'YM', 'ZB', 'ZF', 'ZL', 'ZM', 'ZN', 'ZR', 'ZT', 'ZQ','OC', 'OS', 'OW', 'OZL', 'OZM', 'AD', 'BP', 'CD', 'EC', 'ED', 'ES', 'FC', 'JY', 'LB', 'LC', 'LN', 'M6A', 'M6E', 'MP', 'NE', 'NIY', 'NKD', 'NQ', 'SF', 'OAD', 'OBP', 'OCD', 'OEC', 'OES', 'OJY', 'ONQ', 'GC', 'HG', 'SI', 'OGC', 'DAX', 'ESX', 'GBL', 'GBM', 'GBS', 'CC', 'CT', 'DX', 'KC', 'NTF', 'OJ', 'SB11', 'YG', 'YI', 'LW', 'RC', 'Z', 'AH', 'CA', 'NI', 'PB', 'SN', 'ZS', 'CL', 'HO', 'NG', 'PA', 'PL', 'QM', 'RB', 'TIO', 'OCL', 'ONG', 'CN', 'FE', 'JB', 'MSG', 'NK', 'NU', 'TF', 'CH', 'JRU', 'STF', 'SFE']
        d = ['YW','XB','YC','UBE','TN','EH','C','XBT','VIX','FUPO','FPOL','MHI','PF','QUSDCNY','KCBT WHEAT', 'O', 'S', 'W', 'YM', 'US', 'FV', 'BO', 'SM', 'TY', 'RR', 'TU','ZQ', 'CORN O', 'SOYBEAN O', 'WHEAT O', 'SOYOIL O', 'SOYMEAL O', 'AUD', 'GBP', 'CAN ', 'Euro FX', 'GLBX EURO', 'MINI S&P', 'F CATTLE', 'JPY', 'RL LUMBER', 'L CATTLE', 'LEAN HOG', 'MIC AUDUSD', 'MIC EURUSD', 'MXP', 'NZD', 'NIY', 'NKD', 'MINI NSDQ', 'SWF', 'AUD FX O', 'GBP FX O', 'CAD FX O', 'EUR FX O', 'MINI S&P O', 'JPF FX O', 'MINI NQ O', 'CMX GLD', 'CMX COP', 'CMX SIL', 'CMX GLD  O', 'DAX', 'DJEST50', 'BUND', 'BOBL', 'SHAZ', 'NB COCO', 'NB COTT', 'NB DLR', 'NB COFY', 'MINIRL2000', 'NB FCOJ', 'NB SU11', 'MINIGOLD', 'MINISILVER', 'SUGAR', 'R COFFEE', 'FTSE100', 'AH', 'CA', 'NI', 'PB', 'SN', 'ZS', 'CRUDE', 'HEATOIL', 'NATGAS', 'NYM PAL', 'NYM PLA', 'MINICRUDE', 'GASOLINE', 'IRONORE', 'CRUDE O', 'NATGAS O', 'QXINHUA 50', 'IRON ORE', 'QJG10YR', 'QSG', 'QNK', 'QNU', 'TSR20', 'QCH', 'RUBBER', 'TSR20 SP', 'IRONORE SP']

        Markertmapping = {'CBOT':'CME_CBT','CME':'CME','COMEX':'CME','EUREX':'EUREX','ICUS':'NYBOT','LIFFE':'IPE','LIFFE1':'ICE_UK','LME':'LME','NYMEX':'CME','SGX':'SGXQ','TOCOM':'TOCOM'}
        Codemapping = {'FC': 'F CATTLE', 'NTF': 'MINIRL2000', 'OBP': 'GBP FX O', 'OCD': 'CAD FX O', 'GC': 'CMX GLD', 'NI': 'NI', 'CD': 'CAN ', 'SF': 'SWF', 'OJ': 'NB FCOJ', 'YG': 'MINIGOLD', 'NKD': 'NKD', 'STF': 'TSR20 SP', 'PL': 'NYM PLA', 'CL': 'CRUDE', 'YM': 'DJIA5', 'CN': 'QXINHUA 50', 'SB11': 'NB SU11', 'OW': 'WHEAT O', 'JY': 'JPY', 'RC': 'R COFFEE', 'OZL': 'SOYOIL O', 'TIO': 'IRONORE', 'ZN': '10Y TN', 'SN': 'SN', 'KC': 'NB COFY', 'LN': 'LEAN HOG', 'OS': 'SOYBEAN O', 'OAD': 'AUD FX O', 'YI': 'MINISILVER', 'OGC': 'CMX GLD  O', 'O': 'OATS', 'DX': 'NB DLR', 'OEC': 'EUR FX O', 'W': 'WHEAT', 'OES': 'MINI S&P O', 'RB': 'GASOLINE', 'LW': 'SUGAR', 'CH': 'QCH', 'ZL': 'SOYOIL', 'DAX': 'DAX', 'ZF': '5Y TN', 'OC': 'CORN O', 'MSG': 'QSG', 'AH': 'AH', 'OCL': 'CRUDE O', 'PB': 'PB', 'M6A': 'MIC AUDUSD', 'MP': 'MXP', 'C': 'CORN', 'CC': 'NB COCO', 'ES': 'MINI S&P', 'GBS': 'SHAZ', 'ED': 'GLBX EURO', 'ZM': 'SOYMEAL', 'ZT': '2Y TN', 'HG': 'CMX COP', 'GBL': 'BUND', 'PA': 'NYM PAL', 'NU': 'QNU', 'JRU': 'RUBBER', 'ONQ': 'MINI NQ O', 'NG': 'NATGAS', 'EC': 'Euro FX', 'Z': 'FTSE100', 'AD': 'AUD', 'OZM': 'SOYMEAL O', 'ZB': '30Y TB', 'FE': 'IRON ORE', 'GBM': 'BOBL', 'LB': 'RL LUMBER', 'KW': 'KCBT WHEAT', 'S': 'SOYBEAN', 'ZR': 'RICE', 'OJY': 'JPF FX O', 'NIY': 'NIY', 'CT': 'NB COTT', 'LC': 'L CATTLE', 'HO': 'HEATOIL', 'SI': 'CMX SIL', 'JB': 'QJG10YR', 'BP': 'GBP', 'NK': 'QNK', 'ZS': 'ZS', 'NQ': 'MINI NSDQ', 'QM': 'MINICRUDE', 'NE': 'NZD', 'SFE': 'IRONORE SP', 'ESX': 'DJEST50', 'M6E': 'MIC EURUSD', 'TF': 'TSR20', 'ONG': 'NATGAS O', 'CA': 'CA'}

        def tr_s(s):
            return months[int(s[2:])]+'20'+s[:2]

        def tr_s1(s):
            return months[int(s[2:5])]+'20'+s[0:2]
        def tr_s2(s):
            return s[-7:-2]
        def tr_s3(s):
            return s[-1]
        row = []
        try:
            with open(filepath,'r',encoding='utf-8') as csvfile:
                reader = csv.reader(csvfile)
                for line in reader:
                    if line[0] == 'MAREX':
                        row.append(line)
                    else:
                        pass
                        #print(row2)
        except Exception:
            sys.stderr.write('读取文件发生IO异常！\n')
        finally:
            csvfile.close()
            sys.stderr.write('finnaly读取文件成功！\n')
        #print(row)
        tex.insert(END,row)

        for row1 in row:
            if row1[0] == 'MAREX' and row1[1] !='LME':
                row1[0] = '20868'
            if row1[0] == 'MAREX' and row1[1] == 'LME':
                row1[0] = '20869'
            else:
                pass
            row1.insert(1,'F')
            row1.append('')
            row1.append('')
            row1.append('')
            row1[6] = ''
            row1[7] = ''
            row1[8] = ''
            row1[10] = row1[5]
            row1[5] = ''
            #print(row1)
            for j in row1:
                if j == row1[2]:
                    if row1[2] in a:#市场code转换
                        #print(row1[2])
                        if row1[3] in ["FTSE100"]:
                            row1[2]= "ICE_UK"
                        if row1[2] == "LME":
                            pass
                        else:
                            t = a.index(row1[2])
                            row1[2]= b[t]

                    else:
                        pass
                else:
                    pass
                if j == row1[3]:
                    if row1[3] in c:#产品code转换
                        t2 = c.index(row1[3])
                        row1[3] = d[t2]
                if len(row1[4]) > 6:#期权处理
                    p1 = tr_s1(row1[4])
                    p5 = tr_s2(row1[4])
                    row1[5] = p5
                    p6 = tr_s3(row1[4])
                    row1[6] = p6
                    row1[4] = p1
                    row1[1] = "O"
                elif len(row1[4]) == 4:
                    p1 = tr_s(row1[4])
                    row1[4] = p1
            print(row1)
            tex.insert(END,row1)
        #print(row)
        sFilename = os.path.join(os.path.dirname(filepath),'closeout_MAREX.csv')#创建文件路径

        eFile = open(sFilename,'w',newline='')

        eWriter = csv.writer(eFile,delimiter=',',lineterminator='\r\n')
        eWriter.writerow(['client_no','Com_Type','Exch_cd','Com_cd','Contract_Month','Strike_Price','Call_Put','Val_Date','Trade_Date_Buy','Trade_Date_Sell','Traded_Qty','Traded_Price_Buy','Traded_Price_Sell','Traded_Premium_Buy','Traded_Premium_Sell'])
        for clr in row:
            eWriter.writerow(clr)
        eFile.close()

# 按钮
action = ttk.Button(win, text="生成closeout",state = 'active', command=clickMe)     # 创建一个按钮, text：显示按钮上面显示的文字, command：当这个按钮被点击之后会调用command函数
action.grid(column=2, row=1)    # 设置其在界面中出现的位置  column代表列   row 代表行


#root = Tk()
path = tk.StringVar()

Label(win,text = "目标路径:").grid(row = 2, column = 0)
Entry(win, textvariable = path).grid(row = 2, column = 1)
action2 = ttk.Button(win, text="选择目标文件",state = 'active', command=selectPath)
action2.grid(column=2, row=2)
#Button(win, text = "路径选择", command = selectPath).grid(row = 2, column = 2)
#print(path)

# 文本框
name = tk.StringVar()     # StringVar是Tk库内部定义的字符串变量类型，在这里用于管理部件上面的字符；不过一般用在按钮button上。改变StringVar，按钮上的文字也随之改变。
nameEntered = ttk.Entry(win, width=14, textvariable=name)   # 创建一个文本框，定义长度为12个字符长度，并且将文本框中的内容绑定到上一句定义的name变量上，方便clickMe调用
nameEntered.grid(column=0, row=1)       # 设置其在界面中出现的位置  column代表列   row 代表行
nameEntered.focus()     # 当程序运行时,光标默认会出现在该文本框中

# 创建一个下拉列表
number = tk.StringVar()
numberChosen = ttk.Combobox(win, width=14, textvariable=number)
numberChosen['values'] = ('PHL', 'EDFMAN','ADM')     # 设置下拉列表的值
numberChosen.grid(column=1, row=1)      # 设置其在界面中出现的位置  column代表列   row 代表行
numberChosen.current(0)    # 设置下拉列表默认显示的值，0为 numberChosen['values'] 的下标值
win.geometry('400x300')
#win.center_window(win,340,80)
win.maxsize(400, 200)
win.minsize(300, 100)
#win.iconbitmap('C:/Python33/Scripts/dist/256.ico')
win.mainloop()      # 当调用mainloop()时,窗口才会显示出来

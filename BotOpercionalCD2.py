#!/usr/bin/env python
# coding: utf-8



import os
import time
import datetime
import win32com.client



print("BOT OPERACIONAL - CERTIFICADO\n")
print("A Planilha de Apuração de resultado de metas abre 08:55 e encerra a atualização 09:00\n")
os.startfile('C:\\ONEDRIVE\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Operacional CD\\Apuração de Resultados de Metas.xlsm')






hr = str('09:01')
hr = hr.split(':')
t_desp = datetime.datetime.combine( datetime.datetime.now().date(),
                                    datetime.time( int(hr[0]), int(hr[1])) )


while datetime.datetime.now() < t_desp:

    time.sleep(30)

else:
    xlapp = win32com.client.DispatchEx("Excel.Application")
    xlapp.Visible = True
    xlapp.DisplayAlerts = False



    wb = xlapp.Workbooks.Open("C:\\ONEDRIVE\\OneDrive - CERTIFICA BRASIL SERVICOS DE CERTIFICACAO DIGITAL\\Apuração de Resultados\\Operacional CD\\Ligações Bitrix - CD.xlsm")
    time.sleep(5)


    xlapp.Application.Run("'Ligações Bitrix - CD.xlsm'!AbrirCall")

    time.sleep(90)

    wb.RefreshAll()
    time.sleep(60)

    wb.Save()
    time.sleep(5)
    wb.Close()
    xlapp.Quit()

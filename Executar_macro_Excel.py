import os
import win32com.client


  # Abrir o arquivo excel
    xlapp = win32com.client.DispatchEx("Excel.Application")
    # Excel Visivel
    xlapp.Visible = True
    # Desativar alertas
    xlapp.DisplayAlerts = False


    # Abrir arquivo
    wb = xlapp.Workbooks.Open("enderço do arquivo excel")

    
    # Ativar cálculo manual (só funciona após o arquivo estiver aberto)
    xlapp.Calculation = -4135

    # Executar a macro
    xlapp.Application.Run("'NOME DO ARQUIVO.xlsm'!NOME DA MACRO")
    
    # Ativar cálculo automático
    xlapp.Calculation = -4105

    # Atualizar consultas
    wb.RefreshAll()
    
    #Atualizar consulta específica
    wb.Connections("Consulta - NOME DA CONSULTA").Refresh()

    # Salvar o arquivo excel
    wb.Save()
    
    # Fechar sheets
    wb.Close()
    
    # Fechar excel
    xlapp.Quit()

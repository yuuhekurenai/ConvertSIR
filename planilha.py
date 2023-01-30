"""
    Esse parte do projeto irá transformar a malha conforme a ferramenta
    de dimensionamento do capacity planning após isso estará pronta para
    se utilizada para o dimensionamento de demanda.

"""
from openpyxl import Workbook
import datetime

# Nomear colunas conforme a ferramenta de dimensionamento do capacity

class CriarMalhaDimensionamento:

    def __init__(self):
        self.malha = Workbook()
        self.planilha = self.malha.active
        self.planilha['A1'] = "Dept Sta"
        self.planilha['B1'] = "Arvl Sta"
        self.planilha['C1'] = "Dept Day"
        self.planilha['D1'] = "Arvl Day"
        self.planilha['E1'] = "Dept Time"
        self.planilha['F1'] = "Arvl Time"
        self.planilha['G1'] = "Subfleet"
        self.planilha['H1'] = "FlightNumber"
        self.planilha['I1'] = "Trilho"
        self.planilha['J1'] = "Svc Type"
        self.planilha['K1'] = "Week Day"

    # Escreve os dados na planilha conforme a coluna e a linha


class EscreverDados(CriarMalhaDimensionamento):
    def __init__(self):
        super(EscreverDados, self).__init__()
        coluna = 1
        linha = 2

        deptsta = str('')
        self.planilha.cell(column=coluna, row=linha).value = deptsta
        coluna += 1
        self.arvlsta = str('')
        self.planilha.cell(column=coluna, row=linha).value = self.arvlsta
        coluna += 1
        self.deptday = str('')
        self.planilha.cell(column=coluna, row=linha).value = self.deptday
        coluna += 1
        self.arvlday = str('')
        self.planilha.cell(column=coluna, row=linha).value = self.arvlday
        coluna += 1
        self.depttime = str('')
        self.planilha.cell(column=coluna, row=linha).value = self.depttime
        coluna += 1
        self.arvltime = str('')
        self.planilha.cell(column=coluna, row=linha).value = self.arvltime
        coluna += 1
        self.subfleet = str('')
        self.planilha.cell(column=coluna, row=linha).value = self.subfleet
        coluna += 1
        self.flightnumber = str('')
        self.planilha.cell(column=coluna, row=linha).value = self.flightnumber
        coluna += 1
        self.trilho = 1
        self.planilha.cell(column=coluna, row=linha).value = self.trilho
        coluna += 1
        self.svctype = str('')
        self.planilha.cell(column=coluna, row=linha).value = self.svctype
        coluna += 1
        self.weekday = str('')
        self.planilha.cell(column=coluna, row=linha).value = self.weekday
        coluna -= 10
        linha += 1
        self.malha.save(filename=f"Malha SIR - {datetime.date.today()}.xlsx")

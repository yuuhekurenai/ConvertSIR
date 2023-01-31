"""
    Esse parte do projeto irá transformar a malha conforme a ferramenta
    de dimensionamento do capacity planning após isso estará pronta para
    se utilizada para o dimensionamento de demanda.

"""
from openpyxl import Workbook, load_workbook
import datetime

Base = 'VCP'
time_null = '00:00'

"""Nomear colunas conforme a ferramenta de dimensionamento do capacity"""


class Sir:

    def __init__(self):
        # Pequenos parametros
        self.base = Base
        self.tempo = time_null
        self.linha = 2
        self.coluna = 1
        self.escrever_coluna = 1
        self.escrever_linha = 2

        """Carrega a planilha tratada pelo pandas assim atibuindos respectivos valores das colunas"""

        self.malha = load_workbook(f'SIR - MALHA {datetime.date.today()}.xlsx')
        self.sir = self.malha.active
        self.dptsta = self.sir.cell(column=self.coluna, row=self.linha)
        self.coluna += 1
        self.arlvsta = self.sir.cell(column=self.coluna, row=self.linha)
        self.coluna += 1
        self.deptday = self.sir.cell(column=self.coluna, row=self.linha)
        self.coluna += 1
        self.arvlday = self.sir.cell(column=self.coluna, row=self.linha)
        self.coluna += 1
        self.depttime = self.sir.cell(column=self.coluna, row=self.linha)
        self.coluna += 1
        self.arvltime = self.sir.cell(column=self.coluna, row=self.linha)
        self.coluna += 1
        self.subfleet = self.sir.cell(column=self.coluna, row=self.linha)
        self.coluna += 1
        self.flightnumber = self.sir.cell(column=self.coluna, row=self.linha)
        self.coluna += 2
        self.trilho = self.sir.cell(column=self.coluna, row=self.linha)
        self.coluna += 1
        self.svctype = self.sir.cell(column=self.coluna, row=self.linha)
        self.coluna += 1
        self.weekday = self.sir.cell(column=self.coluna, row=self.linha)

        """Cria e nova planilha e escreve o cabeçalho"""

        self.planilha = self.malha.create_sheet('MALHA', 1)
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

        """Extraindo dados da primeira planilha e transformando para o formato Capacity Planning"""


class EscreverMalha(Sir):

    def __init__(self):
        super().__init__()
        self.planilha.cell(column=self.escrever_coluna, row=self.escrever_linha).value = self.base
        self.escrever_coluna += 1
        self.planilha.cell(column=self.escrever_coluna, row=self.escrever_linha).value = self.arlvsta.value
        self.escrever_coluna += 1
        self.planilha.cell(column=self.escrever_coluna, row=self.escrever_linha).value = self.deptday.value
        self.escrever_coluna += 1
        self.planilha.cell(column=self.escrever_coluna, row=self.escrever_linha).value = self.arvlday.value
        self.escrever_coluna += 1
        self.planilha.cell(column=self.escrever_coluna,
                           row=self.escrever_linha).value = f'{self.deptday.value} {self.depttime.value}'
        self.escrever_coluna += 1
        self.planilha.cell(column=self.escrever_coluna,
                           row=self.escrever_linha).value = f'{self.arvlday.value} {self.tempo}'
        self.escrever_coluna += 1
        self.planilha.cell(column=self.escrever_coluna, row=self.escrever_linha).value = self.subfleet.value
        self.escrever_coluna += 1
        self.planilha.cell(column=self.escrever_coluna, row=self.escrever_linha).value = self.flightnumber.value
        self.escrever_coluna += 1
        self.planilha.cell(column=self.escrever_coluna, row=self.escrever_linha).value = self.trilho.value
        self.escrever_coluna += 1
        self.planilha.cell(column=self.escrever_coluna, row=self.escrever_linha).value = self.svctype.value
        self.escrever_coluna += 1
        self.planilha.cell(column=self.escrever_coluna, row=self.escrever_linha).value = self.weekday.value
        self.escrever_coluna -= 10
        self.escrever_linha += 1


        """Escreve os dados na planilha conforme a coluna e a linha"""

        self.malha.save(filename=f'SIR - MALHA {datetime.date.today()}.xlsx')

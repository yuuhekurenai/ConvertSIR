from openpyxl import load_workbook
from pyxll import xl_macro, xl_app

import datetime
# Importa o arquivo tratado.
from format_tool import formata_txt
# Importa o pandas
import pandas as pd
import datetime


def model_vcp():
    # Definir base

    base = "VCP"

    # Chama a função passando o nome do arquivo txt entre aspas, que será modificado

    formata_txt("teste.txt")

    # Lê o arquivo já modificado

    df = pd.read_csv(

        # Nome do arquivo

        "teste.txt",

        # Usa o espaço entre as palavras para separar por colunas

        delim_whitespace=True,

        # Usa estes nomes como nomes das colunas (alterar a vontade)

        names=["Dept Sta", "Arvl Sta", "Dept Day", "Week-Day", "Pax-Subfleet", "Arvl-Hora", "Dept-Hora", "Svc Type",
               "FlightNumbeR", "FlightNumberArvl"],

        header=None,

    )

    # Imprime as linhas concactena e gera as quatro Keys para filtro (Key, KeyX, Kdept e Karvl)
    df['Concatenar'] = df['Arvl Sta'].astype(str) + df['Dept Sta'].astype(str) + df['Dept Day'].astype(str) + df[
        'Week-Day'].astype(str) + df['Pax-Subfleet'].astype(str) + df['Pax-Subfleet'].astype(str) + df[
                           'Arvl-Hora'].astype(str) + df['Dept-Hora'].astype(str) + df['Svc Type'].astype(str)

    df['Concatenar'] = df['Concatenar'].astype(str)
    df['Key'] = df['Concatenar'].str.len()
    df['KeyX'] = df['Svc Type'].astype(str)
    df['KeyX'] = df['KeyX'].str.len()
    df['Kdept'] = df['Dept-Hora'].astype(str)
    df['Kdept'] = df['Kdept'].str.len()
    df['Karvl'] = df['Arvl-Hora'].astype(str)
    df['Karvl'] = df['Karvl'].str.len()
    df_orig = df

    # Com as Keys geradas, agora é possível buscar todos os padrões e tratá-los.

    df_frame45 = df_orig.query('KeyX == 1  & Key == 45 & Kdept == 7').copy()
    df_frame45['FlightNumbeR'] = df_frame45['Arvl Sta']
    df_frame45['Dept Sta'] = base
    df_frame45['Arvl Sta'] = df_frame45['Dept-Hora'].apply(lambda x: x[4:7])
    df_frame45['Arvl Day'] = df_frame45['Dept Day'].apply(
        lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023' + ' á ' + x[5:7] + '/' + x[7:10] + '/' + '2023')
    df_frame45['Dept Day'] = df_frame45['Dept Day'].apply(
        lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023' + ' á ' + x[5:7] + '/' + x[7:10] + '/' + '2023')
    df_frame45['Trilho'] = ''
    df_frame45['Pax'] = df_frame45['Pax-Subfleet'].apply(lambda x: x[0:3])
    df_frame45['Pax-Subfleet'] = df_frame45['Pax-Subfleet'].apply(lambda x: x[3:6])
    df_frame45['Arvl-Hora'] = '00:00'
    df_frame45['Dept-Hora'] = df_frame45['Dept-Hora'].apply(lambda x: x[0:2] + ':' + x[2:4])
    df_frame45 = df_frame45[
        ['Dept Sta', 'Arvl Sta', 'Dept Day', 'Arvl Day', 'Dept-Hora', 'Arvl-Hora', 'Pax-Subfleet', 'FlightNumbeR',
         'FlightNumberArvl', 'Trilho', 'Svc Type', 'Pax', 'Week-Day']]
    p1 = df_frame45
    df_frame34 = df_orig.query('KeyX == 1  & Key == 34 & Kdept == 7').copy()
    df_frame34['FlightNumbeR'] = df_frame34['Arvl Sta']
    df_frame34['Dept Sta'] = base
    df_frame34['Arvl Sta'] = df_frame34['Dept-Hora'].apply(lambda x: x[4:7])
    df_frame34['Arvl Day'] = df_frame34['Dept Day'].apply(lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023')
    df_frame34['Dept Day'] = df_frame34['Dept Day'].apply(lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023')
    df_frame34['Trilho'] = ''
    df_frame34['Pax'] = df_frame34['Pax-Subfleet'].apply(lambda x: x[0:3])
    df_frame34['Pax-Subfleet'] = df_frame34['Pax-Subfleet'].apply(lambda x: x[3:6])
    df_frame34['Arvl-Hora'] = '00:00'
    df_frame34['Dept-Hora'] = df_frame34['Dept-Hora'].apply(lambda x: x[0:2] + ':' + x[2:4])
    df_frame34 = df_frame34[
        ['Dept Sta', 'Arvl Sta', 'Dept Day', 'Arvl Day', 'Dept-Hora', 'Arvl-Hora', 'Pax-Subfleet', 'FlightNumbeR',
         'FlightNumberArvl', 'Trilho', 'Svc Type', 'Pax', 'Week-Day']]
    p2 = df_frame34
    df_frame48 = df_orig.query('KeyX == 1  & Key == 48 & Kdept == 10').copy()
    df_frame48['FlightNumbeR'] = df_frame48['Arvl Sta']
    df_frame48['Dept Sta'] = base
    df_frame48['Arvl Sta'] = df_frame48['Dept-Hora'].apply(lambda x: x[4:7])
    df_frame48['Arvl Day'] = df_frame48['Dept Day'].apply(
        lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023' + ' á ' + x[5:7] + '/' + x[7:10] + '/' + '2023')
    df_frame48['Dept Day'] = df_frame48['Dept Day'].apply(
        lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023' + ' á ' + x[5:7] + '/' + x[7:10] + '/' + '2023')
    df_frame48['Trilho'] = ''
    df_frame48['Pax'] = df_frame48['Pax-Subfleet'].apply(lambda x: x[0:3])
    df_frame48['Pax-Subfleet'] = df_frame48['Pax-Subfleet'].apply(lambda x: x[3:6])
    df_frame48['Arvl-Hora'] = '00:00'
    df_frame48['Dept-Hora'] = df_frame48['Dept-Hora'].apply(lambda x: x[0:2] + ':' + x[2:4])
    df_frame48 = df_frame48[
        ['Dept Sta', 'Arvl Sta', 'Dept Day', 'Arvl Day', 'Dept-Hora', 'Arvl-Hora', 'Pax-Subfleet', 'FlightNumbeR',
         'FlightNumberArvl', 'Trilho', 'Svc Type', 'Pax', 'Week-Day']]
    p3 = df_frame48
    df_frame37 = df_orig.query('KeyX == 1  & Key == 37 & Kdept == 1 & Karvl == 10').copy()
    df_frame37['FlightNumberArvl'] = df_frame37['Dept Sta'].apply(lambda x: x[1:6])
    df_frame37['Dept Sta'] = df_frame37['Arvl-Hora'].apply(lambda x: x[3:6])
    df_frame37['Arvl Sta'] = base
    df_frame37['Arvl Day'] = df_frame37['Dept Day'].apply(lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023')
    df_frame37['Dept Day'] = df_frame37['Dept Day'].apply(lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023')
    df_frame37['Trilho'] = ''
    df_frame37['Pax'] = df_frame37['Pax-Subfleet'].apply(lambda x: x[0:3])
    df_frame37['Pax-Subfleet'] = df_frame37['Pax-Subfleet'].apply(lambda x: x[3:6])
    df_frame37['Arvl-Hora'] = df_frame37['Arvl-Hora'].apply(lambda x: x[6:8] + ':' + x[8:10])
    df_frame37['Dept-Hora'] = '00:00'
    df_frame37 = df_frame37[
        ['Dept Sta', 'Arvl Sta', 'Dept Day', 'Arvl Day', 'Dept-Hora', 'Arvl-Hora', 'Pax-Subfleet', 'FlightNumbeR',
         'FlightNumberArvl', 'Trilho', 'Svc Type', 'Pax', 'Week-Day']]
    p40 = df_frame37
    df_frame37 = df_orig.query('KeyX == 1  & Key == 37  & Kdept == 10').copy()
    df_frame37['FlightNumbeR'] = df_frame37['Arvl Sta'].apply(lambda x: x[0:8])
    df_frame37['Dept Sta'] = base
    df_frame37['Arvl Sta'] = df_frame37['Dept-Hora'].apply(lambda x: x[4:7])
    df_frame37['Arvl Day'] = df_frame37['Dept Day'].apply(lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023')
    df_frame37['Dept Day'] = df_frame37['Dept Day'].apply(lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023')
    df_frame37['Trilho'] = ''
    df_frame37['Pax'] = df_frame37['Pax-Subfleet'].apply(lambda x: x[0:3])
    df_frame37['Pax-Subfleet'] = df_frame37['Pax-Subfleet'].apply(lambda x: x[3:6])
    df_frame37['Arvl-Hora'] = df_frame37['Dept-Hora'].apply(lambda x: x[0:2] + ':' + x[2:4])
    df_frame37['Dept-Hora'] = '00:00'
    df_frame37 = df_frame37[
        ['Dept Sta', 'Arvl Sta', 'Dept Day', 'Arvl Day', 'Dept-Hora', 'Arvl-Hora', 'Pax-Subfleet', 'FlightNumbeR',
         'FlightNumberArvl', 'Trilho', 'Svc Type', 'Pax', 'Week-Day']]
    p5 = df_frame37
    df_frame36 = df_orig.query('KeyX == 1  & Key == 36 & Kdept == 10').copy()
    df_frame36['FlightNumbeR'] = df_frame36['Arvl Sta']
    df_frame36['Dept Sta'] = base
    df_frame36['Arvl Sta'] = df_frame36['Dept-Hora'].apply(lambda x: x[4:7])
    df_frame36['Arvl Day'] = df_frame36['Dept Day'].apply(lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023')
    df_frame36['Dept Day'] = df_frame36['Dept Day'].apply(lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023')
    df_frame36['Trilho'] = ''
    df_frame36['Pax'] = df_frame36['Pax-Subfleet'].apply(lambda x: x[0:3])
    df_frame36['Pax-Subfleet'] = df_frame36['Pax-Subfleet'].apply(lambda x: x[3:6])
    df_frame36['Arvl-Hora'] = '00:00'
    df_frame36['Dept-Hora'] = df_frame36['Dept-Hora'].apply(lambda x: x[0:2] + ':' + x[2:4])
    df_frame36 = df_frame36[
        ['Dept Sta', 'Arvl Sta', 'Dept Day', 'Arvl Day', 'Dept-Hora', 'Arvl-Hora', 'Pax-Subfleet', 'FlightNumbeR',
         'FlightNumberArvl', 'Trilho', 'Svc Type', 'Pax', 'Week-Day']]
    p6 = df_frame36
    df_frame46 = df_orig.query('KeyX == 1  & Key == 46 & Kdept == 1').copy()
    df_frame46['FlightNumberArvl'] = df_frame46['Dept Sta'].apply(lambda x: x[1:8])
    df_frame46['Dept Sta'] = df_frame46['Arvl-Hora'].apply(lambda x: x[0:3])
    df_frame46['FlightNumbeR'] = df_frame46['Arvl Sta']
    df_frame46['Arvl Day'] = df_frame46['Dept Day'].apply(
        lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023' + ' á ' + x[5:7] + '/' + x[7:10] + '/' + '2023')
    df_frame46['Dept Day'] = df_frame46['Dept Day'].apply(
        lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023' + ' á ' + x[5:7] + '/' + x[7:10] + '/' + '2023')
    df_frame46['Trilho'] = ''
    df_frame46['Pax'] = df_frame46['Pax-Subfleet'].apply(lambda x: x[0:3])
    df_frame46['Pax-Subfleet'] = df_frame46['Pax-Subfleet'].apply(lambda x: x[3:6])
    df_frame46['Dept-Hora'] = df_frame46['Arvl-Hora'].apply(lambda x: x[3:5] + ':' + x[5:7])
    df_frame46['Arvl-Hora'] = '00:00'
    df_frame46['Arvl Sta'] = base
    df_frame46 = df_frame46[
        ['Dept Sta', 'Arvl Sta', 'Dept Day', 'Arvl Day', 'Dept-Hora', 'Arvl-Hora', 'Pax-Subfleet', 'FlightNumbeR',
         'FlightNumberArvl', 'Trilho', 'Svc Type', 'Pax', 'Week-Day']]
    p7 = df_frame46
    df_frame35 = df_orig.query('KeyX == 1  & Key == 35 & Kdept == 1').copy()
    df_frame35['FlightNumberArvl'] = df_frame35['Dept Sta']
    df_frame35['Dept Sta'] = df_frame35['Arvl-Hora'].apply(lambda x: x[0:3])
    df_frame35['Arvl Sta'] = base
    df_frame35['Arvl Day'] = df_frame35['Dept Day'].apply(lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023')
    df_frame35['Dept Day'] = df_frame35['Dept Day'].apply(lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023')
    df_frame35['Trilho'] = ''
    df_frame35['Pax'] = df_frame35['Pax-Subfleet'].apply(lambda x: x[0:3])
    df_frame35['Pax-Subfleet'] = df_frame35['Pax-Subfleet'].apply(lambda x: x[3:6])
    df_frame35['Arvl-Hora'] = df_frame35['Arvl-Hora'].apply(lambda x: x[3:5] + ':' + x[5:7])
    df_frame35['Dept-Hora'] = '00:00'
    df_frame35 = df_frame35[
        ['Dept Sta', 'Arvl Sta', 'Dept Day', 'Arvl Day', 'Dept-Hora', 'Arvl-Hora', 'Pax-Subfleet', 'FlightNumbeR',
         'FlightNumberArvl', 'Trilho', 'Svc Type', 'Pax', 'Week-Day']]
    p8 = df_frame35
    df_frame38 = df_orig.query('KeyX == 1  & Key == 38 & Kdept == 1').copy()
    df_frame38['FlightNumberArvl'] = df_frame38['Dept Sta'].apply(lambda x: x[1:8])
    df_frame38['Dept Sta'] = df_frame38['Arvl-Hora'].apply(lambda x: x[3:6])
    df_frame38['Arvl Sta'] = base
    df_frame38['Arvl Day'] = df_frame38['Dept Day'].apply(lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023')
    df_frame38['Dept Day'] = df_frame38['Dept Day'].apply(lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023')
    df_frame38['Trilho'] = ''
    df_frame38['Pax'] = df_frame38['Pax-Subfleet'].apply(lambda x: x[0:3])
    df_frame38['Pax-Subfleet'] = df_frame38['Pax-Subfleet'].apply(lambda x: x[3:6])
    df_frame38['Arvl-Hora'] = df_frame38['Arvl-Hora'].apply(lambda x: x[6:8] + ':' + x[8:10])
    df_frame38['Dept-Hora'] = '00:00'
    df_frame38 = df_frame38[
        ['Dept Sta', 'Arvl Sta', 'Dept Day', 'Arvl Day', 'Dept-Hora', 'Arvl-Hora', 'Pax-Subfleet', 'FlightNumbeR',
         'FlightNumberArvl', 'Trilho', 'Svc Type', 'Pax', 'Week-Day']]
    p9 = df_frame38
    df_frame49 = df_orig.query('KeyX == 1  & Key == 49 & Kdept == 1').copy()
    df_frame49['FlightNumberArvl'] = df_frame49['Dept Sta'].apply(lambda x: x[1:8])
    df_frame49['Dept Sta'] = df_frame49['Arvl-Hora'].apply(lambda x: x[3:6])
    df_frame49['Arvl Sta'] = base
    df_frame49['Arvl Day'] = df_frame49['Dept Day'].apply(
        lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023' + ' á ' + x[5:7] + '/' + x[7:10] + '/' + '2023')
    df_frame49['Dept Day'] = df_frame49['Dept Day'].apply(
        lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023' + ' á ' + x[5:7] + '/' + x[7:10] + '/' + '2023')
    df_frame49['Trilho'] = ''
    df_frame49['Pax'] = df_frame49['Pax-Subfleet'].apply(lambda x: x[0:3])
    df_frame49['Pax-Subfleet'] = df_frame49['Pax-Subfleet'].apply(lambda x: x[3:6])
    df_frame49['Arvl-Hora'] = df_frame49['Arvl-Hora'].apply(lambda x: x[6:8] + ':' + x[8:10])
    df_frame49['Dept-Hora'] = '00:00'
    df_frame49 = df_frame49[
        ['Dept Sta', 'Arvl Sta', 'Dept Day', 'Arvl Day', 'Dept-Hora', 'Arvl-Hora', 'Pax-Subfleet', 'FlightNumbeR',
         'FlightNumberArvl', 'Trilho', 'Svc Type', 'Pax', 'Week-Day']]
    p10 = df_frame49
    df_frame45 = df_orig.query('KeyX == 2  & Key == 45 & Kdept == 7 & Karvl == 7').copy()
    df_frame45['FlightNumberArvl'] = df_frame45['Arvl Sta']
    df_frame45['FlightNumbeR'] = df_frame45['Dept Sta'].apply(lambda x: x[1:8])
    df_frame45['Dept Sta'] = df_frame45['Arvl-Hora'].apply(lambda x: x[0:3])
    df_frame45['Arvl Sta'] = df_frame45['Dept-Hora'].apply(lambda x: x[4:7])
    df_frame45['Arvl Day'] = df_frame45['Dept Day'].apply(lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023')
    df_frame45['Dept Day'] = df_frame45['Dept Day'].apply(lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023')
    df_frame45['Trilho'] = ''
    df_frame45['Pax'] = df_frame45['Pax-Subfleet'].apply(lambda x: x[0:3])
    df_frame45['Pax-Subfleet'] = df_frame45['Pax-Subfleet'].apply(lambda x: x[3:6])
    df_frame45['Arvl-Hora'] = df_frame45['Arvl-Hora'].apply(lambda x: x[3:5] + ':' + x[5:7])
    df_frame45['Dept-Hora'] = df_frame45['Dept-Hora'].apply(lambda x: x[0:2] + ':' + x[2:4])
    df_frame45 = df_frame45[
        ['Dept Sta', 'Arvl Sta', 'Dept Day', 'Arvl Day', 'Dept-Hora', 'Arvl-Hora', 'Pax-Subfleet', 'FlightNumbeR',
         'FlightNumberArvl', 'Trilho', 'Svc Type', 'Pax', 'Week-Day']]
    p11 = df_frame45
    df_frame46 = df_orig.query('KeyX == 2  & Key == 46 & Kdept == 8 & Karvl == 7').copy()
    df_frame46['FlightNumberArvl'] = df_frame46['Arvl Sta']
    df_frame46['FlightNumbeR'] = df_frame46['Dept Sta'].apply(lambda x: x[1:8])
    df_frame46['Dept Sta'] = df_frame46['Arvl-Hora'].apply(lambda x: x[0:3])
    df_frame46['Arvl Sta'] = df_frame46['Dept-Hora'].apply(lambda x: x[5:8])
    df_frame46['Arvl Day'] = df_frame46['Dept Day'].apply(lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023')
    df_frame46['Dept Day'] = df_frame46['Dept Day'].apply(lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023')
    df_frame46['Trilho'] = ''
    df_frame46['Pax'] = df_frame46['Pax-Subfleet'].apply(lambda x: x[0:3])
    df_frame46['Pax-Subfleet'] = df_frame46['Pax-Subfleet'].apply(lambda x: x[3:6])
    df_frame46['Arvl-Hora'] = df_frame46['Arvl-Hora'].apply(lambda x: x[3:5] + ':' + x[5:7])
    df_frame46['Dept-Hora'] = df_frame46['Dept-Hora'].apply(lambda x: x[0:2] + ':' + x[2:4])
    df_frame46 = df_frame46[
        ['Dept Sta', 'Arvl Sta', 'Dept Day', 'Arvl Day', 'Dept-Hora', 'Arvl-Hora', 'Pax-Subfleet', 'FlightNumbeR',
         'FlightNumberArvl', 'Trilho', 'Svc Type', 'Pax', 'Week-Day']]
    p12 = df_frame46
    df_frame47 = df_orig.query('KeyX == 2  & Key == 47 & Kdept == 7 & Karvl == 7').copy()
    df_frame47['FlightNumberArvl'] = df_frame47['Arvl Sta']
    df_frame47['FlightNumbeR'] = df_frame47['Dept Sta'].apply(lambda x: x[1:8])
    df_frame47['Dept Sta'] = df_frame47['Arvl-Hora'].apply(lambda x: x[0:3])
    df_frame47['Arvl Sta'] = df_frame47['Dept-Hora'].apply(lambda x: x[4:8])
    df_frame47['Arvl Day'] = df_frame47['Dept Day'].apply(lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023')
    df_frame47['Dept Day'] = df_frame47['Dept Day'].apply(lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023')
    df_frame47['Trilho'] = ''
    df_frame47['Pax'] = df_frame47['Pax-Subfleet'].apply(lambda x: x[0:3])
    df_frame47['Pax-Subfleet'] = df_frame47['Pax-Subfleet'].apply(lambda x: x[3:6])
    df_frame47['Arvl-Hora'] = df_frame47['Arvl-Hora'].apply(lambda x: x[3:5] + ':' + x[5:7])
    df_frame47['Dept-Hora'] = df_frame47['Dept-Hora'].apply(lambda x: x[0:2] + ':' + x[2:4])
    df_frame47 = df_frame47[
        ['Dept Sta', 'Arvl Sta', 'Dept Day', 'Arvl Day', 'Dept-Hora', 'Arvl-Hora', 'Pax-Subfleet', 'FlightNumbeR',
         'FlightNumberArvl', 'Trilho', 'Svc Type', 'Pax', 'Week-Day']]
    p13 = df_frame47
    df_frame48 = df_orig.query('KeyX == 2  & Key == 48 & Kdept == 10 & Karvl == 7').copy()
    df_frame48['FlightNumberArvl'] = df_frame48['Arvl Sta']
    df_frame48['FlightNumbeR'] = df_frame48['Dept Sta'].apply(lambda x: x[1:8])
    df_frame48['Dept Sta'] = df_frame48['Arvl-Hora'].apply(lambda x: x[0:3])
    df_frame48['Arvl Sta'] = df_frame48['Dept-Hora'].apply(lambda x: x[4:7])
    df_frame48['Arvl Day'] = df_frame48['Dept Day'].apply(lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023')
    df_frame48['Dept Day'] = df_frame48['Dept Day'].apply(lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023')
    df_frame48['Trilho'] = ''
    df_frame48['Pax'] = df_frame48['Pax-Subfleet'].apply(lambda x: x[0:3])
    df_frame48['Pax-Subfleet'] = df_frame48['Pax-Subfleet'].apply(lambda x: x[3:6])
    df_frame48['Arvl-Hora'] = df_frame48['Arvl-Hora'].apply(lambda x: x[3:5] + ':' + x[5:7])
    df_frame48['Dept-Hora'] = df_frame48['Dept-Hora'].apply(lambda x: x[0:2] + ':' + x[2:4])
    df_frame48 = df_frame48[
        ['Dept Sta', 'Arvl Sta', 'Dept Day', 'Arvl Day', 'Dept-Hora', 'Arvl-Hora', 'Pax-Subfleet', 'FlightNumbeR',
         'FlightNumberArvl', 'Trilho', 'Svc Type', 'Pax', 'Week-Day']]
    p14 = df_frame48
    df_frame48 = df_orig.query('KeyX == 2  & Key == 48 & Kdept == 8 & Karvl == 7').copy()
    df_frame48['FlightNumberArvl'] = df_frame48['Arvl Sta']
    df_frame48['FlightNumbeR'] = df_frame48['Dept Sta'].apply(lambda x: x[1:8])
    df_frame48['Dept Sta'] = df_frame48['Arvl-Hora'].apply(lambda x: x[0:3])
    df_frame48['Arvl Sta'] = df_frame48['Dept-Hora'].apply(lambda x: x[5:8])
    df_frame48['Arvl Day'] = df_frame48['Dept Day'].apply(lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023')
    df_frame48['Dept Day'] = df_frame48['Dept Day'].apply(lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023')
    df_frame48['Trilho'] = ''
    df_frame48['Pax'] = df_frame48['Pax-Subfleet'].apply(lambda x: x[0:3])
    df_frame48['Pax-Subfleet'] = df_frame48['Pax-Subfleet'].apply(lambda x: x[3:6])
    df_frame48['Arvl-Hora'] = df_frame48['Arvl-Hora'].apply(lambda x: x[3:5] + ':' + x[5:7])
    df_frame48['Dept-Hora'] = df_frame48['Dept-Hora'].apply(lambda x: x[0:2] + ':' + x[2:4])
    df_frame48 = df_frame48[
        ['Dept Sta', 'Arvl Sta', 'Dept Day', 'Arvl Day', 'Dept-Hora', 'Arvl-Hora', 'Pax-Subfleet', 'FlightNumbeR',
         'FlightNumberArvl', 'Trilho', 'Svc Type', 'Pax', 'Week-Day']]
    p15 = df_frame48
    df_frame48 = df_orig.query('KeyX == 2  & Key == 48 & Kdept == 7 & Karvl == 10').copy()
    df_frame48['FlightNumberArvl'] = df_frame48['Arvl Sta']
    df_frame48['FlightNumbeR'] = df_frame48['Dept Sta'].apply(lambda x: x[1:8])
    df_frame48['Dept Sta'] = df_frame48['Arvl-Hora'].apply(lambda x: x[0:3])
    df_frame48['Arvl Sta'] = df_frame48['Dept-Hora'].apply(lambda x: x[4:8])
    df_frame48['Arvl Day'] = df_frame48['Dept Day'].apply(lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023')
    df_frame48['Dept Day'] = df_frame48['Dept Day'].apply(lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023')
    df_frame48['Trilho'] = ''
    df_frame48['Pax'] = df_frame48['Pax-Subfleet'].apply(lambda x: x[0:3])
    df_frame48['Pax-Subfleet'] = df_frame48['Pax-Subfleet'].apply(lambda x: x[3:6])
    df_frame48['Arvl-Hora'] = df_frame48['Arvl-Hora'].apply(lambda x: x[6:8] + ':' + x[8:10])
    df_frame48['Dept-Hora'] = df_frame48['Dept-Hora'].apply(lambda x: x[0:2] + ':' + x[2:4])
    df_frame48 = df_frame48[
        ['Dept Sta', 'Arvl Sta', 'Dept Day', 'Arvl Day', 'Dept-Hora', 'Arvl-Hora', 'Pax-Subfleet', 'FlightNumbeR',
         'FlightNumberArvl', 'Trilho', 'Svc Type', 'Pax', 'Week-Day']]
    p16 = df_frame48
    df_frame49 = df_orig.query('KeyX == 2  & Key == 49 & Kdept == 11 & Karvl == 7').copy()
    df_frame49['FlightNumberArvl'] = df_frame49['Arvl Sta']
    df_frame49['FlightNumbeR'] = df_frame49['Dept Sta'].apply(lambda x: x[1:8])
    df_frame49['Dept Sta'] = df_frame49['Arvl-Hora'].apply(lambda x: x[0:3])
    df_frame49['Arvl Sta'] = df_frame49['Dept-Hora'].apply(lambda x: x[5:8])
    df_frame49['Arvl Day'] = df_frame49['Dept Day'].apply(lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023')
    df_frame49['Dept Day'] = df_frame49['Dept Day'].apply(lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023')
    df_frame49['Trilho'] = ''
    df_frame49['Pax'] = df_frame49['Pax-Subfleet'].apply(lambda x: x[0:3])
    df_frame49['Pax-Subfleet'] = df_frame49['Pax-Subfleet'].apply(lambda x: x[3:6])
    df_frame49['Arvl-Hora'] = df_frame49['Arvl-Hora'].apply(lambda x: x[3:5] + ':' + x[5:7])
    df_frame49['Dept-Hora'] = df_frame49['Dept-Hora'].apply(lambda x: x[0:2] + ':' + x[2:4])
    df_frame49 = df_frame49[
        ['Dept Sta', 'Arvl Sta', 'Dept Day', 'Arvl Day', 'Dept-Hora', 'Arvl-Hora', 'Pax-Subfleet', 'FlightNumbeR',
         'FlightNumberArvl', 'Trilho', 'Svc Type', 'Pax', 'Week-Day']]
    p17 = df_frame49
    df_frame49 = df_orig.query('KeyX == 2  & Key == 49 & Kdept == 7 & Karvl == 7').copy()
    df_frame49['FlightNumberArvl'] = df_frame49['Arvl Sta']
    df_frame49['FlightNumbeR'] = df_frame49['Dept Sta'].apply(lambda x: x[1:8])
    df_frame49['Dept Sta'] = df_frame49['Arvl-Hora'].apply(lambda x: x[0:3])
    df_frame49['Arvl Sta'] = df_frame49['Dept-Hora'].apply(lambda x: x[4:8])
    df_frame49['Arvl Day'] = df_frame49['Dept Day'].apply(lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023')
    df_frame49['Dept Day'] = df_frame49['Dept Day'].apply(lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023')
    df_frame49['Trilho'] = ''
    df_frame49['Pax'] = df_frame49['Pax-Subfleet'].apply(lambda x: x[0:3])
    df_frame49['Pax-Subfleet'] = df_frame49['Pax-Subfleet'].apply(lambda x: x[3:6])
    df_frame49['Arvl-Hora'] = df_frame49['Arvl-Hora'].apply(lambda x: x[3:5] + ':' + x[5:7])
    df_frame49['Dept-Hora'] = df_frame49['Dept-Hora'].apply(lambda x: x[0:2] + ':' + x[2:4])
    df_frame49 = df_frame49[
        ['Dept Sta', 'Arvl Sta', 'Dept Day', 'Arvl Day', 'Dept-Hora', 'Arvl-Hora', 'Pax-Subfleet', 'FlightNumbeR',
         'FlightNumberArvl', 'Trilho', 'Svc Type', 'Pax', 'Week-Day']]
    p18 = df_frame49
    df_frame50 = df_orig.query('KeyX == 2  & Key == 50 & Kdept == 10 & Karvl == 7').copy()
    df_frame50['FlightNumberArvl'] = df_frame50['Arvl Sta']
    df_frame50['FlightNumbeR'] = df_frame50['Dept Sta'].apply(lambda x: x[1:8])
    df_frame50['Dept Sta'] = df_frame50['Arvl-Hora'].apply(lambda x: x[0:3])
    df_frame50['Arvl Sta'] = df_frame50['Dept-Hora'].apply(lambda x: x[4:7])
    df_frame50['Arvl Day'] = df_frame50['Dept Day'].apply(lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023')
    df_frame50['Dept Day'] = df_frame50['Dept Day'].apply(lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023')
    df_frame50['Trilho'] = ''
    df_frame50['Pax'] = df_frame50['Pax-Subfleet'].apply(lambda x: x[0:3])
    df_frame50['Pax-Subfleet'] = df_frame50['Pax-Subfleet'].apply(lambda x: x[3:6])
    df_frame50['Arvl-Hora'] = df_frame50['Arvl-Hora'].apply(lambda x: x[3:5] + ':' + x[5:7])
    df_frame50['Dept-Hora'] = df_frame50['Dept-Hora'].apply(lambda x: x[0:2] + ':' + x[2:4])
    df_frame50 = df_frame50[
        ['Dept Sta', 'Arvl Sta', 'Dept Day', 'Arvl Day', 'Dept-Hora', 'Arvl-Hora', 'Pax-Subfleet', 'FlightNumbeR',
         'FlightNumberArvl', 'Trilho', 'Svc Type', 'Pax', 'Week-Day']]
    p19 = df_frame50
    df_frame50 = df_orig.query('KeyX == 2  & Key == 50 & Kdept == 7 & Karvl == 10').copy()
    df_frame50['FlightNumberArvl'] = df_frame50['Arvl Sta']
    df_frame50['FlightNumbeR'] = df_frame50['Dept Sta'].apply(lambda x: x[1:8])
    df_frame50['Dept Sta'] = df_frame50['Arvl-Hora'].apply(lambda x: x[0:3])
    df_frame50['Arvl Sta'] = df_frame50['Dept-Hora'].apply(lambda x: x[4:7])
    df_frame50['Arvl Day'] = df_frame50['Dept Day'].apply(lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023')
    df_frame50['Dept Day'] = df_frame50['Dept Day'].apply(lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023')
    df_frame50['Trilho'] = ''
    df_frame50['Pax'] = df_frame50['Pax-Subfleet'].apply(lambda x: x[0:3])
    df_frame50['Pax-Subfleet'] = df_frame50['Pax-Subfleet'].apply(lambda x: x[3:6])
    df_frame50['Arvl-Hora'] = df_frame50['Arvl-Hora'].apply(lambda x: x[6:8] + ':' + x[8:10])
    df_frame50['Dept-Hora'] = df_frame50['Dept-Hora'].apply(lambda x: x[0:2] + ':' + x[2:4])
    df_frame50 = df_frame50[
        ['Dept Sta', 'Arvl Sta', 'Dept Day', 'Arvl Day', 'Dept-Hora', 'Arvl-Hora', 'Pax-Subfleet', 'FlightNumbeR',
         'FlightNumberArvl', 'Trilho', 'Svc Type', 'Pax', 'Week-Day']]
    p20 = df_frame50
    df_frame51 = df_orig.query('KeyX == 2  & Key == 51 & Kdept == 11 & Karvl == 7').copy()
    df_frame51['FlightNumberArvl'] = df_frame51['Arvl Sta']
    df_frame51['FlightNumbeR'] = df_frame51['Dept Sta'].apply(lambda x: x[1:8])
    df_frame51['Dept Sta'] = df_frame51['Arvl-Hora'].apply(lambda x: x[0:3])
    df_frame51['Arvl Sta'] = df_frame51['Dept-Hora'].apply(lambda x: x[5:8])
    df_frame51['Arvl Day'] = df_frame51['Dept Day'].apply(lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023')
    df_frame51['Dept Day'] = df_frame51['Dept Day'].apply(lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023')
    df_frame51['Trilho'] = ''
    df_frame51['Pax'] = df_frame51['Pax-Subfleet'].apply(lambda x: x[0:3])
    df_frame51['Pax-Subfleet'] = df_frame51['Pax-Subfleet'].apply(lambda x: x[3:6])
    df_frame51['Arvl-Hora'] = df_frame51['Arvl-Hora'].apply(lambda x: x[3:5] + ':' + x[5:7])
    df_frame51['Dept-Hora'] = df_frame51['Dept-Hora'].apply(lambda x: x[0:2] + ':' + x[2:4])
    df_frame51 = df_frame51[
        ['Dept Sta', 'Arvl Sta', 'Dept Day', 'Arvl Day', 'Dept-Hora', 'Arvl-Hora', 'Pax-Subfleet', 'FlightNumbeR',
         'FlightNumberArvl', 'Trilho', 'Svc Type', 'Pax', 'Week-Day']]
    p21 = df_frame51
    df_frame52 = df_orig.query('KeyX == 2  & Key == 52 & Kdept == 11 & Karvl == 10').copy()
    df_frame52['FlightNumberArvl'] = df_frame52['Arvl Sta']
    df_frame52['FlightNumbeR'] = df_frame52['Dept Sta'].apply(lambda x: x[1:8])
    df_frame52['Dept Sta'] = df_frame52['Arvl-Hora'].apply(lambda x: x[0:3])
    df_frame52['Arvl Sta'] = df_frame52['Dept-Hora'].apply(lambda x: x[5:8])
    df_frame52['Arvl Day'] = df_frame52['Dept Day'].apply(lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023')
    df_frame52['Dept Day'] = df_frame52['Dept Day'].apply(lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023')
    df_frame52['Trilho'] = ''
    df_frame52['Pax'] = df_frame52['Pax-Subfleet'].apply(lambda x: x[0:3])
    df_frame52['Pax-Subfleet'] = df_frame52['Pax-Subfleet'].apply(lambda x: x[3:6])
    df_frame52['Arvl-Hora'] = df_frame52['Arvl-Hora'].apply(lambda x: x[6:8] + ':' + x[8:10])
    df_frame52['Dept-Hora'] = df_frame52['Dept-Hora'].apply(lambda x: x[0:2] + ':' + x[2:4])
    df_frame52 = df_frame52[
        ['Dept Sta', 'Arvl Sta', 'Dept Day', 'Arvl Day', 'Dept-Hora', 'Arvl-Hora', 'Pax-Subfleet', 'FlightNumbeR',
         'FlightNumberArvl', 'Trilho', 'Svc Type', 'Pax', 'Week-Day']]
    p22 = df_frame52
    df_frame52 = df_orig.query('KeyX == 2  & Key == 52 & Kdept == 10 & Karvl == 10').copy()
    df_frame52['FlightNumberArvl'] = df_frame52['Arvl Sta']
    df_frame52['FlightNumbeR'] = df_frame52['Dept Sta'].apply(lambda x: x[1:8])
    df_frame52['Dept Sta'] = df_frame52['Arvl-Hora'].apply(lambda x: x[0:3])
    df_frame52['Arvl Sta'] = df_frame52['Dept-Hora'].apply(lambda x: x[4:7])
    df_frame52['Arvl Day'] = df_frame52['Dept Day'].apply(lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023')
    df_frame52['Dept Day'] = df_frame52['Dept Day'].apply(lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023')
    df_frame52['Trilho'] = ''
    df_frame52['Pax'] = df_frame52['Pax-Subfleet'].apply(lambda x: x[0:3])
    df_frame52['Pax-Subfleet'] = df_frame52['Pax-Subfleet'].apply(lambda x: x[3:6])
    df_frame52['Arvl-Hora'] = df_frame52['Arvl-Hora'].apply(lambda x: x[6:8] + ':' + x[8:10])
    df_frame52['Dept-Hora'] = df_frame52['Dept-Hora'].apply(lambda x: x[0:2] + ':' + x[2:4])
    df_frame52 = df_frame52[
        ['Dept Sta', 'Arvl Sta', 'Dept Day', 'Arvl Day', 'Dept-Hora', 'Arvl-Hora', 'Pax-Subfleet', 'FlightNumbeR',
         'FlightNumberArvl', 'Trilho', 'Svc Type', 'Pax', 'Week-Day']]
    p23 = df_frame52
    df_frame53 = df_orig.query('KeyX == 2  & Key == 53 & Kdept == 10 & Karvl == 10').copy()
    df_frame53['FlightNumberArvl'] = df_frame53['Arvl Sta']
    df_frame53['FlightNumbeR'] = df_frame53['Dept Sta'].apply(lambda x: x[1:8])
    df_frame53['Dept Sta'] = df_frame53['Arvl-Hora'].apply(lambda x: x[0:3])
    df_frame53['Arvl Sta'] = df_frame53['Dept-Hora'].apply(lambda x: x[4:7])
    df_frame53['Arvl Day'] = df_frame53['Dept Day'].apply(lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023')
    df_frame53['Dept Day'] = df_frame53['Dept Day'].apply(lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023')
    df_frame53['Trilho'] = ''
    df_frame53['Pax'] = df_frame53['Pax-Subfleet'].apply(lambda x: x[0:3])
    df_frame53['Pax-Subfleet'] = df_frame53['Pax-Subfleet'].apply(lambda x: x[3:6])
    df_frame53['Arvl-Hora'] = df_frame53['Arvl-Hora'].apply(lambda x: x[6:8] + ':' + x[8:10])
    df_frame53['Dept-Hora'] = df_frame53['Dept-Hora'].apply(lambda x: x[0:2] + ':' + x[2:4])
    df_frame53 = df_frame53[
        ['Dept Sta', 'Arvl Sta', 'Dept Day', 'Arvl Day', 'Dept-Hora', 'Arvl-Hora', 'Pax-Subfleet', 'FlightNumbeR',
         'FlightNumberArvl', 'Trilho', 'Svc Type', 'Pax', 'Week-Day']]
    p24 = df_frame53
    df_frame54 = df_orig.query('KeyX == 2  & Key == 54 & Kdept == 11 & Karvl == 10').copy()
    df_frame54['FlightNumberArvl'] = df_frame54['Arvl Sta']
    df_frame54['FlightNumbeR'] = df_frame54['Dept Sta'].apply(lambda x: x[1:8])
    df_frame54['Dept Sta'] = df_frame54['Arvl-Hora'].apply(lambda x: x[0:3])
    df_frame54['Arvl Sta'] = df_frame54['Dept-Hora'].apply(lambda x: x[5:8])
    df_frame54['Arvl Day'] = df_frame54['Dept Day'].apply(lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023')
    df_frame54['Dept Day'] = df_frame54['Dept Day'].apply(lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023')
    df_frame54['Trilho'] = ''
    df_frame54['Pax'] = df_frame54['Pax-Subfleet'].apply(lambda x: x[0:3])
    df_frame54['Pax-Subfleet'] = df_frame54['Pax-Subfleet'].apply(lambda x: x[3:6])
    df_frame54['Arvl-Hora'] = df_frame54['Arvl-Hora'].apply(lambda x: x[6:8] + ':' + x[8:10])
    df_frame54['Dept-Hora'] = df_frame54['Dept-Hora'].apply(lambda x: x[0:2] + ':' + x[2:4])
    df_frame54 = df_frame54[
        ['Dept Sta', 'Arvl Sta', 'Dept Day', 'Arvl Day', 'Dept-Hora', 'Arvl-Hora', 'Pax-Subfleet', 'FlightNumbeR',
         'FlightNumberArvl', 'Trilho', 'Svc Type', 'Pax', 'Week-Day']]
    p25 = df_frame54
    df_frame56 = df_orig.query('KeyX == 2  & Key == 56 & Kdept == 7 & Karvl == 7').copy()
    df_frame56['FlightNumberArvl'] = df_frame56['Arvl Sta']
    df_frame56['FlightNumbeR'] = df_frame56['Dept Sta'].apply(lambda x: x[1:8])
    df_frame56['Dept Sta'] = df_frame56['Arvl-Hora'].apply(lambda x: x[0:3])
    df_frame56['Arvl Sta'] = df_frame56['Dept-Hora'].apply(lambda x: x[4:7])
    df_frame56['Arvl Day'] = df_frame56['Dept Day'].apply(
        lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023' + ' á ' + x[5:7] + '/' + x[7:10] + '/' + '2023')
    df_frame56['Dept Day'] = df_frame56['Dept Day'].apply(
        lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023' + ' á ' + x[5:7] + '/' + x[7:10] + '/' + '2023')
    df_frame56['Trilho'] = ''
    df_frame56['Pax'] = df_frame56['Pax-Subfleet'].apply(lambda x: x[0:3])
    df_frame56['Pax-Subfleet'] = df_frame56['Pax-Subfleet'].apply(lambda x: x[3:6])
    df_frame56['Arvl-Hora'] = df_frame56['Arvl-Hora'].apply(lambda x: x[3:5] + ':' + x[5:7])
    df_frame56['Dept-Hora'] = df_frame56['Dept-Hora'].apply(lambda x: x[0:2] + ':' + x[2:4])
    df_frame56 = df_frame56[
        ['Dept Sta', 'Arvl Sta', 'Dept Day', 'Arvl Day', 'Dept-Hora', 'Arvl-Hora', 'Pax-Subfleet', 'FlightNumbeR',
         'FlightNumberArvl', 'Trilho', 'Svc Type', 'Pax', 'Week-Day']]
    p26 = df_frame56
    df_frame57 = df_orig.query('KeyX == 2  & Key == 57 & Kdept == 8 & Karvl == 7').copy()
    df_frame57['FlightNumberArvl'] = df_frame57['Arvl Sta']
    df_frame57['FlightNumbeR'] = df_frame57['Dept Sta'].apply(lambda x: x[1:8])
    df_frame57['Dept Sta'] = df_frame57['Arvl-Hora'].apply(lambda x: x[0:3])
    df_frame57['Arvl Sta'] = df_frame57['Dept-Hora'].apply(lambda x: x[5:8])
    df_frame57['Arvl Day'] = df_frame57['Dept Day'].apply(
        lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023' + ' á ' + x[5:7] + '/' + x[7:10] + '/' + '2023')
    df_frame57['Dept Day'] = df_frame57['Dept Day'].apply(
        lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023' + ' á ' + x[5:7] + '/' + x[7:10] + '/' + '2023')
    df_frame57['Trilho'] = ''
    df_frame57['Pax'] = df_frame57['Pax-Subfleet'].apply(lambda x: x[0:3])
    df_frame57['Pax-Subfleet'] = df_frame57['Pax-Subfleet'].apply(lambda x: x[3:6])
    df_frame57['Arvl-Hora'] = df_frame57['Arvl-Hora'].apply(lambda x: x[3:5] + ':' + x[5:7])
    df_frame57['Dept-Hora'] = df_frame57['Dept-Hora'].apply(lambda x: x[0:2] + ':' + x[2:4])
    df_frame57 = df_frame57[
        ['Dept Sta', 'Arvl Sta', 'Dept Day', 'Arvl Day', 'Dept-Hora', 'Arvl-Hora', 'Pax-Subfleet', 'FlightNumbeR',
         'FlightNumberArvl', 'Trilho', 'Svc Type', 'Pax', 'Week-Day']]
    p27 = df_frame57
    df_frame58 = df_orig.query('KeyX == 2  & Key == 58 & Kdept == 7 & Karvl == 7').copy()
    df_frame58['FlightNumberArvl'] = df_frame58['Arvl Sta']
    df_frame58['FlightNumbeR'] = df_frame58['Dept Sta'].apply(lambda x: x[1:8])
    df_frame58['Dept Sta'] = df_frame58['Arvl-Hora'].apply(lambda x: x[0:3])
    df_frame58['Arvl Sta'] = df_frame58['Dept-Hora'].apply(lambda x: x[4:8])
    df_frame58['Arvl Day'] = df_frame58['Dept Day'].apply(
        lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023' + ' á ' + x[5:7] + '/' + x[7:10] + '/' + '2023')
    df_frame58['Dept Day'] = df_frame58['Dept Day'].apply(
        lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023' + ' á ' + x[5:7] + '/' + x[7:10] + '/' + '2023')
    df_frame58['Trilho'] = ''
    df_frame58['Pax'] = df_frame58['Pax-Subfleet'].apply(lambda x: x[0:3])
    df_frame58['Pax-Subfleet'] = df_frame58['Pax-Subfleet'].apply(lambda x: x[3:6])
    df_frame58['Arvl-Hora'] = df_frame58['Arvl-Hora'].apply(lambda x: x[3:5] + ':' + x[5:7])
    df_frame58['Dept-Hora'] = df_frame58['Dept-Hora'].apply(lambda x: x[0:2] + ':' + x[2:4])
    df_frame58 = df_frame58[
        ['Dept Sta', 'Arvl Sta', 'Dept Day', 'Arvl Day', 'Dept-Hora', 'Arvl-Hora', 'Pax-Subfleet', 'FlightNumbeR',
         'FlightNumberArvl', 'Trilho', 'Svc Type', 'Pax', 'Week-Day']]
    p28 = df_frame58
    df_frame59 = df_orig.query('KeyX == 2  & Key == 59 & Kdept == 10 & Karvl == 7').copy()
    df_frame59['FlightNumberArvl'] = df_frame59['Arvl Sta']
    df_frame59['FlightNumbeR'] = df_frame59['Dept Sta'].apply(lambda x: x[1:8])
    df_frame59['Dept Sta'] = df_frame59['Arvl-Hora'].apply(lambda x: x[0:3])
    df_frame59['Arvl Sta'] = df_frame59['Dept-Hora'].apply(lambda x: x[4:7])
    df_frame59['Arvl Day'] = df_frame59['Dept Day'].apply(
        lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023' + ' á ' + x[5:7] + '/' + x[7:10] + '/' + '2023')
    df_frame59['Dept Day'] = df_frame59['Dept Day'].apply(
        lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023' + ' á ' + x[5:7] + '/' + x[7:10] + '/' + '2023')
    df_frame59['Trilho'] = ''
    df_frame59['Pax'] = df_frame59['Pax-Subfleet'].apply(lambda x: x[0:3])
    df_frame59['Pax-Subfleet'] = df_frame59['Pax-Subfleet'].apply(lambda x: x[3:6])
    df_frame59['Arvl-Hora'] = df_frame59['Arvl-Hora'].apply(lambda x: x[3:5] + ':' + x[5:7])
    df_frame59['Dept-Hora'] = df_frame59['Dept-Hora'].apply(lambda x: x[0:2] + ':' + x[2:4])
    df_frame59 = df_frame59[
        ['Dept Sta', 'Arvl Sta', 'Dept Day', 'Arvl Day', 'Dept-Hora', 'Arvl-Hora', 'Pax-Subfleet', 'FlightNumbeR',
         'FlightNumberArvl', 'Trilho', 'Svc Type', 'Pax', 'Week-Day']]
    p29 = df_frame59
    df_frame59 = df_orig.query('KeyX == 2  & Key == 59 & Kdept == 8 & Karvl == 7').copy()
    df_frame59['FlightNumberArvl'] = df_frame59['Arvl Sta']
    df_frame59['FlightNumbeR'] = df_frame59['Dept Sta'].apply(lambda x: x[1:8])
    df_frame59['Dept Sta'] = df_frame59['Arvl-Hora'].apply(lambda x: x[0:3])
    df_frame59['Arvl Sta'] = df_frame59['Dept-Hora'].apply(lambda x: x[5:8])
    df_frame59['Arvl Day'] = df_frame59['Dept Day'].apply(
        lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023' + ' á ' + x[5:7] + '/' + x[7:10] + '/' + '2023')
    df_frame59['Dept Day'] = df_frame59['Dept Day'].apply(
        lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023' + ' á ' + x[5:7] + '/' + x[7:10] + '/' + '2023')
    df_frame59['Trilho'] = ''
    df_frame59['Pax'] = df_frame59['Pax-Subfleet'].apply(lambda x: x[0:3])
    df_frame59['Pax-Subfleet'] = df_frame59['Pax-Subfleet'].apply(lambda x: x[3:6])
    df_frame59['Arvl-Hora'] = df_frame59['Arvl-Hora'].apply(lambda x: x[3:5] + ':' + x[5:7])
    df_frame59['Dept-Hora'] = df_frame59['Dept-Hora'].apply(lambda x: x[0:2] + ':' + x[2:4])
    df_frame59 = df_frame59[
        ['Dept Sta', 'Arvl Sta', 'Dept Day', 'Arvl Day', 'Dept-Hora', 'Arvl-Hora', 'Pax-Subfleet', 'FlightNumbeR',
         'FlightNumberArvl', 'Trilho', 'Svc Type', 'Pax', 'Week-Day']]
    p30 = df_frame59
    df_frame60 = df_orig.query('KeyX == 2  & Key == 60 & Kdept == 11 & Karvl == 7').copy()
    df_frame60['FlightNumberArvl'] = df_frame60['Arvl Sta']
    df_frame60['FlightNumbeR'] = df_frame60['Dept Sta'].apply(lambda x: x[1:8])
    df_frame60['Dept Sta'] = df_frame60['Arvl-Hora'].apply(lambda x: x[0:3])
    df_frame60['Arvl Sta'] = df_frame60['Dept-Hora'].apply(lambda x: x[5:8])
    df_frame60['Arvl Day'] = df_frame60['Dept Day'].apply(
        lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023' + ' á ' + x[5:7] + '/' + x[7:10] + '/' + '2023')
    df_frame60['Dept Day'] = df_frame60['Dept Day'].apply(
        lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023' + ' á ' + x[5:7] + '/' + x[7:10] + '/' + '2023')
    df_frame60['Trilho'] = ''
    df_frame60['Pax'] = df_frame60['Pax-Subfleet'].apply(lambda x: x[0:3])
    df_frame60['Pax-Subfleet'] = df_frame60['Pax-Subfleet'].apply(lambda x: x[3:6])
    df_frame60['Arvl-Hora'] = df_frame60['Arvl-Hora'].apply(lambda x: x[3:5] + ':' + x[5:7])
    df_frame60['Dept-Hora'] = df_frame60['Dept-Hora'].apply(lambda x: x[0:2] + ':' + x[2:4])
    df_frame60 = df_frame60[
        ['Dept Sta', 'Arvl Sta', 'Dept Day', 'Arvl Day', 'Dept-Hora', 'Arvl-Hora', 'Pax-Subfleet', 'FlightNumbeR',
         'FlightNumberArvl', 'Trilho', 'Svc Type', 'Pax', 'Week-Day']]
    p31 = df_frame60
    df_frame60 = df_orig.query('KeyX == 2  & Key == 60 & Kdept == 7 & Karvl == 7').copy()
    df_frame60['FlightNumberArvl'] = df_frame60['Arvl Sta']
    df_frame60['FlightNumbeR'] = df_frame60['Dept Sta'].apply(lambda x: x[1:8])
    df_frame60['Dept Sta'] = df_frame60['Arvl-Hora'].apply(lambda x: x[0:3])
    df_frame60['Arvl Sta'] = df_frame60['Dept-Hora'].apply(lambda x: x[4:8])
    df_frame60['Arvl Day'] = df_frame60['Dept Day'].apply(
        lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023' + ' á ' + x[5:7] + '/' + x[7:10] + '/' + '2023')
    df_frame60['Dept Day'] = df_frame60['Dept Day'].apply(
        lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023' + ' á ' + x[5:7] + '/' + x[7:10] + '/' + '2023')
    df_frame60['Trilho'] = ''
    df_frame60['Pax'] = df_frame60['Pax-Subfleet'].apply(lambda x: x[0:3])
    df_frame60['Pax-Subfleet'] = df_frame60['Pax-Subfleet'].apply(lambda x: x[3:6])
    df_frame60['Arvl-Hora'] = df_frame60['Arvl-Hora'].apply(lambda x: x[3:5] + ':' + x[5:7])
    df_frame60['Dept-Hora'] = df_frame60['Dept-Hora'].apply(lambda x: x[0:2] + ':' + x[2:4])
    df_frame60 = df_frame60[
        ['Dept Sta', 'Arvl Sta', 'Dept Day', 'Arvl Day', 'Dept-Hora', 'Arvl-Hora', 'Pax-Subfleet', 'FlightNumbeR',
         'FlightNumberArvl', 'Trilho', 'Svc Type', 'Pax', 'Week-Day']]
    p32 = df_frame60
    df_frame61 = df_orig.query('KeyX == 2  & Key == 61 & Kdept == 10 & Karvl == 7').copy()
    df_frame61['FlightNumberArvl'] = df_frame61['Arvl Sta']
    df_frame61['FlightNumbeR'] = df_frame61['Dept Sta'].apply(lambda x: x[1:8])
    df_frame61['Dept Sta'] = df_frame61['Arvl-Hora'].apply(lambda x: x[0:3])
    df_frame61['Arvl Sta'] = df_frame61['Dept-Hora'].apply(lambda x: x[4:7])
    df_frame61['Arvl Day'] = df_frame61['Dept Day'].apply(
        lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023' + ' á ' + x[5:7] + '/' + x[7:10] + '/' + '2023')
    df_frame61['Dept Day'] = df_frame61['Dept Day'].apply(
        lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023' + ' á ' + x[5:7] + '/' + x[7:10] + '/' + '2023')
    df_frame61['Trilho'] = ''
    df_frame61['Pax'] = df_frame61['Pax-Subfleet'].apply(lambda x: x[0:3])
    df_frame61['Pax-Subfleet'] = df_frame61['Pax-Subfleet'].apply(lambda x: x[3:6])
    df_frame61['Arvl-Hora'] = df_frame61['Arvl-Hora'].apply(lambda x: x[3:5] + ':' + x[5:7])
    df_frame61['Dept-Hora'] = df_frame61['Dept-Hora'].apply(lambda x: x[0:2] + ':' + x[2:4])
    df_frame61 = df_frame61[
        ['Dept Sta', 'Arvl Sta', 'Dept Day', 'Arvl Day', 'Dept-Hora', 'Arvl-Hora', 'Pax-Subfleet', 'FlightNumbeR',
         'FlightNumberArvl', 'Trilho', 'Svc Type', 'Pax', 'Week-Day']]
    p33 = df_frame61
    df_frame61 = df_orig.query('KeyX == 2  & Key == 61 & Kdept == 7 & Karvl == 10').copy()
    df_frame61['FlightNumberArvl'] = df_frame61['Arvl Sta']
    df_frame61['FlightNumbeR'] = df_frame61['Dept Sta'].apply(lambda x: x[1:8])
    df_frame61['Dept Sta'] = df_frame61['Arvl-Hora'].apply(lambda x: x[0:3])
    df_frame61['Arvl Sta'] = df_frame61['Dept-Hora'].apply(lambda x: x[4:7])
    df_frame61['Arvl Day'] = df_frame61['Dept Day'].apply(
        lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023' + ' á ' + x[5:7] + '/' + x[7:10] + '/' + '2023')
    df_frame61['Dept Day'] = df_frame61['Dept Day'].apply(
        lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023' + ' á ' + x[5:7] + '/' + x[7:10] + '/' + '2023')
    df_frame61['Trilho'] = ''
    df_frame61['Pax'] = df_frame61['Pax-Subfleet'].apply(lambda x: x[0:3])
    df_frame61['Pax-Subfleet'] = df_frame61['Pax-Subfleet'].apply(lambda x: x[3:6])
    df_frame61['Arvl-Hora'] = df_frame61['Arvl-Hora'].apply(lambda x: x[6:8] + ':' + x[8:10])
    df_frame61['Dept-Hora'] = df_frame61['Dept-Hora'].apply(lambda x: x[0:2] + ':' + x[2:4])
    df_frame60 = df_frame61[
        ['Dept Sta', 'Arvl Sta', 'Dept Day', 'Arvl Day', 'Dept-Hora', 'Arvl-Hora', 'Pax-Subfleet', 'FlightNumbeR',
         'FlightNumberArvl', 'Trilho', 'Svc Type', 'Pax', 'Week-Day']]
    p34 = df_frame61
    df_frame62 = df_orig.query('KeyX == 2  & Key == 62 & Kdept == 10 & Karvl == 10').copy()
    df_frame62['FlightNumberArvl'] = df_frame62['Arvl Sta']
    df_frame62['FlightNumbeR'] = df_frame62['Dept Sta'].apply(lambda x: x[1:8])
    df_frame62['Dept Sta'] = df_frame62['Arvl-Hora'].apply(lambda x: x[0:3])
    df_frame62['Arvl Sta'] = df_frame62['Dept-Hora'].apply(lambda x: x[4:7])
    df_frame62['Arvl Day'] = df_frame62['Dept Day'].apply(
        lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023' + ' á ' + x[5:7] + '/' + x[7:10] + '/' + '2023')
    df_frame62['Dept Day'] = df_frame62['Dept Day'].apply(
        lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023' + ' á ' + x[5:7] + '/' + x[7:10] + '/' + '2023')
    df_frame62['Trilho'] = ''
    df_frame62['Pax'] = df_frame62['Pax-Subfleet'].apply(lambda x: x[0:3])
    df_frame62['Pax-Subfleet'] = df_frame62['Pax-Subfleet'].apply(lambda x: x[3:6])
    df_frame62['Arvl-Hora'] = df_frame62['Arvl-Hora'].apply(lambda x: x[6:8] + ':' + x[8:10])
    df_frame62['Dept-Hora'] = df_frame62['Dept-Hora'].apply(lambda x: x[0:2] + ':' + x[2:4])
    df_frame62 = df_frame62[
        ['Dept Sta', 'Arvl Sta', 'Dept Day', 'Arvl Day', 'Dept-Hora', 'Arvl-Hora', 'Pax-Subfleet', 'FlightNumbeR',
         'FlightNumberArvl', 'Trilho', 'Svc Type', 'Pax', 'Week-Day']]
    p35 = df_frame62
    df_frame64 = df_orig.query('KeyX == 2  & Key == 64 & Kdept == 10 & Karvl == 10').copy()
    df_frame64['FlightNumberArvl'] = df_frame64['Arvl Sta']
    df_frame64['FlightNumbeR'] = df_frame64['Dept Sta'].apply(lambda x: x[1:8])
    df_frame64['Dept Sta'] = df_frame64['Arvl-Hora'].apply(lambda x: x[0:3])
    df_frame64['Arvl Sta'] = df_frame64['Dept-Hora'].apply(lambda x: x[4:7])
    df_frame64['Arvl Day'] = df_frame64['Dept Day'].apply(
        lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023' + ' á ' + x[5:7] + '/' + x[7:10] + '/' + '2023')
    df_frame64['Dept Day'] = df_frame64['Dept Day'].apply(
        lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023' + ' á ' + x[5:7] + '/' + x[7:10] + '/' + '2023')
    df_frame64['Trilho'] = ''
    df_frame64['Pax'] = df_frame64['Pax-Subfleet'].apply(lambda x: x[0:3])
    df_frame64['Pax-Subfleet'] = df_frame64['Pax-Subfleet'].apply(lambda x: x[3:6])
    df_frame64['Arvl-Hora'] = df_frame64['Arvl-Hora'].apply(lambda x: x[6:8] + ':' + x[8:10])
    df_frame64['Dept-Hora'] = df_frame64['Dept-Hora'].apply(lambda x: x[0:2] + ':' + x[2:4])
    df_frame64 = df_frame64[
        ['Dept Sta', 'Arvl Sta', 'Dept Day', 'Arvl Day', 'Dept-Hora', 'Arvl-Hora', 'Pax-Subfleet', 'FlightNumbeR',
         'FlightNumberArvl', 'Trilho', 'Svc Type', 'Pax', 'Week-Day']]
    p36 = df_frame64
    df_frame64 = df_orig.query('KeyX == 2  & Key == 64 & Kdept == 11 & Karvl == 7').copy()
    df_frame64['FlightNumberArvl'] = df_frame64['Arvl Sta']
    df_frame64['FlightNumbeR'] = df_frame64['Dept Sta'].apply(lambda x: x[1:8])
    df_frame64['Dept Sta'] = df_frame64['Arvl-Hora'].apply(lambda x: x[0:3])
    df_frame64['Arvl Sta'] = df_frame64['Dept-Hora'].apply(lambda x: x[5:8])
    df_frame64['Arvl Day'] = df_frame64['Dept Day'].apply(
        lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023' + ' á ' + x[5:7] + '/' + x[7:10] + '/' + '2023')
    df_frame64['Dept Day'] = df_frame64['Dept Day'].apply(
        lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023' + ' á ' + x[5:7] + '/' + x[7:10] + '/' + '2023')
    df_frame64['Trilho'] = ''
    df_frame64['Pax'] = df_frame64['Pax-Subfleet'].apply(lambda x: x[0:3])
    df_frame64['Pax-Subfleet'] = df_frame64['Pax-Subfleet'].apply(lambda x: x[3:6])
    df_frame64['Arvl-Hora'] = df_frame64['Arvl-Hora'].apply(lambda x: x[3:5] + ':' + x[5:7])
    df_frame64['Dept-Hora'] = df_frame64['Dept-Hora'].apply(lambda x: x[0:2] + ':' + x[2:4])
    df_frame64 = df_frame64[
        ['Dept Sta', 'Arvl Sta', 'Dept Day', 'Arvl Day', 'Dept-Hora', 'Arvl-Hora', 'Pax-Subfleet', 'FlightNumbeR',
         'FlightNumberArvl', 'Trilho', 'Svc Type', 'Pax', 'Week-Day']]
    p37 = df_frame64
    df_frame65 = df_orig.query('KeyX == 2  & Key == 65 & Kdept == 11 & Karvl == 10').copy()
    df_frame65['FlightNumberArvl'] = df_frame65['Arvl Sta']
    df_frame65['FlightNumbeR'] = df_frame65['Dept Sta'].apply(lambda x: x[1:8])
    df_frame65['Dept Sta'] = df_frame65['Arvl-Hora'].apply(lambda x: x[0:3])
    df_frame65['Arvl Sta'] = df_frame65['Dept-Hora'].apply(lambda x: x[5:8])
    df_frame65['Arvl Day'] = df_frame65['Dept Day'].apply(
        lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023' + ' á ' + x[5:7] + '/' + x[7:10] + '/' + '2023')
    df_frame65['Dept Day'] = df_frame65['Dept Day'].apply(
        lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023' + ' á ' + x[5:7] + '/' + x[7:10] + '/' + '2023')
    df_frame65['Trilho'] = ''
    df_frame65['Pax'] = df_frame65['Pax-Subfleet'].apply(lambda x: x[0:3])
    df_frame65['Pax-Subfleet'] = df_frame65['Pax-Subfleet'].apply(lambda x: x[3:6])
    df_frame65['Arvl-Hora'] = df_frame65['Arvl-Hora'].apply(lambda x: x[6:8] + ':' + x[8:10])
    df_frame65['Dept-Hora'] = df_frame65['Dept-Hora'].apply(lambda x: x[0:2] + ':' + x[2:4])
    df_frame65 = df_frame65[
        ['Dept Sta', 'Arvl Sta', 'Dept Day', 'Arvl Day', 'Dept-Hora', 'Arvl-Hora', 'Pax-Subfleet', 'FlightNumbeR',
         'FlightNumberArvl', 'Trilho', 'Svc Type', 'Pax', 'Week-Day']]
    p38 = df_frame65
    df_frame66 = df_orig.query('KeyX == 2  & Key == 67 & Kdept == 11 & Karvl == 10').copy()
    df_frame66['FlightNumberArvl'] = df_frame66['Arvl Sta']
    df_frame66['FlightNumbeR'] = df_frame66['Dept Sta'].apply(lambda x: x[1:8])
    df_frame66['Dept Sta'] = df_frame66['Arvl-Hora'].apply(lambda x: x[0:3])
    df_frame66['Arvl Sta'] = df_frame66['Dept-Hora'].apply(lambda x: x[5:8])
    df_frame66['Arvl Day'] = df_frame66['Dept Day'].apply(
        lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023' + ' á ' + x[5:7] + '/' + x[7:10] + '/' + '2023')
    df_frame66['Dept Day'] = df_frame66['Dept Day'].apply(
        lambda x: x[0:2] + '/' + x[2:5] + '/' + '2023' + ' á ' + x[5:7] + '/' + x[7:10] + '/' + '2023')
    df_frame66['Trilho'] = ''
    df_frame66['Pax'] = df_frame66['Pax-Subfleet'].apply(lambda x: x[0:3])
    df_frame66['Pax-Subfleet'] = df_frame66['Pax-Subfleet'].apply(lambda x: x[3:6])
    df_frame66['Arvl-Hora'] = df_frame66['Arvl-Hora'].apply(lambda x: x[6:8] + ':' + x[8:10])
    df_frame66['Dept-Hora'] = df_frame66['Dept-Hora'].apply(lambda x: x[0:2] + ':' + x[2:4])
    df_frame66 = df_frame66[
        ['Dept Sta', 'Arvl Sta', 'Dept Day', 'Arvl Day', 'Dept-Hora', 'Arvl-Hora', 'Pax-Subfleet', 'FlightNumbeR',
         'FlightNumberArvl', 'Trilho', 'Svc Type', 'Pax', 'Week-Day']]
    p39 = df_frame66
    df_concat1 = pd.concat([p1, p2, p3, p40, p5, p6, p7, p8, p9, p10])
    df_concat2 = pd.concat([p11, p12, p13, p14, p15, p16, p17, p18, p19, p20])
    df_concat3 = pd.concat([p21, p22, p23, p24, p25, p26, p27, p28, p29, p30])
    df_concat4 = pd.concat([p31, p32, p33, p34, p35, p36, p37, p38, p39])
    df_geral = pd.concat([df_concat1, df_concat2, df_concat3, df_concat4])
    df_geral.to_excel(f'SIR - MALHA {datetime.date.today()}.xlsx')
    ler = load_workbook(f'SIR - MALHA {datetime.date.today()}.xlsx')
    planilha = ler
    planilha.active.delete_cols(15, 19)
    planilha.active.delete_cols(1)
    ler.save(f'SIR - MALHA {datetime.date.today()}.xlsx')

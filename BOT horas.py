# coding: latin1
#encoding: utf-8

import schedule
import time
from pynput.keyboard import Key, Controller
import time
import datetime
import pyautogui
import pandas as pd
import os
import numpy as np
import win32com.client

tabela = pd.read_csv(r'C:\Users\bmira\Downloads\Relatorios\transferir.csv', sep='~')
tabela2 = tabela.dropna()
tabela2['Elapsed Time-Date'] = tabela['Elapsed Time-Date'].str[:10]
tabela2 =     tabela2 = tabela2 = tabela2.sort_values('Created By')
tabela2 = tabela2.drop(tabela.loc[tabela2['Elapsed Time-Duration '] == 'sem informacoes'].index , axis=0, inplace=False)
tabela2.to_excel(r'C:\Users\bmira\Downloads\Relatorios\Relatorio de horas - Leonardo Lisboa.xlsx')

# -*- coding: utf-8 -*-
"""
Created on Tue Dec 11 15:45:47 2018

@author: jhoparra
"""

import pandas as pd
import os

#os.chdir("C:/Users/jhoparra/OneDrive - Aerovias del Continente Americano S.A. AVIANCA/Documentos/Data/Monitoreo/Passwords/Periodo 3")
os.chdir("D:/Users/jhoparra/OneDrive - Aerovias del Continente Americano S.A. AVIANCA/Documentos/Data/Monitoreo/Passwords/Periodo 3")

cracked = pd.read_table("Resultados LM/cracked_LM.csv",sep = "|",na_values = "NA",keep_default_na = False,encoding = "iso-8859-1")
palabras_claves_VP = pd.read_table("Resultados LM/palabras_claves_LM.csv",sep = "|",na_values = "NA",keep_default_na = False,encoding = "iso-8859-1")
tipo_usuario_r = pd.read_table("Resultados LM/tipo_usuario_r_LM.csv",sep = "|",na_values = "NA",keep_default_na = False,encoding = "iso-8859-1")
politica_90 = pd.read_table("Resultados LM/politica_90_LM.csv",sep = "|",na_values = "NA",keep_default_na = False,encoding = "iso-8859-1")
expiran = pd.read_table("Resultados LM/expiran_LM.csv",sep = "|",na_values = "NA",keep_default_na = False,encoding = "iso-8859-1")
cambio_inicial = pd.read_table("Resultados LM/cambio_inicial_LM.csv",sep = "|",na_values = "NA",keep_default_na = False,encoding = "iso-8859-1")
complejidad_r = pd.read_table("Resultados LM/complejidad_r_LM.csv",sep = "|",na_values = "NA",keep_default_na = False,encoding = "iso-8859-1")
pwned = pd.read_table("Resultados LM/pwned.csv",sep = "|",na_values = "NA",keep_default_na = False,encoding = "iso-8859-1")

out = pd.ExcelWriter("Resultados LM/Resultados LM.xlsx")
cracked.to_excel(out,sheet_name = "Cracked",index = False, encoding = "iso-8859-1")
palabras_claves_VP.to_excel(out,sheet_name = "Wordcloud VP",index = False, encoding = "iso-8859-1")
tipo_usuario_r.to_excel(out,sheet_name = "Tipos de cuentas",index = False, encoding = "iso-8859-1")
politica_90.to_excel(out,sheet_name = "Política de cambio",index = False, encoding = "iso-8859-1")
expiran.to_excel(out,sheet_name = "No expiran",index = False, encoding = "iso-8859-1")
cambio_inicial.to_excel(out,sheet_name = "Cambio Inicial",index = False, encoding = "iso-8859-1")
complejidad_r.to_excel(out,sheet_name = "Complejidad de passw",index = False, encoding = "iso-8859-1")
pwned.to_excel(out,sheet_name = "Pwned",index = False, encoding = "iso-8859-1")

out.save()
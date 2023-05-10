- 👋 Hi, I’m @Jeffersonfer07
- 👀 I’m interested in ...
- 🌱 I’m currently learning ...
- 💞️ I’m looking to collaborate on ...
- 📫 How to reach me ...

<!---
Jeffersonfer07/Jeffersonfer07 is a ✨ special ✨ repository because its `README.md` (this file) appears on your GitHub profile.
You can click the Preview link to take a look at your changes.
--->

import pyautogui
import pyperclip
import datetime
import time
import os
import pandas as pd
import win32com.client as win32
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, Image
from reportlab.pdfbase import pdfmetrics
import webbrowser
import selenium
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import PySimpleGUI as sg


try:
    df_envio = pd.read_excel('db_fast.xlsx')
    cod0 = df_envio.loc[0, 'codigo']
    cnpj0 = df_envio.loc[0, 'cnpj']
    empresa0 = df_envio.loc[0, 'empresa']
    valor0 = df_envio.loc[0, 'fatura']
    email0 = df_envio.loc[0, 'email']
    cod1 = df_envio.loc[1, 'codigo']
    cnpj1 = df_envio.loc[1, 'cnpj']
    empresa1 = df_envio.loc[1, 'empresa']
    valor1 = df_envio.loc[1, 'fatura']
    email1 = df_envio.loc[1, 'email']
    cod2 = df_envio.loc[2, 'codigo']
    cnpj2 = df_envio.loc[2, 'cnpj']
    empresa2 = df_envio.loc[2, 'empresa']
    valor2 = df_envio.loc[2, 'fatura']
    email2 = df_envio.loc[2, 'email']
    cod3 = df_envio.loc[3, 'codigo']
    cnpj3 = df_envio.loc[3, 'cnpj']
    empresa3 = df_envio.loc[3, 'empresa']
    valor3 = df_envio.loc[3, 'fatura']
    email3 = df_envio.loc[3, 'email']
    cod4 = df_envio.loc[4, 'codigo']
    cnpj4 = df_envio.loc[4, 'cnpj']
    empresa4 = df_envio.loc[4, 'empresa']
    valor4 = df_envio.loc[4, 'fatura']
    email4 = df_envio.loc[4, 'email']
    cod5 = df_envio.loc[5, 'codigo']
    cnpj5 = df_envio.loc[5, 'cnpj']
    empresa5 = df_envio.loc[5, 'empresa']
    valor5 = df_envio.loc[5, 'fatura']
    email5 = df_envio.loc[5, 'email']
    cod6 = df_envio.loc[6, 'codigo']
    cnpj6 = df_envio.loc[6, 'cnpj']
    empresa6 = df_envio.loc[6, 'empresa']
    valor6 = df_envio.loc[6, 'fatura']
    email6 = df_envio.loc[6, 'email']
    cod7 = df_envio.loc[7, 'codigo']
    cnpj7 = df_envio.loc[7, 'cnpj']
    empresa7 = df_envio.loc[7, 'empresa']
    valor7 = df_envio.loc[7, 'fatura']
    email7 = df_envio.loc[7, 'email']
    cod8 = df_envio.loc[8, 'codigo']
    cnpj8 = df_envio.loc[8, 'cnpj']
    empresa8 = df_envio.loc[8, 'empresa']
    valor8 = df_envio.loc[8, 'fatura']
    email8 = df_envio.loc[8, 'email']
    cod9 = df_envio.loc[9, 'codigo']
    cnpj9 = df_envio.loc[9, 'cnpj']
    empresa9 = df_envio.loc[9, 'empresa']
    valor9 = df_envio.loc[9, 'fatura']
    email9 = df_envio.loc[9, 'email']
    cod10 = df_envio.loc[10, 'codigo']
    cnpj10 = df_envio.loc[10, 'cnpj']
    empresa10 = df_envio.loc[10, 'empresa']
    valor10 = df_envio.loc[10, 'fatura']
    email10 = df_envio.loc[10, 'email']
    cod11 = df_envio.loc[11, 'codigo']
    cnpj11 = df_envio.loc[11, 'cnpj']
    empresa11 = df_envio.loc[11, 'empresa']
    valor11 = df_envio.loc[11, 'fatura']
    email11 = df_envio.loc[11, 'email']
    cod12 = df_envio.loc[12, 'codigo']
    cnpj12 = df_envio.loc[12, 'cnpj']
    empresa12 = df_envio.loc[12, 'empresa']
    valor12 = df_envio.loc[12, 'fatura']
    email12 = df_envio.loc[12, 'email']
    cod13 = df_envio.loc[13, 'codigo']
    cnpj13 = df_envio.loc[13, 'cnpj']
    empresa13 = df_envio.loc[13, 'empresa']
    valor13 = df_envio.loc[13, 'fatura']
    email13 = df_envio.loc[13, 'email']
    cod14 = df_envio.loc[14, 'codigo']
    cnpj14 = df_envio.loc[14, 'cnpj']
    empresa14 = df_envio.loc[14, 'empresa']
    valor14 = df_envio.loc[14, 'fatura']
    email14 = df_envio.loc[14, 'email']
    cod15 = df_envio.loc[15, 'codigo']
    cnpj15 = df_envio.loc[15, 'cnpj']
    empresa15 = df_envio.loc[15, 'empresa']
    valor15 = df_envio.loc[15, 'fatura']
    email15 = df_envio.loc[15, 'email']
    cod16 = df_envio.loc[16, 'codigo']
    cnpj16 = df_envio.loc[16, 'cnpj']
    empresa16 = df_envio.loc[16, 'empresa']
    valor16 = df_envio.loc[16, 'fatura']
    email16 = df_envio.loc[16, 'email']
    cod17 = df_envio.loc[17, 'codigo']
    cnpj17 = df_envio.loc[17, 'cnpj']
    empresa17 = df_envio.loc[17, 'empresa']
    valor17 = df_envio.loc[17, 'fatura']
    email17 = df_envio.loc[17, 'email']
    cod18 = df_envio.loc[18, 'codigo']
    cnpj18 = df_envio.loc[18, 'cnpj']
    empresa18 = df_envio.loc[18, 'empresa']
    valor18 = df_envio.loc[18, 'fatura']
    email18 = df_envio.loc[18, 'email']
    cod19 = df_envio.loc[19, 'codigo']
    cnpj19 = df_envio.loc[19, 'cnpj']
    empresa19 = df_envio.loc[19, 'empresa']
    valor19 = df_envio.loc[19, 'fatura']
    email19 = df_envio.loc[19, 'email']
    cod20 = df_envio.loc[20, 'codigo']
    cnpj20 = df_envio.loc[20, 'cnpj']
    empresa20 = df_envio.loc[20, 'empresa']
    valor20 = df_envio.loc[20, 'fatura']
    email20 = df_envio.loc[20, 'email']
    cod21 = df_envio.loc[21, 'codigo']
    cnpj21 = df_envio.loc[21, 'cnpj']
    empresa21 = df_envio.loc[21, 'empresa']
    valor21 = df_envio.loc[21, 'fatura']
    email21 = df_envio.loc[21, 'email']
    cod22 = df_envio.loc[22, 'codigo']
    cnpj22 = df_envio.loc[22, 'cnpj']
    empresa22 = df_envio.loc[22, 'empresa']
    valor22 = df_envio.loc[22, 'fatura']
    email22 = df_envio.loc[22, 'email']
    cod23 = df_envio.loc[23, 'codigo']
    cnpj23 = df_envio.loc[23, 'cnpj']
    empresa23 = df_envio.loc[23, 'empresa']
    valor23 = df_envio.loc[23, 'fatura']
    email23 = df_envio.loc[23, 'email']
    cod24 = df_envio.loc[24, 'codigo']
    cnpj24 = df_envio.loc[24, 'cnpj']
    empresa24 = df_envio.loc[24, 'empresa']
    valor24 = df_envio.loc[24, 'fatura']
    email24 = df_envio.loc[24, 'email']
    cod25 = df_envio.loc[25, 'codigo']
    cnpj25 = df_envio.loc[25, 'cnpj']
    empresa25 = df_envio.loc[25, 'empresa']
    valor25 = df_envio.loc[25, 'fatura']
    email25 = df_envio.loc[25, 'email']
    cod26 = df_envio.loc[26, 'codigo']
    cnpj26 = df_envio.loc[26, 'cnpj']
    empresa26 = df_envio.loc[26, 'empresa']
    valor26 = df_envio.loc[26, 'fatura']
    email26 = df_envio.loc[26, 'email']
    cod27 = df_envio.loc[27, 'codigo']
    cnpj27 = df_envio.loc[27, 'cnpj']
    empresa27 = df_envio.loc[27, 'empresa']
    valor27 = df_envio.loc[27, 'fatura']
    email27 = df_envio.loc[27, 'email']
    cod28 = df_envio.loc[28, 'codigo']
    cnpj28 = df_envio.loc[28, 'cnpj']
    empresa28 = df_envio.loc[28, 'empresa']
    valor28 = df_envio.loc[28, 'fatura']
    email28 = df_envio.loc[28, 'email']
    cod29 = df_envio.loc[29, 'codigo']
    cnpj29 = df_envio.loc[29, 'cnpj']
    empresa29 = df_envio.loc[29, 'empresa']
    valor29 = df_envio.loc[29, 'fatura']
    email29 = df_envio.loc[29, 'email']
    cod30 = df_envio.loc[30, 'codigo']
    cnpj30 = df_envio.loc[30, 'cnpj']
    empresa30 = df_envio.loc[30, 'empresa']
    valor30 = df_envio.loc[30, 'fatura']
    email30 = df_envio.loc[30, 'email']
    cod31 = df_envio.loc[31, 'codigo']
    cnpj31 = df_envio.loc[31, 'cnpj']
    empresa31 = df_envio.loc[31, 'empresa']
    valor31 = df_envio.loc[31, 'fatura']
    email31 = df_envio.loc[31, 'email']
    cod32 = df_envio.loc[32, 'codigo']
    cnpj32 = df_envio.loc[32, 'cnpj']
    empresa32 = df_envio.loc[32, 'empresa']
    valor32 = df_envio.loc[32, 'fatura']
    email32 = df_envio.loc[32, 'email']
    cod33 = df_envio.loc[33, 'codigo']
    cnpj33 = df_envio.loc[33, 'cnpj']
    empresa33 = df_envio.loc[33, 'empresa']
    valor33 = df_envio.loc[33, 'fatura']
    email33 = df_envio.loc[33, 'email']
    cod34 = df_envio.loc[34, 'codigo']
    cnpj34 = df_envio.loc[34, 'cnpj']
    empresa34 = df_envio.loc[34, 'empresa']
    valor34 = df_envio.loc[34, 'fatura']
    email34 = df_envio.loc[34, 'email']
    cod35 = df_envio.loc[35, 'codigo']
    cnpj35 = df_envio.loc[35, 'cnpj']
    empresa35 = df_envio.loc[35, 'empresa']
    valor35 = df_envio.loc[35, 'fatura']
    email35 = df_envio.loc[35, 'email']
    cod36 = df_envio.loc[36, 'codigo']
    cnpj36 = df_envio.loc[36, 'cnpj']
    empresa36 = df_envio.loc[36, 'empresa']
    valor36 = df_envio.loc[36, 'fatura']
    email36 = df_envio.loc[36, 'email']
    cod37 = df_envio.loc[37, 'codigo']
    cnpj37 = df_envio.loc[37, 'cnpj']
    empresa37 = df_envio.loc[37, 'empresa']
    valor37 = df_envio.loc[37, 'fatura']
    email37 = df_envio.loc[37, 'email']
    cod38 = df_envio.loc[38, 'codigo']
    cnpj38 = df_envio.loc[38, 'cnpj']
    empresa38 = df_envio.loc[38, 'empresa']
    valor38 = df_envio.loc[38, 'fatura']
    email38 = df_envio.loc[38, 'email']
    cod39 = df_envio.loc[39, 'codigo']
    cnpj39 = df_envio.loc[39, 'cnpj']
    empresa39 = df_envio.loc[39, 'empresa']
    valor39 = df_envio.loc[39, 'fatura']
    email39 = df_envio.loc[39, 'email']
    cod40 = df_envio.loc[40, 'codigo']
    cnpj40 = df_envio.loc[40, 'cnpj']
    empresa40 = df_envio.loc[40, 'empresa']
    valor40 = df_envio.loc[40, 'fatura']
    email40 = df_envio.loc[40, 'email']
    cod41 = df_envio.loc[41, 'codigo']
    cnpj41 = df_envio.loc[41, 'cnpj']
    empresa41 = df_envio.loc[41, 'empresa']
    valor41 = df_envio.loc[41, 'fatura']
    email41 = df_envio.loc[41, 'email']
    cod42 = df_envio.loc[42, 'codigo']
    cnpj42 = df_envio.loc[42, 'cnpj']
    empresa42 = df_envio.loc[42, 'empresa']
    valor42 = df_envio.loc[42, 'fatura']
    email42 = df_envio.loc[42, 'email']
    cod43 = df_envio.loc[43, 'codigo']
    cnpj43 = df_envio.loc[43, 'cnpj']
    empresa43 = df_envio.loc[43, 'empresa']
    valor43 = df_envio.loc[43, 'fatura']
    email43 = df_envio.loc[43, 'email']
    cod44 = df_envio.loc[44, 'codigo']
    cnpj44 = df_envio.loc[44, 'cnpj']
    empresa44 = df_envio.loc[44, 'empresa']
    valor44 = df_envio.loc[44, 'fatura']
    email44 = df_envio.loc[44, 'email']
    cod45 = df_envio.loc[45, 'codigo']
    cnpj45 = df_envio.loc[45, 'cnpj']
    empresa45 = df_envio.loc[45, 'empresa']
    valor45 = df_envio.loc[45, 'fatura']
    email45 = df_envio.loc[45, 'email']
    cod46 = df_envio.loc[46, 'codigo']
    cnpj46 = df_envio.loc[46, 'cnpj']
    empresa46 = df_envio.loc[46, 'empresa']
    valor46 = df_envio.loc[46, 'fatura']
    email46 = df_envio.loc[46, 'email']
    cod47 = df_envio.loc[47, 'codigo']
    cnpj47 = df_envio.loc[47, 'cnpj']
    empresa47 = df_envio.loc[47, 'empresa']
    valor47 = df_envio.loc[47, 'fatura']
    email47 = df_envio.loc[47, 'email']
    cod48 = df_envio.loc[48, 'codigo']
    cnpj48 = df_envio.loc[48, 'cnpj']
    empresa48 = df_envio.loc[48, 'empresa']
    valor48 = df_envio.loc[48, 'fatura']
    email48 = df_envio.loc[48, 'email']
    cod49 = df_envio.loc[49, 'codigo']
    cnpj49 = df_envio.loc[49, 'cnpj']
    empresa49 = df_envio.loc[49, 'empresa']
    valor49 = df_envio.loc[49, 'fatura']
    email49 = df_envio.loc[49, 'email']
    cod50 = df_envio.loc[50, 'codigo']
    cnpj50 = df_envio.loc[50, 'cnpj']
    empresa50 = df_envio.loc[50, 'empresa']
    valor50 = df_envio.loc[50, 'fatura']
    email50 = df_envio.loc[50, 'email']
    cod51 = df_envio.loc[51, 'codigo']
    cnpj51 = df_envio.loc[51, 'cnpj']
    empresa51 = df_envio.loc[51, 'empresa']
    valor51 = df_envio.loc[51, 'fatura']
    email51 = df_envio.loc[51, 'email']
    cod52 = df_envio.loc[52, 'codigo']
    cnpj52 = df_envio.loc[52, 'cnpj']
    empresa52 = df_envio.loc[52, 'empresa']
    valor52 = df_envio.loc[52, 'fatura']
    email52 = df_envio.loc[52, 'email']
    cod53 = df_envio.loc[53, 'codigo']
    cnpj53 = df_envio.loc[53, 'cnpj']
    empresa53 = df_envio.loc[53, 'empresa']
    valor53 = df_envio.loc[53, 'fatura']
    email53 = df_envio.loc[53, 'email']
    cod54 = df_envio.loc[54, 'codigo']
    cnpj54 = df_envio.loc[54, 'cnpj']
    empresa54 = df_envio.loc[54, 'empresa']
    valor54 = df_envio.loc[54, 'fatura']
    email54 = df_envio.loc[54, 'email']
    cod55 = df_envio.loc[55, 'codigo']
    cnpj55 = df_envio.loc[55, 'cnpj']
    empresa55 = df_envio.loc[55, 'empresa']
    valor55 = df_envio.loc[55, 'fatura']
    email55 = df_envio.loc[55, 'email']
    cod56 = df_envio.loc[56, 'codigo']
    cnpj56 = df_envio.loc[56, 'cnpj']
    empresa56 = df_envio.loc[56, 'empresa']
    valor56 = df_envio.loc[56, 'fatura']
    email56 = df_envio.loc[56, 'email']
    cod57 = df_envio.loc[57, 'codigo']
    cnpj57 = df_envio.loc[57, 'cnpj']
    empresa57 = df_envio.loc[57, 'empresa']
    valor57 = df_envio.loc[57, 'fatura']
    email57 = df_envio.loc[57, 'email']
    cod58 = df_envio.loc[58, 'codigo']
    cnpj58 = df_envio.loc[58, 'cnpj']
    empresa58 = df_envio.loc[58, 'empresa']
    valor58 = df_envio.loc[58, 'fatura']
    email58 = df_envio.loc[58, 'email']
    cod59 = df_envio.loc[59, 'codigo']
    cnpj59 = df_envio.loc[59, 'cnpj']
    empresa59 = df_envio.loc[59, 'empresa']
    valor59 = df_envio.loc[59, 'fatura']
    email59 = df_envio.loc[59, 'email']
    cod60 = df_envio.loc[60, 'codigo']
    cnpj60 = df_envio.loc[60, 'cnpj']
    empresa60 = df_envio.loc[60, 'empresa']
    valor60 = df_envio.loc[60, 'fatura']
    email60 = df_envio.loc[60, 'email']
    cod61 = df_envio.loc[61, 'codigo']
    cnpj61 = df_envio.loc[61, 'cnpj']
    empresa61 = df_envio.loc[61, 'empresa']
    valor61 = df_envio.loc[61, 'fatura']
    email61 = df_envio.loc[61, 'email']
    cod62 = df_envio.loc[62, 'codigo']
    cnpj62 = df_envio.loc[62, 'cnpj']
    empresa62 = df_envio.loc[62, 'empresa']
    valor62 = df_envio.loc[62, 'fatura']
    email62 = df_envio.loc[62, 'email']
    cod63 = df_envio.loc[63, 'codigo']
    cnpj63 = df_envio.loc[63, 'cnpj']
    empresa63 = df_envio.loc[63, 'empresa']
    valor63 = df_envio.loc[63, 'fatura']
    email63 = df_envio.loc[63, 'email']
    cod64 = df_envio.loc[64, 'codigo']
    cnpj64 = df_envio.loc[64, 'cnpj']
    empresa64 = df_envio.loc[64, 'empresa']
    valor64 = df_envio.loc[64, 'fatura']
    email64 = df_envio.loc[64, 'email']
    cod65 = df_envio.loc[65, 'codigo']
    cnpj65 = df_envio.loc[65, 'cnpj']
    empresa65 = df_envio.loc[65, 'empresa']
    valor65 = df_envio.loc[65, 'fatura']
    email65 = df_envio.loc[65, 'email']
    cod66 = df_envio.loc[66, 'codigo']
    cnpj66 = df_envio.loc[66, 'cnpj']
    empresa66 = df_envio.loc[66, 'empresa']
    valor66 = df_envio.loc[66, 'fatura']
    email66 = df_envio.loc[66, 'email']
    cod67 = df_envio.loc[67, 'codigo']
    cnpj67 = df_envio.loc[67, 'cnpj']
    empresa67 = df_envio.loc[67, 'empresa']
    valor67 = df_envio.loc[67, 'fatura']
    email67 = df_envio.loc[67, 'email']
    cod68 = df_envio.loc[68, 'codigo']
    cnpj68 = df_envio.loc[68, 'cnpj']
    empresa68 = df_envio.loc[68, 'empresa']
    valor68 = df_envio.loc[68, 'fatura']
    email68 = df_envio.loc[68, 'email']
    cod69 = df_envio.loc[69, 'codigo']
    cnpj69 = df_envio.loc[69, 'cnpj']
    empresa69 = df_envio.loc[69, 'empresa']
    valor69 = df_envio.loc[69, 'fatura']
    email69 = df_envio.loc[69, 'email']
except:
    sg.popup_error('Erro Ao Coleta Dados', title='Fast System™')
    sg.popup_error('Enter com Contato com o Suporte', title='Fast System™')

try:
    cnpj0 = cnpj0.replace("-","")
    cnpj1 = cnpj1.replace("-","")
    cnpj2 = cnpj2.replace("-","")
    cnpj3 = cnpj3.replace("-","")
    cnpj4 = cnpj4.replace("-","")
    cnpj5 = cnpj5.replace("-","")
    cnpj6 = cnpj6.replace("-","")
    cnpj7 = cnpj7.replace("-","")
    cnpj8 = cnpj8.replace("-","")
    cnpj9 = cnpj9.replace("-","")
    cnpj10 = cnpj10.replace("-","")
    cnpj11 = cnpj11.replace("-","")
    cnpj12 = cnpj12.replace("-","")
    cnpj13 = cnpj13.replace("-","")
    cnpj14 = cnpj14.replace("-","")
    cnpj15 = cnpj15.replace("-","")
    cnpj16 = cnpj16.replace("-","")
    cnpj17 = cnpj17.replace("-","")
    cnpj18 = cnpj18.replace("-","")
    cnpj19 = cnpj19.replace("-","")
    cnpj20 = cnpj20.replace("-","")
    cnpj21 = cnpj21.replace("-","")
    cnpj22 = cnpj22.replace("-","")
    cnpj23 = cnpj23.replace("-","")
    cnpj24 = cnpj24.replace("-","")
    cnpj25 = cnpj25.replace("-","")
    cnpj26 = cnpj26.replace("-","")
    cnpj27 = cnpj27.replace("-","")
    cnpj28 = cnpj28.replace("-","")
    cnpj29 = cnpj29.replace("-","")
    cnpj30 = cnpj30.replace("-","")
    cnpj31 = cnpj31.replace("-","")
    cnpj32 = cnpj32.replace("-","")
    cnpj33 = cnpj33.replace("-","")
    cnpj34 = cnpj34.replace("-","")
    cnpj35 = cnpj35.replace("-","")
    cnpj36 = cnpj36.replace("-","")
    cnpj37 = cnpj37.replace("-","")
    cnpj38 = cnpj38.replace("-","")
    cnpj39 = cnpj39.replace("-","")
    cnpj40 = cnpj40.replace("-","")
    cnpj41 = cnpj41.replace("-","")
    cnpj42 = cnpj42.replace("-","")
    cnpj43 = cnpj43.replace("-","")
    cnpj44 = cnpj44.replace("-","")
    cnpj45 = cnpj45.replace("-","")
    cnpj46 = cnpj46.replace("-","")
    cnpj47 = cnpj47.replace("-","")
    cnpj48 = cnpj48.replace("-","")
    cnpj49 = cnpj49.replace("-","")
    cnpj50 = cnpj50.replace("-","")
    cnpj51 = cnpj51.replace("-","")
    cnpj52 = cnpj52.replace("-","")
    cnpj53 = cnpj53.replace("-","")
    cnpj54 = cnpj54.replace("-","")
    cnpj55 = cnpj55.replace("-","")
    cnpj56 = cnpj56.replace("-","")
    cnpj57 = cnpj57.replace("-","")
    cnpj58 = cnpj58.replace("-","")
    cnpj59 = cnpj59.replace("-","")
    cnpj60 = cnpj60.replace("-","")
    cnpj61 = cnpj61.replace("-","")
    cnpj62 = cnpj62.replace("-","")
    cnpj63 = cnpj63.replace("-","")
    cnpj64 = cnpj64.replace("-","")
    cnpj65 = cnpj65.replace("-","")
    cnpj66 = cnpj66.replace("-","")
    cnpj67 = cnpj67.replace("-","")
    cnpj68 = cnpj68.replace("-","")
    cnpj69 = cnpj69.replace("-","")
except:
     sg.popup_error('Erro ao Converter Do CNPJ',title= 'Fast System™')
     sg.popup_error('Enter com Contato com o Suporte', title='Fast System™')

try:
    valor0 = valor0.replace(',','.')
    valor1 = valor1.replace(',','.')
    valor2 = valor2.replace(',','.')
    valor3 = valor3.replace(',','.')
    valor4 = valor4.replace(',','.')
    valor5 = valor5.replace(',','.')
    valor6 = valor6.replace(',','.')
    valor7 = valor7.replace(',','.')
    valor8 = valor8.replace(',','.')
    valor9 = valor9.replace(',','.')
    valor9 = valor9.replace(',','.')
    valor10 = valor10.replace(',','.')
    valor11 = valor11.replace(',','.')
    valor12 = valor12.replace(',','.')
    valor13 = valor13.replace(',','.')
    valor14 = valor14.replace(',','.')
    valor15 = valor15.replace(',','.')
    valor16 = valor16.replace(',','.')
    valor17 = valor17.replace(',','.')
    valor18 = valor18.replace(',','.')
    valor19 = valor19.replace(',','.')
    valor20 = valor20.replace(',','.')
    valor21 = valor21.replace(',','.')
    valor22 = valor22.replace(',','.')
    valor23 = valor23.replace(',','.')
    valor24 = valor24.replace(',','.')
    valor25 = valor25.replace(',','.')
    valor26 = valor26.replace(',','.')
    valor27 = valor27.replace(',','.')
    valor28 = valor28.replace(',','.')
    valor29 = valor29.replace(',','.')
    valor30 = valor30.replace(',','.')
    valor31 = valor31.replace(',','.')
    valor32 = valor32.replace(',','.')
    valor33 = valor33.replace(',','.')
    valor33 = valor33.replace(',','.')
    valor34 = valor34.replace(',','.')
    valor35 = valor35.replace(',','.')
    valor36 = valor36.replace(',','.')
    valor37 = valor37.replace(',','.')
    valor38 = valor38.replace(',','.')
    valor39 = valor39.replace(',','.')
    valor40 = valor40.replace(',','.')
    valor41 = valor41.replace(',','.')
    valor42 = valor42.replace(',','.')
    valor43 = valor43.replace(',','.')
    valor44 = valor44.replace(',','.')
    valor45 = valor45.replace(',','.')
    valor46 = valor46.replace(',','.')
    valor47 = valor47.replace(',','.')
    valor48 = valor48.replace(',','.')
    valor49 = valor49.replace(',','.')
    valor50 = valor50.replace(',','.')
    valor51 = valor51.replace(',','.')
    valor52 = valor52.replace(',','.')
    valor53 = valor53.replace(',','.')
    valor54 = valor54.replace(',','.')
    valor55 = valor55.replace(',','.')
    valor56 = valor56.replace(',','.')
    valor57 = valor57.replace(',','.')
    valor58 = valor58.replace(',','.')
    valor59 = valor59.replace(',','.')
    valor60 = valor60.replace(',','.')
    valor61 = valor61.replace(',','.')
    valor62 = valor62.replace(',','.')
    valor63 = valor63.replace(',','.')
    valor64 = valor64.replace(',','.')
    valor65 = valor65.replace(',','.')
    valor66 = valor66.replace(',','.')
    valor67 = valor67.replace(',','.')
    valor68 = valor68.replace(',','.')
    valor69 = valor69.replace(',','.')

except:
     sg.popup_error('Erro ao Converter do Valor',title= 'Fast System™')
     sg.popup_error('Enter com Contato com o Suporte', title='Fast System™')

try:
    email0 = email0.casefold()
    email0 = email0.strip()
    email1 = email1.casefold()
    email1 = email1.strip()
    email2 = email2.casefold()
    email2 = email2.strip()
    email3 = email3.casefold()
    email3 = email3.strip()
    email4 = email4.casefold()
    email4 = email4.strip()
    email5 = email5.casefold()
    email5 = email5.strip()
    email6 = email6.casefold()
    email6 = email6.strip()
    email7 = email7.casefold()
    email7 = email7.strip()
    email8 = email8.casefold()
    email8 = email8.strip()
    email9 = email9.casefold()
    email9 = email9.strip()
    email10 = email10.casefold()
    email10 = email10.strip()
    email11 = email11.casefold()
    email11 = email11.strip()
    email13 = email13.casefold()
    email13 = email13.strip()
    email14 = email14.casefold()
    email14 = email14.strip()
    email15 = email15.casefold()
    email15 = email15.strip()
    email16 = email16.casefold()
    email16 = email16.strip()
    email17 = email17.casefold()
    email17 = email17.strip()
    email18 = email18.casefold()
    email18 = email18.strip()
    email19 = email19.casefold()
    email19 = email19.strip()
    email20 = email20.casefold()
    email20 = email20.strip()
    email21 = email21.casefold()
    email21 = email21.strip()
    email22 = email22.casefold()
    email22 = email22.strip()
    email23 = email23.casefold()
    email23 = email23.strip()
    email24 = email24.casefold()
    email24 = email24.strip()
    email25 = email25.casefold()
    email25 = email25.strip()
    email26 = email26.casefold()
    email26 = email26.strip()
    email27 = email27.casefold()
    email27 = email27.strip()
    email28 = email28.casefold()
    email28 = email28.strip()
    email29 = email29.casefold()
    email29 = email29.strip()
    email30 = email30.casefold()
    email30 = email30.strip()
    email31 = email31.casefold()
    email31 = email31.strip()
    email32 = email32.casefold()
    email32 = email32.strip()
    email33 = email33.casefold()
    email33 = email33.strip()
    email34 = email34.casefold()
    email34 = email34.strip()
    email35 = email35.casefold()
    email35 = email35.strip()
    email36 = email36.casefold()
    email36 = email36.strip()
    email37 = email37.casefold()
    email37 = email37.strip()
    email38 = email38.casefold()
    email38 = email38.strip()
    email39 = email39.casefold()
    email39 = email39.strip()
    email40 = email40.casefold()
    email40 = email40.strip()
    email41 = email41.casefold()
    email41 = email41.strip()
    email42 = email42.casefold()
    email42 = email42.strip()
    email43 = email43.casefold()
    email43 = email43.strip()
    email44 = email44.casefold()
    email44 = email44.strip()
    email45 = email45.casefold()
    email45 = email45.strip()
    email46 = email46.casefold()
    email46 = email46.strip()
    email47 = email47.casefold()
    email47 = email47.strip()
    email48 = email48.casefold()
    email48 = email48.strip()
    email49 = email49.casefold()
    email49 = email49.strip()
    email50 = email50.casefold()
    email50 = email50.strip()
    email51 = email51.casefold()
    email51 = email51.strip()
    email52 = email52.casefold()
    email52 = email52.strip()
    email53 = email53.casefold()
    email53 = email53.strip()
    email54 = email54.casefold()
    email54 = email54.strip()
    email55 = email55.casefold()
    email55 = email55.strip()
    email56 = email56.casefold()
    email56 = email56.strip()
    email57 = email57.casefold()
    email57 = email57.strip()
    email58 = email58.casefold()
    email58 = email58.strip()
    email59 = email59.casefold()
    email59 = email59.strip()
    email60 = email60.casefold()
    email61 = email61.strip()
    email61 = email60.casefold()
    email62 = email62.strip()
    email62 = email60.casefold()
    email63 = email63.strip()
    email63 = email60.casefold()
    email64 = email64.strip()
    email64 = email64.casefold()
    email65 = email65.strip()
    email65 = email65.casefold()
    email66 = email66.strip()
    email66 = email66.casefold()
    email67 = email67.strip()
    email67 = email67.casefold()
    email68 = email68.strip()
    email68 = email68.casefold()
    email69 = email69.strip()
    email69 = email69.casefold()

except:
     sg.popup_error('Erro ao Converter E-mail',title= 'Fast System™')
     sg.popup_error('Enter com Contato com o Suporte', title='Fast System™')

try:
    total0 = float(valor0) + float(valor1) + float(valor2) + float(valor3) + float(valor4) + float(valor5) + float(valor6) + float(valor7) + float(valor8) + float(valor9)
    total1 = float(valor10) + float(valor11) + float(valor12) + float(valor13) + float(valor14) + float(valor15) + float(valor16) + float(valor17) + float(valor18) + float(valor19)
    total2 = float(valor20) + float(valor21) + float(valor22) + float(valor23) + float(valor24) + float(valor25) + float(valor26) + float(valor27) + float(valor28) + float(valor29)
    total3 = float(valor30) + float(valor31) + float(valor32) + float(valor33) + float(valor34) + float(valor35) + float(valor36) + float(valor37) + float(valor38) + float(valor39)
    total4 = float(valor40) + float(valor41) + float(valor42) + float(valor43) + float(valor44) + float(valor45) + float(valor46) + float(valor47) + float(valor48) + float(valor49)
    total5 = float(valor50) + float(valor51) + float(valor52) + float(valor53) + float(valor54) + float(valor55) + float(valor56) + float(valor57) + float(valor58) + float(valor59)
    total6 = float(valor60) + float(valor61) + float(valor62) + float(valor63) + float(valor64) + float(valor65) + float(valor66) + float(valor67) + float(valor68) + float(valor69)
    fat_total = total0 + total1 + total2 + total3 + total4 + total5 + total6
except:
    sg.popup_error('Erro Ao Calcular Valor',title= 'Fast System™')
    sg.popup_error('Enter com Contato com o Suporte', title='Fast System™')


def emitir(cnpj,valor,codigo,empresa):
    pyautogui.PAUSE = 1.0
    sg.popup_timed(f'Nota Será Emitida Para: {empresa}', title='Fast System ™')
    time.sleep(2)
    pyautogui.click(x=534, y=334) # Pode Mudar - aba Cnpj
    pyautogui.press('down')
    pyautogui.press('enter')
    time.sleep(3)
    pyautogui.press('tab')
    pyperclip.copy(str(cnpj))
    pyautogui.hotkey('ctrl','v')
    pyautogui.press('tab')
    pyautogui.press('enter')
    time.sleep(3)
    buscar = pyautogui.password(text='Cliente Localizado? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
    while buscar != "1":
        sg.popup_auto_close("Por Favor, Cadaste o Cliente", title='Fast System™')
        time.sleep(10)
        buscar = pyautogui.password(text='Cliente Cadastro? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
    time.sleep(2)
    pyautogui.click(x=1208, y=983) # Pode Mudar - Valores - Descrições
    pyautogui.click(x=534, y=334)  # Clica na Pagina
    pyautogui.press(['tab','tab','tab','tab','tab','tab'])
    pyautogui.press(['down','down','down'])
    pyautogui.press('tab')
    pyautogui.press('down')
    pyautogui.press('tab')
    pyautogui.write('Bahia')
    time.sleep(1)
    pyautogui.press('tab')
    pyautogui.write('RAFAEL')
    time.sleep(1)
    pyautogui.press('tab')
    pyperclip.copy(descricao)
    pyautogui.hotkey('ctrl','v')
    time.sleep(3)
    pyautogui.click(x=1524, y=825)# Pode Mudar - Aba Valores
    time.sleep(3)
    pyautogui.press(['tab','tab','tab'])
    time.sleep(1)
    pyautogui.write(valor)
    time.sleep(1)
    pyautogui.press(['tab','tab'])
    time.sleep(1)
    pyautogui.write(aliquota)
    leitura = pyautogui.password(text='Valor Está Correto? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
    while leitura != "1":
        sg.popup_auto_close("Por Favor, Altere o Valor Manualmente", title='Fast System™')
        time.sleep(10)
        leitura = pyautogui.password(text='Valor Está Correto? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
    time.sleep(1)
    pyautogui.click(x=1560, y=736)# Pode Mudar - Emitir Nota
    time.sleep(1)
    pyautogui.press('enter')
    caminho = r"C:\\"
    pasta_p = caminho + f"\Faturamento Mês {referencia}"
    if (not os.path.exists(pasta_p)):
        os.mkdir(pasta_p)
    pasta_cod = f"C:\Faturamento Mês {referencia}\ {codigo}"
    if (not os.path.exists(pasta_cod)):
        os.mkdir(pasta_cod)
    time.sleep(3)
    try:
        pyautogui.click('salvar.png')
    except:
        pyautogui.click(x=672, y=176)  # Pode Mudar - salvar nota
    time.sleep(2)
    pyautogui.write('Nota Fiscal')
    time.sleep(1)
    pyautogui.press('f4')
    pyautogui.hotkey('ctrl','a')
    pyperclip.copy(pasta_cod)
    pyautogui.hotkey('ctrl','v')
    pyautogui.press('enter')
    pyautogui.click(x=760, y=557) # - Salvar
    salvar = pyautogui.password(text='A Nota Foi Salva? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
    while salvar != "1":
        sg.popup_auto_close("Por Favor, Salve Manualmente.", title='Fast System™')
        salvar = pyautogui.password(text='Tudo OK? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
    time.sleep(1)
    pyautogui.click(x=784, y=17) # Fechar a Aba De baixa
    return sg.popup_timed(f'Nota Fiscal Emitida: {codigo} / {valor}',title='Fast System ™')

def emitir1(cnpj, valor, codigo):
    sg.popup_auto_close("Iniciando a Emissão do Boleto.", title='Fast System™')
    time.sleep(2)
    navegador1.find_element(By.XPATH,'//*[@id="root"]/div/main/div/section/ul/div[3]/div').click()
    time.sleep(3)
    navegador1.find_element(By.XPATH, '//*[@id="searchContacts"]').click()  # Buscar
    time.sleep(2)
    navegador1.find_element(By.XPATH, '//*[@id="searchContacts"]').send_keys(cnpj)
    time.sleep(2)
    pyautogui.click(x=553, y=557)  # Selecionar o Cliente
    buscar = pyautogui.password(text='Cliente Localizado? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
    while buscar != "1":
        navegador1.find_element(By.XPATH, '// *[ @ id = "root"] / div / main / div / div / ul / li / div').click() # Cadastrar
        navegador1.find_element(By.XPATH, '// *[ @ id = "root"] / div / main / div / div / ul / li[2] / div').click() # Juridica
        navegador1.find_element(By.XPATH, '//*[@id="contact-document"]').click() # Cnpj
        navegador1.find_element(By.XPATH, '//*[@id="contact-document"]').send_keys(cnpj)
        navegador1.find_element(By.XPATH, '// *[ @ id = "contact-full-name"]').click()  # nome
        sg.popup_auto_close("Por Favor, Digite o Nome do Cliente", title='Fast System™')
        time.sleep(10)
        navegador1.find_element(By.XPATH, '// *[ @ id = "root"] / div / main / div / form / button').click()
        buscar = pyautogui.password(text='Cliente Cadastro? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
    navegador1.find_element(By.XPATH, '//*[@id="root"]/div/main/div/div/ul/li[1]/div').click()  # Tipo Do Boleto
    time.sleep(1)
    navegador1.find_element(By.XPATH, '//*[@id="bankslip-amount"]').send_keys(valor)
    pyautogui.press('tab')
    time.sleep(1)
    pyautogui.write(vencimento)
    time.sleep(1)
    navegador1.find_element(By.XPATH, '//*[@id="bankslip-name"]').send_keys(descricao)
    navegador1.find_element(By.XPATH, '//*[@id="root"]/div/main/div/form/button').click()  # Avançar
    time.sleep(1)
    navegador1.find_element(By.XPATH, '//*[@id="root"]/div/main/div/form/button').click()  # Avançar
    time.sleep(1)
    # Calculo do boleto
    navegador1.find_element(By.XPATH, '//*[@id="root"]/div/main/div/form/div[2]/div/div[2]/label').click()
    navegador1.find_element(By.XPATH, '//*[@id="root"]/div/main/div/form/div[2]/div/div[3]/div/div/div[2]/label').click()
    time.sleep(1)
    pyautogui.press('tab')
    pyautogui.write("200")
    navegador1.find_element(By.XPATH, '//*[@id="root"]/div/main/div/form/div[3]/div/div[2]/label').click()
    time.sleep(1)
    pyautogui.press('tab')
    pyautogui.write("490")
    navegador1.find_element(By.XPATH, '//*[@id="root"]/div/main/div/form/button').click()  # Avançar
    time.sleep(1)
    navegador1.find_element(By.XPATH, '//*[@id="root"]/div/main/div/section/form/button').click()  # Avançar
    pyautogui.press('tab')
    pyautogui.press('enter')
    time.sleep(3)
    navegador1.find_element(By.XPATH, '//*[@id="root"]/div/main/div/div/div/button').click()  # emitir
    time.sleep(4)
    pyautogui.click(x=154, y=981)
    time.sleep(2)
    pyautogui.click('salvar.png')
    time.sleep(2)
    pyautogui.write('Boleto')
    time.sleep(1)
    pyautogui.press('f4')
    pyautogui.hotkey('ctrl','a')
    pasta_cod = f"C:\Faturamento Mês {referencia}\ {codigo}"
    pyperclip.copy(pasta_cod)
    pyautogui.hotkey('ctrl','v')
    pyautogui.press('enter')
    pyautogui.click(x=760, y=557) # - Salvar
    salvar = pyautogui.password(text='O Boleto  Foi Salvo? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
    while salvar != "1":
        sg.popup_auto_close("Por Favor, Salve Manualmente.", title='Fast System™')
        salvar = pyautogui.password(text='Tudo OK? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
    time.sleep(1)
    pyautogui.click(x=592, y=20)  # Fechar a Guia
    time.sleep(2)
    pyautogui.press('f5')
    return sg.popup_timed(f'Boleto Emitido: {codigo} / {valor}',title='Fast System ™')

def enviar(email, codigo, empresa):
    outlook = win32.Dispatch('outlook.application')
    copia = ('lavadormacaubense04@hotmail.com')
    assunto = f""" 
<p>Segue Lavagens de Veículos Realizadas no Mês: {referencia} <p> 
<p>Cliente: {empresa}<p>
<p>-----------------------------------<p>
<p>Vencimento: {vencimento} <p> 
<p>-----------------------------------<p>
<i> Dados Bancários Para Pagamento<i>
<p>-----------------------------------<p>
<p>Favorecido: Elivando Roque Domingues<p>
<p>CNPJ: 26.909.139/0001-46<p>
<p>Banco Do Brasil <p>
<p>1133-9 | 91.000-7 | Conta Correte<p>
<p>Pix: 26.909.139/0001-46<p>
<p>-----------------------------------<p>

<p> Por Favor. Confirme o Recebimento.<p>

<p>-----------------------------------<p>
<p><i> Faturamento Automático - Fast System™ | Versão 3.0IE.<p><i> 
<p><i>Desenvolvido Pela Empresa Plus Mais Soluções | © Plus Mais Soluções 2023. <p><i>
<p>-----------------------------------<p>
"""

    mail = outlook.CreateItem(0)
    mail.To = (email)
    mail.CC = (copia)
    mail.Subject = ('Faturamento Lavador Macaubense IV')
    mail.HTMLBody = assunto
    attachment0 = r'C:\Faturamento Mês {}\ {}\Nota Fiscal.pdf'.format(referencia, codigo)
    attachment1 = r'C:\Faturamento Mês {}\ {}\Boleto.pdf'.format(referencia, codigo)
    attachment2 = r'C:\Faturamento Mês {}\ {}\{}.pdf'.format(referencia, codigo, codigo)
    mail.Attachments.Add(attachment0)
    mail.Attachments.Add(attachment1)
    mail.Attachments.Add(attachment2)
    mail.Send()
    return sg.popup_timed(f'Conect Database: Fatura Enviada: {empresa}', title='Conexão')


try:
 def rel_fat():
        pdf_file = 'Faturamento.pdf'
        rel = canvas.Canvas(pdf_file)
        rel = canvas.Canvas('Faturamento.pdf')
        rel.setFont('Courier', 18)
        rel.drawString(180, 790, 'Faturamento Mensal')
        rel.setFont('Courier', 12)
        rel.drawString(30, 750, f'Data Emisão: {data}')
        rel.drawString(250, 750, f'Empresa: Lavador Macaubense IV ')
        rel.drawString(30, 730, f'Mês Do Serviço: {referencia} ')
        rel.drawString(250, 730, f'Usuário Solicitante: {loginy} ')
        rel.setFont('Courier', 12)
        rel.drawString(50, 705, 'Codigos')
        rel.drawString(250, 705, 'Clientes')
        rel.drawString(480, 705, 'Valor R$')
        rel.setFont('Courier', 9)
        rel.drawString(50, 690, f'{cod0}'), rel.drawString(130, 690, f'{empresa0}'), rel.drawString(480, 690,
                                                                                                    f'R$ {valor0}')
        rel.drawString(50, 680, f'{cod1}'), rel.drawString(130, 680, f'{empresa1}'), rel.drawString(480, 680,
                                                                                                    f'R$ {valor1}')
        rel.drawString(50, 670, f'{cod2}'), rel.drawString(130, 670, f'{empresa2}'), rel.drawString(480, 670,
                                                                                                    f'R$ {valor2}')
        rel.drawString(50, 660, f'{cod3}'), rel.drawString(130, 660, f'{empresa3}'), rel.drawString(480, 660,
                                                                                                    f'R$ {valor3}')
        rel.drawString(50, 650, f'{cod4}'), rel.drawString(130, 650, f'{empresa4}'), rel.drawString(480, 650,
                                                                                                    f'R$ {valor4}')
        rel.drawString(50, 640, f'{cod5}'), rel.drawString(130, 640, f'{empresa5}'), rel.drawString(480, 640,
                                                                                                    f'R$ {valor5}')
        rel.drawString(50, 630, f'{cod6}'), rel.drawString(130, 630, f'{empresa6}'), rel.drawString(480, 630,
                                                                                                    f'R$ {valor6}')
        rel.drawString(50, 620, f'{cod7}'), rel.drawString(130, 620, f'{empresa7}'), rel.drawString(480, 620,
                                                                                                    f'R$ {valor7}')
        rel.drawString(50, 610, f'{cod8}'), rel.drawString(130, 610, f'{empresa8}'), rel.drawString(480, 610,
                                                                                                    f'R$ {valor8}')
        rel.drawString(50, 600, f'{cod9}'), rel.drawString(130, 600, f'{empresa9}'), rel.drawString(480, 600,
                                                                                                    f'R$ {valor9}')
        rel.drawString(50, 590, f'{cod10}'), rel.drawString(130, 590, f'{empresa10}'), rel.drawString(480, 590,
                                                                                                      f'R$ {valor10}')
        rel.drawString(50, 580, f'{cod11}'), rel.drawString(130, 580, f'{empresa11}'), rel.drawString(480, 580,
                                                                                                      f'R$ {valor11}')
        rel.drawString(50, 570, f'{cod12}'), rel.drawString(130, 570, f'{empresa12}'), rel.drawString(480, 570,
                                                                                                      f'R$ {valor12}')
        rel.drawString(50, 560, f'{cod13}'), rel.drawString(130, 560, f'{empresa13}'), rel.drawString(480, 560,
                                                                                                      f'R$ {valor13}')
        rel.drawString(50, 550, f'{cod14}'), rel.drawString(130, 550, f'{empresa14}'), rel.drawString(480, 550,
                                                                                                      f'R$ {valor14}')
        rel.drawString(50, 540, f'{cod15}'), rel.drawString(130, 540, f'{empresa15}'), rel.drawString(480, 540,
                                                                                                      f'R$ {valor15}')
        rel.drawString(50, 530, f'{cod16}'), rel.drawString(130, 530, f'{empresa16}'), rel.drawString(480, 530,
                                                                                                      f'R$ {valor16}')
        rel.drawString(50, 520, f'{cod17}'), rel.drawString(130, 520, f'{empresa17}'), rel.drawString(480, 520,
                                                                                                      f'R$ {valor17}')
        rel.drawString(50, 510, f'{cod18}'), rel.drawString(130, 510, f'{empresa18}'), rel.drawString(480, 510,
                                                                                                      f'R$ {valor18}')
        rel.drawString(50, 500, f'{cod19}'), rel.drawString(130, 500, f'{empresa19}'), rel.drawString(480, 500,
                                                                                                      f'R$ {valor19}')
        rel.drawString(50, 490, f'{cod20}'), rel.drawString(130, 490, f'{empresa20}'), rel.drawString(480, 490,
                                                                                                      f'R$ {valor20}')
        rel.drawString(50, 480, f'{cod21}'), rel.drawString(130, 480, f'{empresa21}'), rel.drawString(480, 480,
                                                                                                      f'R$ {valor21}')
        rel.drawString(50, 470, f'{cod22}'), rel.drawString(130, 470, f'{empresa22}'), rel.drawString(480, 470,
                                                                                                      f'R$ {valor22}')
        rel.drawString(50, 460, f'{cod23}'), rel.drawString(130, 460, f'{empresa23}'), rel.drawString(480, 460,
                                                                                                      f'R$ {valor23}')
        rel.drawString(50, 450, f'{cod24}'), rel.drawString(130, 450, f'{empresa24}'), rel.drawString(480, 450,
                                                                                                      f'R$ {valor24}')
        rel.drawString(50, 440, f'{cod25}'), rel.drawString(130, 440, f'{empresa25}'), rel.drawString(480, 440,
                                                                                                      f'R$ {valor25}')
        rel.drawString(50, 430, f'{cod26}'), rel.drawString(130, 430, f'{empresa26}'), rel.drawString(480, 430,
                                                                                                      f'R$ {valor26}')
        rel.drawString(50, 420, f'{cod27}'), rel.drawString(130, 420, f'{empresa27}'), rel.drawString(480, 420,
                                                                                                      f'R$ {valor27}')
        rel.drawString(50, 410, f'{cod28}'), rel.drawString(130, 410, f'{empresa28}'), rel.drawString(480, 410,
                                                                                                      f'R$ {valor28}')
        rel.drawString(50, 400, f'{cod29}'), rel.drawString(130, 400, f'{empresa29}'), rel.drawString(480, 400,
                                                                                                      f'R$ {valor29}')
        rel.drawString(50, 390, f'{cod30}'), rel.drawString(130, 390, f'{empresa30}'), rel.drawString(480, 390,
                                                                                                      f'R$ {valor30}')
        rel.drawString(50, 380, f'{cod31}'), rel.drawString(130, 380, f'{empresa31}'), rel.drawString(480, 380,
                                                                                                      f'R$ {valor31}')
        rel.drawString(50, 370, f'{cod32}'), rel.drawString(130, 370, f'{empresa32}'), rel.drawString(480, 370,
                                                                                                      f'R$ {valor32}')
        rel.drawString(50, 360, f'{cod33}'), rel.drawString(130, 360, f'{empresa33}'), rel.drawString(480, 360,
                                                                                                      f'R$ {valor33}')
        rel.drawString(50, 350, f'{cod34}'), rel.drawString(130, 350, f'{empresa34}'), rel.drawString(480, 350,
                                                                                                      f'R$ {valor34}')
        rel.drawString(50, 340, f'{cod35}'), rel.drawString(130, 340, f'{empresa35}'), rel.drawString(480, 340,
                                                                                                      f'R$ {valor35}')
        rel.drawString(50, 330, f'{cod36}'), rel.drawString(130, 330, f'{empresa36}'), rel.drawString(480, 330,
                                                                                                      f'R$ {valor36}')
        rel.drawString(50, 320, f'{cod37}'), rel.drawString(130, 320, f'{empresa37}'), rel.drawString(480, 320,
                                                                                                      f'R$ {valor37}')
        rel.drawString(50, 310, f'{cod38}'), rel.drawString(130, 310, f'{empresa38}'), rel.drawString(480, 310,
                                                                                                      f'R$ {valor38}')
        rel.drawString(50, 300, f'{cod39}'), rel.drawString(130, 300, f'{empresa39}'), rel.drawString(480, 300,
                                                                                                      f'R$ {valor39}')
        rel.drawString(50, 290, f'{cod40}'), rel.drawString(130, 290, f'{empresa40}'), rel.drawString(480, 290,
                                                                                                      f'R$ {valor40}')
        rel.drawString(50, 280, f'{cod41}'), rel.drawString(130, 280, f'{empresa41}'), rel.drawString(480, 280,
                                                                                                      f'R$ {valor41}')
        rel.drawString(50, 270, f'{cod42}'), rel.drawString(130, 270, f'{empresa42}'), rel.drawString(480, 270,
                                                                                                      f'R$ {valor42}')
        rel.drawString(50, 260, f'{cod43}'), rel.drawString(130, 260, f'{empresa43}'), rel.drawString(480, 260,
                                                                                                      f'R$ {valor43}')
        rel.drawString(50, 250, f'{cod44}'), rel.drawString(130, 250, f'{empresa44}'), rel.drawString(480, 250,
                                                                                                      f'R$ {valor44}')
        rel.drawString(50, 240, f'{cod45}'), rel.drawString(130, 240, f'{empresa45}'), rel.drawString(480, 240,
                                                                                                      f'R$ {valor45}')
        rel.drawString(50, 230, f'{cod46}'), rel.drawString(130, 230, f'{empresa46}'), rel.drawString(480, 230,
                                                                                                      f'R$ {valor46}')
        rel.drawString(50, 220, f'{cod47}'), rel.drawString(130, 220, f'{empresa47}'), rel.drawString(480, 220,
                                                                                                      f'R$ {valor47}')
        rel.drawString(50, 210, f'{cod48}'), rel.drawString(130, 210, f'{empresa48}'), rel.drawString(480, 210,
                                                                                                      f'R$ {valor48}')
        rel.drawString(50, 200, f'{cod49}'), rel.drawString(130, 200, f'{empresa49}'), rel.drawString(480, 200,
                                                                                                      f'R$ {valor49}')
        rel.drawString(50, 190, f'{cod50}'), rel.drawString(130, 190, f'{empresa50}'), rel.drawString(480, 190,
                                                                                                      f'R$ {valor50}')
        rel.drawString(50, 180, f'{cod51}'), rel.drawString(130, 180, f'{empresa51}'), rel.drawString(480, 180,
                                                                                                      f'R$ {valor51}')
        rel.drawString(50, 170, f'{cod52}'), rel.drawString(130, 170, f'{empresa52}'), rel.drawString(480, 170,
                                                                                                      f'R$ {valor52}')
        rel.drawString(50, 160, f'{cod53}'), rel.drawString(130, 160, f'{empresa53}'), rel.drawString(480, 160,
                                                                                                      f'R$ {valor53}')
        rel.drawString(50, 150, f'{cod54}'), rel.drawString(130, 150, f'{empresa54}'), rel.drawString(480, 150,
                                                                                                      f'R$ {valor54}')
        rel.drawString(50, 140, f'{cod55}'), rel.drawString(130, 140, f'{empresa55}'), rel.drawString(480, 140,
                                                                                                      f'R$ {valor56}')
        rel.drawString(50, 130, f'{cod56}'), rel.drawString(130, 130, f'{empresa56}'), rel.drawString(480, 130,
                                                                                                      f'R$ {valor56}')
        rel.drawString(50, 120, f'{cod57}'), rel.drawString(130, 120, f'{empresa57}'), rel.drawString(480, 120,
                                                                                                      f'R$ {valor57}')
        rel.drawString(50, 110, f'{cod58}'), rel.drawString(130, 110, f'{empresa58}'), rel.drawString(480, 110,
                                                                                                      f'R$ {valor58}')
        rel.drawString(50, 100, f'{cod59}'), rel.drawString(130, 100, f'{empresa59}'), rel.drawString(480, 100,
                                                                                                      f'R$ {valor59}')
        rel.drawString(50, 90, f'{cod60}'), rel.drawString(130, 90, f'{empresa60}'), rel.drawString(480, 90,
                                                                                                    f'R$ {valor60}')
        rel.drawString(50, 80, f'{cod61}'), rel.drawString(130, 80, f'{empresa61}'), rel.drawString(480, 80,
                                                                                                    f'R$ {valor61}')
        rel.drawString(50, 70, f'{cod62}'), rel.drawString(130, 70, f'{empresa62}'), rel.drawString(480, 70,
                                                                                                    f'R$ {valor62}')
        rel.drawString(50, 60, f'{cod63}'), rel.drawString(130, 60, f'{empresa63}'), rel.drawString(480, 60,
                                                                                                    f'R$ {valor63}')
        rel.drawString(50, 50, f'{cod64}'), rel.drawString(130, 50, f'{empresa64}'), rel.drawString(480, 50,
                                                                                                    f'R$ {valor64}')
        rel.drawString(50, 40, f'{cod65}'), rel.drawString(130, 40, f'{empresa65}'), rel.drawString(480, 40,
                                                                                                    f'R$ {valor65}')
        rel.rect(20, 700, 545, 16, fill=False, stroke=True)  # Linha titulo da tabela
        rel.rect(20, 725, 545, 40, fill=False, stroke=True)  # cabecario 2
        rel.rect(20, 785, 545, 20, fill=False, stroke=True)  # cabecario 1
        rel.rect(20, 35, 545, 680, fill=False, stroke=True)  # tabela
        rel.rect(120, 35, 350, 680, fill=False, stroke=True)  # Linha cabeçario
        rel.setFont('Times-Roman', 6)
        rel.rect(20, 14, 545, 0, fill=True, stroke=False)  # ultima linha
        rel.drawString(20, 6,
                       'Faturamento Automático - Fast System™ | Versão 3.0IE. Desenvolvido Pela Empresa Plus Mais Soluções | © Plus Mais Soluções 2023. ')
        rel.drawString(540, 6, "Pagina 01")
        rel.showPage()

        rel.setFont('Courier', 18)
        rel.drawString(180, 790, 'Faturamento Mensal')
        rel.setFont('Courier', 12)
        rel.drawString(30, 750, f'Data Emisão: {data} ')
        rel.drawString(250, 750, f'Empresa: Lavador Macaubense IV ')
        rel.drawString(30, 730, f'Mês Do Serviço: {referencia} ')
        rel.drawString(250, 730, f'Usuário Solicitante: {loginy}')
        rel.setFont('Courier', 11)
        rel.drawString(50, 705, 'Codigos')
        rel.drawString(250, 705, 'Clientes')
        rel.drawString(480, 705, 'Valor R$')
        rel.setFont('Courier', 9)
        rel.drawString(50, 690, f'{cod66}'), rel.drawString(130, 690, f'{empresa66}'), rel.drawString(480, 690,
                                                                                                      f'R$ {valor66}')
        rel.drawString(50, 680, f'{cod67}'), rel.drawString(130, 680, f'{empresa67}'), rel.drawString(480, 680,
                                                                                                      f'R$ {valor67}')
        rel.drawString(50, 670, f'{cod68}'), rel.drawString(130, 670, f'{empresa68}'), rel.drawString(480, 670,
                                                                                                      f'R$ {valor68}')
        rel.drawString(50, 660, f'{cod69}'), rel.drawString(130, 660, f'{empresa69}'), rel.drawString(480, 660,
                                                                                                      f'R$ {valor69}')
        rel.rect(20, 700, 545, 20, fill=False, stroke=True)  # Linha titulo da tabela
        rel.rect(20, 725, 545, 40, fill=False, stroke=True)  # cabecario 2
        rel.rect(20, 785, 545, 20, fill=False, stroke=True)  # cabecario 1
        rel.rect(20, 330, 545, 390, fill=False, stroke=True)  # tabela
        rel.rect(120, 330, 350, 390, fill=False, stroke=True)  # Linha cabeçario
        rel.rect(20, 250, 545, 0, fill=True, stroke=False)  # linha total
        rel.rect(20, 230, 545, 0, fill=True, stroke=False)  # linha total
        rel.setFont('Courier', 12)
        rel.drawString(180, 235, 'Faturamento Total: R$ {:.2f}'.format(fat_total))
        rel.setFont('Times-Roman', 6)
        rel.drawString(220, 220, f'Protocolo De Impressão {data1}')
        rel.setFont('Times-Roman', 6)
        rel.rect(20, 14, 545, 0, fill=True, stroke=False)  # ultima linha
        rel.drawString(20, 6,
                       'Faturamento Automático - Fast System™ | Versão 3.0IE. Desenvolvido Pela Empresa Plus Mais Soluções | © Plus Mais Soluções 2023.')
        rel.drawString(540, 6, "Pagina 02")
        rel.showPage()
        rel.save()
        webbrowser.open('Faturamento.pdf')
except:
    sg.popup_error('Erro Ao Gerar Relatório', title='Fast System™')
    sg.popup_error('Enter com Contato com o Suporte', title='Fast System™')

data1 = datetime.datetime.now()
data = (str(data1.day)+ " / "+ str(data1.month)+" / "+ str(data1.year))
descricao1 = "Lavagens De Veículos"
aliquota = '2,00'
loginx = 'Elivando'
senhax = "123456"

def janela_login():
    sg.theme('Default')
    layout =[
         [sg.Text(f'Empresa: Lavador Macaubense IV -  {data}')],
         [sg.Text('Login: '),sg.InputText(key='login')],
         [sg.Text('Senha:'),sg.InputText(key='senha',password_char='*')],
         [sg.Button('Entrar'),sg.Button('Sair'),sg.Text('-----------------------------------------------'),sg.Text('Versão 3.0EI',font='Verdana 8')],
    ]
    return sg.Window('Fast System™',layout=layout,finalize=True)

def janela_info():
    sg.theme('Default')
    layout =[
         [sg.Text(f'Empresa: Lavador Macaubense IV -  {data}')],
         [sg.Text('Faturamento: '),sg.InputText(key='referencia')],
         [sg.Text(' Vencimento: '), sg.InputText(key='vencimento')],
         [sg.Button('Avançar'),sg.Button('Sair'),sg.Text('-----------------------------------------------'),sg.Text('Versão 3.0EI',font='Verdana 8')],
    ]
    return sg.Window('Fast System™',layout=layout,finalize=True)
def janela_menu():
    sg.theme('Default')
    layout =[
         [sg.Text(f'Empresa: Lavador Macaubense IV -  {data}'),sg.Text('--------------------------------------'),sg.Text('Versão 3.0EI',font='Verdana 8')],
         [sg.Button("Emissão Das Nota Fiscal"),sg.Button("Emissão Dos Boletos"),sg.Button("Envio Dos Faturamentos"),sg.Button("Relatório")],
         [sg.Button('Ajuda'),sg.Text('----------------------------------------------------------------------------------'),sg.Button('Sair')],
    ]
    return sg.Window('Menu',layout=layout,finalize=True)


#try:
#Janela
janela,janela1,janela2 = janela_login(),None,None
#Ler os eventos
while True:
    windon, event, values = sg.read_all_windows()
    if windon == janela and event == sg.WIN_CLOSED or event == 'Sair':
        windon.close()
        janela.close()
        break
    if windon == janela and event == 'Entrar':
        if values['login'] == loginx and values['senha'] == senhax:
            loginy = values['login']
            sg.popup_timed(f'Conect Database: Bem Vindo: {loginy} ', title='Conexão')
            try:
                sg.popup_auto_close('Conect Database: Aguarde...Conectando Com o Banco de Dados.', title='Conexão')
                df_envio = pd.read_excel('db_fast.xlsx')
                sg.popup_timed('Database Response: Banco de Dados Conectado.', title='Conexão')
                janela1 = janela_info()
                janela.close()
            except:
                sg.popup_error('Database Response: Sem Conexão com o Banco de Dados', title='Fast System™')
                sg.popup_error('Enter com Contato com o Suporte', title='Fast System™')
                windon.close()
        else:
            sg.popup_timed('Usuário ou Senha Inválida',title = 'Login')

    if windon == janela1 and event == 'Avançar':
        if values['referencia'] != "" and values ['vencimento'] != "":
            sg.popup_timed('Conect Database: Dados Coletados Com Sucesso', title='Conexão')
            referencia = values['referencia']
            vencimento = values['vencimento']
            descricao = descricao1 + " " + referencia + f" Vencimento {vencimento}"
            janela2 = janela_menu()
            janela1.close()
        else:
            sg.popup_timed('Preenchar o Campo', title='Conexão')

    if windon == janela2 and event == 'Emissão Das Nota Fiscal':
        sg.popup_auto_close('Conect Database: Aguarde...Configurando a Emissão da Nota.', title='Conexão')
        servico = Service(ChromeDriverManager().install())
        navegador = webdriver.Chrome(service=servico)
        navegador.get("http://www.webnota.net.br")
        pyautogui.click(x=1208, y=31)
        navegador.find_element(By.XPATH,
                               '/html/body/table/tbody/tr/td/form/table[1]/tbody/tr[3]/td[2]/font/select').click()
        navegador.find_element(By.XPATH,
                               '/html/body/table/tbody/tr/td/form/table[1]/tbody/tr[3]/td[2]/font/select/option[4]').click()
        navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/form/div[3]/input').click()
        navegador.find_element(By.XPATH,
                               '/html/body/table/tbody/tr/td/form/table[1]/tbody/tr[3]/td[2]/font/input').send_keys(
                    '01644578522')
        navegador.find_element(By.XPATH,
                               '/html/body/table/tbody/tr/td/form/table[1]/tbody/tr[5]/td[2]/font/input').send_keys(
            'A8H4K')
        pyautogui.click(x=1208, y=31)
        navegador.find_element(By.XPATH,
                               '/html/body/table/tbody/tr/td/form/table[1]/tbody/tr[7]/td[2]/label/input[2]').click()
        navegador.find_element(By.XPATH, '/html/body/table/tbody/tr/td/form/div[3]/input[2]').click()
        leitura = pyautogui.password(text='O Site Foi Carregado Com Sucesso? Digite: 1- Sim / 2- Não',
                                     title='Fast System™', mask='')
        while leitura != "1":
            sg.popup_auto_close("Por Favor Aguarde... Recarreguando a Página.", title='Fast System™')
            pyautogui.press('f5')
            time.sleep(2)
            leitura = pyautogui.password(text='Log com Sucesso? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')

        pergunta0 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa0}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta0 == '1':
            emitir(cnpj0, valor0, cod0, empresa0)
        pergunta1 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa1}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta1 == '1':
            emitir(cnpj1, valor1, cod1, empresa1)
        pergunta2 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa2}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta2 == '1':
            emitir(cnpj2, valor2, cod2, empresa2)
        pergunta3 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa3}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta3 == '1':
            emitir(cnpj3, valor3, cod3, empresa3)
        pergunta4 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa4}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta4 == '1':
            emitir(cnpj4, valor4, cod4, empresa4)
        pergunta5 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa5}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta5 == '1':
            emitir(cnpj5, valor5, cod5, empresa5)
        if sg.popup_ok_cancel('Dejesa Continuar a Emissão') != 'OK':
            sg.popup_auto_close('Conect Database: Fechando o Navagedor...', title='Conexão')
            navegador.close()
        pergunta6 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa6}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta6 == '1':
            emitir(cnpj6, valor6, cod6, empresa6)
        pergunta7 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa7}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta7 == '1':
            emitir(cnpj7, valor7, cod7, empresa7)
        pergunta8 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa8}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta8 == '1':
            emitir(cnpj8, valor8, cod8, empresa8)
        pergunta9 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa9}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta9 == '1':
            emitir(cnpj9, valor9, cod9, empresa9)
        pergunta10 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa10}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta10 == '1':
            emitir(cnpj10, valor10, cod10 , empresa10)
        if sg.popup_ok_cancel('Dejesa Continuar a Emissão') != 'OK':
            sg.popup_auto_close('Conect Database: Fechando o Navagedor...', title='Conexão')
            navegador.close()
        pergunta11 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa11}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta11 == '1':
            emitir(cnpj11, valor11, cod11, empresa11)
        pergunta12 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa12}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta12 == '1':
            emitir(cnpj12, valor12, cod12, empresa12)
        pergunta13 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa13}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta13 == '1':
            emitir(cnpj13, valor13, cod13, empresa13)
        pergunta14 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa14}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta14 == '1':
            emitir(cnpj14, valor14, cod14, empresa14)
        pergunta15 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa15}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta15 == '1':
            emitir(cnpj15, valor15, cod15, empresa15)
        if sg.popup_ok_cancel('Dejesa Continuar a Emissão') != 'OK':
            sg.popup_auto_close('Conect Database: Fechando o Navagedor...', title='Conexão')
            navegador.close()
        pergunta16 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa16}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta16 == '1':
            emitir(cnpj16, valor16, cod16, empresa16)
        pergunta17 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa17}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta17 == '1':
            emitir(cnpj17, valor17, cod17, empresa17)
        pergunta18 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa18}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta18 == '1':
            emitir(cnpj18, valor18, cod18, empresa18)
        pergunta19 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa19}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta19 == '1':
            emitir(cnpj19, valor19, cod19, empresa19)
        pergunta20 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa20}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta20 == '1':
            emitir(cnpj20, valor20, cod20 , empresa20)
        if sg.popup_ok_cancel('Dejesa Continuar a Emissão') != 'OK':
            sg.popup_auto_close('Conect Database: Fechando o Navagedor...', title='Conexão')
            navegador.close()
        pergunta21 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa21}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta21 == '1':
            emitir(cnpj21, valor21, cod21, empresa21)
        pergunta22 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa22}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta22 == '1':
            emitir(cnpj22, valor22, cod22, empresa22)
        pergunta23 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa23}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta23 == '1':
            emitir(cnpj23, valor23, cod23, empresa23)
        pergunta24 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa24}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta24 == '1':
            emitir(cnpj24, valor24, cod24, empresa24)
        pergunta25 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa25}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta25 == '1':
            emitir(cnpj25, valor25, cod25, empresa25)
        if sg.popup_ok_cancel('Dejesa Continuar a Emissão') != 'OK':
            sg.popup_auto_close('Conect Database: Fechando o Navagedor...', title='Conexão')
            navegador.close()
        pergunta26 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa26}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta26 == '1':
            emitir(cnpj26, valor26, cod26, empresa26)
        pergunta27 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa27}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta27 == '1':
            emitir(cnpj27, valor27, cod27, empresa27)
        pergunta28 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa28}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta28 == '1':
            emitir(cnpj28, valor28, cod28, empresa28)
        pergunta29 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa29}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta29 == '1':
            emitir(cnpj29, valor29, cod29, empresa29)
        pergunta30 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa30}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta30 == '1':
            emitir(cnpj30, valor30, cod30 , empresa30)
        if sg.popup_ok_cancel('Dejesa Continuar a Emissão') != 'OK':
            sg.popup_auto_close('Conect Database: Fechando o Navagedor...', title='Conexão')
            navegador.close()
        pergunta31 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa31}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta31 == '1':
            emitir(cnpj31, valor31, cod31, empresa31)
        pergunta32 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa32}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta32 == '1':
            emitir(cnpj32, valor32, cod32, empresa32)
        pergunta33 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa33}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta33 == '1':
            emitir(cnpj33, valor33, cod33, empresa33)
        pergunta34 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa34}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta34 == '1':
            emitir(cnpj34, valor34, cod34, empresa34)
        pergunta35 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa35}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta35 == '1':
            emitir(cnpj35, valor35, cod35, empresa35)
        if sg.popup_ok_cancel('Dejesa Continuar a Emissão') != 'OK':
            sg.popup_auto_close('Conect Database: Fechando o Navagedor...', title='Conexão')
            navegador.close()
        pergunta36 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa36}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta36 == '1':
            emitir(cnpj36, valor36, cod36, empresa36)
        pergunta37 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa37}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta37 == '1':
            emitir(cnpj37, valor37, cod37, empresa37)
        pergunta38 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa38}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta38 == '1':
            emitir(cnpj8, valor8, cod8, empresa8)
        pergunta39 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa39}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta39 == '1':
            emitir(cnpj9, valor9, cod9, empresa9)
        pergunta40 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa40}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta40 == '1':
            emitir(cnpj40, valor40, cod40 , empresa40)
        if sg.popup_ok_cancel('Dejesa Continuar a Emissão') != 'OK':
            sg.popup_auto_close('Conect Database: Fechando o Navagedor...', title='Conexão')
            navegador.close()
        pergunta41 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa41}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta41 == '1':
            emitir(cnpj41, valor41, cod41, empresa41)
        pergunta42 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa42}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta42 == '1':
            emitir(cnpj42, valor42, cod42, empresa42)
        pergunta43 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa43}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta43 == '1':
            emitir(cnpj43, valor43, cod43, empresa43)
        pergunta44 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa44}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta44 == '1':
            emitir(cnpj44, valor44, cod44, empresa44)
        pergunta45 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa45}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta45 == '1':
            emitir(cnpj45, valor45, cod45, empresa45)
        if sg.popup_ok_cancel('Dejesa Continuar a Emissão') != 'OK':
            sg.popup_auto_close('Conect Database: Fechando o Navagedor...', title='Conexão')
            navegador.close()
        pergunta46 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa46}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta46 == '1':
            emitir(cnpj46, valor46, cod46, empresa46)
        pergunta47 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa47}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta47 == '1':
            emitir(cnpj47, valor47, cod47, empresa47)
        pergunta48 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa48}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta48 == '1':
            emitir(cnpj48, valor48, cod48, empresa48)
        pergunta49 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa49}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta49 == '1':
            emitir(cnpj49, valor49, cod49, empresa49)
        pergunta50 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa50}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta50 == '1':
            emitir(cnpj50, valor50, cod50 , empresa50)
        if sg.popup_ok_cancel('Dejesa Continuar a Emissão') != 'OK':
            sg.popup_auto_close('Conect Database: Fechando o Navagedor...', title='Conexão')
            navegador.close()
        pergunta51 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa51}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta51 == '1':
            emitir(cnpj51, valor51, cod51, empresa51)
        pergunta52 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa52}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta52 == '1':
            emitir(cnpj52, valor52, cod52, empresa52)
        pergunta53 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa53}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta53 == '1':
            emitir(cnpj53, valor53, cod53, empresa53)
        pergunta54 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa54}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta54 == '1':
            emitir(cnpj54, valor54, cod54, empresa54)
        pergunta55 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa55}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta55 == '1':
            emitir(cnpj55, valor55, cod55, empresa55)
        if sg.popup_ok_cancel('Dejesa Continuar a Emissão') != 'OK':
            sg.popup_auto_close('Conect Database: Fechando o Navagedor...', title='Conexão')
            navegador.close()
        pergunta56 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa56}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta56 == '1':
            emitir(cnpj56, valor56, cod56, empresa56)
        pergunta57 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa57}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta57 == '1':
            emitir(cnpj57, valor57, cod57, empresa57)
        pergunta58 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa58}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta58 == '1':
            emitir(cnpj58, valor58, cod58, empresa58)
        pergunta59 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa59}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta59 == '1':
            emitir(cnpj59, valor59, cod59, empresa59)
        pergunta60 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa60}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta60 == '1':
            emitir(cnpj60, valor60, cod60 , empresa60)
        if sg.popup_ok_cancel('Dejesa Continuar a Emissão') != 'OK':
            sg.popup_auto_close('Conect Database: Fechando o Navagedor...', title='Conexão')
            navegador.close()
        pergunta61 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa61}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta61 == '1':
            emitir(cnpj61, valor61, cod61, empresa61)
        pergunta62 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa62}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta62 == '1':
            emitir(cnpj62, valor62, cod62, empresa62)
        pergunta63 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa63}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta63 == '1':
            emitir(cnpj63, valor63, cod63, empresa63)
        pergunta64 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa64}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta64 == '1':
            emitir(cnpj64, valor64, cod64, empresa64)
        pergunta65 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa65}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta65 == '1':
            emitir(cnpj65, valor65, cod65, empresa65)
        if sg.popup_ok_cancel('Dejesa Continuar a Emissão') != 'OK':
            sg.popup_auto_close('Conect Database: Fechando o Navagedor...', title='Conexão')
            navegador.close()
        pergunta66 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa66}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta66 == '1':
            emitir(cnpj66, valor66, cod66, empresa66)
        pergunta67 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa67}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta67 == '1':
            emitir(cnpj67, valor67, cod67, empresa67)
        pergunta68 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa68}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta68 == '1':
            emitir(cnpj68, valor68, cod68, empresa68)
        pergunta69 = pyautogui.password(text=f'Deseja Emitir a Nota: {empresa69}? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        if pergunta69 == '1':
            emitir(cnpj69, valor69, cod69, empresa69)
        sg.popup_auto_close('Conect Database: Processo De Emissão Concluído... Fechando o Navegador.', title='Conexão')
        navegador.close()
        janela2 = janela_menu()

    if windon == janela2 and event == 'Emissão Dos Boletos':
        sg.popup_auto_close('Conect Database: Aguarde...Configurando a Emissão do Boleto.', title='Conexão')
        servico1 = Service(ChromeDriverManager().install())
        navegador1 = webdriver.Chrome(service=servico1)
        navegador1.get("https://app.cora.com.br/")
        pyautogui.click(x=1208, y=31)
        sg.popup_auto_close("Por Favor, Leia o Qrcode", title='Fast System™')
        leitura = pyautogui.password(text='Leitura com Sucesso? Digite: 1- Sim / 2- Não', title='Fast System™', mask='')
        while leitura != "1":
            sg.popup_auto_close("Por Favor Aguarde... Recarreguando a Página.", title='Fast System™')
            pyautogui.press('f5')
            time.sleep(2)
            leitura = pyautogui.password(text='Leitura com Sucesso? Digite: 1- Sim / 2- Não', title='Fast System™',
                                         mask='')
        time.sleep(1)
        navegador1.find_element(By.XPATH, '//*[@id="root"]/div/main/div/div/div[2]/div/div[4]/button').click()  # Emitir

        try:
            emitir1(cnpj0, valor0, cod0)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa0}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            emitir1(cnpj1, valor1, cod1)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa1}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            emitir1(cnpj2, valor2, cod2)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa2}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            emitir1(cnpj3, valor3, cod3)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa3}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            emitir1(cnpj4, valor4, cod4)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa4}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            if sg.popup_ok_cancel('Dejesa Continuar a Emissão') != 'OK':
                sg.popup_auto_close('Conect Database: Fechando o Navagedor...', title='Conexão')
                navegador1.close()
            emitir1(cnpj5, valor5, cod5)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa5}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            emitir1(cnpj6, valor6, cod6)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa6}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            emitir1(cnpj7, valor7, cod7)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto {empresa7}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            emitir1(cnpj8, valor8, cod8)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa8}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            emitir1(cnpj9, valor9, cod9)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa9}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            if sg.popup_ok_cancel('Dejesa Continuar a Emissão') != 'OK':
                sg.popup_auto_close('Conect Database: Fechando o Navagedor...', title='Conexão')
                navegador1.close()
            emitir1(cnpj10, valor10, cod10)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa10}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            emitir1(cnpj11, valor11, cod11)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa11}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            emitir1(cnpj12, valor12, cod12)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa12}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            emitir1(cnpj13, valor13, cod13)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa13}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            emitir1(cnpj14, valor14, cod14)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa14}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            if sg.popup_ok_cancel('Dejesa Continuar a Emissão') != 'OK':
                sg.popup_auto_close('Conect Database: Fechando o Navagedor...', title='Conexão')
                navegador1.close()
            emitir1(cnpj15, valor15, cod15)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa15}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            emitir1(cnpj16, valor16, cod16)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa16}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            emitir1(cnpj17, valor17, cod17)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto {empresa17}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            emitir1(cnpj18, valor18, cod18)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa18}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            emitir1(cnpj19, valor19, cod19)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa19}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            if sg.popup_ok_cancel('Dejesa Continuar a Emissão') != 'OK':
                sg.popup_auto_close('Conect Database: Fechando o Navagedor...', title='Conexão')
                navegador1.close()
            emitir1(cnpj20, valor20, cod20)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa20}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            emitir1(cnpj21, valor21, cod21)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa21}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            emitir1(cnpj22, valor22, cod22)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa22}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            emitir1(cnpj23, valor23, cod23)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa23}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            emitir1(cnpj24, valor24, cod24)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa24}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            if sg.popup_ok_cancel('Dejesa Continuar a Emissão') != 'OK':
                sg.popup_auto_close('Conect Database: Fechando o Navagedor...', title='Conexão')
                navegador1.close()
            emitir1(cnpj25, valor25, cod25)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa25}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            emitir1(cnpj26, valor26, cod26)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa26}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            emitir1(cnpj27, valor27, cod27)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto {empresa27}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            emitir1(cnpj28, valor28, cod28)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa28}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            emitir1(cnpj29, valor29, cod29)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa29}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            if sg.popup_ok_cancel('Dejesa Continuar a Emissão') != 'OK':
                sg.popup_auto_close('Conect Database: Fechando o Navagedor...', title='Conexão')
                navegador1.close()
            emitir1(cnpj30, valor30, cod30)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa30}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            emitir1(cnpj31, valor31, cod31)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa31}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            emitir1(cnpj32, valor32, cod32)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa32}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            emitir1(cnpj33, valor33, cod33)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa33}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            emitir1(cnpj34, valor34, cod34)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa34}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            if sg.popup_ok_cancel('Dejesa Continuar a Emissão') != 'OK':
                sg.popup_auto_close('Conect Database: Fechando o Navagedor...', title='Conexão')
                navegador1.close()
            emitir1(cnpj35, valor35, cod35)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa35}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            emitir1(cnpj36, valor36, cod36)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa36}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            emitir1(cnpj37, valor37, cod37)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto {empresa37}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            emitir1(cnpj38, valor38, cod38)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa38}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            emitir1(cnpj39, valor39, cod39)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa39}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            if sg.popup_ok_cancel('Dejesa Continuar a Emissão') != 'OK':
                sg.popup_auto_close('Conect Database: Fechando o Navagedor...', title='Conexão')
                navegador1.close()
            emitir1(cnpj40, valor40, cod40)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa40}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            emitir1(cnpj41, valor41, cod41)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa41}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            emitir1(cnpj42, valor42, cod42)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa42}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            emitir1(cnpj43, valor43, cod43)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa43}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            emitir1(cnpj44, valor44, cod44)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa44}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            if sg.popup_ok_cancel('Dejesa Continuar a Emissão') != 'OK':
                sg.popup_auto_close('Conect Database: Fechando o Navagedor...', title='Conexão')
                navegador1.close()
            emitir1(cnpj45, valor45, cod45)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa45}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            emitir1(cnpj46, valor46, cod46)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa46}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            emitir1(cnpj47, valor47, cod47)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto {empresa47}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            emitir1(cnpj48, valor48, cod48)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa48}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            emitir1(cnpj49, valor49, cod49)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa49}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            if sg.popup_ok_cancel('Dejesa Continuar a Emissão') != 'OK':
                sg.popup_auto_close('Conect Database: Fechando o Navagedor...', title='Conexão')
                navegador1.close()
            emitir1(cnpj50, valor50, cod50)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa50}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            emitir1(cnpj51, valor51, cod51)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa51}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            emitir1(cnpj52, valor52, cod52)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa52}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            emitir1(cnpj53, valor53, cod53)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa53}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            emitir1(cnpj54, valor54, cod54)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa54}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            if sg.popup_ok_cancel('Dejesa Continuar a Emissão') != 'OK':
                sg.popup_auto_close('Conect Database: Fechando o Navagedor...', title='Conexão')
                navegador1.close()
            emitir1(cnpj55, valor55, cod55)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa55}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            emitir1(cnpj56, valor56, cod56)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa56}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            emitir1(cnpj57, valor57, cod57)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto {empresa57}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            emitir1(cnpj58, valor58, cod58)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa58}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            emitir1(cnpj59, valor59, cod59)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa59}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            if sg.popup_ok_cancel('Dejesa Continuar a Emissão') != 'OK':
                sg.popup_auto_close('Conect Database: Fechando o Navagedor...', title='Conexão')
                navegador1.close()
            emitir1(cnpj60, valor60, cod60)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa60}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            emitir1(cnpj61, valor61, cod61)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa61}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            emitir1(cnpj62, valor62, cod62)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa62}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            emitir1(cnpj63, valor63, cod63)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa63}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            emitir1(cnpj64, valor64, cod64)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa64}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            if sg.popup_ok_cancel('Dejesa Continuar a Emissão') != 'OK':
                sg.popup_auto_close('Conect Database: Fechando o Navagedor...', title='Conexão')
                navegador1.close()
            emitir1(cnpj65, valor65, cod65)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa65}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            emitir1(cnpj66, valor66, cod66)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa66}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            emitir1(cnpj67, valor67, cod67)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto {empresa67}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            emitir1(cnpj68, valor68, cod68)
        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa68}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        try:
            emitir1(cnpj69, valor69, cod69)

        except:
            sg.popup_error(f'Conect Database: Erro Ao Emitir Boleto: {empresa69}', title='Conexão')
            pyautogui.click('menu.png')
            pyautogui.click('menu2.png')
            navegador1.find_element(By.XPATH, '//*[@id="root"]/div/nav/div/ul[1]/li[1]').click()  # Menu
        janela2 = janela_menu()

    if windon == janela2 and event == 'Envio Dos Faturamentos':
          sg.popup_auto_close('Conect Database: Aguarde...Enviandos Os Faturamentos.', title='Conexão')
          try:
           enviar(email0, cod0, empresa0)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa0}', title='Conexão')
          try:
              enviar(email1, cod1, empresa1)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa1}', title='Conexão')
          try:
              enviar(email2, cod2, empresa2)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa2}', title='Conexão')
          try:
              enviar(email3, cod3, empresa3)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa3}', title='Conexão')
          try:
              enviar(email4, cod4, empresa4)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa4}', title='Conexão')
          try:
              enviar(email5, cod5, empresa5)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa5}', title='Conexão')
          try:
              enviar(email6, cod6, empresa6)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa6}', title='Conexão')
          try:
              enviar(email7, cod7, empresa7)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa7}', title='Conexão')
          try:
              enviar(email8, cod8, empresa8)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa8}', title='Conexão')
          try:
              enviar(email9, cod9, empresa9)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa9}', title='Conexão')
          try:
              enviar(email10, cod10, empresa10)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa10}', title='Conexão')
          try:
              enviar(email11, cod11, empresa11)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa11}', title='Conexão')
          try:
              enviar(email12, cod12, empresa12)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa12}', title='Conexão')
          try:
              enviar(email13, cod13, empresa13)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa13}', title='Conexão')
          try:
              enviar(email14, cod14, empresa14)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa14}', title='Conexão')
          try:
              enviar(email15, cod15, empresa15)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa15}', title='Conexão')
          try:
              enviar(email16, cod16, empresa16)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa16}', title='Conexão')
          try:
              enviar(email17, cod17, empresa17)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa17}', title='Conexão')
          try:
              enviar(email18, cod18, empresa18)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa18}', title='Conexão')
          try:
              enviar(email19, cod19, empresa19)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa19}', title='Conexão')
          try:
              enviar(email20, cod20, empresa20)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa20}', title='Conexão')
          try:
              enviar(email21, cod21, empresa21)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa21}', title='Conexão')
          try:
              enviar(email22, cod22, empresa22)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa22}', title='Conexão')
          try:
              enviar(email23, cod23, empresa23)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa23}', title='Conexão')
          try:
              enviar(email24, cod24, empresa24)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa24}', title='Conexão')
          try:
              enviar(email25, cod25, empresa25)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa25}', title='Conexão')
          try:
              enviar(email26, cod26, empresa26)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa26}', title='Conexão')
          try:
              enviar(email27, cod27, empresa27)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa27}', title='Conexão')
          try:
              enviar(email28, cod28, empresa28)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa28}', title='Conexão')
          try:
              enviar(email29, cod29, empresa29)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa29}', title='Conexão')
          try:
              enviar(email30, cod30, empresa30)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa30}', title='Conexão')
          try:
              enviar(email41, cod41, empresa41)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa41}', title='Conexão')
          try:
              enviar(email42, cod42, empresa42)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa42}', title='Conexão')
          try:
              enviar(email43, cod43, empresa43)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa43}', title='Conexão')
          try:
              enviar(email44, cod44, empresa44)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa44}', title='Conexão')
          try:
              enviar(email45, cod45, empresa45)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa45}', title='Conexão')
          try:
              enviar(email46, cod46, empresa46)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa46}', title='Conexão')
          try:
              enviar(email47, cod47, empresa47)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa47}', title='Conexão')
          try:
              enviar(email48, cod48, empresa48)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa48}', title='Conexão')
          try:
              enviar(email49, cod49, empresa49)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa49}', title='Conexão')
          try:
              enviar(email40, cod40, empresa40)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa40}', title='Conexão')
          try:
              enviar(email51, cod51, empresa51)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa51}', title='Conexão')
          try:
              enviar(email52, cod52, empresa52)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa52}', title='Conexão')
          try:
              enviar(email53, cod53, empresa53)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa53}', title='Conexão')
          try:
              enviar(email54, cod54, empresa54)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa54}', title='Conexão')
          try:
              enviar(email55, cod55, empresa55)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa55}', title='Conexão')
          try:
              enviar(email56, cod56, empresa56)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa56}', title='Conexão')
          try:
              enviar(email57, cod57, empresa57)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa57}', title='Conexão')
          try:
              enviar(email58, cod58, empresa58)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa58}', title='Conexão')
          try:
              enviar(email59, cod59, empresa59)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa59}', title='Conexão')
          try:
              enviar(email60, cod60, empresa60)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa60}', title='Conexão')
          try:
              enviar(email61, cod61, empresa61)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa61}', title='Conexão')
          try:
              enviar(email62, cod62, empresa62)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa62}', title='Conexão')
          try:
              enviar(email63, cod63, empresa63)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa63}', title='Conexão')
          try:
              enviar(email64, cod64, empresa64)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa64}', title='Conexão')
          try:
              enviar(email65, cod65, empresa65)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa65}', title='Conexão')
          try:
              enviar(email66, cod66, empresa66)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa66}', title='Conexão')
          try:
              enviar(email67, cod67, empresa67)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa67}', title='Conexão')
          try:
              enviar(email68, cod68, empresa68)
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa68}', title='Conexão')
          try:
              enviar(email69, cod69, empresa69)
              janela2 = janela_menu()
          except:
              sg.popup_auto_close(f'Conect Database: Erro Ao Enviar: {empresa69}', title='Conexão')
              janela2 = janela_menu()

    if windon == janela2 and event == 'Relatório':
        sg.popup_auto_close('Conect Database: Aguarde...Gerando Relatório.', title='Conexão')
        rel_fat()

    if windon == janela2 and event == 'Ajuda':
        webbrowser.open('https://wa.me/message/YPI2UVEWCJORJ1')


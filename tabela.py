from fpdf import FPDF
from tkinter import *
from tkinter import ttk
from tkinter.tix import *
from csv import reader
from tkscrolledframe import ScrolledFrame
import pandas as pd
from datetime import date
from tkinter import messagebox
import os
import datetime
import glob
from reportlab.pdfgen import canvas
from PyPDF2 import PdfFileWriter, PdfFileReader
from time import sleep
import win32com.client as win32


outlook = win32.Dispatch('outlook.application')

# criar um email
email = outlook.CreateItem(0)

faturamento = 1500
qtde_produtos = 10
ticket_medio = faturamento / qtde_produtos

# configurar as informações do seu e-mail
email.To = "loliveira@gasbrasiliano.com.br"
email.Subject = "E-mail automático Análise Tributária TESTE"
email.HTMLBody = f"""


<p>A Análise Tributária DV-004-2021 está disponível para assinatura.</p>


<p>Abs,</p>
<p>Código Python</p>
"""

# anexo = "C://Users/joaop/Downloads/arquivo.xlsx"
# email.Attachments.Add(anexo)

email.Send()
print("Email Enviado")
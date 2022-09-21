#ctypes          para minimizar terminal
#cv2             para Camera + Leitura QR
#datetime        para operação com data
#mysql.connector para conectar ao banco
#openpyxl        para Editar XLSX
#os              para reiniciar o programa
#ping3           para Ping
#png             para gerar png do QRCode
#psutil          para trabalhar com PID
#pyqrcode        para gerar QRCode com ID
#shutil          para mover arquivos XLSX
#time            para utilizar pausa
#tkinter         para fazer janelas
#requests        para fazer janelas
#win32gui        para minimizar terminal
#win32con        para minimizar terminal
#webrowser       para Download XLSX HTTP

import ctypes
import cv2
import datetime
import mysql.connector
import openpyxl
import os
import ping3
import png
import psutil
import pyqrcode
import shutil
import time
import tkinter
import requests 
#import win32gui   -> DESATIVADO
#import win32con   -> DESATIVADO
#import webbrowser -> DESATIVADO

from datetime             import date
from datetime             import datetime
from datetime             import timedelta
from openpyxl             import load_workbook
from os                   import listdir
from ping3                import ping
from pyqrcode             import QRCode
from tkinter              import *
from tkinter.scrolledtext import ScrolledText
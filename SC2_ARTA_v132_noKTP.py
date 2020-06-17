import tkinter as tk
import sys
import time
import calendar
import random
import datetime as dt
from tkinter import *
from PIL import Image, ImageTk
import locale
import os
import cv2
import numpy as np
import pytesseract
import imutils
import warnings
warnings.filterwarnings("ignore")
import serial
import xlwt
from datetime import datetime
from playsound import playsound
from SC4_ARTA_v131 import main as main_SC4
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

sys.path.append("..")

ardLoadcell = serial.Serial('com5',9600,timeout=5)
ard1 = serial.Serial('com11',9600,timeout=5)
TB = 0
BB = 0
tempTubuh1 = 0

timeKTP = 5
timeSuhu = 5
timeAll = timeKTP + timeSuhu
timeAll1 = timeKTP + timeSuhu
teksWarning = "Siapkan KTP, lalu taruh di depan kamera, \ndan kami akan hitung mundur: "

baris = 1
style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on',
    num_format_str='#,##0.00')
style1 = xlwt.easyxf(num_format_str='D-MMM-YY')
wb = xlwt.Workbook()

def cekKTP():
    global timeKTP, teksWarning, teksNama, teksNIK, teksTempat, teksTanggal, umur, teksSex
    cam = cv2.VideoCapture(0)
    s, im = cam.read()
    image_cropped = im
    image = imutils.resize(image_cropped,1024,665)

    #try:
    imageNIK = image[50:145,195:770]
    imageNama = image[120:180,245:760]
    imageTTL = image[170:220,245:700]
    imageSex = image[210:255,245:430]

    teksNIK = "Sesuai KTP" #pytesseract.image_to_string(imageNIK, lang = 'eng')
    teksNama = "Sesuai KTP" #pytesseract.image_to_string(imageNama, lang = 'eng')
    teksTTL = pytesseract.image_to_string(imageTTL, lang = 'eng')
    teksSex = "Sesuai KTP" #pytesseract.image_to_string(imageSex, lang = 'eng')
    teksTTL = teksTTL.split(',')
    teksTempat = "Sesuai KTP" #teksTTL[0]
    teksTanggal = "Sesuai KTP" #teksTTL[1]        
    teksTahun = "Sesuai KTP" #teksTanggal.split(' ')
    teksTahun1 = "Sesuai KTP" #teksTahun[2]
    umur = "Sesuai KTP" #2020 - int(teksTahun1)
            
    print("IDENTITAS USER: ")
    print("   ")
    print("NIK:  " + teksNIK)
    print("Nama:  " + teksNama)
    print("Tempat lahir:  " + teksTempat)
    print("Tanggal lahir:  " + teksTanggal)
    print("Umur:  " + str(umur) + ' tahun')
    print("Jenis kelamin:  " + teksSex)

    output_path = 'KTP/'+teksNama+'.jpg'
    cv2.imwrite(output_path,image)
    teksWarning = "    Dekatkan dahi anda di depan sensor, \n    dan kami akan hitung mundur: "
'''
    except:
        print("Maaf! KTP tidak terdeteksi")
        print("Silahkan coba lagi...")
        print()
        teksWarning = "Maaf, KTP anda tidak terdeteksi, silahkan coba lagi: "
        timeKTP = 5
        return'''

def cekTubuh():
    global TB,tempTubuh1
    ard1.write(b't') #kirim perintah ke arduino
    msg = ard1.readline() #baca data serial dari arduino
    if (len(msg) > 5): #jika panjang karakter data > 5
        strmsg = str(msg) #ini data dari serial yg sudah string
        data = strmsg.split(',') #karena formatnya "data1,data2,data3", maka dipisah berdasar tanda koma
        tempTubuh1 = int(data[1].replace("\\r\\n'","")) #biar cuma tinggal angkanya
        TB = data[0].replace("b'","") #biar cuma tinggal angkanya
        print(TB,tempTubuh1)
        print()
        
        ard1.flush()

def cekBeratTubuh():
    global BB
    ardLoadcell.write(b't') #kirim perintah ke arduino
    msg1 = ardLoadcell.readline() #baca data serial dari arduino
    if (len(msg1) > 2): #jika panjang karakter data > 5
        strmsg1 = str(msg1) #ini data dari serial yg sudah string
        BB = strmsg1.replace("\\r\\n'","") #biar cuma tinggal angkanya
        BB = BB.replace("b'","") #biar cuma tinggal angkanya
        print(BB)
        print()
        ardLoadcell.flush()
        
        
def simpanExcel():
    ws = wb.add_sheet('Data Pasien ARTA 1')
    ws.write(0, 0, 1)
    ws.write(0, 1, 'Nama')
    ws.write(0, 2, 'NIK')
    ws.write(0, 3, 'Tempat Lahir')
    ws.write(0, 4, 'Tanggal Lahir')
    ws.write(0, 5, 'Umur')
    ws.write(0, 6, 'Jenis Kelamin')
    ws.write(0, 7, 'Tinggi Badan (cm)')
    ws.write(0, 8, 'Berat Badan (kg)')
    ws.write(0, 9, 'Suhu Badan (C)')
    ws.write(0, 10, 'G1b')
    ws.write(0, 11, 'G2')
    ws.write(0, 12, 'G3')
    ws.write(0, 13, 'F1')
    ws.write(0, 14, 'F2a')
    ws.write(0, 15, 'F2b')
    ws.write(0, 16, 'F2c')
    ws.write(0, 17, 'F2d')
    ws.write(baris, 0, datetime.now(), style1)
    ws.write(baris, 1, teksNama)
    ws.write(baris, 2, teksNIK)
    ws.write(baris, 3, teksTempat)
    ws.write(baris, 4, teksTanggal)
    ws.write(baris, 5, umur)
    ws.write(baris, 6, teksSex)
    ws.write(baris, 7, TB)
    ws.write(baris, 8, BB)
    ws.write(baris, 9, tempTubuh1)
    wb.save('Data Pasien ARTA 1.xls')
    print("sudah tersimpan...")

class Counter(tk.Label):
    def __init__(self, parent=None, seconds=True, colon=False):
        tk.Label.__init__(self, parent)

        self.display_seconds = seconds
        if self.display_seconds:
            self.time     = teksWarning + str(timeKTP)
        else:
            self.time     = teksWarning + str(timeKTP)
        self.display_time = self.time
        self.configure(text=self.display_time)

        if colon:
            self.blink_colon()

        self.after(1000, self.tick)
                                              
    def tick(self):
        global timeKTP, timeAll

        if (timeAll == timeAll1):
            playsound('Rekaman ARTA\Slide 2 (Sebelum Memulai).mp3')
            
        timeKTP -= 1
        timeAll -= 1
        if self.display_seconds:
            if timeKTP > 0:
                new_time = teksWarning + str(timeKTP)
            else:
                new_time = teksWarning + str(timeAll)
        else:
            if timeKTP > 0:
                new_time = teksWarning + str(timeKTP)
            else:
                new_time = teksWarning + str(timeAll)
        if new_time != self.time:
            self.time = new_time
            self.display_time = self.time
            self.config(text=self.display_time)

        

        if (timeAll == timeAll - timeKTP):
            cekBeratTubuh()
            cekKTP()

        if (timeAll == (timeAll - timeKTP)-2):
            playsound('Rekaman ARTA\BEEP.mp3')
            playsound('Rekaman ARTA\BEEP.mp3')
            playsound('Rekaman ARTA\BEEP.mp3')

        if (timeAll == 0):
            cekTubuh()
            simpanExcel()
            ardLoadcell.close()
            root.quit()
            main_SC4()
            # import SC4_ARTA_v131
            

        print(timeAll)
            
            
        self.after(1000, self.tick)


def main():
    global root
    root=tk.Tk()
    root.geometry("1366x768") #Width x Height
    root.attributes('-fullscreen', True)

    load = Image.open('Komponen UI\scene_2\LayoutUI_scene 2.jpg')
    render = ImageTk.PhotoImage(load)
    img = Label(image=render)
    img.image = render
    img.place(x=0, y=0)


    clock2 = Counter(root)
    clock2.configure(fg="black", bg="white",font=("helvetica",25))
    clock2.place(x=390, y=345)

    root.mainloop()
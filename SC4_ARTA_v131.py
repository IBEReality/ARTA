import tkinter as tk
import threading
import serial                                      # add Serial library for serial communication
import pyautogui
import datetime as dt
from tkinter import *
from PIL import Image, ImageTk
from xlrd import open_workbook
from xlutils.copy import copy
import numpy
import time
from playsound import playsound
import ctypes
ctypes.windll.user32.ShowWindow( ctypes.windll.kernel32.GetConsoleWindow(), 6 )



Arduino_Serial = serial.Serial('com3',9600) 
suara = ["Rekaman ARTA\Slide 4 - Pertanyaan Pertama.mp3",\
         "Rekaman ARTA\Slide 5 - Pertanyaan Kedua.mp3",\
         "Rekaman ARTA\Slide 6 - Pertanyaan Ketiga.mp3",\
         "Rekaman ARTA\Slide 7 - Pertanyaan Keempat.mp3",\
         "Rekaman ARTA\Slide 8 - Pertanyaan Kelima.mp3",\
         "Rekaman ARTA\Slide 9 - Pertanyaan Keenam.mp3",\
         "Rekaman ARTA\Slide 10 - Pertanyaan Ketujuh.mp3",\
         "Rekaman ARTA\Slide 11 - Pertanyaan Kedelapan.mp3",\
         "Rekaman ARTA\Slide 12 - Anamnesi Selesai.mp3"]
         

tanya = ["Apakah ada riwayat demam dalam 14 hari terakhir?",\
         "Apakah ada riwayat batuk/pilek/nyeri tenggorokan dalam 14 hari terakhir?",\
         "Apakah ada riwayat sesak napas dalam 14 hari terakhir?",\
         "Apakah ada riwayat perjalanan ke luar negeri atau wilayah red zone dalam 14 hari terakhir?",\
         "Apakah ada riwayat kontak dengan pasien COVID-19 dalam 14 hari terakhir?",\
         "Apakah ada riwayat bekerja/berkunjung ke fasilitas kesehatan yang berhubungan dengan pasien COVID-19 dalam 14 hari terakhir?",\
         "Apakah anda Punya riwayat kontak dengan hewan penular?",\
         "Apakah anda Punya riwayat kontak dengan orang yang punya riwayat perjalanan ke luar negeri atau wilayah red zone dalam 14 hari terakhir?",\
         "\n\nAnamnesi selesai, mohon tunggu hasilnya..."]

iTanya = 1
dataJawaban = numpy.zeros(len(tanya)+2)

#wb = xlwt.Workbook()

baris = 1
book=open_workbook('Data Pasien ARTA 1.xls')
sheet=book.sheet_by_index(0)
Nama = sheet.cell(baris,1).value
NIK = sheet.cell(baris,2).value
Tempat = sheet.cell(baris,3).value
Tanggal = sheet.cell(baris,4).value
Umur = str(sheet.cell(baris,5).value)
Sex  = sheet.cell(baris,6).value
TB = str(sheet.cell(baris,7).value)
BB = str(sheet.cell(baris,8).value)
Suhu = sheet.cell(baris,9).value

inputText = "  Identitas anda :\n"+"  Nama             : "+Nama+"\n" \
            +"  NIK                 : "+NIK+"\n"\
            +"  Tempat Lahir  : "+Tempat+"\n"\
            +"  Tanggal Lahir : "+Tanggal+"\n"\
            +"  Umur              : "+Umur+"\n"\
            +"  Jenis Kelamin : "+str(Sex)+"\n"\
            +"  Tinggi / Berat Badan : "+TB+"(cm) / "+BB+"(kg)\n"\
            +"  Swipe kanan untuk mulai tes... "

class App1(threading.Thread):
    def __init__(self):
        threading.Thread.__init__(self)
        self.start()

    def callback(self):
        self.root.quit()

    def quit(self):
       self.root.destroy 

    def run(self):
        global E1, T, root, load
        self.root = tk.Tk()
        v1 = tk.StringVar()
        self.root.protocol("SC4_ARTA_v13", self.callback)

        self.root.geometry("1366x768") #Width x Height
        self.root.attributes('-fullscreen', True)
        self.root.lift()
        load = Image.open('Komponen UI\scene_3\img_3_1.png')
        render = ImageTk.PhotoImage(load)
        img = Label(image=render)
        img.image = render
        img.place(x=0, y=0)

        load = Image.open('Komponen UI2\scene_4\img_4_5.png')
        render = ImageTk.PhotoImage(load)
        img = Label(image=render)
        img.image = render
        img.place(x=1050, y=460)

        load = Image.open('Komponen UI2\scene_4\img_4_6.png')
        render = ImageTk.PhotoImage(load)
        img = Label(image=render)
        img.image = render
        img.place(x=80, y=460)

        load = Image.open('Komponen UI2\scene_3\img_3_5.png')
        render = ImageTk.PhotoImage(load)
        img = Label(image=render)
        img.image = render
        img.place(x=410, y=490)


        T = tk.Text(self.root, height=9, width=50)
        T.configure(fg="black", bg="white",font=("helvetica",25))
        T.insert(tk.END, inputText)
        T.place(x=280, y=100)
        
        #E1 = tk.Entry(self.root, bd =5)
        #E1.place(x=20, y=50)

        self.root.mainloop()

def simpanExcel():
    wb = copy(book)
    ws = wb.get_sheet(0)
    ws.write(baris, 10, Tr[1])
    ws.write(baris, 11, Tr[2])
    ws.write(baris, 12, Tr[3])
    ws.write(baris, 13, Tr[4])
    ws.write(baris, 14, Tr[5])
    ws.write(baris, 15, Tr[6])
    ws.write(baris, 16, Tr[7])
    ws.write(baris, 17, Tr[8])
    ws.write(baris, 18, hasilTr[0])
    wb.save('Data Pasien ARTA 1.xls')
    print("sudah tersimpan...")

def cekAnamnesi():
    global Tr, hasilTr
    Tr = [Suhu,0,0,0,0,0,0,0,0]
    Tr[1] = dataJawaban[1]
    Tr[2] = dataJawaban[2]
    Tr[3] = dataJawaban[3]
    Tr[4] = dataJawaban[4]
    Tr[5] = dataJawaban[5]
    Tr[6] = dataJawaban[6]
    Tr[7] = dataJawaban[7]
    Tr[8] = dataJawaban[8]
    print(Tr)
    hasilTr = [0,0,0,0,0,0] #[H,HG1,HG2,HG3,HF1,HF2]
    H = []

    sub2()
    sub3()
    sub4()

    print(hasilTr)
    time.sleep(2)

    if (hasilTr[0] == 3):
        T.delete(1.0, tk.END)
        T.insert(tk.END, "\n\nSilahkan menuju loket A untuk pemeriksaan selanjutnya...")
        print("Silahkan menuju loket A untuk pemeriksaan selanjutnya")
    elif (hasilTr[0] == 2):
        T.delete(1.0, tk.END)
        T.insert(tk.END, "\n\nSilahkan menuju loket B untuk pemeriksaan selanjutnya...")
        print("Silahkan menuju loket B untuk pemeriksaan selanjutnya")
    else:
        T.delete(1.0, tk.END)
        T.insert(tk.END, "\n\nAnda bebas penyakit...")
        print("Anda bebas")

def sub2():
    global Tr, hasilTr
    G1a = int(Tr[0])
    G1b = Tr[1]
    G2 = Tr[2]
    G3 = Tr[3]

    if (G1a >= 38):
        HG1 = 1
        hasilTr[1] = HG1
    else:
        if (G1b == 1):
            HG1 = 1
            hasilTr[1] = HG1
        else:
            HG1 = 0
            hasilTr[1] = HG1

    if (G2 == 1):
        HG2 = 1
        hasilTr[2] = HG2
    else:
        HG2 = 0
        hasilTr[2] = HG2

    if (G3 == 1):
        HG3 = 1
        hasilTr[3] = HG3
    else:
        HG3 = 0
        hasilTr[3] = HG3

def sub3():
    global Tr, hasilTr
    F1 = Tr[4]
    F2a = Tr[5]
    F2b = Tr[6]
    F2c = Tr[7]
    F2d = Tr[8]

    if (F1 == 1):
        HF1 = 1
        hasilTr[4] = HF1
    else:
        HF1 = 0
        hasilTr[4] = HF1

    if (F2a == 1):
        HF2 = 1
        hasilTr[5] = HF2
    else:
        if (F2b == 1):
            HF2 = 1
            hasilTr[5] = HF2
        else:
            if (F2c == 1):
                HF2 = 1
                hasilTr[5] = HF2
            else:
                if (F2d == 1):
                    HF2 = 1
                    hasilTr[5] = HF2
                else:
                    HF2 = 0
                    hasilTr[5] = HF2

def sub4():
    global Tr, hasilTr
    HG1 = hasilTr[1]
    HG2 = hasilTr[2]
    HG3 = hasilTr[3]
    HF1 = hasilTr[4]
    HF2 = hasilTr[5]

    if ((HG1 == 1) and (HG2 == 1) and (HG3 == 1) and (HF1 == 1)):
        H = 3
        hasilTr[0] = H
    else:
        if ((HG1 == 1) and (HF2 == 1)):
            H = 3
            hasilTr[0] = H
        else:
            if ((HG2 == 1) and (HF2 == 1)):
                H = 3
                hasilTr[0] = H
            else:
                if ((HG1 == 1) and (HF1 == 1)):
                    H = 2
                    hasilTr[0] = H
                else:
                    if ((HG2 == 1) and (HF1 == 1)):
                        H = 2
                        hasilTr[0] = H
                    else:
                        H = 1
                        hasilTr[0] = H


def main():
    global iTanya
    app = App1()
    print('Now we can continue running code while mainloop runs!')
    while 1:
        if (iTanya == 1):
            playsound('Rekaman ARTA\Slide 3 (Petunjuk 1-3).mp3')
            #playsound('Rekaman ARTA\Petunjuk Melambaikan tangan.mp3')

        incoming_data = str (Arduino_Serial.readline())
        dataIn = incoming_data.replace("b'","")
        dataIn1 = dataIn.replace("\r\n","")
        dataIn2 = dataIn1[0]
        print(dataIn2)
        if (iTanya == (len(tanya))):
            T.delete(1.0, tk.END)
            T.insert(tk.END, tanya[len(tanya)-1])
            playsound(suara[iTanya-1])
            break
        else:
            if (dataIn2 == "1"):
                T.delete(1.0, tk.END)
                T.insert(tk.END, "\n\nPertanyaan " + str(iTanya) + " : \n" + tanya[iTanya-1])
                playsound(suara[iTanya-1])
                dataJawaban[iTanya] = 0
                iTanya += 1
            elif (dataIn2 == "2"):
                T.delete(1.0, tk.END)
                T.insert(tk.END, "\n\nPertanyaan " + str(iTanya) + " : \n" + tanya[iTanya-1])
                playsound(suara[iTanya-1])
                dataJawaban[iTanya] = 1
                iTanya += 1
            else:
                T.delete(1.0, tk.END)
                T.insert(tk.END, "\n\nPertanyaan " + str(iTanya) + " : \n" + tanya[iTanya-1]\
                         +"\nSwipe dengan benar!")
                playsound(suara[iTanya-1])

        print(iTanya, dataJawaban)

    cekAnamnesi()
    simpanExcel()
    print(iTanya, dataJawaban)

def quit():
    global root
    root.quit()
#time.sleep(5)
if __name__ == '__main__':
    main()


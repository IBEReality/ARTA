import tkinter as tk
import sys
import time
import calendar
import random
import datetime as dt
from tkinter import *
from PIL import Image, ImageTk
import locale
import serial #library serial
import socket
from playsound import playsound
import ctypes
ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 6 )
from SC2_ARTA_v132_noKTP import main as main_SC2

ard = serial.Serial('com10',9600,timeout=5) #deklarasi port komunikasi serial

dis = 0
temp = 0
hum = 0

class Clock(tk.Label):
    def __init__(self, parent=None, seconds=True, colon=False):
        tk.Label.__init__(self, parent)

        self.display_seconds = seconds
        if self.display_seconds:
            self.time     = time.strftime('%I:%M:%S')
        else:
            self.time     = time.strftime('%I:%M:%S').lstrip('0')
        self.display_time = self.time
        self.configure(text=self.display_time)

        if colon:
            self.blink_colon()

        self.after(200, self.tick)
                                              
    def tick(self):
        if self.display_seconds:
            new_time = time.strftime('%I:%M:%S')
        else:
            new_time = time.strftime('%I:%M:%S').lstrip('0')
        if new_time != self.time:
            self.time = new_time
            self.display_time = self.time
            self.config(text=self.display_time)
        
        self.after(200, self.tick)


    def blink_colon(self):
        if ':' in self.display_time:
            self.display_time = self.display_time.replace(':',' ')
        else:
            self.display_time = self.display_time.replace(' ',':',1)
        self.config(text=self.display_time)
        self.after(1000, self.blink_colon)

class Sensor(tk.Label):
    def __init__(self, parent=None, seconds=True, colon=False):
        tk.Label.__init__(self, parent)

        self.display_seconds = seconds
            
        if self.display_seconds:
            self.time     = "Sedang menghubungkan sensor..."
        else:
            self.time     = "Sedang menghubungkan sensor..."
            
        self.display_time = self.time
        self.configure(text=self.display_time)

        if colon:
            self.blink_colon()

        self.after(200, self.tick)
                                             
    def tick(self):
        ard.write(b't') #kirim perintah ke arduino
        msg = ard.readline() #baca data serial dari arduino
        #data = msg.split(',')
        #print (len(msg))
        if (len(msg) > 5): #jika panjang karakter data > 5
            global dis,temp,hum
            strmsg = str(msg) #ini data dari serial yg sudah string
            data = strmsg.split(',') #karena formatnya "data1,data2,data3", maka dipisah berdasar tanda koma
            dis = int(data[2].replace("\\r\\n'","")) #biar cuma tinggal angkanya
            temp = data[1] #sudah tinggal angka
            hum = data[0].replace("b'","") #biar cuma tinggal angkanya
            print(data)

            if (dis < 150):
                playsound('Rekaman ARTA\Slide 1 - Introduction ARTA.mp3')
                #playsound('Rekaman ARTA\Slide 2 (Sebelum Memulai).mp3')
                ard.close()
                root.quit()
                main_SC2()





        degree_sign= u'\N{DEGREE SIGN}'
        
        if self.display_seconds:
            new_time = "Suhu: "+ str(temp) + " " + degree_sign + "C | Kelembaban: " + str(hum) +" %"
        else:
            new_time = "Suhu: "+ str(temp) +" " + degree_sign + "C | Kelembaban: " + str(hum) +" %"
        if new_time != self.time:
            self.time = new_time
            self.display_time = self.time
            self.config(text=self.display_time)

        self.after(200, self.tick)

class FullScreenApp(object):
    def __init__(self, master, **kwargs):
        self.master=master
        pad=3
        self._geom='200x200+0+0'
        master.geometry("{0}x{1}+0+0".format(
            master.winfo_screenwidth()-pad, master.winfo_screenheight()-pad))
        master.bind('<Escape>',self.toggle_geom)            
    def toggle_geom(self,event):
        geom=self.master.winfo_geometry()
        print(geom,self._geom)
        self.master.geometry(self._geom)
        self._geom=geom



def main ():
    global root
    # Root is the name of the Tkinter Window. This is important to remember.
    root=tk.Tk()
    root.geometry("1366x768") #Width x Height
    root.attributes('-fullscreen', True)

    load = Image.open('Komponen UI\scene_1\LayoutUI_scene 2.jpg')
    render = ImageTk.PhotoImage(load)
    img = Label(image=render)
    img.image = render
    img.place(x=0, y=0)


    clock1 = Clock(root)
    clock1.configure(fg="black", bg="white",font=("helvetica",20))
    clock1.place(x=1240, y=20)

    clock2 = Sensor(root)
    clock2.configure(fg="black", bg="white",font=("helvetica",16))
    clock2.place(x=1010, y=60)

    date = dt.datetime.now()
    locale.setlocale(locale.LC_TIME, "IND")
    format_date = f"{date:%A, %d %B %Y | }"

    w = Label(root, text=format_date, fg="black", bg="white", font=("arial", 20))
    w.place(x=970, y=20)

    root.mainloop()

if __name__ == '__main__':
    main()
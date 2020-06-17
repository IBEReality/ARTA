from fpdf import FPDF
import time
import datetime as dt
import locale
from xlrd import open_workbook
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
import smtplib
import base64
from WS_ARTA_v131 import main as main_WS
def main():
    tgl = time.strftime("%Y-%m-%d")
    wkt = time.strftime("%H:%M")

    hari = int(tgl[8:10])
    bulan = int(tgl[5:7])
    tahun = int(tgl[0:4])

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
    G1b = sheet.cell(baris,10).value
    G2 = sheet.cell(baris,11).value
    G3 = sheet.cell(baris,12).value
    F1 = sheet.cell(baris,13).value
    F2a = sheet.cell(baris,14).value
    F2b = sheet.cell(baris,15).value
    F2c = sheet.cell(baris,16).value
    F2d = sheet.cell(baris,17).value
    hasilTr = int(sheet.cell(baris,18).value)

    if (F2a == 1):
        F2ay = "o"
        F2at = " "
    else:
        F2ay = " "
        F2at = "o"
    if (F2b == 1):
        F2by = "o"
        F2bt = " "
    else:
        F2by = " "
        F2bt = "o"
    if (F2c == 1):
        F2cy = "o"
        F2ct = " "
    else:
        F2cy = " "
        F2ct = "o"
    if (F2d == 1):
        F2dy = "o"
        F2dt = " "
    else:
        F2dy = " "
        F2dt = "o"

    date = dt.datetime.now()
    locale.setlocale(locale.LC_TIME, "IND")
    format_date = f"{date:%d %B %Y}"

    pdf = FPDF('P', 'mm', 'A4')
    pdf.set_margins(150, 200, 10)
    pdf.add_page()
    pdf.set_font('Times', 'B', 12)

    pdf.set_font('Times', '', 12)
    pdf.set_xy(15, 15)
    pdf.multi_cell(0, 5, "Nama              : " + Nama,0,'L',0)
    pdf.set_xy(15, 20)
    pdf.multi_cell(0, 5, "NIK                 : " + NIK,0,'L',0)
    pdf.set_xy(15, 25)
    pdf.multi_cell(0, 5, "Tanggal Lahir : " + Tanggal,0,'L',0)

    pdf.set_font('Times', 'B', 12)
    pdf.set_xy(0, 35)
    pdf.multi_cell(0, 5, "FORMULIR DETEKSI DINI CORONA VIRUS DESEASE (COVID-19)",0,'C',0)

    pdf.set_font('Times', '', 12)
    pdf.set_xy(15, 45)
    degree_sign= u'\N{DEGREE SIGN}'
    pdf.multi_cell(0, 5, "Suhu Badan      : " + str(Suhu)+ " " + degree_sign + "C",0,'L',0)
    pdf.set_xy(15, 50)
    pdf.multi_cell(0, 5, "Berat Badan     : " + str(BB) + " kg",0,'L',0)
    pdf.set_xy(15, 55)
    pdf.multi_cell(0, 5, "Tinggi Badan    : " + str(TB) + " cm",0,'L',0)

    pdf.set_font('Times', 'B', 12)
    pdf.set_xy(15, 65)
    pdf.multi_cell(0, 5, "GEJALA",0,'L',0)

    pdf.set_xy(15, 70)
    pdf.multi_cell(20, 5, "No.",1,'C',0)
    pdf.set_xy(35, 70)
    pdf.multi_cell(120, 5, "Pertanyaan",1,'C',0)
    pdf.set_xy(155, 70)
    pdf.multi_cell(20, 5, "Ya",1,'C',0)
    pdf.set_xy(175, 70)
    pdf.multi_cell(20, 5, "Tidak",1,'C',0)

    pdf.set_font('Times', '', 12)
    baris = 70
    baris += 5
    pdf.set_xy(15, baris)
    pdf.multi_cell(20, 5, "1",1,'C',0)
    pdf.set_xy(35, baris)
    pdf.multi_cell(120, 5, "Demam / riwayat demam",1,'L',0)
    pdf.set_xy(155, baris)
    if (G1b == 1):
        pdf.multi_cell(20, 5, "o",1,'C',0)
        pdf.set_xy(175, baris)
        pdf.multi_cell(20, 5, " ",1,'C',0)
    else:
        pdf.multi_cell(20, 5, " ",1,'C',0)
        pdf.set_xy(175, baris)
        pdf.multi_cell(20, 5, "o",1,'C',0)

    baris += 5
    pdf.set_xy(15, baris)
    pdf.multi_cell(20, 5, "2",1,'C',0)
    pdf.set_xy(35, baris)
    pdf.multi_cell(120, 5, "Batuk / pilek / nyeri tenggorokan",1,'L',0)
    pdf.set_xy(155, baris)
    if (G1b == 1):
        pdf.multi_cell(20, 5, "o",1,'C',0)
        pdf.set_xy(175, baris)
        pdf.multi_cell(20, 5, " ",1,'C',0)
    else:
        pdf.multi_cell(20, 5, " ",1,'C',0)
        pdf.set_xy(175, baris)
        pdf.multi_cell(20, 5, "o",1,'C',0)

    baris += 5
    pdf.set_xy(15, baris)
    pdf.multi_cell(20, 5, "3",1,'C',0)
    pdf.set_xy(35, baris)
    pdf.multi_cell(120, 5, "Sesak napas",1,'L',0)
    pdf.set_xy(155, baris)
    if (G1b == 1):
        pdf.multi_cell(20, 5, "o",1,'C',0)
        pdf.set_xy(175, baris)
        pdf.multi_cell(20, 5, " ",1,'C',0)
    else:
        pdf.multi_cell(20, 5, " ",1,'C',0)
        pdf.set_xy(175, baris)
        pdf.multi_cell(20, 5, "o",1,'C',0)

    pdf.set_font('Times', 'B', 12)
    baris += 10
    pdf.set_xy(15, baris)
    pdf.multi_cell(0, 5, "FAKTOR RESIKO",0,'L',0)
    baris += 5
    pdf.set_xy(15, baris)
    pdf.multi_cell(20, 5, "NO",1,'C',0)
    pdf.set_xy(35, baris)
    pdf.multi_cell(120, 5, "Pertanyaan",1,'C',0)
    pdf.set_xy(155, baris)
    pdf.multi_cell(20, 5, "Ya",1,'C',0)
    pdf.set_xy(175, baris)
    pdf.multi_cell(20, 5, "Tidak",1,'C',0)

    pdf.set_font('Times', '', 12)
    baris += 5
    pdf.set_xy(15, baris)
    pdf.multi_cell(20, 5, "1\n ",1,'C',0)
    pdf.set_xy(35, baris)
    pdf.multi_cell(120, 5, "Riwayat perjalanan ke luar negeri atau wilayah red zone dalam waktu 14 hari sebelum timbul gejala",1,'L',0)
    pdf.set_xy(155, baris)
    if (G1b == 1):
        pdf.multi_cell(20, 5, "o\n ",1,'C',0)
        pdf.set_xy(175, baris)
        pdf.multi_cell(20, 5, " \n ",1,'C',0)
    else:
        pdf.multi_cell(20, 5, " \n ",1,'C',0)
        pdf.set_xy(175, baris)
        pdf.multi_cell(20, 5, "o\n ",1,'C',0)


    baris += 10
    pdf.set_xy(15, baris)
    pdf.multi_cell(20, 5, "2\n\n\n\n\n\n\n\n",1,'C',0)
    pdf.set_xy(35, baris)
    pdf.multi_cell(120, 5, "Memiliki riwayat paparan salah satu atau lebih:\n"\
                   +"  a. Riwayat kontak erat dengan dengan kasus konfirmasi\n"\
                   +"  b. Bekerja atau mengunjungi fasilitas yang berhubungan dengan dengan kasus konfirmasi\n"\
                   +"  c. Memiliki riwayat kontak dengan hewan penular ATAU\n"\
                   +"  d. Memiliki riwayat kontak dengan orang yang punya riwayat perjalanan ke luar negeri atau wilayah red zone dalam 14 hari terakhir"\
                   ,1,'L',0)
    pdf.set_xy(155, baris)
    pdf.multi_cell(20, 5, "Ya\n"+F2ay+"\n"+F2by+"\n\n"+F2cy+"\n"+F2dy+"\n\n\n",1,'C',0)
    pdf.set_xy(175, baris)
    pdf.multi_cell(20, 5, "Tidak\n"+F2at+"\n"+F2bt+"\n\n"+F2ct+"\n"+F2dt+"\n\n\n",1,'C',0)

    baris += 45
    if (hasilTr == 3):
        hasilTr1 = "Pasien Dalam Pengawasan (PDP)"
        TL = "Rujuk IGD"
    elif (hasilTr == 2):
        hasilTr1 = "Orang Dalam Pengawasan (ODP)"
        TL = "Rujuk IGD"
    else:
        hasilTr1 = "Tanpa Gejala"
        TL = "Silahkan pulang dan jaga kesehatan"

    pdf.set_font('Times', 'B', 12)
    pdf.set_xy(15, baris)
    pdf.multi_cell(0, 5, "KESIMPULAN : " + hasilTr1,0,'L',0)
    baris += 10
    pdf.set_xy(15, baris)
    pdf.multi_cell(0, 5, "TINDAK LANJUT : " + TL,0,'L',0)

    baris += 30
    pdf.set_xy(15, baris)
    pdf.multi_cell(0, 5, "Nb: Lampiran KTP",0,'L',0)

    pdf.set_xy(130, 180)
    pdf.cell(0, 0, 'Surabaya, '+format_date)

    pdf.set_xy(115, 190)
    pdf.multi_cell(90, 5, "Tanda Tangan Petugas Skrining",0,'C',0)
    pdf.set_xy(115, 225)
    pdf.multi_cell(90, 5, ". . . . . . . . . . . . . . . . . . . . . . . . . .",0,'C',0)

    image = 'KTP/'+Nama+'.jpg'
    baris += 10
    pdf.image(image, x=15, y=baris, w=90, h=50)

    pdf.set_font('Times', '', 6)
    pdf.set_xy(170, 270)
    pdf.cell(0, 0, "digitally signed at "+wkt+", "+tgl)

    namaFile = "SURAT\HASIL TRIASE_"+Nama+'_'+NIK+".pdf"
    pdf.output(namaFile, 'F')

    print('save pdf oke')



    #############################################################################
    msg = MIMEMultipart()

    message = "  Identitas Pasien :\n"+"  Nama             : "+Nama+"\n" \
                +"  NIK                 : "+NIK+"\n"\
                +"  Tempat Lahir  : "+Tempat+"\n"\
                +"  Tanggal Lahir : "+Tanggal+"\n"\
                +"  Umur              : "+Umur+"\n"\
                +"  Jenis Kelamin : "+str(Sex)+"\n"\
                +"  Tinggi / Berat Badan : "+TB+"(cm) / "+BB+"(kg)\n"\
                +"\n"\
                +"  Hasil Tes Triase: "+hasilTr1

    # setup the parameters of the message
    password = "Megantoro1"
    msg['From'] = "megantoro.prisma@gmail.com"
    msg['To'] = "megantoro.prisma@yahoo.co.id"
    msg['Subject'] = "DATA PASIEN ARTA 2"

    # add in the message body
    msg.attach(MIMEText(message, 'plain'))
    fp = open(namaFile, 'rb')
    attach = MIMEApplication(fp.read(), 'pdf')
    fp.close()
    attach.add_header('Content-Disposition', 'attachment', filename = "HASIL TRIASE_"+Nama+'_'+NIK+".pdf")
    msg.attach(attach)

    #create server
    server = smtplib.SMTP('smtp.gmail.com: 587')

    server.starttls()

    # Login Credentials for sending the mail
    server.login(msg['From'], password)


    # send the message via the server.
    server.sendmail(msg['From'], msg['To'], msg.as_string())

    server.quit()

    print("successfully sent email to %s:" % (msg['To']))


if __name__ == '__main__':
    main()
    main_WS()
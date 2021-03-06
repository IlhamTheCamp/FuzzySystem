import matplotlib.pyplot as plot
import pandas as pd
import xlsxwriter

# # Penghasilan
h = [0, 5.5, 8, 13, 15.5, 22]

def penghasilanRendah(x):
    if(x < h[1]):
        return 1
    elif(x >= h[2]):
        return 0
    elif((x >= h[1]) and (x < h[2])):
        return (h[2] - x)/(h[2] - h[1])

def penghasilanSedang(x):
    if((x <= h[1]) or (x >= h[4])):
        return 0
    elif((x > h[2]) and (x < h[3])):
        return 1
    elif((x > h[1]) and (x <= h[2])):
        return (x - h[1])/(h[2] - h[1])
    elif((x >= h[3]) and (x < h[4])):
        return (h[4] - x)/(h[4] - h[3])

def penghasilanTinggi(x):
    if(x <= h[3]):
        return 0
    elif((x > h[3]) and (x <= h[4])):
        return (x - h[3])/(h[4] - h[3])
    elif(x > h[4]):
        return 1

# #Plot Penghasilan
y1 = [1, 1, 0, 0, 0, 0]
y2 = [0, 0, 1, 1, 0, 0]
y3 = [0, 0, 0, 0, 1, 1]

plot.plot(h, y1, label="Penghasilan Rendah")
plot.plot(h, y2, label="Penghasilan Sedang")
plot.plot(h, y3, label="Penghasilan Tinggi")
plot.legend()
plot.show()

# # Pengeluaran
k = [0, 3.5, 5, 7, 8.5, 11]
#
def pengeluaranRendah(x):
    if(x < k[1]):
        return 1
    elif(x >= k[2]):
        return 0
    elif((x >= k[1]) and (x < k[2])):
        return (k[2] - x)/(k[2] - k[1])

def pengeluaranSedang(x):
    if((x <= k[1]) or (x >= k[4])):
        return 0
    elif((x > k[2]) and (x < k[3])):
        return 1
    elif((x > k[1]) and (x <= k[2])):
        return (x - k[1])/(k[2] - k[1])
    elif((x >= k[3]) and (x < k[4])):
        return (k[4] - x)/(k[4] - k[3])

def pengeluaranTinggi(x):
    if(x <= k[3]):
        return 0
    elif((x > k[3]) and (x <= k[4])):
        return (x - k[3])/(k[4] - k[3])
    elif(x > k[4]):
        return 1

# #Plot Pengeluaran
y1 = [1, 1, 0, 0, 0, 0]
y2 = [0, 0, 1, 1, 0, 0]
y3 = [0, 0, 0, 0, 1, 1]

plot.plot(k, y1, label="Pengeluaran Rendah")
plot.plot(k, y2, label="Pengeluaran Sedang")
plot.plot(k, y3, label="Pengeluaran Tinggi")
plot.legend()
plot.show()

# #Inferensi Rules
def basedRules(hasil, keluar, id):
    inferensi = []

    inferensi.append(["Tidak", (hasil[id][2] and keluar[id][1])])
    inferensi.append(["Tidak", (hasil[id][3]) and (keluar[id][1])])
    inferensi.append(["Tidak", (hasil[id][3]) and (keluar[id][2])])

    inferensi.append(["Mungkin", (hasil[id][1] and keluar[id][1])])
    inferensi.append(["Mungkin", (hasil[id][2] and keluar[id][2])])
    inferensi.append(["Mungkin", (hasil[id][3]) and (keluar[id][3])])

    inferensi.append(["Iya", (hasil[id][1] and keluar[id][2])])
    inferensi.append(["Iya", (hasil[id][1] and keluar[id][3])])
    inferensi.append(["Iya", (hasil[id][2]) and (keluar[id][3])])

    return inferensi

# #Defuzzy dengan Model Sugeno
def defuzzification(inferensi):
    a = ((inferensi[0] * 30) + (inferensi[1] * 60) + (inferensi[2] * 100))
    b = inferensi[0] + inferensi[1] + inferensi[2]
    return a / b

# #Plot Kelayakan Model Sugeno
plot.axvline(30, 0, 1, color="red", label='tidak')
plot.axvline(60, 0, 1, color="green", label='mungkin')
plot.axvline(100, 0, 1, color="blue", label='iya')
plot.legend()
plot.show()
#
# #Main Program
# #Baca file Excel
excel = pd.read_excel('Mahasiswa.xls')

nilaiPenghasilan = []
nilaiPengeluaran = []

# #Mengambil nilai fuzzy Penghasilan & Pengeluaran
for i in range(len(excel)):
    nilai = []
    nilai.append(excel["Id"][i])
    nilai.append(penghasilanRendah(excel["Penghasilan"][i]))
    nilai.append(penghasilanSedang(excel["Penghasilan"][i]))
    nilai.append(penghasilanTinggi(excel["Penghasilan"][i]))
    nilaiPenghasilan.append(nilai)
    nilai = []
    nilai.append(excel["Id"][i])
    nilai.append(pengeluaranRendah(excel["Pengeluaran"][i]))
    nilai.append(pengeluaranSedang(excel["Pengeluaran"][i]))
    nilai.append(pengeluaranTinggi(excel["Pengeluaran"][i]))
    nilaiPengeluaran.append(nilai)

# #Pengaplikasian Inferensi Rules
defuzzy = []
for i in range(len(excel)):
    inferensi = []
    temp = basedRules(nilaiPenghasilan, nilaiPengeluaran, i)
    iya = []
    mungkin = []
    tidak = []
    for k in range(len(temp)):
        if (temp[k][0] == "Iya"):
            iya.append(temp[k][1])
        if (temp[k][0] == "Mungkin"):
            mungkin.append(temp[k][1])
        if (temp[k][0] == "Tidak"):
            tidak.append(temp[k][1])
    inferensi.append(tidak[0] or tidak[1] or tidak[2])
    inferensi.append(mungkin[0] or mungkin[1] or mungkin[2])
    inferensi.append(iya[0] or iya[1] or iya[2])
    defuzzy.append([defuzzification(inferensi), i+1])

#Mengurutkan dan mengambil 20 ID yang paling layak
defuzzy.sort(reverse=True)
layak = []
for i in range (20):
    layak.append(defuzzy[i][1])

# # Buat file output
workbook = xlsxwriter.Workbook('Bantuan.xls')
worksheet = workbook.add_worksheet()
row = column = 0
for item in layak :
    worksheet.write(row, column, item)
    row += 1
workbook.close()
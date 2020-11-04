import matplotlib.pyplot as plot
import pandas as pd
import xlsxwriter

#Baca file Excel
excel = pd.read_excel('Mahasiswa.xls')

# Penghasilan
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

#Plot Penghasilan
y1 = [1, 1, 0, 0, 0, 0]
y2 = [0, 0, 1, 1, 0, 0]
y3 = [0, 0, 0, 0, 1, 1]

plot.plot(h, y1, label="Penghasilan Rendah")
plot.plot(h, y2, label="Penghasilan Sedang")
plot.plot(h, y3, label="Penghasilan Tinggi")
plot.legend()
plot.show()

# Pengeluaran
k = [0, 3.5, 5, 7, 8.5, 11]

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

#Plot Pengeluaran
y1 = [1, 1, 0, 0, 0, 0]
y2 = [0, 0, 1, 1, 0, 0]
y3 = [0, 0, 0, 0, 1, 1]

plot.plot(h, y1, label="Pengeluaran Rendah")
plot.plot(h, y2, label="Pengeluaran Sedang")
plot.plot(h, y3, label="Pengeluaran Tinggi")
plot.legend()
plot.show()

#Inferensi Rules
def basedRules(hasil, keluar, id):
    literasi = ["Rendah", "Sedang", "Tinggi"]
    litHasil = literasi
    litKeluar = literasi
    inferensi = []
    if (litHasil[0] == "Rendah") and (litKeluar[0] == "Rendah"):
        inferensi.append(["Mungkin", (hasil[id][1] and keluar[id][1])])
    if (litHasil[0] == "Rendah") and (litKeluar[1] == "Sedang"):
        inferensi.append(["Iya", (hasil[id][1] and keluar[id][2])])
    if (litHasil[0] == "Rendah") and (litKeluar[2] == "Tinggi"):
        inferensi.append(["Iya", (hasil[id][1] and keluar[id][3])])
    if (litHasil[1] == "Sedang") and (litKeluar[0] == "Rendah"):
        inferensi.append(["Tidak", (hasil[id][2] and keluar[id][1])])
    if (litHasil[1] == "Sedang") and (litKeluar[1] == "Sedang"):
        inferensi.append(["Mungkin", (hasil[id][2] and keluar[id][2])])
    if (litHasil[1] == "Sedang") and (litKeluar[2] == "Tinggi"):
        inferensi.append(["Iya", (hasil[id][2]) and (keluar[id][3])])
    if (litHasil[2] == "Tinggi") and (litKeluar[0] == "Rendah"):
        inferensi.append(["Tidak", (hasil[id][3]) and (keluar[id][1])])
    if (litHasil[2] == "Tinggi") and (litKeluar[1] == "Sedang"):
        inferensi.append(["Tidak", (hasil[id][3]) and (keluar[id][2])])
    if (litHasil[2] == "Tinggi") and (litKeluar[2] == "Tinggi"):
        inferensi.append(["Tidak", (hasil[id][3]) and (keluar[id][3])])
    return inferensi

#Main Program
nilaiPenghasilan = []
nilaiPengeluaran = []

#Mengambil nilai fuzzy Penghasilan
for i in range(len(excel)):
    nilai = []
    nilai.append(excel["Id"][i])
    nilai.append(penghasilanRendah(excel["Penghasilan"][i]))
    nilai.append(penghasilanSedang(excel["Penghasilan"][i]))
    nilai.append(penghasilanTinggi(excel["Penghasilan"][i]))
    nilaiPenghasilan.append(nilai)

#Mengambil nilai fuzzy Pengeluaran
for i in range(len(excel)):
    nilai = []
    nilai.append(excel["Id"][i])
    nilai.append(pengeluaranRendah(excel["Pengeluaran"][i]))
    nilai.append(pengeluaranSedang(excel["Pengeluaran"][i]))
    nilai.append(pengeluaranTinggi(excel["Pengeluaran"][i]))
    nilaiPengeluaran.append(nilai)

#Pengaplikasian Inferensi Rules
defuzzy = []
for i in range(len(excel)):
    temp = basedRules(nilaiPenghasilan, nilaiPengeluaran, i)
    inferensi = iya = mungkin = tidak = []
    for k in range(len(temp)):
        if (temp[k][0] == "Iya"):
            iya.append(temp[k][1])
        if (te)

# penghasilan = 7.84
# print("Rendah : ", penghasilanRendah(penghasilan))
# print("Sedang : ", penghasilanSedang(penghasilan))
# print("Tinggi : ", penghasilanTinggi(penghasilan))
#
# pengeluaran = 4.89
# print("Rendah : ", pengeluaranRendah(pengeluaran))
# print("Sedang : ", pengeluaranSedang(pengeluaran))
# print("Tinggi : ", pengeluaranTinggi(pengeluaran))

# Buat file output
# workbook = xlsxwriter.Workbook('Bantuan.xls')
# worksheet = workbook.add_worksheet()
#
# row = 0
# column = 0
#
# content = [11,2,4,1,4,6,5,5,3,53]
#
# for item in content :
#     worksheet.write(row, column, item)
#     row += 1
# workbook.close()
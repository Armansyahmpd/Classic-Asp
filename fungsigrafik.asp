'-------------------------------------------------------------------- 
'---------- Fungsi Penghitungan Persentase dan Gambar Statistik ----- 
'---------- Diprogram oleh Armansyah, S.Kom; Rabu 18 Juni 2003 ------ 
'-------------------------------------------------------------------- 

Dim PanjangGambar1 
Dim PanjangGambar2 
Dim PanjangGambar3 
Dim PanjangGambar4 
Dim TotalPersen 
Dim Total 

    Total = Total_Semua_OEFSELESAI + Total_Semua_OEFBLMSELESAI + Total_Semua_OEFBatal 
    
Function HitungPersen(Angka) 
    TotalPersen = (angka/total) * 100 
    TotalPersen = Cint(TotalPersen) 
    HitungPersen = TotalPersen 
End Function 

Function Gambar1(x) 
    PanjangGambar1 = (x*2) 
    Gambar1 = PanjangGambar1 
End Function 

Function Gambar2(y) 
    PanjangGambar2 = (y*2) 
    Gambar2 = PanjangGambar2 
End Function 

Function Gambar3(z) 
    PanjangGambar3 = (z*2) 
    Gambar3 = PanjangGambar3 
End Function 

Function Gambar4(Arman_Mita) 
    PanjangGambar4 = (Arman_Mita*2) 
    Gambar4 = PanjangGambar4 
End Function

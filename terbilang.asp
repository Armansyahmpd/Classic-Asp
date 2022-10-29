<%
'---------------------------------------------------------------------------
' Menampilkan tanda terima angka uang dalam bentuk abjad
' Seperti tanda terima kwitansi
' Oleh : Armansyah
' Kamis, 14 Juli 2005 23:28 Wib
'---------------------------------------------------------------------------

Public Function terbilang(x)
Dim triliun
Dim milyar
Dim juta
Dim ribu
Dim satu
Dim sen
Dim baca

'Jika x adalah 0, maka dibaca sebagai 0
If x = 0 Then
baca = angka(0, 1)
Else

'Pisah masing-masing bagian untuk triliun, milyar, juta, ribu, rupiah, dan sen

triliun = Int(x * 0.001 ^ 4)
milyar = Int((x - triliun * 1000 ^ 4) * 0.001 ^ 3)
juta = Int((x - triliun * 1000 ^ 4 - milyar * 1000 ^ 3) / 1000 ^ 2)
ribu = Int((x - triliun * 1000 ^ 4 - milyar * 1000 ^ 3 - juta * 1000 ^ 2) / 1000)
satu = Int(x - triliun * 1000 ^ 4 - milyar * 1000 ^ 3 - juta * 1000 ^ 2 - ribu * 1000)
sen = Int((x - Int(x)) * 100)

'Baca bagian triliun dan ditambah akhiran triliun
If triliun > 0 Then
baca = ratus(triliun, 5) + "triliun "
End If

'Baca bagian milyar dan ditambah akhiran milyar
If milyar > 0 Then
baca = ratus(milyar, 4) + "milyar "
End If

'Baca bagian juta dan ditambah akhiran juta
If juta > 0 Then
baca = baca + ratus(juta, 3) + "juta "
End If

'Baca bagian ribu dan ditambah akhiran ribu
If ribu > 0 Then
baca = baca + ratus(ribu, 2) + "ribu "
End If

'Baca bagian rupiah dan ditambah akhiran rupiah
If satu > 0 Then
baca = baca + ratus(satu, 1) + "rupiah "
End If

'Baca bagian sen dan ditambah akhiran sen
If sen > 0 Then
baca = baca + ratus(sen, 0) + "sen"
End If
End If
terbilang = UCase(Left(baca, 1)) & LCase(Mid(baca, 2))
End Function

Function ratus(x, posisi)
Dim a100, a10 , a1
Dim baca

a100 = Int(x * 0.01)
a10 = Int((x - a100 * 100) * 0.1)
a1 = Int(x - a100 * 100 - a10 * 10)

'Baca Bagian Ratus
If a100 = 1 Then
baca = "Seratus "
Else
If a100 > 0 Then
baca = angka(a100, 2) + "ratus "
End If
End If

'Baca Bagian Puluh dan Satuan
If a10 = 1 Then
baca = baca + angka(a10 * 10 + a1, 2)
Else
If a10 > 0 Then
baca = baca + angka(a10, 2) + "puluh "
End If
If a1 > 0 Then
If posisi = 2 And a100 = 0 And a10 = 0 Then
baca = baca + angka(a1, 1)
Else
baca = baca + angka(a1, 2)
End If
End If
End If
ratus = baca
End Function

Function angka(x, posisi)
Select Case x
Case 0: angka = "Nol"
Case 1:
If posisi = 2 Then
angka = "Satu "
Else
angka = "Se"
End If
Case 2: angka = "Dua "
Case 3: angka = "Tiga "
Case 4: angka = "Empat "
Case 5: angka = "Lima "
Case 6: angka = "Enam "
Case 7: angka = "Tujuh "
Case 8: angka = "Delapan "
Case 9: angka = "Sembilan "
Case 10: angka = "Sepuluh "
Case 11: angka = "Sebelas "
Case 12: angka = "Duabelas "
Case 13: angka = "Tigabelas "
Case 14: angka = "Empatbelas "
Case 15: angka = "Limabelas "
Case 16: angka = "Enambelas "
Case 17: angka = "Tujuhbelas "
Case 18: angka = "Delapanbelas "
Case 19: angka = "Sembilanbelas "
End Select
End Function

'------------- Untuk Test ----------------
'Dim Test, nomorik
' nomorik = 2002003
' Test = terbilang(nomorik)

'Response.Write formatnumber(nomorik,0)
'Response.Write "
"
'Response.Write test

'------------- Untuk Test ----------------

%>

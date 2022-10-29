'-------------- Hitung Jumlah data/Records yang di-isi hari ini saja
'Armansyah

Dim adoCon_Hari_ini
Set adoCon_Hari_ini = Server.CreateObject("ADODB.Connection")
adoCon_Hari_ini.Open "DRIVER={Microsoft Access Driver (*.mdb)};uid=;pwd=dunia_it_armansyah;DBQ=" & Server.MapPath("ordermarket.mdb")
Dim rsOEFToday
Set rsOEFToday = Server.CreateObject("ADODB.Recordset")

'---------- Tentukan dulu variabel Tanggal, Bulan, Tahun ----
Dim Tanggal
Dim Bulan
Dim Tahun
Bulan = Month(date())
Tanggal = Day(date())

Tahun = "1999"
'Bisa diganti > Tahun = year(now)

Dim SQLOEFToday
SQLOEFTODAY = "Select * from Tabel1" & _
"Where ((day(WaktuPengisian)= " &Tanggal & ") And " & _
"(month(WaktuPengisian)= " &Bulan & ") And " & _
"(Year(WaktuPengisian)= " &Tahun & "))"
rsOEFToday.Open SQLOEFToday, adoCon_Hari_ini, adOpenStatic
Dim TotalOEFToday
TotalOEFToday = rsOEFToday.recordcount

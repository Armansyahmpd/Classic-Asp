<% 

  '-------------------------------------------------------------------- 
  'File ini bernama : persen_order_2_2006.asp 
  'Untuk Kelompok Kompas Gramedia Palembang 
  'Sistem Informasi Manajemen Terpadu (Simradu) 
  'Perhitungan data Order Entry File (OEF)
  'Tanggal Pembuatan : Rabu, 18 Juni 2003 Pkl. 15:22 
  'Programmer : Armansyah, S.Kom 
    '-------------------------------------------------------------------- 
    
    '-------------- Hitung Jumlah Semua data/Records yang ada ditabel OEF 
    Dim ActiveXDataObject_Connection 
    set ActiveXDataObject_Connection = Server.CreateObject("ADODB.Connection") 
    ActiveXDataObject_Connection.Open "DRIVER={Microsoft Access Driver (*.mdb)};uid=;pwd=akulahthebestcoy;DBQ=" & Server.MapPath("databasenyo.mdb") 
    
    Dim RecordSet_OEF Set RecordSet_OEF = Server.CreateObject("ADODB.Recordset") 
    Dim SQL_TotalSemuaOef 
    
        SQL_TotalSemuaOef = "Select * from Tabel1" 
        RecordSet_OEF.Open SQL_TotalSemuaOef, ActiveXDataObject_Connection, adOpenStatic 
        
        Dim Total_Semua_OEF 
        Total_Semua_OEF = RecordSet_OEF.recordcount RecordSet_OEF.Close 
        
  '-------------- Hitung Jumlah data/Records yang sudah selesai proses 
  
     Dim SQL_Order_Selesai 
         SQL_Order_Selesai = "Select * from Tabel1 Where Tabel1.StatusOrder = 'SUDAH SELESAI'" 
         RecordSet_OEF.Open SQL_Order_Selesai, ActiveXDataObject_Connection, adOpenStatic 
         
     Dim Total_Semua_OEFSELESAI 
     Total_Semua_OEFSELESAI = RecordSet_OEF.recordcount 
     RecordSet_OEF.Close 
     
     '-------------- Hitung Jumlah data/Records yang belum selesai proses 
     
     Dim SQL_Order_Blm_Selesai 
         SQL_Order_Blm_Selesai = "Select * from Tabel1 Where Tabel1.StatusOrder <> 'SUDAH SELESAI'" 
         RecordSet_OEF.Open SQL_Order_Blm_Selesai, ActiveXDataObject_Connection, adOpenStatic 
     
     Dim Total_Semua_OEFBLMSELESAI 
         Total_Semua_OEFBLMSELESAI = RecordSet_OEF.recordcount 
         RecordSet_OEF.Close 
         
 '-------------- Hitung Jumlah data/Records yang batal  
 
    Dim SQL_Order_yang_batal 
        SQL_Order_yang_batal = "Select * from Tabel1 Where Tabel1.StatusOrder = 'BATAL'" 
        RecordSet_OEF.Open SQL_Order_yang_batal, ActiveXDataObject_Connection, adOpenStatic 
        
    Dim Total_Semua_OEFBatal 
        Total_Semua_OEFBatal = RecordSet_OEF.recordcount 
    RecordSet_OEF.Close 
    
 %>

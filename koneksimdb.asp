 <% 
 
 'Contoh Koneksi Classic ASP dengan Access 2007 
 'Oleh : Armansyah 
 
 Dim koneksi 
 Dim rekamdata 
 Dim databasenya 
 
    databasenya= "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & _ Server.MapPath("database/accdb/07/test.accdb") 
    Set koneksi = Server.CreateObject("ADODB.Connection") 
    koneksi.open databasenya 
    Set rekamdata = Server.CreateObject("ADODB.Recordset") 
    
 Dim Sqlnya 
    Sqlnya = "select * from bukutamu" rekamdata.CursorLocation = 3 
    rekamdata.CursorType = 2 
    rekamdata.LockType = 3 
    rekamdata.Open Sqlnya, koneksi rekamdata.AddNew 
    rekamdata.Fields("nama") = request.form("form_nama") 
    rekamdata.Fields("alamat") = request.form("form_alamat") 
    rekamdata.Fields("WaktuKmptr") = now() 
    rekamdata.Update 
    rekamdata.Close 
    Set rekamdata = Nothing 
    Set koneksi= Nothing 
    
    '--- Akhiri dengan senyum keikhlasan .... :-) 
    '--- By : Armansyah 
    
    %>

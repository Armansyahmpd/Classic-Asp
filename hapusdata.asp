<% 

'Contoh kasus koding hapus data dari classic ASP ke database Access mdb
'KKG (Kelompok Kompas Gramedia Palembang) 2004
'Armansyah

  Dim adoCon 
  Dim rsDeleteEntry 
  Dim strSQL 
  Dim Waktu1 
  
      Waktu1=Request.Form("Waktu") 
      Set adoCon = Server.CreateObject("ADODB.Connection") 
      adoCon.Open "DRIVER={Microsoft Access Driver (*.mdb)};uid=;pwd=password; DBQ=" & Server.MapPath("dunia_it_armansyah.mdb") 
      Set rsDeleteEntry = Server.CreateObject("ADODB.Recordset") 
      StrSql = "Select * from Tabel_tabelan Where Tabel_tabelan.waktunyo = '" & Waktu1 & "'" 
      
      rsDeleteEntry.LockType = 3 
      rsDeleteEntry.Open strSQL, adoCon 
      
      rsDeleteEntry.Delete 
      rsDeleteEntry.Close 
      
      Set rsDeleteEntry = Nothing 
      Set adoCon = Nothing Response.Redirect "okebos.asp" 
      
%>

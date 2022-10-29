<%
'Contoh proses logoff dan simpan kedatabase
'Armansyah 2005

Dim adoLog
Dim rsLog
Dim ConLog
Dim LacakKukis
LacakKukis = Request.Cookies("kapanmasoknyo")

Set adoLog = Server.CreateObject("ADODB.Connection")
adoLog.Open "DRIVER={Microsoft Access Driver (*.mdb)};uid=;pwd=manomatokau;DBQ=" & Server.MapPath("../.../Database/mdb/masuklahoiy.mdb")
Set rsLog = Server.CreateObject("ADODB.Recordset")

Dim SqlLog
SqlLog = "Select * from logoff Where waktuloginid Like '%" & LacakKukis & "%' "

rsLog.CursorType = 2
rsLog.LockType = 3
rsLog.Open SqlLog, adoLog

Dim Bla2
Dim DatNow2
DatNow2 = now()
Bla2 = Server.HTMLEncode(FormatDateTime(datNow2,3))

Dim TanggalSekarangId
TanggalSekarangId = Server.HTMLEncode(FormatTanggal(now()))

Dim TanggalIndonesia
TanggalIndonesia = TanggalSekarangId + " " + Bla2

rsLog.Fields("waktulogoffid") = TanggalIndonesia
rsLog.Fields("waktulogoff") = now()

rsLog.Update
rsLog.Close
Set rsLog = Nothing
Set adoLog = Nothing

'-------- akhir proses penyimpanan data login

If Session("iyodakye") = True then
Session.Contents.Remove("iyodakye")


Else

If Session("iyodakye") = False or IsNull(Session("iyodakye")) = True then

End If
End If
%>

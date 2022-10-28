<%
Dim fso, MyFile


Set fso = CreateObject("Scripting.FileSystemObject")

Set MyFile = fso.GetFile(Server.MapPath("data_arman.mdb"))
MyFile.Copy (Server.MapPath("backupofdata_armanmdb"))

%>

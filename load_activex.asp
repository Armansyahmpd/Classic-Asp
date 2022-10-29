<%
'Contoh integrasi ActiveX-VB6 dengan Classic ASP

Dim komponenvbku
set komponenvbku = Server.CreateObject("roundmin2.VB6forASP")

'Pastikan komponen roundmin2.dll sudah diregister kesystem
'roundmin2.dll adalah component fungsi round -2 yg dibuat dengan VB6 untuk ASP

if MBruto <>
MPPH = 0
else

if MBruto <= MSubsII Then
MPPH = komponenvbku.FungsiRound(MPPH-MPPHSub*1, -2)

else
MPPH = komponenvbku.FungsiRound(MPPH*1, -2)

end if
end if

%>

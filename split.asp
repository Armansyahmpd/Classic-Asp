<% 
'Contoh kasus koding di KKG 2003

Dim Test_Split, Data1, Test_Split2,Data2, Data3, Test_Split3, SisipTeks1, SisipTeks2 

    SisipTeks = " EDISI " 
    SisipTeks2 = date() 
    
    Data1 = "SKH KOMPAS ED. 16 - 30 NOVEMBER 2004" 
    Test_Split = split(Data1,"ED") 
    
    Data2 = "SKH. TRANSPARAN ED. NOVEMBER 2004" 
    Test_Split2 = split(Data2,"ED") 
    
    Data3 = "SKH SRIPO ED. NOVEMBER 2004" 
    Test_Split3 = split(Data3,"ED") 
    
    Response.Write Data1 
    Response.Write "" 
    Response.Write "Sesudah di split hasilnya > " Response.Write Test_Split(0) 
    Response.Write ""

    Response.Write Data2 Response.Write "" 
    Response.Write "Sesudah di split hasilnya > " Response.Write Test_Split2(0) 
    Response.Write ""
    Response.Write Data3 
    Response.Write "" 
    
    Response.Write "Sesudah di split hasilnya > " Response.Write Test_Split3(0) 
    Response.Write SisipTeks 
    Response.Write SisipTeks2 
    
  %>

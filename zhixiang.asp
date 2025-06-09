<% 
Function gen_key(digits)
dim char_array(50)
char_array(0) = "0"
char_array(1) = "1"
char_array(2) = "2"
char_array(3) = "3"
char_array(4) = "4"
char_array(5) = "5"
char_array(6) = "6"
char_array(7) = "7"
char_array(8) = "8"
char_array(9) = "9"
randomize
do while len(output) < digits
cnum = char_array(Int((10 - 0 + 1) * Rnd + 0))
output = output + cnum
loop
gen_key = output
End Function
comm=gen_key(9)
xiangjia=request.Form("xiangjia")
xiangshu=request.Form("xiangshu")
use_dep=trim(request.Form("use_dep"))
uptime=trim(request.Form("uptime"))

set rs=server.createobject("adodb.recordset") 
sql="select * from chengping_out_store"
rs.open sql,conn,1,3
rs.addnew
rs("fenlei") = "包装"
rs("pinming") = "纸箱"
rs("comm") =comm
rs("uptime") = request.Form("uptime")
rs("use_dep") = trim(request.Form("use_dep"))
rs("Loan_num") = xiangshu
rs("Loan_price") =xiangjia
rs("Loan_Amount") =xiangshu*xiangjia
rs.update
rs.close
set rs=nothing
end if
set rs2=server.createobject("adodb.recordset") 
sql2="select * from JS_out_store"
rs2.open sql2,conn,1,3
rs2.addnew
rs2("fenlei") = "包装"
rs2("pinming") = "纸箱"
rs2("uptime") = request.Form("uptime")
rs2("use_dep") = use_dep
rs2("Loan_num") = xiangshu
rs2("comm") =comm
rs2("Loan_price") = xiangjia
rs2("Loan_Amount") = xiangshu*xiangjia
rs2.update
rs2.close
set rs2=nothing

set rs3=server.createobject("adodb.recordset")
sql3="select * from JS_huzong where kehu='"&use_dep&"'"
rs3.open sql3,conn,2,3
if not rs3.eof then
rs3("Loan_Amount")=rs3("Loan_Amount")+(xiangshu*xiangjia)
else 
rs3.addnew
rs3("Loan_Amount")=xiangshu*xiangjia
rs3("kehu") = trim(request.Form("use_dep"))
rs3.update
rs3.close
set rs3=nothing
            

%>
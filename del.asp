<!--#include file="conn.asp"-->
<%
set rs1=server.createobject("adodb.recordset")
sql1="delete from in_store where Loan_num>0"
rs1.open sql1,conn,1,3
set rs1=nothing
response.Write "<script language=javascript>alert('É¾³ý³É¹¦£¡');</script>"
response.end
%>


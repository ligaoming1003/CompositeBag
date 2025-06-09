<!--#include file="conn.asp"-->
<!--#include file="checkuser.asp"-->
<html> 
<head> 
<script LANGUAGE="javascript"> 
<!-- 
function openwin() { 
window.open ("page.html", "newwindow", "height=400, width=400, toolbar=no, menubar=no, scrollbars=no, resizable=no, location=no, status=no") 
//写成一行 
} 
//--> 
</script> 
</head> 
<body onload="openwin()"> 
<%
set rs=server.createobject("adodb.recordset") 
sql="select * from fuliao_store"
rs.open sql,conn,1,3
response.Write "高压粒籽不多了！只剩下rs("number")公斤了。"
rs.update
rs.close
set rs=nothing
%>
</body> 
</html> 

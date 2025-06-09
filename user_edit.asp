<!--#include file="conn.asp"-->
<!--#include file="checkuser.asp"-->
<!--#include file="md5.asp" -->
<%
dim action
action=request("action")
select case action
case "edit" 
passwd=request.form("pwd")
passwd1=md5(trim(request.form("pwd")))   
set rs=server.createobject("adodb.recordset") 
sql="select * from userinfo where userid="&userid
rs.open sql,conn,1,3


if passwd<>rs("pwd") then
rs("pwd")=passwd1

end if
rs.update
rs.close
set rs=nothing
response.Write "<script language=javascript>alert('修改成功！');</script>"
response.Write "<script language=javascript>window.history.go(-2)</script>"
response.end
end select
%>
<html>
<head>
<title>∷嘉美管理系统∷</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="css/main.css" rel="stylesheet" type="text/css">

<SCRIPT language=JavaScript src="css/User_Info_Modify.js"></SCRIPT>
<style type="text/css">
<!--
td {  font-family: "宋体"; font-size: 9pt}
body {  font-family: "宋体"; font-size: 9pt}
select {  font-family: "宋体"; font-size: 9pt}
A {text-decoration: none; color: #336699; font-family: "宋体"; font-size: 9pt}
A:hover {text-decoration: underline; color: #FF0000; font-family: "宋体"; font-size: 9pt} 
-->
</style>
</head>
<BODY topMargin=0 rightMargin=0 leftMargin=0>       
<!--#include file="top.asp"--><br>
<br>
<fieldset style="width:50%"><legend>编辑个人资料</legend>
<table width="300" border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#D4D0C8">
          <form name="powersearch" method="post" action="user_edit.asp" ><input name="action" value="edit" type="hidden">
		  <%set rs=server.CreateObject("ADODB.RecordSet")
          sql="select * from userinfo where userid="&userid
          rs.open sql,conn,1,1
		  %>
		  <tr><td width="100" align="right" height="22">用户名：&nbsp;</td><td width="200">&nbsp;<%=name%></td></tr>
		  <tr><td width="100" align="right" height="22">新密码：&nbsp;</td><td width="200">&nbsp;<input name="pwd"  type="password"  size="10"  value='<%= rs("pwd")%>'></td></tr>
		  <tr><td width="100" align="right" height="22">　</td><td width="200">　</td></tr>
		  <tr><td width="100" align="right" height="22">　</td><td width="200">　</td></tr>
			<tr><td colspan="4" height="30" align="center"><input type="submit"  value="编辑"></td></tr>
			<%
			rs.close
			set rs=nothing
			%>
			</form>
		  </table>
</fieldset>
</body>
</html>

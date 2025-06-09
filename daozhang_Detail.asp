<!--#include file="conn.asp"-->
<!--#include file="checkuser.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="css/main.css" rel="stylesheet" type="text/css">
<title>∷嘉美管理系统∷</title>
<!--media=print 这个属性可以在打印时有效--> 
<style media=print> 
.Noprint{display:none;} 
.PageNext{page-break-after: always;} 
</style>

</head>

<body leftmargin="30" topmargin="20">
<%
dim SeachModel
SeachModel=request.form("SeachModel")
yinhang=trim(request.form("yinhang"))
kehu=request.form("kehu")
fanshi=request.form("fanshi")

uptime=cstr(trim(request.form("uptime")))
uptime1=cstr(trim(request.form("uptime1")))
end_time=cstr(trim(request.form("end_time")))
end_time1=cstr(trim(request.form("end_time1")))
nian=cstr(trim(request.form("nian")))
'查询代码开始
names=Split(SeachModel,", ")
i=0
sql="select * from daozhang where"
for each name in names
if names(i)="1" then
sql=sql+" and kehu = '"&kehu&"'"
end if
if names(i)="2" then
sql=sql+" and fanshi like '%"&fanshi&"%'"
end if
if names(i)="3" then
sql=sql+" and uptime  between  '"&uptime&"' and '"&uptime1&"'"
end if
if names(i)="4" then
sql=sql+" and yinhang = '"&yinhang&"'"
end if
if names(i)="5" then
sql=sql+" and uptime  between  '"&uptime&"' and '"&uptime1&"'"
end if

if names(i)="6" then
sql=sql+" and year(uptime) = '"&nian&"'"
end if



i=i+1
next
sql=sql+" order by -id"
set rs=server.createobject("adodb.recordset")
sql=Replace(sql, "where and", "where")

%>

<table width="800" border="1" cellpadding="0" cellspacing="0" bgcolor="#ffffff"  bordercolor="#000000">
          <tr><td height="35" colspan="7" align="center"><b>
			<font face="黑体" size="4">货</font></b><font size="4" face="黑体"><b>款到账汇总明细</b></font></td></tr>
		  <tr> 
            <td width="10%" height="22" align="center" >出库日期</td>
            <td width="10%" align="center" >类 别</td>
            <td width="20%" align="center" >客 户</td>
			<td width="15%" align="center" >金 额</td>
            <td width="10%" align="center" >方 式</td>
			<td width="20%" align="center">银 行</td>
			<td width="15%" align="center">备 注</td>
          </tr>
          
                  <% 
				  dim i
				  i=0
				  Loan_num=0
				  Loan_Amount=0
         rs.open sql,conn,1,1
		 if rs.eof and rs.bof then
		response.write "<tr><td><br><br><center>还没有符合条件的查询！</center><br><br></td></tr>"
		else
		 do while not rs.eof
		 i=i+1
		   %>
                  <tr>
<td height="18" width="10%" align="center" ><%=rs("uptime")%></td>
<td align="center" ><%=rs("leixin")%></td>
<td   align="center"><%=rs("kehu")%></td>
<td   align="center"><%=rs("Loan_num")%></td>
<td   align="center"><%=rs("fanshi")%></td>
<td align="center"><%=rs("yinhang")%></td>
<td align="center"></td>
</tr>
				  
                  <%
Loan_num=Loan_num+rs("Loan_num")

rs.movenext
loop
end if
%>
                   <tr><td colspan="2" height="30"> 记录总共为 <font color="#cc0000"><%=i%></font> 条</td><td  align="right">总计：</td><td align="center"><%=Loan_num%></td><td align="center"></td><td>　</td><td></td></tr>
				</table>
				<table>
				<tr><td height="30" colspan="10" align="center"><OBJECT id=WebBrowser classid=CLSID:8856F961-340A-11D0-A96B-00C04FD705A2 height=0 width=0 VIEWASTEXT> 
</OBJECT> 
<input type=button value=打印 onClick="document.all.WebBrowser.ExecWB(6,1)" class="NOPRINT"> 
<input type=button value=直接打印 onClick="document.all.WebBrowser.ExecWB(6,6)" class="NOPRINT"> 
<input type=button value=页面设置 onClick="document.all.WebBrowser.ExecWB(8,1)" class="NOPRINT"> 
<input type=button value=打印预览 onClick="document.all.WebBrowser.ExecWB(7,1)" class="NOPRINT">
</td></tr>
				</table>
				
</body>
</html>
<!--#include file="conn.asp"-->
<!--#include file="checkuser.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="css/main.css" rel="stylesheet" type="text/css">
<title>�˼�������ϵͳ��</title>
<!--media=print ������Կ����ڴ�ӡʱ��Ч--> 
<style media=print> 
.Noprint{display:none;} 
.PageNext{page-break-after: always;} 
</style>

</head>

<body leftmargin="30" topmargin="20">
<%
dim SeachModel
SeachModel=request.form("SeachModel")
pinming=trim(request.form("pinming"))
Loan_Amount=trim(request.form("Loan_Amount"))
use_dep=request.form("use_dep")
hkdz=request.form("hkdz")
uptime=cstr(trim(request.form("uptime")))
uptime1=cstr(trim(request.form("uptime1")))
nian=cstr(trim(request.form("nian")))

'��ѯ���뿪ʼ
names=Split(SeachModel,", ")
i=0
sql="select * from JS_out_Store where pihao='ok'"
for each name in names
if names(i)="1" then
sql=sql+" and use_dep = '"&use_dep&"'"
end if
if names(i)="2" then
sql=sql+" and uptime  between  '"&uptime&"' and '"&uptime1&"'"
end if
if names(i)="3" then
sql=sql+" and uptime  between  '"&uptime&"' and '"&uptime1&"'"
end if
if names(i)="4" then
sql=sql+" and year(uptime) = '"&nian&"'"
end if
i=i+1
next
sql=sql+" order by -id"
set rs=server.createobject("adodb.recordset")
sql=Replace(sql, "where and", "where")

%>

<table width="750" border="1" cellpadding="0" cellspacing="0" bgcolor="#ffffff"  bordercolor="#000000">
          <tr><td height="35" colspan="6" align="center">
			<font size="4" face="����"><b>�ͻ�������ϸ</b></font></td></tr>
		  <tr> 
            <td width="10%" height="22" align="center" >����</td>
            <td width="10%" align="center" >�ͻ�</td>
            <td width="31%" align="center" >��&nbsp;&nbsp;&nbsp; ��</td>
			<td width="12%" align="center" >�ͻ�����</td>
            <td width="12%" align="center" >�ͻ��շ�</td>
            <td width="22%" align="center" >��ע</td>
          </tr>
          
                  <% 
				  dim i
				  i=0
				  Loan_Amount=0
				  hkdz=0
         rs.open sql,conn,1,1
		 if rs.eof and rs.bof then
		response.write "<tr><td><br><br><center>��û�з��������Ĳ�ѯ��</center><br><br></td></tr>"
		else
		 do while not rs.eof
		 i=i+1
		   %>
                  <tr>
<td height="18" width="10%" align="center" ><%=rs("uptime")%></td>
<td align="center" ><%=rs("use_dep")%></td>
<td   align="center"><%=rs("pinming")%></td>
<td   align="center"><%=rs("Loan_Amount")%></td>
<td   align="center"><%=rs("hkdz")%></td>
<td   align="center">��</td>
</tr>
				  
                  <%
Loan_Amount=Loan_Amount+rs("Loan_Amount")
hkdz=hkdz+rs("hkdz")

rs.movenext
loop
end if
%>
                   <tr><td colspan="2" height="30"> ��¼�ܹ�Ϊ <font color="#cc0000"><%=i%></font> ��</td><td  align="right">�ܼƣ�</td><td align="center"><%=Loan_Amount%></td><td align="center"><%=hkdz%></td><td align="center"><%if Loan_Amount-hkdz>0 then%>�ͻ�Ƿ�<%=Loan_Amount-hkdz%>Ԫ<%else%>�ͻ����<%=hkdz-Loan_Amount%>Ԫ<%end if%></td></tr>
				</table>
				<table>
				<tr><td height="30" colspan="10" align="center"><OBJECT id=WebBrowser classid=CLSID:8856F961-340A-11D0-A96B-00C04FD705A2 height=0 width=0 VIEWASTEXT> 
</OBJECT> 
<input type=button value=��ӡ onClick="document.all.WebBrowser.ExecWB(6,1)" class="NOPRINT"> 
<input type=button value=ֱ�Ӵ�ӡ onClick="document.all.WebBrowser.ExecWB(6,6)" class="NOPRINT"> 
<input type=button value=ҳ������ onClick="document.all.WebBrowser.ExecWB(8,1)" class="NOPRINT"> 
<input type=button value=��ӡԤ�� onClick="document.all.WebBrowser.ExecWB(7,1)" class="NOPRINT">
</td></tr>
				</table>
				
</body>
</html>
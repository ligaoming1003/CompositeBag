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
id=request.form("id")
SeachModel=request.form("SeachModel")
cname=trim(request.form("cname"))
pihao=request.form("pihao")
uptime=cstr(trim(request.form("uptime")))
uptime1=cstr(trim(request.form("uptime1")))
end_time=cstr(trim(request.form("end_time")))
end_time1=cstr(trim(request.form("end_time1")))
nian=cstr(trim(request.form("nian")))
'��ѯ���뿪ʼ
names=Split(SeachModel,", ")
i=0
sql="select * from fuhe_in where"
for each name in names
if names(i)="1" then
sql=sql+" and pihao = '"&pihao&"'"
end if

if names(i)="5" then
sql=sql+" and cname like '%"&cname&"%'"
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
sql=sql+""
set rs=server.createobject("adodb.recordset")
sql=Replace(sql, "where and", "where")

%>

<table width="600" border="1" cellpadding="0" cellspacing="0" bgcolor="#ffffff"  bordercolor="#000000">
          <tr><td height="35" colspan="12" align="center">
			<font size="4" face="����"><b>�������Ʒ������ϸ</b></font></td></tr>
		  <tr> 
            <td width="10%" height="22" align="center" >��������</td>
            <td width="50%" align="center" >Ʒ ��</td>
            <td width="10%" align="center" >��Ʒ</td>
            <td width="10%" align="center" >���Ʒ</td>
            <td width="10%" align="center" >Ա��</td>
			<td width="10%" align="center" >�� λ</td>

          </tr>
          
                  <% 
				  dim i
				  i=0
				  loan_oan=0
                  one_fu=0
         rs.open sql,conn,1,1
		 if rs.eof and rs.bof then
		response.write "<tr><td><br><br><center>��û�з��������Ĳ�ѯ��</center><br><br></td></tr>"
		else
		 do while not rs.eof
		 i=i+1
		   %>
                  <tr>
<td height="18" width="10%" align="center" ><%=rs("uptime")%></td>
<td  align="center" ><%=rs("cname")%></td>
<td  align="center"><%if rs("loan_oan")="0" then%>&nbsp;<%else%><%=rs("loan_oan")%><%end if%></td>
<td  align="center"><%if rs("one_fu")="0" then%>&nbsp;<%else%><%=rs("one_fu")%><%end if%></td>
<td  align="center"><%=rs("pihao")%></td>
<td   align="center">ǧ��</td>

</tr>
				  
                  <%
loan_oan=loan_oan+rs("loan_oan")
one_fu=one_fu+rs("one_fu")
numm=one_fu+loan_oan
rs.movenext
loop
end if
%>
                   <tr><td colspan="3" height="30"> ���ϲ�ѯ�����ĳ����¼�ܹ�Ϊ <font color="#cc0000"><%=i%></font> ��</td><td  align="right">�ܼƣ�</td><td align="center"><%=numm%>&nbsp;</td><td align="center">
��</td></tr>
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


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
use_dep=request.form("use_dep")
lingliaoyuan=request.form("lingliaoyuan")
fenlei1=request.form("fenlei")
guige=request.form("guige")
cname=request.form("cname")
uptime=cstr(trim(request.form("uptime")))
uptime1=cstr(trim(request.form("uptime1")))
end_time=cstr(trim(request.form("end_time")))
end_time1=cstr(trim(request.form("end_time1")))
nian=cstr(trim(request.form("nian")))
'��ѯ���뿪ʼ
names=Split(SeachModel,", ")
i=0
sql="select * from cailiao_out_Store where"
for each name in names
if names(i)="1" then
sql=sql+" and fenlei = '"&fenlei1&"'"
end if
if names(i)="5" then
sql=sql+" and uptime  between  '"&uptime&"' and '"&uptime1&"'"
end if
if names(i)="6" then
sql=sql+" and lingliaoyuan = '"&lingliaoyuan&"'"
end if
if names(i)="7" then
sql=sql+" and uptime  between  '"&uptime&"' and '"&uptime1&"'"
end if
if names(i)="9" then
sql=sql+" and cname = '"&cname&"'"
end if
if names(i)="8" then
sql=sql+" and year(uptime) = '"&nian&"'"
end if
if names(i)="2" then
sql=sql+" and pinming like '%"&pinming&"%'"
end if
if names(i)="4" then
sql=sql+" and use_dep = '"&use_dep&"'"
end if
if names(i)="3" then
sql=sql+" and guige = '"&guige&"'"
end if
i=i+1
next
sql=sql+" order by -id"
set rs=server.createobject("adodb.recordset")
sql=Replace(sql, "where and", "where")

%>

<table width="960" border="1" cellpadding="0" cellspacing="0" bgcolor="#ffffff"  bordercolor="#000000">
          <tr><td height="35" colspan="9" align="center"><font size="4" face="����"><b>������ϻ�����ϸ</b></font></td></tr>
		  <tr> 
            <td width="70" height="22" align="center" >��������</td>
            <td width="83" align="center" >�� ��</td>
            <td width="69" align="center" >Ʒ ��</td>
            <td width="92" align="center" >�� ��</td>
            <td width="155" align="center" >��Ʒ��</td>
            <td width="157" align="center" >�� ��</td>
			<td width="61" align="center" >�� λ</td>

			<td width="107" align="center" >���ò���</td>

			<td width="100" align="center">�� ע</td>
          </tr>
          
                  <% 
				  dim i
				  i=0
				  Loan_num=0
				  Loan_Amount=0
         rs.open sql,conn,1,1
		 if rs.eof and rs.bof then
		response.write "<tr><td><br><br><center>��û�з��������Ĳ�ѯ��</center><br><br></td></tr>"
		else
		 do while not rs.eof
		 i=i+1
		   %>
                  <tr>
<td height="18" width="10%" align="center" ><%=rs("uptime")%></td>
<td align="center" ><%=rs("fenlei")%></td>
<td  align="center"><%=rs("pinming")%></td>
<td  align="center"><%=rs("guige")%>*<%=rs("houdou")%></td>
<td  align="center"><%=rs("cname")%><%if rs("cname")="" then%>&nbsp;<%end if%></td>
<td align="center"><%=Formatnumber(rs("Loan_num"),2,-1,0,0)%></td>
<td   align="center"><%=rs("Unit")%></td>

<td   align="center"><%if rs("use_dep")<>"" then%><%=rs("use_dep")%><%else%>��Ϊ<%=rs("kufang")%><%=rs("old_fenlei")%><%end if%></td>

<td align="center"><%=rs("content")%>&nbsp;</td>
</tr>
				  
                  <%
Loan_num=Loan_num+rs("Loan_num")
Loan_Amount=Loan_Amount+rs("Loan_Amount")
rs.movenext
loop
end if
%>
                   <tr><td colspan="4" height="30"> ���ϲ�ѯ�����ĳ����¼�ܹ�Ϊ <font color="#cc0000"><%=i%></font> ��</td><td  align="right">�ܼƣ�</td><td align="center"><%=Loan_num%>&nbsp;</td><td align="center"><%=Loan_Amount%></td></tr>
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
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
pinming=trim(request.form("pinming"))
Supplier=request.form("Supplier")
fenlei1=request.form("fenlei")
caigouyuan=request.form("caigouyuan")
guige=request.form("guige")
uptime=cstr(trim(request.form("uptime")))
uptime1=cstr(trim(request.form("uptime1")))
end_time=cstr(trim(request.form("end_time")))
end_time1=cstr(trim(request.form("end_time1")))
nian=cstr(trim(request.form("nian")))
'��ѯ���뿪ʼ
names=Split(SeachModel,", ")
i=0
sql="select * from chengping_In_Store where"
for each name in names
if names(i)="1" then
sql=sql+" and fenlei = '"&fenlei1&"'"
end if
if names(i)="5" then
sql=sql+" and uptime  between  '"&uptime&"' and '"&uptime1&"'"
end if
if names(i)="6" then
sql=sql+" and caigouyuan = '"&caigouyuan&"'"
end if
if names(i)="7" then
sql=sql+" and uptime  between  '"&uptime&"' and '"&uptime1&"'"
end if

if names(i)="8" then
sql=sql+" and year(uptime) = '"&nian&"'"
end if
if names(i)="2" then
sql=sql+" and pinming like '%"&pinming&"%'"
end if
if names(i)="4" then
sql=sql+" and Supplier = '"&Supplier&"'"
end if
if names(i)="3" then
sql=sql+" and guige = '"&guige&"'"
end if
i=i+1
next
sql=sql+""
set rs=server.createobject("adodb.recordset")
sql=Replace(sql, "where and", "where")

%>

<table width="600" border="1" cellpadding="0" cellspacing="0" bgcolor="#ffffff"  bordercolor="#000000">
          <tr><td height="35" colspan="12" align="center"><font size="4" face="����"><b>����Ʒ������ϸ</b></font></td></tr>
		  <tr> 
            <td width="15%" height="22" align="center" >�������</td>
            <td width="15%" align="center" >�� ��</td>
            <td width="20%" align="center" >Ʒ ��</td>
            <td width="20%" align="center" >�� ��</td>
            <td width="20%" align="center" >�� ��</td>
			<td width="10%" align="center" >�� λ</td>

          </tr>
          
                  <% 
				  dim i
				  i=0
				  use_num=0
				  use_Amount=0
         rs.open sql,conn,1,1
		 if rs.eof and rs.bof then
		response.write "<tr><td><br><br><center>��û�з��������Ĳ�ѯ��</center><br><br></td></tr>"
		else
		 do while not rs.eof
		 i=i+1
		   %>
                  <tr>
<td height="18" align="center" ><%=rs("uptime")%></td>
<td  align="center" ><%=rs("fenlei")%></td>
<td  align="center"><%=rs("pinming")%></td>
<td  align="center"><%=rs("guige")%>*<%=rs("kuang")%></td>
<td  align="center"><%=Formatnumber(rs("use_num"),2,-1,0,0)%></td>
<td   align="center"><%=rs("Unit")%></td>

</tr>
				  
                  <%
use_num=use_num+rs("use_num")
use_Amount=use_Amount+rs("use_Amount")
rs.movenext
loop
end if
%>
                   <tr><td colspan="3" height="30"> ���ϲ�ѯ����������¼�ܹ�Ϊ <font color="#cc0000"><%=i%></font> ��</td><td  align="right">�ܼƣ�</td><td align="center"><%=use_num%>&nbsp;</td><td align="center">
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

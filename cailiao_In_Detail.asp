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
danjia=request.form("danjia")
fenlei1=request.form("fenlei")
jinge=request.form("jinge")
guige=request.form("guige")
gonghuodanwei=request.form("gonghuodanwei")
uptime=cstr(trim(request.form("uptime")))
uptime1=cstr(trim(request.form("uptime1")))
end_time=cstr(trim(request.form("end_time")))
end_time1=cstr(trim(request.form("end_time1")))
nian=cstr(trim(request.form("nian")))
'��ѯ���뿪ʼ
names=Split(SeachModel,", ")
i=0
sql="select * from cailiao_In_Store where"
for each name in names
if names(i)="1" then
sql=sql+" and fenlei = '"&fenlei1&"'and pinming<>'���ڿͻ�����' and pinming<>'���ڿͻ�Ƿ��'"
end if
if names(i)="5" then
sql=sql+" and uptime  between  '"&uptime&"' and '"&uptime1&"'and pinming<>'���ڿͻ�����' and pinming<>'���ڿͻ�Ƿ��'"
end if
if names(i)="6" then
sql=sql+" and caigouyuan = '"&caigouyuan&"'and pinming<>'���ڿͻ�����' and pinming<>'���ڿͻ�Ƿ��'"
end if
if names(i)="7" then
sql=sql+" and uptime  between  '"&uptime&"' and '"&uptime1&"'and pinming<>'���ڿͻ�����' and pinming<>'���ڿͻ�Ƿ��'"
end if

if names(i)="8" then
sql=sql+" and year(uptime) = '"&nian&"'and pinming<>'���ڿͻ�����' and pinming<>'���ڿͻ�Ƿ��'"
end if
if names(i)="2" then
sql=sql+" and pinming='"&pinming&"'and pinming<>'���ڿͻ�����' and pinming<>'���ڿͻ�Ƿ��'"
end if
if names(i)="4" then
sql=sql+" and gonghuodanwei='"&gonghuodanwei&"'and pinming<>'���ڿͻ�����' and pinming<>'���ڿͻ�Ƿ��'"
end if
if names(i)="3" then
sql=sql+" and guige = '"&guige&"'and pinming<>'���ڿͻ�����' and pinming<>'���ڿͻ�Ƿ��'"
end if
i=i+1
next
sql=sql+" order by -id"
set rs=server.createobject("adodb.recordset")
sql=Replace(sql, "where and", "where")

%>

<table width="650" border="1" cellpadding="0" cellspacing="0" bgcolor="#ffffff"  bordercolor="#000000">
          <tr><td height="35" colspan="12" align="center"><font size="4" face="����"><b>�����ϻ�����ϸ</b></font></td></tr>
		  <tr> 
            <td width="60" height="22" align="center" >�������</td>
            <td width="40" align="center" >�� ��</td>
            <td width="50" align="center" >Ʒ ��</td>
            <td width="60" align="center" >�� ��</td>
            <td width="40" align="center" >�� λ</td>
			<td width="60" align="center" >�� ��</td>
			<td width="60" align="center" >����</td>
			<td width="60" align="center" >���</td>
			<td width="90" align="center" >������λ</td>
			<td width="60" align="center">�� ע</td>
          </tr>
          
                  <% 
				  dim i
				  i=0
				  use_num=0
                  jinge=0
         rs.open sql,conn,1,1
		 if rs.eof and rs.bof then
		response.write "<tr><td><br><br><center>��û�з��������Ĳ�ѯ��</center><br><br></td></tr>"
		else
		 do while not rs.eof
		 i=i+1
		   %>
                  <tr>
<td height="18" width="10%" align="center" ><%=rs("uptime")%></td>
<td  align="center" ><%=rs("fenlei")%></td>
<td  align="center"><%=rs("pinming")%></td>
<td  align="center"><%=rs("guige")%>*<%=rs("houdou")%></td>
<td  align="center"><%=rs("Unit")%></td>
<td  align="center"><%=Formatnumber(rs("use_num"),2,-1,0,0)%></td>
<td  align="center"><%=Formatnumber(rs("danjia"),2,-1,0,0)%></td>
<td  align="center"><%=Formatnumber(rs("jinge"),2,-1,0,0)%></td>
<td  align="center"><%if rs("old_fenlei")<>"" then%>��<%=rs("old_fenlei")%><%=rs("end_time")%><%=rs("gonghuodanwei")%><%else%><%=rs("gonghuodanwei")%><%end if%></td>
<td align="center">&nbsp;<%=rs("content")%></td>
</tr>
				  
                  <%
use_num=use_num+rs("use_num")
jinge=jinge+rs("jinge")
rs.movenext
loop
end if
%>
                   <tr><td colspan="4" height="30"> ���ϲ�ѯ����������¼�ܹ�Ϊ <font color="#cc0000"><%=i%></font> ��</td><td  align="right">�ܼƣ�</td><td align="center"><%=use_num%></td>
                   <td align="center"><%=Formatnumber(jinge/use_num,2,-1,0,0)%></td><td>
			<p align="center"><%=jinge%></td><td></td><td></td></tr>
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
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
SeachModel=request.Form("SeachModel")
pinming=trim(request.Form("pinming"))
fenlei1=request.Form("fenlei")
guige=request.Form("guige")


'��ѯ���뿪ʼ
names=Split(SeachModel,", ")
i=0
sql="select * from chengping_Store where number<>0"
for each name in names
if names(i)="1" then
sql=sql+" and fenlei = '"&fenlei1&"'"
end if
if names(i)="2" then
sql=sql+" and pinming like '%"&pinming&"%'"
end if

if names(i)="3" then
sql=sql+" and guige = '"&guige&"'"
end if
i=i+1
next
sql=sql+""
set rs=server.createobject("adodb.recordset")
sql=Replace(sql, "where and", "where")

%><br />

<table width="600" border="1" cellpadding="0" cellspacing="0" bgcolor="#ffffff"  bordercolor="#000000">
          <tr><td height="35" colspan="10" align="center"><font size="4" face="����"><b>����Ʒ������ϸ</b></font></td></tr>
		  <tr> 
            <td width="120" align="center"  height="22"><b>�� ��</b></td>
            <td width="150" align="center" ><b>Ʒ ��</b></td>
            <td width="120" align="center" ><b>�� ��</b></td>
            <td width="110" align="center" ><b>�� ��</b></td>
			<td width="100" align="center" ><b>�� λ</b></td>
          </tr>
          
                  <% 
				  dim i
				  i=0
				  num=0
         rs.open sql,conn,1,1
		 if rs.eof and rs.bof then
		response.write "<tr><td><br><br><center>��û�з��������Ĳ�ѯ��</center><br><br></td></tr>"
		else
		 do while not rs.eof
		 i=i+1
		   %>
                  <tr>
<td height="22" align="center" ><%=rs("fenlei")%><%if rs("fenlei")="" then%>&nbsp;<%end if%></td>
<td  align="center">&nbsp;<%=rs("pinming")%></td>
<td  align="center">&nbsp;<%=rs("guige")%>*<%=rs("kuang")%></td>
<td align="center"><%=Formatnumber(rs("number"),2,-1,0,0)%></td>
<td   align="center"><%=rs("Unit")%><%if rs("Unit")="" then%>&nbsp;<%end if%></td>
</tr>
				  
                  <%
num=num+rs("number")
rs.movenext
loop
end if
%>
                   <tr><td colspan="2" height="30"> &nbsp;���ϲ�ѯ�����Ŀ���¼�ܹ�Ϊ <font color="#cc0000"><%=i%></font> ��</td><td   align="center">�ܼƣ�</td><td  align="center"><%=num%>&nbsp;</td><td  align="center">��</td></tr>
				</table>
				<table>
				<tr><td height="30" colspan="10" align="center"><OBJECT id=WebBrowser classid=CLSID:8856F961-340A-11D0-A96B-00C04FD705A2 height=0 width=0 VIEWASTEXT> 
</OBJECT> 
<input type=button value=��ӡ onclick="document.all.WebBrowser.ExecWB(6,1)" class="NOPRINT"> 
<input type=button value=ֱ�Ӵ�ӡ onclick="document.all.WebBrowser.ExecWB(6,6)" class="NOPRINT"> 
<input type=button value=ҳ������ onclick="document.all.WebBrowser.ExecWB(8,1)" class="NOPRINT"> 
<input type=button value=��ӡԤ�� onclick="document.all.WebBrowser.ExecWB(7,1)" class="NOPRINT">
</td></tr>
				</table>
				
</body>
</html>

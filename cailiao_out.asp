<!--#include file="conn.asp"-->
<!--#include file="checkuser.asp"-->
<html>
<head>
<title>�˼�������ϵͳ��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="css/main.css" rel="stylesheet" type="text/css">
<SCRIPT language=JavaScript src="css/User_Info_Modify.js"></SCRIPT>

<SCRIPT language=javascript  src="js/selectcity.js"></script>
<script language="JavaScript" src="js/validate.js" type="text/JavaScript"></script>

<SCRIPT language=javascript src="css/init.js"></SCRIPT>

<style type="text/css">
<!--
td {  font-family: "����"; font-size: 9pt}
body {  font-family: "����"; font-size: 9pt}
select {  font-family: "����"; font-size: 9pt}
A {text-decoration: none; color: #336699; font-family: "����"; font-size: 9pt}
A:hover {text-decoration: underline; color: #FF0000; font-family: "����"; font-size: 9pt} 
-->
</style>




<SCRIPT LANGUAGE=javascript>
<!--
function Juge(myform)
{

	
}


function SelectAll() {
	for (var i=0;i<document.selform.selBigClass.length;i++) {
		var e=document.selform.selBigClass[i];
		e.checked=!e.checked;
	}
}
//-->
</script>


<!--media=print ������Կ����ڴ�ӡʱ��Ч--> 
<style media=print> 
.Noprint{display:none;} 
.PageNext{page-break-after: always;} 
</style>
</HEAD>
<BODY topMargin=0 rightMargin=0 leftMargin=0>       
        <table width="700" border="0" cellpadding="0" style="border-collapse: collapse" >
        <tr>
 <td colspan="2" align=center valign=middle height=70><font size="3">���ռ�����װ���޹�˾</font><br>
	<b><font size="6">�� �� �� �� ��</font></b></td>
</tr>
<tr>
<td colspan="2" height=23 align=left><font size=2>&nbsp;&nbsp; ���ڣ�<%=date()%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;<b><font size=4 color="#FF0000">&nbsp;&nbsp;��:<%
Response.Write(year(date()) & "" & month(date()) & "" & day(date()) & "" & "" & mid("",weekday(date()),1))
%><%Function gen_key(digits) 
dim char_array(50) 
char_array(0) = "0" 
char_array(1) = "1" 
char_array(2) = "2" 
char_array(3) = "3" 
char_array(4) = "4" 
char_array(5) = "5" 
char_array(6) = "6" 
char_array(7) = "7" 
char_array(8) = "8" 
char_array(9) = "9" 
randomize 
do while len(output) < digits 
nums = char_array(Int((9 - 0 + 1) * Rnd + 0)) 
output = output + nums 
loop 
gen_key = output 
End Function 
response.write "" & gen_key(3) & "" & vbcrlf 
%> 
</font></b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
</tr>

<tr>
<td>
<table width="700" border="1" cellpadding="0" style="border-collapse: collapse">

<tr>
<td height="30" align=center width=10%><b><font size="2">���</font></b></td>
<td height="30" align=center width=35%><b><font size="2">Ʒ��</font></b></td>
<td height="30" align=center width=10%><b><font size="2">���</font></b></td>
<td height="30" align=center width=15%><b><font size="2">����</font></b></td>
<td height="30" align=center width=10%><b><font size="2">��λ</font></b></td>
<td height="30" align=center width=10%><b><font size="2">����</font></b></td>
<td height="30" align=center width=10%><b><font size="2">���ó���</font></b></td>

</tr>
<% 

Sum = trim(request.form("selBigClass").count)
if Sum =0 then 
response.write "<script language=javascript>alert('δѡ����¼��');history.go(-1);</script>"
response.end
elseif sum>6 then
response.write "<script language=javascript>alert('���ֻ��ѡ���');history.go(-1);</script>"
response.end
else
for i = 1 to Sum
selBigClass= request.form("selBigClass")(i)
set rs=server.createobject("adodb.recordset")
sql="select * from cailiao_out_store where id in ("&selBigClass&")"
rs.open sql,conn,1,1

  %>

<tr>
<td height="30" align=center ><%=rs("fenlei")%><%=rs("pinming")%></td>
<td height="30" align=center ><%if rs("cname")="0000" then%>&nbsp;<%else%><%=rs("cname")%><%end if%></td>

<td height="30" align=center ><%=rs("guige")%>*<%=rs("houdou")%></td>

<td height="30" align=center><font size="2"><%=rs("loan_num")%></font></td>
<td height="30" align=center ><%=rs("unit")%></td>
<td height="30" align=center ><%if rs("pihao")<>"" then%><%=rs("pihao")%><%else%><%=rs("content")%><%end if%></td>
<td height="30" align=center ><%=rs("use_dep")%></td></tr>
<%
next
end if
%>




</table>








</td>
</tr>
<tr>
 <td colspan="8" align=right valign=bottom height=30>
	<p align="left"><br><font size="2"><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; �����ˣ�_______________&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�� �� �ˣ�</b><%=name%></font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
</tr>

				
	<table><tr><td height="30" colspan="8" align="center"><OBJECT id=WebBrowser classid=CLSID:8856F961-340A-11D0-A96B-00C04FD705A2 height=0 width=0 VIEWASTEXT> 
</OBJECT> 
<input type=button value=��ӡ onClick="document.all.WebBrowser.ExecWB(6,1)" class="NOPRINT"> 
<input type=button value=ֱ�Ӵ�ӡ onClick="document.all.WebBrowser.ExecWB(6,6)" class="NOPRINT"> 
<input type=button value=ҳ������ onClick="document.all.WebBrowser.ExecWB(8,1)" class="NOPRINT"> 
<input type=button value=��ӡԤ�� onClick="document.all.WebBrowser.ExecWB(7,1)" class="NOPRINT">
</td></tr>

				
				</table>     
		
</BODY></HTML>


<!--#include file="conn.asp"-->
<!--#include file="top.asp"-->
<!--#include file="checkuser.asp"-->
<head>
<style type="text/css">
<!--
.select {
	color: #FFFFFF;
	background-color: #08246b;
}
.offline {
	filter: Gray;
}
-->
</style>
<title>�˼�������ϵͳ��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="css/main.css" rel="stylesheet" type="text/css">

<SCRIPT language=JavaScript src="css/User_Info_Modify.js"></SCRIPT>

</head>
<body >

<TABLE width=800 align=center border="0" cellspacing="0" cellpadding="0" bordercolor="#0055E6">
  <tr><td align="center">
  <%if request("action")<>"search" then%>

  <table width="800"  border="1" cellpadding="0" cellspacing="0" bordercolor="#0055E6" align="center">

    <tr> 
      
      <td width="800" align="center" valign="top">
	        <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">��ϸ���ϳ�����Ϣ��ѯ</font></b>

</td>
</tr>
</table>

<table width="800" border="0" cellpadding="0" cellspacing="0">
	<form name="powersearch" method="post" action="fuhe_compiles.asp?action=search" >
  <tr><td height="30" >
	<p align="center"><input type="checkbox" name="SeachModel"   value=1>��Ա��</td><td>
  <select name='pihao'  style='width:130'>
			<option selected value="">��ѡ��Ʒ��</option>
			<%			
			set rs2=server.createobject("adodb.recordset")
		   rs2.open "select distinct(pihao) from fuhe_in order by pihao",conn,2,2
		   if not rs2.eof then
		   do while not rs2.eof
			%>
			<option value="<%=rs2("pihao")%>"><%=rs2("pihao")%></option>
			<%
			rs2.movenext
			loop
			end if
			rs2.close
			set rs2=nothing
			%></select></td></tr>
<tr>
 
  
<td height="30" >
<p align="center"><input type="checkbox" name=SeachModel value=2>������</td>
<td ><INPUT  name="uptime" onFocus="show_cele_date(uptime,'','',uptime)" value="<%=year(now())%>-<%=month(now())%>-<%=day(now())%>"  size="8"> 
				�� <INPUT  name="uptime1" onFocus="show_cele_date(uptime1,'','',uptime1)" value="<%=year(now())%>-<%=month(now())%>-<%=day(now())%>"  size="8"></td>

</tr>

<tr><td align="center" height="30" colspan="2"><input type="submit" name="Submit" value="��ѯ" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" title="��ѯ"></td></tr>
			</form>
	</table> 
</td>
</tr>
</table>
<br><br>
<TABLE width=800 align=center border="0" cellspacing="0" cellpadding="0" bordercolor="#0055E6">
<tr>
<td >
  <table width="800"  border="1" cellpadding="0" cellspacing="0" bordercolor="#0055E6" align="center">

    <tr> 
      
      <td width="800" align="center" valign="top">
	        <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">���³�����ܲ�ѯ</font></b>

</td>
</tr>
</table>

<table width="90%" border="0" cellpadding="0" cellspacing="0" bordercolor="#D4D0C8">
	<form name="powersearch" method="post" action="fuhe_compiles.asp?action=search" >
<tr><td  height="40" valign="middle" width="30%" align="center"><input type="checkbox" name=SeachModel value=3>ѡ���ѯ���ڣ�1����Ȼ�£� </td><td  width="25%"> <INPUT  name="uptime" onFocus="show_cele_date(uptime,'','',uptime)" value="<%=year(now())%>-<%=month(now())%>-<%=day(now())%>"  size="8"> �� <INPUT  name="uptime1" onFocus="show_cele_date(uptime1,'','',uptime1)" value="<%=year(now())%>-<%=month(now())%>-<%=day(now())%>"  size="8"></td><td  height="30"  width="50%"><input type="submit" name="Submit" value="��ѯ" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" title="��ѯ"><font color=red>* �磺2007��10����ѡ��2007-10-1��2007-10-31</font> </td></tr>
<tr ><td colspan="3" align=center></td></tr>

			</form>
	</table> 

</td>
</tr>
</table>
<br><br>
<TABLE width=800 align=center border="0" cellspacing="0" cellpadding="0" bordercolor="#0055E6">
<tr>
<td >

  <table width="800"  border="1" cellpadding="0" cellspacing="0" bordercolor="#0055E6" align="center">

    <tr> 
      
      <td width="800" align="center" valign="top">
	        <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">��ȳ�����ܲ�ѯ</font></b>

</td>
</tr>
</table>


<table width="90%" border="0" cellpadding="0" cellspacing="0" bordercolor="#D4D0C8">
	<form name="powersearch" method="post" action="fuhe_compiles.asp?action=search" >
<tr><td  height="40" valign="middle" width="30%" align="center"><input type="checkbox" name=SeachModel value=4>�������ѯ��� </td><td  width="30%"> <INPUT  name="nian"  value="<%=year(now())%>" onKeyUp="value=value.replace(/[^\d\.]/g,'')" maxlength="4" size="6"> </td><td  height="30"  width="40%"><input type="submit" name="Submit" value="��ѯ" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" title="��ѯ"></td></tr>
			</form>
	</table> 
	<%else%>

<%
dim SeachModel
SeachModel=request.form("SeachModel")
pihao=trim(request.form("pihao"))
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
<table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">�� ѯ �� �� �� ��</font></b>

</td>
</tr>
</table>
<table width="800" border="1" cellpadding="0" cellspacing="1" bordercolor="#EFEADD">
          <tr class="but"> 
                        <td width="10%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">����</td>
                        <td width="50%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">Ʒ ��</td>
                        <td width="10%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">��Ʒ</td>
                        <td width="10%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">���Ʒ</td>
                        <td width="10%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">Ա��</td>
			<td width="10%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">�� λ</td>

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
<tr "onmouseover="this.bgColor='#73A475';" onmouseout="this.bgColor=''"">
<td height="18" align="center"><%=rs("uptime")%></td>
<td height="18" align="center"><font color="#CC3300"><%=rs("cname")%></font><%if rs("cname")="" then%>&nbsp;<%end if%></td>
<td height="18" align="center"><%if rs("loan_oan")="0" then%>&nbsp;<%else%><%=rs("loan_oan")%><%end if%></td>
<td height="18" align="center"><%if rs("one_fu")="0" then%>&nbsp;<%else%><%=rs("one_fu")%><%end if%></td>
<td height="18" align="center"><%if rs("pihao")="" then%>&nbsp;<%else%><%=rs("pihao")%><%end if%></td>
<td height="18" align="center">ǧ��</td>


</tr>

                  <%
loan_oan=loan_oan+rs("loan_oan")
one_fu=one_fu+rs("one_fu")
numm=one_fu+loan_oan
rs.movenext
loop
end if
%>
                   <tr><td colspan="3" height="30" bgcolor="#FFFFFF">  ���ϲ�ѯ����������¼�ܹ�Ϊ <font color="#cc0000"><%=i%></font> ��</td>
					<td  align="right" bgcolor="#FFFFFF">
			�ܼƣ�</td><td align="center" bgcolor="#FFFFFF"><font color="#cc0000">
<%=numm%></font></td><td align="center" bgcolor="#FFFFFF">ǧ��</td></tr>
		
        </table>
		<br>
				<table cellpadding="0">
				<form name="search" method="post" action="fuhe_Detail.asp" target="_blank" >
				  <tr><td>
				  <input name="SeachModel" type="hidden" value="<%=SeachModel%>">
				  <input name="cname" type="hidden" value="<%=cname%>">
				  <input name="pihao" type="hidden" value="<%=pihao%>">
				  <input name="uptime" type="hidden" value="<%=uptime%>">
				  <input name="uptime1" type="hidden" value="<%=uptime1%>">
				  <input name="end_time" type="hidden" value="<%=end_time%>">
				  <input name="end_time1" type="hidden" value="<%=end_time1%>">
                                  <input name="nian" type="hidden" value="<%=nian%>">
				  <input type="submit" name="sub"  value="������ϸ��"></td></tr></form></table>
	<%end if%>
	</TD></TR>
  </table>
</body>
</html>
<!--#include file="footer.htm"--></TD>
 </BODY></HTML>

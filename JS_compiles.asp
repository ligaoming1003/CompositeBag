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
<b><font color="#ffffff">������Ϣ</font></b></td>
</tr>
</table>

<table width="800" border="0" cellpadding="0" cellspacing="0">
	<form name="powersearch" method="post" action="JS_Compiles.asp?action=search" >
  
<tr>
 
  <td height="30" >
<input type="checkbox" name="SeachModel"   value=1>���ͻ�</td>
<td><select name="use_dep"  style='width:130'>
			<option selected value="">��ѡ��ͻ�</option>
			<%			
			set rs=server.createobject("adodb.recordset")
		   rs.open "select distinct(use_dep) from JS_out_store order by use_dep",conn,2,2
		   if not rs.eof then
		   do while not rs.eof
			%>
			<option value="<%=rs("use_dep")%>"><%=rs("use_dep")%></option>
			<%
			rs.movenext
			loop
			end if
			rs.close
			set rs=nothing
			%></select></td>
        <td  height="30" ><input type="checkbox" name=SeachModel value=2>���������</td><td ><INPUT  name="uptime" onFocus="show_cele_date(uptime,'','',uptime)" value="<%=year(now())%>-<%=month(now())%>-<%=day(now())%>"  size="8"> �� <INPUT  name="uptime1" onFocus="show_cele_date(uptime1,'','',uptime1)" value="<%=year(now())%>-<%=month(now())%>-<%=day(now())%>"  size="8"></td>

</tr>

<tr><td align="center" height="30" colspan="4"><input type="submit" name="Submit" value="��ѯ" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" title="��ѯ"></td></tr>
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
<b><font color="#ffffff">���������ܲ�ѯ</font></b>

</td>
</tr>
</table>

<table width="90%" border="0" cellpadding="0" cellspacing="0" bordercolor="#D4D0C8">
	<form name="powersearch" method="post" action="JS_Compiles.asp?action=search" >
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
<b><font color="#ffffff">��������ܲ�ѯ</font></b>

</td>
</tr>
</table>


<table width="90%" border="0" cellpadding="0" cellspacing="0" bordercolor="#D4D0C8">
	<form name="powersearch" method="post" action="JS_Compiles.asp?action=search" >
<tr><td  height="40" valign="middle" width="30%" align="center"><input type="checkbox" name=SeachModel value=4>�������ѯ��� </td><td  width="30%"> <INPUT  name="nian"  value="<%=year(now())%>" onKeyUp="value=value.replace(/[^\d\.]/g,'')" maxlength="4" size="6"> </td><td  height="30"  width="40%"><input type="submit" name="Submit" value="��ѯ" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" title="��ѯ"></td></tr>
			</form>
	</table> 
	<%else%>

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
<table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">�� ѯ �� �� �� ��</font></b>

</td>
</tr>
</table>
<table width="800" border="1" cellpadding="0" cellspacing="1" bordercolor="#EFEADD">
          <tr class="but"> 
                        <td width="80" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">����</td>
                        <td width="80" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
						��&nbsp;&nbsp; ��</td>
                        <td width="160" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
						Ʒ&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; �� </td>
                        <td width="100" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
						�ͻ�����</td>
                        <td width="80" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
						�ͻ��շ�</td>
                        <td width="120" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
						��ע</td>
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
<tr "onmouseover="this.bgColor='#73A475';" onmouseout="this.bgColor=''"">
<td height="18" align="center"><%=rs("uptime")%></td>
<td align="center"><font color="#CC3300"><%=rs("use_dep")%></font></td>
<td align="center"><%=rs("pinming")%></td>
<td align="center"><%=Formatnumber(rs("Loan_Amount"),2,-1,0,0)%></td>
<td align="center"><%=Formatnumber(rs("hkdz"),2,-1,0,0)%></td>
<td align="center">��</td>

</tr>

                  <%
Loan_Amount=Loan_Amount+rs("Loan_Amount")
hkdz=hkdz+rs("hkdz")
rs.movenext
loop
end if
%>
                   <tr><td colspan="3" height="30" bgcolor="#FFFFFF">  
					<p align="right">��Ϊ <font color="#cc0000"><%=i%></font> �����ܼƣ�</td>
<td align="center" bgcolor="#FFFFFF"><font color="#cc0000">
<%=Loan_Amount%></font></td><td align="center" bgcolor="#FFFFFF"><%=hkdz%></td><td align="center" bgcolor="#FFFFFF"><%if Loan_Amount-hkdz>0 then%>�ͻ�Ƿ�<%=Loan_Amount-hkdz%>Ԫ<%else%>�ͻ����<%=hkdz-Loan_Amount%>Ԫ<%end if%></td></tr>
		
        </table>
		<br>
				<table cellpadding="0">
				<form name="search" method="post" action="JS_Detail.asp" target="_blank" >
				  <tr><td>
				  <input name="SeachModel" type="hidden" value="<%=SeachModel%>">
				  <input name="pinming" type="hidden" value="<%=pinming%>">
				  <input name="use_dep" type="hidden" value="<%=use_dep%>">
				  <input name="Loan_Amount" type="hidden" value="<%=Loan_Amount%>">
				  <input name="hkdz" type="hidden" value="<%=hkdz%>">
                  <input name="uptime" type="hidden" value="<%=uptime%>">
				  <input name="uptime1" type="hidden" value="<%=uptime1%>">
                  <input name="nian" type="hidden" value="<%=nian%>">
				  <input type="submit" name="sub"  value="������ϸ��"></td></tr></form></table>
	<%end if%>
	</TD></TR>
  </table>
</body>
</html>
<!--#include file="footer.htm"--></TD>
 </BODY></HTML>
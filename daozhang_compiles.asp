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

<TABLE width=800 align=center border="1" cellspacing="0" cellpadding="0" bordercolor="#0055E6">
  <tr><td align="center">
  <%if request("action")<>"search" then%>

	        <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">�������Ϣ��ѯ</font></b>

</td>
</tr>
</table>


<table width="800" border="0" cellpadding="0" cellspacing="0" bordercolor="#D4D0C8">
	<form name="powersearch" method="post" action="daozhang_Compiles.asp?action=search" >
  <tr>
 
  <td width="20%" height="30"><input type="checkbox" name=SeachModel value=1>���ͻ�</td>
            <td><select name="kehu" style='width:130' >
            	<option selected value="">��ѡ��ͻ�</option>
	<%			
			set rs2=server.createobject("adodb.recordset")
		   rs2.open "select distinct(kehu) from daozhang order by kehu",conn,1,2
		   if not rs2.eof then
		   do while not rs2.eof
			%>	
             <option value="<%=rs2("kehu")%>"><%=rs2("kehu")%></option>
<%
			rs2.movenext
			loop
			end if
			rs2.close
			set rs2=nothing
			%>	
</select>            </td>

<td height="30" >
<input type="checkbox" name="SeachModel"   value=2>�����˷�ʽ</td>
<td><select name="fanshi"  style='width:130'>
			<option selected value="">��ѡ��ʽ</option>
			<%			
			set rs3=server.createobject("adodb.recordset")
		   rs3.open "select distinct(fanshi) from daozhang order by fanshi",conn,2,2
		   if not rs3.eof then
		   do while not rs3.eof
			%>
			<option value="<%=rs3("fanshi")%>"><%=rs3("fanshi")%></option>
			<%
			rs3.movenext
			loop
			end if
			rs3.close
			set rs3=nothing
			%></select></td>
          </tr>
		 
			
			
<tr><td  height="30" ><input type="checkbox" name=SeachModel value=3>����������</td>
<td ><INPUT  name="uptime" onFocus="show_cele_date(uptime,'','',uptime)" value="<%=year(now())%>-<%=month(now())%>-<%=day(now())%>"  size="8"> �� <INPUT  name="uptime1" onFocus="show_cele_date(uptime1,'','',uptime1)" value="<%=year(now())%>-<%=month(now())%>-<%=day(now())%>"  size="8"></td>

<td  height="30" ><input type="checkbox" name="SeachModel"   value=4>����������</td>
<td>
  <select name='yinhang'  style='width:130'>
			<option selected value="">��ѡ������</option>
			<%			
			set rs2=server.createobject("adodb.recordset")
		   rs2.open "select distinct(yinhang) from daozhang order by yinhang",conn,2,2
		   if not rs2.eof then
		   do while not rs2.eof
			%>
			<option value="<%=rs2("yinhang")%>"><%=rs2("yinhang")%></option>
		<%
			rs2.movenext
			loop
			end if
			rs2.close
			set rs2=nothing
			%></select></td>

</tr><tr><td align="center" height="30" colspan="4"><input type="submit" name="Submit" value="��ѯ" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" title="��ѯ"></td></tr>
			</form>
	</table>

</td>
</tr>
</table>
<br><br>
<TABLE width=800 align=center border="1" cellspacing="0" cellpadding="0" bordercolor="#0055E6">
<tr>
<td>

	        <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">����˻��ܲ�ѯ</font></b>

</td>
</tr>
</table>

<table width="800" border="0" cellpadding="0" cellspacing="0" bordercolor="#D4D0C8">
	<form name="powersearch" method="post" action="daozhang_Compiles.asp?action=search" >
<tr><td  height="40" valign="middle" width="30%" align="center"><input type="checkbox" name=SeachModel value=5>ѡ���ѯ���ڣ�1����Ȼ�£� </td><td  width="25%"> <INPUT  name="uptime" onFocus="show_cele_date(uptime,'','',uptime)" value="<%=year(now())%>-<%=month(now())%>-<%=day(now())%>"  size="8"> �� <INPUT  name="uptime1" onFocus="show_cele_date(uptime1,'','',uptime1)" value="<%=year(now())%>-<%=month(now())%>-<%=day(now())%>"  size="8"></td><td  height="30"  width="50%"><input type="submit" name="Submit" value="��ѯ" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" title="��ѯ"><font color=red>* �磺2007��10����ѡ��2007-10-1��2007-10-31</font> </td></tr>
<tr ><td colspan="3" align=center></td></tr>

			</form>
	</table> 

</td>
</tr>
</table>
<br><br>
<TABLE width=800 align=center border="1" cellspacing="0" cellpadding="0" bordercolor="#0055E6">
<tr>
<td>


	        <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">��ȵ��˻��ܲ�ѯ</font></b>

</td>
</tr>
</table>


<table width="800" border="0" cellpadding="0" cellspacing="0" bordercolor="#D4D0C8">
	<form name="powersearch" method="post" action="daozhang_Compiles.asp?action=search" >
<tr><td  height="40" valign="middle" width="30%" align="center"><input type="checkbox" name=SeachModel value=6>�������ѯ��� </td><td  width="30%"> <INPUT  name="nian"  value="<%=year(now())%>" onKeyUp="value=value.replace(/[^\d\.]/g,'')" maxlength="4" size="6"> </td><td  height="30"  width="40%"><input type="submit" name="Submit" value="��ѯ" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" title="��ѯ"></td></tr>
			</form>
	</table> </fieldset>
	<%else%>

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
'��ѯ���뿪ʼ
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
<table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">�� ѯ �� �� �� ��</font></b>

</td>
</tr>
</table>
<table id = "PrintA" width="800" border="1" cellpadding="0" cellspacing="1" bordercolor="#EFEADD">
          <tr class="but"> 
                        <td width="10%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">��&nbsp;&nbsp; ��</td>
                        <td width="10%" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'"> ��&nbsp;&nbsp; ��</td>
                        <td width="20%" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">��&nbsp;&nbsp;&nbsp;&nbsp; ��</td>
                        <td width="15%" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">��&nbsp;&nbsp; ��</td>
			            <td width="10%" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">��&nbsp; ʽ</td>
			            <td width="20%" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">��&nbsp; ��</td>		
			            <td width="15%" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
						��&nbsp; ע</td>

			
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
<tr "onmouseover="this.bgColor='#73A475';" onmouseout="this.bgColor=''"">

<td height="18" align="center" ><%=rs("uptime")%></td>
<td align="center"><font color="#CC3300"><%=rs("leixin")%></font><%if rs("leixin")="" then%>&nbsp;<%end if%></td>
<td align="center"><a href="daozhang.Asp?id=<%=rs("ID")%>&action=edit"><%=rs("kehu")%></a></td>
<td align="center"><%=rs("Loan_num")%></td>
<td align="center"><font color="#009900"><%=rs("fanshi")%></font><%if rs("fanshi")="" then%>&nbsp;<%end if%></font></td>
<td align="center"><font color="#666666"><%=rs("yinhang")%></font><%if rs("yinhang")="" then%>&nbsp;<%end if%></font></td>
<td align="center"><font color="#666666"></font></td>

</tr>
				  
                  <%
Loan_num=Loan_num+rs("Loan_num")


rs.movenext
loop
end if
%>
                   <tr><td colspan="2" height="30" bgcolor="#FFFFFF">  ��¼�ܹ�Ϊ <font color="#cc0000"><%=i%></font> ��</td>
					<td align="center" bgcolor="#FFFFFF"><font color="#cc0000">�ϼ�:</font></td>
					<td align="center" bgcolor="#FFFFFF"><font color="#cc0000"><%=Loan_num%></font></td><td align="center" bgcolor="#FFFFFF"></td><td align="center" bgcolor="#FFFFFF"></td><td align="center" bgcolor="#FFFFFF"><font color="#cc0000"></font></td>
</tr>
		
        </table>
		<br>
				<table cellpadding="0">
				<form name="search" method="post" action="daozhang_Detail.asp" target="_blank" >
				  <tr><td>
				  <input name="SeachModel" type="hidden" value="<%=SeachModel%>">
				  <input name="leixin" type="hidden" value="<%=leixin%>">
				  <input name="kehu" type="hidden" value="<%=kehu%>">
				  <input name="fanshi" type="hidden" value="<%=fanshi%>">
				  <input name="yinhang" type="hidden" value="<%=yinhang%>">
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
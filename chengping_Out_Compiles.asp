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
<title>∷嘉美管理系统∷</title>
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
<b><font color="#ffffff">详细销售信息查询</font></b>

</td>
</tr>
</table>


<table width="800" border="0" cellpadding="0" cellspacing="0" bordercolor="#D4D0C8">
	<form name="powersearch" method="post" action="chengping_out_Compiles.asp?action=search" >
  <tr><td width="20%" height="30" ><input type="checkbox" name="SeachModel"   value=1>按分类</td><td>
  <select name='fenlei'  style='width:130'>
			<option selected value="">请选择类别</option>
			<%			
			set rs1=server.createobject("adodb.recordset")
		   rs1.open "select * from fenlei1",conn,1,1
		   if not rs1.eof then
		   do while not rs1.eof
			%>
			<option value="<%=rs1("fenlei")%>"><%=rs1("fenlei")%></option>
			<%
			rs1.movenext
			loop
			end if
			rs1.close
			set rs1=nothing
			%></select></td><td width="20%" height="30" ><input type="checkbox" name="SeachModel"   value=2>按品名</td><td>
  <select name='pinming'  style='width:130'>
			<option selected value="">请选择品名</option>
			<%			
			set rs2=server.createobject("adodb.recordset")
		   rs2.open "select distinct(pinming) from chengping_out_store order by pinming",conn,2,2
		   if not rs2.eof then
		   do while not rs2.eof
			%>
			<option value="<%=rs2("pinming")%>"><%=rs2("pinming")%></option>
		<%
			rs2.movenext
			loop
			end if
			rs2.close
			set rs2=nothing
			%></select></td></tr>
<tr>
 
  <td width="20%" height="30"><input type="checkbox" name=SeachModel value=4>按客户</td>
            <td><select name="use_dep" style='width:130' >
            	<option selected value="">请选择客户</option>
	<%			
			set rs2=server.createobject("adodb.recordset")
		   rs2.open "select distinct(use_dep) from chengping_out_store order by use_dep",conn,1,2
		   if not rs2.eof then
		   do while not rs2.eof
			%>	
             <option value="<%=rs2("use_dep")%>"><%=rs2("use_dep")%></option>
<%
			rs2.movenext
			loop
			end if
			rs2.close
			set rs2=nothing
			%>	
</select>            </td>

<td height="30" >
<input type="checkbox" name="SeachModel"   value=3>按规格</td>
<td><select name="guige"  style='width:130'>
			<option selected value="">请选择袋长</option>
			<%			
			set rs3=server.createobject("adodb.recordset")
		   rs3.open "select distinct(guige) from chengping_out_store order by guige",conn,2,2
		   if not rs3.eof then
		   do while not rs3.eof
			%>
			<option value="<%=rs3("guige")%>"><%=rs3("guige")%></option>
			<%
			rs3.movenext
			loop
			end if
			rs3.close
			set rs3=nothing
			%></select></td>
          </tr>
		 
			
			
			<tr><td  height="30" ><input type="checkbox" name=SeachModel value=5>按出库日期</td><td ><INPUT  name="uptime" onFocus="show_cele_date(uptime,'','',uptime)" value="<%=year(now())%>-<%=month(now())%>-<%=day(now())%>"  size="8"> 到 <INPUT  name="uptime1" onFocus="show_cele_date(uptime1,'','',uptime1)" value="<%=year(now())%>-<%=month(now())%>-<%=day(now())%>"  size="8"></td>

<td  height="30" >　</td>
<td>　</td>

</tr><tr><td align="center" height="30" colspan="8"><input type="submit" name="Submit" value="查询" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" title="查询"></td></tr>
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
<b><font color="#ffffff">单月销售汇总查询</font></b>

</td>
</tr>
</table>

<table width="800" border="0" cellpadding="0" cellspacing="0" bordercolor="#D4D0C8">
	<form name="powersearch" method="post" action="chengping_out_Compiles.asp?action=search" >
<tr><td  height="40" valign="middle" width="30%" align="center"><input type="checkbox" name=SeachModel value=7>选择查询日期（1个自然月） </td><td  width="25%"> <INPUT  name="uptime" onFocus="show_cele_date(uptime,'','',uptime)" value="<%=year(now())%>-<%=month(now())%>-<%=day(now())%>"  size="8"> 到 <INPUT  name="uptime1" onFocus="show_cele_date(uptime1,'','',uptime1)" value="<%=year(now())%>-<%=month(now())%>-<%=day(now())%>"  size="8"></td><td  height="30"  width="50%"><input type="submit" name="Submit" value="查询" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" title="查询"><font color=red>* 如：2007年10月则选择2007-10-1到2007-10-31</font> </td></tr>
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
<b><font color="#ffffff">年度销售汇总查询</font></b>

</td>
</tr>
</table>


<table width="800" border="0" cellpadding="0" cellspacing="0" bordercolor="#D4D0C8">
	<form name="powersearch" method="post" action="chengping_out_Compiles.asp?action=search" >
<tr><td  height="40" valign="middle" width="30%" align="center"><input type="checkbox" name=SeachModel value=8>请输入查询年份 </td><td  width="30%"> <INPUT  name="nian"  value="<%=year(now())%>" onKeyUp="value=value.replace(/[^\d\.]/g,'')" maxlength="4" size="6"> </td><td  height="30"  width="40%"><input type="submit" name="Submit" value="查询" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" title="查询"></td></tr>
			</form>
	</table> </fieldset>
	<%else%>

<%
dim SeachModel
SeachModel=request.form("SeachModel")
pinming=trim(request.form("pinming"))
use_dep=request.form("use_dep")
lingliaoyuan=request.form("lingliaoyuan")
fenlei1=request.form("fenlei")
guige=request.form("guige")
uptime=cstr(trim(request.form("uptime")))
uptime1=cstr(trim(request.form("uptime1")))
end_time=cstr(trim(request.form("end_time")))
end_time1=cstr(trim(request.form("end_time1")))
nian=cstr(trim(request.form("nian")))
'查询代码开始
names=Split(SeachModel,", ")
i=0
sql="select * from chengping_out_Store where"
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
sql=sql+" and guige = "&guige&""
end if
i=i+1
next
sql=sql+"order by -id "
set rs=server.createobject("adodb.recordset")
sql=Replace(sql, "where and", "where")

%>
<table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">查 询 结 果 列 表</font></b>

</td>
</tr>
</table>
<table width="800" border="1" cellpadding="0" cellspacing="1" bordercolor="#EFEADD">
          <tr class="but"> 
                        <td width="80" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">日期</td>
                        <td width="50" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'"> 类 别</td>
                        <td width="50" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">客户名称</td>
                        <td width="271" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">品 名</td>
                        <td width="56" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">规 格</td>
                        <td width="80" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">数 量</td>
			            <td width="40" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">单 位</td>
			            <td width="52" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
						单价</td>

			
			            <td width="80" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
						金额</td>

			
          </tr>
          
                  <% 
				  dim i
				  i=0
				  Loan_num=0
				  Loan_Amount=0
         rs.open sql,conn,1,1
		 if rs.eof and rs.bof then
		response.write "<tr><td><br><br><center>还没有符合条件的查询！</center><br><br></td></tr>"		                       
         else
		 do while not rs.eof
		 i=i+1
		   %>
<%
if i>5500 then
response.Write "<script language=javascript>alert('查询记录太多，请分段查询!');</script>"
response.write "<meta http-equiv=""refresh"" content=""0;url=chengping_out_compiles.asp"">"
response.end
end if 

		   %>
<tr "onmouseover="this.bgColor='#73A475';" onmouseout="this.bgColor=''"">

<td height="18" align="center" ><a href="chengping_out_Store.Asp?id=<%=rs("ID")%>&action=edit"><%=rs("uptime")%></a></td>

<td align="center"><font color="#CC3300"><%=rs("fenlei")%></font><%if rs("fenlei")="" then%>&nbsp;<%end if%></td>
<td align="center"><font color="#666666"><%=rs("use_dep")%></font></td>
<td align="center"><a href="chengping_out_Store.Asp?id=<%=rs("ID")%>&action=edit"><font color="#000099"><%=rs("pinming")%></font></a></td>
<td align="center">&nbsp;<%=rs("guige")%>*<%=rs("kuang")%></td>
<td align="center"><%=rs("Loan_num")%></td>
<td align="center">&nbsp;<font color="#009900"><%=rs("Unit")%></font><%if rs("Unit")="" then%>&nbsp;<%end if%></font></td>
<td align="center"><font color="#666666"><%=Formatnumber(rs("Loan_price"),4,-1,0,0)%></font></td>
<td align="center"><font color="#666666"><%=Formatnumber(rs("Loan_Amount"),2,-1,0,0)%></font></td>

</tr>
				  
<%
Loan_amount=Loan_amount+rs("Loan_amount")
rs.movenext
loop
end if
%>
                   <tr><td colspan="4" height="30" bgcolor="#FFFFFF">  符合查询条件的出库记录总共为 <font color="#cc0000"><%=i%></font> 条</td>
					<td align="center" bgcolor="#FFFFFF"><font color="#cc0000">合计:</font></td>
					<td align="center" bgcolor="#FFFFFF"><font color="#cc0000"></font></td><td align="center" bgcolor="#FFFFFF"></td><td align="center" bgcolor="#FFFFFF"></td><td align="center" bgcolor="#FFFFFF"><font color="#cc0000"><%=Formatnumber(Loan_amount,2,-1,0,0)%></font></td>
</tr>
		
        </table>
		<br>
				<table cellpadding="0">
				<form name="search" method="post" action="chengping_out_Detail.asp" target="_blank" >
				  <tr><td>
				  <input name="SeachModel" type="hidden" value="<%=SeachModel%>">
				  <input name="pinming" type="hidden" value="<%=pinming%>">
				  <input name="fenlei" type="hidden" value="<%=fenlei1%>">
				  <input name="guige" type="hidden" value="<%=guige%>">
				  <input name="use_dep" type="hidden" value="<%=use_dep%>">
				  <input name="lingliaoyuan" type="hidden" value="<%=lingliaoyuan%>">
				  <input name="uptime" type="hidden" value="<%=uptime%>">
				  <input name="uptime1" type="hidden" value="<%=uptime1%>">
				  <input name="end_time" type="hidden" value="<%=end_time%>">
				  <input name="end_time1" type="hidden" value="<%=end_time1%>">
                  <input name="nian" type="hidden" value="<%=nian%>">
				  <input type="submit" name="sub"  value="生成明细表"></td></tr></form></table>
	<%end if%>
	</TD></TR>
  </table>
</body>
</html>
<!--#include file="footer.htm"--></TD>
 </BODY></HTML>
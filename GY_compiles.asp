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

<TABLE width=800 align=center border="0" cellspacing="0" cellpadding="0" bordercolor="#0055E6">
  <tr><td align="center">
  <%if request("action")<>"search" then%>

  <table width="800"  border="1" cellpadding="0" cellspacing="0" bordercolor="#0055E6" align="center">

    <tr> 
      
      <td width="800" align="center" valign="top">
	        <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">结算信息</font></b></td>
</tr>
</table>

<table width="800" border="0" cellpadding="0" cellspacing="0">
	<form name="powersearch" method="post" action="GY_Compiles.asp?action=search" >
  
<tr>
 
  <td height="30" >
<input type="checkbox" name="SeachModel"   value=1>按供应商</td>
<td><select name="gonghuodanwei"  style='width:130'>
			<option selected value="">请选择供方</option>
			<%			
			set rs=server.createobject("adodb.recordset")
		   rs.open "select distinct(gonghuodanwei) from cailiao_in_store order by gonghuodanwei",conn,2,2
		   if not rs.eof then
		   do while not rs.eof
			%>
			<option value="<%=rs("gonghuodanwei")%>"><%=rs("gonghuodanwei")%></option>
			<%
			rs.movenext
			loop
			end if
			rs.close
			set rs=nothing
			%></select></td>
        <td  height="30" ><input type="checkbox" name=SeachModel value=2>按入库日期</td><td ><INPUT  name="uptime" onFocus="show_cele_date(uptime,'','',uptime)" value="<%=year(now())%>-<%=month(now())%>-<%=day(now())%>"  size="8"> 到 <INPUT  name="uptime1" onFocus="show_cele_date(uptime1,'','',uptime1)" value="<%=year(now())%>-<%=month(now())%>-<%=day(now())%>"  size="8"></td>

</tr>

<tr><td align="center" height="30" colspan="4"><input type="submit" name="Submit" value="查询" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" title="查询"></td></tr>
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
<b><font color="#ffffff">单月入库汇总查询</font></b>

</td>
</tr>
</table>

<table width="90%" border="0" cellpadding="0" cellspacing="0" bordercolor="#D4D0C8">
	<form name="powersearch" method="post" action="GY_Compiles.asp?action=search" >
<tr><td  height="40" valign="middle" width="30%" align="center"><input type="checkbox" name=SeachModel value=3>选择查询日期（1个自然月） </td><td  width="25%"> <INPUT  name="uptime" onFocus="show_cele_date(uptime,'','',uptime)" value="<%=year(now())%>-<%=month(now())%>-<%=day(now())%>"  size="8"> 到 <INPUT  name="uptime1" onFocus="show_cele_date(uptime1,'','',uptime1)" value="<%=year(now())%>-<%=month(now())%>-<%=day(now())%>"  size="8"></td><td  height="30"  width="50%"><input type="submit" name="Submit" value="查询" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" title="查询"><font color=red>* 如：2007年10月则选择2007-10-1到2007-10-31</font> </td></tr>
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
<b><font color="#ffffff">年度入库汇总查询</font></b>

</td>
</tr>
</table>


<table width="90%" border="0" cellpadding="0" cellspacing="0" bordercolor="#D4D0C8">
	<form name="powersearch" method="post" action="GY_Compiles.asp?action=search" >
<tr><td  height="40" valign="middle" width="30%" align="center"><input type="checkbox" name=SeachModel value=4>请输入查询年份 </td><td  width="30%"> <INPUT  name="nian"  value="<%=year(now())%>" onKeyUp="value=value.replace(/[^\d\.]/g,'')" maxlength="4" size="6"> </td><td  height="30"  width="40%"><input type="submit" name="Submit" value="查询" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" title="查询"></td></tr>
			</form>
	</table> 
	<%else%>

<%
dim SeachModel
SeachModel=request.form("SeachModel")
pinming=trim(request.form("pinming"))
jinge=trim(request.form("jinge"))
guige=trim(request.form("guige"))
houdou=trim(request.form("houdou"))
gonghuodanwei=request.form("gonghuodanwei")
fukuang=request.form("fukuang")
uptime=cstr(trim(request.form("uptime")))
uptime1=cstr(trim(request.form("uptime1")))
nian=cstr(trim(request.form("nian")))
'查询代码开始
names=Split(SeachModel,", ")
i=0
sql="select * from cailiao_in_Store where zhuangtai='ok'"
for each name in names
if names(i)="1" then
sql=sql+" and gonghuodanwei = '"&gonghuodanwei&"'"
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
<b><font color="#ffffff">查 询 结 果 列 表</font></b>

</td>
</tr>
</table>
<table width="800" border="1" cellpadding="0" cellspacing="1" bordercolor="#EFEADD">
          <tr class="but"> 
                        <td width="80" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">日期</td>
                        <td width="80" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
						供 应 商</td>
                        <td width="160" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
						品&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 名 </td>
                        <td width="100" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
						供方长款</td>
                        <td width="80" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
						供方欠款</td>
                        <td width="120" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
						备注</td>
          </tr>
                           <% 
				  dim i
				  i=0
				  jinge=0
                  fukuang=0
         rs.open sql,conn,1,1
		 if rs.eof and rs.bof then
		response.write "<tr><td><br><br><center>还没有符合条件的查询！</center><br><br></td></tr>"
		else
		 do while not rs.eof
		 i=i+1
		   %>
<tr "onmouseover="this.bgColor='#73A475';" onmouseout="this.bgColor=''"">
<td height="18" align="center"><%=rs("uptime")%></td>
<td align="center"><font color="#CC3300"><%=rs("gonghuodanwei")%></font></td>
<td align="center"><b><%=rs("pinming")%></b><font color="#FF0000">(<%=rs("guige")%>*<%=rs("houdou")%>)</font></td>
<td align="center"><%=Formatnumber(rs("jinge"),2,-1,0,0)%></td>
<td align="center"><%=Formatnumber(rs("fukuang"),2,-1,0,0)%></td>
<td align="center">　</td>

</tr>

                  <%
jinge=jinge+rs("jinge")
fukuang=fukuang+rs("fukuang")
rs.movenext
loop
end if
%>
                   <tr><td colspan="3" height="30" bgcolor="#FFFFFF">  
					<p align="right">共为 <font color="#cc0000"><%=i%></font> 条，总计：</td>
<td align="center" bgcolor="#FFFFFF"><font color="#cc0000">
<%=jinge%></font></td><td align="center" bgcolor="#FFFFFF"><%=fukuang%></td><td align="center" bgcolor="#FFFFFF"><%if jinge-fukuang>0 then%>供方长款：<%=jinge-fukuang%>元<%else%>供方欠款：<%=fukuang-jinge%>元<%end if%></td></tr>
		
        </table>
		<br>
				<table cellpadding="0">
				<form name="search" method="post" action="GY_Detail.asp" target="_blank" >
				  <tr><td>
				  <input name="SeachModel" type="hidden" value="<%=SeachModel%>">
				  <input name="pinming" type="hidden" value="<%=pinming%>">
				  <input name="gonghuodanwei" type="hidden" value="<%=gonghuodanwei%>">
				  <input name="guige" type="hidden" value="<%=guige%>">
                  <input name="houdou" type="hidden" value="<%=houdou%>">
				  <input name="jinge" type="hidden" value="<%=jinge%>">
				  <input name="fukuang" type="hidden" value="<%=fukuang%>">
                  <input name="uptime" type="hidden" value="<%=uptime%>">
				  <input name="uptime1" type="hidden" value="<%=uptime1%>">
                  <input name="nian" type="hidden" value="<%=nian%>">
				  <input type="submit" name="sub"  value="生成明细表"></td></tr></form></table>
	<%end if%>
	</TD></TR>
  </table>
</body>
</html>
<!--#include file="footer.htm"--></TD>
 </BODY></HTML>
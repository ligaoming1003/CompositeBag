

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

  <table width="100%" border="0" cellspacing="0" bordercolor="#D6D3CE">
  <tr><td align="center">
  <%if request("action")<>"search" then%>
<fieldset style="width:798px; height:76px"><legend>库存信息普通查询</legend>
<table width="100%" border="0" cellpadding="0" cellspacing="0" bordercolor="#D4D0C8">
	<form name="powersearch" method="post" action="cailiao_Compiles.asp?action=search" >
  <tr><td width="20%" height="30" ><input type="checkbox" name="SeachModel"   value=1>按分类</td><td>
  <select name='fenlei'  style='width:130'>
			<option selected value="">请选择类别</option>
			<%			
			set rs1=server.createobject("adodb.recordset")
		   rs1.open "select * from fenlei",conn,1,1
		   if not rs1.eof then
		   do while not rs1.eof
			%>
			<option value="<%=rs1("fenleiname")%>"><%=rs1("fenleiname")%></option>
			<%
			rs1.movenext
			loop
			end if
			rs1.close
			set rs1=nothing
			%></select></td>
			<td width="20%" height="30" >
			<input type="checkbox" name="SeachModel"   value=2>按品名</td><td>
  <select name='pinming'  style='width:130'>
			<option selected value="">请选择品名</option>
			<%			
			set rs2=server.createobject("adodb.recordset")
		   rs2.open "select distinct(pinming) from cailiao_store order by pinming",conn,2,2
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
<tr><td height="30" >
<input type="checkbox" name="SeachModel"   value=3>按规格</td>
<td><select name="guige"  style='width:130'>
			<option selected value="">请选择宽度</option>
			<%			
			set rs3=server.createobject("adodb.recordset")
		   rs3.open "select distinct(guige) from cailiao_store order by guige",conn,2,2
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
 
  <td align="center">
	<p align="left">
			<input type="checkbox" name="SeachModel"   value=4>按规格</td>
 
  <td align="center">
	<p align="left">
  <select name="houdou"  style='width:130'>
			<option selected value="">请选择厚度</option>
			<%			
			set rs4=server.createobject("adodb.recordset")
		   rs4.open "select distinct(houdou) from cailiao_store order by houdou",conn,2,2
		   		   if not rs4.eof then
		   do while not rs4.eof
			%>
			<option value="<%=rs4("houdou")%>"><%=rs4("houdou")%></option>
			<%
			rs4.movenext
			loop
			end if
			rs4.close
			set rs4=nothing
			%></select></td></tr>
<tr><td height="30" colspan="4" >
<p align="center"><input type="submit" name="Submit" value="查询" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" title="查询"></td>
	</tr>
			</form>
	</table> </fieldset>
	<br><br><br>

	<%else%>

<%
dim SeachModel
SeachModel=request.form("SeachModel")
pinming=trim(request.form("pinming"))
fenlei1=request.form("fenlei")
guige=request.form("guige")
houdou=request.form("houdou")

'查询代码开始
names=Split(SeachModel,", ")
i=0
sql="select * from cailiao_Store where number<>0"
for each name in names
if names(i)="1" then
sql=sql+" and fenlei = '"&fenlei1&"'"
end if
if names(i)="2" then
sql=sql+" and pinming like '%"&pinming&"%'"
end if
if names(i)="3" then
sql=sql+" and guige = "&guige&""
end if
if names(i)="4" then
sql=sql+" and houdou = "&houdou&""
end if

i=i+1
next
sql=sql+" order by pinming"
set rs=server.createobject("adodb.recordset")
sql=Replace(sql, "where and", "where")

%>

<table width="800" border="0" cellpadding="0" cellspacing="0" bgcolor="#D4D0C8" class="tddown">
          <tr class="but"> 
            
            <td width="20%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">类 别</td>
            <td width="20%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">品 名</td>
            <td width="20%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">规 格</td>
            <td width="20%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">数 量</td>
	    <td width="7%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">单价</td>
	    <td width="7%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">金额</td>
	    <td width="6%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">单 位</td>
          </tr>
          <tr valign="top" bgcolor="#FFFFFF"> 
            <td colspan="7" align="center" class="iframe"> 
                <table width="800" border="0" cellpadding="0" cellspacing="0" class="mouse">
                  <% 
				  dim i
				  i=0
				  num=0
         rs.open sql,conn,1,1
		 if rs.eof and rs.bof then
		response.write "<tr><td><br><br><center>还没有符合条件的查询！</center><br><br></td></tr>"
		else
		 do while not rs.eof
		 i=i+1
		   %>
                  <tr>
<td height="18" width="20%" align="center"  class="but1" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but1'"><font color="#CC3300"><%=rs("fenlei")%></font><%if rs("fenlei")="" then%>&nbsp;<%end if%></td>
<td width="20%" align="center" class="but1" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but1'">&nbsp;<%=rs("pinming")%></td>
<td width="20%" align="center" class="but1" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but1'">&nbsp;<%=rs("guige")%>*<%=rs("houdou")%></td>
<td width="20%" align="center" class="but1" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but1'"><%=Formatnumber((rs("number")-rs("anpai")),2,-1,0,0)%></td>
<td width="7%" align="center"  class="but1" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but1'"><%=Formatnumber(rs("danjia"),2,-1,0,0)%></td>
<td width="7%" align="center"  class="but1" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but1'"><%=Formatnumber(rs("jinge"),2,-1,0,0)%></td>
<td width="6%" align="center"  class="but1" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but1'"><font color="#009900"><%=rs("Unit")%></font><%if rs("Unit")="" then%>&nbsp;<%end if%>　</td>
</tr>
				  
                  <%
num=num+rs("number")
jinge=jinge+rs("jinge")
 rs.movenext

loop

end if
%>
                   <tr><td colspan="2" height="30">  符合查询条件的库存记录总共为 <font color="#cc0000"><%=i%></font> 条</td><td  align="center">总计：</td><td align="center"><font color="#cc0000"><%=num%></font><%=rs("Unit")%></td>
                   <td align="center"></td>
                   <td align="center"><font color="#cc0000"><%=Formatnumber(jinge,2,-1,0,0)%></font></td>
                   </tr>
				</table>
				
              </td>
          </tr>
        </table>
		<br>

	<%end if%>
	</TD></TR>
  </table>
</body>
</html>
</TD>
 </BODY></HTML><!--#include file="footer.htm"-->
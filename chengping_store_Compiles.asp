

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
	<form name="powersearch" method="post" action="chengping_store_Compiles.asp?action=search" >
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
		   rs2.open "select distinct(pinming) from chengping_store order by pinming",conn,2,2
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
			<option selected value="">请选择袋长</option>
			<%			
			set rs3=server.createobject("adodb.recordset")
		   rs3.open "select distinct(guige) from chengping_store order by guige",conn,2,2
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
 
  <td align="center" colspan="2"><input type="submit" name="Submit" value="查询" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" title="查询"></td></tr>
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

'查询代码开始
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
sql=sql+" and guige = "&guige&""
end if
i=i+1
next
sql=sql+""
set rs=server.createobject("adodb.recordset")
sql=Replace(sql, "where and", "where")

%>

<table width="800" border="0" cellpadding="0" cellspacing="0" bgcolor="#D4D0C8" class="tddown">
          <tr class="but"> 
            
            <td width="20%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">类 别</td>
            <td width="20%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">品 名</td>
            <td width="20%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">规 格</td>
            <td width="20%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">数 量</td>
	       <td width="20%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">单 位</td>
          </tr>
          <tr valign="top" bgcolor="#FFFFFF"> 
            <td colspan="5" align="center" class="iframe"> 
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
<td width="20%" align="center" class="but1" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but1'">&nbsp;<a href="Store.Asp?id=<%=rs("ID")%>&action=edit"><font color="#000099"><%=rs("pinming")%></font></a></td>
<td width="20%" align="center" class="but1" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but1'">&nbsp;<%=rs("guige")%>*<%=rs("kuang")%></td>
<td width="20%" align="center" class="but1" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but1'"><%=Formatnumber(rs("number"),2,-1,0,0)%></td>
<td width="20%" align="center"  class="but1" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but1'"><font color="#009900"><%=rs("Unit")%></font><%if rs("Unit")="" then%>&nbsp;<%end if%></td>
</tr>
				  
                  <%
num=num+rs("number")
rs.movenext
loop
end if
%>
                   <tr><td colspan="2" height="30">  符合查询条件的库存记录总共为 <font color="#cc0000"><%=i%></font> 条</td><td  align="center">总计：</td><td align="center"><font color="#cc0000"><%=num%></font><%=rs("Unit")%></td></tr>
				</table>
				
              </td>
          </tr>
        </table>
		<br>
				<table cellpadding="0">
				<form name="search" method="post" action="chengping_store_Detail.asp" target="_blank" >
				  <tr><td>
				  <input name="SeachModel" type="hidden" value="<%=SeachModel%>">
				  <input name="pinming" type="hidden" value="<%=pinming%>">
				  <input name="fenlei" type="hidden" value="<%=fenlei1%>">
				  <input name="guige" type="hidden" value="<%=guige%>">
				  <input type="submit" name="sub"  value="生成明细表"></td></tr></form></table>
	<%end if%>
	</TD></TR>
  </table>
</body>
</html>
</TD>
 </BODY></HTML><!--#include file="footer.htm"-->
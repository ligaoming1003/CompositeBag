

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
	<form name="powersearch" method="post" action="dangang_Compiles.asp?action=search" >
  <tr><td width="20%" height="30" ><input type="checkbox" name="SeachModel"   value=1>按分类</td><td>
  <select name='fenlei'  style='width:130'>
			<option selected value="">请选择类别</option>
			<%			
			set rs=server.createobject("adodb.recordset")
		   rs.open "select distinct(fenlei) from dangang order by fenlei",conn,2,2
		   if not rs.eof then
		   do while not rs.eof
			%>
			<option value="<%=rs("fenlei")%>"><%=rs("fenlei")%></option>
			<%
			rs.movenext
			loop
			end if
			rs.close
			set rs=nothing
			%></select></td>
			<td width="20%" height="30" >
			<input type="checkbox" name="SeachModel"   value=2>按品名</td><td>
  <input name='pinming' type="text"  size="21"></td></tr>
<tr><td height="30" >
<input type="checkbox" name="SeachModel"   value=3>按面膜</td>
<td><select name="nianbang"  style='width:130'>
			<option selected value="">请选择宽度</option>
			<%			
			set rs3=server.createobject("adodb.recordset")
		   rs3.open "select distinct(nianbang) from dangang order by nianbang",conn,2,2
		   if not rs3.eof then
		   do while not rs3.eof
			%>
			<option value="<%=rs3("nianbang")%>"><%=rs3("nianbang")%></option>
			<%
			rs3.movenext
			loop
			end if
			rs3.close
			set rs3=nothing
			%></select></td>
 
  <td align="center">
	<p align="left">
			<input type="checkbox" name="SeachModel"   value=4>按客户</td>
 
  <td align="center">
	<p align="left">
  <select name='kehu'  style='width:130'>
			<option selected value="">请选择客户</option>
			<%			
			set rs4=server.createobject("adodb.recordset")
		   rs4.open "select distinct(kehu) from dangang order by kehu",conn,2,2
		   		   if not rs4.eof then
		   do while not rs4.eof
			%>
			<option value="<%=rs4("kehu")%>"><%=rs4("kehu")%></option>
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
fenlei=request.form("fenlei")
nianbang=request.form("nianbang")
kehu=request.form("kehu")

'查询代码开始
names=Split(SeachModel,", ")
i=0
sql="select * from dangang "
for each name in names
if names(i)="1" then
sql=sql+" and fenlei = '"&fenlei&"'"
end if
if names(i)="2" then
sql=sql+" and pinming like '%"&pinming&"%'"
end if
if names(i)="3" then
sql=sql+" and nianbang = "&nianbang&""
end if
if names(i)="4" then
sql=sql+" and kehu = '"&kehu&"'"
end if

i=i+1
next
sql=sql+" order by pinming"
set rs=server.createobject("adodb.recordset")
sql=Replace(sql, "where and", "where")

%>

<table width="800" border="0" cellpadding="0" cellspacing="0" bgcolor="#D4D0C8" class="tddown">
          <tr class="but"> 
            
            <td width="12%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">类 别</td>
            <td width="40%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">品 名</td>
            <td width="12%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">规 格</td>
            <td width="29%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">材&nbsp;&nbsp;&nbsp;&nbsp; 料</td>
	        <td width="6%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">客 户</td>
          </tr>
          <tr valign="top" bgcolor="#FFFFFF"> 
            <td colspan="6" align="center" class="iframe"> 
                <table width="800" border="0" cellpadding="0" cellspacing="0" class="mouse">
                
 <% 
				  dim i
				  i=0
         rs.open sql,conn,1,1
		 if rs.eof and rs.bof then
		response.write "<tr><td><br><br><center>还没有符合条件的查询！</center><br><br></td></tr>"
		else
		 do while not rs.eof
		 i=i+1
		   %>

                  <tr>
<td height="18" width="12%" align="center"  class="but1" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but1'"><font color="#CC3300"><%=rs("fenlei")%></font><%if rs("fenlei")="" then%>&nbsp;<%end if%></td>
<td width="40%" align="center" class="but1" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but1'">&nbsp;<a href="dangang.Asp?id=<%=rs("ID")%>&action=edit"><font color="#000099"><%=rs("pinming")%></font></a></td>
<td width="12%" align="center" class="but1" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but1'">&nbsp;<%=rs("chang")%>*<%=rs("kuang")%></td>
<td width="29%" align="center" class="but1" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but1'">&nbsp;<%=rs("name1")%><% if rs("name2")="0" then%><%else%>/<%=rs("name2")%><%end if%><% if rs("name3")="0" then%><%else%>/<%=rs("name3")%><%end if%><% if rs("name4")="0" then%><%else%>/<%=rs("name4")%><%end if%></td>

<td width="6%" align="center"  class="but1" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but1'"><font color="#009900"><%=rs("kehu")%></font><%if rs("kehu")="" then%>&nbsp;<%end if%>　</td>
</tr>

<%
rs.movenext
loop
end if
%>
				  
                  
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